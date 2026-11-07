//! Native Windows Office COM automation used by Audit and the optional converter.
//!
//! The interface is deliberately small: one private Excel/WPS application,
//! sequential workbooks, full calculation, verified save, and deterministic
//! close.  It does not attach to the user's active Office session.

use anyhow::{bail, ensure, Context, Result};
use std::{
    collections::HashSet,
    ffi::c_void,
    mem,
    path::Path,
    ptr, thread,
    time::{Duration, Instant},
};
use windows_sys::{
    core::GUID,
    Win32::{
        Foundation::{
            CloseHandle, SysAllocString, SysFreeString, SysStringLen, FILETIME, VARIANT_BOOL,
            VARIANT_FALSE, VARIANT_TRUE,
        },
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize,
                EXCEPINFO, CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS,
            },
            Diagnostics::ToolHelp::{
                CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W,
                TH32CS_SNAPPROCESS,
            },
            Ole::DISPID_PROPERTYPUT,
            Threading::{
                GetProcessTimes, OpenProcess, TerminateProcess, WaitForSingleObject,
                PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
            },
            Variant::{
                VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I2, VT_I4,
                VT_NULL, VT_R8,
            },
        },
        UI::WindowsAndMessaging::GetWindowThreadProcessId,
    },
};

const IID_IDISPATCH: GUID = GUID::from_u128(0x0002040000000000_c000000000000046);
const IID_NULL: GUID = GUID::from_u128(0);
const LOCALE_USER_DEFAULT: u32 = 0x0400;
const XL_CALCULATION_AUTOMATIC: i32 = -4105;
const XL_CALCULATION_MANUAL: i32 = -4135;
const XL_DONE: i32 = 0;
const XL_PENDING: i32 = 2;
const PROCESS_SYNCHRONIZE_ACCESS: u32 = 0x0010_0000;
const WAIT_OBJECT_0: u32 = 0;

#[repr(C)]
struct DispatchInterface {
    vtable: *const DispatchVtable,
}

#[repr(C)]
struct DispatchVtable {
    query_interface: unsafe extern "system" fn(
        this: *mut DispatchInterface,
        iid: *const GUID,
        object: *mut *mut c_void,
    ) -> i32,
    add_ref: unsafe extern "system" fn(this: *mut DispatchInterface) -> u32,
    release: unsafe extern "system" fn(this: *mut DispatchInterface) -> u32,
    get_type_info_count:
        unsafe extern "system" fn(this: *mut DispatchInterface, count: *mut u32) -> i32,
    get_type_info: unsafe extern "system" fn(
        this: *mut DispatchInterface,
        index: u32,
        lcid: u32,
        type_info: *mut *mut c_void,
    ) -> i32,
    get_ids_of_names: unsafe extern "system" fn(
        this: *mut DispatchInterface,
        iid: *const GUID,
        names: *mut *const u16,
        count: u32,
        lcid: u32,
        ids: *mut i32,
    ) -> i32,
    invoke: unsafe extern "system" fn(
        this: *mut DispatchInterface,
        id: i32,
        iid: *const GUID,
        lcid: u32,
        flags: u16,
        params: *mut DISPPARAMS,
        result: *mut VARIANT,
        exception: *mut EXCEPINFO,
        argument_error: *mut u32,
    ) -> i32,
}

struct Dispatch {
    object: *mut DispatchInterface,
    /// DISPID lookup crosses the COM channel into Excel and is measurably
    /// expensive per call; member names on a dispatch are stable for its
    /// lifetime, so resolve each name once and reuse the id.
    dispid_cache: std::cell::RefCell<std::collections::HashMap<String, i32>>,
}

impl Dispatch {
    unsafe fn from_raw(raw: *mut c_void) -> Result<Self> {
        ensure!(
            !raw.is_null(),
            "Excel COM returned a null IDispatch interface"
        );
        Ok(Self {
            object: raw.cast(),
            dispid_cache: std::cell::RefCell::new(std::collections::HashMap::new()),
        })
    }

    fn invoke(&self, name: &str, arguments: &[OwnedVariant], flags: u16) -> Result<OwnedVariant> {
        let member_id = {
            let mut cache = self.dispid_cache.borrow_mut();
            if let Some(id) = cache.get(name) {
                *id
            } else {
                let name_wide = wide(name);
                let mut name_pointer = name_wide.as_ptr();
                let mut member_id = -1_i32;
                let get_ids = unsafe { (*(*self.object).vtable).get_ids_of_names };
                let result = unsafe {
                    get_ids(
                        self.object,
                        &IID_NULL,
                        &mut name_pointer,
                        1,
                        LOCALE_USER_DEFAULT,
                        &mut member_id,
                    )
                };
                ensure_hresult(result, format!("Excel COM lookup {name}"))?;
                cache.insert(name.to_owned(), member_id);
                member_id
            }
        };

        let mut raw_arguments = arguments
            .iter()
            .rev()
            .map(|argument| argument.0)
            .collect::<Vec<_>>();
        let mut named_put_id = DISPID_PROPERTYPUT;
        let is_property_put = flags & DISPATCH_PROPERTYPUT != 0;
        let mut parameters = DISPPARAMS {
            rgvarg: if raw_arguments.is_empty() {
                ptr::null_mut()
            } else {
                raw_arguments.as_mut_ptr()
            },
            rgdispidNamedArgs: if is_property_put {
                &mut named_put_id
            } else {
                ptr::null_mut()
            },
            cArgs: raw_arguments.len() as u32,
            cNamedArgs: u32::from(is_property_put),
        };
        let mut value = OwnedVariant::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        let invoke = unsafe { (*(*self.object).vtable).invoke };
        let result = unsafe {
            invoke(
                self.object,
                member_id,
                &IID_NULL,
                LOCALE_USER_DEFAULT,
                flags,
                &mut parameters,
                &mut value.0,
                &mut exception,
                &mut argument_error,
            )
        };
        if result < 0 {
            if let Some(fill_in) = exception.pfnDeferredFillIn {
                unsafe {
                    let _ = fill_in(&mut exception);
                }
            }
            let source = unsafe { exception_bstr_to_string(exception.bstrSource) };
            let description = unsafe { exception_bstr_to_string(exception.bstrDescription) };
            unsafe {
                for value in [
                    exception.bstrSource,
                    exception.bstrDescription,
                    exception.bstrHelpFile,
                ] {
                    if !value.is_null() {
                        SysFreeString(value);
                    }
                }
            }
            let details = match (source.trim(), description.trim()) {
                ("", "") => String::new(),
                ("", description) => format!("；COM 说明：{description}"),
                (source, "") => format!("；COM 来源：{source}"),
                (source, description) => {
                    format!("；COM 来源：{source}；COM 说明：{description}")
                }
            };
            bail!(
                "Excel COM call {name}失败（HRESULT 0x{:08X}{details}）",
                result as u32
            );
        }
        Ok(value)
    }

    fn invoke_method(&self, name: &str, arguments: &[OwnedVariant]) -> Result<OwnedVariant> {
        // Excel 的部分成员（Worksheets.Item、Workbook.Calculate 等）在
        // IDispatch 下以 METHOD|PROPERTYGET 组合暴露；VB 自动化客户端同样
        // 传组合标志。单 METHOD 会在这些成员上返回 0x80020003/0x80020006。
        self.invoke(name, arguments, DISPATCH_METHOD | DISPATCH_PROPERTYGET)
    }

    fn invoke_method_only(&self, name: &str, arguments: &[OwnedVariant]) -> Result<OwnedVariant> {
        // Workbooks.Open is a method, not a property getter. Excel raises
        // DISP_E_EXCEPTION (“Cannot get the Open property of the Workbooks
        // class”) when this member is invoked with METHOD|PROPERTYGET.
        self.invoke(name, arguments, DISPATCH_METHOD)
    }

    fn get(&self, name: &str) -> Result<OwnedVariant> {
        self.invoke(name, &[], DISPATCH_PROPERTYGET)
    }

    fn set(&self, name: &str, value: OwnedVariant) -> Result<()> {
        self.invoke(name, &[value], DISPATCH_PROPERTYPUT)?;
        Ok(())
    }

    fn get_dispatch(&self, name: &str) -> Result<Self> {
        self.get(name)?.into_dispatch()
    }

    fn get_i32(&self, name: &str) -> Result<i32> {
        self.get(name)?.as_i32(name)
    }
}

unsafe fn exception_bstr_to_string(value: windows_sys::core::BSTR) -> String {
    if value.is_null() {
        return String::new();
    }
    let length = unsafe { SysStringLen(value) as usize };
    let units = unsafe { std::slice::from_raw_parts(value, length) };
    String::from_utf16_lossy(units)
}

impl Drop for Dispatch {
    fn drop(&mut self) {
        if !self.object.is_null() {
            unsafe { ((*(*self.object).vtable).release)(self.object) };
        }
    }
}

struct OwnedVariant(VARIANT);

impl OwnedVariant {
    fn empty() -> Self {
        Self(VARIANT::default())
    }

    fn string(value: &str) -> Result<Self> {
        let wide = wide(value);
        let bstr = unsafe { SysAllocString(wide.as_ptr()) };
        ensure!(!bstr.is_null(), "为 Excel COM 字符串分配 BSTR 失败");
        let mut result = Self::empty();
        unsafe {
            result.0.Anonymous.Anonymous.vt = VT_BSTR;
            result.0.Anonymous.Anonymous.Anonymous.bstrVal = bstr;
        }
        Ok(result)
    }

    fn i32(value: i32) -> Self {
        let mut result = Self::empty();
        unsafe {
            result.0.Anonymous.Anonymous.vt = VT_I4;
            result.0.Anonymous.Anonymous.Anonymous.lVal = value;
        }
        result
    }

    fn boolean(value: bool) -> Self {
        let mut result = Self::empty();
        let value: VARIANT_BOOL = if value { VARIANT_TRUE } else { VARIANT_FALSE };
        unsafe {
            result.0.Anonymous.Anonymous.vt = VT_BOOL;
            result.0.Anonymous.Anonymous.Anonymous.boolVal = value;
        }
        result
    }

    fn as_i32(&self, label: &str) -> Result<i32> {
        let value = unsafe { &self.0.Anonymous.Anonymous };
        match value.vt {
            VT_I4 => Ok(unsafe { value.Anonymous.lVal }),
            VT_I2 => Ok(i32::from(unsafe { value.Anonymous.iVal })),
            other => bail!("Excel COM {label} 返回意外 Variant 类型 {other}"),
        }
    }

    fn as_optional_i32(&self, label: &str) -> Result<Option<i32>> {
        let value = unsafe { &self.0.Anonymous.Anonymous };
        match value.vt {
            VT_I4 => Ok(Some(unsafe { value.Anonymous.lVal })),
            VT_I2 => Ok(Some(i32::from(unsafe { value.Anonymous.iVal }))),
            // Excel often marshals Interior.Color / DisplayFormat color as
            // VT_R8 even though the RGB value is an integral 24-bit number.
            VT_R8 => {
                let number = unsafe { value.Anonymous.dblVal };
                ensure!(
                    number.is_finite()
                        && number.fract() == 0.0
                        && number >= f64::from(i32::MIN)
                        && number <= f64::from(i32::MAX),
                    "Excel COM {label} 返回非整数颜色值 {number}"
                );
                Ok(Some(number as i32))
            }
            VT_NULL | VT_EMPTY => Ok(None),
            other => bail!("Excel COM {label} 返回意外 Variant 类型 {other}"),
        }
    }

    fn as_string(&self, label: &str) -> Result<String> {
        let value = unsafe { &self.0.Anonymous.Anonymous };
        ensure!(
            value.vt == VT_BSTR,
            "Office COM {label} 返回意外 Variant 类型 {}",
            value.vt
        );
        let bstr = unsafe { value.Anonymous.bstrVal };
        if bstr.is_null() {
            return Ok(String::new());
        }
        let length = unsafe { SysStringLen(bstr) } as usize;
        let units = unsafe { std::slice::from_raw_parts(bstr, length) };
        String::from_utf16(units).context("Office COM 返回无效 UTF-16 文本")
    }

    fn into_dispatch(mut self) -> Result<Dispatch> {
        let value = unsafe { &mut self.0.Anonymous.Anonymous };
        ensure!(
            value.vt == VT_DISPATCH,
            "Excel COM 属性应返回 IDispatch，实际 Variant 类型为 {}",
            value.vt
        );
        let pointer = unsafe { value.Anonymous.pdispVal };
        value.vt = windows_sys::Win32::System::Variant::VT_EMPTY;
        unsafe { Dispatch::from_raw(pointer) }
    }
}

impl Drop for OwnedVariant {
    fn drop(&mut self) {
        unsafe { VariantClear(&mut self.0) };
    }
}

/// A private Office/WPS instance reused across sequential S1-F1 workbooks.
pub struct ExcelComSession {
    app: Option<Dispatch>,
    _apartment: ComApartment,
    application_name: String,
    process_id: Option<u32>,
    owns_process: bool,
    initialization_timings: serde_json::Value,
}

impl ExcelComSession {
    pub fn create(application_names: &[&str]) -> Result<Self> {
        let create_started = Instant::now();
        let (apartment, app, application_name, process_id, mut initialization_timings) =
            create_isolated_application(application_names)?;

        let mut session = Self {
            app: Some(app),
            _apartment: apartment,
            application_name,
            process_id: Some(process_id),
            owns_process: true,
            initialization_timings: serde_json::Value::Null,
        };
        let configuration_timings = session.configure()?;
        initialization_timings["application_configuration"] = configuration_timings;
        initialization_timings["session_create_ms"] =
            serde_json::json!(create_started.elapsed().as_millis());
        session.initialization_timings = initialization_timings;
        Ok(session)
    }

    pub fn application_name(&self) -> &str {
        &self.application_name
    }

    pub fn process_id(&self) -> Option<u32> {
        self.process_id
    }

    pub fn process_created(&self) -> Option<u64> {
        self.process_id.and_then(process_creation_ticks)
    }

    pub fn owns_process(&self) -> bool {
        self.owns_process
    }

    pub fn initialization_timings(&self) -> &serde_json::Value {
        &self.initialization_timings
    }

    /// Quit the private Office application and release this thread's COM
    /// apartment before the worker process exits. Returns per-stage timings
    /// and the Quit error, if Office refused the request.
    pub fn shutdown(&mut self) -> (serde_json::Value, Option<String>) {
        let shutdown_started = Instant::now();
        let mut quit_call_ms = 0_u64;
        let mut interface_release_ms = 0_u64;
        let quit_error = if let Some(app) = self.app.take() {
            let quit_started = Instant::now();
            let result = app.invoke_method("Quit", &[]);
            quit_call_ms = quit_started.elapsed().as_millis() as u64;
            let release_started = Instant::now();
            drop(app);
            interface_release_ms = release_started.elapsed().as_millis() as u64;
            result.err().map(|error| format!("{error:#}"))
        } else {
            None
        };
        let uninitialize_started = Instant::now();
        self._apartment.uninitialize();
        let com_uninitialize_ms = uninitialize_started.elapsed().as_millis() as u64;
        let shutdown_stages = serde_json::json!({
            "quit_call_ms": quit_call_ms,
            "interface_release_ms": interface_release_ms,
            "com_uninitialize_ms": com_uninitialize_ms,
            "total_ms": shutdown_started.elapsed().as_millis() as u64,
        });
        (shutdown_stages, quit_error)
    }

    /// Convert a spreadsheet through the same private native COM instance.
    pub fn convert_to_xlsx(&mut self, input: &Path, output: &Path) -> Result<()> {
        ensure!(
            input.is_file() && !output.exists(),
            "格式转换输入不存在或输出已存在"
        );
        let app = self.app.as_ref().expect("COM app initialized");
        let workbook = app
            .get_dispatch("Workbooks")?
            .invoke_method_only(
                "Open",
                &[
                    OwnedVariant::string(&input.to_string_lossy())?,
                    OwnedVariant::i32(0),
                    OwnedVariant::boolean(true),
                ],
            )?
            .into_dispatch()?;
        let result = (|| -> Result<()> {
            workbook.invoke_method(
                "SaveAs",
                &[
                    OwnedVariant::string(&output.to_string_lossy())?,
                    OwnedVariant::i32(51),
                ],
            )?;
            Ok(())
        })();
        let close_result = workbook.invoke_method("Close", &[OwnedVariant::boolean(false)]);
        result?;
        close_result.context("关闭格式转换工作簿")?;
        Ok(())
    }

    /// Read actual conditional-format fill rendering for the named scan areas.
    /// A single private Office instance is reused across all workbooks in the
    /// S1-F1 batch. Uniform rectangles are checked in bulk and recursively
    /// subdivided only when their stored/display colors differ within them.
    pub fn scan_conditional_display_colors(
        &mut self,
        input: &Path,
        scan_areas: &[String],
    ) -> Result<crate::conditional_format::ConditionalDisplayScan> {
        use crate::conditional_format::ConditionalDisplayScan;

        ensure!(input.is_file(), "条件格式显示色扫描输入不存在");
        ensure!(
            !scan_areas.is_empty(),
            "模板没有配置条件格式扫描区域；无法按实际显示色审核"
        );
        let started = Instant::now();
        let app = self.app.as_ref().context("Office COM 会话已关闭")?;
        let workbooks = app.get_dispatch("Workbooks")?;
        let workbook = workbooks
            .invoke_method_only(
                "Open",
                &[
                    OwnedVariant::string(&input.to_string_lossy())?,
                    OwnedVariant::i32(0),
                    OwnedVariant::boolean(true),
                ],
            )?
            .into_dispatch()
            .with_context(|| format!("Excel COM 无法只读打开条件格式来源 {}", input.display()))?;

        let is_wps = self.application_name.eq_ignore_ascii_case("ket.Application")
            || self.application_name.eq_ignore_ascii_case("KET.Application");
        let scan_result = (|| -> Result<ConditionalDisplayScan> {
            let sheets = workbook.get_dispatch("Worksheets")?;
            let mut scan = ConditionalDisplayScan::default();
            let mut seen_changes = HashSet::new();
            if is_wps {
                // WPS may return the underlying Interior color or an empty
                // DisplayFormat color while launched hidden with screen
                // updates disabled. The Python Audit refreshes WPS rendering
                // by activating the workbook/sheet and calculating before
                // reading DisplayFormat; mirror that behavior here.
                if let Err(error) = app.set("ScreenUpdating", OwnedVariant::boolean(true)) {
                    scan.diagnostics
                        .push(format!("WPS 条件格式刷新无法启用屏幕更新：{error:#}"));
                }
                if let Err(error) = workbook.invoke_method("Activate", &[]) {
                    scan.diagnostics
                        .push(format!("WPS 条件格式刷新无法激活工作簿：{error:#}"));
                }
            }
            let mut refreshed_wps_sheets = HashSet::new();
            for specification in scan_areas {
                let Some((raw_sheet_name, range_address)) = specification.rsplit_once('!') else {
                    scan.diagnostics.push(format!(
                        "条件格式扫描范围格式无效：{specification}"
                    ));
                    continue;
                };
                let sheet_name = raw_sheet_name.trim_matches('\'').replace("''", "'");
                let sheet = match sheets
                    .invoke_method("Item", &[OwnedVariant::string(&sheet_name)?])
                    .and_then(OwnedVariant::into_dispatch)
                {
                    Ok(sheet) => sheet,
                    Err(error) => {
                        scan.diagnostics.push(format!(
                            "条件格式扫描工作表不存在或不可读：{sheet_name}：{error:#}"
                        ));
                        continue;
                    }
                };
                if is_wps && refreshed_wps_sheets.insert(sheet_name.to_ascii_lowercase()) {
                    if let Err(error) = sheet.invoke_method("Activate", &[]) {
                        scan.diagnostics.push(format!(
                            "WPS 条件格式刷新无法激活工作表 {sheet_name}：{error:#}"
                        ));
                    }
                    match app.invoke_method("CalculateFullRebuild", &[]) {
                        Ok(_) => {}
                        Err(full_rebuild_error) => {
                            if let Err(calculate_error) = app.invoke_method("Calculate", &[]) {
                                scan.diagnostics.push(format!(
                                    "WPS 条件格式刷新计算失败：{sheet_name}；全量重建：{full_rebuild_error:#}；普通计算：{calculate_error:#}"
                                ));
                            }
                        }
                    }
                }
                let range = match sheet
                    .invoke_method("Range", &[OwnedVariant::string(range_address.trim())?])
                    .and_then(OwnedVariant::into_dispatch)
                {
                    Ok(range) => range,
                    Err(error) => {
                        scan.diagnostics.push(format!(
                            "条件格式扫描区域不可读：{specification}：{error:#}"
                        ));
                        continue;
                    }
                };
                let areas = match range.get_dispatch("Areas") {
                    Ok(areas) => areas,
                    Err(error) => {
                        scan.diagnostics.push(format!(
                            "条件格式扫描区域无法枚举：{specification}：{error:#}"
                        ));
                        continue;
                    }
                };
                let area_count = match areas.get_i32("Count") {
                    Ok(count) => count.max(0),
                    Err(error) => {
                        scan.diagnostics.push(format!(
                            "条件格式扫描区域数量无法读取：{specification}：{error:#}"
                        ));
                        continue;
                    }
                };
                for area_index in 1..=area_count {
                    let area = match areas
                        .invoke_method("Item", &[OwnedVariant::i32(area_index)])
                        .and_then(OwnedVariant::into_dispatch)
                    {
                        Ok(area) => area,
                        Err(error) => {
                            scan.diagnostics.push(format!(
                                "条件格式扫描区域第 {area_index} 段无法读取：{specification}：{error:#}"
                            ));
                            continue;
                        }
                    };
                    let dimensions = (|| -> Result<(i32, i32, i32, i32)> {
                        let row = area.get_i32("Row")?;
                        let column = area.get_i32("Column")?;
                        let rows = area.get_dispatch("Rows")?.get_i32("Count")?;
                        let columns = area.get_dispatch("Columns")?.get_i32("Count")?;
                        ensure!(row > 0 && column > 0 && rows > 0 && columns > 0, "Excel 返回了空条件格式扫描范围");
                        Ok((row, column, rows, columns))
                    })();
                    let (row, column, rows, columns) = match dimensions {
                        Ok(dimensions) => dimensions,
                        Err(error) => {
                            scan.diagnostics.push(format!(
                                "条件格式扫描区域尺寸无法读取：{specification}：{error:#}"
                            ));
                            continue;
                        }
                    };
                    if let Err(error) = scan_display_rectangle(
                        &sheet_name,
                        &sheet,
                        row,
                        column,
                        rows,
                        columns,
                        &mut scan,
                        &mut seen_changes,
                    ) {
                        scan.diagnostics.push(format!(
                            "条件格式实际显示色读取失败：{specification}：{error:#}"
                        ));
                    }
                }
            }
            scan.elapsed_ms = started.elapsed().as_millis();
            scan.engine_name = self.application_name.clone();
            scan.engine_process_id = self.process_id;
            Ok(scan)
        })();
        let reset_screen_result = if is_wps {
            app.set("ScreenUpdating", OwnedVariant::boolean(false))
        } else {
            Ok(())
        };
        let scan_result = match (scan_result, reset_screen_result) {
            (Ok(scan), Ok(())) => Ok(scan),
            (Ok(mut scan), Err(error)) => {
                scan.diagnostics
                    .push(format!("WPS 条件格式扫描后无法关闭屏幕更新：{error:#}"));
                Ok(scan)
            }
            (Err(error), Ok(())) => Err(error),
            (Err(error), Err(reset_error)) => Err(error.context(format!(
                "WPS 条件格式扫描后也无法关闭屏幕更新：{reset_error:#}"
            ))),
        };
        let close_result = workbook.invoke_method("Close", &[OwnedVariant::boolean(false)]);
        match (scan_result, close_result) {
            (Ok(scan), Ok(_)) => Ok(scan),
            (Err(error), _) => Err(error),
            (Ok(_), Err(error)) => Err(error).context("关闭只读条件格式工作簿"),
        }
    }

    pub fn recalculate(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
        stage_path: &Path,
    ) -> Result<serde_json::Value> {
        self.recalculate_with_mode(input, output, timeout, stage_path, "full", false)
    }

    /// `mode` selects the calculation invocation so benchmarks can attribute
    /// cost to each semantic: `full` is the production sequence (explicit
    /// automatic mode + Application.CalculateFull + completion polling),
    /// `full_no_set` skips the redundant mode set, and `sheet` mirrors the
    /// legacy Python production semantics (workbook-level Worksheets.Calculate).
    pub fn recalculate_with_mode(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
        stage_path: &Path,
        mode: &str,
        save_in_place: bool,
    ) -> Result<serde_json::Value> {
        write_stage(stage_path, "opening_workbook");
        let result = self.recalculate_inner(
            input,
            output,
            timeout,
            stage_path,
            mode,
            save_in_place,
        );
        if result.is_err() {
            write_stage(stage_path, "closing_failed_workbook");
        }
        result.map(|(stages,)| stages)
    }

    fn configure(&mut self) -> Result<serde_json::Value> {
        let app = self.app.as_ref().expect("COM app initialized");
        let total_started = Instant::now();
        let visible_ms = set_timed(app, "Visible", OwnedVariant::boolean(false))?;
        let display_alerts_ms = set_timed(app, "DisplayAlerts", OwnedVariant::boolean(false))?;
        // 与旧 Python 生产通道一致的 Excel 端开销配置：关闭屏幕刷新与事件、
        // AskToUpdateLinks 关闭（UpdateLinks 已在 Open 参数显式控制）、
        // 计算模式设为手动以消除 Workbooks.Open 对脏公式的隐式自动重算
        // （审核副本全部公式为待算状态；显式 Calculate 随后一次性同步完成，
        // 旧 Python 生产语义与此完全一致，Calculation 为应用级属性不随
        // SaveAs 写入输出文件）。
        let screen_updating_ms = set_timed(app, "ScreenUpdating", OwnedVariant::boolean(false))?;
        let enable_events_ms = set_timed(app, "EnableEvents", OwnedVariant::boolean(false))?;
        let ask_to_update_links_ms =
            set_timed(app, "AskToUpdateLinks", OwnedVariant::boolean(false))?;
        // 某些含数据模型/OLAP 连接的 Excel 实例禁止通过 COM 修改
        // Application.Calculation（0x800A03EC）；与旧 Python 一致降级忽略，
        // 显式 Calculate 仍完成重算。
        let calculation_started = Instant::now();
        let calculation_manual_succeeded = app
            .set("Calculation", OwnedVariant::i32(XL_CALCULATION_MANUAL))
            .is_ok();
        let calculation_manual_ms = calculation_started.elapsed().as_millis() as u64;
        Ok(serde_json::json!({
            "visible_ms": visible_ms,
            "display_alerts_ms": display_alerts_ms,
            "screen_updating_ms": screen_updating_ms,
            "enable_events_ms": enable_events_ms,
            "ask_to_update_links_ms": ask_to_update_links_ms,
            "calculation_manual_ms": calculation_manual_ms,
            "calculation_manual_succeeded": calculation_manual_succeeded,
            "total_ms": total_started.elapsed().as_millis() as u64,
        }))
    }

    fn recalculate_inner(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
        stage_path: &Path,
        mode: &str,
        save_in_place: bool,
    ) -> Result<(serde_json::Value,)> {
        // Stage-level timing (goal-mandated): every COM phase is measured so
        // hot/cold differences against the Python channel are attributable.
        let mut stages = serde_json::Map::new();
        let mut record = |name: &str, started: Instant| {
            stages.insert(
                name.to_owned(),
                serde_json::json!(started.elapsed().as_millis() as u64),
            );
        };

        let app = self.app.as_ref().expect("COM app initialized");
        let open_started = Instant::now();
        let workbooks = app.get_dispatch("Workbooks")?;
        let input_text = input.to_string_lossy();
        let workbook = workbooks
            .invoke_method_only(
                "Open",
                &[
                    OwnedVariant::string(&input_text)?,
                    OwnedVariant::i32(0),
                    // The engine input is a calculation-only copy. Open it
                    // read-only so workbook write protection/lock state cannot
                    // block formula evaluation or cause Excel to modify it.
                    OwnedVariant::boolean(!save_in_place),
                ],
            )
            .context("Office COM Workbooks.Open 调用失败")?
            .into_dispatch()
            .context("Excel COM Workbooks.Open 未返回工作簿")?;
        record("open", open_started);

        let calculation = (|| -> Result<()> {
            // No explicit Calculation-mode set: a fresh Excel instance already
            // defaults to automatic mode, CalculateFull is mode-independent,
            // and the old Python production channel never set it either. The
            // set was measured at ~150ms per workbook (it forces a synchronous
            // recalculation pass on the just-opened dirty workbook).

            // 逐工作表 Calculate：与旧 Python 生产通道一致（应用级
            // CalculateFull 会把所有已开工作簿纳入同一计算队列），并在
            // 计算完成后仍以 CalculationState 轮询拒绝 xlPending。
            let calc_started = Instant::now();
            write_stage(stage_path, "calculating_workbook_sheets");
            // 逐工作表 Calculate（集合级 Calculate 在 IDispatch 按名解析下
            // 返回 0x80020006，仅逐表调用可用；与旧 Python 生产语义一致）。
            let sheets = workbook.get_dispatch("Worksheets")?;
            let sheet_count = sheets.get_i32("Count")?;
            for sheet_index in 1..=sheet_count {
                let worksheet = sheets
                    .invoke_method("Item", &[OwnedVariant::i32(sheet_index)])?
                    .into_dispatch()
                    .with_context(|| {
                        format!("Excel COM Worksheets.Item({sheet_index}) 未返回工作表")
                    })?;
                worksheet.invoke_method("Calculate", &[])?;
            }
            record("sheets_calculate", calc_started);

            let connections_started = Instant::now();
            let connections = workbook.get_dispatch("Connections")?;
            if connections.get_i32("Count")? > 0 {
                write_stage(stage_path, "waiting_for_async_queries");
                app.invoke_method("CalculateUntilAsyncQueriesDone", &[])?;
            }
            record("connections_check", connections_started);

            let wait_started = Instant::now();
            let started = Instant::now();
            let mut last_state = -1;
            let mut full_rebuild_retried = false;
            loop {
                let state = app.get_i32("CalculationState")?;
                if state == XL_DONE {
                    write_stage(stage_path, "calculation_complete");
                    break;
                }
                if state == XL_PENDING && !full_rebuild_retried {
                    full_rebuild_retried = true;
                    write_stage(stage_path, "retrying_full_rebuild_after_pending");
                    app.invoke_method("CalculateFullRebuild", &[])?;
                    thread::sleep(Duration::from_millis(200));
                    continue;
                }
                ensure!(
                    state != XL_PENDING,
                    "Excel calculation remains xlPending after full dependency rebuild; refusing to save"
                );
                ensure!(
                    started.elapsed() < timeout,
                    "Excel calculation did not reach xlDone within {} seconds; last state={state}",
                    timeout.as_secs()
                );
                if state != last_state {
                    write_stage(
                        stage_path,
                        &format!("waiting_for_calculation_state:{state}"),
                    );
                    last_state = state;
                }
                thread::sleep(Duration::from_millis(200));
            }

            let save_started = Instant::now();
            write_stage(stage_path, "saving_output");
            if save_in_place {
                workbook.invoke_method("Save", &[])?;
            } else {
                let output_text = output.to_string_lossy();
                workbook.invoke_method(
                    "SaveAs",
                    &[OwnedVariant::string(&output_text)?, OwnedVariant::i32(51)],
                )?;
            }
            record("save_as", save_started);

            let close_started = Instant::now();
            write_stage(stage_path, "output_saved");
            // SaveAs 已落盘；Close 传 SaveChanges=False 避免 Close 阶段
            // 触发二次保存检查（与旧 Python 通道一致）。
            workbook.invoke_method("Close", &[OwnedVariant::boolean(false)])?;
            write_stage(stage_path, "workbook_closed");
            record("close", close_started);
            Ok(())
        })();

        if calculation.is_err() {
            let _ = workbook.invoke_method("Close", &[OwnedVariant::boolean(false)]);
        }
        calculation.map(|()| (serde_json::Value::Object(stages),))
    }
}

fn scan_display_rectangle(
    sheet_name: &str,
    sheet: &Dispatch,
    row: i32,
    column: i32,
    rows: i32,
    columns: i32,
    scan: &mut crate::conditional_format::ConditionalDisplayScan,
    seen_changes: &mut HashSet<String>,
) -> Result<()> {
    use crate::conditional_format::ConditionalDisplayColorChange;

    let end_row = row
        .checked_add(rows - 1)
        .context("条件格式范围行号溢出")?;
    let end_column = column
        .checked_add(columns - 1)
        .context("条件格式范围列号溢出")?;
    let address = format!(
        "{}:{}",
        excel_cell_address(column, row)?,
        excel_cell_address(end_column, end_row)?
    );
    let range = sheet
        .invoke_method("Range", &[OwnedVariant::string(&address)?])?
        .into_dispatch()?;
    let base_color = range_fill_color(&range, false)?;
    let display_color = range_fill_color(&range, true)?;

    if let (Some(base_color), Some(display_color)) = (base_color, display_color) {
        let cell_count = (rows as usize)
            .checked_mul(columns as usize)
            .context("条件格式扫描单元格数量溢出")?;
        scan.scanned_cells = scan
            .scanned_cells
            .checked_add(cell_count)
            .context("条件格式扫描单元格累计数量溢出")?;
        if base_color != display_color {
            for cell_row in row..=end_row {
                for cell_column in column..=end_column {
                    let cell = excel_cell_address(cell_column, cell_row)?;
                    let key = format!("{}!{}", sheet_name.to_ascii_lowercase(), cell);
                    if seen_changes.insert(key) {
                        scan.changes.push(ConditionalDisplayColorChange {
                            sheet: sheet_name.to_owned(),
                            cell,
                            base_fill_color: base_color,
                            display_fill_color: display_color,
                        });
                    }
                }
            }
        }
        return Ok(());
    }

    if rows == 1 && columns == 1 {
        scan.scanned_cells = scan.scanned_cells.saturating_add(1);
        // A one-cell range cannot contain mixed formatting. If ColorIndex is
        // nevertheless unavailable, use that cell's scalar Color directly.
        // On larger ranges Color alone is unsafe: Excel can return a scalar
        // even when cells in the range have different displayed fills.
        let base_color = range_fill_color_value(&range, false)?;
        let display_color = range_fill_color_value(&range, true)?;
        if let (Some(base_color), Some(display_color)) = (base_color, display_color) {
            if base_color != display_color {
                let cell = excel_cell_address(column, row)?;
                let key = format!("{}!{}", sheet_name.to_ascii_lowercase(), cell);
                if seen_changes.insert(key) {
                    scan.changes.push(ConditionalDisplayColorChange {
                        sheet: sheet_name.to_owned(),
                        cell,
                        base_fill_color: base_color,
                        display_fill_color: display_color,
                    });
                }
            }
            return Ok(());
        }
        scan.diagnostics.push(format!(
            "{}!{} 的基础/实际显示填充色不是单一可读颜色",
            sheet_name,
            excel_cell_address(column, row)?
        ));
        return Ok(());
    }

    if rows >= columns && rows > 1 {
        let first_rows = rows / 2;
        let second_rows = rows - first_rows;
        scan_display_rectangle(
            sheet_name,
            sheet,
            row,
            column,
            first_rows,
            columns,
            scan,
            seen_changes,
        )?;
        scan_display_rectangle(
            sheet_name,
            sheet,
            row + first_rows,
            column,
            second_rows,
            columns,
            scan,
            seen_changes,
        )
    } else {
        let first_columns = columns / 2;
        let second_columns = columns - first_columns;
        scan_display_rectangle(
            sheet_name,
            sheet,
            row,
            column,
            rows,
            first_columns,
            scan,
            seen_changes,
        )?;
        scan_display_rectangle(
            sheet_name,
            sheet,
            row,
            column + first_columns,
            rows,
            second_columns,
            scan,
            seen_changes,
        )
    }
}

fn range_fill_color(range: &Dispatch, display_format: bool) -> Result<Option<i32>> {
    let interior = if display_format {
        range.get_dispatch("DisplayFormat")?.get_dispatch("Interior")?
    } else {
        range.get_dispatch("Interior")?
    };
    let index_label = if display_format {
        "DisplayFormat.Interior.ColorIndex"
    } else {
        "Interior.ColorIndex"
    };
    // ColorIndex is VT_NULL for a mixed-format range, while Color can still
    // misleadingly return one cell's color. Use it as the uniformity check.
    if interior.get("ColorIndex")?.as_optional_i32(index_label)?.is_none() {
        return Ok(None);
    }
    range_fill_color_value(range, display_format)
}

fn range_fill_color_value(range: &Dispatch, display_format: bool) -> Result<Option<i32>> {
    let interior = if display_format {
        range.get_dispatch("DisplayFormat")?.get_dispatch("Interior")?
    } else {
        range.get_dispatch("Interior")?
    };
    interior.get("Color")?.as_optional_i32(if display_format {
        "DisplayFormat.Interior.Color"
    } else {
        "Interior.Color"
    })
}

fn excel_cell_address(column: i32, row: i32) -> Result<String> {
    ensure!(
        column > 0 && row > 0 && column <= 16_384 && row <= 1_048_576,
        "条件格式单元格坐标越界"
    );
    let mut value = column;
    let mut letters = String::new();
    while value > 0 {
        value -= 1;
        letters.insert(0, char::from(b'A' + (value % 26) as u8));
        value /= 26;
    }
    Ok(format!("{letters}{row}"))
}

/// A private Word/WPS COM instance used for document conversion and checks.
pub struct WordComSession {
    app: Option<Dispatch>,
    /// 本会话创建（并拥有）的 Word/WPS 进程 ID，作为批次复用证据外露。
    process_id: u32,
    _apartment: ComApartment,
}

impl WordComSession {
    pub fn create(application_names: &[&str]) -> Result<Self> {
        let (apartment, app, _, process_id, _) = create_isolated_application(application_names)?;
        let session = Self {
            app: Some(app),
            process_id,
            _apartment: apartment,
        };
        let app = session.app.as_ref().expect("COM app initialized");
        app.set("Visible", OwnedVariant::boolean(false))?;
        app.set("DisplayAlerts", OwnedVariant::i32(0))?;
        Ok(session)
    }

    /// 本会话拥有的 Word/WPS 进程 ID（批次会话复用证据）。
    pub fn process_id(&self) -> u32 {
        self.process_id
    }

    fn open_read_only(&self, input: &Path) -> Result<Dispatch> {
        ensure!(input.is_file(), "Word 格式转换输入不存在");
        self.app
            .as_ref()
            .expect("COM app initialized")
            .get_dispatch("Documents")?
            .invoke_method(
                "Open",
                &[
                    OwnedVariant::string(&input.to_string_lossy())?,
                    OwnedVariant::boolean(false),
                    OwnedVariant::boolean(true),
                ],
            )?
            .into_dispatch()
    }

    pub fn convert_to_docx(&self, input: &Path, output: &Path) -> Result<()> {
        ensure!(!output.exists(), "Word 格式转换输出已存在");
        let document = self.open_read_only(input)?;
        let result = (|| -> Result<()> {
            let args = [
                OwnedVariant::string(&output.to_string_lossy())?,
                OwnedVariant::i32(16),
            ];
            if document.invoke_method("SaveAs2", &args).is_err() {
                document.invoke_method("SaveAs", &args)?;
            }
            Ok(())
        })();
        let close_result = document.invoke_method("Close", &[OwnedVariant::boolean(false)]);
        result?;
        close_result.context("关闭格式转换文档")?;
        Ok(())
    }

    pub fn read_text(&self, input: &Path) -> Result<String> {
        let document = self.open_read_only(input)?;
        let result = document
            .get_dispatch("Content")
            .and_then(|content| content.get("Text"))
            .and_then(|value| value.as_string("Content.Text"));
        let close_result = document.invoke_method("Close", &[OwnedVariant::boolean(false)]);
        let text = result?;
        close_result.context("关闭文本校验文档")?;
        Ok(text)
    }
}

impl Drop for WordComSession {
    fn drop(&mut self) {
        if let Some(app) = self.app.take() {
            let _ = app.invoke_method("Quit", &[]);
            drop(app);
        }
    }
}

impl Drop for ExcelComSession {
    fn drop(&mut self) {
        let _ = self.shutdown();
    }
}

struct ComApartment {
    initialized: bool,
}

impl ComApartment {
    fn initialize() -> Result<Self> {
        let result = unsafe {
            CoInitializeEx(
                ptr::null(),
                (COINIT_APARTMENTTHREADED | 4) as u32, // COINIT_DISABLE_OLE1DDE
            )
        };
        ensure_hresult(result, "initialize the Excel COM apartment")?;
        Ok(Self { initialized: true })
    }

    fn uninitialize(&mut self) {
        if self.initialized {
            unsafe { CoUninitialize() };
            self.initialized = false;
        }
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        self.uninitialize();
    }
}

fn create_application(
    application_names: &[&str],
) -> Result<(Dispatch, String, serde_json::Value, u128)> {
    let mut failures = Vec::new();
    let mut progid_lookup_ms = 0_u64;
    let mut cocreate_instance_ms = 0_u64;
    let mut dispatch_wrap_ms = 0_u64;
    for application_name in application_names {
        let mut prog_id = wide(application_name);
        let mut class_id = GUID::default();
        let lookup_started = Instant::now();
        let result = unsafe { CLSIDFromProgID(prog_id.as_mut_ptr(), &mut class_id) };
        progid_lookup_ms += lookup_started.elapsed().as_millis() as u64;
        if result < 0 {
            failures.push(format!(
                "{application_name}: ProgID HRESULT 0x{:08X}",
                result as u32
            ));
            continue;
        }
        let mut object = ptr::null_mut();
        let activation_started_unix_ms = system_time_unix_ms();
        let activation_started = Instant::now();
        let result = unsafe {
            CoCreateInstance(
                &class_id,
                ptr::null_mut(),
                CLSCTX_LOCAL_SERVER,
                &IID_IDISPATCH,
                &mut object,
            )
        };
        cocreate_instance_ms += activation_started.elapsed().as_millis() as u64;
        if result < 0 {
            failures.push(format!(
                "{application_name}: activation HRESULT 0x{:08X}",
                result as u32
            ));
            continue;
        }
        let wrap_started = Instant::now();
        let dispatch = unsafe { Dispatch::from_raw(object)? };
        dispatch_wrap_ms += wrap_started.elapsed().as_millis() as u64;
        return Ok((
            dispatch,
            (*application_name).to_owned(),
            serde_json::json!({
                "progid_lookup_ms": progid_lookup_ms,
                "cocreate_instance_ms": cocreate_instance_ms,
                "dispatch_wrap_ms": dispatch_wrap_ms,
            }),
            activation_started_unix_ms,
        ));
    }
    bail!("未能启动 Excel/WPS COM：{}", failures.join("；"))
}

fn create_isolated_application(
    application_names: &[&str],
) -> Result<(ComApartment, Dispatch, String, u32, serde_json::Value)> {
    let total_started = Instant::now();
    let com_initialize_started = Instant::now();
    let apartment = ComApartment::initialize()?;
    let com_initialize_ms = com_initialize_started.elapsed().as_millis() as u64;
    let snapshot_started = Instant::now();
    let processes_before = process_names().context("无法确认启动前的 Office 进程清单")?;
    let process_snapshot_ms = snapshot_started.elapsed().as_millis() as u64;
    let (app, application_name, mut activation_timings, activation_started_unix_ms) =
        create_application(application_names)?;
    let hwnd_started = Instant::now();
    let hwnd_lookup = app.get_i32("Hwnd");
    let process_id = match hwnd_lookup {
        Ok(hwnd) => {
            let mut process_id = 0_u32;
            unsafe {
                GetWindowThreadProcessId(hwnd as u32 as usize as *mut c_void, &mut process_id)
            };
            process_id
        }
        // Word/WPS Writer 没有 Hwnd 属性（实测 HRESULT 0x80020006）：
        // 退化为启动前后进程快照差集 + 可执行名匹配（Rust-Tools e2973e6）。
        Err(_) => {
            let exe_names: &[&str] = if application_names
                .iter()
                .any(|name| name.eq_ignore_ascii_case("Word.Application"))
            {
                &["winword.exe", "wps.exe"]
            } else {
                &["excel.exe", "et.exe"]
            };
            owned_process_by_name_diff(&processes_before, exe_names)
                .context("Word/WPS 实例进程归属确认失败")?
        }
    };
    let hwnd_process_lookup_ms = hwnd_started.elapsed().as_millis() as u64;
    ensure!(process_id != 0, "无法确认 Office COM 实例所属进程");
    let verification_started = Instant::now();
    let process_created = process_creation_ticks(process_id);
    ensure!(
        !processes_before.contains_key(&process_id),
        "Office COM 连接到了已存在的进程；为避免影响用户打开的 Office，已拒绝复用"
    );
    let process_verification_ms = verification_started.elapsed().as_millis() as u64;
    let activation_ms = activation_timings["cocreate_instance_ms"]
        .as_u64()
        .unwrap_or_default();
    let launch_offset_ms = process_created
        .and_then(filetime_to_unix_ms)
        .map(|created_ms| created_ms - activation_started_unix_ms as i128);
    let after_launch_ms = launch_offset_ms
        .filter(|offset| *offset >= 0 && *offset <= i128::from(activation_ms))
        .map(|offset| i128::from(activation_ms) - offset);
    activation_timings["excel_process_launch_after_activation_start_ms_estimate"] =
        serde_json::json!(launch_offset_ms);
    activation_timings["excel_com_ready_after_process_creation_ms_estimate"] =
        serde_json::json!(after_launch_ms);
    let startup_timings = serde_json::json!({
        "com_initialize_ms": com_initialize_ms,
        "pre_activation_process_snapshot_ms": process_snapshot_ms,
        "com_activation": activation_timings,
        "excel_hwnd_pid_lookup_ms": hwnd_process_lookup_ms,
        "excel_process_ownership_verification_ms": process_verification_ms,
        "create_before_configuration_ms": total_started.elapsed().as_millis() as u64,
    });
    Ok((
        apartment,
        app,
        application_name,
        process_id,
        startup_timings,
    ))
}

fn set_timed(app: &Dispatch, property: &str, value: OwnedVariant) -> Result<u64> {
    let started = Instant::now();
    app.set(property, value)?;
    Ok(started.elapsed().as_millis() as u64)
}

fn system_time_unix_ms() -> u128 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}

fn filetime_to_unix_ms(ticks: u64) -> Option<i128> {
    const WINDOWS_EPOCH_OFFSET_TICKS: u64 = 116_444_736_000_000_000;
    ticks
        .checked_sub(WINDOWS_EPOCH_OFFSET_TICKS)
        .map(|unix_ticks| i128::from(unix_ticks / 10_000))
}

fn ensure_hresult(result: i32, action: impl Into<String>) -> Result<()> {
    ensure!(
        result >= 0,
        "{}失败（HRESULT 0x{:08X}）",
        action.into(),
        result as u32
    );
    Ok(())
}

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

/// 进程快照：PID → 小写可执行名（winword.exe / excel.exe 等）。
fn process_names() -> Option<std::collections::HashMap<u32, String>> {
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot.is_null() || snapshot as isize == -1 {
        return None;
    }
    let mut result = std::collections::HashMap::new();
    let mut entry = PROCESSENTRY32W::default();
    entry.dwSize = mem::size_of::<PROCESSENTRY32W>() as u32;
    let mut found = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while found {
        let name = String::from_utf16_lossy(&entry.szExeFile)
            .trim_end_matches(char::from(0))
            .to_ascii_lowercase();
        result.insert(entry.th32ProcessID, name);
        found = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    unsafe { CloseHandle(snapshot) };
    Some(result)
}

fn process_ids() -> Option<std::collections::HashSet<u32>> {
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot.is_null() || snapshot as isize == -1 {
        return None;
    }
    let mut result = std::collections::HashSet::new();
    let mut entry = PROCESSENTRY32W::default();
    entry.dwSize = mem::size_of::<PROCESSENTRY32W>() as u32;
    let mut found = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while found {
        result.insert(entry.th32ProcessID);
        found = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    unsafe { CloseHandle(snapshot) };
    Some(result)
}

/// 在启动后快照里确认本任务新创建的 Office 进程：可执行名匹配且不在启动
/// 前快照中；多于一个候选时无法唯一归属，保守报错（移植自
/// Rust-Tools e2973e6 的进程探针思路，适配 Toolhelp 快照）。
fn owned_process_by_name_diff(
    before: &std::collections::HashMap<u32, String>,
    exe_names: &[&str],
) -> Result<u32> {
    let after = process_names().context("无法确认启动后的 Office 进程清单")?;
    let mut candidates: Vec<u32> = after
        .iter()
        .filter(|(process_id, name)| {
            !before.contains_key(process_id)
                && exe_names.iter().any(|exe| name == exe)
        })
        .map(|(process_id, _)| *process_id)
        .collect();
    candidates.sort_unstable();
    ensure!(
        candidates.len() == 1,
        "无法唯一确认本任务创建的 Office 进程（候选 {} 个）",
        candidates.len()
    );
    Ok(candidates[0])
}

fn creation_ticks_for_handle(handle: *mut c_void) -> Option<u64> {
    let mut created = FILETIME::default();
    let mut exited = FILETIME::default();
    let mut kernel = FILETIME::default();
    let mut user = FILETIME::default();
    let success =
        unsafe { GetProcessTimes(handle, &mut created, &mut exited, &mut kernel, &mut user) } != 0;
    success.then_some((u64::from(created.dwHighDateTime) << 32) | u64::from(created.dwLowDateTime))
}

fn process_creation_ticks(process_id: u32) -> Option<u64> {
    let handle = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, process_id) };
    if handle.is_null() {
        return None;
    }
    let result = creation_ticks_for_handle(handle);
    unsafe { CloseHandle(handle) };
    result
}

/// Keep a verified handle to this exact process instance so timeout cleanup
/// cannot accidentally terminate a later process that reused the same PID.
pub fn open_verified_process(process_id: u32, expected_creation: u64) -> Option<usize> {
    let handle = unsafe {
        OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_TERMINATE | PROCESS_SYNCHRONIZE_ACCESS,
            0,
            process_id,
        )
    };
    if handle.is_null() {
        return None;
    }
    if creation_ticks_for_handle(handle) != Some(expected_creation) {
        unsafe { CloseHandle(handle) };
        return None;
    }
    if unsafe { WaitForSingleObject(handle, 0) } == WAIT_OBJECT_0 {
        // OpenProcess may still succeed for a signaled process object while
        // another component holds a handle. It is no longer a live process
        // and should not be handed to session cleanup as an active instance.
        unsafe { CloseHandle(handle) };
        return None;
    }
    Some(handle as usize)
}

/// Wait for the exact process represented by a verified handle to exit.
/// `true` means the process object was signaled, not merely that its PID was
/// reused or could no longer be queried.
pub fn wait_verified_process_exit(handle: usize, timeout: Duration) -> bool {
    let timeout_ms = timeout.as_millis().min(u128::from(u32::MAX)) as u32;
    unsafe { WaitForSingleObject(handle as *mut c_void, timeout_ms) == WAIT_OBJECT_0 }
}

pub fn close_verified_process(handle: usize, terminate: bool) -> bool {
    let handle = handle as *mut c_void;
    let exit_confirmed = if terminate {
        unsafe { TerminateProcess(handle, 1) };
        unsafe { WaitForSingleObject(handle, 5_000) == WAIT_OBJECT_0 }
    } else {
        unsafe { WaitForSingleObject(handle, 0) == WAIT_OBJECT_0 }
    };
    unsafe { CloseHandle(handle) };
    exit_confirmed
}

fn write_stage(path: &Path, stage: &str) {
    let _ = std::fs::write(path, stage);
}
