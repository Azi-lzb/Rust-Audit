//! Native Windows Office COM automation used by Audit and the optional converter.
//!
//! The interface is deliberately small: one private Excel/WPS application,
//! sequential workbooks, full calculation, verified save, and deterministic
//! close.  It does not attach to the user's active Office session.

use anyhow::{bail, ensure, Context, Result};
use std::{
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
            CloseHandle, SysAllocString, SysStringLen, FILETIME, VARIANT_BOOL, VARIANT_FALSE,
            VARIANT_TRUE,
        },
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize,
                CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS,
            },
            Diagnostics::ToolHelp::{
                CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W,
                TH32CS_SNAPPROCESS,
            },
            Ole::DISPID_PROPERTYPUT,
            Threading::{
                GetProcessTimes, OpenProcess, TerminateProcess, PROCESS_QUERY_LIMITED_INFORMATION,
                PROCESS_TERMINATE,
            },
            Variant::{VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I2, VT_I4},
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
        exception: *mut c_void,
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
                ptr::null_mut(),
                &mut argument_error,
            )
        };
        ensure_hresult(result, format!("Excel COM call {name}"))?;
        Ok(value)
    }

    fn invoke_method(&self, name: &str, arguments: &[OwnedVariant]) -> Result<OwnedVariant> {
        // Excel 的部分成员（Worksheets.Item、Workbook.Calculate 等）在
        // IDispatch 下以 METHOD|PROPERTYGET 组合暴露；VB 自动化客户端同样
        // 传组合标志。单 METHOD 会在这些成员上返回 0x80020003/0x80020006。
        self.invoke(name, arguments, DISPATCH_METHOD | DISPATCH_PROPERTYGET)
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
}

impl ExcelComSession {
    pub fn create(application_names: &[&str]) -> Result<Self> {
        let (apartment, app, application_name, process_id) =
            create_isolated_application(application_names)?;

        let mut session = Self {
            app: Some(app),
            _apartment: apartment,
            application_name,
            process_id: Some(process_id),
            owns_process: true,
        };
        session.configure()?;
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

    /// Convert a spreadsheet through the same private native COM instance.
    pub fn convert_to_xlsx(&mut self, input: &Path, output: &Path) -> Result<()> {
        ensure!(
            input.is_file() && !output.exists(),
            "格式转换输入不存在或输出已存在"
        );
        let app = self.app.as_ref().expect("COM app initialized");
        let workbook = app
            .get_dispatch("Workbooks")?
            .invoke_method(
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

    pub fn recalculate(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
        stage_path: &Path,
    ) -> Result<serde_json::Value> {
        self.recalculate_with_mode(input, output, timeout, stage_path, "full")
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
    ) -> Result<serde_json::Value> {
        write_stage(stage_path, "opening_workbook");
        let result = self.recalculate_inner(input, output, timeout, stage_path, mode);
        if result.is_err() {
            write_stage(stage_path, "closing_failed_workbook");
        }
        result.map(|(stages,)| stages)
    }

    fn configure(&mut self) -> Result<()> {
        let app = self.app.as_ref().expect("COM app initialized");
        app.set("Visible", OwnedVariant::boolean(false))?;
        app.set("DisplayAlerts", OwnedVariant::boolean(false))?;
        // 与旧 Python 生产通道一致的 Excel 端开销配置：关闭屏幕刷新与事件、
        // AskToUpdateLinks 关闭（UpdateLinks 已在 Open 参数显式控制）、
        // 计算模式设为手动以消除 Workbooks.Open 对脏公式的隐式自动重算
        // （审核副本全部公式为待算状态；显式 Calculate 随后一次性同步完成，
        // 旧 Python 生产语义与此完全一致，Calculation 为应用级属性不随
        // SaveAs 写入输出文件）。
        app.set("ScreenUpdating", OwnedVariant::boolean(false))?;
        app.set("EnableEvents", OwnedVariant::boolean(false))?;
        app.set("AskToUpdateLinks", OwnedVariant::boolean(false))?;
        // 某些含数据模型/OLAP 连接的 Excel 实例禁止通过 COM 修改
        // Application.Calculation（0x800A03EC）；与旧 Python 一致降级忽略，
        // 显式 Calculate 仍完成重算。
        if let Err(error) = app.set("Calculation", OwnedVariant::i32(XL_CALCULATION_MANUAL)) {
            let _ = error;
        }
        Ok(())
    }

    fn recalculate_inner(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
        stage_path: &Path,
        mode: &str,
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
            .invoke_method(
                "Open",
                &[
                    OwnedVariant::string(&input_text)?,
                    OwnedVariant::i32(0),
                    OwnedVariant::boolean(false),
                ],
            )?
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
                    .with_context(|| format!("Excel COM Worksheets.Item({sheet_index}) 未返回工作表"))?;
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
            let output_text = output.to_string_lossy();
            workbook.invoke_method(
                "SaveAs",
                &[OwnedVariant::string(&output_text)?, OwnedVariant::i32(51)],
            )?;
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

/// A private Word/WPS COM instance used for document conversion and checks.
pub struct WordComSession {
    app: Option<Dispatch>,
    _apartment: ComApartment,
}

impl WordComSession {
    pub fn create(application_names: &[&str]) -> Result<Self> {
        let (apartment, app, _, _) = create_isolated_application(application_names)?;
        let session = Self {
            app: Some(app),
            _apartment: apartment,
        };
        let app = session.app.as_ref().expect("COM app initialized");
        app.set("Visible", OwnedVariant::boolean(false))?;
        app.set("DisplayAlerts", OwnedVariant::i32(0))?;
        Ok(session)
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
        if let Some(app) = self.app.take() {
            let _ = app.invoke_method("Quit", &[]);
            drop(app);
        }
    }
}

struct ComApartment;

impl ComApartment {
    fn initialize() -> Result<Self> {
        let result = unsafe {
            CoInitializeEx(
                ptr::null(),
                (COINIT_APARTMENTTHREADED | 4) as u32, // COINIT_DISABLE_OLE1DDE
            )
        };
        ensure_hresult(result, "initialize the Excel COM apartment")?;
        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

fn create_application(application_names: &[&str]) -> Result<(Dispatch, String)> {
    let mut failures = Vec::new();
    for application_name in application_names {
        let mut prog_id = wide(application_name);
        let mut class_id = GUID::default();
        let result = unsafe { CLSIDFromProgID(prog_id.as_mut_ptr(), &mut class_id) };
        if result < 0 {
            failures.push(format!(
                "{application_name}: ProgID HRESULT 0x{:08X}",
                result as u32
            ));
            continue;
        }
        let mut object = ptr::null_mut();
        let result = unsafe {
            CoCreateInstance(
                &class_id,
                ptr::null_mut(),
                CLSCTX_LOCAL_SERVER,
                &IID_IDISPATCH,
                &mut object,
            )
        };
        if result < 0 {
            failures.push(format!(
                "{application_name}: activation HRESULT 0x{:08X}",
                result as u32
            ));
            continue;
        }
        return Ok((
            unsafe { Dispatch::from_raw(object)? },
            (*application_name).to_owned(),
        ));
    }
    bail!("未能启动 Excel/WPS COM：{}", failures.join("；"))
}

fn create_isolated_application(
    application_names: &[&str],
) -> Result<(ComApartment, Dispatch, String, u32)> {
    let apartment = ComApartment::initialize()?;
    let processes_before = process_ids().context("无法确认启动前的 Office 进程清单")?;
    let (app, application_name) = create_application(application_names)?;
    let hwnd = app.get_i32("Hwnd").context("读取 Office COM 窗口句柄")?;
    let mut process_id = 0_u32;
    unsafe { GetWindowThreadProcessId(hwnd as u32 as usize as *mut c_void, &mut process_id) };
    ensure!(process_id != 0, "无法确认 Office COM 实例所属进程");
    ensure!(
        !processes_before.contains(&process_id),
        "Office COM 连接到了已存在的进程；为避免影响用户打开的 Office，已拒绝复用"
    );
    Ok((apartment, app, application_name, process_id))
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
            PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_TERMINATE,
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
    Some(handle as usize)
}

pub fn close_verified_process(handle: usize, terminate: bool) {
    let handle = handle as *mut c_void;
    if terminate {
        unsafe { TerminateProcess(handle, 1) };
    }
    unsafe { CloseHandle(handle) };
}

fn write_stage(path: &Path, stage: &str) {
    let _ = std::fs::write(path, stage);
}
