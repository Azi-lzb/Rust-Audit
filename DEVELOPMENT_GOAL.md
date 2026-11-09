# Rust-Audit 独立开发目标与状态

**数值尾差不再硬失败：容差 1e-15 → 1e-14 + 数值漂移降级为警告（2026-11-09，用户确认）**

现场反馈：openpyxl 编辑读回时「计算前后非公式数据不一致」导致整次失败，
给的样本 `0.0012083234879418526 → 0.00120832348794185` 实测相对差 2.15e-15
——**恰好等于 Excel 15 位有效数字截断**（`float(f"{a:.15g}") == b`），属合理差距。

- **容差统一**：`excel-editor` 新增 `NUMERIC_VALUE_TOLERANCE = 1e-14`，替换
  openpyxl 读回校验三处 1e-15（planned copy、`validate_value`、`json_value_equal`）。
  依据：同类实测最坏 2.76e-15，1e-14 有约 3.6 倍余量，仍比 Excel 15 位精度
  固有误差紧一个数量级，1 元 / 1% 级真实差异照常拦截。
- **比较口径修正**（`numeric_values_equal`）：补零带绝对差（原实现要求两侧
  非零，`0` 与极小正数会被误判）、整数必须精确相等（大数值下相对容差会
  放过 ±1）、非有限值只接受完全相等。
- **超容差数值漂移降级为警告**：`validate_preserved_values` 遇到「未修改格
  两侧同为数值且超容差」时记警告并放行（打印坐标 + 两侧数值 + 相对差），
  不再 fail-closed。**结构性变化仍硬拒**：单元格增删、工作表名/序变、
  公式文本变、类型变化、文本/布尔/错误值变化一律 `bail!`。
  警告按线程隔离收集（`NUMERIC_DRIFT_WARNINGS`），每轮校验开始时清空。
- **S1-F02 写回自校验同步放宽**：`s1_summary_merge::floats_equal` 与
  `s1_region_summary::values_equal` 的 1e-15 一并提到 1e-14 并补零带/整数保护
  （同一类 Excel 15 位尾差风险点）。
- 未改动 Rust-Tools（用户明确不需要改那边）；Rust-Audit 不做 xls→xlsx 转换校验。

回归：excel-editor lib **80/80**（新增 3 项容差/漂移测试）、audit-core lib **289/289**；
`cargo fmt --check` 干净。

**产品改名 AuditTools + 版本号体系 + 旧 audit 常见问题迁入（2026-11-09，用户要求）**

用户要求：英文名改为 `AuditTools`、中文名保持「审核工具」、按旧 audit 逻辑编写版本号、
把旧 audit 的常见问题搬进来，最后打 exe/deb 共 4 个包。

- **版本清单** `VERSION_MANIFEST.json`（仓库根，仿旧 audit `版本管理/VERSION_MANIFEST.json`）：
  `product.id=AuditTools`、`product.name=审核工具`、`product.version=26.3.0.0`；
  格式 `YY.MAJOR.MINOR.PATCH`（年份两位、各段 ≤65535），与旧 audit 的
  `app_identity.py` 校验口径一致。Rust 侧 `app_identity` 模块读取安装根清单，
  缺失时回退编译期内嵌的同一份（单测强制常量与清单一致）。
- **界面标题**：`__AUDIT_APP_TITLE__` 由常量「审核工具」改为
  `app_identity::app_title()` → **「审核工具 V26.3.0.0」**；`appInfo.title`、
  浏览器标题、关于对话框、系统信息摘要同步。
- **英文名**：CLI 名 `AuditTools`、Cargo bin 名 `AuditTools`、发行 exe `AuditTools.exe`、
  deb 包名 `audittools`、安装路径 `/opt/audittools/`、命令 `/usr/bin/audittools`、
  资源目录 `libexec/audittools/openpyxl/`（`PRODUCT_DIR`，满足 worker 解析器的
  小写/数字/连字符约束）、状态目录 `%LOCALAPPDATA%\audittools` 与
  `~/.local/state/audittools`。`run.bat`/`build-release.bat`/`package.bat`/
  `package-windows.ps1`/`package-linux-deb.sh` 全部同步；打包时把
  `VERSION_MANIFEST.json` 放进发行包根（配置更新模块也读它）。
- **常见问题**：旧 audit `docs/常见问题.json` 的 23 条问答原样迁入
  `audit-core/src/faq_data.json`（逐字段校验与旧文件完全一致），新增
  `get_faq` API；S1S2 发行版取 `answerS1S2`（无则回退完整答案），全量版取
  `answer`。前端 FAQ 页签（分类筛选/搜索/Q 编号）此前已就绪，本次接上真实数据。

回归：audit-core lib **289/289**（新增 app_identity 3 项、faq 2 项）；`cargo fmt --check` 干净。

**config-real 默认绑定播种进 settings.json（2026-11-09，用户反馈「配置还是没有绑定上去」）**

现象：S3 四个配置框（通用/执行比较/本期数值核对/指标比较拆分）显示「未选择」。
原因：启动默认依赖 `resolve_default_config_path` 找 config-real，用户直接双击
`dist\open-data-audit.exe` 时 CWD=dist、exe 同级无 config-real → 解析失败；
且默认值只存在于内存状态，settings.json 里没有。

- **默认绑定表**（`app_settings::BUNDLED_CONFIG_DEFAULTS`，键 → config-real 相对路径）：
  templateDir → `模板文件_命名区域`（目录）；external（S1 大集中系统数据）与
  pcCentral（S2 大集中系统数据）→ `外部文件-大集中相关数据.xlsx`；
  historyConfig → `1.逐笔统计系统_配置.xlsx`；pcConfig → `2.报表采集系统_配置.xlsx`；
  centralCommonConfig/centralComparisonConfig/centralCrossConfig/centralFormsConfig →
  `3.0/3.1/3.2/3.3` 四份配置。
- **播种**：serve 启动即调 `seed_bundled_config_defaults()`——仅当字段缺失/为空/
  失效时填充并**写盘 settings.json**；用户手动选择的有效路径永不覆盖。
  S1S2 包没有 3.x 配置时对应字段留空。
- **config-real 目录解析**（`config_real_directory`）：`AUDIT_CONFIG_REAL_DIR` 显式
  覆盖 → `<exe>/config-real`（发行包）→ `<exe>/../audit-core/config-real` 与
  `<exe>/../../audit-core/config-real`（开发目录双击 dist exe）→ CWD 两个候选；
  目录须含 1./2./3.0 标记文件之一。`resolve_default_config_path` 同步补 exe 上级候选。
- **打包**：发行包 config-real 新增「外部文件-大集中相关数据.xlsx」与
  「模板文件_命名区域/」文件夹（全量与 S1S2 都带，S1S2 仍无 3.x）；
  package-windows.ps1 与 package-linux-deb.sh 同步。

回归：audit-core lib **284/284**（新增播种单测：填缺失、换失效、不动用户选择）。

## 2026-11-07 对齐基线（da7b1c3）八功能真实数据对比：S1F1 逐条一致、S3F4 逐格一致、S3F1 收敛至 15 处

- **对齐基线**：3.open-data-audit@da7b1c3 新增 rustaudit_alignment 基线 JSON（8 功能真实数据生产入口跑出）；驱动 `tests/compare/rustaudit_alignment_p1.py`（每功能两跑、367 输入 SHA-256 前后核验）。引擎口径：calc=auto(Excel COM)、S1 条件格式提取来源=source、空白格纳入、centralTolerance=100 元。
- **S1F1 ✅ 完全对齐**：129/129 结构通过；公式 3345 + 条件格式 257 = 3602 条与基线逐条一致（Rust-only 0、Baseline-only 0、字段值差 0），两次运行一致。关键修复链：excel-open-repair 第三类缺陷——模板校验区域公式残留 `[1]表名!` 悬空外部引用（模板自带 12 处 [1] 外链，外部表导入为本地 sheet 后链接悬空），Excel COM 打开即拒绝（0x80020009）、WPS 忽略断链但重算丢公式（W152）→ 剥离 `[N]本地表!` 前缀后债券 7 本恢复真实重算、13 条 CF 归位。该修复已固化进 repair 模块（第三类缺陷，非本地表外链不动），RustTools 等自研 OOXML 副本管线同样需要此归一。
- **S3F1 ✅ 逐格 0 差异**（原收敛 15 → 7 → 0）：三项修复——排序三级键不区分大小写（VBA EndCompare/casefold）；表达式语法口径 legacy-eight-segment（源生产 run_comparison 默认 LEGACY_8）；「跨期核对」组规则从五段表借机构/地区 scope、机构列为空 → scope_missing 整条跳过。行数 3868 一致、两次运行一致。收口链：① 比较运算 Decimal 精确路径（_decimal_comparison_value 口径，消除 IEEE 尾差误判 <>）→ 15→7；② 累计指标 `_accu_active` 门控补齐（同年且上期非 1 月 1 日时只按 change<0 报通用文案，不检查负值/不走日期分支）→ 7→**0**。两跑逐格 0 差异、367 输入 SHA-256 不变、填充色分布与基线一致。
- **S3F4 ✅ 逐格 0 差异**（33 行核对结果，两跑一致）。S3F2/F3 ✅（0 行/0 回写为该期真实业务结果）。S1F2/S2F1 ✅（首轮口径，505 行/1087 条）。S3F5 ❌ 未对齐：三个后端均不支持 DeleteColumn（Umya 有删列逻辑但定义名称守卫拒绝；ooxml import 后端显式 UnsupportedBeforeWrite；openpyxl worker 未接线）。路由决策（用户确认）：**S3-F05 修改管线 umya 退出；ooxml DeleteColumn 优先**（已实现并验证：ooxml import 后端原生支持 DeleteColumn；8 份表单两跑逐格 0 差异；ZIP 级直改、格式零损失）；**openpyxl worker 备选**（openpyxl 原生 delete_cols() 与旧 audit render_mode=OPENPYXL 同款语义）。**实现进展（2026-11-07）**：S3-F05 已切到 openpyxl 优先路由（umya 退出，见 `s3_f05_modifier`；`production_modifier` 与 `production_modifier_with_edit_engine` 均已接线）；worker 新增 delete_column / set_cell_number_format / set_row_hidden / set_sheet_visibility 四个操作（含白名单与读回校验），数值读回改为相对误差 1e-14 容差（Excel 15 位有效数字序列化）。**S3F5 ✅ 逐格 0 差异达成（2026-11-07 收口）**。实现路径：路由切到 openpyxl 优先（umya 退出 S3-F05；`s3_f05_modifier`，两条入口均接线）；worker 新增 delete_column（openpyxl `delete_cols` 与旧 audit OPENPYXL 渲染同款）、set_cell_number_format、set_row_hidden、set_sheet_visibility 四个操作（含白名单与逐操作读回）；数值读回改相对误差 1e-14（Excel 15 位有效数字序列化）；PlanRowMap 增加删列位移映射；被删工作表的 DefinedNames/RowAndColumnDimensions 显式登记 drop allowance。
**样式保真口径（业务确认 B）**：S3-F05 的 `s3_f5_required_preservation` 收窄为业务关键项（单元格值/类型、公式、条件格式、数据验证、合并、定义名称、工作表身份与顺序、批注），不再把 Styles / 行高列宽 / 页面设置 / 页眉页脚 / Relationships / 冻结窗格 / AutoFilter 列为硬门槛——openpyxl 重写必然规范化这些渲染元数据且不影响表单业务含义。后端校验器本身未放宽（excel-editor 64/64 全部通过，真实数据丢失仍被拦截）。
**验收**：8 份机构表单文件名与基线逐一对应、逐格内容 0 差异；全量 lib 273/273。



**openpyxl 路由换架构压缩 37.3s → 4.6s（2026-11-08 第五轮，用户确认换架构）**：逐段计时（AUDIT_S3F5_TIMING）定位 apply=19-28s / validate=6.8-8.2s。两处纯收益优化：①验证器输入侧缓存（同批次 8 表单 plan.input 字节相同，`(path,size,mtime)` 复用值转储/保真快照/差分样式，validate 9-11s→6.8-8.2s）；②worker 惰性加载（`data_only=True` 整簿二次加载只有 copy_range 需要，改为按需，S3-F05 省每表 ~10s）。但墙钟仍 37s——12 核跑 8 个 Python 进程 + 8 个 Rust 验证线程已近满载，且开销在数千次逐格 Python 赋值。实测证明瓶颈在逐格操作而非 openpyxl 本身（**整簿往返仅 0.75-0.88s/份**）。按用户决策换架构：新增 `fast_render` worker 协议（`mode="fast_roundtrip"`）——Rust 侧完成 FAST 回填，worker 只做「openpyxl 加载→保存」往返，包由 openpyxl 自己的序列化器重建（保持与 FAST 直改字节的实现独立性），并保留样式表完整性校验与源哈希不变校验。路由：`edit_engine=openpyxl` → `FastRenderMode::OpenpyxlRoundtrip`；auto/ooxml → `Direct`。**结果：openpyxl 路由 37.3s → 4.6-5.0s（8 表全批次，单表 2.3s，降 8×），快于旧 audit OPENPYXL 11.2s 的 2.2×；FAST 直出维持 1.7-2.0s；三方逐格仍 0 差异、输入 SHA-256 不变；76/76 + 278/278。**










**高级设置文案纠错（2026-11-09）**

用户指出「条件格式判定方式」标着 PYTHON 容易误解——实际是**内置 Rust 求值器**
（`conditional_format.rs::evaluate_expression`），PYTHON 只是沿用旧 Python 实现的
语义口径名。全量审计高级设置后修正 6 处：

| 位置 | 原文案问题 | 修正 |
| --- | --- | --- |
| 条件格式判定方式（说明） | 「PYTHON=直读 OOXML 规则求值」暗示用 Python | 「规则求值（默认）=由内置 Rust 求值器…与旧 Python 实现口径一致」 |
| 条件格式判定方式（下拉） | 选项直接显示 `PYTHON`/`COM_DISPLAY` | 显示「规则求值（内置 Rust）」/「COM_DISPLAY（Excel/WPS 显示色）」/「COM_EVALUATE（未实现）」；内部枚举值不变（兼容已存设置） |
| 公式重算（只读行） | 写「原生 COM / LibreOffice（按设置）」，漏掉默认 betteroffice | 「betteroffice（内置）/ COM / LibreOffice（按设置）」 |
| 条件格式（只读行） | 「OOXML 规则判断（内置）」 | 「Rust 规则求值（内置）」 |
| 生成文件方式（S3-F05） | 写「由 Excel 文件修改引擎填充」，实际是自研 FAST 渲染器 | 「由内置 OOXML 渲染器按模板逐机构生成…ZIP 级直改、机构级并行」 |
| S1-F01 公式重算引擎 | 只写 COM/LibreOffice，未提 betteroffice 优先 | 说明与「校验计算引擎」配合：默认 betteroffice，本下拉为回退链起点 |

运行时可读文本同步修正（`COM_EVALUATE` 未实现的报错与设置更新日志），
内部枚举值 `"PYTHON"` 保留不动以免破坏已保存设置。

回归：audit-core lib 283/283。

**发行包预填配置路径（2026-11-09）**

目标：用户只需选「源数据目录」，模板与外部文件自动识别，不必在多个配置框里挑文件。

- 打包已含 `config-real/`（全量 6 份 + 外部数据；S1S2 版 2 份）。启动时
  `resolve_default_config_path` 找到 `<exe 目录>/config-real`，`AppState::new`
  据此推导其余配置的绝对路径并预填到界面（S1 historyConfig、S3 四项配置）。
- **新增重定位**：`apply_preferences_to_state` 遇到失效的保存路径时，按「同文件名」
  在 `<exe 目录>/config-real/` 下重新定位，用户整体移动发布包后无需重选。
  实测：删掉旧目录后启动 → historyConfig 自动指向新位置的 config-real。
- 实测首次启动（无 settings.json）：6 项配置全部预填到包内路径；S1S2 包里
  S3 配置为空（预期，包内无 S3 配置）。

回归：audit-core lib 283/283。

**界面功能收缩：F3/F4/F5 按钮、组合设置导航、金融表单目录、彩蛋（2026-11-09，用户确认）**

- **主界面 F3/F4/F5 按钮删除**：组合工作表（核查表）/ 组合工作表（按配置）/
  制作联合模板三个按钮及 `dagExtrasCard` 卡片从主界面移除；设置中心「组合设置」
  导航页签删除（导航只剩 常规设置 / 操作指引 / 帮助与更新 / 高级设置）。
  `start_dag_flow` 后端 API 保留——S1-F01/F02 主按钮仍走该路由。
- **金融表单输出目录字段删除**：S3-F05 输出目录与 S3 输出目录合并为同一个
  `centralOutput`（用户确认「金融表单目录就和输出结果是同一个」）；
  `centralFormOutput*` 状态字段、联动逻辑与「最近输出目录：」回显行全部移除。
- **五连击彩蛋删除**：低频功能入口已不存在，`auditExtrasUnlocked` /
  `applyEggState` / `dag-extras-locked` CSS / `uiAuditExtrasUnlocked` 存储键、
  操作指引中 `egg:true` 的 S1F5 条目一并删除；`createTemplateMerge` /
  `doCreateTemplateMerge`（后端路由早已删除的死代码）同步清理。
- **帮助与更新精简**：保留「常见问题」与「安装 LibreOffice」两个子页
  （betteroffice 内置后 LO 仍是 Linux 主回退，安装入口保留）。
- **残留文案清理**：重算失败日志不再引用已删除的 S1-F3 入口，改为完整指引
  「用 Excel 打开审核副本 → Ctrl+Alt+F9 强制重算 → Ctrl+S 保存 → 在 Excel 中
  查看校验区域结果」；恢复发布路径的内部上下文改称 S1-F01（该路径现在是
  S1-F01 失败保留流程的一部分）；测试符号 `write_s1_f3_recovery_fixture` →
  `write_s1_recovery_fixture` 等同步重命名。前端 index.html 内联 JS 全量
  语法解析通过；`cargo fmt --check` 干净。

回归：audit-core lib **283/283**。

**S1-F3 恢复功能整体删除（2026-11-09，用户确认）**

多重兜底（betteroffice → COM/WPS → LibreOffice）下重算失败已属罕见，S1-F3
「从审核副本汇总校验核查表」入口不再需要，按用户要求**整体删除**（不只是隐藏）：

- 前端：删除恢复按钮/警告/面板、`renderS1F3RecoveryEntry` / `toggleS1F3RecoveryPanel`
  / `startS1F3Recovery` 三个函数、CSS 规则、引导条目、标题映射，以及「组合设置」
  页的 S1-F3/F4 说明与 profile DAG 工具行
- 后端：删除 `start_s1_f3_workflow` / `list_s1_f3_tasks` 两个 API 路由及其实现、
  `run_s1_f3_worker` 系列、`s1_f3_publish_stage`、`apply_s1_f3_recovery_state`，
  以及 `s1F3Recovery*` / `s1F3LostCopy*` 状态字段（`get_state` 不再每次扫描磁盘）
- 测试：删除 5 个 S1-F3 专属测试；`write_s1_f3_recovery_fixture` 保留（S1-F01 测试仍在用）
- **保留**：S1-F01 重算失败时的「审核副本已保留，等待人工重算」日志提示与副本
  （必要的失败兜底，用户可自行在 Excel 中重算；只是不再有界面入口回收）

回归：audit-core lib **283/283**（-5 为删除的 S1-F3 测试）。

**未验证项全部补齐（2026-11-09）**

| 原未验证项 | 结果 | 证据 |
| --- | --- | --- |
| Linux deb 打包（glibc 2.28 / 麒麟 V10 / UOS V20） | ✅ **通过** | zigbuild target `x86_64-unknown-linux-gnu.2.28`；`objdump` 实测最高符号 GLIBC_2.28；新增 `package-linux-deb.sh`（含 glibc 上限硬校验）；WSL 实测 打包→安装→运行→卸载 全通，deb 9.4MB |
| 循环引用拒绝分支 | ✅ 拒绝 | `存在 2 个循环引用格` |
| #NAME? 拒绝分支 | ✅ 拒绝 | `1 个公式格求值为 #NAME?（betteroffice 未实现的函数）`；触发路径是**未定义名称**（GEOMEAN/UNIQUE 已被实现或静默降级） |
| 真实带 VBA 的 xlsm | ✅ **通过** | `统计工具-大集中工具v3.3.8.xlsm`（38 表 / 3106 公式 / `vbaProject.bin` 713728 B）：重算成功、无 #NAME?、**VBA 部件逐字节保留** |
| 3 轮以上收敛 | ✅ 已收敛 | ！100 / ！90 / ！50 三模板：round1 有变化 → round2 changed=0 → round3/4/5 恒 0 |
| WPS COM 引擎 | ✅ **核心能力已验**（异常转移路径未构造样本） | WPS COM（ket.Application v12.0）实测可打开、CalculateFull、SaveAs，产物 46 部件 / 551 公式可读。**两种接入方式**：①未安装 Excel 时 `ExcelComSession::create` 的候选列表 `[Excel.Application, ket.Application, KET.Application]` 按序跳过 Excel，**WPS 直接作为首选 COM 引擎**（不是"回退"）；②装了 Excel 但 `Workbooks.Open` 失败时，`should_retry_with_wps` 关闭 Excel 会话、切 WPS 重试。②这条异常转移路径未构造样本（repair 已覆盖已知全部 Excel 拒开原因，真损坏则 WPS 同样打不开） |

**Windows 重算引擎完整优先级**（代码实证，`native_recalculation.rs:31` +
`s1_recalculation.rs:692`）：

```
betteroffice（纯 Rust，默认）
  ↓ 拒绝（预算耗尽 / #NAME? / 循环引用 / 打不开）
Excel.Application（若已安装；未安装则候选列表直接跳过）
  ↓ 已安装但 Workbooks.Open 失败 → should_retry_with_wps 显式切 WPS
ket.Application / KET.Application（WPS）
  ↓ 也失败
LibreOffice（Windows 通常未安装）
  ↓
明确报错（不静默）
```

日常使用不依赖任何 Office 软件：20 个真实模板在 betteroffice 下单路由全部通过，
只有 betteroffice 拒绝时才需要 COM/WPS/LibreOffice 兜底。

**Linux glibc 2.28 端到端**：20 模板 betteroffice **20/20 全部成功**，与 Windows 逐字段完全相同（实质差异 0）——此前 Linux 侧 ！40/！70 失败的真因是
`recalculate_with_betteroffice` 与 sidecar session 分支仍用**未规范化的 input** 作比对基准，
shared ref 归一被误判成「重算改变公式」；修正后两平台一致。

**顺带修的错误信息缺陷**：betteroffice 拒绝时若同时缺 sidecar，最终错误只显示
「未配置 sidecar」，吞掉了真实原因（预算耗尽/#NAME?/循环引用）。现已在批量与单本
两个入口把拒绝原因拼入错误文本。

**OOXML 规范化与计算引擎解耦（2026-11-09，Rust-Tools ③ 对齐）**

Rust-Tools 指出：若只在 COM 分支做 shared formula ref 归一，betteroffice 侧副本
仍带错位 ref，两侧口径不一致 → 跨路由对拍照样假阳性。本项目实测确认此判断：
修复前 betteroffice 副本 `ref="C13:C48"`、COM 副本 `ref="C14:C48"`；虽然
公式文本守卫没报错（betteroffice 不改 `<f>`），但**两侧副本字节不同**，任何
后续跨路由比较都会受影响。

改为「与引擎无关」：`excel_open_repair` 在**任何引擎打开副本之前**执行，
三处入口全部去门控——
- `s1_formula_copy::execute_s1_formula_copy_with_recalculator`（公式复制后）
- `S1RecalculationSession::recalculate_betteroffice`（新增，此前缺失）
- `S1RecalculationSession::recalculate_native` / `recalculate_s1_formula_copy`

实测：修复后两路由副本均 `ref="C14:C48"`，完全一致；betteroffice 路由日志
打印 `shared_ref=1` 留痕。20 模板双路由仍 20/20、实质差异 0。

**重算公式文本守卫可配置（2026-11-09）**

新增高级设置项「重算公式文本守卫」（`recalcFormulaGuard`，默认开启）。
三个计算引擎（betteroffice / Excel COM / LibreOffice）现在都能正确计算，
但各自的公式文本规范化口径不同（shared formula ref、TRUE/FALSE 字面量、
表名引号），逐格比对可能产生假阳性把整簿判死。关闭该项只跳过**第一层**
（公式文本/位置比对），三层底线检查**始终生效**：
① 公式数量一致；② 缓存节点完整；③ 结果可通过 Calamine 读取。
关闭时运行日志打印明确警告，不会静默降级。

实测（！40 委托贷款，COM 路由）：守卫 on → hits=1 成功（3507ms）；
守卫 off → hits=1 成功（3340ms）——shared ref 修复后守卫本身已不误报，
开关作为未来遇到新规范化差异时的应急手段。

**未验证项补齐 + ！40 真因修复（2026-11-09）**

**！40 失败真因（与悬空外链无关）**：模板存在 `ref="C13:C48"` 而 shared formula
主格在 C14（ref 起点落在主格之前）。calamine 按 ref 起点换算从属格相对引用，
源副本上 C15 被读成 `$D16`，而 Excel 输出是 `$D15` → 重算守卫误判
「重算改变公式 C15」整簿失败。betteroffice 因字节级不动公式而绕过该误判。
修复：`excel_open_repair` 新增第四类修复（shared ref 起点归一到主格，Excel 打开
保存的既有行为，公式语义不变）+ 纳入 `needs_repair` 检测 + repair 接入
`recalculate_native` 与 `recalculate_s1_formula_copy` 两条入口。
**结果：20 模板双路由 20/20 全部成功，实质差异 0（原 19/20）**——优于
Rust-Tools 的 18/20 基线。

**② 拒绝分支实测（补充 Rust-Tools 未验证项）**：

| 样本 | 结果 |
| --- | --- |
| 循环引用（A1=B1, B1=A1） | ✅ 拒绝：`存在 2 个循环引用格` |
| #NAME?（`=不存在的函数名(1,2)`） | ✅ 拒绝：`1 个公式格求值为 #NAME?（模板使用了 betteroffice 未实现的函数）` |

注：`GEOMEAN`/`UNIQUE` 均已被 betteroffice 实现或被静默降级，不产生 #NAME?；
真正的触发路径是未定义名称。

**③ 重算保真实测**：

| 项 | 结果 |
| --- | --- |
| xlsm | ✅ 2 公式正常重算，readback 一致 |
| 条件格式 / 数据验证 / 合并 / 定义名 / 跨表引用 | ✅ 逐要素前后完全相等（cf=1 dv=1 merge=1 definedName=1，sheets 2→2 formulas 11→11） |
| strip 样式空格（内容 5 行 + 样式刷到 5000 行） | ✅ 剥离 4995 格纯样式空格；`SUM(A:A)` 两种模式结果一致（content_range 修复生效） |

**④ 多轮收敛实测（1~5 轮）**：

| 模板 | round1 | round2 | round3 | round4 | round5 |
| --- | ---: | ---: | ---: | ---: | ---: |
| ！100 单位存款 | 20 | **0** | 0 | 0 | 0 |
| ！90 同业业务 | 21 | **0** | 0 | 0 | 0 |
| ！50 票据业务 | 1 | **0** | 0 | 0 | 0 |

第 2 轮即 changed=0 收敛，3/4/5 轮保持 0——默认 2 轮的口径在本项目真实负载上成立。

**回归**：audit-core lib **286/286**（+2 shared ref 单测）。

**betteroffice GUI 高级配置 5 开关 + 配置优先级（2026-11-08 第六轮）**

- **5 个设置项**（`PREFERENCE_SPECS` 白名单 + `AppSettings` 访问器 + 前端「校验计算引擎（betteroffice）」卡片）：
  `boCalcEngine`（betteroffice 默认｜com｜off）、`boCalcMaxVisits`（0=默认 1000 万）、
  `boCalcRounds`（1~5，默认 2）、`boCalcStripStyle`（默认否）、
  `boCalcFallback`（auto 默认｜com｜libreoffice｜none）。前端 5 个 setter 走 `bridge.api.update`，
  与既有偏好同一条持久化链路。
- **配置优先级**：GUI/配置存储 > 环境变量 > 内置默认（`BoCalcConfig::from_app`）。
  逐键覆盖而非整体替换——GUI 未设置的键让 env 继续生效；「引擎=com」对回退方向有
  最终决定权（锁定 COM-only）。
- **#ERROR! 字面检测**：betteroffice 把求值错误写成 `<v>#ERROR!</v>`，OOXML 只允许
  7 种标准错误码（Excel 0x80020009 拒开）。落盘前扫描字节，命中即视为
  `BoCalcRefusal::Unusable` 回退，绝不把打不开的文件交给下游。

**GUI 闭环实测（5 项逐个改值 → 重跑 → 行为切换）**：

| 设置 | 实测引擎 | 耗时 |
| --- | --- | ---: |
| `boCalcEngine=com` | **Excel.Application**（betteroffice 被跳过） | 7141ms |
| `boCalcEngine=betteroffice` | **betteroffice (pure Rust)** | 162ms |
| `boCalcMaxVisits=10` | `Excel.Application；betteroffice 回退（重算预算耗尽（158 个公式格未求值，保留旧值不可信））` | 3554ms |
| `boCalcFallback=none` + 预算耗尽 | **明确报错**（不静默） | — |
| `boCalcRounds=1` / `boCalcStripStyle=true` | betteroffice 正常出结果 | 185 / 199ms |

**回归**：audit-core lib **284/284**（+3：`error_literal_detection`、
`settings_take_priority_over_env_and_defaults`、环境锁串行化的既有测试修正）；
20 模板双路由对拍维持 **19/20 实质差异 0**（！40 仍是 COM 侧既有
「重算改变公式」守卫，betteroffice 字节级未改公式）。

**环境变量仍是兜底**（与 Rust-Tools 同名）：`PYTOOLS_BO_CALC` / `_MAX_VISITS` /
`_ROUNDS` / `_STRIP_STYLE` / `_FALLBACK`，仅在 GUI 未显式设置对应键时生效。

**S1-F01 校验区域重算：betteroffice 优先路由（2026-11-08）**

按 Rust-Tools 已验收的同源改造（919750e/925630b/a23b2ca）移植：

- **vendor**：`vendor/betteroffice-xlsx-{model,calc}` 0.5.1 修改版（Apache-2.0，修改点见 `vendor/VENDOR_NOTES.md`）；根 `Cargo.toml` 用 `[patch.crates-io]` path 替换，版本保持 0.5.1，两目录入 `workspace.exclude`；`--offline` 构建通过，`Cargo.lock` 中两条已无 source/checksum。
- **bo_calc**：新增 `audit-core/src/bo_calc.rs`（等价 pytools-core 实现）——open → 设预算 → N 轮 `recalculate_all` → save → 统计受限/错误格；`now_serial` 显式注入（系统时钟→Excel 1900 序列号）；默认 2 轮；`strip_style_only_cells` 同步移植到 `crates/ooxml-patch`。
- **路由**：`S1RecalculationSession::recalculate` 改为 **betteroffice → COM → LibreOffice → 明确报错**；`PYTOOLS_BO_CALC_FALLBACK=none` 时拒绝静默采用受限结果；回退原因拼入报告 `engine`（如 `Excel.Application；betteroffice 回退（重算预算耗尽（1329 个公式格未求值…））`）。引擎白名单扩充为接受 betteroffice 及其回退串。
- **配置**（与 Rust-Tools 同名）：`PYTOOLS_BO_CALC`（auto/off）、`PYTOOLS_BO_CALC_MAX_VISITS`、`PYTOOLS_BO_CALC_ROUNDS`（默认 2，上限 5）、`PYTOOLS_BO_CALC_STRIP_STYLE`（默认 off）、`PYTOOLS_BO_CALC_FALLBACK`（auto/com/libreoffice/none）。

**20 模板双路由对拍**（`audit-core/examples/s1f1_bo_parity.rs`，只走「重算→扫描校验区域」以隔离引擎变量）：

| 结果 | 数量 |
| --- | ---: |
| 实质差异 0（逐字段完全相同） | **19 / 20** |
| 失败态不一致 | 1（！40 委托贷款） |

- **！40 说明**：betteroffice 侧字节级未改任何公式（`C15`/`W7` 重算前后完全一致）；COM 侧被既有守卫 `重算改变公式`（`native_recalculation.rs`，`f2eed96` 引入，本任务未改）拦下——Excel 保存时把 shared formula 从属格展开。属既有 COM 侧行为，非本任务回归。
- 校验区域扫描有效（！40 命中 `W7 = #DIV/0!`）；其余模板 hits=0 为真实业务状态（双路由一致）。

**三条配置分支实测**：

| 分支 | 结果 |
| --- | --- |
| `PYTOOLS_BO_CALC=off` | ✅ 纯 COM（engine=Excel.Application，3950ms） |
| `MAX_VISITS=10` | ✅ 预算耗尽→自动回退 COM，engine 留痕 `betteroffice 回退（重算预算耗尽（1329 个公式格未求值…））` |
| `FALLBACK=none` + 预算不足 | ✅ 明确报错，不静默采用受限结果 |

**性能**（重算段，20 模板合计）：betteroffice **4.80s** vs COM **16.69s**，**快 3.5×**；单模板 30-705ms vs 203-4309ms。

**回归**：audit-core lib **281/281**（+3 bo_calc 测试）、excel-editor 76、excel-reader 26、excel-writer 1、ooxml-patch 22，全绿。`s1_audit_directory_cli` 2 个集成测试失败为**既有失败**（已用 HEAD 版本源码复现确认，与本任务无关）。

**FAST 渲染器热点压缩至 1.7s（2026-11-08 第四轮，反超旧 audit）**：逐段计时（AUDIT_S3F5_TIMING=1）定位：每机构渲染仅 ~650ms（已快于 Python），**全部墙钟时间在模板编译 41.8s**——`xml_attr`/`cell_text` 每格每次调用都 `Regex::new`，全模板数十万次编译占编译期 99%。修复：属性正则按名缓存（Mutex<HashMap>）、`<v>`/`<t>` 提升为 `OnceLock<Regex>` 静态 → **compile 41.8s→890ms（47×），每机构 650ms→121-146ms（5×），全批次 40.6s→1.7s**。solo 三方逐格 0 差异保持、输入 SHA-256 不变、278/278。终态对比：**Rust 1.7s vs 旧 audit FAST_OOXML 5.4s（快 3.2×）/ OPENPYXL 11.2s（快 6.6×）/ 自身 EditPlan 路线 513-601s（快 ~300×）**。

**FAST_OOXML 移植 Rust + 性能定论（2026-11-08 第三轮）**：逐段计时定论（CLI developer-view）：S3-F05 耗时 >99% 在 EditPlan 编辑管线+重验证器（read 133ms/planning 510ms 可忽略；EditPlan 路由每表单 133-145s），"逐格计划"架构本身是热点，Rust 实现快慢无关。按用户指示**整体移植旧 audit `fast_ooxml.py`（1335 行）算法为 Rust 新模块 `s3_fast_ooxml.rs`+`s3_fast_ooxml_render.rs`**：模板编译一次（CompiledSheet 行结构/keep_cols/indicator_rows/new_dimension + 无表头/无数据列空占位表跳过）→ 每机构纯字符串重建 sheetData（编译行内容复用+计划目标覆盖，无逐格扫描）→ SheetPlan 删表/降级隐藏（公式引用迭代收敛 + definedNames owner 判定/localSheetId 重映射/越界丢弃 + 保留表公式名称引用否决救回）→ 一次落盘 + 轻量完整性校验。语义锚点修正：Python `EMBEDDED_CONFIG_SHEETS=("使用说明","转表设置")`——报表清单是保留页（此前 EditPlan 路由误删）。路由：`edit_engine` auto/ooxml → FAST 渲染（fail-closed、机构级并行保留）；openpyxl → 编辑计划不变。**solo 三方验证逐格 0 差异（两跑/vs openpyxl/vs Python 基线）、输入 SHA-256 不变；耗时 513-601s（初始）→ 141-144s（批量化+并行）→ 40.6/40.7s（FAST 移植），与旧 audit 同数量级（旧 OPENPYXL 11.2s / FAST 5.4s）**。excel-editor 76/76、audit-core lib 278/278。

**S3F5 性能优化：机构级并行 + ooxml 按表批量回填（2026-11-08 第二轮）**：对齐旧 audit FAST_OOXML（5.4s/8 表，8 进程+模板占位符回填）的实测促使两步改造——①`s3_indicator_split` 机构循环改 `std::thread::scope` 并行（线程数=min(机构数,可用核)；进度经 mpsc 汇流到主线程 FnMut 回调；结果按机构序聚合，任一失败整体 fail-closed）；②ooxml 后端值/样式操作按表累积成批：单遍扫描改写既有格（自闭合分支防吞邻居）、缺失格复用单格路径插补、styles.xml 全局去重（同码 numFmt/同色 fill 只解析一次 xf，样式格保留除 s= 外全部属性）。solo 三方验证保持逐格 0 差异；**ooxml 路由 513–601s → 141–144s（3.6×），openpyxl 路由 150s → 37.6s（4×）**。与旧 audit 11s 仍有差距（其 8 进程 + 模板重绘粒度），进一步追平需把编辑计划改回"模板+回填"语义。excel-editor 76/76、audit-core lib 278/278。

**S3F5 ooxml 直改后端真实验证与路由事实澄清（2026-11-08）**：此前"ooxml DeleteColumn 已验证"实为**回退掩盖的假象**——旧实现含 rust regex 不支持的 look-ahead 正则，ooxml 一执行即报错，auto 路由静默回退 openpyxl，此前"0 差异"实际是 openpyxl 跑出的。本轮用 `probe/verify_s3f5_engine_solo.py`（强制 `edit_engine=ooxml`，仅注册 ooxml 后端、fail-closed 无回退）做到**真 solo 验证**：ooxml 两跑 + vs openpyxl 强制跑 + vs Python 基线三方**逐格 0 差异**（8 份表单×三方），输入 SHA-256 不变；32 条 dimension 元数据备注全部来自**未触碰空表的字节保留语义**（模板空表原样保留宽 dimension，非格值差异）。逐层修复的真实缺陷（每个均有单测）：①look-ahead 正则×3；②自闭合单元格被第一分支吞噬邻居；③样式写入丢值且不闭合标签；④workbook.xml `<sheet>` 重建丢 `name=` 属性；⑤属性间缺空格（无效 XML）；⑥**DeleteColumn 只匹配 `<c r=` 前缀**——删除格属性/内容漏进尾部输出、左移格把行内后续文本重复复制（真实 3.3 模板复现，此前从未真正执行过）；⑦插入超界单元格不同步 dimension（openpyxl iter_rows 类读取器按 dimension 截取会丢列）；⑧删表后 definedNames 不归一（localSheetId 失配，校验器按 openpyxl worksheet-scope normalization 拒收）。excel-editor 75/75、audit-core lib 278/278。性能基线（优化前）：ooxml-solo 513–601s vs openpyxl 150s。

**八功能对齐最终验收（2026-11-08，全量两跑）**：`tests/compare/rustaudit_alignment_p1.py --runs 2 --features all` 对基线 3.open-data-audit@da7b1c3 逐功能验证，**8/8 通过、367 输入 SHA-256 不变**：
- S1F1 3602 条逐条一致（3345 公式 + 257 条件格式）
- S1F2 505 行逐格 0 差异（任务说明 130 / 业务说明 333 / 本地校验结果 45）
- S2F1 1087 行逐格 0 差异（729 + 357 + 4）
- S3F1 3868 行逐格 0 差异；S3F2 单表逐格 0 差异；S3F3 exit 0 两跑一致；S3F4 33 行逐格 0 差异；S3F5 8 份表单逐格 0 差异
本轮为达成 8/8 补的三处对齐：①S1F2 功能名匹配 `汇总_` 前缀 + 驱动改传「任意行汇总」（原误传工作表名）；②S2F1 外部表 label 探测列范围（1..=4）与源生产 `range(min(4,len))` 一致，缺失时按 0 兜底；③驱动 S2F1 外部文件改用基线同款 `config/外部文件-大集中相关数据20261016.xlsx`。

**S2F1 / S3F2 对齐修正 + S3F3 真实闭环（2026-11-08）**

- **S2F1**：外部数据导入提示（原 Q002「外部数据导入」8 条）此前被并入 findings，导致「校验规则」表 12 行 vs 基线 4 行。源生产 `service.py` 明确其属**运行诊断**，不计入审核发现或异常统计——已改为只计数并逐条写进度日志，`s2_service` 与 `s2_analysis` 两处调用点同步移除并入逻辑。findings 恢复 1087 条（729+357+4 行）。
- **S3F2**：工作表名改为按比较结果标题推导 `(说明)YYYY.MM 说明文件`（对齐 `_explanation_sheet_title`）；删除自加的「生成信息」诊断表（基线只有单表）；输出校验器同步改为单表 + 名称前缀断言。
- **S3F3 真实闭环验证**（不使用伪造反馈）：从真实 S3F1 结果中选 5 行「是否说明=是」的真实行，按业务操作在「输出说明」列标记 → S3F2 导出 5 行到 `(说明)2026.10 说明文件` 单表 → 在导出文件「变动原因」列填写说明 → S3F3 原地回写：**imported 5 / unmatched 0 / conflicts 0**，回写后 5 行变动原因全部非空，更新前备份已生成。
- 已知环境限制：真实机构反馈原文件为 `.xls`，Rust S3-F03 与 Python openpyxl 均不支持该格式（基线 S3F3 的 feedback 实为 xls→xlsx 转换件）；实际业务需先另存为 .xlsx。

**外部引用剥离逻辑对齐 Python v2.7.7（2026-11-07）**：Rust 原先只处理索引式 `[N]表名!`，漏掉路径式三种形态。现按 `formula_region_writer`/`excel_open_repair` 逐条对齐，覆盖四形态：①`'C:\路径\[簿.xlsx]表名'!` ②`'[簿.xlsx]表名'!` ③`[簿.xlsx]表名!` ④`[1]表名!`（含 `'表名'` 变体）；扩展名 .xlsx/.xlsm/.xlsb/.xls；表名支持空格/全角空格/`''` 转义；函数内出现同样处理；本簿内引用不动。**索引式与路径式一律无条件剥离**（业务口径：副本只允许工作簿内引用；本地无同名表时公式成 `#REF!` 按异常计入，优于整簿被 Excel 拒绝）。编码兼容：表名匹配同时覆盖字面 UTF-8、`&#N;`、`&#xN;`（openpyxl 保存 NCR 实体的文件），剥离后解码为字面。**检测与剥离同源**：残留判定 `has_any_external_ref` 同时识别索引式与路径式，避免误删仍被引用的外链定义。外链清理：确认无残留时移除 `<externalReferences>`、workbook rels 的 externalLink 关系、Content_Types 覆盖与 `xl/externalLinks/` 全部部件；有残留则保持原样。
真实文件验收（毒模板 `！70.金融基础数据采集（债券业务）202608.xlsx`）：剥离 **12 处**、移除 **2 个**外链部件、`[N]` 残留 0、externalReferences/rels/CT 声明全部清除、Excel COM 打开 **7 表**成功、公式为 `INDEX(集中系统数据!$1:$1048576,...)`、原始文件 SHA-256 不变。单测 17/17（四形态 × 引号/无引号 × 函数内 × NCR 编码 × 空格表名 × 本簿引用不动 × 缺失表日志 × 检测同源）。

**毒模板专项验收（新增，可复跑）**：`tests/compare/poison_template_acceptance.py` —— 以自带 externalLink 的 `！70.金融基础数据采集（债券业务）202608.xlsx`（模板内 12 处 `[1]` 外链，指向旧机器断链路径）为输入，走 S1-F01 正式流程生成 7 个审核副本并断言：副本本地表 `[N]` 残留 0、Excel COM 只读打开 7/7、对照组手工写回 `[1]` 后必打不开、原始源文件与模板 SHA-256 全程不变。报告 `probe/reports/poison_template_acceptance.json`。该批条件格式异常 13 条，与 S1F1 全量对齐中此 7 本的 CF 数量一致。


**S3F1 表达式语法口径更正**：源生产 `run_comparison` 未显式传 schema 时用 **LEGACY_8（八段兼容表）**；五段式为 GUI 设置中心可选（central_expression_schema 默认 FIVE_SEGMENT_V1 仅供配置检查）。对齐驱动已改为 legacy-eight-segment。
- **高级设置中心（P2）**：五页签（新增组合设置空态）、updates 四子页（更新配置由 configUpdateSupported 门控，Windows 恒 false）、get/apply_config_update 桥接（备份+哈希校验+重启提示）、条件格式判定方式（PYTHON 默认/COM_DISPLAY/COM_EVALUATE 显式报错未实现）、空白格开关、大集中容差（元）、指标名去空格、完成自动打开、extra_files 手动追加后端（修复自死锁）、PREFERENCE_SPECS 新增 FilePathList/FloatRange。
- **验证**：全量 lib 271/271；对齐驱动 s3f1/s1f1 两跑一致证据在 `probe/reports/rustaudit_alignment_report.json`；诊断取证 `probe/reports/bond_cacheless_excel_reject_diagnosis.md`。已推送 origin/main 至 090a7e2 及后续。
- **遗留（下一轮）**：S3F5 ooxml 直改性能优化（单遍批处理，见 2026-11-08 性能实测）或路由调整决策；S3F5 五操作能力证据登记（production_capability_verification）；GUI 人工验收。

## 2026-11-07 高级设置对齐旧 Audit：手动追加文件、空白格开关、大集中容差/去空格/自动打开

- **手动追加待处理文件（extra_files）补齐后端**：前端按钮此前调用的 `append_extra_files`/`remove_extra_file` 接口在 Rust 桥接层不存在（点击即报"尚未实现接口"）。现已实现：批量追加源目录外工作簿（存在性校验、规范化去重、跨会话持久化到设置 JSON 白名单 `extraFiles`）、单个移除；追加/移除后自动重建源文件清单并并入模板匹配（直接位于输入目录下的追加文件由目录扫描覆盖，不重复并入），清单行带 `extra` 标记。定向单测覆盖追加去重/不存在文件跳过/持久化/移除。
- **空白格命中开关**：新增高级设置"空白格命中计入结果"（默认开启，与 Excel/WPS COM DisplayFormat 实际渲染口径一致）。关闭时在条件格式提取链末端排除目标单元格为空白（值 None 或空文本）的命中，其余口径不变，并写入一条 `blank_cells` 诊断说明排除数量；打开行为与既有结果完全一致。旧 Audit 同名开关口径一致（其过滤谓词即值 None/""）。
- **大集中核对容差全局默认**：3.1 表达式规则「容差值(万元)/Thd」列未填时，由原硬编码回退 0.01 改为读取设置中心"核对容差（元）"（默认 100 元 = 0.01 万元，范围 0–1,000,000，手工改坏的设置文件回落默认）；配置了容差的规则仍按配置值（万元→元 ×10000）执行，非数值仍报错。注：旧 Audit 同名设置只作用于跨期比较的大集中核对；其 3.x 复杂规则空 Thd 回退 0.01 是旧代码口径，本次按用户确认改为全局默认。
- **指标名称去空格（默认关）**：开启后 S3-F01 比较结果工作簿的指标名称显示列（明细第 8 列、读回校验同列）去首尾空格；只影响结果显示，不影响规则匹配、说明文字或核对逻辑。经逐格扫描当前正式 `config-real/3.1大集中执行比较_配置.xlsx`：全部字符串无一例首尾空格，此开关目前是兼容保险。
- **完成后自动打开结果（默认关）**：开启后大集中执行比较成功时用系统默认程序自动打开比较结果工作簿，失败只写日志不打断。
- **设置管道扩展**：`PREFERENCE_SPECS` 白名单新增 5 个键（`extraFiles` 路径数组、`conditionalFormatIncludeBlankCells`、`centralTolerance` 浮点范围、`centralTrimIndicatorNames`、`centralAutoOpenResult`），`PreferenceKind` 新增 `FilePathList`/`FloatRange` 类别（校验、状态回填、持久化全覆盖）；GUI 高级设置页新增"空白格命中计入结果"（计算与条件格式组）与"大集中执行与结果"组（容差输入 + 两个开关），双布局（按功能/按系统）均已接入。
- **修复 remove_extra_file 自死锁**：首版实现持有 UI 锁调用会再次加锁的清单刷新；被新增单测暴露（测试挂起 60 秒以上），已改为先释放锁再刷新。
- 验证：lib 全量单测 270/270 通过（12 忽略）、s1f1_ooxml_route 集成 2/2、debug lib+bin 构建通过；rustfmt 已将 `web_bridge.rs`/`app_settings.rs`/`s3_execution_compare.rs`/`excel_open_repair.rs` 整体格式化（其中包含此前在途改动的纯格式整理，内容不变）。未做 Release 构建/GUI 人工验收/提交。

## 2026-11-07 S1-F1 审核副本 Excel 打开修复 + 条件格式提取来源默认切到审核副本

- **Excel COM 打开修复（与旧 Python Audit `excel_open_repair.py` 逐条一致）**：报送文件中两类非规范内容会让 Excel COM 拒绝打开——公式裸引用含全角括号等强制引号字符的表名（如 `个人贷款发生额信息（隐藏发放）!A1`），以及冻结窗格 `topLeftCell` 行号落在 `ySplit` 冻结区内。新增 `audit-core/src/excel_open_repair.rs` 在 S1-F01 公式副本准备完成之后、重算之前对审核副本做无害规范化：`<f>`/`<formula>`/`<c:f>`/`<definedNames>` 四类上下文中 `(?<!')` 的裸引用补单引号（表名内 `'` 加倍），冻结窗格 `topLeftCell` 行号改为 ySplit+1。只改审核副本，原始报送文件永不改写；廉价预检（无问题不重写）；临时文件原子替换；修复失败吞掉异常不打断审核（WPS/LO 本来就能打开）。接入点按“路由到 Excel COM 提取校验时修复”门控：Windows 且计算引擎 `auto|excel` 时生效，WPS/LO 与 Linux LO 批次路径不修复。GUI 运行日志输出修复明细，结果 JSON 增加 `excelOpenRepair` 统计；新增 CLI `excel-open-repair --workbook` 子命令。
- **验收（28 本 Excel 打不开的报送文件）**：修复统计与旧 Python 实现逐本一致（补引号 456/688/2928、修窗格 1）28/28；修复副本 Excel COM 只读打开 28/28 成功；对照组原样副本抽样 6/6 仍失败；负向控制无问题文件抽样 12/12 `changed=false` 且字节不变。单测 11/11 通过。报告：`probe/reports/excel_open_repair_acceptance.json`，驱动脚本 `tests/compare/excel_open_repair_acceptance.py`（临时副本已清理，原始文件只读）。
- **条件格式提取来源默认改为“复制校验区域并重算后的审核副本”**：审核副本现在先修复再重算，扫描区域提取、Rust 求值与 COM 渲染都在同一份修复后的副本上进行，对拍口径统一。`s1_conditional_format_source()` 缺省值由 `source` 改为 `audit_copy`，GUI 初始状态与设置页默认项同步调整；`原始源文件` 仍可在高级设置中显式选择用于对照旧口径。顺序保持：提取发生在复制+修复+重算之后。S1-F3 仍按恢复任务记录的口径继续提取，旧恢复任务没有该字段时维持源文件口径。本机用户设置 JSON 中遗留的旧默认值 `source` 已迁移为 `audit_copy`。此前 129 本 CF 对拍（Rust 求值 vs Excel/WPS COM 渲染 MATCH）就是在审核副本口径上完成的；当 Rust 求值器完全覆盖 Excel/WPS COM 后可只保留 Rust 求值。
- 状态：debug 构建（lib + bin）与定向单测通过；`cargo fmt --check` 本任务改动块全部干净。未做 Release 构建/打包、未提交（工作树尚有此前在途改动），未跑 GUI 人工验收。

## 2026-11-07 00:50 S1-F1 条件格式显示色复跑

- 按业务确认，条件格式异常以所选工作簿中 Excel `DisplayFormat` 的实际填充色变化为准；空白值若实际变红也计为异常。条件格式扫描优先使用所选源文件/审核副本中与映射匹配的“条件格式区域”命名区域；没有匹配命名区域时，只使用该工作簿 OOXML 条件格式规则的 `sqref`，不扩展到整张工作表。没有条件格式规则时按空结果完成。
- 修复了混合颜色矩形的误报：Excel 对多单元格混合范围可能仍返回一个标量 `Interior.Color`，Rust 之前将该颜色套用到整块。现在先检查 `ColorIndex` 是否为单一值；混合范围递归拆分，逐格比较。单元格级读取仍可在 `ColorIndex` 不可用时回退读标量颜色。
- Release 单文件冒烟：广东惠东单位存款源文件扫描 2,701 格，COM 显示 47 个颜色变化；Python 同输入有 46 个条件格式项，46 个位置全部一致。唯一 Rust-only 项是空值但实际显示红色的 `单位存款发生额分产品期限!AX10`，符合已确认的新口径。该运行共 10 条公式异常、47 条条件格式异常，源输入哈希未变，用时 18.6 秒。
- Release 全量 129 本：129 个结构通过，公式异常 3,345 条（与 Python 的 3,345 条位置完全一致），条件格式异常 257 条；Python 基线为 264 条。对照为共有 256 个条件格式位置、Python-only 8、Rust-only 1（AX10）；Rust 总计 3,602，Python 总计 3,609。Rust 8 条暂缺均为 4 本“单位贷款置换旧债发生额”工作簿在 Excel COM 显示色扫描时 `Workbooks.Open` 返回 `HRESULT 0x80020009`，每本对应统计表 `C36` 和 `D36`；它们不能视为 Rust 已验证无异常。另有 7 本同类型工作簿状态为部分完成（其中 3 本此轮未形成 Python-only 异常），需要后续决定是否补充安全的 COM 打开回退。
- 全量耗时 301.275 秒，生成 129 个审核副本及结果表；成功 122、部分完成 7、失败 0、跳过 0。全部验收输入 SHA-256 未变化；Excel 进程数量运行前后回到 3，未额外遗留进程。该耗时不与前一版 2,427 条误报运行作性能对比。
- 定向扫描路由测试 3/3 通过；`cargo check -p audit-core --bin open-data-audit` 与 Release 离线构建通过。未运行全套 Cargo 测试，未进行 GUI 人工点击验收。完整记录见 `audit-core/probe/reports/s1f1_real_sample_parity_20261106.md` 的 2026-11-07 补验，以及 `F:/RustBuild/audit/tmp/s1f1-rust-current-template-129-20261107-run6/acceptance/`。

## 2026-11-06 18:21 业务复核更正：条件格式标签与票据浮点残值

- 五条单位存款条件格式异常都由目标 F 单元格自己的数值触发，条件为 `AND(Fn<>"",OR(Fn<=-30,Fn>=30))`。目标值是数值；上方 F38/F40/F43、F26 或 F11 的 `#DIV/0!` 与该规则触发判断无关，也不应追加到 `check_field`。`check_field` 按结构区域中的指标行文字生成，例如 `客户地区（6）_陕西省｜余额_环比增减(%)`。旧 Python 的 `#DIV/0!` 后缀是标签扫描误读上方错误缓存，当前源码继续使用结构标签；最新 129 本验收仍待用当前源码重建/复跑。
- 重新逐格检查 11 个票据位置的 Python 与 Rust 审核副本：两边公式文本相同，Python 副本中这 11 个 W 单元格的公式缓存均为空；Rust 经 Excel 重算后的缓存为 10 个 `-3.63797880709171E-12`（TCL）和 1 个 `+3.63797880709171E-12`（广东博罗）。因此旧 Python 汇总未收录它们与旧副本缓存为空相符；不能说 Python 算出了零，也没有证据表明 Python 有专门的残差过滤。
- 以上 11 个绝对残差均小于 `1E-10`。建议由模板公式显式改成 `ABS(差值)>1E-10` 才报错；当前模板的精确 `<>0` 仍按原意执行。未改测试模板，也未在 Rust 提取层全局过滤。是否正式采用该容差由业务方决定。
- Markdown 明细表中的 `\\|` 是为在表格内显示字面竖线而做的转义；Excel 中的公式文本和缓存字符串都只有 `|`，不包含反斜杠。旧公式 `K134` 的真实缓存 `错|结清合计||结清合计为0，请核实。` 按现有解析规则填入错误类型、校验类型、空校验指标和描述，数值字段留空；后续新模板会统一输出七段 `错|校验类型|校验指标|详细描述|本期值|对比值|差值`。
- 只读核对 Python/Rust 汇总及 11 个审核副本公式/缓存；本次未运行 Cargo、未重建 Release、未改源/模板工作簿。

## 2026-11-06 业务口径补充：S1-F1 条件格式标签、公式结果字段

- 五条单位存款条件格式异常仍应按目标 F 单元格的数值和既定条件格式规则提取；上方标签候选中的 `#DIV/0!` 不参与指标名。指标名应由表结构范围解析，例如 `客户地区（6）_陕西省｜余额_环比增减(%)`。已撤回“把错误缓存转成 `#DIV/0!` 拼入 `check_field`”的局部改动；定向单测 `error_valued_header_is_not_used_as_indicator_text` 通过 1/1。现有 129 本 Release 产物是前一规则生成的，尚未用当前代码重建/重跑；触发数量与值逻辑不变，但五条字段和 `issue_id` 会回到结构标签。
- Rust 的公式结果解析顺序已经是七段：`错误级别|校验类型|校验指标|详细描述|本期值|对比值|差值`，第 4–7 段依次进入对应结果字段。用户示例对应的旧模板 `个人贷款发生额信息!K134` 实际生成 `错|结清合计||结清合计为0，请核实。`；公式以 `E134` 作为第三段，但本样本 `E134` 为空，且公式没有生成“数值”标签或本期/对比/差值三段。因此 Rust 填入校验类型和描述、把校验指标及数值列留空是按实际公式内容解析，并非丢弃已生成内容。类似 `K69` 生成 `错|发生合计||结清合计为0，请核实。`。未来模板统一七段后将由现有解析器逐段填列。
- 对 11 条约 `±3.64E-12` 票据残值，建议将业务阈值设为 `1E-10`，但应写入对应模板公式（差值绝对值不超过阈值即通过），不要由 Rust 对所有公式结果做通用抑制。当前实现仍严格执行现有 `<>0` 模板公式；本轮没有修改测试模板或数值判断。
- 独立 Agent 只读复核确认：95 条应保留；11 条若采用容差应改模板、不能全局过滤；五条条件格式应使用结构字段。严格逐字追平旧 Python 与这些已确认语义冲突，当前旧 Python 在 95 条漏报、在五条字段上追加错误缓存文字。17:41 Release 和 17:25 运行是旧标签实现，不能代替当前口径验收；本轮未重建或重跑。

## 2026-11-06 17:25 增量：S1-F1 结果差异复核

- Rust 条件格式表头扫描现按旧 Python 规则把 Excel 错误缓存转为显示文本。新 Release 对原 129 本、19 个实际命中模板及相同配置/外部数据复跑：129/129 完成，0 跳过、0 失败、0 待人工重算；公式异常 3,373、条件格式异常 349；157 个输入 SHA-256 未变；耗时 283,415 ms。编辑后端 OOXML Patch 129 本，重算后端 Excel 122 本、WPS 7 本。
- Python/Rust 对照：共有 3,616 条记录现在全部字段一致（之前 5 条条件格式 `check_field`/`rule_id` 差异已解决）；Python-only 0；Rust-only 106 个位置全部保留。95 条是 Python 提取器可复现 `IndexError` 漏掉的模板校验区域公式，11 条是按业务决定保留的票据 `<>0` 精确比较残值，因此**S1-F1 尚未达到与旧 Python 的字面全量一致，不能标记完成**。字段名修复会改变 5 条历史 Rust `issue_id`，后续历史连续性需人工留意。
- 7 本债券业务副本均无 externalLinks 部件；`参照表!D2:D211` 每本 210 个公式均引用本地 `机构参照表`。与外部参照工作簿三张辅助表逐格比较，公式/常量和缓存均为 0 差异。
- 定向单测 1/1 通过，`build-release.bat` 成功。全 workspace 回归和 GUI 手动点击未执行。结果及逐条证据见 `audit-core/probe/reports/s1f1_real_sample_parity_20261106.md` 和 `F:/RustBuild/audit/tmp/s1f1-error-label-parity-129-20261106_172554/`。

## 2026-11-06 增量：S1-F1 差异复核与引擎默认顺序确认

- **票据规则**：按业务方明确要求保留模板公式的精确 `<>0` 比较，不在代码中添加数值容差；如需容差，由业务方调整模板公式。
- **复制引擎默认值**：当前源码默认 `自动`，高级设置列出 `自动 / OOXML / openpyxl / Umya`。自动顺序为 OOXML Patch → 环境可用的 openpyxl worker → Umya；只有前一后端在写入前明确拒绝时才从原输入副本尝试下一后端，写后验证失败即停止。该次核对为源码和既有定向测试证据；本次差异复核没有重新构建或进行 GUI 点击验收。
- **13 份旧版提取器异常样本**：逐项核对的 95/95 个 Rust-only 公式异常都落在模板定义的“校验区域”内，且 Rust 输出公式文本与模板完全一致，没有发现区域外误报。样本源路径、Rust 重算副本及单元格坐标见 `F:/RustBuild/audit/tmp/s1f1-corrected-inputs-129-20261106/s1f1_python_extractor_indexerror_13_paths.csv`；95 条异常详情见同目录 `s1f1_python_extractor_indexerror_95_items.csv`。这支持 Rust 结果符合模板规则，但仍保留给业务方查看具体数据与结论。
- **旧版 35 本复制失败复核**：历史批次为 7 本含过期 externalLinks 的债券模板故障、28 本公式复制写后校验失败。当前完整 Rust 批次对应 35 本均有成功结果，129 本复制区域检查无差异；旧 Python 使用业务方修正后的当前债券模板重跑 7 本，25 条异常与 Rust 逐字段一致，保护输入哈希未变化。完整证据见 `audit-core/probe/reports/s1f1_real_sample_parity_20261106.md`。历史完整批次未记录每本实际选用后端，因此不把 28 本成功归因到某单一引擎。
- **当前 Release 自动路由复跑（2026-11-06 14:45）**：使用 13:38 构建、晚于引擎路由源码修改的 Release EXE，按相同 129 本输入/19 类模板/配置/外部数据重新运行，状态 completed，模板匹配及结构检查 129/129、跳过 0，公式异常 3,373、条件格式异常 349、总异常 3,722；157 个输入 SHA-256 全部未变，耗时 332,639 ms，生成 129 个审核副本和汇总工作簿。自动顺序为 OOXML → 可用 openpyxl → Umya；逐副本记录实际编辑引擎 OOXML 129、openpyxl 0、Umya 0；重算引擎为 Excel 122、WPS 7。此前 35 本失败文件全部包含在本轮 129/129 成功集合。本轮没有重跑 Python；剔除批次号、时间和输出路径后，本轮与上一 Rust 129 批次的逐文件异常集合及业务字段 0 差异，因此此前 Python/Rust 对照差异仍是 95 条 Python 提取器漏项、11 条票据精确比较残值及 5 条共有条件格式描述字段差异。当前运行没有实际调用 openpyxl，不能据此声称它的整批回退已经验收。产物：`F:/RustBuild/audit/tmp/s1f1-auto-route-current-release-129-20261106/acceptance/`；本轮 13 份样本/95 条异常已重新映射到当前副本：`s1f1_python_extractor_indexerror_13_paths_current_run.csv`、`s1f1_python_extractor_indexerror_95_items_current_run.csv`（同目录）。
- **当前模板口径下 Python 全量复跑（2026-11-06 15:10）**：旧 Python GUI 的 19 个业务批次在同一组 129 个源文件、当前 19 份模板、原配置及外部数据上全部完成。19 份历史配置副本 SHA-256 均与旧/新正式配置相同；源文件、模板、配置和外部数据共 168 个保护文件运行前后哈希 0 变化。当前模板与上一对拍版本仅债券模板哈希不同，符合业务方已修正模板外链的说明。Python 产出 3,616 条，和 14:45 Release Rust 产出的 3,722 条对比：Python-only 0、Rust-only 106；共有项 3,611 条的业务字段一致，另 5 条为条件格式 `check_field`/规则编号文字（Python 追加 `#DIV/0!`）。汇总表中 60 个数值显示差异经数值解析后相等（如 `1825.0` 与 `1825`），不算业务差异。106 条 Rust-only 的位置集合与前次完全相同：95 条在个人/单位贷款命名校验区域、旧 Python 提取器可复现 `IndexError`；11 条为票据模板精确 `<>0` 触发的浮点残值。Python 19 批次总用时 428.385 秒、墙钟 429.81 秒；Rust 本批 332.801 秒，因 Python 按 19 批启动引擎而 Rust 是一次批次，不能直接作为严谨性能结论。完整 13 列对照：`F:/RustBuild/audit/tmp/s1f1-python-current-template-129-20261106-run2/python_vs_current_rust_summary_parity_normalized.json`；Python 执行结果与前后哈希见同目录 `python-results.json`、`input_hashes_before.json`、`input_hashes_after.json`。
- **校验范围纠正**：业务方确认模板“校验区域”外的文字、数值和公式可能被临时修改来观察规则触发，不应复制或提取。S1 复制计划仍只导入命名“校验区域”；公式复制完整性门槛现只比较实际导入区域，不再因源工作簿其他公式文本差异而阻断。旧错误 `存量同业借贷!F45` 是此前整簿公式保护比较报出的区域外位置；旧失败候选未留存前后公式文本，无法还原当时差值；当前源文件、模板和 Rust 副本的 F45 公式文本一致。定向单测 `s1_copy_validation_checks_imported_formulas_but_ignores_other_formula_text` 通过（1/1）。条件格式值默认从源核查表读取，结构/表头检查单独执行。
- **修正后的 Rust Release 全量复跑（2026-11-06 15:43）**：修复了公式复制校验对校验区域外公式变化的误阻断，并将 GUI 链路的条件格式求值输入改为原始源工作簿；公式异常仍从已重算审核副本提取。`build-release.bat` 成功，条件格式路径定向测试 1/1 通过。用新 Release 对同一 129 本、19 份模板、配置及外部数据重跑：129/129 完成、结构通过 129、跳过/失败/待人工重算均为 0；公式异常 3,373、条件格式异常 349、总计 3,722，耗时 292,626 ms，所有验收输入哈希未变。新 Python/Rust 汇总对照仍为 3,616/3,722，Python-only 0、Rust-only 106、共有项 3,611 字段一致、5 条条件格式元数据差异；Rust-only 的 106 个位置与前次完全相同，条件格式改读原始源文件后没有新增或丢失异常。产物：`F:/RustBuild/audit/tmp/s1f1-source-cf-current-129-20261106/acceptance/`；逐项对照 `F:/RustBuild/audit/tmp/s1f1-source-cf-current-129-20261106/python_vs_rust_after_source_cf.json`。剩余差异仍是 95 条旧 Python 可复现 `IndexError` 漏项、11 条模板精确 `<>0` 票据残值、5 条共有条件格式字段/规则编号文字；这些不等于 13 列结果完全一致，仍需按已确认业务口径验收。
- **条件格式取值来源选项**：已在高级设置中增加“原始源文件（默认）/复制校验区域并重算后的审核副本”选择，使用现有用户设置 JSON 记忆。S1-F1 将当前选择写入恢复任务；S1-F3 按该任务记录继续提取，旧恢复任务没有此字段时默认使用源文件。定向路径选择与设置持久化测试各 1/1 通过；真实 129 本批次仍按默认源文件口径验收，本选项的审核副本模式尚未做真实工作簿验收。
- **整体差异仍未达到逐条输出完全相等**：Python 与 Rust 异常总数为 3,616 / 3,722；95 条来自可复现的 Python 提取器 `IndexError`，11 条票据精确公式残值触发差异，另有 5 条共有条件格式异常的 `check_field` 元数据差异（旧 Python 拼入 `#DIV/0!`）。旧版对 7 本债券的基准因过期模板链接作废，已用修正模板重跑并得到上述 25/25 一致结果。

## 2026-11-06 增量：S1-F02 全量真实样本旧版对齐（当前模板口径已确认）

- **样本范围**：`测试数据/S1F2全量`，递归处理其中 129 本嵌套源工作簿；根目录 `汇总信息.xlsx` 是既有汇总结果，作为结构/内容参考，不得再次作为源文件；目录内 2 份 DOCX 不属于旧 Audit 的 `.xlsx` 源文件扫描范围。
- **旧版对齐结果**：通过旧 GUI 正式 DAG 入口和 Rust S1-F02 分别处理同一批 129 本源文件。初次对比发现数据行和值一致，但 Rust 工作表顺序、字段顺序不同；Rust 已改为按 Python 定义名称作用域/工作表顺序及模板字段顺序处理。复跑后，四张工作表、表头、行列位置及全部 505 行数据完全一致；结构预检部分继续 129、整本跳过 0、缺失表组合 175。旧 Python 与 Rust 均使用默认/AUTO 源读取和写入路径。
- **输入路径**：源目录 `测试数据/S1F2全量`；模板使用旧项目 `config-real/模板文件_命名区域/！汇总校验结果及报送说明.xlsx`（只读）；历史配置以旧 Audit 的 `config-real/1.逐笔统计系统_配置.xlsx` 为同输入（只读）。所有运行结果写入 `F:/RustBuild/audit/tmp/` 下独立目录，不覆盖源文件或现有 `汇总信息.xlsx`。
- **模板名称差异（业务方已确认）**：28 本源工作簿含“在线抽验结果”，当前模板定义“在线校验结果”；旧 Python 和 Rust 均依模板生成空的“在线校验结果”，不把源表自动映射过去。该结果符合当前模板口径，S1-F02 新旧同步已确认；模板名称由业务方后续自行调整。如模板改动后需要汇总该类数据，再按新模板重新运行即可。目录原有 `汇总信息.xlsx` 含“在线抽验结果”（123×15），不作为本次旧版 GUI 输出结构的判定依据。详见 [S1-F02 全量对照报告](audit-core/probe/reports/s1f2_full_sample_parity_20261106.md)。
- **验证**：S1-F02 Rust 单元测试 6/6、CLI 集成测试 6/6；Release 构建成功并更新 `dist/open-data-audit.exe`。源文件、旧版模板、历史配置和根级汇总共 132 个文件运行前后 SHA-256 不变。未运行全项目测试或 GUI 点击验收；单轮耗时仅作观察，不作为性能基准。

## 2026-11-06 增量：S1-F1 真实样本完整运行、Python 对照及复制后端更新（本轮验证完成）

- **真实业务范围**：按模板配置映射的“校验区域”命名区域复制值/文字、公式文本及单元格批注（位置、作者和内容），也覆盖目标区域空白；模板校验区域外内容可能为规则试验值，不复制、不扫描、不作为复制完整性门槛。源工作簿校验区域外既有公式也不参与复制阶段公式差异阻断。条件格式审核读取源核查表；结构/表头检查独立执行。普通样式差异不扩大范围。此次真实批次使用命名区域模板，未覆盖结构化“审核规则”模板分支。
- **Rust 完整批次**：使用当前源码构建后的 S1-F01 执行器、129 个源工作簿、对应命名区域模板、逐笔系统配置及旧 Python 项目的大集中外部数据。129/129 模板匹配、结构通过并完成重算；跳过 0、待人工重算 0、失败 0。发现公式异常 3,373 条、条件格式异常 349 条，共 3,722 条；生成 129 个审核副本及 1 个汇总工作簿。端到端用时 1,819,953 ms（约 30 分 20 秒），157 个输入文件均未改变。产物目录：F:/RustBuild/audit/tmp/s1f1-corrected-inputs-129-20261106/acceptance/steps/S1-F01/S1-F1_审核结果_20261106_111058_317；运行报告位于 acceptance/acceptance-report.json。
- **旧 Python 同输入对照**：Python 3,616 条、Rust 3,722 条；Python-only 0、Rust-only 106（全部是公式异常）。共有 3,616 条，3,611 条所有对比字段一致，另 5 条条件格式 check_field 仅有旧 Python 多出的 #DIV/0! 字符串不同；349 条条件格式异常的触发与说明一致。该结果不能宣称完全对齐，差异已经定位为：95 条 Rust-only 公式落在 13 个旧版副本/提取失败样本中，复跑旧 Python 提取函数时触发 IndexError，旧提取器漏掉这些公式异常；另 11 条票据公式以 <>0 精确比较，Excel 重算产生约 ±3.63797880709171E-12 并触发，旧 Python 缓存为空。部分样本还显示旧 Python 副本清除了公式引用的静态标签或明细数据，造成结果与保留原数据的 Rust 副本不同。应按校验区域和实际公式结果由业务人员核定，不为追求计数相同静默删除 Rust 发现。
- **命名区域复制对照**：129 本 Rust 输出与模板配置映射的区域逐格核验，无内容/公式差异；范围内比较了 16,898 个公式、896 个文字/常量及空格。旧 Python 有 21 本少 42 个区域末尾公式，主要因源表短于模板公式范围。命名区域批注按用户要求复制到 Rust 审核副本；旧 Python Direct OOXML 默认写入路径没有复制模板批注，因此当前批注行为按用户澄清执行。
- **外部表依赖差异待核**：旧 Python 按模板全簿公式识别所需外部表；Rust 从校验区域公式开始递归追踪依赖，并复制被选工作表 UsedRange 的值/公式。若校验区外普通公式单独引用某外部表，两端可能复制不同的辅助表。本次 129 本完整运行通过，但尚未证实该类模板依赖在所有模板上完全一致；需抽查模板外部表引用。
- **修改引擎与 openpyxl 备用路径**：S1-F01 高级设置提供自动、OOXML、openpyxl、Umya 四项。自动顺序为 OOXML Patch → 可用的 openpyxl worker → Umya；判断可用时检查 Python、worker 文件及 `import openpyxl`。仅前一后端明确在写入前拒绝时，才从原始输入副本尝试下一项；写后错误或候选验证失败会停止并拒绝发布。显式选项强制单一后端；该设置只影响 S1-F01。旧设置中的 `umya` 会迁移为 Auto，保留此前默认 OOXML 优先的行为。129 本既有批次通过 OOXML 完成，未完成 openpyxl 自动回退的整批验收。
- **本轮验证**：设置与旧值迁移测试 7/7、S1-F01 引擎路由测试 5/5、高级设置 API/UI 选项测试 2/2；`cargo check -p audit-core --bin open-data-audit` 通过；`build-release.bat` 成功并更新 `dist/open-data-audit.exe`。未重跑 129 本完整批次，未运行全 workspace 测试或 GUI 人工点击验收；Linux/Kylin/UOS 实机也未验证。此前 229/0/12 等记录属于前一版代码，不作为本次变更的全量回归证据。
- **待业务验收**：票据保持模板的精确 `<>0` 比较，Rust 不加入容差；需要容差时由业务方修改模板公式。请按 13 份旧 Python 提取器失败路径清单，抽查源工作簿、Rust 重算副本及异常位置：`F:/RustBuild/audit/tmp/s1f1-corrected-inputs-129-20261106/s1f1_python_extractor_indexerror_13_paths.csv`；95 条逐项异常记录在同目录的 `s1f1_python_extractor_indexerror_95_items.csv`。完整聚合报告：audit-core/probe/reports/s1f1_real_sample_parity_20261106.md。129 本按 19 类业务归类的清单：F:/RustBuild/audit/tmp/s1f1-python-classification-20261106/核查表业务分类.csv；原核查表和模板均未移动或修改。
## 2026-11-04 增量：S1-F1 Linux LibreOffice 批次复用

Linux GUI 现在会先准备本轮可审核源文件的独立公式副本，再调用一次隔离的 LibreOffice 批次重算；不再由 GUI 对每本工作簿单独调用 LO。WSL Ubuntu + LO 24.2.7.2 合成测试 2/10/33 本均通过，分别节省约 16%/68%/73%，输入未改且公式缓存有效。批次重算失败时继续保留无缓存审核副本供 S1-F3 人工重算。此实现是单次 `soffice --convert-to` 批处理，不是持久 UNO 会话。真实机构文件、Linux 桌面 GUI 和 Kylin/UOS 仍待人工验收；详见 [性能基线](audit-core/PERFORMANCE_BASELINE.md)、[迁移记录](audit-core/MIGRATION_INVENTORY.md)和专项报告。

历史状态摘要：2026-10-30（S1-F1 Windows 原生 COM 批量性能优化已完成最终验收；条件格式异常指标定位、结构预检/条件审核共用源读取及模板/源结构预检投影读取优化已完成；源结构预检同输入 Rust 对照为 501.532→237.579 ms，结构判断逐项一致；完整审核链三轮小型合成端到端复测：Python/Rust 业务步骤中位数 7,226.75/6,890 ms（1.049×，近似持平），结果、公式缓存、规则表批注作者和输入保护均通过；worker 墙钟中位数 32,747.97/6,965.34 ms，Python 值包含新 Excel 进程退出等待，不能解释为业务计算 4.70×；峰值工作集 502.141/297.863 MB；Rust acceptance 已补微秒阶段计时；并发负载下单轮诊断显示 Excel 重算约 3.38 秒、会话关闭约 2.93 秒，不纳入正式倍率基线；S5 配置审核仍按用户要求暂停开发。COM 批量结果见 [S1-F1 COM 最终验收报告](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.md)，源结构投影对照见 [专项报告](audit-core/probe/reports/s1_structure_source_projection_20261030.md)。）

状态只使用“已实现，待人工验收”“部分实现，需要继续开发”“尚未实现”“已明确暂缓”或“存在已知问题”。同一功能可有主状态和已知问题；代码、入口或局部测试不等于完整验收。实现、单元/冒烟、Windows 真实样本人工验收和 Linux/deb 分别记录。

## 2026-10-30 增量：S1、S2、S3 性能地图与持续优化

- **范围**：本持续目标覆盖 S1-F1/F2、S2-F1、S3-F1 至 F5。S5 已按 2026-11-05 用户决定从产品 UI 和代码中移除；历史报告仍保留作记录；S1-F1 原生 COM 批次与 S3-F1 的性能优化按用户指定暂时结项，不重复已完成比较。S1-F1 条件格式异常提取另发现生产结构区域指标定位瓶颈，本次已定点优化。
- **性能台账**：已建立 [性能基线](audit-core/PERFORMANCE_BASELINE.md)、[性能历史](audit-core/PERFORMANCE_HISTORY.md) 和 [benchmark 索引](audit-core/benchmarks/README.md)。基线先汇总已有 2026-10-29/30 的报告、机器数据及最终 COM 验收，不重复运行既有对比。
- **当前可用结论**：S1-F1 原生 COM 批次 2/10/33 册均与 Python 平价；S3-F1 最新优化报告为约 1.02×，两项暂结。S1-F1 带结构区域的条件格式异常流程中，指标组装由 15,523 ms 降至 57 ms；结构预检+条件审核 Rust developer event 耗时（不含进程启动）由 16,044 ms 降至 498 ms 后，进一步通过共享同一次源工作簿读取降至 390 ms（比前版再省 108 ms / 21.7%）；新状态五轮 8,610 条异常摘要完全一致，峰值工作集约 85.95 MB，与原状态相当。随后为结构模板增加专用 Calamine 投影读取，在 20k 合成结构/条件样本中模板预检读取由 528 ms 降至 224 ms（-304 ms / 57.6%）；三轮 43,051 条异常摘要与旧版完全一致，峰值内存仍由条件审核全量读取主导。S3-F3 已移除原地导入中临时 ExcelEditPlan JSON 的磁盘往返，同负载 Rust 内部中位数从 946 ms 降至 203 ms（减少 78.54%），优化后 Python/Rust CLI 墙钟为 912.62/235.54 ms，3.88×；输出、备份和输入保护对拍无差异。另有 1,000 条全冲突压力样本：Python/Rust CLI 1,332.49/149.89 ms（8.89×），冲突清单与目标工作簿值/结构无差异，Rust 冲突清单写出阶段约 18–21 ms。S1-F2、S3-F2/F4 在已测合成规模明显更快；S3-F2 新增 2,610 行输出规模剖析，Rust CLI 墙钟快 8.11×且 3/3 全输出对拍一致。S3-F4 的 10k 行、20k 输出结果在三轮 340,017 单元格对拍无差异下快 7.15×（CLI）；S3-F5 修正基准后，在 30×40×4 的有效回填负载下快约 2.28×（CLI），30 个工作簿逐格对拍无业务差异；Python/Rust 进程树峰值工作集约 475/24 MB。S2-F1 外部规则正常/异常样本在常驻服务计时中快约 1.75×，独立 CLI 计入启动后慢约 18%；三轮输出逐格与结构对拍无业务差异。S1-F1 完整链路最近三轮合成对照为 Python/Rust 业务步骤 7,226.75/6,890 ms（1.049×，近似持平），峰值工作集 502.141/297.863 MB；结果、公式缓存、规则表批注作者、Excel 打开及输入保护均 3/3 通过。Python worker 墙钟包含新建 Excel 进程退出等待，存在显著轮间波动。
- **数据质量说明**：2026-10-30 的原始 JSON 与汇总 Markdown 在 S2-F1、S3-F1/F2/F3 等行存在不一致；台账优先从逐轮 JSON 重算中位数，并单列 S3-F1 优化后逐轮数据缺失。另发现 2029/30 两份 S3-F5 旧基准表名导致 0 项业务命中，已将旧数字标为无业务代表性；当前基准生成器改用 A1 表名并设有命中数量门槛。历史报告和逐轮 JSON 保留原样，不覆盖。
- **尚缺性能证据**：历史记录有 S2 42 份 XLS 真实批次和 S3-F2 真实工作簿的阶段观察，但 S2 原始 21+21 份输入未保留，不能复跑同输入 Python/Rust A/B；旧 S3-F2 真实 2,610 行工作簿也未保留，不能解释其 1,186 ms 写出阶段对应的输入结构。S3-F2 合成放大剖析现包含阶段耗时与峰值工作集（Python/Rust 43.688/44.340 MB），没有显示内存优化空间；S1-F2 轻负载峰值工作集为 Python+Excel COM/Rust 300.078/10.719 MB。新增 Rust-only 阶段诊断记录了 S1-F2 的发现、读取、计算、写出和验证时间，但测量时检测到 21 个后台进程，明确不纳入性能基线；Python/Rust 三轮总耗时和完整输出对拍仍由既有报告支撑。S3-F1 10k 两期 CSV 峰值工作集为 Python/Rust 133.477/119.035 MB，耗时结项不变。S2 外部规则、S1-F2、S3-F1、S3-F2、S3-F4 和有效 S3-F5 专项记录了进程/进程树峰值工作集。S1-F1 完整业务链路已完成三轮合成耗时和峰值工作集测量；acceptance 现已提供请求准备、预检、复制/编辑、重算、异常提取、结果输出和 COM 会话关闭微秒计时，阶段剖析目前为并发负载下单轮诊断；其他 S1/S2/S3 功能仍有部分阶段需要同口径 A/B 记录。结构预检与条件格式现共用源工作簿的一次读取，4,000 行样本单次解析约 93 ms；尚未把这份投影跨公式副本/重算阶段保留，也未验证多文件场景的内存扩张。性能台账标注了数据规模、输入文件字节、输出正确性验证及其缺失项。
- **S1 读取证据边界**：历史脱敏真实工作簿仍记录 Python 77 ms、Rust 读取 983 ms（1,075 个固定标签）；该输入不可从当前夹具复跑，也没有文件字节数/峰值内存。本轮同输入 Rust 路径对照显示源结构预检完整读取+比对 501.532 ms、投影 237.579 ms（2.11×），五轮结果一致；模板预检另由 528 ms 降至 224 ms。两项都是合成数据的局部读取路径证据，未证明历史真实文件或完整 S1-F1 链路有相同比例提升；源投影对照未采集 RSS。
- **S1-F1 同模板批量优化**：模板/结构预检按 canonical 模板路径在单批次内缓存，避免每本源文件重复读取同一模板；模板结构读取现使用专用 Calamine 投影，在 20,000 行 Rust-only 样本中由 528 ms 降至 224 ms。三文件 GUI 合成路径均通过结构预检，因本机 Excel COM worker 未启动且无 LibreOffice，按预期保留人工重算副本。20k Python A/B 因参考路径耗时过长中止；本轮无跨引擎倍率，也未覆盖公式重算成功的完整 GUI 批次；后续三轮 S1-F1 全链路合成对照已通过，见 [报告](audit-core/probe/reports/s1f1_full_chain_20261030.md)。
- **S1-F1/F2 模板自动匹配元数据复用**：多文件 GUI/汇总入口原先逐文件扫描模板目录并重复读取每个模板的 Sheet 目录；现改为每次批次加载一次模板档案，源簿仍逐本读取并独立匹配。21 个源簿、5 个模板簿合成对照中，Release profile 中位数由 30.281 ms 降至 5.908 ms（省 24.373 ms，5.13×），21/21 建议完全一致。此项只减少匹配子步骤约 24 ms，不代表整项功能达到 5×。
- **本增量状态**：新增 S2-F1 外部规则、S3-F4 输出值级和 S3-F5 有效回填性能剖析；修正当前 S3-F5 基准生成器的 A1 表名映射，并为专用驱动添加非零命中门槛。新增 S2-F1 配置读取复用及输出轻量验证优化：配置解析 40→15 ms，21+21 合成批次保存阶段 312→98 ms、服务总耗时 659→429 ms；相关对拍通过。历史真实 21+21 文件已不可用，合成结果不替代真实文件验收。S3-F1/F5 历史数字的基准有效性限制已记录。S1/S2/S3 性能地图、现有瓶颈排序及有明确收益的优化均已完成；S1-F1、S1-F2 阶段诊断中的并发环境限制，以及缺失的历史原始输入/数据均在性能基线中标出。本性能目标按当前可复核证据结项；S5 已从产品代码中移除；历史性能数据仅供追溯。

## 2026-10-30 增量：S1-F1 结构预检与条件审核复用源读取

- **修改**：S1-F1 结构预检与条件格式审核针对同一本源工作簿共用一次 `WorkbookDump` 读取；结构比较和条件格式异常提取都使用该只读投影，随后即释放，不跨文件缓存，也不跨公式副本/COM 重算阶段持有。单文件 CLI、目录审核和条件审核入口均已接入。
- **实测**：4,000 行合成文件、32,008 个非空格、4 条条件格式规则、16,000 个扫描格、8,610 条异常，预热 1 轮、交错 5 轮。Rust developer event 内部耗时（不含进程启动）从前一实现 498 ms 降至 390 ms，减少 108 ms（21.7%）；结构/条件处理共用的源读取从约 99+97 ms 降至 93 ms 一次。Python/Rust 内部计时为 19,222.839/390 ms；49.29×仅适用于该专项。前后 XLSX 容器大小差 1 byte，数据与结构由同一生成器产生。
- **正确性和内存**：五轮每轮 8,610 条异常规范化 SHA-256 多重集完全一致，结构标签 8/8 匹配；本轮源和模板测量前后哈希不变。峰值工作集为 Python/Rust 77.312/85.949 MB，Rust 前一实现约 86.562 MB；没有观测到共享读取增加峰值内存。
- **验证边界**：定向 Rust 测试及 Release 构建通过，完整 S1-F1 GUI 链路、真实机构文件、全工作区回归和 Linux 环境没有在本轮验证；总性能目标继续进行。详见 [专项报告](audit-core/probe/reports/s1_source_read_reuse_optimization_20261030.md) 和 [五轮原始数据](audit-core/probe/reports/s1_preflight_reads_20261030_shared_read_after.json)。

## 2026-10-30 增量：S1-F1 条件格式异常提取性能剖析

- **负载与口径**：在 Windows 使用同一生成工作簿，4,000 行、32,008 个非空单元格、4 条条件格式规则、16,000 个扫描单元格；预热 1 轮，随后交错运行 5 轮。这个早期条件格式微基准没有把模板的 `structure_ranges` 传给两端，故未执行生产指标定位；计时为 Python/Rust 服务内部时间，不含子进程启动或完整结果 JSON 序列化，另以独立子进程采样峰值工作集。
- **结果**：Python 中位数 1,343 ms，Rust 246 ms，约 5.46×。Rust 读取/判定/异常组装中位数为 102/94/24 ms；Python 规则规范化/判定/异常组装约 62/844/453 ms。Rust 峰值工作集约 85.53 MB，Python 约 87.94 MB。
- **正确性与决策**：五轮均生成 8,610 条异常，异常字段逐项一致、诊断数为 0，源工作簿 SHA-256 未变化。该合成规模没有发现 Rust 瓶颈，不改业务代码。不过此前真实小负载单次观测为 Python 375 ms、Rust 618 ms，其中 Rust 读表 578 ms；该输入路径未保留、无法做重复对照，不能据此直接改读表路径，需有可复测业务样本再决定。此专项只覆盖条件格式子步骤；在该专项完成时，S1-F1 完整 GUI/业务链路仍未测量，后续已补三轮全链路对照（见 `probe/reports/s1f1_full_chain_20261030.md`）；COM 最终批次性能结论不变。
- **后续剖析候选**：当时静态路径显示 S1-F1 结构预检与条件格式评估各自读取源工作簿；下一节已针对真实结构区域测得两次读表成本，并定位到指标前缀重复扫描。跨公式副本/重算阶段缓存整簿投影的内存代价仍未测量。
- **验证范围**：专用性能驱动本身完成五轮对拍；使用既有 Release CLI SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。未运行 Cargo build/test，也未做全工作区回归；总性能目标继续进行。

## 2026-10-30 增量：S1-F1 条件格式指标定位优化

- **瓶颈**：结构区域启用时，每条触发异常都会重新扫描本列从结构起点到当前行的所有前缀标签。4,000 行样本中约 8,610 个触发项让 Rust `issue_build` 中位数达到 15,523 ms；Python 参考流程也存在相同的重复前缀读取。
- **修改**：Rust 按工作表、目标列对触发位置分组，每列按行一次累计相同的有序去重文本标签，再按原触发顺序组装异常。保留左侧行标签、上方列标签、合并单元格锚点、数值过滤、去重顺序和字段拼接，不改审核规则或 Python 项目。
- **性能**：相同 4,000 行结构区域样本，Rust 异常组装中位数 15,523→57 ms（约 272.3×）；结构预检+条件格式阶段 CLI 中位数 16,044→498 ms（约 32.2×）。优化后 Python/Rust 分别为 19,264.7/498 ms（38.68×，Rust 计 CLI 墙钟、Python 不含解释器启动）。峰值工作集约 86.61→86.56 MB。
- **正确性与验证**：五轮每轮均得到 8,610 条异常；规范化异常字段 SHA-256 多重集 5/5 相同，8/8 结构标签匹配，源文件及模板 SHA-256 不变。`cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture` 通过 5/5；`cargo build -p audit-core --bin open-data-audit --release` 成功。未运行全 workspace 回归，也未执行真实机构文件或 GUI 手工验收。
- **剩余项**：结构预检和条件审核各自读取源文件约 99/97 ms。因两阶段之间会执行副本复制和 COM 重算，当前不跨阶段保留整簿数据；需在完整链路和多文件内存样本中量化后再决定。报告及前后逐轮 JSON 见 [性能基线](audit-core/PERFORMANCE_BASELINE.md) 和 [专项报告](audit-core/probe/reports/s1_indicator_resolution_optimization_20261030.md)。

## 2026-10-30 增量：S3-F3 原地写回计划传递优化

- **瓶颈和修改**：S3-F3 原地执行原先先将已构造的 ExcelEditPlan 写入临时 JSON 并同步磁盘，再立即读回交给 Excel-Core。现在执行路径保留经过校验的内存计划并直接交给 Excel-Core；显式“生成计划”入口仍按原契约持久化 JSON。匹配、冲突处理、输入哈希保护、备份、原子发布及输出规则不变。
- **优化前后**：同一 1,000 条唯一反馈、无冲突负载，Rust 内部中位数 946→203 ms（减少 78.54%），CLI 墙钟 975.25→235.54 ms（减少 75.85%）；优化后 Python/Rust CLI 墙钟为 912.62/235.54 ms。Python 与 Rust 输出差异为 0；备份 SHA-256 与原目标一致，仅 H3:H1002 发生业务值变化，反馈源保持不变。
- **验证**：定向库测试 7/7（新增执行冲突路径用例）、计划 CLI 集成测试 3/3 通过，Release CLI 构建成功。未运行全工作区回归；全仓格式检查仍因已有格式差异未通过，没有格式化无关代码。工作区未提交、未 push。
- **剩余证据**：已补 1,000 键、100% 冲突的合成压力样本；该结果不代表真实冲突率。实际机构文件规模及 GUI 常驻服务负载仍待测。详见 [性能基线](audit-core/PERFORMANCE_BASELINE.md)、[历史](audit-core/PERFORMANCE_HISTORY.md)、[无冲突报告](audit-core/probe/reports/s3f3_parity_optimized_20261030.md) 和 [冲突报告](audit-core/probe/reports/s3f3_conflict_profile_20261030.md)。

### 2026-10-30 增量：S3-F3 冲突输出剖析

- **补充负载**：对 1,000 个匹配键分别从两份反馈表提供不同原因，预期 1,000 条冲突、0 条导入。Python/Rust CLI 三轮完整输出对拍均为 0 差异；冲突明细值/结构一致，目标原字节不变、备份 SHA-256 与目标一致，反馈输入未变。
- **性能**：Python 服务中位数 1,332.49 ms，Rust CLI 墙钟 149.89 ms，约 8.89×；Rust 服务内部 120 ms。Rust 冲突清单写出 18–21 ms；独立进程树峰值工作集约 41.04/16.39 MB（Python/Rust）。100% 冲突是压力样本，不能外推为生产冲突率。
- **可观测性及验证**：执行报告现暴露 `conflict_write_time_ms`，CLI JSON 可直接记录冲突清单写出阶段。S3-F3 定向库测试 7/7 通过，计划 CLI 集成测试 3/3 通过；当前 Release 构建成功，SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。全工作区回归未运行；代码和文档未提交、未 push。

## 2026-10-30 增量：S1-F1 COM 最终性能与批量稳定性验收

- **性能优化状态：完成。** Python `DispatchEx` 与 Rust 原生 COM worker 均创建独立 Excel 实例；2/10/33 册批次均只创建一个 Excel 实例。最终 Release 二进制 A/B 中位数为 Python/Rust：2 册 7,570.60/7,171.91 ms，10 册 12,246.33/12,139.69 ms，33 册 27,006.11/27,138.94 ms；Rust 在不同负载下交替领先，差异约 0.5–5.3%，结论为实测平价。
- **差距来源：** Excel 新实例冷启动约 2.4–2.5 秒；COM 激活中位数随样本/负载变化（本轮 `DispatchEx` 2,524 ms、`CoCreateInstance` 2,391 ms）。Rust worker `spawn` 与 ready 协议残差合计约 43 ms/批。应用属性设置约 74/73 ms，批内逐册打开、计算、保存、关闭及缓存验证持平。旧端到端口径只计 COM 释放约 0.2 秒；本轮完整计入 Excel 进程退出等待，两边约 2.7 秒。
- **正确性与稳定性：** 18 个正式批次（每引擎 9 批）共验证 96 个输出/引擎、576,000 个公式结果/引擎；Python 与 Rust 的公式文本哈希、结果哈希配对差异为 0，6,000 公式/7,200 个缓存值样本全部通过。每批仅一个专属 Excel PID；18 个 worker 所有权 PID 均正常退出，Rust worker 均退出，15 个基线 Excel PID 保留。监视快照另见短暂的非本 harness 隐藏 Excel PID，之后查询时已退出；未对其采取清理。无强制终止、原输入哈希变化或错误结果。损坏工作簿单文件失败后，批次继续处理其后的有效文件。
- **超时：** 定向 Windows Excel COM 测试 1/1 通过，受控响应延迟超过 35 秒后，父进程结束 worker 和确切归属的 Excel 进程，源文件未变且未发布输出。该测试模拟计算成功后的响应延迟，不能证明 Excel 内部 COM 计算永久卡死时的所有行为。
- **回归与构建：** Release 构建成功；全 workspace 82 个测试目标合计 501 passed / 0 failed / 34 ignored，另有两项真实 Excel COM integration tests 分别单独运行通过。忽略项包含需要外部业务样本或连接的测试。最终报告与原始逐轮 JSON：[`报告`](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.md)、[`数据`](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.json)。
- **尚待业务验收：** 本轮合成工作簿验证不替代机构文件上的 S1-F1 GUI 人工业务验收。S1-F1 性能优化已完成；没有发现值得以牺牲隔离或正确性换取的低成本优化。S5 已按用户要求从产品代码中移除。

## 2026-10-30 历史性能优化与复测记录（早期微基准，非最终批次验收口径）

- **S1-F1 Windows 原生 COM 热会话**：GUI/CLI 批次使用 Audit 自身的 COM worker，一批两册复用同一 Excel 进程；通过 Release 构建。定向合成会话测试 1/1 通过，确认进程 PID 复用、公式结果有效、源文件字节未变。没有重做完整 S1-F1 真实 GUI 验收，因此不把这项测试当作整链路人工验收。
- **S1-F1 性能范围**：早期三轮微基准测量两册、每册 7,200 个公式的 COM 重算/保存子步骤；报告中 Rust 4,950.6 ms、Python `Application.CalculateFull` 3,547.7 ms（0.72×），后续优化记录另列 Rust 4,668.5 ms 对 Python 3,595.5 ms（0.77×）。该组数字是特定热会话测量，不覆盖模板匹配、公式复制、最终审核结果输出；它与上方最终批量验收不是同一夹具和计时边界。最终验收每册为 6,000 个公式和 7,200 个非空缓存，批次 2/10/33 册均实测平价。两组证据均只涉及 COM 重算子步骤，不代表 S1-F1 整条业务链路。
- **S1-F2 性能与结果对拍**：补齐合成夹具所需模板结构后运行三轮。Python 中位数 4,320 ms，Rust 59 ms；逐项工作表、维度、单元格值及主要结构检查通过，列宽最大差约 0.54。输入文件未变；该规模仅 9 行输出且 Python 使用 Excel COM，不可泛化为语言级加速比。
- **其他已迁移功能性能复测**：S3-F1 10k 行 0.74x；S3-F2 5k/1,667 行 7.90x；S3-F3 1k 行 0.95x；S3-F4 10k 行 6.46x（后续专项逐格对拍为 0 差异）；S3-F5 30×40×4 的 3.79x 后确认旧输入没有任何业务命中，已作废并由有效 A1 表单负载替代；S2-F1 0.77x；S5 小负载 0.34x。S2 旧基准只核验输出表结构，S3-F3 核验原地备份/导入，S3-F2/F4/F5 的具体结果边界见性能历史报告。
- **构建与测试状态**：Release 构建成功；定向 `s1_native_com_session` 测试通过 1/1。没有运行全量 Cargo 测试、打包或 Linux 测试。完整性能口径、逐轮数据和限制记录于 `audit-core/probe/reports/perf_comparison_20261030.md` 与同名 JSON。
- **S5**：本轮只读取现有配置执行性能测量，没有修改或恢复 S5 业务开发。

## S1 逐笔统计审核

- **S1-F1 模板校验与审核：生产链路已打通，隔离真实样本的 Release GUI 点击验收通过，待用户按机构业务人工验收。** GUI worker 串起逐文件模板匹配、模板体检、结构不符跳过及原因记录、独立审核副本公式复制、外部数据引用导入、真实重算、公式/校验区域异常与重算副本上实际触发的条件格式异常提取、两类 Issue 合流、历史归类和 13 列结果写入。Windows Excel COM 16.0 合成 E2E 通过：独立预期为 `检查表!D1=#DIV/0!`、`D2=#REF!` 各 1 条公式异常及 `B2` 1 条实际触发的条件格式异常；输出可重开且源文件、模板、预重算副本哈希不变。2026-10-29 在 COM `xlPending` 修复后，用旧 Python S1-F1 的真实源文件、模板、外部数据、映射配置及同机构 Python 审核副本复跑 3 个真实探针，3/3 通过：约 210 个 `参照表!D2:D211` VLOOKUP 公式及依赖保留；模板 1,072 个公式文本一致，Excel COM 重算后 1,072 个存储结果与旧 Python 审核副本逐格一致（695 个非空，0 差异）；外部表独立重算基线一致；分段导入和完整副本均可由 Excel COM 打开；源文件、模板、外部表、配置哈希不变。Umya 将样本中 211 个显式空字符串单元格序列化为空白；其他非空值未变，且本样本 1,072 个公式结果仍与 Python 基线一致。2026-10-29 最新 Release GUI 使用隔离的真实源、模板、映射、外部数据实际点击执行并生成可重开的 13 列结果；该样本没有触发条件格式异常，触发型异常由合成 E2E 覆盖。该证据不替代用户对机构审核结论的业务验收。Windows 与 Linux x86_64 均已生成生产能力证据（Linux：`audit-core/capability_evidence/umya_linux_x86_64_20261028.json`），未登记平台仍失败关闭。校验区域扫描当前 GUI 使用重算副本、CLI 使用源文件缓存；统一口径待用户决定（AL-03）。
- **S1-F1 复制后运行时校验：已补齐并通过真实样本复验。** Excel-Core 生成的临时候选在缓存清理/真实重算前，按 S1 编辑计划重建预期公式与非公式输入值，比较所有工作表名称和顺序、公式单元格位置/文本及复制值；仅容忍已知零长度字符串规范化为空白，任一实质差异失败关闭。首次真实复验发现 Umya 对映射区域中缺失的源单元格未清除目标旧公式/值；已修正 Umya ImportRange 的空白覆盖语义，并由新的回归和候选校验验证。合成校验 15/15 通过，覆盖公式丢失/变化、引用表值变化、额外工作表和 210 条 VLOOKUP；随后真实旧 Python 工作簿完整探针 3/3 通过。真实 GUI 业务验收继续待办。
- **S1-F2 审核结果说明汇总：已实现，待人工验收。** CLI 与 GUI 入口已有模板/结构预检、任意行汇总和历史富化。历史记录有合成和部分真实样本证据；还需用户核对不同行数、异常跳过、历史关联及输出和原文件保护。普通列宽等外观差异不单独作为阻塞。
- **S1-F1 外部表公式复制：真实样本复制与计算通过，待 GUI 业务验收。** 旧 Python 的 Direct OOXML 优先、遇公式回退 openpyxl 属于 S1-F1。Rust 按本轮用户确认改为 Umya 主引擎、OOXML Patch 只承担已实现且语义相同的定点操作、Excel COM/LibreOffice 真重算。最新 Windows 真实样本探针覆盖纯数据源、普通公式和跨 Sheet 引用、约 210 个 VLOOKUP、同时复制「参照表」/「机构参照表」、模板公式及输入保护，并实际经过新增候选验证器。此前匹配样本已验证 1,072 个可读公式结果与同机构 Python 审核副本一致。2026-10-29 当前真实样本重跑中，纯数据往返与 Excel COM 分段/整本打开通过；公式复制/重算探针在未提供 Python 对照时通过，独立外部表重算结果、公式文本及缓存可读性检查均通过。使用同名 Python 对照时仅读到 973 个公式缓存结果，未达到探针要求的 1,000 个覆盖数，因此该次不形成 Python 结果对拍结论；此前已通过的 Python 基线对拍仍是独立历史证据。另一个 7 月目录中的同名外部文件与本次 `config-real` 映射下的预期非公式值不符，候选验证器在重算前明确拒绝该输入；未放宽校验。Umya 对缺失源单元格的导入现在清除目标旧值/公式。OOXML Patch 没有跨工作簿公式工作表复制能力，故 S1-F1 不登记该回退；Umya 对该操作失败时明确失败关闭。真实 GUI 整批审核输出仍待人工验收。
- **S1-F1 当前 Release 真实样本复验（2026-10-29）：3/3 通过。** 使用只读旧 Python 项目中的单位贷款报送样本及匹配的 `config-real` 模板、映射、外部数据，当前 `dist/lo-convert.exe` 经 Excel COM 完成重算。探针验证纯数据 Umya 往返、模板与外部表导入、至少 1,000 个模板公式位置/文本、外部 `参照表!D2:D211` 至少 200 个 VLOOKUP 公式以及其独立重算值、`机构参照表` 公式/值、完整副本公式缓存和原始四个输入文件哈希。未提供 Python 输出基线，因此此轮不宣称全部模板结果与 Python 逐格一致；另有此前 1,072 项完整基线对拍记录。首次选用的另一旧工作簿因已含同名“集中系统数据”页而被既有安全检查拒绝，且该候选独立 Umya 往返会改写其公式；未绕过校验，改用不含冲突页的报送样本后通过。日志：`F:\RustBuild\audit\s1f1-real-configreal-source-candidate-20261029.log`；被拒候选记录：`F:\RustBuild\audit\s1f1-real-current-goal-20261029.log`。GUI 整批输出仍待人工验收。
- **S1-F1 引擎路由与跨 Sheet 覆盖复验（2026-10-29）：3/3 真实探针通过。** 当前真实工作簿测试现在直接断言完整模板/外部公式导入批次由 `UmyaSpreadsheet` 唯一选中并成功执行；`参照表!D2:D211` 的真实 VLOOKUP 必须存在对「机构参照表」的跨 Sheet 引用，随后验证引用表与依赖表共同导入、独立 Excel COM 重算结果、输出重开和原始输入哈希。当前同语义公式工作表 `ImportRange` 没有 OOXML Patch 后备能力，故不尝试伪回退。`excel-editor` 单元测试 48 项通过（包括写入前拒绝后从原始副本切换到已声明 OOXML 能力、双方均不支持时失败关闭及写入后错误不回退）；公式坐标/区域集成测试 5 项通过；S1 公式复制单测 15 项通过。此次未提供 Python 审核副本作逐公式结果基线；真实 GUI 整批输出仍待人工验收。日志：`F:\RustBuild\audit\s1f1-umya-route-cross-sheet-current-20261029.log`、`F:\RustBuild\audit\excel-editor-engine-policy-current-20261029.log`、`F:\RustBuild\audit\excel-editor-import-formula-current-20261029.log`、`F:\RustBuild\audit\audit-core-s1-formula-copy-current-20261029.log`。
- **S1-F1 当前 Release 业务 worker 真实样本复验（2026-10-29）：已完成，待 GUI 人工验收。** `build-release.bat` 使用 `F:\RustBuild\audit` 成功构建并更新 `dist` 主程序与 sidecar；随后通过 Release `acceptance-run` 调用 GUI 使用的同一 `start_dag_flow`/S1-F1 worker，在隔离副本上执行一个真实单位贷款样本。结构匹配 406/406 个标签，1/1 文件通过、跳过 0、失败 0、部分 0；生产 Umya 导入 4 张外部表，处理 1,282 个公式，Excel COM 完成重算。运行器报告 `completed`，3 条公式异常、0 条条件格式异常、总计 3 条；源、模板、映射、外部数据和清单均哈希未变。正式 `汇总核查表校验.xlsx` 已由 Excel 只读打开，`本期审核结果` 为 13 列、4 行。此真实样本没有触发条件格式异常，触发型条件格式仍由既有合成 E2E 覆盖；本次未传 Python 结果基线，也没有执行真实 GUI 点击验收。报告目录：`F:\RustBuild\audit\s1f1-release-acceptance-20261029-065200\output`。
- **S1-F1 Release GUI 真实样本点击验收（2026-10-29）：已通过隔离样本。** 最新 Release `serve` 页面实际选择源目录、模板目录、`mapping.xlsx`、`external.xlsx` 和固定输出目录，刷新后自动匹配 1 个模板并点击 F1。真实源文件结构匹配 406/406；Umya 导入 4 张外部表并准备 1,282 个公式，Excel COM 完成重算；任务 `COMPLETED`，提取 3 条公式异常、0 条条件格式异常，输出 `本期审核结果` 4 行×13 列且由 Excel 只读重开成功。源文件、模板、映射、外部数据 SHA-256 与此前隔离验收基线完全一致。结果位于 `F:\RustBuild\audit\s1-final-release-20261029-070522\gui-output-ui-retest\S1-F1_审核结果_20261029_073237_123\汇总核查表校验.xlsx`。该样本未触发条件格式异常，本次也没有新的 Python 逐格结果基线；触发型条件格式由既有合成 E2E 覆盖。GUI 还修正了 F1 “复制能力未验证”的过时标签、输出跟随/固定状态 API 及确认框的实际输出路径，并防止打开其它输入选择器时误固定默认输出。新增输出模式回归 1 项通过。全工作区回归 80 个目标、484 passed / 0 failed / 32 ignored；日志 `F:\RustBuild\audit\workspace-final-goal-ui-output-20261029.log`。最新 Windows Release 构建成功；日志 `F:\RustBuild\audit\release-final-goal-ui-output-20261029-r2.log`。以上是隔离真实样本的自动化 GUI 点击验收，不等同于用户原始文件和审核规则的业务签收。
- **S1-F3 人工重算恢复：已实现，worker 测试及当前 Release GUI 恢复点击通过；待用户手工重算与业务验收。** F1 无法自动重算时持久化审核副本路径、模板/配置摘要及精确和兼容公式指纹；隐藏的 S1-F3 入口仅在存在待恢复任务时显示。用户在 Excel/WPS/LibreOffice 中强制重算、保存并关闭副本后，F3 检查公式位置/文本、全部缓存节点及每个公式结果是否可由 Calamine 读取，再复用 F1 公式/条件格式提取及结果写入，不重新读源文件或复制公式。公式保护仅容忍 LibreOffice 将 `TRUE`/`FALSE` 写成等价的 `TRUE()`/`FALSE()`；公式位置、引用、区域及其余文本必须一致。恢复状态跨程序重启保留；相同审核结论复用同一输出，输出已发布但状态尚未提交时不会再生成重复结果；审核结论变化会产生新结果文件且不覆盖旧版。2026-10-29 新增 worker 端到端合成回归：重建状态存储后从待处理任务执行 F3、提取 1 条公式异常并发布可重开的 13 列结果，输入哈希不变、任务完成后待处理数归零；另验证缓存缺失时保持 `WAITING_MANUAL_RECALC` 且不发布结果。另新增 Windows 真实样本探针：Umya 生成 1,282 个公式的 S1-F1 独立副本，Excel COM 在分离输出文件中重算后将已保存副本交由新建状态存储的 F3 worker 恢复，成功提取 3 条公式异常，生成可重开的 13 列结果并完成任务，源、模板、外部表和配置哈希不变。本轮用当前 Release 服务做了隔离进程级状态验证：空状态待处理数为 0；单个待恢复任务为 1；结束服务并重启后仍为 1，恢复读取错误为空。该验证覆盖持久任务状态/API，不执行 F3 恢复按钮，也不替代用户在 Excel/WPS/LibreOffice 手工重算、保存、关闭及 GUI 恢复验收。缓存节点只能证明缓存存在；typed 可读性证明值可被消费，仍不单独证明数学正确。
- **S1-F3 Release 跨进程恢复（2026-10-29）：真实工作簿流程通过。** 新的隔离 Release `acceptance-run` 在刻意使用不兼容 sidecar 后，将 1 个真实工作簿保留为 `WAITING_MANUAL_RECALC`，保存 1,282 个公式的审核副本及恢复清单；首次运行进程退出后，由 `lo-convert.exe recalculate --engine com` 使用 Excel COM 重算并保存副本。随后启动全新 Release `serve` 进程，GUI 使用的 `/api/list_s1_f3_tasks` 在进程重启后读到该待处理任务，`/api/start_s1_f3_workflow` 调用正式恢复 worker，完成 3 条公式异常提取、0 条条件格式异常、生成 Excel 可重开的 13 列结果并将队列归零。此项验证了 Release 进程之间的持久化与 GUI API 恢复路径；人工重算由 sidecar 模拟，未模拟用户按键或真实 GUI 点击。材料：`F:\RustBuild\audit\s1f3-process-restart-20261029-065923\`。
- **S1-F3 输入完整性校验补强：** 待恢复任务额外持久化非公式单元格值、条件格式语义、合并区域和命名区域摘要；F3 允许公式缓存更新，并按规则内容/相对优先顺序规范化 LibreOffice 的条件格式样式 ID、默认属性与命名区域序列化顺序。发现审核输入/规则变化时拒绝恢复。旧清单保留原公式校验兼容行为。合成回归覆盖缓存清除不改变摘要、改动普通输入导致摘要变化、恢复校验拒绝变化及 LibreOffice 重存语义兼容；真实 Windows Excel COM 工作簿恢复 worker 探针已通过，真实 LibreOffice 样本及 GUI 手工打开/重算/保存/重启恢复仍待人工验收。
- **S1-F3 恢复副本异常可见性：已实现。** 重启后发现待恢复清单引用的副本丢失，会持久化为 `FAILED` 并在 GUI 明确提示重新运行 S1-F1；存在但损坏的副本会保留原字节，恢复记录为 `RESUME_FAILED`，不发布审核结果并继续出现在可修复重试列表。GUI 的恢复清单读取错误也有独立提示。测试覆盖丢失/损坏副本、状态列表与页面渲染。
- **S1-F4/F5：已明确暂缓。** 不自动开发组合工作表或联合模板。

## Excel 引擎策略及验证状态

- **读取：Calamine 主读值。** OOXML 只补命名区域、条件格式、批注等需要的元数据；未实现通用 OOXML/openpyxl 读值回退。Calamine 报告特定未知 Excel 错误 token 时，才复制原 XLSX/XLSM 到临时包、只将该错误单元格缓存 token 作为文本规范化，再由 Calamine 重读；临时副本自动清理，原文件不改。当前源码 `excel-reader` 19 个单元测试和 2 个集成测试通过，覆盖纯文本/公式错误值、公式文本/普通数据保留、输入哈希、缓存检查、不安全 XML 拒绝及文件缺失/权限/格式/工作表读取/损坏分类。其他损坏、权限、格式、工作表读取错误不走此兼容回退；分类错误信息保留原始错误链。
- **修改：Umya 主引擎 + 能力限定的 OOXML Patch。** Excel-Core 按操作和能力证据选路，只有 Umya 在写入前明确报告不支持，且 OOXML Patch 对同一操作有通过证据时，才允许从干净原始副本回退；部分写入、保存失败、输出校验失败或未知错误均停止。合成路由测试覆盖“Umya 预拒绝→OOXML 干净副本回退”和“两者均不支持→失败关闭”；2026-10-29 对当前源码定向复跑两项均通过。该测试验证路由协议，不把未实现的 OOXML 单元格编辑登记为生产能力。S1-F1 的真实公式 ImportRange 已由 Umya 满足，OOXML Patch 不实现冗余公式复制，故此业务操作不配置伪回退；S1 业务层现于候选落盘后核验公式/引用、工作表顺序和复制的非公式值，再允许进入外部重算。
- **公式计算：Windows PowerShell + Excel COM；Linux LibreOffice。** Windows `auto` 顺序为 Excel COM、再 LibreOffice；Linux `auto` 仅 LibreOffice；人工重算是引擎不可用/失败时的 S1-F3 恢复路径。计算前在独立副本清缓存并请求全量重算；COM 专用副本将 `fullCalcOnLoad` 和 `forceFullCalc` 设为 false，避免 Excel 在两标记同时开启时于 `CalculateFullRebuild` 后仍持续报告 `xlPending`；COM 仍执行全量计算和依赖重建，并要求 `xlDone` 才保存。LibreOffice 路由继续使用两个全量重算标记。保存后核对工作表顺序/名称、公式单元格位置、公式文本（只允许 `TRUE`/`FALSE` 与 `TRUE()`/`FALSE()` 的等价序列化）、所有缓存节点存在，且 Calamine 能为每个公式读取带类型的结果（包括错误值和空字符串）；这证明外部引擎写出可消费的结果缓存，不单独证明数学结果正确，正确性仍由业务预期值/基线和审核规则验证。Windows Release sidecar 的 S1-F1 合成 E2E 和修复后真实样本 E2E 均通过；真实样本中 1,072 个模板公式结果与 Python 审核副本无差异。Linux LibreOffice 实际重算集成覆盖 VLOOKUP、跨表公式、错误值、空字符串和工作表结构。目标机与真实 GUI 验收仍待完成。
- **重算超时进程清理：已实现，Windows/Linux 回归通过。** Windows 超时先对 sidecar 启动的进程执行 `taskkill /T /F`，再直接 kill/wait 父进程；Linux 为引擎建立独立进程组，超时对该组发送 `SIGKILL`，再回收父进程。若等待状态读取出错也走相同清理。新增测试启动真实子进程并断言超时后子进程已退出。2026-10-29 Windows sidecar 全量结果：14 单测通过；集成入口因 Windows 无 `soffice` 按默认策略跳过。Ubuntu WSL：14 单测及实际 LibreOffice CLI 重算集成均通过，后者由 `LO_CONVERT_REQUIRE_LO_TEST=1` 强制启用。Windows Release 已重建；新 `dist/lo-convert.exe` 的 S1-F1 Excel COM 合成端到端回归 1/1 通过。测试仅覆盖进程树清理与既有小型 LibreOffice 样本；目标 Linux 发行版的进程行为仍待设备验收。
- **条件格式及批次隔离：** OOXML 读取条件格式规则，由 Rust 支持的子集求值；不支持/求值失败需带规则、范围和单元格诊断，标记部分结果。2026-10-29 加固了工作簿内部隔离：单个条件格式元数据节点格式错误只跳过该节点并保留同表其他有效规则；单元格值无法安全转换时跳过该表的条件格式判断并记录坐标诊断；单个触发单元格无法转为异常时跳过该异常并继续处理其他触发项。整个条件格式提取阶段意外失败时，该工作簿已提取出的公式异常仍会写入部分结果，文件标为 `Partial` 并记录条件格式失败原因；公式异常提取或重算校验本身失败仍按文件失败处理。批次级已验证独立文件失败不会停止其他文件。该项早期全工作区回归为 80 个测试目标、478 passed / 0 failed / 31 ignored；最新源码总回归见本文件末尾。Release 构建及主程序版本/sidecar 帮助冒烟通过；使用新 Release sidecar 的 Excel COM 合成 S1-F1 端到端为 1 passed，日志 `F:\RustBuild\audit\s1f1-com-e2e-release-partial-condition-20261029.log`。未知条件格式规则仍只形成诊断，不能声称已被完整求值。真实业务文件及 GUI 批次仍待人工验收。

## Excel 引擎与 S1-F3 Goal 完成度核对（2026-10-29）

- **Calamine 读取错误分类与定点兼容：已实现并有测试。** 普通读值仍以 Calamine 为准；只有特定未知 Excel 错误 token 会在原文件的临时 OOXML 副本中规范化后再交 Calamine 读取。损坏、权限、格式、工作表读取等错误不会被扩大为通用重试。`excel-reader` 有 19 个单测和 2 个集成测试覆盖分类、回退边界及原文件保护。
- **工作簿修改路由：已实现；S1-F1 的主路径已有真实工作簿证据。** 生产路由按平台能力证据注册 Umya；ExcelModifier 只在后端写入前明确拒绝且同语义回退能力已验证时才从原始副本切换引擎。S1-F1 真实公式 `ImportRange` 由 Umya 独立完成；约 210 个 VLOOKUP 与跨 Sheet 依赖、公式/非公式内容及输出重开均已检查。OOXML Patch 没有相同语义的公式工作表复制能力，因此该操作没有伪造的 Patch 回退，Umya 失败时失败关闭。S1-F1 生产路由不调用 Python/openpyxl；仓库保留的隔离兼容 worker 不是 Audit 业务逻辑，也不在当前 Windows `dist` 中。
- **公式重算与验证：实现和合成/真实 Windows 样本已通过；Linux 真实业务样本未验。** Windows 通过 PowerShell + Excel COM，Linux 通过 LibreOffice；保存前要求计算完成，保存后核对公式结构、缓存节点和 Calamine typed 结果。Windows COM 已覆盖真实 S1-F1 公式结果及既有 Python 审核副本对拍；2026-10-29 本轮再次在 Ubuntu WSL 实际运行 LibreOffice CLI 集成（15 单测 + 1 集成，进程退出码 0），覆盖 VLOOKUP、跨表引用、公式错误和空字符串结果。Linux 真实机构工作簿与麒麟 V10/UOS V20 设备仍未验证。
- **超时清理、条件格式与批次隔离：实现和自动化回归已通过。** Windows 清理 COM/sidecar 进程树，Linux 清理 LibreOffice 进程组；条件格式不支持项带诊断，局部失败保留可靠公式结果并允许独立工作簿继续。具体边界和失败状态见本节前文及迁移清单。
- **S1-F3 人工重算恢复：代码、worker、跨进程状态和当前 Release GUI 模拟流程已通过；用户亲自重算仍待验收。** 待处理任务持久化，入口按任务状态显隐；恢复时验证公式身份、输入语义和可读缓存，然后复用 F1 的异常提取/结果发布；跨重启、幂等、丢失/损坏副本及混合状态有回归覆盖。此前 Release GUI 恢复操作使用 Excel COM 计算的隔离副本，不能替代用户在 Excel/WPS/LibreOffice 中手工强制重算、保存、关闭后再点击恢复。
- **Release 与 Goal 状态：Windows Release 已构建并当前冒烟通过；总 Goal 仍未完成。** 当前 `dist/open-data-audit.exe --version` 为 `0.1.0`，`dist/lo-convert.exe --help` 含 COM/LibreOffice 重算入口。最新主程序 EXE 时间为 2026-10-29 08:32（晚于嵌入 GUI 资源修改）；sidecar 源码时间为 08:06、Release EXE 时间为 08:12，构建日志为 `F:\RustBuild\audit\release-excel-engine-s1f3-completion-20261029.log`。最终全工作区结果见本文件末尾。尚缺用户提供的真实机构文件/审核结论签收、用户亲自执行的 S1-F3 重算恢复，以及目标 Linux 设备验收；这些缺口不能由合成测试或隔离 Release 点击替代。

## S2 报表采集审核

- **S2-F1：已实现，待人工验收。** 配置加载与校验、报表导入、指标/机构匹配、单位转换、本期与上期比较、外部数据核对、配置规则和结果导出已串入 Rust service、CLI 和 Web API。配置表达式可处理本期/上期数据；“同期比较”应按实际配置规则核验，不能仅凭名称宣称专门的同比流程已实现。
- 历史有真实 XLS 两期及外部数据的 CLI 样本记录，也有 Web API 和合成正常/异常样例；这些不等于当前独立构建的完整人工验收。仍需核对真实机构/指标映射、三类规则及外部异常、输出和原文件保护。

## S3 大集中统计系统：按原 Python GUI 的五项编号

原 Python `config-real/` 的 3.3 报表清单有 43 条非空启用记录；Rust 仓库 `audit-core/config/` 中的 3.3 是空占位，不能替代真实配置验收。3.0–3.3 的工作表和关键字段见 [S3 配置结构核对](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-配置工作簿的实际工作表与关键字段只读核对)。

原 Python GUI 的正式范围为五项：F1 两期数据→比较文件、F2 比较文件→说明文件、F3 机构说明→说明文件、F4 本期数值核对、F5 比较文件→金融表单。此前将 F4/F5 分别压成 Rust 目标 F2/F3、并把 F2/F3 当作额外旧流程的描述有误，现按 Python 实际编号统一。数据流为 F1→F2→F3，F1→F5；F4 独立运行。F1/F4/F5 依赖的 3.0–3.3 配置实际表头及入口、输出结构均以 [V2 需求的 S3 对照表](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-f1-至-f5-原-python-与-rust-业务对照) 为准。

- **S3-F1 两期数据→比较文件：已实现，待人工验收。** Rust `s3_execution_compare`、CLI `s3-execution-compare` 与 Web API `start_central_comparison` 已接通；读取本期/上期数据、3.0 通用配置和 3.1 执行比较配置，执行配对、单位换算、环比警戒、规则动作、五段/八段表达式和软性规则颜色处理，输出 21 个业务列及独立“输出说明”列。没有单独外部数据文件输入；GUI 每期选择一个文件。Linux Debian 10 / glibc 2.28 容器以合成异常样本实跑 CLI：1 行输出、1 条动作规则触发；Windows GUI 和真实机构样本仍待人工验收。
- **S3-F2 比较文件→说明文件：已实现，待人工验收。** Rust `s3_explanation_export`、CLI `s3-export-explanations` 与 Web API `start_central_explanation_export` 已接通。以比较文件中的“输出说明”人工标记筛选行；3.0 的“机构地区参照”用于修正展示名称和按旧规则拼接可区分机构代码，生成八字段“说明文件”及“生成信息”表。现有文件级对拍证据仍待用户按真实机构业务验收。
- **S3-F3 机构说明→说明文件：生产写回链路已打通，待人工验收。** 匹配反馈、备份、暂存写入、失败回滚和目标原地更新已接通 GUI/Web API 与 `s3-import-explanation-feedback` CLI；写回前制作不覆盖的时间戳备份，候选写入失败时恢复原件并保留备份。Windows 和 Linux 能力 harness 均验证目标原地更新、备份与反馈输入哈希保护。已按用户确认对齐 Python 输出契约；runner 若接收用户目标，会先复制到验收输出目录再执行。
- **S3-F4 本期数值核对：已实现，待人工验收。** Rust `s3_current_value_check`、CLI `s3-check-current-values` 与 Web API `start_central_cross_period` 已接通，使用当前期数据、3.0 和 3.2 配置，生成规则核对结果。GUI 实际只提供本期文件；服务/CLI 另有可选上期数据参数，但这条能力没有 GUI 选择入口。CSV/XLSX/XLS 已有局部核对证据，真实边界和 GUI 操作尚待验收；时长单元格目前显式拒绝。
- **S3-F5 比较文件→金融表单：生产编辑链路已打通，CLI 已接通，待人工验收。** `s3_indicator_split` 把规划结果转换为 Excel-Core 编辑计划，按机构写入独立暂存文件，全部成功后发布 `指标比较拆分_时间戳` 目录；GUI/Web API 和 `s3-generate-financial-forms` CLI 已接通。Windows/Linux harness 验证编辑批次和 16 项保留契约；CLI 成功/失败路径及 Release 隔离冒烟已有证据。3.3“分析文件数据单位”按用户确认属于历史遗留字段，继续沿用 3.0 默认目标单位。

Rust 现有 F1、F2、F4 逻辑和入口继续保留；F3、F5 的部分实现也保留并补齐各自剩余链路，不因编号校正重写。完整逐项业务和输出对照见 [V2 需求](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md)，历史实现证据见 [迁移清单](audit-core/MIGRATION_INVENTORY.md)。

编号映射补充：Rust 当前 `s3_explanation_feedback` 是 Python S3-F3；`s3_indicator_split` 及表单编辑操作测试实际对应 Python S3-F5。遗留的 `s3_f3_edit_operations.rs` 文件名和 `AUDIT_S3_F3_*` 环境变量仍带旧编号；该测试函数/说明已按表单编辑操作归入 F5，并在 2026-10-28 补充了测试源码，但未运行，也不构成 F5 生产写入能力或验收证据，更不能作为 F3 业务实现证据。

S3 的辅助配置检查入口已补齐：GUI 显示 3.0–3.3 四个检查按钮，Web bridge 现注册全部四个只读检查 API（`check_central_common_config`、`check_central_expression_config`、`check_central_cross_formulas`、`check_central_forms_config`），并有成功/缺参单元测试。这些按钮仍只是辅助诊断，不作为额外 S3-F 编号或 F1–F5 已完成的证据。

## S5：已从当前产品移除

- S5 导航、前端页面、Web API、CLI、业务模块、验收样本生成代码及 S5 专属 Rust/Python 测试辅助代码已删除。
- S1、S2、S3 保持当前产品范围。既有 S5 报告、人工验收文件、基准结果和修改记录作为历史证据保留，不表示当前产品仍提供 S5。

## 其他范围与交付

- **S4：已移除（2026-11-04）。** 此前无任何实现，仅前端占位页与导航按钮，已按用户要求删除；如恢复需重新立项。保留当前 Rust + HTML GUI；TUI 是可选未来方向，不在本轮实施。
- 项目内置 Excel-Core 源码和独立 Cargo workspace；外部第三方 crates 仍需要相应 Cargo 缓存或正常依赖获取。旧 Python、Main 和 Codex 工作树不是构建或运行路径。可选外部计算引擎/隔离 worker 只用于其明确支持的功能，不能推断所有路径都可在无外部引擎时执行。
- Windows：按 Excel 引擎 Goal 的最终复验再次运行 `build-release.bat`，成功生成主程序与更新后的重算 sidecar 并复制到 `dist`，日志 `F:\RustBuild\audit\release-excel-engine-s1f3-completion-20261029.log`；`open-data-audit.exe --version` 为 0.1.0。新 Release sidecar 的 Excel COM 合成 E2E 1/1 通过；F1 真实 Umya 探针 3/3、F3 Release 真实恢复探针 1/1 通过。ZIP `dist/open-data-audit-windows-x64-20261028-122102.zip` 仍是 2026-10-28 的包，未在本轮重打包。
- Linux：`linux-x86_64` 生产能力证据已生成，未知平台继续失败关闭。最终 musl deb 为 `dist/linux-delivery/open-data-audit_0.1.0_amd64.deb`（13,498,066 字节），在 Debian GLIBC 2.28 / dpkg-deb 1.19.8 等价容器通过安装、静态链接、CLI、Web API、S3-F1/F2/F5 合成 CLI 冒烟；无 `soffice` 时重算侧车明确报错。麒麟 V10/UOS V20 实机尚未验收。
- 验收运行器：在隔离输出中运行 S3-F1→F2 和 S5 合成样本，步骤均完成、所有源文件与配置哈希不变，报告 JSON/Markdown 生成于 `F:\RustBuild\audit\acceptance-runner-synthetic-0b813b7c867b4f979e155fccad36f151`。
- 完整回归：当前源码 `CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -j 1 -- --quiet` 为 **80 个测试目标，489 passed / 0 failed / 32 ignored**，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-final-20261029-r5.log`。忽略项包括真实样本/人工操作及外部运行器，未计为通过。WSL x86_64 LibreOffice 重新计算套件以 `LO_CONVERT_REQUIRE_LO_TEST=1` 为 17 单测和 1 实际 CLI 重算集成通过；目标发行版设备验收仍待完成。
- 公开配置是结构占位，真实配置/样本保持本地忽略。audit-core/manual_acceptance/20261027/ 中大量 Excel 手测样本与输出当前只存在本机、被 Git 忽略；文档和部分数据被跟踪，不等于全套样本随仓库提交。

## 2026-10-29 S1-F3 Release GUI 验收补充

当前源码的 Release 重新构建成功。对隔离的 Excel COM 已重算审核副本执行了当前 Release GUI 完整恢复点击：S1 页显示 1 个待处理任务，打开恢复面板并确认后生成核查表；状态更新为 `COMPLETED`，包含 3 条公式异常、0 条条件格式异常，审核副本 SHA-256 与执行前复制源一致。结果目录：`F:\RustBuild\audit\s1f3-current-release-gui-e2e-20261029\output\S1-F1_审核结果_20261029_073237_123`。

该 Release 测试同时发现并修复了“待处理数为 0 时按钮仍可见”：`.btn` 的 `display` 样式覆盖了 HTML `hidden` 属性。加入 S1-F3 专属隐藏规则及页面断言后，定向页面测试 1 项通过；重新构建的 Release 在待处理数为 1 时显示按钮，恢复后待处理数归零且按钮从无障碍树与实际画面中消失。该 CSS 修复后的最新源码已重新执行全工作区回归：80 个测试目标、485 passed / 0 failed / 32 ignored，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-current-20261029-r4.log`。测试夹具目录隔离加固及首次两个不完整 deflate 错误的排查见迁移清单最新增量。

GUI 点击测试使用的是此前由 Excel COM 重算并保存的隔离副本；没有模拟用户亲自在 Excel/WPS/LibreOffice 按 `Ctrl+Alt+F9`。用户手工重算并对机构业务结果验收、麒麟 V10/UOS V20 设备验收仍待完成。

## 对齐空间与引擎高级设置（2026-10-29）

- **对齐决策清单**：`audit-core/docs/ALIGNMENT_DECISION_LIST_20261029.md`——矩阵 104 个失败场景分三类（96 建议直接采纳 / 6 需用户业务决策 / 2 项 S5 移出）+ 9 条对齐空间（AL-01…AL-09）。待用户决策项见清单第二部分与 AL-04/05/09。
- **引擎高级设置**：设置中心“高级”页新增“Rust 引擎（高级）”卡片。S1-F01 公式重算引擎四通道：auto（Windows=Excel COM→WPS 表格探测；Linux=LibreOffice 无头）/ 强制 Microsoft Excel（COM）/ 强制 WPS 表格（COM，ProgID 探测 ket.Application→KET.Application）/ 强制 LibreOffice 无头（经 lo-convert sidecar）；Excel-Core 修改引擎两通道：Umya（默认，已验证）/ openpyxl（高级兼容路径，证据驱动注册）。选择持久化到状态目录 settings.json 并在日志回显；强制引擎不可用即报错，不自动切换。
- **openpyxl 高级路径**：负面证据（capability_evidence/openpyxl_windows_x86_64_20261029.json，available=false）——openpyxl 重写无法满足三个业务批次的生产保留契约（OOXML 关系、列尺寸、包部件、DeleteColumn），全部批次路由失败关闭；GUI 置灰。未来后端补齐契约后重跑 AUDIT_OPENPYXL_CAPABILITY_REPORT 证据即可解锁。
- **对拍**：S3-F02 合成 Python/Rust 对拍通过（probe/reports/s3_f02_synthetic_parity_20261029.json，4/4 行一致、differences=[]、9 项检查全过）。S1-F02 对拍器夹具落后于当前表结构预检（0/3 被跳过），为对拍器维护项。S3-F1 数值口径评估：旧 Python 基准即 IEEE-754，不建议 Decimal 化（AL-04，待确认）。
- **编号澄清**：Rust 内部“S1-F3”是 S1-F01 人工重算恢复 worker（s1_recovery.rs），与 Python 菜单暂缓的 S1-F3（组合工作表）撞名；是否重命名待用户确认（AL-05）。
- S1-F01 校验区域扫描双口径（CLI 源缓存 / GUI 重算副本）已由 tests/s1_scan_scope_contract.rs 固定，统一口径待用户决定（AL-03）。此前误标为 S1-F02。

## 下一步

本轮已按用户授权完成 COM `xlPending` 定点修复、公式结果可读性/规范化校验、S1-F1 Umya 空白覆盖修复、复制后候选公式/输入值校验、真实旧 Python 工作簿复验、Release GUI 输出目录状态修复、全工作区回归及最新 Windows Release 构建；隔离真实样本的 GUI 点击、Excel COM 重算和 13 列结果重开已通过。另在当前 Release 上完成 S1-F3 GUI 任务选择、确认恢复、结果写出及完成态刷新；修复完成任务后按钮仍可见的问题。这些结果不替代用户机构文件/规则的业务验收。下一步是用户选定文件的人工验收：

1. S1-F1 以用户选定的机构源文件、模板、外部数据和配置在 GUI 执行审核；核对模板匹配/跳过原因、13 列结果与实际公式/条件格式异常，并确认源文件不变。隔离真实样本 GUI 路径已经过验证。
2. S1-F3 的 Release GUI 恢复点击流程已在隔离的已重算副本上通过：任务显示→选择→确认→结果写出→待处理数归零；仍待用户实际使用 Excel/WPS/LibreOffice 对由 S1-F1 保留的副本强制重算、保存并关闭，再通过 GUI 恢复，并检查真实业务结论、重启持久化和重复恢复幂等。
3. S3-F1 至 F5 使用真实 3.0–3.3 配置与业务数据验收菜单操作、输入/输出结构及 F1→F2→F3、F1→F5 数据流；F3 核对“先备份、再原地更新”，F5 按已确认口径沿用 3.0 默认目标单位。S2 使用真实配置和输入文件核对结果及原文件保护。
4. 在麒麟 V10/UOS V20 x86_64 目标机安装 deb 并复核 GUI、配置发现、S1 缺少 Office 引擎时的提示及业务入口。

## 2026-11-07 增量：WPS 显示色能力探测与 95 条差异归因更正

- 对上一轮 Excel COM 打不开的 7 本“单位贷款置换旧债发生额”源文件，用 WPS COM 做了真实探测。WPS 虽能打开/重算，但 `DisplayFormat` 填充色读取没有提供可用颜色：每本 116 个目标格均产生诊断、颜色变化 0；整批 7 本均为 `PARTIAL`。这次不能作为条件格式验收通过，且运行中 WPS 自动化进程没有随批次全部退出；已按本次运行时间、`/Automation -Embedding` 命令行和进程归属确认并清理 6 个本次测试遗留进程。7 本输入哈希均未变化。
- 因此 WPS 不再作为条件格式显示色真值来源。S1-F1 仅接受无诊断的 Excel COM `DisplayFormat` 扫描；WPS 结果或不完整 Excel 扫描转入现有静态规则求值，并明确标记部分完成，防止把“颜色不可读”当成零异常。新增可信度门控的 3 个单测通过；条件格式扫描区域路由测试 3/3 通过；`cargo check` 和 Windows Release 构建通过。Release 位于 `F:\RustBuild\audit\release\open-data-audit.exe`。
- 没有用本次 WPS 结果覆盖之前的 129 本基准，也没有重跑全量。此前 run6 的有效 Excel 结果仍是：122 成功、7 部分；Python-only 的 8 个条件格式位置均为四家机构的 `单位贷款置换旧债发生额统计表!C36/D36`，Rust-only 为按用户要求保留的空白红色 `单位存款发生额分产品期限!AX10`。8 条尚未完成真实显示色验收，S1-F1 条件格式未完全对齐。
- 95 条公式差异的 `IndexError` 归因需要更正：用 `81505fb` 未修改基线和正确区域映射重跑 13 份副本时，95/95 均可提取，没有复现 `IndexError`；此前一次缺少区域映射的尝试无效。当前没有证据将这 95 条归因于内存不足或某一已确认的越界。它们仍是原 Python 整批输出与 Rust 输出的差异，但根因未定，不能宣称已修复或完全对齐。
- 截图中的 Excel 修复日志 `C:\Users\AZI\AppData\Local\Temp\error345880_02.xml` 指向原始输入 `测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx`，内容是 Excel 从 `sheet1.xml` 删除公式。文件含 86 个 I5:I90 公式；同一批 7 家机构该表公式文本哈希相同。此文件 SHA-256 为 `618f7e195e1a52bae03dc19c985402aea12549904a5feddaaa67c3a296da3c70`，与 129 本 run6 验收的运行前/后哈希相同，文件最后修改时间为 2026-10-19。证据证明 11-07 这次 Rust 批次没有改动输入；修复日志没有指出被删除的具体单元格，具体公式问题及其更早来源仍未定位。
- 本次没有运行全工作区测试、GUI 手动验收、Kylin/UOS 设备验收；未提交、未 push。完整 WPS 试验及失败原因见 `audit-core/probe/reports/s1f1_real_sample_parity_20261106.md` 的 2026-11-07 补充记录。

### 2026-11-07 续验：条件格式静态求值边界

- 再次核对并恢复了 S1-F1 原生 COM worker 的 `scan_conditional_display_colors` 请求分支。Windows 条件格式真值仍取隔离 Excel COM 的 `Interior.Color` 与 `DisplayFormat.Interior.Color` 差异；扫描范围优先来自所选工作簿的命名“条件格式区域”，无匹配命名区域时只读 OOXML `sqref`。没有条件格式规则时才返回零扫描。
- 试验将空白目标也交给静态规则求值后，7 本样本得到 26 个候选位置。旧 Python run4 的 batch16 实际使用 WPS COM，输出 8 项（四家机构各 C36/D36）；直接调用旧 Python OOXML 求值器则同样得到 26 项。两者说明纯规则求值与渲染颜色结果有分歧，但没有可用 Excel COM 观察，不能把多出的 18 项断言为误报或真实红格。Rust 的空白目标扩展已撤回；静态回退继续保守跳过空白并标记部分完成。
- 当前边界：Excel COM 无诊断时按显示填充色变化判定（含空白红格 AX10）；COM 不可用或扫描不完整时保留静态结果并标记部分完成，不能把 WPS 或未验证的静态空白推断当作显示色基准。7 本工作簿的 Excel 打开失败和其中 8 个 Python-only 条件格式位置仍待解决；没有重新运行 129 本批次。
- 按照旧 Python WPS 的刷新顺序（激活表后 `CalculateFullRebuild`）做了一次 Rust worker 单文件试验；90 秒无响应，停止试验并清理了该试验启动的隐藏 WPS 进程。此改动已从源码撤回，WPS 不纳入颜色判定。
- 验证：定向显示色门控测试 3/3、扫描范围测试 3/3、空白静态回退测试 1/1 通过；回退 WPS 试验补丁后的 `cargo check --offline --locked -p audit-core --bin open-data-audit --quiet` 和 Windows Release 离线构建通过。没有跑全工作区测试或 GUI 手动验收；未提交、未 push。
- 旧 Python run4 后续已恢复完成：19/19 批次成功、129/129 文件成功、3,345 条公式异常及 264 条条件格式异常，共 3,609 条；输入哈希变化为空。批次耗时相加 934,233 ms，恢复后的墙钟耗时 513,250 ms。该结果确认旧 Python 运行已结束，但不消除 Rust 7 本 Excel COM 打开失败造成的颜色基准缺口。
- 用刚构建的 Release EXE 直接做原生 COM worker 冒烟：只读打开广东惠东单位存款源文件、扫描 `单位存款发生额分产品期限!AX10`，返回基础填充色 `16777215`、显示填充色 `255`，检测到 1 个变化、0 诊断；专用 Excel PID `34256` 正常退出，源文件 SHA-256 与 run6 验收记录一致。该测试验证 worker 通道与空白红格识别，不替代 129 本复验。

### 2026-11-07 续验：COM 红色基准与静态求值器语义修正

- Windows 真值以候选区域内 Excel `Interior.Color` 与 `DisplayFormat.Interior.Color` 的实际差异为准；优先使用所选工作簿的“条件格式区域”命名映射，没有映射时才使用该工作簿 OOXML 的条件格式 `sqref`。不会扩大到整张表；空白但实际显示为红色的单元格也计入。
- 在 `audit-core/src/conditional_format.rs` 修正离线静态求值：空白目标不再跳过；空白引用参与算术时按 Excel 规则转为 0；除数为 0 产生 `#DIV/0!` 错误值并沿比较、`AND`/`OR`/`SUM` 传播；补充 `IF`、`IFERROR`、`ISBLANK` 基础求值。错误条件不触发格式，`IFERROR` 可按规则处理错误。Audit 生产逻辑保持纯 Rust，没有新增 Python/LibreOffice 运行依赖。静态求值读取保存的单元格值，不执行公式重算；缓存过期时不能等同于 Excel 当前画面。
- 广东惠东单位存款真实源文件 COM 复核：独立 Excel 实例只读扫描 OOXML `sqref` 的 2,701 格，47 格基础填充 `16777215` 变为显示填充 `255`，47 格全部为红色；包含空白的 `单位存款发生额分产品期限!AX10`。输入 SHA-256 前后不变。当前 Release 静态求值器对同一源文件也得到 47/47，无诊断，触发位置一致。
- 对 7 本“单位贷款置换旧债发生额”源文件和 run6 生成的 7 本审核副本逐本尝试 Excel COM 普通只读及修复模式打开，均返回 `Workbooks.Open` HRESULT `0x80020009`；将一份源文件复制到 ASCII 临时路径仍失败，所有输入哈希未变。静态读取可解析 7 条规则并覆盖 116 格，规则 7/7 受支持；按 Excel 除零语义重算为 0 条。Python run4 的 WPS COM 为 8 条，旧 Python OOXML 静态结果为 26 条。没有可用 Excel 基准，0/8/26 不能相互验证或视作最终业务结论；这 7 本仍为部分完成。
- 条件格式相关库单测 28/28 通过，其中包含空白目标命中、`0/0` 错误不触发及 `IFERROR` 恢复；`cargo check --offline --locked -p audit-core --bin open-data-audit --quiet` 和 Windows Release 离线构建通过。只跑条件格式定向测试和样本复核，未重跑 129 本全量或 GUI 手动验收。
- Python COM 探针仅用于取得 Excel 显示色基准；Audit 生产条件格式逻辑在 Rust 实现。没有修改源文件，没有提交或 push；工作区其他并行改动保留。

## 2026-10-29 增量：真实 S1-F1 LibreOffice 兼容与 Release 复验

真实 Linux S1-F1 探针首次运行时，在外部工作簿基线重算阶段失败，尚未进入 Umya 模板复制。对隔离副本直接执行 LibreOffice 并比较 OOXML 后确认：`参照表!D2:D211` 共 210 条 VLOOKUP 公式文本完全未变，LibreOffice 仅添加 `aca="false"` 序列化属性；原本为固定错误值的 `参照表!C211` 则被 LibreOffice 写成缓存值相同的 `#REF!` 公式。没有放宽通用公式校验。

现新增能力受限的 OOXML Patch `restore_promoted_error_constants`：只接受源工作表中确为 `t="e"` 的错误常量，并且计算输出在同一坐标仍为相同的受支持错误 token、公式属性仅为可选的 `aca="false"`、缓存错误值完全相同；仅从新的候选副本移除该单元格的公式节点并保留错误值。sidecar 随后仍要求工作表、所有原公式位置/文本、结果缓存节点和 Calamine 类型化结果严格通过现有校验。新公式、被改写公式、引用变化、错误值不同或公式属性异常均失败关闭。该补丁不是跨工作簿公式复制引擎；S1-F1 ImportRange 仍由 Umya 单独完成，不配置未经验证的 OOXML 公式复制回退。

最终证据：`ooxml-patch` 定向单测 2/2 通过；Windows `lo-convert` 重算单测 17/17 通过；Ubuntu WSL 以 `LO_CONVERT_REQUIRE_LO_TEST=1` 执行 `cargo test -p lo-convert -- --quiet`，17 单测和 1 个真实 LibreOffice CLI 集成通过。集成样本新增固定错误常量，确认输出仍保留原错误值且不变成公式。使用 WSL LibreOffice 对隔离真实 S1-F1 工作簿执行完整探针 1/1 通过：覆盖纯数据、普通/跨表公式、约 210 个 VLOOKUP、共同复制「参照表」及「机构参照表」、模板公式复制、重算后值比较、输出重开和原输入保护。Excel-Core 路由回退、双引擎不支持失败关闭、失败后从原副本重试三项定向测试均通过。随后 `build-release.bat` 成功更新 Windows `dist`，`--version`/sidecar `--help` 冒烟通过；最新 Release sidecar 经 Excel COM 对同一真实 S1-F1 探针 1/1 通过。最新 Windows 全工作区回归为 80 个测试目标、489 passed / 0 failed / 32 ignored，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-final-20261029-r5.log`。其他日志：`F:\RustBuild\audit\s1-real-linux-lo-probe-20261029-r3.log`、`F:\RustBuild\audit\lo-convert-wsl-final-20261029.log`、`F:\RustBuild\audit\s1-real-release-com-error-restore-20261029.log`、`F:\RustBuild\audit\release-s1f1-error-constant-normalization-20261029.log`。三个修改源文件的 `rustfmt --check` 通过。用户机构业务签收、S1-F3 用户手工重算恢复和 Kylin/UOS 目标设备验收仍待完成；总 Goal 保持 active。

## 2026-10-29 增量：当前源码真实样本 S1-F3 worker 复验

当前环境没有预设 `AUDIT_S1_REAL_*` 路径，因此全工作区测试中的真实 S1-F3 案例按要求保持 ignored。随后从本机既有隔离样本目录显式提供只读源文件、模板、外部数据、映射及当前 `dist/lo-convert.exe`，在当前 Release 测试二进制中手动启用该测试，1/1 通过：当前 Umya 生成 1,282 公式的审核副本，Excel COM 将计算结果保存至分离副本，再恢复为原审核副本；新建恢复 Store 后执行 S1-F3 worker，状态为 `COMPLETED`，提取 3 条公式异常、0 条条件格式异常、0 条诊断，生成含预期 13 列标题的结果，源文件/模板/外部数据/映射哈希未变化。日志 `F:\RustBuild\audit\s1-f3-real-worker-current-20261029.log`。此后再次运行当前 `dist` 冒烟：主程序 `--version` 返回 `0.1.0`、sidecar `--help` 显示 `recalculate --engine auto|lo|com`，二者退出码均为 0；主程序/sidecar 分别为 09:46:26/09:46:50 构建，晚于当前相关源码修改时间，无需重复 Release 构建。这证明当前真实工作簿 worker 与重启后恢复路径可工作，但 Excel COM 仍是对人工重算步骤的自动模拟，不等于用户本人在 Office 按 `Ctrl+Alt+F9`、保存关闭并从 GUI 恢复；真实机构业务结论签收和 Kylin/UOS 设备验收继续待办，Goal 保持 active。
