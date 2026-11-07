# Rust-Audit：GUI 运行日志系统 + 当前处理对象实时展示（收尾）

正式工作目录：`C:\Users\AZI\Desktop\VibeCode\Rust-Audit`
分支 `main`，接手时 HEAD `5c74fbb`

> 本文件是交接目标。前一轮会话额度耗尽中断，工作区有未提交内容与一个未合并索引项，
> 必须先按「第零步」恢复干净状态再继续，不要重做已完成部分。

## 第零步：恢复工作区（必做，5 分钟）

当前 Git 状态（已核实）：

```
HEAD = 5c74fbb
git status:
  UU  audit-core/src/web_bridge.rs          ← 索引未合并（工作区已解决冲突，只是没 git add）
  M   audit-core/frontend/web/index.html    ← 已暂存（日志 UI 改造）
  M   DEVELOPMENT_GOAL.md / audit-core/AGENTS.md / MIGRATION_INVENTORY.md
  M   PERFORMANCE_BASELINE.md / PERFORMANCE_HISTORY.md
  M   audit-core/src/{app_settings,period_pairs,s1_formula_copy,s1_recalculation,
      s1_region_summary,s2_import,s2_service}.rs
  M   audit-core/../rules/13_*.md
  M   sidecar/lo-convert/src/{lo_batch,main}.rs
  M   sidecar/lo-convert/tests/recalculate_libreoffice.rs
  stash@{0}: concurrent-wip-round2
```

**关键背景**：上述 17 个「并发会话改动」来自**另一个 Codex 会话**（在跑 S1 Linux LO 批次复用 / 性能基线相关工作），不属于日志任务。
- `web_bridge.rs` 的 `UU` 是前一轮 `git stash pop` 产生：**冲突已在工作区手工解决**（`grep -c "^<<<<<<<"` = 0），`cargo check -p audit-core --bin open-data-audit` 通过，只是索引没标记解决。
- 不要 `git checkout`/`git reset` 任何这些文件——会丢并发会话的工作。

恢复动作：
1. `git add audit-core/src/web_bridge.rs` 标记冲突已解决（工作区内容是正确的，已编译验证）
2. 确认 `cargo check -p audit-core --bin open-data-audit` 通过
3. 若 `git stash list` 仍有 `concurrent-wip-round2`，其内容已 pop 进工作区，确认无重复后可 drop
4. **不要提交那 17 个并发文件**——它们属于另一个会话

## 一、已完成的日志任务（已提交，勿重做）

### 提交 `4520f77`：分级 + 时间戳日志 + 前端筛选/保存/清空

- 后端 `audit-core/src/web_bridge.rs`：
  - `LOG_CAPACITY = 2000`（原 200），`LOG_CAPACITY` 满时 `drain`
  - `log_timestamp()` → `HH:MM:SS`
  - 统一入口 `push_log_entry(state, message, level, detail)`；条目结构 `{text, detail, level, time}`
  - `push_log`（info）/ `push_detail_log`（info+detail）/ `push_warn_log` / `push_error_log`
  - 11 条 worker 失败行升级 error（S1-F01/F3/F02、S3-F01/F02/F04/F05 启动失败与业务失败）
  - S1-F01 逐文件结果按状态分级：`failed`→error、`skipped*`→warn、其余 info
  - S1-F01 开始日志带「N 通过 / M 跳过」，status 带总数与阶段说明；完成摘要带总耗时秒
  - `cancel_execution` 补齐（原为幽灵 API，前端调用但后端无 dispatch）：worker 无取消检查点，故如实返回「不支持中途终止」并记一条 warn，不伪造取消
- 前端 `audit-core/frontend/web/index.html`：
  - 日志容器 `textarea#log` → `div#log`（每行一个 `.log-line`）
  - ERROR 红/加粗、WARN 琥珀、时间前缀灰（`.log-time`）
  - 级别筛选下拉 `#logLevelFilter`（全部/警告及以上/仅错误）
  - 「保存日志」（`saveLogToFile`，Blob 下载 .txt）、「清空显示」（`clearLogView`）
  - `copyLog` 适配 div（`innerText` + 隐藏 `#logCopyBuffer` 兜底）
  - 签名未变跳过重建、底部跟随滚动保留

### 提交 `5c74fbb`：currentWork 当前处理对象状态条

- 状态新增 `currentWork` 槽（初始 `Value::Null`），与历史日志分离——正常文件/表切换只更新状态不写日志
- 后端 helpers：`CurrentWorkSlot = Arc<Mutex<Value>>`、`current_work_update`、`current_work_publish`、`current_work_finish`
- S1-F01：`execute_s1_f1_workflow(request, accepted, work_slot, state)` 签名扩展，闭包内每文件开始写 `{workbook: 文件名, stage: "复制公式", index: n}` 并 publish；Ok 分支写 `{stage:"完成", done, failed, partial, waitingManualRecalc}` 后 finish；Err 分支写 `{stage:"失败", error: 截断 500 字}` 后 finish
- S2-F01：`on_progress` 里写 `{feature:"S2-F01", stage: event.label, stageKey: event.stage, index: current, total}`；结束 `{stage:"完成"}`
- S3-F01：步骤回调写 `{feature:"S3-F01", workbook: 文件名, stage: step}`；完成写 `{stage:"完成", rows}`
- S3-F02/F04/F05：任务开始写 `{feature, stage}`（单文件 worker，不虚构 sheet）
- 前端：状态点旁 pill `#currentWork`，渲染 `功能 · 阶段 · 工作簿(截断,title 全名) · 工作表 · n/total · 成功/失败计数`；空闲隐藏；签名比对避免重复 DOM 写

**已知未覆盖**：`currentWork.sheet` 字段前端已支持渲染，但**没有任何功能写入真实工作表名**——因为执行链路上没有现成的「当前 Sheet」事件源。这是本任务的剩余核心项。

## 二、剩余工作（按优先级）

### P1. 修复随机失败的测试（必须先做，否则无法通过提交门禁）

现状：`cargo test -p audit-core --lib` 有 1 个失败，**每次失败的具体测试不同**，隔离单跑必过。根因已定位到 `AUDIT_STATE_DIR` 是**进程级环境变量**：

- `audit-core/src/app_settings.rs` 的测试与 `web_bridge.rs` 的多个测试都读写它
- 自建目录的测试（`advanced_settings_update_persists_and_logs_selection`）结束时 `remove_dir_all` 自己的目录，而并行测试此时正因 env 指向该目录而写入 → `os error 3/5`
- 前一轮已加 `AUDIT_STATE_ENV_LOCK: Mutex<()>` 并让两个测试持锁，但**锁只覆盖了 2 个测试**，其他触碰白名单键的测试（如 `s1_output_directory_tracks_input_until_explicitly_pinned`、`choosing_and_clearing_period_directories_refreshes_pair_preview`）也会触发持久化

要求：
1. 找出所有会触发 `persist_state_preferences` 的测试（凡 dispatch `update`/`choose_*` 且带白名单键的都算）
2. 用同一把锁串行化，或改用不删目录的策略（保留目录、进程退出时由系统清理）
3. 验证标准：`cargo test -p audit-core --lib` **连续 3 轮 284/284 全绿**

### P2. 补齐 currentWork 的 sheet 与更细阶段（本任务核心剩余项）

用户原始要求（补充目标）：
> 用户在批量运行时，能够实时知道正在处理哪个工作簿、哪张工作表、处于哪个执行阶段。
> 若当前没有具体工作表，不要虚构名称。

现状与缺口：

| 功能 | workbook | stage | sheet |
|---|---|---|---|
| S1-F01 | ✅ 有 | ✅ 「复制公式」/「完成」/「失败」 | ❌ 无 |
| S2-F01 | ❌ 无（需从 request 取目录/文件名） | ✅ event.label | ❌ 无 |
| S3-F01 | ✅ 有 | ✅ step | ❌ 无 |
| S3-F02/F04/F05 | ❌ 无 | ✅ 粗粒度 | ❌ 无 |

实现建议（只读现有数据，不新增 Excel 读取）：

1. **S1-F01 的 sheet**：`s1_formula_copy.rs` 构造 `ExcelEditPlan` 时已有 `EditOperation::ImportRange { target_sheet, .. }` / `ImportRangeWithFormulaTranslation { target_sheet }`；应在执行编辑前把**当前正在写入的目标表名**推给 currentWork。可选实现：
   - 在 `execute_s1_formula_copy_with_manual_recovery_session` 前，从 plan 的 operations 里取第一个 `ImportRange*` 的 `target_sheet`（这是真实将写入的表，不是预估）
   - 或在 `ExcelModifier::execute` 内部加一个 `on_operation` 回调（改动面较大，需评估）
   - **不允许**用 `read_xlsx_sheet_names` 之类的额外读取凑数据
2. **S2 的 workbook**：`S2ProgressEvent` 只有 stage/label/current/total，没有文件名。需评估 `s2_service.rs` 是否能在事件里带当前文件（若数据结构改动会波及测试，先评估再动）；**若无法在不新增读取的前提下拿到，就明确不显示**，不要编造
3. **S3 的 sheet**：S3-F01 读取的是 CSV/XLSX 混合，`run_execution_compare_files` 内部可能有 sheet 遍历；先 grep 确认有无现成变量，无则不显示
4. 阶段细化：S1-F01 目前只有「复制公式」，重算与异常提取阶段没有独立事件源；评估 `execute_s1_f1_workflow_with_dependencies` 是否有可挂钩的真实节点

### P3. 真实验收（用户明确要求）

用户要求「提供实际 GUI 截图、代表性异常日志和测试结果」。前一轮只做了 API 层的冒烟（`get_state` 里 `currentWork: null`、S3-F04 前置校验失败不写状态），**没有真实业务运行的截图**。

需完成：
1. 用测试副本跑 S1 批量（≥2 个文件）、S2、S3-F04 各一次
2. 截图：任务进行中的状态条（显示功能/阶段/工作簿/n-total）、日志窗口（含 ERROR 红色行）
3. 记录代表性异常日志：文件读取失败、工作表处理失败、COM/LibreOffice 计算失败各一条，确认含**文件名 + 阶段 + 原因**
4. 记录任务结束后 `currentWork` 被清空（不残留"正在处理"）

### P4. 提交

- 只提交日志任务的 2 个文件：`audit-core/src/web_bridge.rs`、`audit-core/frontend/web/index.html`
- **不要**提交那 17 个并发会话文件
- 提交前必须：`cargo check` 通过 + `cargo test -p audit-core --lib` 连续 3 轮 0 failed
- **禁止 push**（本轮未获授权）

## 三、约束

- 只改日志与状态展示，不改 S1/S2/S3 业务算法、不改 Excel 引擎选型、不新增 Excel 读取/计算/保存
- 不新增第三方依赖
- 优先复用现有 `web_bridge` 状态结构与 800ms 轮询，不引入 WebSocket/SSE
- 保留现有日志级别/筛选/保存/复制/环形缓冲能力
- 保留 `push_log` / `push_detail_log` 现有调用点语义（detail 行默认隐藏）
- 不伪造进度、不伪造 sheet 名、不把失败显示为成功
- 不删测试、不放宽断言
- 不碰 rustTools、旧 Python 项目、S5

## 四、完成标准

1. `cargo test -p audit-core --lib` 连续 3 轮 284/284，0 failed
2. `cargo check -p audit-core --bin open-data-audit` 通过
3. S1 批量运行时 `currentWork` 随真实执行更新（workbook + 真实 stage + n/total）
4. 有真实 Sheet 名的功能能显示；拿不到的功能如实不显示
5. 失败日志含文件名 + 阶段 + 原因（代表性样本 3 条）
6. 任务结束/失败后 `currentWork` 清空，无"正在处理"残留
7. 打包产物 GUI 冒烟通过（`dist/OpenDataAudit/`，双击即起 GUI）
8. 提交仅含日志任务的 2 个文件，不含并发会话文件
9. 最终报告如实列出：已完成 / 未完成及原因 / 未跑的验收 / 环境阻碍

## 五、快速定位索引（避免重复调研）

```
日志基础设施        audit-core/src/web_bridge.rs  push_log_entry / LOG_CAPACITY / push_error_log
currentWork helpers audit-core/src/web_bridge.rs  CurrentWorkSlot / current_work_update / _publish / _finish
S1 worker           audit-core/src/web_bridge.rs  fn run_s1_f1_worker (~L2689)
S1 执行+闭包        audit-core/src/web_bridge.rs  fn execute_s1_f1_workflow (~L3543)
S2 worker           audit-core/src/web_bridge.rs  fn run_s2_worker (~L6564) + on_progress
S3-F01 worker       audit-core/src/web_bridge.rs  fn run_s3_execution_compare_worker (~L4973)
S3-F02/F04/F05      audit-core/src/web_bridge.rs  ~L5211 / ~L6140 / ~L5672
Excel 编辑计划      crates/excel-editor/src/plan.rs  ExcelEditPlan / EditOperation::ImportRange{target_sheet}
S1 公式复制+plan    audit-core/src/s1_formula_copy.rs (~L749 target_sheet: sheet.sheet_name)
前端日志渲染        audit-core/frontend/web/index.html  renderLog / renderCurrentWork / saveLogToFile
前端状态条          audit-core/frontend/web/index.html  #currentWork / .current-work
设置持久化          audit-core/src/app_settings.rs  AppSettings::write / PREFERENCE_SPECS
```

## 六、验收证据现状（诚实记录）

| 项 | 状态 |
|---|---|
| `cargo check`（日志任务基线） | ✅ 通过 |
| `cargo test -p audit-core --lib` | ⚠️ 283/284，1 个随机失败（见 P1） |
| 打包 exe 含新代码 | ✅ 二进制字符串核验（currentWork ×5、renderCurrentWork ×2） |
| 页面新元素 | ✅ 7 处引用齐全（logLevelFilter/saveLogToFile/clearLogView/cw-file…） |
| API 层 currentWork 冒烟 | ✅ 初始 null；前置校验失败不写状态；无 busy 残留 |
| 真实业务运行截图 | ❌ 未做（P3） |
| sheet 级展示 | ❌ 未实现（P2） |
| 真实异常日志样本 | ❌ 未收集（P3） |
