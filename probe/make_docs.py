"""生成两份 docx：
1. 使用说明（面向小白，只讲界面可见内容）→ release 内
2. 功能架构表（开发者/AI 用，不入 release）
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit")


def style_doc(doc: Document, title: str) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "微软雅黑"
    normal.font.size = Pt(10.5)
    heading = doc.styles["Heading 1"]
    heading.font.name = "微软雅黑"
    heading.font.size = Pt(16)
    heading.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    doc.add_heading(title, level=0)


def para(doc: Document, text: str, *, bold: bool = False, size: float | None = None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    return p


def bullet(doc: Document, text: str):
    doc.add_paragraph(text, style="List Bullet")


def numbered(doc: Document, text: str):
    doc.add_paragraph(text, style="List Number")


def table(doc: Document, headers: list[str], rows: list[list[str]]):
    grid = doc.add_table(rows=1, cols=len(headers))
    grid.style = "Light Grid Accent 1"
    for index, header in enumerate(headers):
        cell = grid.rows[0].cells[index]
        cell.text = header
        for run in cell.paragraphs[0].runs:
            run.bold = True
    for row in rows:
        cells = grid.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = value
    return grid


# ---------------------------------------------------------------- 使用说明

def build_user_guide(path: Path, edition: str) -> None:
    doc = Document()
    style_doc(doc, f"Open-Data-Audit 数据审核工具 使用说明（{edition}）")

    para(doc, "本说明面向日常使用者，只介绍软件界面上能直接看到的操作。"
              "不需要安装 Excel、WPS、Python 或其它软件即可运行。"
              "原始报送文件在整个过程中不会被修改。", size=10)

    doc.add_heading("一、启动与界面", level=1)
    numbered(doc, "解压全部文件到一个文件夹（不要只解压 exe 出来，配置文件要一起）。")
    numbered(doc, "双击 run.bat（或直接双击 open-data-audit.exe）。")
    numbered(doc, "浏览器会自动打开操作界面。界面左侧是三个系统入口，右侧是当前系统的操作区。")

    doc.add_heading("二、三个系统分别做什么", level=1)
    table(doc, ["系统", "用途", "典型任务"], [
        ["逐笔统计系统", "审核单家机构的逐笔明细表", "汇总核查表校验、汇总校验结果说明"],
        ["报表采集系统", "对两期报表做环比、校验与外部核对", "执行报表规则"],
        ["大集中统计系统", "比较两期导数并生成金融表单", "两期数据→比较文件、说明文件、金融表单"],
    ])

    doc.add_heading("三、逐笔统计系统", level=1)

    doc.add_heading("3.1 汇总核查表校验", level=2)
    para(doc, "适用：审核同一种基础数据报表，生成本期审核结果和各机构审核副本。", bold=True)
    para(doc, "源数据获取：逐笔统计系统 → 待办 → 筛选对应任务 → 使用「批量导出核查表」取得源文件。")
    para(doc, "操作步骤：")
    numbered(doc, "选择「源数据目录」——存放源报送文件的文件夹。")
    numbered(doc, "首次选择后，软件会自动识别模板；识别不对时手动选择。")
    numbered(doc, "在「待处理文件」列表里取消本次不需要处理的文件。")
    numbered(doc, "如需把外部数据表并入，点击「追加文件」选择外部工作簿。")
    numbered(doc, "点击「F1 汇总核查表校验」。")
    para(doc, "输出：执行结果目录下的「本期审核结果」工作簿；必要时保留审核副本。")
    para(doc, "提示：原始报送文件不会被修改；文件都在源目录里时不必勾选「递归子文件夹」。")

    doc.add_heading("3.2 汇总校验结果说明", level=2)
    para(doc, "适用：汇总机构报送的任务说明、业务说明、本地校验结果。", bold=True)
    para(doc, "源数据获取：逐笔统计系统 → 待办 → 使用「批量导出校验结果说明」。")
    para(doc, "操作步骤：")
    numbered(doc, "选择包含校验结果与说明文件的源数据目录。")
    numbered(doc, "机构文件放在子文件夹时，勾选「递归子文件夹」。")
    numbered(doc, "核对待处理文件清单。")
    numbered(doc, "点击「F2 汇总校验结果说明」。")
    para(doc, "输出：「校验结果与报送说明汇总」工作簿。")

    doc.add_heading("四、报表采集系统", level=1)
    para(doc, "适用：对两期报送数据执行环比、校验规则和外部核对，生成三张分类审核表。", bold=True)
    para(doc, "源数据获取：金融基础数据报表采集 → 业务处理 → 报表数据查看 → "
              "勾选机构和 20201/20202/20203 报表 → 使用默认「单 Sheet」导出。")
    para(doc, "操作步骤：")
    numbered(doc, "打开「报表采集系统」。")
    numbered(doc, "选择「本期目录」「上期目录」和输出目录。")
    numbered(doc, "点击「检查配置」，确认没有阻断问题。")
    numbered(doc, "需要大集中核对时，再选择「大集中数据」与「比较配置」。")
    numbered(doc, "点击「执行报表规则」。")
    para(doc, "输出：「报表审核清单_时间戳.xlsx」，含「环比规则」「外部核对规则」「校验规则」三张表。")

    doc.add_heading("五、大集中统计系统", level=1)

    doc.add_heading("5.1 两期数据 → 比较文件", level=2)
    para(doc, "适用：比较两期大集中导数，按配置执行环比、警戒和指标校验。", bold=True)
    para(doc, "源数据获取：大集中系统 → 指标查询 → 选择机构及所有指标 → "
              "分别导出本期和上期数据（不要混在同一个文件里）。")
    para(doc, "操作步骤：")
    numbered(doc, "打开「大集中统计系统」。")
    numbered(doc, "选择「本期数据」「上期数据」和输出目录。")
    numbered(doc, "点击「两期数据 ⇨ 比较文件」。")
    para(doc, "输出：「比较结果_时间戳.xlsx」。")

    doc.add_heading("5.2 生成说明文件", level=2)
    para(doc, "适用：把比较结果中需要说明的行导出成说明文件，发给机构填写。", bold=True)
    para(doc, "操作步骤：选择上一步的比较结果，选择输出目录，点击生成。")
    para(doc, "输出：「说明文件」工作簿，机构填写「变动原因」后回传。")

    doc.add_heading("5.3 导入反馈", level=2)
    para(doc, "适用：把机构填好的说明文件回收，把「变动原因」写回说明工作簿。", bold=True)
    para(doc, "操作步骤：选择机构回传的反馈文件与目标说明工作簿，点击导入。")
    para(doc, "提示：目标工作簿建议是 .xlsx；导入前软件会自动备份。")

    doc.add_heading("5.4 生成金融表单", level=2)
    para(doc, "适用：按比较结果给每家机构生成一份金融表单。", bold=True)
    para(doc, "操作步骤：选择比较结果、通用配置、表单配置与输出目录，点击生成。")
    para(doc, "输出：输出目录下一家机构一个 xlsx 文件。")

    doc.add_heading("六、常见问题", level=1)
    table(doc, ["现象", "原因与处理"], [
        ["提示「没有匹配到模板」", "确认模板目录里存在与源文件同类型的模板；或手动选择模板文件。"],
        ["提示「源文件格式不支持」", "逐笔与报表系统的源文件需为 .xlsx；旧版 .xls 请先用 Excel 另存为 .xlsx。"],
        ["提示「数据期不唯一」", "同一批源文件里混了不同月份。请按月份分开处理。"],
        ["按钮点了没有反应", "确认没有另一个窗口正在执行任务（界面会显示进度条）。"],
        ["结果目录找不到", "界面会显示本次输出目录的完整路径，可直接点击打开。"],
    ])

    doc.add_heading("七、注意事项", level=1)
    bullet(doc, "原始源文件只读，软件不会修改它们。")
    bullet(doc, "执行期间请不要移动或重命名源文件。")
    bullet(doc, "输出目录不要与源目录相同，避免覆盖。")
    bullet(doc, "遇到无法处理的文件，软件会在运行日志里给出具体原因。")

    doc.save(path)
    print("已生成:", path)


# ---------------------------------------------------------------- 功能架构表

def build_architecture(path: Path) -> None:
    doc = Document()
    style_doc(doc, "Open-Data-Audit 功能架构表（开发者 / AI 参考）")

    para(doc, "版本：2026-11-09　｜　用途：开发维护与 AI 辅助开发的权威索引。"
              "所有结论均以当前代码为准，不描述历史实现。", size=10)

    doc.add_heading("一、仓库结构", level=1)
    table(doc, ["路径", "职责"], [
        ["audit-core/", "业务逻辑、HTTP 桥、CLI、GUI 前端（内嵌 index.html）"],
        ["crates/excel-reader/", "calamine 封装：值/公式/注释/定义名/条件格式读取"],
        ["crates/excel-writer/", "rust_xlsxwriter 封装：新建工作簿"],
        ["crates/excel-editor/", "编辑后端：OOXML 直改 / openpyxl worker / Umya"],
        ["crates/ooxml-patch/", "ZIP 级 OOXML 补丁：公式缓存清理、区域复制、样式剥离"],
        ["sidecar/lo-convert/", "LibreOffice 重算桥（Linux 回退）"],
        ["vendor/", "betteroffice 0.5.1 修改版（content_range + 预算可配）"],
    ])

    doc.add_heading("二、八个功能与实现路径", level=1)
    para(doc, "S1 = 逐笔统计系统，S2 = 报表采集系统，S3 = 大集中统计系统。"
              "「注册表」指 ExcelEditPlan 操作序列；「直算」指纯 Rust 内存处理。", size=10)
    table(doc, ["功能", "阶段链路", "引擎/路径", "实测耗时"], [
        ["S1-F01 汇总核查表校验",
         "模板预检 → 数据期识别 → 生成审核副本（注册表）→ OOXML 规范化 → 重算 → "
         "公式异常提取 → 校验区域扫描 → 条件格式提取 → 历史说明富化 → 汇总发布",
         "编辑：OOXML→openpyxl→Umya；重算：betteroffice→COM/WPS→LibreOffice",
         "189s / 129 本"],
        ["S1-F02 汇总校验结果说明",
         "输入校验 → 模板命名区域解析 → 源结构预检 → 区域聚合 → 三段列（任务/业务/本地校验）→ 写簿",
         "纯 Rust（calamine + excel-writer）", "1.0s"],
        ["S2-F01 执行报表规则",
         "配置校验 → 读本期/上期 → 导入检查 → 环比规则 → 校验规则 → 读大集中 → 外部核对 → 三表导出",
         "纯 Rust（calamine + 内存规则引擎 + excel-writer）", "0.2s"],
        ["S3-F01 两期数据→比较文件",
         "读本期/上期/3.0/3.1 配置 → 表达式解析（legacy-8）→ 导入校验 → 逐规则比较 → 排序分组 → 写",
         "纯 Rust", "2.3s"],
        ["S3-F02 生成说明文件", "读比较结果 → 筛「是否说明=是」→ 推导表名 → 写单表", "纯 Rust", "1.1s"],
        ["S3-F03 导入反馈", "读反馈 → 读目标说明 → 匹配与冲突检测 → 生成回写计划 → 执行回写 → 备份",
         "编辑：OOXML→Umya", "0.2s"],
        ["S3-F04 本期数值核对", "读本期 → 读核对规则 → 逐条核对（容差/单位换算）→ 写", "纯 Rust", "1.0s"],
        ["S3-F05 生成金融表单",
         "读比较结果/3.3/3.0 → 规划机构 → 模板编译（一次）→ 并行渲染（N 线程）→ "
         "SheetPlan（删表/降级/definedNames）→ 轻量校验 → ZIP 写出",
         "FAST_OOXML 自研渲染器；可选 openpyxl 整簿往返", "1.7s / 40s→4.6s"],
    ])

    doc.add_heading("三、重算引擎路由（S1-F01）", level=1)
    para(doc, "优先级：betteroffice → Excel COM（未装则跳过）→ WPS COM → LibreOffice → 明确报错。", bold=True)
    table(doc, ["层", "实现", "说明"], [
        ["betteroffice", "audit-core/src/bo_calc.rs",
         "vendor 0.5.1 修改版；默认 2 轮（第 2 轮 changed=0 收敛）；now_serial 必须显式注入"],
        ["Excel COM", "native_recalculation.rs",
         "Workbook 打开失败时触发 should_retry_with_wps"],
        ["WPS COM", "同文件，候选列表 ket.Application / KET.Application",
         "未装 Excel 时直接作为首选 COM 引擎（不是回退）"],
        ["LibreOffice", "sidecar/lo-convert", "Linux 主回退；Windows 通常未安装"],
    ])
    para(doc, "拒绝条件（任一命中整簿回退，禁止部分采用）：budget_exhausted / limited_cells 非空、"
              "出现 #NAME?、存在循环引用、落盘前发现 #ERROR! 字面。", size=10)
    para(doc, "配置项（GUI > 环境变量 > 默认）：boCalcEngine / boCalcMaxVisits / "
              "boCalcRounds / boCalcStripStyle / boCalcFallback；"
              "环境变量同名兜底 PYTOOLS_BO_CALC*。", size=10)

    doc.add_heading("四、OOXML 规范化（与引擎无关）", level=1)
    para(doc, "位置：audit-core/src/excel_open_repair.rs，在任何引擎打开副本之前执行。", size=10)
    table(doc, ["类", "内容"], [
        ["一", "表名裸引用补单引号（全角括号等强制引号字符）"],
        ["二", "冻结窗格 topLeftCell 落在冻结区内 → 修正"],
        ["三", "外部工作簿引用剥离（四形态：路径式/索引式 × 引号/裸名）"],
        ["四", "shared formula 的 ref 起点归一到主格（glibc/Excel 从属格展开一致）"],
    ])
    para(doc, "安全边界：规范化只作用于**临时工作副本**，绝不回写调用方文件/源文件/模板。"
              "重算后比对基准也必须是「已规范化基线」，否则规范化会被误判成「重算改变公式」。", size=10)

    doc.add_heading("五、编辑后端", level=1)
    table(doc, ["后端", "实现", "适用"], [
        ["OOXML 直改", "ooxml_import_backend.rs",
         "默认首选；支持 11 种操作；写前拒绝不支持批次"],
        ["openpyxl worker", "openpyxl_backend.rs + python/openpyxl_worker.py",
         "备选；fast_roundtrip 模式用于 S3-F05 整簿往返"],
        ["Umya", "umya_backend.rs", "完整兜底；S3-F05 因定义名守卫退出该功能"],
    ])

    doc.add_heading("六、S3-F05 渲染器", level=1)
    para(doc, "模块：s3_fast_ooxml.rs（编译）+ s3_fast_ooxml_render.rs（渲染）。", size=10)
    table(doc, ["阶段", "说明", "实测"], [
        ["模板编译", "一次：CompiledSheet 行结构、keep_cols、indicator_rows、new_dimension；"
                    "空占位表跳过；正则静态缓存", "890ms"],
        ["并行渲染", "std::thread::scope，每表纯字符串重建 sheetData", "~110ms/表"],
        ["SheetPlan", "删内嵌配置页/禁用表单/空表；公式引用迭代收敛降级隐藏；"
                     "definedNames owner 判定 + localSheetId 重映射", "—"],
        ["轻量校验", "核心 XML 可解析 + sheet↔rels↔part 无悬空", "—"],
        ["ZIP 写出", "zip crate，按 part_order 跳过已删部件", "~20ms/表"],
    ])

    doc.add_heading("七、关键不变量", level=1)
    bullet(doc, "原始源文件、模板、调用方传入副本：全程只读，SHA-256 不变。")
    bullet(doc, "所有引擎读的都是临时副本；规范化只改临时副本。")
    bullet(doc, "fail-closed：验证失败不发布、不静默采用受限结果。")
    bullet(doc, "对拍基线：Python 旧 audit（S1F1 3602 条 / 其余逐格 0 差异）。")

    doc.add_heading("八、验证命令", level=1)
    table(doc, ["用途", "命令"], [
        ["全量单测", "cargo test -p audit-core --lib"],
        ["S1F1 双引擎对拍", "cargo run --release -p audit-core --example s1f1_bo_parity <templates> <work> auto"],
        ["对拍比较", "... --example s1f1_bo_parity --compare a.json b.json"],
        ["S3F5 三方验证", "python -I probe/verify_s3f5_engine_solo.py"],
        ["八功能对齐", "python tests/compare/rustaudit_alignment_p1.py --runs 2 --features all"],
        ["Linux 构建(2.28)", "cargo zigbuild --release --target x86_64-unknown-linux-gnu.2.28 -p audit-core --bin open-data-audit"],
        ["Linux deb", "./package-linux-deb.sh <binary> <product-dir> <version>"],
    ])

    doc.add_heading("九、当前限制", level=1)
    table(doc, ["项", "状态"], [
        ["S1-F3/F4/F5", "deferred，已从界面移除（AGENTS.md 标注）"],
        ["WPS COM", "核心能力已验（打开/CalculateFull/SaveAs）；Excel 打开失败→切 WPS 的异常转移未构造样本"],
        ["S3-F05 openpyxl 往返", "需 worker 可用；不可用时回退直出"],
        ["xls 输入", "S1/S2 源文件需 .xlsx；S3-F03 反馈支持 .xls（含伪装 xlsx）"],
    ])

    doc.save(path)
    print("已生成:", path)


def main() -> int:
    out = ROOT / "docs"
    out.mkdir(exist_ok=True)
    build_user_guide(out / "使用说明-全量版.docx", "全量版 S1+S2+S3")
    build_user_guide(out / "使用说明-S1S2版.docx", "S1+S2 版")
    build_architecture(out / "功能架构表-开发者.docx")
    return 0


if __name__ == "__main__":
    sys.exit(main())
