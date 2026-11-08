"""构造 betteroffice 拒绝分支的样本（循环引用 / #NAME?）与保真样本（条件格式）。

用法：python -I probe/make_bo_edge_samples.py <out_dir>
输出三个 xlsx：
  cycle.xlsx       — A1=B1, B1=A1（循环引用）
  name_error.xlsx  — 使用 betteroffice 未实现的函数（预期 #NAME?）
  cformat.xlsx     — 条件格式 + 数据验证 + 合并 + 定义名（重算保真）
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation


def build_cycle(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "循环"
    sheet["A1"] = "=B1"
    sheet["B1"] = "=A1"
    sheet["C1"] = "=1+1"
    workbook.save(path)


def build_name_error(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "未实现"
    # betteroffice 未实现的函数（Excel 有、上游引擎没实现 → #NAME?）
    sheet["A1"] = "=GEOMEAN(1,2,3)"
    sheet["A2"] = "=1+1"
    workbook.save(path)


def build_cformat(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "表单"
    sheet["A1"] = "指标"
    sheet["B1"] = "余额"
    for row in range(2, 12):
        sheet[f"A{row}"] = f"指标{row - 1}"
        sheet[f"B{row}"] = row * 10
        sheet[f"C{row}"] = f"=B{row}*2"
    # 条件格式：>50 标红
    sheet.conditional_formatting.add(
        "B2:B11",
        CellIsRule(operator="greaterThan", formula=["50"],
                   fill=PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")),
    )
    # 数据验证
    validation = DataValidation(type="whole", operator="between", formula1="0", formula2="1000")
    sheet.add_data_validation(validation)
    validation.add("B2:B11")
    # 合并
    sheet.merge_cells("D1:E1")
    sheet["D1"] = "合并表头"
    # 定义名
    workbook.defined_names["余额区间"] = DefinedName("余额区间", attr_text="表单!$B$2:$B$11")
    # 第二个表做跨表引用（验证跨表链重算）
    other = workbook.create_sheet("汇总")
    other["A1"] = "=SUM(表单!C2:C11)"
    workbook.save(path)


def main() -> int:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "probe/reports/bo_edge")
    out.mkdir(parents=True, exist_ok=True)
    build_cycle(out / "cycle.xlsx")
    build_name_error(out / "name_error.xlsx")
    build_cformat(out / "cformat.xlsx")
    print(f"样本已生成：{out}")
    for name in ("cycle.xlsx", "name_error.xlsx", "cformat.xlsx"):
        print(" ", out / name, (out / name).stat().st_size, "bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
