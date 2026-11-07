# S1-F01 债券副本 Excel 拒绝打开：诊断取证（2026-11-07）

对象：`F:/RustBuild/audit/tmp/bond-cacheless-copy.xlsx`（广东博罗·债券业务·清缓存中间副本，
来源 = 探针跑 `F:/RustBuild/audit/tmp/s1f1-single-probe/` 的 WAITING_MANUAL_RECALC 保留件）。
现象：Excel COM `Workbooks.Open` 报 0x80020009“Open 方法无效”；WPS 能开（重算丢公式）；
同款 excel-open-repair 预检“未发现需要修复的内容”（该文件无全角表名裸引用/窗格矛盾两类缺陷）。

## 二分定位实验矩阵（全部用 Excel COM 实测）

| 变体 | 结果 |
|---|---|
| 剥离 comments5+vmlDrawing5+sheet5 rels+legacyDrawing | FAIL |
| 源壳 + 副本全部同名部件（hybrid） | FAIL |
| 源壳 + 仅副本 sheet1 | FAIL |
| 源壳 + 仅副本 sheet2 | **OK** |
| 源壳 + 仅副本 sheet3 | FAIL |
| 源壳 + 仅副本 sheet4 | **OK** |
| 源壳 + 仅副本 sheet5（rels 悬空，实验伪影） | FAIL |
| 副本 − sheet6/7 导入（workbook/rels/CT 同步剥离） | FAIL |
| 副本去 forceFullCalc/fullCalcOnLoad | FAIL |
| **openpyxl 全量重写副本** | **FAIL**（排除一切 XML 写法问题） |
| openpyxl 重写 + 只留 sheet1 | FAIL（问题在 sheet1 语义内容） |
| openpyxl 重写 + 清空全部单元格值 | **OK** |
| openpyxl 重写 + 删定义名称（保留单元格值） | FAIL |
| 清值 + 回填 sheet1 全部 650 公式的前半/后半 | 均 FAIL |

## 已排除清单

包结构（部件/Content_Types/三层 rels/元素顺序/行地址/共享公式主从/样式索引越界/
非法字符/实体/全角残留/inlineStr 结构）、f+is 组合、公式长度（最长 337 字符）、
强制重算、外部导入 sheet6/7、批注与 VML（数量/结构与 Python 可打开副本同构）、
定义名称（两侧一致）。

## 已锁定范围

- 杀手在 **sheet1（存量债券投资）的单元格语义内容**：清空全部值（含公式）→ 打开成功；
  保留任一 20 行段的值/公式 → 失败（四段全失败 → 多个触发点或系统性组合）。
- 自动二分收敛到含 **N8**（`IF(AND(G8<>"",ABS(G8)>5000),"错|"&A8&B8&"|"&G8&"|注意…")`）
  的候选集，但 **N8 单条、N8+第 8 行 C8/F8/G8 组**均能打开——是多公式组合触发。
- 650 条公式按前半/后半回填均失败 → 两半各含触发成员。
- 单条探针：`(-105020879.61--102311726.56)<>…` 的 IEEE 尾差与本问题无关
  （probe_minus_minus_comparison 测试已固化）。

## 下一轮最短路径

1. 从“清空版（可打开）”出发做**逐条公式回填的自动二分**，但改用
   **“保留行”二分**（对行而非公式集合），并允许多杀手：对 1-40/41-80
   分别递归（killers ≥ 2）。
2. 定位最小触发组合后，与旧 Python 副本同位置内容对比（该副本 Excel 可开，
   在 `3.open-data-audit/test-output/s1f1_com_isolation_20261107/isolated/python/audit_copies/`），
   差异即修复口径（大概率是模板校验区域公式与原有公式的某种组合或写入顺序）。
3. 修复落点候选：`ooxml_patch` 的校验区域导入（inlineStr 公式/值写入顺序、
   是否将模板公式与既有 shared 公式体系交错）。
