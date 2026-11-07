Attribute VB_Name = "通用"
Option Explicit

'
'   统计路漫漫，手动脑壳疼
'
'   造此物能省则省
'
'   如有需求可再交流
'
'                   by cyc
'

Public gloT As Variant      '计时器
Public gloEt As New CMyExcelTool
Public gloFt As New CMyFileTool
Public gloCft As New CMyConfigTool
Public gloDe As New CMyDebugTool
Public glo_log As New CMylogging    '用来记录程序工作的一些问题
Public Const SYSSTARTROW As Integer = 5 '面板默认标题行
Public glo_SforbidCol As Integer
Public glo_CforbidCol As Integer
Public glo_fileFilter As String
Public glo_iOrgMode As Integer
Public Const OLD_SATUS_NUM As Integer = 10
Public Const SUPPROT_WSNAME As String = "yc辅助区"
Public glo_dcz As Object      '存放程序运行参数

Public Const HZ_XUHAO_FIELD As String = "序号"
Public Const HZ_LONGNAME_FIELD As String = "源文件名"
Public Const HZ_SHORTNAME_FIELD As String = "源文件"
Public Const HZ_BIAOZHUN_FIELD As String = "标准化"
Public Const HZ_NEWNAME_FIELD As String = "新文件名"
Public Const HZ_JIAOYAN_FIELD As String = "校验结果"
Public Const HZ_ZHIXING_FIELD As String = "执行结果"
Public Const HZ_BEIZHU_FIELD As String = "备注"

'
' 读取执行面板参数
'
Public Function InitMBcz() As Object
    On Error Resume Next
    Set glo_dcz = CreateObject("scripting.dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("报表跨期面板")
    
    Dim arr As Variant, i As Long, j As Long
    arr = gloEt.GetUsedRange(ws).value
    
    For i = 1 To UBound(arr)
        If CStr(arr(i, 1)) <> "" Then
            glo_dcz(CStr(arr(i, 1))) = CStr(arr(i, 2))
        End If
    Next
End Function


Sub SplitByLastComma()
    Dim s As String
    s = "a,b,c"  ' 你可以换成任意字符串

    Dim pos As Long
    pos = InStrRev(s, ",")

    Dim part1 As String, part2 As String
    If pos > 0 Then
        part1 = left(s, pos - 1)
        part2 = Mid(s, pos + 1)
    Else
        part1 = s
        part2 = ""  ' 没有逗号的情况
    End If

    Debug.Print "前部分: " & part1
    Debug.Print "后部分: " & part2
End Sub

Public Function ycS(str As Variant) As String
    On Error Resume Next
    ycS = CStr(str)
End Function

Public Function GetStrippedValue(v As Variant) As Variant
    GetStrippedValue = v
End Function

Public Function ActiveCondition(rng As Range, wb As Workbook) As Integer
    Dim Ndx As Long
    Dim FC As Object 'FormatCondition
    Dim temp As Variant
    Dim Temp2 As Variant
    Dim colorindex As Integer
    If rng.FormatConditions.count = 0 Then
        ActiveCondition = 0
    Else
       ' On Error Resume Next
        For Ndx = 1 To rng.FormatConditions.count
            Set FC = rng.FormatConditions(Ndx)
            colorindex = CInt(ycV(FC.Interior.colorindex))
            If colorindex = 0 Then colorindex = 1
            Select Case FC.Type
                Case xlCellValue
                Select Case FC.Operator
                    Case xlBetween
                        temp = GetStrippedValue(FC.Formula1)
                        Temp2 = GetStrippedValue(FC.Formula2)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) >= ycV(FC.Formula1) And _
                               ycV(rng.value) <= ycV(FC.Formula2) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                       Else
                          If rng.value >= temp And _
                             rng.value <= Temp2 Then
                             ActiveCondition = colorindex
                             Exit Function
                          End If
                       End If
                    Case xlGreater
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) > ycV(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If rng.value > temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                    Case xlEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) = ycV(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                        Else
                           If temp = rng.value Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
    
                    Case xlGreaterEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) >= ycV(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                        Else
                           If rng.value >= temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                   
                    Case xlLess
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                            If ycV(rng.value) < ycV(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                            End If
                        Else
                            If rng.value < temp Then
                               ActiveCondition = colorindex
                               Exit Function
                            End If
                        End If
                    Case xlLessEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) <= ycV(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If rng.value <= temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
    
                    Case xlNotEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(rng.value) <> ycV(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If temp <> rng.value Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                   Case xlNotBetween
                        temp = GetStrippedValue(FC.Formula1)
                        Temp2 = GetStrippedValue(FC.Formula2)
                        If IsNumeric(temp) Then
                           If Not (ycV(rng.value) <= ycV(wb.Application.Evaluate(replace(FC.Formula1, "", 0)))) And _
                              (ycV(rng.value) >= ycV(wb.Application.Evaluate(replace(FC.Formula2, "", 0)))) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If Not rng.value <= temp And _
                              rng.value >= Temp2 Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                
                   Case Else
                        Debug.Print "UNKNOWN OPERATOR"
               End Select
    
            Case xlExpression
                wb.Activate
                rng.Parent.Activate
                rng.Select
                On Error Resume Next
                If wb.Application.Evaluate(FC.Formula1) Then
                    If err.Number = 0 Then
                        ActiveCondition = colorindex
                        Exit Function
                    Else
                        err.Clear
                    End If
                End If
            Case Else
                Debug.Print "UNKNOWN TYPE"
           End Select
        Next Ndx
    End If
    ActiveCondition = 0
End Function


'
'   兼容性
'
Public Function ycV(va As Variant) As Double
    If VarType(va) = vbError Then
        ycV = 0
    Else
        ycV = val(va)
    End If
End Function

'
'   兼容性1
'
Public Function ycD(va As Variant) As Date
    If VarType(va) = vbError Then
        ycD = Date
    Else
        ycD = CDate(va)
    End If
End Function

Public Function ycG(d As Variant, name As String) As Variant
    On Error Resume Next
    ycG = d(name)
End Function

Public Function ycIsDatae(d As Variant) As Boolean
    
End Function

'
'   去掉左右两边的空格，以及去除第一个"'"
'
Public Function ycTrim(v As Variant) As String
    On Error GoTo err
    Dim str As String
    str = Trim(CStr(v))
    If Len(str) > 1 And left(str, 1) = "'" Then
        str = right(str, Len(str) - 1)
    End If
    ycTrim = str
Exit Function
err:
    ycTrim = ""
End Function


Public Function ycGetDate(lstr As String)
    Dim y As Integer, m As Integer, str As String
    str = right(lstr, Len(lstr) - InStrRev(lstr, "\"))
    y = CInt(left(str, 4))
    m = CInt(Mid(str, 6, InStr(str, " ") - InStr(str, ".")))
    
    ycGetDate = DateSerial(y, m + 1, 0)
End Function

Public Function ycGetOrgName2(str As String)
    On Error Resume Next
    '先判断是否是日期+调查+机构类+机构名称的形式，以是否包含两个空格进行判断
    If Len(str) - Len(replace(str, " ", "")) < 2 Then
        ycGetOrgName2 = ""
        Exit Function
    End If
    
    Dim r1 As Integer, r2 As Integer
    r2 = VBA.InStrRev(str, " ")
    r1 = VBA.InStrRev(left(str, r2 - 1), " ")
    
    ycGetOrgName2 = Mid(str, r1 + 1, r2 - r1 - 1)

    If err.Number <> 0 Then
        ycGetOrgName2 = ""
    End If
End Function

Public Function ycGetOrgName(str As String)
    On Error Resume Next
    '先判断是否是日期+调查+机构类+机构名称的形式，以是否包含两个空格进行判断
    If Len(str) - Len(replace(str, " ", "")) < 2 Then
        ycGetOrgName = ""
        Exit Function
    End If
    
    If VBA.InStrRev(str, ".") < VBA.InStrRev(str, " ") Then
        ycGetOrgName = Mid(str, VBA.InStrRev(str, " ") + 1, Len(str) - VBA.InStrRev(str, " "))
    Else
        ycGetOrgName = Mid(str, VBA.InStrRev(str, " ") + 1, VBA.InStrRev(str, ".") - VBA.InStrRev(str, " ") - 1)
    End If
    If err.Number <> 0 Then
        ycGetOrgName = ""
    End If
End Function


'
'   该字符串左边起到哪个字符串
'
Public Function ycLeft(str As String, findstr As String, Optional num As Long = 1, Optional bh As Boolean = True)
    Dim lf As Long
    lf = CLng(ycV(ycInStr(str, findstr, num)))
    
    If lf = 0 Then ycLeft = "": Exit Function
    If bh = False And lf = 1 Then ycLeft = "": Exit Function
    
    If bh = True Then
        ycLeft = left(str, lf)
    Else
        ycLeft = left(str, lf - 1)
    End If
End Function

Public Function ycMid(str As String, lfindstr As String, lnum As Long, lbh As Boolean, _
                         rfindstr As String, rnum As Long, rbh As Boolean)
    Dim lf As Long, lr As Long
    lf = CLng(ycV(ycInStr(str, lfindstr, lnum)))
    lr = CLng(ycV(ycInStr(str, rfindstr, rnum)))
    
    If lf = 0 Then ycMid = "": Exit Function    '左边区域没有找到
    If lbh = False And lf = Len(str) Then ycMid = "": Exit Function   '左边区域找到的话，但是不包含,因此不能是最后一个
    If rbh = False And lr = 1 Then ycMid = "": Exit Function    '右边区域找到的话，但是不包含，因此不能是第一个
    
    If lr = 0 Then rbh = True: lr = Len(str)    '如有右边没找到的话，就相当于全部
    If lbh = False Then lf = lf + Len(lfindstr)             '如果左边不包含，则向右移动1，+1
    If rbh = False Then lr = lr - 1             '如果右边不包含，则向左边移动1，-1
    If lr < lf Then ycMid = "": Exit Function
    
    ycMid = Mid(str, lf, lr - lf + 1)   '取数
End Function

Public Function ycRight(str As String, findstr As String, Optional num As Long = 1, Optional bh As Boolean = True)
    Dim lf As Long
    lf = CLng(ycV(ycInStrRev(str, findstr, num)))
    
    If lf = 0 Then ycRight = "": Exit Function
    If bh = False And lf = 1 Then ycRight = "": Exit Function
    
    If bh = True Then
        ycRight = right(str, lf)
    Else
        ycRight = right(str, lf - Len(findstr))
    End If
End Function
'
'   从右边开始算，findstr出现第num次，在字符串从左边开始算起的什么位置
'
Public Function ycInStrRev(str As String, findstr As String, Optional num As Long = 1)
    Dim i As Long, strTmp As String, r As Long, rel As Long
    strTmp = CStr(str)
    rel = 0
    For i = 1 To num Step 1
        r = InStrRev(strTmp, findstr)
        If r <= 0 Then ycInStrRev = "": Exit Function
        strTmp = left(strTmp, r - 1)
        rel = r
    Next i
    'num 出现次数
    ycInStrRev = Len(str) - rel + 1
End Function

'
'   从左边开始算，findstr出现第num次，在字符串从左边开始算起的什么位置
'
Public Function ycInStr(str As String, findstr As String, Optional num As Long = 1)
    Dim i As Long, strTmp As String, r As Long, rel As Long
    rel = 0
    strTmp = CStr(str)
    For i = 1 To num Step 1
        r = InStr(strTmp, findstr)
        If r <= 0 Then ycInStr = "": Exit Function
        strTmp = Mid(strTmp, r + 1, Len(strTmp))
        rel = rel + r
    Next i
    'num 出现次数
    ycInStr = rel
End Function

'
'   是否执行
'
Public Function IsContinue(str As String) As Boolean
    Dim r As Variant
    r = MsgBox(str, vbYesNo, Application.Caption)
    If r = vbYes Then
        IsContinue = True
    Else
        Call gloCft.OpenSomething
        IsContinue = False
    End If
End Function




Public Sub ShowChangeName()
    Call 修改名称.Show(0)
End Sub

Sub SetTrustInVBE()
    Dim blChgset As Boolean
    On Error Resume Next
    '陷阱测试，下面第一句blChgset赋值无实质意义
    blChgset = ThisWorkbook.VBProject.Protection
    '如果发生错误则说明需要设置信任工程
    blChgset = err
    If blChgset Then
        err.Clear
        '设置信任工程访问
        SendKeys "%TMS%T%V{ENTER}"
        DoEvents    '更新状态
    End If
    '要执行的操作....
    '.....
    'Application.VBE.MainWindow.Visible = False
    '操作完成后还原操作前的状态
    'If blChgset Then
    '    SendKeys "%TMS%T%V{ENTER}"
    'End If
End Sub

Public Function GetUnitVal(str As String) As Long
    Dim unit As Long
    unit = 1
    
    Select Case str
        Case "元"
            unit = 1
        Case "十元"
            unit = 10
        Case "百元"
            unit = 100
        Case "千元"
            unit = 1000
        Case "万元"
            unit = 10000
        Case "十万元"
            unit = 100000
        Case "百万元"
            unit = 1000000
        Case "千万"
            unit = 10000000
        Case "亿元"
            unit = 100000000
    End Select
    
    GetUnitVal = unit
End Function

'
'   去掉字符串的空格及分号
'
Public Function CTR(str As Variant) As String
    CTR = replace(Trim(Asc(CStr(str))), "'", "")
End Function


'
'   转换为日期
'
Public Function ChangeToCDate(v As Variant) As Variant
    On Error Resume Next
    v = CDate(v)
    If err.Number <> 0 Then err.Clear
    
    ChangeToCDate = v
End Function

'
'   日期转换为 yyyy.mm
'
Public Function ycDateToStr(v2 As Variant) As Variant
    On Error Resume Next
    Dim v As Variant
    v = v2
    v = CDate(v)
    v = Format(v, "yyyy.mm")
    If err.Number <> 0 Then err.Clear
    
    ycDateToStr = CStr(v)
End Function

Public Function ycIsNum(t As Variant) As Boolean
    Dim IsRight As Boolean
    IsRight = True
    
    If VarType(t) = vbEmpty Or VarType(t) = vbSingle Or _
        VarType(t) = vbLong Or VarType(t) = vbInteger Or _
        VarType(t) = vbDouble Or VarType(t) = vbDecimal Then
        '说明是数值，忽略
    ElseIf VarType(t) = vbError Then
        IsRight = False
    Else
        If ycV(t) <> 0 Then
            If ycTrim(ycV(t)) <> ycTrim(t) Then
                IsRight = False
            End If
        End If
    End If
    ycIsNum = IsRight
End Function

Public Function ycIsDate(v As Variant) As Boolean
    On Error Resume Next
    Dim IsRight As Boolean, t As Variant
    IsRight = True
    
    If VarType(v) = vbDate Then Exit Function
        
    t = CDate(v)
    If err.Number <> 0 Then
        err.Clear
        IsRight = False
    Else
        '防止填的是数值
        If Format(t, "yyyy-mm-dd") <> CStr(v) Then
            IsRight = False
        End If
    End If
    ycIsDate = IsRight
End Function

Public Function OutputExcelbk(wsCur As Worksheet, arrCur As Variant) As Boolean
    On Error GoTo ErrorHandler

    ' 确保变量声明明确
    Dim col As Long
    Dim rows As Long
    Dim cur_rows As Long
    Dim new_row As Long
    Dim i As Long, j As Long
    Dim ws_temp As Worksheet
    Dim arr_temp As Variant
    Dim t As Double
    Dim current_org As String
    Dim last_org As String
    Dim page_start As Long, page_end As Long

    ' 获取行数和列数
    rows = UBound(arrCur, 1)
    col = UBound(arrCur, 2)

    ' 判断行数是否需要分页
    If rows <= 60000 Then
        ' 小于50000行，直接写入
        wsCur.Range(wsCur.Cells(1, 1), wsCur.Cells(rows, col)).value = arrCur
        wsCur.Cells.EntireRow.WrapText = False
    Else
        ' 超过50000行分页输出
        page_start = 2
        i = 0

        Do While page_start <= rows
            Call gloCft.ShowStatusMessage_Up("正在输出文件：页 " & CStr(i + 1))
            
            ' 确定分页范围
            page_end = page_start + 59999
            If page_end >= rows Then
                page_end = rows
            Else
                ' 检查 C 列的机构名称，确保机构完整性
                last_org = CStr(arrCur(page_end, 3))
                current_org = last_org
                Do While page_end < rows And CStr(arrCur(page_end + 1, 3)) = current_org
                    page_end = page_end + 1
                Loop
            End If
            
            ' 计算当前分页行数
            cur_rows = page_end - page_start + 1

            ' 定义分页数组，行数加1用于列标题
           
            ReDim arr_temp(1 To cur_rows + 1, 1 To col)
            
            ' 复制列标题
            For j = 1 To col
                arr_temp(1, j) = arrCur(1, j)
            Next

            ' 复制当前分页数据
            For new_row = 1 To cur_rows
                For j = 1 To col
                    arr_temp(new_row + 1, j) = arrCur(page_start + new_row - 1, j)
                Next
            Next

            ' 判断是当前Sheet还是新建Sheet
            If i = 0 Then
                Set ws_temp = wsCur
            Else
                Set ws_temp = wsCur.Parent.Worksheets.add
            End If
            
            ' 设置新Sheet名称，避免命名冲突
            On Error Resume Next
            ws_temp.name = "说明(" & i & ")"
            On Error GoTo ErrorHandler

            ' 输出分页数据到当前Sheet
            ws_temp.Cells.EntireRow.WrapText = False
            ws_temp.Cells.NumberFormat = "@"
            If col > 16 Then
                ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(cur_rows + 1, 16)).value = arr_temp
            Else
                ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(cur_rows + 1, col)).value = arr_temp
            End If
            ws_temp.Cells.EntireRow.WrapText = False
            ws_temp.Range("g:j").NumberFormat = "0.00"
            ' 清空数组
            Erase arr_temp
            
            ' 更新起始行
            page_start = page_end + 1
            i = i + 1
        Loop
    End If

    ' 成功返回
    OutputExcelbk = True
    Exit Function

ErrorHandler:
    OutputExcelbk = False
    MsgBox "Error: " & err.Description, vbCritical
End Function



Public Function OutputExcel(wsCur As Worksheet, arrCur As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 是否启用排序
    Dim sortedIndexArr() As Long
    If CStr(glo_dcz("是否启用排序")) = "是" Then
        Dim sortCols(0 To 1) As Variant, ascendingOrders(0 To 1) As Variant
        sortCols(0) = 3: sortCols(1) = 4
        ascendingOrders(0) = True: ascendingOrders(1) = True
        sortedIndexArr = SortIndexOnly(arrCur, sortCols, ascendingOrders, True)
    Else
        ' 如果不排序，则按原始顺序输出
        OutputExcel = OutputExcelbk(wsCur, arrCur)
        Exit Function
    End If

    ' 确保变量声明明确
    Dim col As Long, rows As Long, cur_rows As Long, new_row As Long
    Dim j As Long, page_start As Long, page_end As Long, i As Long
    Dim ws_temp As Worksheet, arr_temp As Variant
    Dim current_org As String, last_org As String

    ' 获取行数和列数
    rows = UBound(arrCur, 1)
    col = UBound(arrCur, 2)

    ' 判断行数是否需要分页
    If rows <= 60000 Then
        Dim sortedData As Variant
        ReDim sortedData(1 To rows, 1 To col)
        
        ' 保留标题行
        For j = 1 To col
            sortedData(1, j) = arrCur(1, j)
        Next j
        
        ' 复制排序后数据
        For i = 2 To rows
            For j = 1 To col
                sortedData(i, j) = arrCur(sortedIndexArr(i), j)
            Next j
        Next i
        
        wsCur.Range(wsCur.Cells(1, 1), wsCur.Cells(rows, col)).NumberFormat = "@"
        wsCur.Range(wsCur.Cells(1, 1), wsCur.Cells(rows, col)).value = sortedData
        wsCur.Range("G1:" & "J" & CStr(rows)).NumberFormat = "0.00"
        Erase sortedData
        wsCur.Cells.EntireRow.WrapText = False
    Else
        ' 超过50000行分页输出
        page_start = 2
        i = 0

        Do While page_start <= rows
            Call gloCft.ShowStatusMessage_Up("正在输出文件：页 " & CStr(i + 1))
            
            ' 确定分页范围
            page_end = page_start + 59999
            If page_end >= rows Then
                page_end = rows
            Else
                ' 确保机构完整性（C列）
                last_org = CStr(arrCur(sortedIndexArr(page_end), 3))
                current_org = last_org
                Do While page_end < rows And CStr(arrCur(sortedIndexArr(page_end + 1), 3)) = current_org
                    page_end = page_end + 1
                Loop
            End If
            
            ' 计算当前分页行数
            cur_rows = page_end - page_start + 1

            ' 定义分页数组
            ReDim arr_temp(1 To cur_rows + 1, 1 To col)
            
            ' 复制列标题
            For j = 1 To col
                arr_temp(1, j) = arrCur(1, j)
            Next

            ' 复制当前分页数据（按排序后的索引）
            For new_row = 1 To cur_rows
                For j = 1 To col
                    arr_temp(new_row + 1, j) = arrCur(sortedIndexArr(page_start + new_row - 1), j)
                Next
            Next

            ' 处理分页 Sheet
            If i = 0 Then
                Set ws_temp = wsCur
            Else
                Set ws_temp = wsCur.Parent.Worksheets.add
            End If
            
            ' 设置新 Sheet 名称，避免冲突
            On Error Resume Next
            ws_temp.name = "说明(" & i & ")"
            On Error GoTo ErrorHandler

            ' 输出分页数据
            ws_temp.Cells.NumberFormat = "@"
            ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(cur_rows + 1, col)).value = arr_temp
            ws_temp.Range("G:J").NumberFormat = "0.00"
            ws_temp.Cells.EntireRow.WrapText = False
            
            ' 清空数组
            Erase arr_temp
            
            ' 更新分页起始行
            page_start = page_end + 1
            i = i + 1
        Loop
    End If
    Erase sortedIndexArr
    ' 成功返回
    OutputExcel = True
    Exit Function

ErrorHandler:
    OutputExcel = False
    Erase sortedIndexArr
    MsgBox "Error: " & err.Description, vbCritical
End Function

Public Function SortIndexOnly(ByRef arr As Variant, ByVal sortCols As Variant, ByVal ascendingOrders As Variant, Optional ByVal isChinese As Boolean = False) As Variant
    Dim indexArr() As Long
    Dim i As Long, rows As Long
    
    ' 初始化索引数组
    rows = UBound(arr, 1)
    ReDim indexArr(2 To rows)
    For i = 2 To rows
        indexArr(i) = i
    Next i

    ' 只对索引排序
    Sort2DArrayMulti indexArr, arr, sortCols, ascendingOrders, isChinese
    
    SortIndexOnly = indexArr
End Function

Public Sub Sort2DArrayMulti(ByRef indexArr As Variant, ByRef arr As Variant, ByVal sortCols As Variant, ByVal ascendingOrders As Variant, Optional ByVal isChinese As Boolean = False)
    Dim stack() As Variant, top As Long
    Dim lower As Long, upper As Long, left As Long, right As Long
    
    lower = LBound(indexArr, 1)
    upper = UBound(indexArr, 1)

    ' 预防数组过大，动态分配栈空间
    ReDim stack(1 To (upper - lower + 1), 2)
    top = 1
    stack(top, 1) = lower
    stack(top, 2) = upper

    Do While top > 0
        left = stack(top, 1)
        right = stack(top, 2)
        top = top - 1

        If left < right Then
            Dim pivotIndex As Long
            pivotIndex = PartitionIndex(indexArr, arr, left, right, sortCols, ascendingOrders, isChinese)

            If left < pivotIndex - 1 Then
                top = top + 1
                stack(top, 1) = left
                stack(top, 2) = pivotIndex - 1
            End If
            If pivotIndex + 1 < right Then
                top = top + 1
                stack(top, 1) = pivotIndex + 1
                stack(top, 2) = right
            End If
        End If
    Loop
End Sub

Private Function PartitionIndex(ByRef indexArr As Variant, ByRef arr As Variant, ByVal low As Long, ByVal high As Long, ByVal sortCols As Variant, ByVal ascendingOrders As Variant, ByVal isChinese As Boolean) As Long
    Dim pivot As Long, i As Long, j As Long, temp As Long
    pivot = indexArr(high)
    i = low - 1
    
    For j = low To high - 1
        If CompareRows(arr, indexArr(j), pivot, sortCols, ascendingOrders, isChinese) <= 0 Then
            i = i + 1
            temp = indexArr(i)
            indexArr(i) = indexArr(j)
            indexArr(j) = temp
        End If
    Next j
    
    temp = indexArr(i + 1)
    indexArr(i + 1) = indexArr(high)
    indexArr(high) = temp
    
    PartitionIndex = i + 1
End Function

' ====== 比较两行数据，返回 -1、0、1（支持中文）======
Private Function CompareRows(ByRef arr As Variant, ByVal row1 As Long, ByVal row2 As Long, ByVal sortCols As Variant, ByVal ascendingOrders As Variant, ByVal isChinese As Boolean) As Long
    Dim k As Long, col As Long, order As Boolean, cmpResult As Long
    
    For k = LBound(sortCols) To UBound(sortCols)
        col = sortCols(k) ' 当前列
        order = ascendingOrders(k) ' 是否升序
        
        ' 处理中文或普通比较
        If isChinese Then
            cmpResult = StrComp(CStr(arr(row1, col)), CStr(arr(row2, col)), vbTextCompare) ' 中文拼音排序
        Else
            ' 兼容数字和文本的比较
            If IsNumeric(arr(row1, col)) And IsNumeric(arr(row2, col)) Then
                If CDbl(arr(row1, col)) < CDbl(arr(row2, col)) Then cmpResult = -1
                If CDbl(arr(row1, col)) > CDbl(arr(row2, col)) Then cmpResult = 1
            Else
                cmpResult = StrComp(CStr(arr(row1, col)), CStr(arr(row2, col)), vbBinaryCompare)
            End If
        End If
        
        ' 按排序方式返回
        If cmpResult < 0 Then
            CompareRows = IIf(order, -1, 1)
            Exit Function
        ElseIf cmpResult > 0 Then
            CompareRows = IIf(order, 1, -1)
            Exit Function
        End If
    Next k
    
    ' 所有列都相等
    CompareRows = 0
End Function


