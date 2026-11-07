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


Public Const HZ_XUHAO_FIELD As String = "序号"
Public Const HZ_LONGNAME_FIELD As String = "源文件名"
Public Const HZ_SHORTNAME_FIELD As String = "源文件"
Public Const HZ_BIAOZHUN_FIELD As String = "标准化"
Public Const HZ_NEWNAME_FIELD As String = "新文件名"
Public Const HZ_JIAOYAN_FIELD As String = "校验结果"
Public Const HZ_ZHIXING_FIELD As String = "执行结果"
Public Const HZ_BEIZHU_FIELD As String = "备注"

Public Function ycS(str As Variant) As String
    On Error Resume Next
    ycS = CStr(str)
End Function

Public Function GetStrippedValue(v As Variant) As Variant
    GetStrippedValue = v
End Function

Public Function ActiveCondition(Rng As Range, wb As Workbook) As Integer
    Dim Ndx As Long
    Dim FC As Object 'FormatCondition
    Dim temp As Variant
    Dim Temp2 As Variant
    Dim colorindex As Integer
    If Rng.FormatConditions.count = 0 Then
        ActiveCondition = 0
    Else
       ' On Error Resume Next
        For Ndx = 1 To Rng.FormatConditions.count
            Set FC = Rng.FormatConditions(Ndx)
            colorindex = CInt(ycV(FC.Interior.colorindex))
            If colorindex = 0 Then colorindex = 1
            Select Case FC.Type
                Case xlCellValue
                Select Case FC.Operator
                    Case xlBetween
                        temp = GetStrippedValue(FC.Formula1)
                        Temp2 = GetStrippedValue(FC.Formula2)
                        If IsNumeric(temp) Then
                           If CDbl(Rng.Value) >= CDbl(FC.Formula1) And _
                               CDbl(Rng.Value) <= CDbl(FC.Formula2) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                       Else
                          If Rng.Value >= temp And _
                             Rng.Value <= Temp2 Then
                             ActiveCondition = colorindex
                             Exit Function
                          End If
                       End If
                    Case xlGreater
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If CDbl(Rng.Value) > CDbl(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If Rng.Value > temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                    Case xlEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If CDbl(Rng.Value) = CDbl(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                        Else
                           If temp = Rng.Value Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
    
                    Case xlGreaterEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If CDbl(Rng.Value) >= CDbl(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                           End If
                        Else
                           If Rng.Value >= temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                   
                    Case xlLess
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                            If ycV(Rng.Value) < ycV(FC.Formula1) Then
                               ActiveCondition = colorindex
                               Exit Function
                            End If
                        Else
                            If Rng.Value < temp Then
                               ActiveCondition = colorindex
                               Exit Function
                            End If
                        End If
                    Case xlLessEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If CDbl(Rng.Value) <= CDbl(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If Rng.Value <= temp Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
    
                    Case xlNotEqual
                        temp = GetStrippedValue(FC.Formula1)
                        If IsNumeric(temp) Then
                           If ycV(Rng.Value) <> CDbl(FC.Formula1) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If temp <> Rng.Value Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                   Case xlNotBetween
                        temp = GetStrippedValue(FC.Formula1)
                        Temp2 = GetStrippedValue(FC.Formula2)
                        If IsNumeric(temp) Then
                           If Not (ycV(Rng.Value) <= ycV(wb.Application.Evaluate(replace(FC.Formula1, "", 0)))) And _
                              (ycV(Rng.Value) >= ycV(wb.Application.Evaluate(replace(FC.Formula2, "", 0)))) Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        Else
                           If Not Rng.Value <= temp And _
                              Rng.Value >= Temp2 Then
                              ActiveCondition = colorindex
                              Exit Function
                           End If
                        End If
                
                   Case Else
                        Debug.Print "UNKNOWN OPERATOR"
               End Select
    
            Case xlExpression
                wb.Activate
                Rng.Parent.Activate
                Rng.Select
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
    If lbh = False Then lf = lf + 1             '如果左边不包含，则向右移动1，+1
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
        ycRight = right(str, lf - 1)
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


'
'   3. 对应菜单 系统导数比较
'
Public Sub ShowDoCompare()
    glo_SforbidCol = 10
    glo_CforbidCol = 5
    glo_iOrgMode = 1
    glo_fileFilter = "Excel 文件 (*.xls; *.csv), *.xls; *.csv,所有文件(*.*),*.*"
    Call 导数比较窗体.Show(0)
    'glo_fileFilter = ""
End Sub

Public Sub ShowDoCompare2()
    '对于金融机构
    glo_SforbidCol = 11
    glo_CforbidCol = 9
    glo_iOrgMode = 2
    glo_fileFilter = "IJ 文件 (*.IDX;), *.IDX;,所有文件(*.*),*.*"
    Call 导数比较窗体.Show(0)
    'glo_fileFilter = ""
End Sub

Public Sub ShowDoReload()
    glo_SforbidCol = 10
    glo_CforbidCol = 5
    glo_iOrgMode = 1
    glo_fileFilter = "Excel 文件 (*.xls; *.csv), *.xls; *.csv,所有文件(*.*),*.*"
    Call 系统导数转EXCLE窗体.Show(0)
End Sub

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
    If err.Number <> 0 Then
        err.Clear
        v = Now()
    End If
    ChangeToCDate = v
End Function

'
'   日期转换为 yyyy.mm
'
Public Function ycDateToStr(v2 As Variant, Optional date_format As String = "yyyy.mm") As Variant
    On Error Resume Next
    Dim v As Variant
    v = v2
    v = CDate(v)
    v = Format(v, date_format)
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


