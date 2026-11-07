VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMylogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'   用来记录日志，并输出到指定工作簿
'   也可以作为一个记录台账
'

Private log_data As Object
Private dlog_title As Object
Private cur_row As Long     '储存标题
Private cur_col As Integer  '储存当前最大的列
Private Const MAXCOL As Integer = 256

'类初始化
Private Sub Class_Initialize()
    Set log_data = Nothing
    Set log_data = CreateObject("scripting.dictionary")
    Set dlog_title = Nothing
    Set dlog_title = CreateObject("scripting.dictionary")
    
    log_data(1) = Split("", "|")
    cur_row = 1
    cur_col = 0
End Sub

'类销毁
Private Sub Class_Terminate()
    Set log_data = Nothing
    Set dlog_title = Nothing
End Sub

Public Sub clear_log()
    Set log_data = Nothing
    Set log_data = CreateObject("scripting.dictionary")
    Set dlog_title = Nothing
    Set dlog_title = CreateObject("scripting.dictionary")
    cur_row = 1
    log_data(1) = Split("", "|")
End Sub

Public Function get_log_count() As Long
    get_log_count = cur_row
End Function

'
'   判断系统日志中是否正确
'
Public Function is_right_log() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = gloEt.FindWs(ThisWorkbook, "log", 0, True)
    
    If gloEt.IsBlankSheet(ws) = True Then
        is_right_log = True
    Else
        Dim arr As Variant, i As Long, err_count As Long
        arr = gloEt.GetUsedRange(ws).value
        err_count = 0
        For i = 2 To UBound(arr)
            If InStr(CStr(arr(i, 3)), "有误") > 0 Or InStr(CStr(arr(i, 3)), "错误") Or InStr(CStr(arr(i, 3)), "错") > 0 Then
                err_count = err_count + 1
            End If
        Next
        
        If err_count = 0 Then
            is_right_log = True
        Else
            is_right_log = False
        End If
    End If
End Function


'
'   创建日记标题
'
'
Public Sub create_log_title(ParamArray strs() As Variant)
    'Debug.Print "varytpe:" & VarType(log_data(1))
    Call clear_log
    Set dlog_title = CreateObject("scripting.dictionary")
    log_data(1) = strs
    If UBound(strs) - LBound(strs) + 1 > cur_col Then cur_col = UBound(strs) - LBound(strs) + 1
    
    Dim str As Variant, i As Integer
    i = 1
    For Each str In strs
        dlog_title(ycS(str)) = i
        i = i + 1
    Next
    cur_col = i - 1
    
    'Debug.Print "类型:" & VarType(strs) & ";大小:" & UBound(strs) & ";varytpe:" & VarType(log_data(1)) & ";col:" & col
End Sub


'
'   添加一条记录，不限参数
'   插入一个字典对象
'
Public Sub add_log(ParamArray strs() As Variant)
    On Error GoTo errhanlde
    If UBound(strs) < 0 Then Exit Sub
    Dim arr() As Variant, i As Integer, key As Variant, e As Integer
    If UBound(strs) = 0 Then
        If VarType(strs(0)) = 9 Then
            Dim d As Object
            Set d = strs(0)
            For Each key In d.keys
                If Not dlog_title.exists(key) Then
                    cur_col = cur_col + 1
                    dlog_title(key) = cur_col
                End If
            Next
            
            ReDim arr(1 To cur_col) As Variant
            For Each key In d.keys
                arr(dlog_title(key)) = d(key)
            Next
        Else
            '就是组合的形式进来的那种
            Dim arr2 As Variant
            arr2 = strs(0)
            e = UBound(arr2)
            If e Mod 2 <> 0 Then e = e - 1
            For i = 1 To e Step 2
                If Not dlog_title.exists(arr2(i)) Then
                    cur_col = cur_col + 1
                    dlog_title(strs(i)) = cur_col
                End If
            Next i
            ReDim arr(1 To cur_col) As Variant
            For i = 1 To e Step 2
                arr(dlog_title(arr2(i))) = arr2(i + 1)
            Next i
        End If
    Else    ' 1 以上的数量，剔除奇数
        e = UBound(strs) + 1
        If e Mod 2 <> 0 Then e = e - 1
        For i = 0 To e - 1 Step 2
            If Not dlog_title.exists(strs(i)) Then
                cur_col = cur_col + 1
                dlog_title(strs(i)) = cur_col
            End If
        Next i
        ReDim arr(1 To cur_col) As Variant
        For i = 0 To e - 1 Step 2
            arr(dlog_title(strs(i))) = strs(i + 1)
        Next i
    End If
    log_data(cur_row + 1) = arr
    cur_row = cur_row + 1
    '判断标题是否需要更新
    If UBound(log_data(1)) < cur_col Then
        '说明新增了列，需要补充
        ReDim arr(1 To cur_col) As Variant
        For Each key In dlog_title.keys
            arr(dlog_title(key)) = key
        Next
        log_data(1) = arr
    End If
Exit Sub
errhanlde:
    Debug.Print "add_log error" & err.Description
    err.Clear
End Sub

'
'   不判断，直接插入
'
Public Sub add_log_stand(ParamArray strs() As Variant)
    If UBound(strs) < 0 Then Exit Sub
    Dim arr() As Variant, i As Integer, key As Variant
    
    log_data(cur_row + 1) = strs
    cur_row = cur_row + 1
    If UBound(strs) + 1 > cur_col Then cur_col = UBound(strs) + 1
End Sub

Public Sub export_log_to_log(Optional mode As String = "w", Optional start_row As Long = 1, Optional start_col As Long = 1)
    Dim ws As Worksheet
    Set ws = gloEt.FindWs(ThisWorkbook, "log", 0, True)
    Call export_log_to_ws(ws, mode, start_row, start_col)
End Sub

Private Function clear_job_name(job_name As String) As String
    Dim r As String
    r = replace(job_name, "金融基础数据采集（", "")
    r = replace(r, "）", "")
    r = replace(r, "金融基础数据-", "")
    
    clear_job_name = r
End Function

'
'   导出到指定工作表
'   1 - 清空，0 - 追加
'
Public Sub export_log_to_nwb_jcsj(mod_name As String, Optional jy_link As Integer = 0)
    On Error Resume Next
    ' 新建工作簿
    Dim wb As Workbook, fn As String, ws As Worksheet, arr_info As Variant
    Dim items As Variant, i As Long, j As Long, d_data As Object, d_temp As Object, arr As Variant, d_temp2 As Object
    Dim d_history As Object, key As String, date_str As String, d_org_area As Object
    
    items = log_data.items
    
    Set d_history = Nothing
    Set d_org_area = Nothing
    Set d_data = Nothing
    Set d_temp = Nothing
    Set d_org_area = CreateObject("scripting.dictionary")
    Set d_data = CreateObject("scripting.dictionary")
    Set d_history = CreateObject("scripting.dictionary")
    ' Call glo_log.create_log_title("模块", "工作簿", "工作表", "校验结果类型", "行位置", "行名称", "列位置", "列名称", "批注(规则描述)", "校验区域类型", "校验字段", "校验值", "详细描述", "列4", "列5", "说明")
    
    ' 读取历史说明
    Set ws = gloEt.FindWs(ThisWorkbook, "历史核查表审核结果", 0, True)
    arr = gloEt.GetUsedRange(ws).value
    For i = 2 To UBound(arr)
        Set d_temp = CreateObject("scripting.dictionary")
        For j = 1 To UBound(arr, 2)
            d_temp(CStr(arr(1, j))) = CStr(arr(i, j))
        Next
        key = d_temp("机构名称") & d_temp("任务名称") & d_temp("表单名称") & d_temp("校验结果类型") & d_temp("校验区域类型") & d_temp("校验字段") & d_temp("批注(规则描述)") & d_temp("行名称") & d_temp("列名称")
        Set d_history(key) = d_temp
        Set d_temp = Nothing
    Next
    
    '读取地区参照
    Set ws = gloEt.FindWs(ThisWorkbook, "基础数据参照表", 0, True)
    arr = gloEt.GetUsedRange(ws).value
    For i = 2 To UBound(arr)
        Set d_org_area(CStr(arr(i, 1))) = CStr(arr(i, 2))
    Next
    
    For i = 1 To UBound(items)
        Set d_temp = CreateObject("scripting.dictionary")
        For j = 0 To UBound(items(i))
            d_temp(CStr(items(0)(j))) = CStr(items(i)(j))
        Next

        arr_info = Split(gloFt.getShortFileName(CStr(d_temp("工作簿"))), "_")  ' 获取了基本信息
        '针对性修改
        d_temp("文件名") = d_temp("工作簿")
        d_temp("机构名称") = CStr(arr_info(0))
        d_temp("任务名称") = clear_job_name(CStr(arr_info(1)))
        d_temp("表单名称") = d_temp("工作表")
        date_str = replace(CStr(arr_info(2)), "校验结果及报送说明", "")
        If ycIsDate(date_str) = False Then
            date_str = replace(CStr(arr_info(3)), "校验结果及报送说明", "")
        End If
        d_temp("数据日期") = date_str
        d_temp("说明") = ""
        
        key = d_temp("机构名称") & d_temp("任务名称") & d_temp("表单名称") & d_temp("校验结果类型") & d_temp("校验区域类型") & d_temp("校验字段") & d_temp("批注(规则描述)") & d_temp("行名称") & d_temp("列名称")
        If d_history.exists(key) Then
            Set d_temp2 = d_history(key)
            d_temp("历史是否通过") = d_temp2("历史是否通过")
            d_temp("历史说明") = d_temp2("历史说明")
            d_temp("历史判断") = d_temp2("历史判断")
            Set d_temp2 = Nothing
        End If
        
        If d_org_area.exists(CStr(d_temp("机构名称"))) Then
            d_temp("地区") = d_org_area(CStr(d_temp("机构名称")))
        End If
        
        Set d_data(i) = d_temp
        
        Set d_temp = Nothing
    Next
    
    
    Set wb = Workbooks.add
    fn = ThisWorkbook.path & "\" & "报表工具-校验结果-" & mod_name & " " & Format(now, "yyyymmdd hhmmss") & ".xls"
    
    Set ws = wb.Worksheets(1)
    ws.name = "校验"
    
    ws.Cells.NumberFormat = "@"
    ws.Cells.Font.Size = 12
    ws.Cells.Font.name = "微软雅黑"
    ws.Range("N:N").NumberFormat = "0.00"
    
    arr = dictsTolists(Split(gloCft.GetConfig("基础数据日记输出"), ","), d_data, True)
    
    ws.Range("A1").Resize(UBound(arr), UBound(arr, 2)).value = arr
    
    ' 增加超链接
    If jy_link <> 0 Then
        For i = 2 To UBound(arr)
            ws.Hyperlinks.add anchor:=ws.Cells(i, 14), address:=CStr(arr(i, 1)), SubAddress:=CStr(arr(i, 5)) & "!" & ws.Cells(ycV(arr(i, 9)), ycV(arr(i, 11))).address
        Next

    End If
    
    ws.Cells.WrapText = False
    Call AddMacroCode_Check_jcsj(wb)
    Set d_data = Nothing
    Set d_history = Nothing
    Set d_org_area = Nothing
    
    '冻结首行
    ws.Parent.Windows(1).FreezePanes = False
    ws.Parent.Windows(1).SplitRow = 1
    ws.Parent.Windows(1).FreezePanes = True
    wb.SaveAs fn
    
    If err.Number <> 0 Then
        MsgBox "导出校验结果的时候存在异常。" & err.Description
        err.Clear
    End If
End Sub

'给新的工作簿增加宏代码
Private Sub AddMacroCode_Check_jcsj(wbDes As Workbook)
    Dim MyCodes As String, fullName As String
    
    On Error GoTo errhandle
    
    '插入启动菜单
    MyCodes = gloCft.GetConfig("基础数据校验文件菜单宏")
    With wbDes.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines 2, MyCodes
    End With
    
    MyCodes = gloCft.GetConfig("基础数据校验文件工作宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work"
        .CodeModule.InsertLines 2, MyCodes
    End With
Exit Sub

errhandle:
    err.Clear
    MsgBox "说明文件功能受限。需要在EXCEL菜单 工具-宏－安全性－可靠发行商　中勾选【信任对于VB项目的访问】"
End Sub

'
'  按指定格式转换成数组
'
Private Function dictsTolists(titles As Variant, ByRef d_data As Object, Optional is_dict_title As Boolean = False) As Variant
    On Error Resume Next
    Dim c As Long, arr As Variant, i As Long, j As Long, v As Variant, v2 As Variant, d_title As Object, k2 As Variant
    c = 1 + d_data.count
    
    ReDim arr(1 To c, 1 To UBound(titles) + 1) As Variant
    Set d_title = CreateObject("scripting.dictionary")
    ' 标题
    For j = 0 To UBound(titles)
        arr(1, j + 1) = CStr(titles(j))
        d_title(CStr(titles(j))) = j + 1
    Next
    
    If is_dict_title = False Then
        v = d_data.items
        For i = 0 To UBound(v)
            v2 = v(i).items
            For j = 0 To UBound(v2)
                arr(i + 2, j + 1) = CStr(v2(j))
            Next
        Next
    Else
        v = d_data.items
        For i = 0 To UBound(v)
            v2 = v(i).items
            k2 = v(i).keys
            For j = 0 To UBound(k2)
                If d_title.exists(CStr(k2(j))) Then
                    arr(i + 2, d_title(CStr(k2(j)))) = CStr(v2(j))
                End If
            Next
        Next
    End If
    
    dictsTolists = arr
End Function

'
'   导出到新工作文件
'
Public Sub export_log_to_nwb(mod_name As String)
    ' 新建工作簿
    Dim wb As Workbook, fn As String, ws As Worksheet
    Set wb = Workbooks.add
    fn = ThisWorkbook.path & "\" & "报表工具-校验结果-" & mod_name & " " & Format(now, "yyyymmdd hhmmss") & ".xls"
    wb.SaveAs fn
    Set ws = wb.Worksheets(1)
    ws.name = "校验"
    Call export_log_to_ws(ws)
End Sub

'
'   导出到指定工作表
'   1 - 清空，0 - 追加
'
Public Sub export_log_to_ws(ws As Worksheet, Optional mode As String = "w", Optional start_row As Long = 1, Optional start_col As Long = 1, Optional is_insert As Boolean = True)
    On Error Resume Next
    Dim items As Variant
    items = log_data.items
    If UBound(items) <= 0 Then Exit Sub
    
    Dim arr() As Variant, col As Integer, row As Integer, s_row As Long
    If mode = "w" Then
        ws.Cells.Clear
        s_row = start_row
    Else
        If gloEt.IsBlankSheet(ws) Then
            s_row = start_row
        Else
            s_row = gloEt.GetCurRows(ws) + 1
            s_row = s_row + 1 '间隔一行
        End If
    End If
    '超过列数
    col = cur_col
    If col + start_col > ws.Cells.columns.count Then col = ws.Cells.columns.count - start_col  '剩余多少列可以使用
    
    '超过行数
    row = cur_row
    If row + s_row + 5 > ws.Cells.rows.count Then row = ws.Cells.rows.count - s_row - 5 '剩余多少行可以使用
    
    ReDim arr(1 To row, 1 To col) As Variant
    
    Dim args As Variant, i As Integer, j As Integer, arg As Variant
    i = 1
    For Each args In items
        j = 1
        If VarType(args) <> vbEmpty Then
            For Each arg In args
                '复制列
                arr(i, j) = arg
                j = j + 1
            Next
        End If
        i = i + 1
    Next

    '如果是插入模式，则起始列调整未1
    If is_insert = True Then
        ws.rows(1).Resize(row + 5).Insert
        s_row = 1
    End If
    '输出到工作表
    ws.Cells(s_row, start_col).Resize(row, col).Clear
    ws.Cells(s_row, start_col).Resize(row, col).NumberFormat = "@"
    If col > 11 Then
        ws.Cells(s_row, col - 4).Resize(row, 4).NumberFormat = "0.0000"
        
    End If
    
    ws.Cells(s_row, start_col).Resize(1, col).Interior.colorindex = 6
    ws.Cells.WrapText = False
    'ws.Cells.ShrinkToFit = False
    ws.Cells.Font.Size = 12
    ws.Cells.Font.name = "微软雅黑"
    ws.Cells(s_row, start_col).Resize(row, col).value = arr
    ws.Cells.WrapText = False
    
'    If InStr(CStr(arr(2, 2)), "金融基础数据") > 0 Or InStr(CStr(arr(2, 2)), "数据核对表") > 0 Then
'        '
'        Dim arr_bc As Variant, d As Object, arr_jc As Variant, ws_jc As Worksheet
'        Set ws_jc = gloEt.FindWs(ThisWorkbook, "基础数据参照表", 0)
'        If Not ws_jc Is Nothing Then
'            If gloEt.IsBlankSheet(ws_jc) = False Then
'                ReDim arr_bc(1 To UBound(arr), 1 To 3) As Variant
'                Set d = CreateObject("Scripting.Dictionary")
'                arr_jc = gloEt.GetUsedRange(ws_jc).value
'                For i = 2 To UBound(arr_jc)
'                    d(CStr(arr_jc(i, 1))) = CStr(arr_jc(i, 2))
'                Next
'
'                '填充
'                arr_bc(1, 1) = "机构名称": arr_bc(1, 2) = "所属地区": arr_bc(1, 3) = "历史说明"
'                For i = 2 To UBound(arr)
'                    arr_bc(i, 1) = ycLeft(CStr(arr(i, 2)), "_", 1, 0)
'                    If d.exists(CStr(arr_bc(i, 1))) Then arr_bc(i, 2) = CStr(d(CStr(arr_bc(i, 1))))
'                Next
'                '输出
'                ws.Cells(s_row, col + 1).Resize(row, 3).NumberFormat = "@"
'                ws.Cells(s_row, col + 1).Resize(row, 3).value = arr_bc
'            End If
'        End If
'    End If
    ws.Activate
    ws.Range("A1").Select
End Sub


'
'   给上一条记录添加颜色
'
Public Sub AddColorDebug(ws As Worksheet, color As Integer, Optional c As Integer = 0)
    With ws
        If c = 0 Then
            .rows(2).Interior.colorindex = color
        Else
            .Cells(2, c).Interior.colorindex = color
        End If
    End With
End Sub

Public Sub AddDeHypeLinks(ws As Worksheet, col As Integer, path As String, showName As String)
    With ws
       .Cells(2, col).Hyperlinks.add anchor:=.Cells(2, col), address:=path, SubAddress:="", ScreenTip:="", TextToDisplay:=showName
    End With
End Sub

Public Sub SetDeFoucs(ws As Worksheet)
    On Error GoTo FoucsErrHandle
    Dim wb As Workbook
    Set wb = ws.Parent
    wb.Activate
    ws.Activate
    ws.Range("A1").Activate
Exit Sub

FoucsErrHandle:
    Debug.Print "设置焦点日志工作表焦点转移功能失败，不影响其他使用。" & err.Description
    err.Clear
End Sub


