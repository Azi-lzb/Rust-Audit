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
        arr = gloEt.GetUsedRange(ws).Value
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
    If col + start_col > ws.Cells.Columns.count Then col = ws.Cells.Columns.count - start_col  '剩余多少列可以使用
    
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
    ws.Cells(s_row, start_col).Resize(row, col).Value = arr
    ws.Cells.WrapText = False
    
    If InStr(CStr(arr(2, 2)), "金融基础数据") > 0 Then
        '
        Dim arr_bc As Variant, d As Object, arr_jc As Variant, ws_jc As Worksheet
        Set ws_jc = gloEt.FindWs(ThisWorkbook, "基础数据参照表", 0)
        If Not ws_jc Is Nothing Then
            If gloEt.IsBlankSheet(ws_jc) = False Then
                ReDim arr_bc(1 To UBound(arr), 1 To 2) As Variant
                Set d = CreateObject("Scripting.Dictionary")
                arr_jc = gloEt.GetUsedRange(ws_jc).Value
                For i = 2 To UBound(arr_jc)
                    d(CStr(arr_jc(i, 1))) = CStr(arr_jc(i, 2))
                Next
                
                '填充
                arr_bc(1, 1) = "机构名称": arr_bc(1, 2) = "所属地区"
                For i = 2 To UBound(arr)
                    arr_bc(i, 1) = ycLeft(CStr(arr(i, 2)), "_", 1, 0)
                    If d.exists(CStr(arr_bc(i, 1))) Then arr_bc(i, 2) = CStr(d(CStr(arr_bc(i, 1))))
                Next
                '输出
                ws.Cells(s_row, col + 1).Resize(row, 2).NumberFormat = "@"
                ws.Cells(s_row, col + 1).Resize(row, 2).Value = arr_bc
            End If
        End If
    End If
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
       .Cells(2, col).Hyperlinks.add Anchor:=.Cells(2, col), Address:=path, SubAddress:="", ScreenTip:="", TextToDisplay:=showName
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


