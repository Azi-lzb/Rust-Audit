Attribute VB_Name = "台账拆分"
Option Explicit

Sub ClearShts()
    'EH技术论坛。VBA编程学习与实践。看见星光
    Dim sht As Worksheet, Mypath$
    Application.DisplayAlerts = False
    '取消显示系统警告和消息，避免重名工作簿无法保存。当有重名工作簿时，会直接覆盖保存。
    Application.ScreenUpdating = False
    '取消屏幕刷新
    For Each sht In Worksheets
        If InStr(sht.name, "拆分") Then
        '遍历工作表
            If Worksheets.count > 1 Then
                sht.Delete
            End If
        End If
    Next
    Application.ScreenUpdating = True '恢复屏幕刷新
    Application.DisplayAlerts = True '恢复显示系统警告和消息
End Sub

Sub SplitShts()
    Call ClearShts
    
    Dim d As Object, sht As Worksheet
    Dim aData, aResult, aTemp, aKeys, i&, j&, k&, x&
    Dim rngData As Range, rngGist As Range
    Dim lngTitleCount&, lngGistCol&, lngColCount&
    Dim rngFormat As Range
    Dim strKey As String
    Set d = CreateObject("scripting.dictionary")
    Set rngGist = Application.InputBox("请框选拆分依据列！只能选择单列单元格区域！", title:="提示", Type:=8)
    '========用户选择的拆分依据列
    lngGistCol = rngGist.Column
    '========拆分依据列的列标
    lngTitleCount = val(Application.InputBox("请输入总表标题行的行数？"))
    '========用户设置总表的标题行数
    If lngTitleCount < 0 Then MsgBox "标题行数不能为负数，程序退出。": Exit Sub
    Set rngData = ActiveSheet.UsedRange
    '========总表的数据区域
    Set rngFormat = ActiveSheet.Cells
    '========总表的单元格集用于粘贴总表格式
    aData = rngData.Value
    lngGistCol = lngGistCol - rngData.Column + 1
    '========计算依据列在数组中的位置
    lngColCount = UBound(aData, 2)
    '========数据源的列数
    For i = lngTitleCount + 1 To UBound(aData)
        If aData(i, lngGistCol) = "" Then aData(i, lngGistCol) = "单元格空白"
        strKey = aData(i, lngGistCol)
    '========统一转换为字符串格式
        If Not d.exists(strKey) Then
    '========字典中不存在关键字时将行号装入字典
            d(strKey) = i
        Else
            d(strKey) = d(strKey) & "," & i
    '========如果字段存在关键字则合并行号
        End If
    Next
    Application.DisplayAlerts = False

    aKeys = d.keys
    '========字典的key集
    Application.ScreenUpdating = False
    For i = 0 To UBound(aKeys)
        If aKeys(i) <> "" Then
            aTemp = Split(d(aKeys(i)), ",")
    '========取出item里储存的行号
            ReDim aResult(1 To UBound(aTemp) + 1, 1 To lngColCount)
    '========声明放置结果的数组aResult
            k = 0
            For x = 0 To UBound(aTemp)
                k = k + 1
                For j = 1 To lngColCount
                    aResult(k, j) = aData(aTemp(x), j)
                Next
            Next
            With Worksheets.add(, Sheets(Sheets.count))
    '========新建一个工作表
                .name = "拆分 " & aKeys(i)
                .[a1].Resize(UBound(aData), lngColCount).NumberFormat = "@"
    '========设置单元格为文本格式
                If lngTitleCount > 0 Then .[a1].Resize(lngTitleCount, lngColCount) = aData
    '========标题行
                .[a1].Offset(lngTitleCount, 0).Resize(k, lngColCount) = aResult
    '========数据
                rngFormat.Copy
                .[a1].PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '========复制粘贴总表的格式
                .[a1].Offset(lngTitleCount + k, 0).Resize(UBound(aData) - k - lngTitleCount, 1).EntireRow.Delete
    '========删除多余的格式单元格
                .[a1].Select
            End With
        End If
    Next
    rngData.Parent.Activate
    '========激活总表
    Application.ScreenUpdating = True
    Set d = Nothing
    Set rngData = Nothing
    Set rngGist = Nothing
    Set rngFormat = Nothing
    Erase aData: Erase aResult
    MsgBox "数据拆分完成！"
End Sub

Sub Newbooks()
    'EH技术论坛。VBA编程学习与实践。看见星光
    Dim sht As Worksheet, Mypath$
    With Application.FileDialog(msoFileDialogFolderPicker)
   '选择保存工作薄的文件路径
        .AllowMultiSelect = False
        '不允许多选
        If .Show Then
            Mypath = .SelectedItems(1)
            '读取选择的文件路径
        Else
            Exit Sub
            '如果没有选择保存路径，则退出程序
        End If
    End With
    If right(Mypath, 1) <> "\" Then Mypath = Mypath & "\"
    Application.DisplayAlerts = False
    '取消显示系统警告和消息，避免重名工作簿无法保存。当有重名工作簿时，会直接覆盖保存。
    Application.ScreenUpdating = False
    '取消屏幕刷新
    For Each sht In Worksheets
        If InStr(sht.name, "拆分") Then
        '遍历工作表
            sht.Copy
            '复制工作表，工作表单纯复制后，会成为活动工作薄
            With ActiveWorkbook
                .SaveAs Mypath & replace(sht.name, "拆分", "") & ".xls"
                '保存活动工作薄到指定路径下，以默认文件格式
                .Close True '关闭工作薄并保存
            End With
        End If
    Next
    MsgBox "处理完成。", , "提醒"
    Application.ScreenUpdating = True '恢复屏幕刷新
    Application.DisplayAlerts = True '恢复显示系统警告和消息
End Sub

Sub CancelMerge(ws As Worksheet, lngTitleCount As Long, Optional mode As Integer = 1)
    Dim r As Long, c As Long, rnT As Range, t As Variant, rn As Range, add As String
    '指标一样的话
    
    '比如如果是 $a2:c4,标题行是2的话
    add = gloEt.GetDeleteTileArea(ws, lngTitleCount, mode)
    If add = "" Then Exit Sub
    Set rn = ws.Range(add)
    If IsNull(rn.MergeCells) = False Then
        If CBool(rn.MergeCells) = True Then
            GoTo jixu
        End If
        
        Exit Sub    '没有合并单元格则退出
    End If

jixu:
    '去除合并单元格
    For Each rnT In rn
        If rnT.MergeArea.count > 1 Then
            t = rnT.MergeArea(1).Value
            add = rnT.MergeArea.Address
            rnT.UnMerge
            ws.Range(add).Value = t
        End If
    Next
End Sub

Sub SplitShts2()
    On Error Resume Next
    Call ClearShts
    
    ThisWorkbook.Worksheets("台账面板").Activate
    Dim Mypath$
    
    Mypath = gloFt.GetOutputFolder
    If Mypath = "" Then Exit Sub
    
    If right(Mypath, 1) <> "\" Then Mypath = Mypath & "\"
    
    Dim d As Object, sht As Worksheet, shtT As Worksheet
    Dim aData, aResult, aTemp, aKeys, i&, j&, k&, x&
    Dim rngData As Range, rngGist As Range
    Dim lngTitleCount&, lngGistCol&, lngColCount&
    Dim rngFormat As Range
    Dim strKey As String
    Set d = CreateObject("scripting.dictionary")
    Set rngGist = Application.InputBox("请框选拆分依据列！只能选择单列单元格区域！", title:="提示", Type:=8)
    '========用户选择的拆分依据列
    lngGistCol = rngGist.Column
    '========拆分依据列的列标
    lngTitleCount = val(Application.InputBox("请输入总表标题行的行数？"))
    '========用户设置总表的标题行数
    If lngTitleCount < 0 Then MsgBox "标题行数不能为负数，程序退出。": Exit Sub
    '这里考虑取消合并单元格？这里仅使用currentRegion
    Call CancelMerge(ActiveSheet, lngTitleCount, 1)
    
    Set rngData = ActiveSheet.UsedRange
    '========总表的数据区域
    Set rngFormat = ActiveSheet.Cells
    '========总表的单元格集用于粘贴总表格式
    aData = rngData.Value
    lngGistCol = lngGistCol - rngData.Column + 1
    '========计算依据列在数组中的位置
    lngColCount = UBound(aData, 2)
    '========数据源的列数
    For i = lngTitleCount + 1 To UBound(aData)
        If aData(i, lngGistCol) = "" Then aData(i, lngGistCol) = "单元格空白"
        strKey = aData(i, lngGistCol)
    '========统一转换为字符串格式
        If Not d.exists(strKey) Then
    '========字典中不存在关键字时将行号装入字典
            d(strKey) = i
        Else
            d(strKey) = d(strKey) & "," & i
    '========如果字段存在关键字则合并行号
        End If
    Next
    Application.DisplayAlerts = False
    
    aKeys = d.keys
    '========字典的key集
    Application.ScreenUpdating = False
    '新建一个工作表
    Set shtT = Worksheets.add(, Sheets(Sheets.count))
    For i = 0 To UBound(aKeys)
        If aKeys(i) <> "" Then
            aTemp = Split(d(aKeys(i)), ",")
    '========取出item里储存的行号
            ReDim aResult(1 To UBound(aTemp) + 1, 1 To lngColCount)
    '========声明放置结果的数组aResult
            k = 0
            For x = 0 To UBound(aTemp)
                k = k + 1
                For j = 1 To lngColCount
                    aResult(k, j) = aData(aTemp(x), j)
                Next
            Next
            
            With shtT
    '========新建一个工作表
                .Cells.Clear
                .name = "拆分 " & aKeys(i)
                .[a1].Resize(UBound(aData), lngColCount).NumberFormat = "@"
    '========设置单元格为文本格式
                If lngTitleCount > 0 Then .[a1].Resize(lngTitleCount, lngColCount) = aData
    '========标题行
                .[a1].Offset(lngTitleCount, 0).Resize(k, lngColCount) = aResult
    '========数据
                rngFormat.Copy
                .[a1].PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '========复制粘贴总表的格式
                .[a1].Offset(lngTitleCount + k, 0).Resize(UBound(aData) - k - lngTitleCount, 1).EntireRow.Delete
    '========删除多余的格式单元格
                .[a1].Select
                
                .Copy
                '复制工作表，工作表单纯复制后，会成为活动工作薄
                With ActiveWorkbook
                    .SaveAs Mypath & CStr(aKeys(i)) & ".xls"
                    '保存活动工作薄到指定路径下，以默认文件格式
                    .Close False '关闭工作薄并保存
                End With
                '.Delete
                Application.CutCopyMode = False
            End With
        End If
    Next
    shtT.Delete
    rngData.Parent.Activate
    '========激活总表
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set d = Nothing
    Set rngData = Nothing
    Set rngGist = Nothing
    Set rngFormat = Nothing
    Erase aData: Erase aResult
    MsgBox "数据拆分完成，" & "共拆分出" & CStr(i) & "个文件。"
End Sub

