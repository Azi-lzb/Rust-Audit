Attribute VB_Name = "面板辅助"
Option Explicit

Private clsmymenu As CMySysMenu

'
'   追加汇总模板文件作为候选
'
Public Sub LoadModFilesSelect()
    Dim ws As Worksheet, ws2 As Worksheet, files As String, arr_file As Variant, i As Long, cur_row As Long
    
    Set ws2 = gloEt.FindWs(ThisWorkbook, "模板列表", 0, True)
    
    ws2.Cells.NumberFormat = "@"
    ws2.Range("b1").value = "模板文件名"
    ws2.Range("a1").value = "模板文件简称"
    
    files = gloFt.SelectFiles("选择批量追加的模板文件")
    
    If files = "" Then
        MsgBox "未选择源文件"
    Else
        arr_file = Split(files, ",")
        
        For i = 0 To UBound(arr_file)
            cur_row = ws2.Cells(ws2.Cells.rows.count, 1).End(xlUp).row + 1
            ws2.Cells(cur_row, 2).value = CStr(arr_file(i))
            ws2.Cells(cur_row, 1).value = gloFt.getShortFileName2(CStr(arr_file(i)))
        Next
    End If
End Sub

'
'   选择汇总模板文件
'
Public Sub SelHzModFile()
    'Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelModFile
    Set clsmymenu = Nothing
    'Call gloCft.OpenSomething
End Sub



Public Sub SelHzModFile_form_list()
    Dim ws2 As Worksheet
    
    Set ws2 = gloEt.FindWs(ThisWorkbook, "模板列表", 0, True)
    
    ws2.Cells.NumberFormat = "@"
    ws2.Range("b1").value = "模板文件名"
    ws2.Range("a1").value = "模板文件简称"
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelModFile_form_list
    Set clsmymenu = Nothing
    
    
End Sub

Public Sub SelExtendFile()
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelExtendFile
    Set clsmymenu = Nothing
End Sub



'
'   选择汇总源数据
'
Public Sub SelHzSrcFiles()
    'Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelSrcFiles
    Set clsmymenu = Nothing
    'Call gloCft.OpenSomething
End Sub

'
'   批量选择汇总源数据
'
Public Sub SelHzSrcFilesInFolder()
    'Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelSrcFilesByFolder
    'Call clsmymenu.SelSrcFilesInFolder
    Set clsmymenu = Nothing
    'Call gloCft.OpenSomething
End Sub

'
'   批量选择汇总源数据
'
Public Sub SelHzSrcFilesInFolder2()
    'Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SelSrcFilesByFolderWithFolder
    Set clsmymenu = Nothing
    'Call gloCft.OpenSomething
End Sub
