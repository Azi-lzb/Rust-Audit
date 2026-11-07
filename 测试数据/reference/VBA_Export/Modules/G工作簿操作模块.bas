Attribute VB_Name = "G工作簿操作模块"
Option Explicit

Private clsmymenu As CMySysMenu
Private bErrAvailable As Boolean   '用于存储工作簿是否正确
Private srcfiles As Variant '源数据二维数组
Private dsrcfiles As Object '源数据字典

Private Sub EraseSomething()
    Set clsmymenu = Nothing
    Set dsrcfiles = Nothing
End Sub

Private Sub SetMenuTitleValue(title As String, i As Long, value As String)
    srcfiles(i, clsmymenu.GetMenuTitleCol(title)) = value
End Sub

Private Function GetSrcFiles() As Boolean
    srcfiles = clsmymenu.GetSysMenuAreaValue
    Set dsrcfiles = clsmymenu.GetSysMenuSrcs
    
    If dsrcfiles.count = 0 Then
        GetSrcFiles = False
    Else
        GetSrcFiles = True
    End If
End Function

'
'   将多个工作簿中的工作表汇集在一个工作表
'
Public Sub GuijiWb()
    Call gloCft.CloseSomething
    
    On Error GoTo errhandle
    Set clsmymenu = New CMySysMenu
    clsmymenu.SetSysWorkSheet ("执行面板")
    If GetSrcFiles = False Then Exit Sub
    
    Dim fnOut As String
    fnOut = gloFt.GetOutputFile
    If fnOut = "" Then
        Exit Sub
    End If
    
    '初始化面板
    Call clsmymenu.SetMenuTitle("校验结果", "汇总结果")

    
    '判断是否有列指标
    Dim wsdes As Worksheet, wbDes As Workbook, arrSplit As Variant, arrSplit2 As Variant, hzNum As Long, bsplit As Boolean, i As Long, j As Integer
    Set wbDes = gloEt.GetWb(fnOut)
    If wbDes Is Nothing Then
        Set wbDes = Workbooks.add
        wbDes.SaveAs fnOut
    End If

    Dim wbsrc As Workbook, wsSrc As Worksheet, curRow As Long, arrSrc As Variant, k As Integer, t_wsName As String
    
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    '初始化源数据显示
    Call clsmymenu.ClearMenuTitle
    
    arrSplit = Split(clsmymenu.GetConfig("设置不参与的工作表"), ";")
    arrSplit2 = Split(clsmymenu.GetConfig("设置参与的工作表"), ";")
    
    hzNum = 0
    For i = 2 To UBound(srcfiles) Step 1
        If Not dsrcfiles.exists(CStr(srcfiles(i, 2))) Then GoTo NextHzSrcFiles
        
        Set wbsrc = gloEt.GetWb(CStr(srcfiles(i, 2)))
         
        If wbsrc Is Nothing Then
            Call SetMenuTitleValue("汇总结果", i, "F")
            Call gloDe.InsertNewDebug("工作簿归集", CStr(srcfiles(i, 2)), "-", "不参与汇总", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        
        For Each wsSrc In wbsrc.Worksheets
            If wbDes.Worksheets.count > 254 Then
                MsgBox "工作簿数量已达到" & CStr(wbDes.Worksheets.count) & "停止汇总"
                Call gloCft.OpenSomething
                Exit Sub
            End If
            
            If UBound(arrSplit2) = -1 Then
                bsplit = True
                For j = 0 To UBound(arrSplit) Step 1
                    If InStr(CStr(arrSplit(j)), wsSrc.name) > 0 Then
                        bsplit = False
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit(j)) & "*" Then
                        bsplit = False
                        Exit For
                    End If
                Next j
            Else
                bsplit = False
                For j = 0 To UBound(arrSplit2) Step 1
                    If InStr(CStr(arrSplit2(j)), wsSrc.name) > 0 Then
                        bsplit = True
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit2(j)) & "*" Then
                        bsplit = True
                        Exit For
                    End If
                Next j
            End If
            
            If UBound(arrSplit2) = -1 And UBound(arrSplit) = -1 Then bsplit = True
            
            '如果没有被筛选，那么则进行汇总,空表也汇总，说不定留底需要
            If bsplit = True Then
                
                arrSrc = wsSrc.UsedRange.value
                If VarType(arrSrc) = vbString Then
                    Dim temp As Variant
                    temp = arrSrc
                    ReDim arrSrc(1 To 1, 1 To 1) As Variant
                    arrSrc(1, 1) = temp
                End If
                
                '判断工作簿中是否已经存在该表
                k = 0
                Do While True
                    If k = 0 Then
                        t_wsName = gloFt.getShortFileName(wbsrc.name) & "-" & wsSrc.name
                    Else
                        t_wsName = gloFt.getShortFileName(wbsrc.name) & "-" & wsSrc.name & "(" & CStr(k) & ")"
                    End If
                    If Len(t_wsName) > 25 Then
                        t_wsName = right(t_wsName, 25)
                    End If
                    Set wsdes = gloEt.FindWs(wbDes, t_wsName, 0)
                    If wsdes Is Nothing Then
                        Set wsdes = wbDes.Worksheets.add
                        wsdes.name = t_wsName
                        Exit Do
                    End If
                    k = k + 1
                Loop
                If Not wsdes Is Nothing Then
                    If gloEt.IsBlankSheet(wsSrc) = False Then Call gloEt.CopyAndPaste(wsSrc, wsdes)
                    'wsDes.Range(wsSrc.UsedRange.address).value = arrSrc
                    Call gloDe.InsertNewDebug("工作簿归集", wbsrc.name, wsSrc.name, "汇总成功", "归集名为:" & wsdes.name)
                Else
                    Call gloDe.InsertNewDebug("工作簿归集", wbsrc.name, wsSrc.name, "汇总失败")
                End If
            Else
                Call gloDe.InsertNewDebug("工作簿归集", wbsrc.name, wsSrc.name, "不参与汇总")
            End If
        Next
        Call SetMenuTitleValue("汇总结果", i, "T")
        Call gloDe.InsertNewDebug("工作簿归集", wbsrc.name, "-", "汇总成功")
        Call gloDe.AddColorDebug(15)
        hzNum = hzNum + 1

        wbsrc.Close True
NextHzSrcFiles:

    If i > UBound(srcfiles) Then Exit For
    Next i

    Call gloDe.InsertNewDebug("工作簿归集", wsdes.Parent.name, "-", "汇总结束", "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbDes.Save
    Call EraseSomething
    Call gloCft.OpenSomething
Exit Sub
errhandle:

    If Not wbsrc Is Nothing And Not wsSrc Is Nothing Then
        Call gloDe.InsertNewDebug("工作簿归集", wbsrc.name, wsSrc.name, "汇总失败" & err.Description)
    Else
        Call gloDe.InsertNewDebug("工作簿归集", "-", "-", "汇总出现异常，需检查汇总结果" & err.Description)
    End If
    err.Clear
    GoTo NextHzSrcFiles
End Sub

'
'   将工作簿的表拆分出去
'
Public Sub ChaifenWb()
    Call gloCft.CloseSomething
    
    On Error GoTo errhandle
    Set clsmymenu = New CMySysMenu
    clsmymenu.SetSysWorkSheet ("执行面板")
    If GetSrcFiles = False Then Exit Sub
    
    Dim folder As String
    folder = gloFt.GetOutputFolder
    
    If folder = "" Then
        Call gloCft.OpenSomething
        Exit Sub
    End If
    
    '初始化面板
    Call clsmymenu.SetMenuTitle("校验结果", "拆分结果")
    
    '判断是否有列指标
    Dim arrSplit As Variant, arrSplit2 As Variant, hzNum As Long, bsplit As Boolean, i As Long, j As Integer
    Dim wbsrc As Workbook, wsSrc As Worksheet, curRow As Long, kfn As String
    
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    '初始化源数据显示
    Call clsmymenu.ClearMenuTitle
    
    arrSplit = Split(clsmymenu.GetConfig("设置不参与的工作表"), ";")
    arrSplit2 = Split(clsmymenu.GetConfig("设置参与的工作表"), ";")
    hzNum = 0
    For i = SYSSTARTROW + 1 To UBound(srcfiles) Step 1
        '这么加就能对接了
        If Not dsrcfiles.exists(CStr(srcfiles(i, 2))) Then GoTo NextHzSrcFiles
        
        Set wbsrc = gloEt.GetWb(CStr(srcfiles(i, 2)))
         
        If wbsrc Is Nothing Then
            Call SetMenuTitleValue("拆分结果", i, "F")
            Call gloDe.InsertNewDebug("工作簿拆分", CStr(srcfiles(i, 2)), "-", "不参与拆分", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        
        For Each wsSrc In wbsrc.Worksheets
            
            If UBound(arrSplit2) = -1 Then
                bsplit = True
                For j = 0 To UBound(arrSplit) Step 1
                    If InStr(CStr(arrSplit(j)), wsSrc.name) > 0 Then
                        bsplit = False
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit(j)) & "*" Then
                        bsplit = False
                        Exit For
                    End If
                Next j
            Else
                bsplit = False
                For j = 0 To UBound(arrSplit2) Step 1
                    If InStr(CStr(arrSplit2(j)), wsSrc.name) > 0 Then
                        bsplit = True
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit2(j)) & "*" Then
                        bsplit = True
                        Exit For
                    End If
                Next j
            End If
            
            If UBound(arrSplit2) = -1 And UBound(arrSplit) = -1 Then bsplit = True
                 
            '如果没有被筛选，那么则进行拆分，空表也拆分
            If bsplit = True Then
                ' 工作簿名称 + 工作表名称
                If Len(gloFt.getShortFileName(wbsrc.name) & " " & wsSrc.name & ".xls") > 180 Then
                    kfn = folder & "\" & right(gloFt.getShortFileName(wbsrc.name) & " " & wsSrc.name & ".xls", 180)
                Else
                    kfn = folder & "\" & gloFt.getShortFileName(wbsrc.name) & " " & wsSrc.name & ".xls"
                End If
                
                If gloEt.SaveSheetAsWorkbook(kfn, wsSrc) = True Then
                    Call gloDe.InsertNewDebug("工作簿拆分", wbsrc.name, wsSrc.name, "拆分成功", "拆分为：" & gloFt.getShortFileName(kfn))
                End If
            Else
                Call gloDe.InsertNewDebug("工作簿拆分", wbsrc.name, wsSrc.name, "不参与拆分")
            End If
        Next
        Call SetMenuTitleValue("拆分结果", i, "T")
        Call gloDe.InsertNewDebug("工作簿拆分", wbsrc.name, "-", "拆分成功")
        Call gloDe.AddColorDebug(15)
        hzNum = hzNum + 1

        wbsrc.Close True
NextHzSrcFiles:
    Next i

    Call gloDe.InsertNewDebug("工作簿拆分", "-", "-", "拆分结束", "共拆分" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Call EraseSomething
    Call gloCft.OpenSomething
Exit Sub
errhandle:
    MsgBox "拆分停止 -- " & err.Description
    err.Clear
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub
