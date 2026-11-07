Attribute VB_Name = "C系统导数转表"
Option Explicit

Private Type 警戒值                  '定义一个新数据类型记录
    序号 As Integer
    变幅 As Double
    备注文字 As String
    备注颜色 As Integer
    是否整行填充 As Integer
End Type

'
'   orgs    -   CMyReloadOrg
'       CMyReloadOrg    - dExportData
'                               -频度批次   -   指标代码 - 币种 - 数据属性
'                                                                       -   （1）当期 (2)上期
'
Public dOrgs As Object '机构类字典 机构类代码 + 地区代码   值为    CMyReloadOrg类
Private dModTables As Object    '模板文件 频度批次
Private dTtoP As Object '报表 - 频度 批次
Private dOrgRef As Object

Private modname As String
Private dataNames() As String
Private IsHideRow As Boolean
Private IsDeleteWs As Boolean
Private IsDeleteFm As Boolean
Private savePath As String
Private arrAlert() As 警戒值    '警戒值数组
Private strUnit As String   '源数据文件
Private 分析文件数据单位 As String
Private CurDate As Date

Private UniqueDate As String
Private SrcUnit As String
Private Const start_col As Integer = 4

Public Sub 系统导数转表控制台(sp As String)
'    modName = "C:\Users\Administrator\Desktop\分行工具083122\工具\金融表单180512.xls"
'    dataNames = Split("C:\Users\Administrator\Desktop\分行工具083122\data\数据对比结果2018年7月22.xls", ",")
'    savePath = "C:\Users\Administrator\Desktop\分行工具083122\data\输出"
'    IsHideRow = True
'    IsDeleteWs = False
    modname = 系统导数转EXCLE窗体.tb1.Value
    dataNames = Split(系统导数转EXCLE窗体.tb2.Value, ",")
    savePath = sp
    IsHideRow = 系统导数转EXCLE窗体.chkHideRow.Value
    IsDeleteWs = 系统导数转EXCLE窗体.chkDelEmptyWs.Value
    IsDeleteFm = 系统导数转EXCLE窗体.chkDeleteFm.Value
    '分析文件数据单位 = gloCft.GetConfig("分析文件数据单位")
    分析文件数据单位 = 系统导数转EXCLE窗体.说明文件单位
    UniqueDate = ""
    SrcUnit = ""
    'If Application.Version <= 11 Then   '如果EXCEL版本小于11，则应该要去格式化。
    If IsDeleteWs = True Then
        IsDeleteFm = True
    End If
    On Error GoTo errHandle
    Application.Calculation = xlCalculationManual
    
    Call 警戒区间校验
    Call InitOrgRef
    Call InitModTable
    Call InitData
    Call 重载选择窗体.Show
    
    Call gloCft.CloseSomething
    
    If dOrgs.count > 0 Then
        Call gloCft.CloseSomething
        Call gloDe.CreateDebugTitle("模块", "工作簿", "批次", "表格", "错误类型", "描述")
        gloT = Timer
        Call Reload
        Call gloDe.InsertNewDebug("执行耗时:" & CStr(Timer - gloT) & "s")
        Call gloDe.SetDeFoucs
    End If
    
    Call EraseSomething
    Application.Calculation = xlCalculationAutomatic
Exit Sub

errHandle:
    MsgBox err.Description
    err.Clear
    Call EraseSomething
End Sub

'
'   初始化机构信息表
'
Private Sub InitOrgRef()
    Dim arr() As Variant, i%, ws As Worksheet, d As Object
    
    Set dOrgRef = CreateObject("scripting.dictionary")
    
    Set ws = gloEt.FindWs(ThisWorkbook, "机构参照表", 0)
    If ws Is Nothing Then Exit Sub
    
    With ws
        arr = .Range("A1:L" & CStr(.Range("A65536").End(xlUp).row))
    End With
    For i = 2 To UBound(arr) Step 1
        If dOrgRef.exists(replace(CStr(arr(i, 1)), "'", "")) Then
            Set d = dOrgRef(replace(CStr(arr(i, 1)), "'", ""))
        Else
            Set d = CreateObject("scripting.dictionary")
        End If
        d(replace(CStr(arr(i, 3)), "'", "")) = CStr(arr(i, 5))
        Set dOrgRef(replace(CStr(arr(i, 1)), "'", "")) = d
        Set d = Nothing
    Next i
End Sub

'
'   清除内存空间消耗
'
Private Sub EraseSomething()
    Erase arrAlert
    Set dOrgs = Nothing
    Set dModTables = Nothing
    Set dTtoP = Nothing
    Set dOrgRef = Nothing
    Call gloEt.GetWb(modname).Close(SaveChanges:=False)
End Sub

'
'   初始化模板文件,获取频度批次及对应的表
'
Private Sub InitModTable()
    Dim wbMod As Workbook, strErr As String, arrTables As Variant
    Dim arrTtoP As Variant
    Dim i As Integer, fre As String
    
    strErr = ""
    Set wbMod = gloEt.GetWb(modname)
    
    With wbMod.Worksheets("报表清单")
        arrTables = .Range("A1:D" & .Range("A1").End(xlDown).row)
    End With
    
    
    
    Set dTtoP = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arrTables) Step 1
        If InStr(arrTables(i, 1), "A") > 0 Then dTtoP(arrTables(i, 1)) = arrTables(i, 3) & "报" & arrTables(i, 4) & "批"
    Next i
    
    Set dModTables = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arrTables) Step 1
        If InStr(arrTables(i, 1), "A") > 0 Then
            '判断是否存在此表
            If gloEt.FindWs(wbMod, CStr(arrTables(i, 1)), 0) Is Nothing Then
                strErr = strErr + "未找到 " & arrTables(i, 1)
            Else
                fre = arrTables(i, 3) & arrTables(i, 4)
                If Not dModTables.exists(fre) Then
                    Set dModTables(fre) = New Collection
                End If
                dModTables(fre).add (arrTables(i, 1))
            End If
        End If
    Next i
    
    If strErr <> "" Then
        err.Raise Number:=1002, Description:="初始化模板错误:" & strErr
    End If
End Sub

'
'   依次读取文件进 dOrgs
'
Private Sub InitData()
    Dim dn As Variant, ws As Worksheet, wb As Workbook, arrData As Variant
    Dim i As Long, clsOrg As CMyReloadOrg
    Dim orgRcode As String, frePcode As String, dataKey As String
    
    Set dOrgs = CreateObject("scripting.dictionary")
    For Each dn In dataNames
        Set wb = gloEt.GetWb(CStr(dn))
        Set ws = gloEt.FindWs(wb, "对比结果", 1)
        Call gloEt.ShowData(ws)
        Call CheckSomethingUnieque(ws)
        arrData = ws.Range("A1").CurrentRegion
        CurDate = ChangeToCDate(ws.Cells(2, 2).Value)
        For i = 2 To UBound(arrData) Step 1
            orgRcode = arrData(i, 3) & arrData(i, 5)
            orgRcode = replace(orgRcode, "'", "")
            '如果不存在，则添加这个机构
            If Not dOrgs.exists(orgRcode) Then
                Set clsOrg = New CMyReloadOrg
                clsOrg.clsOrgCode = replace(arrData(i, 3), "'", "")
                clsOrg.clsOrgName = arrData(i, 4)
                clsOrg.clsAreaCode = replace(arrData(i, 5), "'", "")
                clsOrg.clsAreaName = arrData(i, 6)
                clsOrg.clsDate = arrData(i, 2)
                Set dOrgs(orgRcode) = clsOrg
            End If
            
            frePcode = arrData(i, 12) & arrData(i, 13)  '频度 & 批次
            dataKey = arrData(i, 1) & arrData(i, 8) & arrData(i, 10) & arrData(i, 11)   '业务类 & 指标代码 & 数据属性 & 币种
            dataKey = replace(dataKey, "'", "")
            
            Set clsOrg = dOrgs(orgRcode)
            With clsOrg
                '如果不存在，则添加这个频度批次字典
                If Not .clsDicExportData.exists(frePcode) Then
                    Set .clsDicExportData(frePcode) = CreateObject("scripting.dictionary")
                End If
                '添加数据值，0-当期值，1-上期值
                If Not .clsDicExportData(frePcode).exists(dataKey) Then .clsDicExportData(frePcode).add dataKey, Array(arrData(i, 14), arrData(i, 15))
            End With
        Next i
        Set ws = Nothing
        wb.Close SaveChanges:=False
        Set wb = Nothing
    Next
End Sub

'
'   检查源数据日期的一致性
'
Private Sub CheckSomethingUnieque(ws As Worksheet)
    Dim row As Long
    With ws
        row = .Range("A1").End(xlDown).row
        
        If SrcUnit = "" Then
            SrcUnit = replace(replace(.Cells(1, 16).Value, "增减额(", ""), ")", "")
        Else
            If SrcUnit <> replace(replace(.Cells(1, 16).Value, "增减额(", ""), ")", "") Then
                ws.Parent.Close SaveChanges:=False
                err.Raise Number:=1001, Description:="CheckDateUnieque:" & ws.name & "比较文件单位不一致。"
            End If
        End If
        
        If UniqueDate = "" Then
            UniqueDate = Year(.Cells(2, 2).Value) & "." & Month(.Cells(2, 2).Value)
        Else
            If UniqueDate <> (Year(.Cells(2, 2).Value) & "." & Month(.Cells(2, 2).Value)) Then
                ws.Parent.Close SaveChanges:=False
                err.Raise Number:=1001, Description:="CheckDateUnieque:" & ws.name & "文件数据日期不一致。"
            End If
        End If
        
        If Application.WorksheetFunction.CountIf(.Range(.Cells(2, 2), .Cells(row, 2)), .Cells(2, 2).Value2) <> row - 1 Then
            ws.Parent.Close SaveChanges:=False
            err.Raise Number:=1001, Description:="CheckDateUnieque:" & ws.name & "数据日期不一致。"
        End If
    End With
End Sub

'
'   遍历 dOrgs
'
Private Sub Reload()
    Dim i As Integer, j As Integer, k As Integer, wb As Workbook, ws As Worksheet
    Dim dModKey As Variant, colT As Variant, Cmt As Comment
    Dim dOrgItem As Variant, dFreKey As Variant
    Dim dOrgDataKey As Variant
    '遍历每个机构
    For Each dOrgItem In dOrgs.items
        '打开模板文件
        Set wb = gloEt.GetWb(modname)
        wb.Worksheets("报表清单").Cells.Hyperlinks.Delete
        '遍历每个机构的数据频度批次
        For Each dModKey In dModTables.keys
            Set colT = dModTables(dModKey)
            If dOrgItem.clsDicExportData.exists(dModKey) Then
                '存在，则获取dOrgItem相应数据，读取数据
                If Not colT Is Nothing Then Call ReloadData(wb, colT, dOrgItem.clsDicExportData(dModKey))
                'Call HideRowCol(wb, colT, 1)
            Else
                '要是不存在，根据选项，将相应的工作表全部隐藏或删除
                'Call HideRowCol(wb, colT)
            End If
        Next    '下一个频度
        
        '先重计算
        For Each ws In wb.Worksheets
            If InStr(ws.name, "校验") > 0 Then
                '天幻一些数值
                For Each Cmt In ws.Comments
                    'Set rn = Cmt.Parent
                    If Cmt.text = "数值单位" Then
                        Cmt.Parent.Value = CStr(分析文件数据单位)
                    ElseIf Cmt.text = "单位转元乘数" Then
                        Cmt.Parent.Value = CLng(GetUnitVal(分析文件数据单位))
                    End If
                Next
                
                ws.Calculate
                If IsDeleteFm Then
                    ws.Range("A1:L" & ws.Range("A65536").End(xlUp).row).Value2 = ws.Range("A1:L" & ws.Range("A65536").End(xlUp).row).Value2
                End If
            End If
        Next
        '处理表格样式
        Call BeautfSheet(wb)
        '处理表
        Call HideRowCol(wb)
        '输出校验
        Call DebugRule(wb, dOrgItem)
        Call AddMacroCode(wb, dOrgItem)
        Call EndOrgReload(wb, dOrgItem)
    Next  '下一个机构
End Sub

Private Sub BeautfSheet(wb As Workbook)
    On Error Resume Next
    Dim tName As Variant, ws As Worksheet, rows As Integer, cols As Integer, iTotal As Long, i As Integer

    '遍历每一张表
    For Each tName In dTtoP.keys
        '处理表头
        Set ws = wb.Worksheets(CStr(tName))
        With ws
            cols = .Range("A1").End(xlToRight).Column
            rows = .Range("A1").End(xlDown).row
            ws.Range("A1").CurrentRegion.Offset(1, 0).EntireRow.RowHeight = 15
            ws.Range("A1").EntireRow.WrapText = True
            'ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
            
            If cols > 3 Then
                .Columns(1).ColumnWidth = 8
                .Columns(2).ColumnWidth = 30
                '.columns(3).ColumnWidth = 5
                '.columns("4:" & CStr(cols)).EntireColumn.AutoFit
            End If
        End With
    Next
End Sub

'
'   结束保存工作簿
'
Private Sub EndOrgReload(wb As Workbook, org As Variant)
    '遍历一遍设置首单元格
    On Error Resume Next
    Dim ws As Worksheet, arrUnitRef() As Variant
    
    wb.Activate
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ws.Range("A1").Activate
        End If
    Next
    
    Set ws = ThisWorkbook.Worksheets("单位不需要转换指标参照表")
    With ws
        arrUnitRef = .Range("A1:C" & CStr(.Range("B1").End(xlDown).row))
    End With
    
    With wb.Worksheets.add
        .name = ws.name
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)).NumberFormatLocal = "@"
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)) = arrUnitRef
        .Visible = xlSheetHidden
    End With
    
    Set ws = ThisWorkbook.Worksheets("指标顺序参照")
    With ws
        arrUnitRef = .Range("A1:J" & CStr(.Range("B1").End(xlDown).row))
    End With
    
    With wb.Worksheets.add
        .name = ws.name
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)).NumberFormatLocal = "@"
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)) = arrUnitRef
        .Visible = xlSheetHidden
    End With
    
    Set ws = ThisWorkbook.Worksheets("指标参照")
    With ws
        arrUnitRef = .Range("A1:C" & CStr(.Range("B1").End(xlDown).row))
    End With
    
    With wb.Worksheets.add
        .name = ws.name
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)).NumberFormatLocal = "@"
        .Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)) = arrUnitRef
        .Visible = xlSheetHidden
    End With
    
    With wb.Worksheets.add
        .name = "宏代码"
        .Cells(1, 1).Value = gloCft.GetConfig("说明文件菜单模块宏代码")
        .Cells(2, 1).Value = gloCft.GetConfig("输出说明单元格处理宏")
        .Cells(3, 1).Value = gloCft.GetConfig("输出说明word处理宏")
        .Visible = xlSheetHidden
    End With
    
    With wb.Worksheets("报表清单")
        .Activate
        .Range("A1").Activate
    End With
    '保存
    wb.SaveAs savePath & "\" & Year(org.clsDate) & "." & Month(org.clsDate) & " " & org.clsOrgName & " " & org.clsAreaName & ".xls"
    Call gloDe.AddDeHypeLinks(2, wb.fullName, "校验")
    wb.Close SaveChanges:=True
Exit Sub

errHandle:
    '保存
    wb.SaveAs savePath & "\" & Year(org.clsDate) & "." & Month(org.clsDate) & " " & org.clsOrgName & " " & org.clsAreaName & ".xls"
    Call gloDe.AddDeHypeLinks(2, wb.fullName, "校验")
    wb.Close SaveChanges:=True
    
    Call gloDe.InsertNewDebug("设置焦点", "工作表焦点转移失败，不影响其他使用。" & err.Description)
    err.Clear
End Sub

Private Sub AddMacroCode(wbDes As Workbook, org As Variant)
    Dim MyCodes As String, fullName As String, orgOutName As String, d As Object
    'Dim org As CMyReloadOrg
    On Error GoTo errHandle
    
    '插入启动菜单
    MyCodes = gloCft.GetConfig("导出文件_菜单")
    With wbDes.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines 2, MyCodes
    End With
    
    If dOrgRef.exists(org.clsOrgCode) Then
        Set d = dOrgRef(org.clsOrgCode)
    
        If d.exists(org.clsAreaCode) Then
            orgOutName = d(org.clsAreaCode)
        ElseIf d.exists("*") Then
            If CStr(d("*")) = "" Then
                orgOutName = org.clsOrgName
            Else
                orgOutName = CStr(d("*"))
            End If
        Else
            orgOutName = org.clsOrgName & " " & org.clsAreaName
        End If
    Else
        orgOutName = org.clsOrgName & " " & org.clsAreaName
    End If
    
    If orgOutName = "" Then orgOutName = org.clsOrgName & " " & org.clsAreaName
    
    
    '插入执行命令
    MyCodes = "public Const unit as Double = " + CStr(GetUnitVal(SrcUnit) / GetUnitVal(分析文件数据单位)) + Chr(10) + Chr(13) _
            + "public Const orgName as String = " + Chr(34) + org.clsOrgName + Chr(34) + Chr(10) + Chr(13) _
            + "public Const orgCode as String = " + Chr(34) + org.clsOrgCode + Chr(34) + Chr(10) + Chr(13) _
            + "public Const areaName as String = " + Chr(34) + org.clsAreaName + Chr(34) + Chr(10) + Chr(13) _
            + "public Const areaCode as String = " + Chr(34) + org.clsAreaCode + Chr(34) + Chr(10) + Chr(13) _
            + "public Const orgOutName as String = " + Chr(34) + orgOutName + Chr(34) + Chr(10) + Chr(13) _
            + "public Const 说明文件名 as String = " + Chr(34) + CStr(Year(org.clsDate)) + "." + CStr(Month(org.clsDate)) + " 说明文件.xls" + Chr(34) + Chr(10) + Chr(13) _
            + "public Const 分析文件数据单位 as String = " + Chr(34) + CStr(分析文件数据单位) + Chr(34) + Chr(10) + Chr(13) _
            + "public Const 当期数据日期 as String = " + Chr(34) + CStr(CurDate) + Chr(34) + Chr(10) + Chr(13)
    MyCodes = MyCodes + gloCft.GetConfig("输出说明通用宏")
    With wbDes.VBProject.VBComponents.add(1)
        .name = "common"
        .CodeModule.InsertLines 2, MyCodes
    End With
    
    MyCodes = gloCft.GetConfig("输出说明单元格处理宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "DealCell"
        .CodeModule.InsertLines 2, MyCodes
    End With

    MyCodes = gloCft.GetConfig("输出说明工作宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work"
        .CodeModule.InsertLines 2, MyCodes
    End With
    
    
    MyCodes = gloCft.GetConfig("导出文件_说明处理")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work_explain"
        .CodeModule.InsertLines 2, MyCodes
    End With
Exit Sub

errHandle:
    err.Clear
    MsgBox "说明文件功能受限。需要在EXCEL菜单 工具-宏－安全性－可靠发行商　中勾选【信任对于VB项目的访问】"
End Sub

'
'   输出校验结果
'
Private Sub DebugRule(wb As Workbook, org As Variant)
    'Dim org As CMyReloadOrg

    Dim ws As Worksheet, tName As String, fre As String, str As String
    Dim iTotal As Integer, iSoftTotal As Integer, iRigidTotal As Integer, Cmt As Comment, rn As Range
    
    
    For Each ws In wb.Worksheets
        If InStr(ws.name, "校验") > 0 Then
            'ws.Calculate
            With ws
                For Each Cmt In .Comments
                    Set rn = Cmt.Parent
                    If rn.EntireRow.Hidden = False And InStr(Cmt.text, "校验") > 0 Then
                        tName = .Range(.Cells(rn.row, 1).Address).CurrentRegion(1).Value
                        iTotal = iTotal + 1
                        If dTtoP.exists(tName) Then
                            fre = CStr(dTtoP(tName))
                        Else
                            fre = "*"
                        End If
                        
                        '正确得，则为空或者T，其他视作错误。
                        If rn.text <> "" And rn.text <> "T" Then
                            If InStr(rn.text, "软性") > 0 Or InStr(Cmt.text, "软性") > 0 Then '不是软性则是硬性
                                str = replace(rn.text, "软性。", "") & " -- " & Cmt.text
                                rn.Interior.colorindex = 22
                                Call gloDe.InsertNewDebug("校验", org.clsOrgName & " " & org.clsAreaName, fre, tName, "软性", str)
                                iSoftTotal = iSoftTotal + 1
                            Else
                                str = replace(rn.text, "硬性。", "") & " -- " & Cmt.text
                                rn.Interior.colorindex = 3
                                Call gloDe.InsertNewDebug("校验", org.clsOrgName & " " & org.clsAreaName, fre, tName, "硬性", str)
                                Call gloDe.AddColorDebug(3, 6)
                                iRigidTotal = iRigidTotal + 1
                            End If
                        End If
                    End If
                Next 'Cmt
            End With
        ElseIf InStr(ws.name, "报表清单") > 0 Then
            Exit For
        Else

        End If
    Next
    
    '校验完毕总结
    str = "校验" & CStr(iTotal) & "条，正确" & CStr(iTotal - iSoftTotal - iRigidTotal) & "，错误" & CStr(iSoftTotal + iRigidTotal) & "条。"
    If iSoftTotal > 0 Or iRigidTotal > 0 Then str = str & "其中，"
    If iSoftTotal > 0 Then str = str & "软性错误" & CStr(iSoftTotal) & "条。"
    If iRigidTotal > 0 Then str = str & "硬性错误" & CStr(iRigidTotal) & "条。"
    
    Call gloDe.InsertNewDebug("校验", Year(org.clsDate) & "." & Month(org.clsDate) & " " & org.clsOrgName & " " & org.clsAreaName, "", "", "总结", str)
    Call gloDe.AddColorDebug(15)

End Sub

'
'   arrT 对应频度批次表
'   dEpData 对应频度批次数据
'
Private Sub ReloadData(wb As Workbook, colT As Variant, dEpData As Variant)
    Dim tName As Variant, ws As Worksheet
    Dim tBtype As String, dataProType As String, coinType As String, dataType As String
    Dim arrData As Variant, i As Integer, j As Integer, str As String, colStep As Integer, col As Integer
    Dim dColorRn As Object, colorId As Integer, colorKey As Variant
    'On Error GoTo errHandle
    '遍历每一张表
    For Each tName In colT
        '处理表头
        Set ws = wb.Worksheets(CStr(tName))
        Set dColorRn = CreateObject("scripting.dictionary")
        dColorRn.RemoveAll
        With ws
'            If CStr(tName) = "A3302" Then
'                Debug.Print "A3302"
'            End If
            tBtype = JudgeTheBusinessType(CStr(tName))
            '获取表数据
            Call DealTableTitle(ws)
            arrData = .Range("A1").CurrentRegion

            col = UBound(arrData, 2)
            colStep = (col - 3) / 4
            
            For j = 4 To colStep + 3 Step 1
                dataProType = JudgeTheDataProperty(CStr(arrData(1, j)))
                coinType = JudgeTheDataCurrency(tBtype, CStr(arrData(1, j)))
                
                For i = 2 To UBound(arrData) Step 1
                    '数值公式
                    str = tBtype & arrData(i, 1) & dataProType & coinType
                    
                    If dEpData.exists(str) Then '如果源数据中，存在这个指标
                        '本期
                        arrData(i, j) = dEpData(str)(0)
                        '上期
                        arrData(i, j + 1 * colStep) = dEpData(str)(1)
                        '比上期
                        arrData(i, j + 2 * colStep) = ycV(arrData(i, j)) - ycV(arrData(i, j + 1 * colStep))
                        If val(arrData(i, j + 2 * colStep)) = 0 Then arrData(i, j + 2 * colStep) = ""
                        '增幅
                        If val(arrData(i, j + 2 * colStep)) <> 0 Then
                            If ycV(arrData(i, j + 1 * colStep)) <> 0 Then     '如果上期值不为空
                                arrData(i, j + 3 * colStep) = val(arrData(i, j + 2 * colStep)) / ycV(arrData(i, j + 1 * colStep))
                                If val(arrData(i, j + 3 * colStep)) = 0 Then arrData(i, j + 3 * colStep) = ""
                            End If
                        End If
                        
                        colorId = 0
                        If CStr(arrData(i, j)) = "" Then
                            '本期无，上期有
                            colorId = 6
                        ElseIf CStr(arrData(i, j + 1 * colStep)) = "" Then
                            '本期有，上期无
                            colorId = 3
                        ElseIf CStr(arrData(i, j + 3 * colStep)) <> "" Then
                            '判断加入哪个色块
                            colorId = AddMarkColor(CDbl(arrData(i, j + 3 * colStep)))
                        End If
                        If colorId <> 0 Then
                            If Not dColorRn.exists(colorId) Then
                                Set dColorRn(colorId) = Application.Union(.Cells(i, j), .Cells(i, j + 1 * colStep) _
                                                    , .Cells(i, j + 2 * colStep), .Cells(i, j + 3 * colStep))
                            Else
                                 Set dColorRn(colorId) = Application.Union(dColorRn(colorId), .Cells(i, j), .Cells(i, j + 1 * colStep) _
                                                    , .Cells(i, j + 2 * colStep), .Cells(i, j + 3 * colStep))
                            End If
                        End If
                    End If
                Next i
            Next j
            
            '遍历上色
            For Each colorKey In dColorRn.keys
                dColorRn(colorKey).Interior.colorindex = CInt(colorKey)
            Next
            
            .Range("A1").CurrentRegion = arrData
            Erase arrData
            Set dColorRn = Nothing
        End With
    Next
'Exit Sub
'
'errHandle:
'    Call gloDe.InsertNewDebug("系统导数模块", wb.name, "-", "导数错误", "出现点意外 -- " & CStr(str))
'    Resume
End Sub

Private Sub AddHypeLink(ws As Worksheet, wb As Workbook)
    On Error Resume Next
    Dim ws_c As Worksheet, ws_l As Worksheet
    Set ws_c = gloEt.FindWs(wb, "校验", 0)
    Set ws_l = wb.Worksheets("报表清单")
    'If ws_c Is Nothing Then Exit Sub
    
    
    Dim key_rn1 As Range, key_rn2 As Range
    Set key_rn1 = ws_c.UsedRange.SpecialCells(xlCellTypeVisible).find(ws.name, LookIn:=xlValues, LookAt:=xlWhole)
    Set key_rn2 = ws_l.UsedRange.SpecialCells(xlCellTypeVisible).find(ws.name, LookIn:=xlValues, LookAt:=xlWhole)
    If Not key_rn2 Is Nothing Then
        ws_l.Hyperlinks.add Anchor:=key_rn2.Offset(0, 1), Address:="", SubAddress:=ws.name + "!" + ws.Range("A1").Address, ScreenTip:="跳转到报表"
        ws.Hyperlinks.add Anchor:=ws.Range("A1"), Address:="", SubAddress:=ws_l.name + "!" + key_rn2.Offset(0, 1).Address, ScreenTip:="返回清单"
        If Not key_rn1 Is Nothing Then
            ws_l.Hyperlinks.add Anchor:=key_rn2, Address:="", SubAddress:=ws_c.name + "!" + key_rn1.Address, ScreenTip:="跳转到校验"
            ws_c.Hyperlinks.add Anchor:=key_rn1, Address:="", SubAddress:=ws_l.name + "!" + key_rn2.Address, ScreenTip:="返回清单"
        End If
    End If
End Sub

'
'   隐藏行
'
Private Sub HideRowCol(wb As Workbook)
    Dim tName As Variant, ws As Worksheet, rows As Integer, cols As Integer, iTotal As Long, i As Integer
    Dim empty_rng As Range, t_em_rn As Range

    '遍历每一张表
    For Each tName In dTtoP.keys
        '处理表头
        Set ws = wb.Worksheets(CStr(tName))
        With ws
            cols = .Range("A1").End(xlToRight).Column
            rows = .Range("A1").End(xlDown).row
            iTotal = gloEt.GetNotEmptyCellCount(.Range(.Cells(2, 4), .Cells(rows, cols)))
            If iTotal = 0 Then
                Call hideCheckArea(wb, CStr(tName))
                If IsDeleteWs Then
                    ws.Delete
                Else
                    ws.Visible = xlSheetHidden
                End If
            Else
                '处理空列
                Set empty_rng = Nothing
                For i = 4 To cols Step 1
                    '判断是否空列
                    iTotal = gloEt.GetNotEmptyCellCount(.Range(.Cells(2, i), .Cells(rows, i)))
                    If iTotal = 0 Then
                        If empty_rng Is Nothing Then
                            Set empty_rng = .Cells(2, i)
                        Else
                            Set empty_rng = Application.Union(empty_rng, .Cells(2, i))
                        End If
                    End If
                Next i
                If Not empty_rng Is Nothing Then empty_rng.EntireColumn.ColumnWidth = 5  '没有隐藏
                
                '处理空行
                If IsHideRow Then
                    Set empty_rng = Nothing
                    For i = 2 To rows Step 1
                        Set t_em_rn = .Range(.Cells(i, 4), .Cells(i, cols))
                        iTotal = gloEt.GetNotEmptyCellCount(t_em_rn)
                        If iTotal = 0 Then
                            If empty_rng Is Nothing Then
                                Set empty_rng = t_em_rn
                            Else
                                Set empty_rng = Application.Union(empty_rng, t_em_rn)
                            End If
                        End If
                    Next i
                
                    If Not empty_rng Is Nothing Then empty_rng.EntireRow.Hidden = True
                End If
                
                '添加超链接
                Call AddHypeLink(ws, wb)
            End If
        End With
    Next
End Sub

Private Sub hideCheckArea(wb As Workbook, key As String)
    Dim ws As Worksheet, key_rn As Range
    
    For Each ws In wb.Worksheets
        If InStr(ws.name, "校验") > 0 Then
            With ws
                Set key_rn = .UsedRange.find(key, LookIn:=xlValues, LookAt:=xlWhole)
            
                If Not key_rn Is Nothing Then
                    '不对，现在无论如何都要去格式了，因为没有使用iferror,如果要删除，则提前去格式化
                    'If IsDeleteWs Then .Range(CStr(key_rn.row) & ":" & CStr(key_rn.row + key_rn.CurrentRegion.rows.count - 1)).Value2 = .Range(CStr(key_rn.row) & ":" & CStr(key_rn.row + key_rn.CurrentRegion.rows.count - 1)).Value2
                    .Range(CStr(key_rn.row) & ":" & CStr(key_rn.row + key_rn.CurrentRegion.rows.count - 1)).EntireRow.Hidden = True
                End If
            End With
        ElseIf InStr(ws.name, "报表清单") > 0 Then
            Exit For
        Else
        
        End If
    Next
End Sub

'
'   获取标记的颜色
'
Private Function AddMarkColor(h As Double) As Integer
    Dim i%, c%, colorId As Integer
    c = UBound(arrAlert)
    
    colorId = 0
    For i = c To LBound(arrAlert) Step -1
        If Round(h - arrAlert(i).变幅, 10) >= 0 Then
            colorId = arrAlert(i).备注颜色
            Exit For
        End If
    Next i
    AddMarkColor = colorId
End Function


'
'   处理表格表头
'
Private Sub DealTableTitle(ws As Worksheet)
     '删除一些无用列

    Dim cols As Integer, newCols As Integer, colCnt As Integer, rows As Integer, i As Integer, val As String
    cols = ws.Range("A1").End(xlToRight).Column
    rows = ws.Range("A1").End(xlDown).row
    
    If cols <= start_col - 1 Then
        Exit Sub
    End If
    
    With ws
        For i = 5 To cols Step 1
            val = .Cells(1, i).Value
            If InStr(val, "剥离") > 0 Or InStr(val, "退出") > 0 Or InStr(val, "上期") > 0 Or val = "" Then
                Exit For
            End If
        Next i
        If i <= cols Then .Range(.Cells(1, i), .Cells(1, cols)).Clear
        cols = .Range("A1").End(xlToRight).Column
        colCnt = cols - start_col + 1
        .Activate
        '数据值格式
        '.Range(.Cells(2, START_COL), .Cells(rows, cols)).NumberFormatLocal = "0.00"
        
        '上期值列标题
        .Range(.Cells(1, cols - colCnt + 1), .Cells(1, cols)).Copy .Cells(1, cols + 1)
        
        '将新增的列,添加上相应前缀啊
        newCols = .Range("A1").End(xlToRight).Column
        For i = cols + 1 To newCols Step 1
            .Cells(1, i).Value = "上期-" + .Cells(1, i).Value
            '.Range(.Cells(2, i), .Cells(rows, i)).NumberFormatLocal = "0.00"
        Next i
        
        
        '比上期列标题
        cols = newCols
        .Range(.Cells(1, cols - 2 * colCnt + 1), .Cells(1, cols - colCnt)).Copy .Cells(1, cols + 1)
        
        '将新增的列,添加上相应前缀啊
        newCols = .Range("A1").End(xlToRight).Column
        For i = cols + 1 To newCols Step 1
            .Cells(1, i).Value = "增减额-" + .Cells(1, i).Value
            '.Range(.Cells(2, i), .Cells(rows, i)).NumberFormatLocal = "0.00"
        Next i
        
         '比上期百分比列标题
        cols = newCols
        .Range(.Cells(1, cols - 3 * colCnt + 1), .Cells(1, cols - 2 * colCnt)).Copy .Cells(1, cols + 1)
        
        '将新增的列,添加上相应前缀啊
        newCols = .Range("A1").End(xlToRight).Column
        For i = cols + 1 To newCols Step 1
            .Cells(1, i).Value = "环比-" + .Cells(1, i).Value
            '.Range(.Cells(2, i), .Cells(rows, i)).NumberFormatLocal = "0.00%"
        Next i
        .Range(.Cells(2, start_col), .Cells(rows, newCols - (newCols - start_col) / 4)).NumberFormatLocal = "0.00"
        .Range(.Cells(2, newCols + 1 - (newCols - start_col) / 4), .Cells(rows, newCols)).NumberFormatLocal = "0.00%"
    End With
End Sub

'
'   获取表格业务类
'
Private Function JudgeTheBusinessType(wsName As String) As String
    Dim businessType As String
    
    If InStr(wsName, "A1") > 0 Then
        businessType = "人民币"
    ElseIf InStr(wsName, "A2") > 0 Then
        businessType = "外币"
    Else
        businessType = "本外币"
    End If
    
    JudgeTheBusinessType = businessType
End Function

'
'   通过表单名，行标题 判断币种
'   人民币、外币、本外币
'
Private Function JudgeTheDataCurrency(businessType As String, colName As String) As String
    Dim dataCurrency As String
    
    If businessType = "人民币" Then
        dataCurrency = "人民币"
    ElseIf businessType = "外币" Then
        dataCurrency = "美元合计"
    Else
        '本外币比较复杂，判断是否包含
        If InStr(colName, "人民币") > 0 Then
            dataCurrency = "人民币"
        ElseIf InStr(colName, "本外币") > 0 Then
            dataCurrency = "本外币"
        ElseIf InStr(colName, "外币") > 0 Then
            dataCurrency = "美元合计"
        Else
            '可能部分出现余额的情况
            dataCurrency = "人民币"
        End If
    End If
    
    JudgeTheDataCurrency = dataCurrency
End Function

'
'   通过 列标题 判断数据属性
'   返回 余额 发生额
'
Private Function JudgeTheDataProperty(colName As String) As String
    Dim dataProperty As String
    
    If InStr(colName, "发生额") > 0 Then
        dataProperty = "发生额"
    ElseIf InStr(colName, "余额") > 0 Then
        dataProperty = "余额"
    Else
        dataProperty = "余额"
    End If
    
    JudgeTheDataProperty = dataProperty
End Function

'
'   警戒区间校验
'
Private Sub 警戒区间校验()
    Dim i As Integer, wsAlert As Worksheet, row As Integer
    Dim 警戒值行数1, 警戒值行数2 As Integer
    Dim 警戒值是否正确
    
    Set wsAlert = ThisWorkbook.Worksheets("警戒区间")
    

    '警戒值设定逻辑检查
    With wsAlert
        If .Range("A1").Value <> "默认警戒区间" Then
            err.Raise Number:=1001, Description:="警戒区间区域被破坏，请正确设置"
        ElseIf .Range("B1").End(xlDown).row <> .Range("A1").End(xlDown).row Then
            err.Raise Number:=1001, Description:="警戒区间数值上标、下标设置不全！"
        Else
            row = .Range("A1").End(xlDown).row
            ReDim arrAlert(row - 2)
            '检测变幅是否从小到大
            For i = 2 To row Step 1
                If i = 2 Then
                    arrAlert(i - 2).序号 = .Cells(i, 1).Value2
                    arrAlert(i - 2).变幅 = .Cells(i, 2).Value2
                    arrAlert(i - 2).备注文字 = .Cells(i, 3).Value2
                    arrAlert(i - 2).备注颜色 = .Cells(i, 3).Interior.colorindex
                    arrAlert(i - 2).是否整行填充 = CInt(val(.Cells(i, 4).Value2))
                Else
                    If .Cells(i, 2).Value2 > .Cells(i, 3).Value2 Then
                        err.Raise Number:=1001, Description:="警戒区间阀值顺序不对，应该从小到大"
                    Else
                        arrAlert(i - 2).序号 = .Cells(i, 1).Value2
                        arrAlert(i - 2).变幅 = .Cells(i, 2).Value2
                        arrAlert(i - 2).备注文字 = .Cells(i, 3).Value2
                        arrAlert(i - 2).备注颜色 = .Cells(i, 3).Interior.colorindex
                        arrAlert(i - 2).是否整行填充 = CInt(val(.Cells(i, 4).Value2))
                    End If
                End If
            Next i
        End If
    End With
End Sub
