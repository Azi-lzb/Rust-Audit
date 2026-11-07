VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMySysMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private glodHzMod As Object   '存储汇总模板
Private glodExtendMod As Object '存储外部文件
Private m_OldTitle As String    '用于重绘
Private m_dTitle As Object
Private m_wsSys As Worksheet

'类初始化
Private Sub Class_Initialize()
    '不存在则增加
    Set glodHzMod = Nothing
    Set glodExtendMod = Nothing
    Set m_dTitle = Nothing
End Sub

'类销毁
Private Sub Class_Terminate()
    
End Sub

Public Function GetMyHzMod() As Object
    Set GetMyHzMod = glodHzMod
End Function

Public Function GetMyExtendMod() As Object
    Set GetMyExtendMod = glodExtendMod
End Function

Public Function GetMyHzModWb() As Workbook
    Set GetMyHzModWb = gloEt.GetWb(m_wsSys.Cells(1, 2).value)
End Function


    

Public Sub RecoverSysMenu(srcfiles As Variant)
    m_wsSys.Range("A1").Resize(UBound(srcfiles), UBound(srcfiles, 2)) = srcfiles
    m_wsSys.Cells.WrapText = False
    m_wsSys.Cells.Font.Size = 12
    m_wsSys.Cells.Font.name = "微软雅黑"
    m_wsSys.Cells.NumberFormat = "@"
End Sub

Public Sub SetSysWorkSheet(name As String)
    Set m_wsSys = gloEt.FindWs(ThisWorkbook, name, 0, True)
End Sub

Public Sub SetSysWorkSheetFormat()
    m_wsSys.Cells.WrapText = False
End Sub


'
'   获取菜单面板源数据可用区域。
'
Public Function GetSysMenuAreaValue() As Variant
    Dim arr As Variant, row As Long, col As Integer
    
    row = gloEt.GetCurRows(m_wsSys, "C")
    col = m_wsSys.Cells(5, 256).End(xlToLeft).Column
    
    If col < 9 Then col = 9
    If row < 5 Then row = SYSSTARTROW
    
    arr = m_wsSys.Range(m_wsSys.Cells(1, 1), m_wsSys.Cells(row, col)).value
    GetSysMenuAreaValue = arr
End Function

Public Function GetMenuTitleDic() As Object
    Dim dset As Object, srcfiles As Variant, i As Integer
    srcfiles = GetSysMenuAreaValue
    Set dset = CreateObject("scripting.dictionary")
    For i = 1 To UBound(srcfiles, 2) Step 1
        dset(CStr(srcfiles(SYSSTARTROW, i))) = i
    Next i
    Set GetMenuTitleDic = dset
End Function

Public Function GetMenuTitleCol(title) As Integer
    Dim col As Integer
    col = 4
    If Not m_dTitle Is Nothing Then
        If m_dTitle.exists(title) Then
            col = CInt(m_dTitle(title))
        End If
    End If
    
    GetMenuTitleCol = col
End Function

Public Sub addhypelink(title As String, i As Long, path As String)
    On Error Resume Next
    If Not m_dTitle Is Nothing Then
        If m_dTitle.exists(title) Then
            m_wsSys.Hyperlinks.add m_wsSys.Cells(i, CInt(m_dTitle(title))), path
        End If
    End If
End Sub

'
'   给相应的标题设置相应值
'
Public Sub SetMenuTitleValue(ByRef srcfiles As Variant, title As String, i As Long, value As String)
    If Not m_dTitle Is Nothing Then
        If m_dTitle.exists(title) Then
            srcfiles(i, CInt(m_dTitle(title))) = value
        Else
            '增加一列
            Dim col As Integer
            For col = 2 To UBound(srcfiles, 2) Step 1
                If VarType(srcfiles(SYSSTARTROW, col)) = vbEmpty Or CStr(srcfiles(SYSSTARTROW, col)) = "" Then
                    Exit For
                End If
            Next col
            
            If col > UBound(srcfiles, 2) Then
                ReDim Preserve srcfiles(1 To UBound(srcfiles), 1 To UBound(srcfiles, 2) + 1)
            End If
            If col < 4 Then col = 4
            
            m_dTitle(title) = col
            srcfiles(i, CInt(m_dTitle(title))) = value
           srcfiles(SYSSTARTROW, col) = title
        End If
    End If
End Sub

'
'   给相应的位置设置批注
'
Public Sub SetMenuTitleCommentValue(title As String, i As Long, value As String)
    If Not m_dTitle Is Nothing Then
        If Not m_dTitle.exists(title) Is Nothing Then
            Call gloEt.InsertComment(m_wsSys, value, m_wsSys.Cells(i, CInt(m_dTitle(title))).address)
        End If
    End If
End Sub

'
'   清除非设置标题区域
'
Public Sub ClearMenuTitle(Optional srclear As Integer = 4)
    Dim col As Integer, row As Long
    With m_wsSys
        For col = srclear To 256 Step 1
            If VarType(.Cells(SYSSTARTROW, col).value) = vbEmpty Or CStr(.Cells(SYSSTARTROW, col).value) = "" Or InStr(CStr(.Cells(SYSSTARTROW, col).value), "设置") > 0 Then
                col = col - 1
                Exit For
            End If
        Next col
        If col > 256 Then col = 256
        row = gloEt.GetCurRows(m_wsSys, "C")
        If row < SYSSTARTROW + 1 Then row = SYSSTARTROW + 1
        
        .Range(.Cells(SYSSTARTROW + 1, srclear), .Cells(row, col)).ClearContents
    End With
End Sub

'
'   获取相应配置
'
Public Function GetConfig(key As String) As String
    Dim rn As Range, c As Range
    Set rn = m_wsSys.Range("1:5")
    With rn
        Set c = .find(key, LookIn:=xlValues, LookAt:=xlWhole)
        If c Is Nothing Then
            GetConfig = ""
        Else
            GetConfig = CStr(m_wsSys.Cells(c.row, c.Column + 1).value)
        End If
    End With
End Function

'
'   获取源数据,以字典形式返回
'
Public Function GetSysMenuSrcs() As Object
    Dim arr As Variant, row As Range, col As Integer, i As Long, dSrc As Object, curRow As Long
    Dim rnRe As Range, dHideSrc As Object
    
    Set rnRe = Nothing
    arr = GetSysMenuAreaValue
    Set dSrc = Nothing
    Set dSrc = CreateObject("scripting.dictionary")
    Set dHideSrc = CreateObject("scripting.dictionary")
    '在这里考虑去除重复项目
    With m_wsSys
        '将隐藏的，重复的删除
        For i = SYSSTARTROW To UBound(arr) Step 1
            If .Cells(i, 1).EntireColumn.Hidden = True Then
                If dHideSrc.exists(CStr(arr(i, 2))) Then
                    If rnRe Is Nothing Then
                        Set rnRe = .Cells(i, 1)
                    Else
                        Set rnRe = Application.Union(rnRe, .Cells(i, 1))
                    End If
                Else
                    dHideSrc(CStr(arr(i, 2))) = i
                End If
            End If
        Next i
        
        '将隐藏的，又重新添加的，将隐藏的重的删除
        For i = SYSSTARTROW To UBound(arr) Step 1
            If .Cells(i, 1).EntireColumn.Hidden = False Then
                '如果没有隐藏，则读取,形成 文件->行号
                If dSrc.exists(CStr(arr(i, 2))) Then
                    If rnRe Is Nothing Then
                        Set rnRe = .Cells(i, 1)
                    Else
                        Set rnRe = Application.Union(rnRe, .Cells(i, 1))
                    End If
                Else
                    If dHideSrc.exists(CStr(arr(i, 2))) Then
                        If rnRe Is Nothing Then
                            Set rnRe = .Cells(CLng(dHideSrc(CStr(arr(i, 2)))), 1)
                        Else
                            Set rnRe = Application.Union(rnRe, .Cells(CLng(dHideSrc(CStr(arr(i, 2)))), 1))
                        End If
                    End If
                    dSrc(CStr(arr(i, 2))) = i
                End If
            End If
        Next i
        
        If Not rnRe Is Nothing Then rnRe.EntireRow.Delete
        Set dHideSrc = Nothing
        '到这里清除重复项完毕，重新获取一次
        
        dSrc.RemoveAll
        arr = GetSysMenuAreaValue
        For i = SYSSTARTROW + 1 To UBound(arr) Step 1
            If .Cells(i, 1).EntireColumn.Hidden = False Then
                dSrc(CStr(arr(i, 2))) = i
            End If
        Next i
        curRow = gloEt.GetCurRows(m_wsSys, "C")
        
        '插入批注的形式
        Call gloEt.InsertComment(m_wsSys, "面板源数据：" & CStr(curRow - SYSSTARTROW) & "个。读取源数据：" & CStr(dSrc.count) & "个。", .Cells(SYSSTARTROW, 3).address)
    End With
    
    Set GetSysMenuSrcs = dSrc
End Function

Public Function SelExtendFile()
    Dim fileName As Variant
    fileName = Application.GetOpenFilename(fileFilter:="Excel 文件 (*.xls; *.xlsx; *.csv), *.xls; *.xlsx; *.csv", title:="选择外部文件", MultiSelect:=False)
    
    If VarType(fileName) = vbBoolean Then
        
    ElseIf VarType(fileName) = vbString Then
        With m_wsSys
            .Cells(2, 1).value = "外部文件"
            .Cells(2, 2).value = CStr(fileName)
            .Cells(2, 3).value = gloFt.getShortFileName2(CStr(fileName))
            m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(2, 2), address:=CStr(fileName)
        End With
        Call gloCft.CloseSomething
        Call InitModFileMessage
        Call IsAvailableModTableSetting
        Call gloCft.OpenSomething
    Else
        MsgBox "出现不兼容"
    End If
End Function

Private Function get_mod_path_from_list() As String
    ' 返回选择的单元格行号（第一行）
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim targetSheetName As String
     
    ' 1. 判断是否在 "模板文件候选" 工作表中
    targetSheetName = "模板列表"
    
    ' 获取当前活动工作表
    Set ws = ActiveSheet
    
    ' 如果不在目标工作表，跳转过去
    If ws.name <> targetSheetName Then
        Set ws = ThisWorkbook.Worksheets(targetSheetName)
        ' 激活目标工作表
        ws.Activate
        ws.Select
    End If
    
    ' 弹出输入框让用户选择
    On Error Resume Next
    Set rng = Application.InputBox("选择模板所在行！", title:="提示", Type:=8)
    
    ' 判断是否取消
    If rng Is Nothing Then
        get_mod_path_from_list = ""
        Exit Function
    End If
    
    ' 验证是否在目标工作表
    If Not rng.Parent.name = targetSheetName Then
        MsgBox "请选择【" & targetSheetName & "】工作表中的单元格！"
        get_mod_path_from_list = ""
        Exit Function
    End If
    
    ' 返回行号
    get_mod_path_from_list = ws.Cells(rng.row, 2).value
    
End Function

Public Function SelModFile_form_list()
    ' 判断是不是在 "模板文件候选" 工作表中，如果不在的话，就跳转过去
    
    ' 弹出单元格引用框，点击模板位置单元格后返回
    
    ' 判断选择内容是否为空，不为空的话，提示成功
    Dim fileName As String
    fileName = get_mod_path_from_list
    
    If VarType(fileName) = vbBoolean Then
        
    ElseIf VarType(fileName) = vbString Then
        If fileName <> "" Then
            With m_wsSys
                .Cells(1, 1).value = "模板文件"
                .Cells(1, 2).value = CStr(fileName)
                .Cells(1, 3).value = gloFt.getShortFileName2(CStr(fileName))
                m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(1, 2), address:=CStr(fileName)
                
                m_wsSys.Activate
                m_wsSys.Select
            End With
        
            Call gloCft.CloseSomething
            Call InitModFileMessage
            Call IsAvailableModTableSetting
        Else
            MsgBox "该单元格区域无模板内容"
        End If
        Call gloCft.OpenSomething
    Else
        MsgBox "出现不兼容"
    End If
End Function

'
'   选择模板,strs 为指定列,默认带有源文件，及，源文件缩写
'
Public Function SelModFile()
    
    Dim fileName As Variant
    fileName = Application.GetOpenFilename(fileFilter:="Excel 文件 (*.xls; *.xlsx; *.csv), *.xls; *.xlsx; *.csv", title:="选择汇总模板文件", MultiSelect:=False)
    
    If VarType(fileName) = vbBoolean Then
        
    ElseIf VarType(fileName) = vbString Then
        With m_wsSys
            .Cells(1, 1).value = "模板文件"
            .Cells(1, 2).value = CStr(fileName)
            .Cells(1, 3).value = gloFt.getShortFileName2(CStr(fileName))
            m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(1, 2), address:=CStr(fileName)
        End With
        Call gloCft.CloseSomething
        Call InitModFileMessage
        Call IsAvailableModTableSetting
        Call gloCft.OpenSomething
    Else
        MsgBox "出现不兼容"
    End If
End Function

Public Function IsAvailableModTableSetting() As Boolean
    If glodHzMod Is Nothing Then
        IsAvailableModTableSetting = False
        Exit Function
    End If
    
    Dim clshzT As CMyHzModTable, v As Variant, str As String, j As Integer
    v = glodHzMod.items
    str = ""
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        str = IIf(str = "", clshzT.TableSettingAvailableDescription, IIf(clshzT.TableSettingAvailableDescription = "", str, str & Chr(10) & clshzT.TableSettingAvailableDescription))
    Next j
    
    If str = "" Then
        IsAvailableModTableSetting = True
        Call gloEt.InsertComment(m_wsSys, "模板正确。", "C1")
    Else
        MsgBox "模板表格设置有误，请检查，详情看模板文件单元格批注。"
        Call gloEt.InsertComment(m_wsSys, "模板有误。" & str, "C1")
        IsAvailableModTableSetting = False
    End If
End Function

Public Function IsHzRowColumnToData() As Boolean
    If glodHzMod Is Nothing Then
        IsHzRowColumnToData = False
        Exit Function
    End If
    
    Dim clshzT As CMyHzModTable, v As Variant, r As Boolean, j As Integer
    v = glodHzMod.items

    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        If clshzT.IsRowColumnToData = False Then
            MsgBox "存在数据、文字、公式区域找不到对应行列。"
            IsHzRowColumnToData = False
            Exit Function
        End If
    Next j
    
    IsHzRowColumnToData = True
End Function


'
' 选择源数据
'
Public Function SelSrcFiles()
    
    Dim fileName As Variant

    fileName = Application.GetOpenFilename(fileFilter:="Excel 文件 (*.xls; *.xlsx; *.csv), *.xls; *.xlsx; *.csv,所有文件(*.*),*.*", title:="选择汇总源数据", MultiSelect:=True)
    If VarType(fileName) = vbBoolean Then
        
    ElseIf VarType(fileName) = vbArray Or VarType(fileName) = 8200 Or VarType(fileName) = 8204 Then
        Call gloCft.CloseSomething
        Call InitSrcFiles(fileName)
    Else
        
    End If
    Call gloCft.OpenSomething
End Function



Public Sub SelSrcFilesByFolder()
    Dim folder As String, arrfile As Variant
    Dim v As Variant, curRow As Long, i As Long
    
    folder = gloFt.GetOutputFolder("请选择源文件夹")
    If folder <> "" Then
        arrfile = gloFt.GetFileList(folder)
        curRow = gloEt.GetCurRows(m_wsSys, "C") + 1
        If curRow <= SYSSTARTROW Then curRow = SYSSTARTROW + 1
        
        With m_wsSys
            .Cells(SYSSTARTROW, 2).value = "源文件名"
            .Cells(SYSSTARTROW, 3).value = "源文件"
        End With
        For i = 0 To UBound(arrfile)
            m_wsSys.Cells(curRow, 2).value = CStr(arrfile(i))
            m_wsSys.Cells(curRow, 3).value = gloFt.getShortFileName2(CStr(arrfile(i)))
            m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(curRow, 2), address:=CStr(arrfile(i))
            curRow = curRow + 1
        Next
        
        Call GetSysMenuSrcs
    End If
End Sub

Public Sub SelSrcFilesByFolderWithFolder()
    Dim folder As String, arrfile As Variant
    Dim v As Variant, curRow As Long, i As Long
    
    folder = gloFt.GetOutputFolder("请选择源文件夹")
    If folder <> "" Then
        arrfile = gloFt.GetFileList(folder, True)
        curRow = gloEt.GetCurRows(m_wsSys, "C") + 1
        If curRow <= SYSSTARTROW Then curRow = SYSSTARTROW + 1
        
        With m_wsSys
            .Cells(SYSSTARTROW, 2).value = "源文件名"
            .Cells(SYSSTARTROW, 3).value = "源文件"
        End With
        For i = 0 To UBound(arrfile)
            m_wsSys.Cells(curRow, 2).value = CStr(arrfile(i))
            m_wsSys.Cells(curRow, 3).value = gloFt.getShortFileName2(CStr(arrfile(i)))
            m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(curRow, 2), address:=CStr(arrfile(i))
            curRow = curRow + 1
        Next
        
        Call GetSysMenuSrcs
    End If
End Sub

'
'   批量选择汇总源数据
'
Public Sub SelSrcFilesInFolder()
    
    Dim folder As String
    folder = gloFt.GetOutputFolder("请选择源文件夹")

    If folder <> "" Then
        Call gloCft.CloseSomething
        With Application.FileSearch
            .NewSearch
            .LookIn = folder
            .SearchSubFolders = True
            .fileName = "*.*"
            '.FileType = msoFileTypeExcelWorkbooks
            If .Execute > 0 Then
                Call InitSrcFiles2(.FoundFiles)
            End If
        End With
    End If
    
    Call gloCft.OpenSomething
End Sub

'
'   批量选择含文件夹
'
Public Sub SelSrcFilesInFolder2()
    
    Dim folder As String

    folder = gloFt.GetOutputFolder("请选择源文件夹")
    
    If folder <> "" Then
        Call gloCft.CloseSomething
        Dim fso As Object, objFolder As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set objFolder = fso.GetFolder(folder)
        Call GetFileFolderList(objFolder, m_wsSys)
        
        Call GetSysMenuSrcs
        Set objFolder = Nothing
        Set fso = Nothing
    End If
    
    Call gloCft.OpenSomething
End Sub

Private Sub GetFileFolderList(objFolder As Object, ws As Worksheet)
    Dim SubFolders As Object, SubFolder As Object
    Dim files As Object, file As Object, curRow As Long
    'strTmp = "文件夹:" & vbTab & objfolder.path & vbCrLf
    
    curRow = gloEt.GetCurRows(m_wsSys, "C") + 1
    If curRow <= SYSSTARTROW Then curRow = SYSSTARTROW + 1

    ws.Cells(curRow, 2).value = CStr(objFolder.path)
    ws.Cells(curRow, 3).value = CStr(objFolder.path)
    curRow = curRow + 1
    
    Set files = objFolder.files
    If files.count <> 0 Then
        For Each file In files
            ws.Cells(curRow, 2).value = CStr(file.path)
            ws.Cells(curRow, 3).value = gloFt.getShortFileName2(CStr(file.path))
            curRow = curRow + 1
        Next
    End If
    Set SubFolders = objFolder.SubFolders
    If SubFolders.count <> 0 Then
        For Each SubFolder In SubFolders
            Call GetFileFolderList(SubFolder, ws)
        Next
    End If
End Sub


'
'   将原数据转化到面板中
'
Private Sub InitSrcFiles2(fileName As Variant)
    Dim v As Variant, curRow As Long, i As Long
    
    If IsEmpty(fileName) Then Exit Sub
    
    curRow = gloEt.GetCurRows(m_wsSys, "C") + 1
    If curRow <= SYSSTARTROW Then curRow = SYSSTARTROW + 1
    
    With m_wsSys
        .Cells(SYSSTARTROW, 2).value = "源文件名"
        .Cells(SYSSTARTROW, 3).value = "源文件"

        For i = 1 To fileName.count
            m_wsSys.Cells(curRow, 2).value = CStr(fileName(i))
            m_wsSys.Cells(curRow, 3).value = gloFt.getShortFileName2(CStr(fileName(i)))
            curRow = curRow + 1
        Next i

    End With
    Call GetSysMenuSrcs
End Sub

'
'   将原数据转化到面板中
'
Private Sub InitSrcFiles(fileName As Variant)
    Dim v As Variant, curRow As Long
    
    If IsEmpty(fileName) Then Exit Sub
    
    curRow = gloEt.GetCurRows(m_wsSys, "C") + 1
    If curRow <= SYSSTARTROW Then curRow = SYSSTARTROW + 1
    
    With m_wsSys
        .Cells(SYSSTARTROW, 2).value = "源文件名"
        .Cells(SYSSTARTROW, 3).value = "源文件"
        For Each v In fileName
            m_wsSys.Cells(curRow, 2).value = CStr(v)
            m_wsSys.Cells(curRow, 3).value = gloFt.getShortFileName2(CStr(v))
            m_wsSys.Hyperlinks.add anchor:=m_wsSys.Cells(curRow, 2), address:=CStr(v)
            curRow = curRow + 1
        Next
    End With
    Call GetSysMenuSrcs
End Sub


'
'   设置标题
'
Public Function SetMenuTitle(ParamArray strs() As Variant)
    m_OldTitle = "源文件名,源文件,新文件名,标准化,校验结果,执行结果,校验结果汇总"
    Dim i As Integer, v As Variant, chmt As CMyHzModTable, v2 As Variant
    
    
    For i = 0 To UBound(strs) Step 1
        m_OldTitle = m_OldTitle + "," + strs(i)
    Next i
    
    '默认执行到这里，已经时要选过模板文件的了，遍历模板文件
    If Not glodHzMod Is Nothing Then
        v = glodHzMod.items
        For i = 0 To glodHzMod.count - 1 Step 1
            Set chmt = v(i)
            For Each v2 In chmt.dSetAreaParse.keys
                 m_OldTitle = m_OldTitle + "," + CStr(v2)
            Next
        Next i
    End If
    
    Call RebuildMenuTitle
End Function

Public Sub ReloadMenuTitleDic()
    Dim srcfiles As Variant, i As Integer
    srcfiles = GetSysMenuAreaValue
    
    Set m_dTitle = Nothing
    Set m_dTitle = CreateObject("scripting.dictionary")
    
    For i = 1 To UBound(srcfiles, 2) Step 1
        m_dTitle(CStr(srcfiles(SYSSTARTROW, i))) = i
    Next i
End Sub
    

'
'   重建菜单
'
Private Sub RebuildMenuTitle()
    Dim arr As Variant, i As Integer
    arr = Split(m_OldTitle, ",")
    
    Set m_dTitle = Nothing
    Set m_dTitle = CreateObject("scripting.dictionary")
    
    '清除第五行
    With m_wsSys
        .Range("A5").EntireRow.ClearContents
        
        '原来的
        For i = 0 To UBound(arr) Step 1
            m_dTitle(CStr(arr(i))) = i + 2
            .Cells(SYSSTARTROW, i + 2).value = CStr(arr(i))
        Next i
    End With
End Sub

'
'   初始化模板
'
Public Function InitModFileMessage() As Boolean
    On Error GoTo errhandle
    Dim modname As String
    Set glodHzMod = Nothing
    Set glodHzMod = CreateObject("scripting.dictionary")
    Set glodExtendMod = Nothing
    Set glodExtendMod = CreateObject("scripting.dictionary")
    
    modname = CStr(m_wsSys.Cells(1, 2).value)
    m_wsSys.Cells(1, 3).value = gloFt.getShortFileName2(modname)
    
    Dim rnT As Range, t As CMyHzModTable, ws As Worksheet, i As Integer, wb As Workbook
    Dim cmtText As String, arrTemp As Variant
    Set wb = gloEt.GetWb(modname)
    For i = 1 To wb.Worksheets.count Step 1
        Set ws = wb.Worksheets(i)
        
        '遍历模板表的非额外信息表格
        cmtText = gloEt.GetCommentStr(ws.Range("A1"))
        If InStr(cmtText, "外部文件") > 0 Then
            '需要得到外部文件对应的工作表
            arrTemp = Split(cmtText, " ")
            glodExtendMod(ws.name) = replace(CStr(arrTemp(0)), "外部文件|", "")  '得到对应的工作表名
        ElseIf InStr(ws.name, "额外") > 0 Or InStr(cmtText, "不参与") > 0 Or InStr(ws.Range("a1").value, "隐藏") > 0 Then
        
        Else
            Set t = New CMyHzModTable
            Call t.InitTableInfo(ws)
            'If InStr(ws.Name, "人民币") > 0 Then t.testdRowParse
            Set glodHzMod(ws.name) = t
        End If
    Next i
    wb.Close False
    
    Call SetMenuTitle
    InitModFileMessage = True
Exit Function

errhandle:
    MsgBox "InitModFileMessage初始化模板设置异常-----" & err.Description
    err.Clear
    InitModFileMessage = False
End Function

Public Function InitExtendModFileMessage() As Boolean
    On Error Resume Next
    '
    If glodExtendMod.count = 0 Then InitExtendModFileMessage = True: Exit Function
    
    Dim extendFileName As String, k As Variant, i As Long
    Dim modname As String
    Dim wb_extend As Workbook, ws_extend As Worksheet, wb_mod As Workbook, ws_mod As Worksheet
    
    modname = CStr(m_wsSys.Cells(1, 2).value)
    extendFileName = CStr(m_wsSys.Cells(2, 2).value)
    
    If extendFileName = "" Then
        MsgBox "模板设置中存在需要读取的外部数据，但未选择外部文件。"
        InitExtendModFileMessage = False
        Exit Function
    End If
    '打开工作簿
    Set wb_extend = gloEt.GetWb(extendFileName)
    If wb_extend Is Nothing Then
        MsgBox "打开外部文件失败。" & extendFileName
        InitExtendModFileMessage = False
        Exit Function
    End If
    Set wb_mod = gloEt.GetWb(modname)
    If wb_mod Is Nothing Then
        MsgBox "打开模板文件失败。" & modname
        InitExtendModFileMessage = False
        Exit Function
    End If
    
'    Dim reg As Object, mc As Object, m As Object, strPattern As String
'    Set reg = CreateObject("VBSCRIPT.REGEXP")
    
    Dim result As String, arr As Variant, rn As Range, is_find As Boolean
    result = ""
    For Each k In glodExtendMod.keys
        is_find = False
        Set ws_mod = gloEt.FindWs(wb_mod, CStr(k), 0)
        ws_mod.Cells.ClearContents  '清除内容
'        With reg
'            .Global = True: .ignorecase = False: .Pattern = CStr(glodExtendMod(k))
'        End With
        For Each ws_extend In wb_extend.Worksheets
'            If reg.text(ws_extend.name) = True Then
            If InStr(ws_extend.name, CStr(glodExtendMod(k))) > 0 Then
                Set rn = gloEt.GetUsedRange(ws_extend)
                If Not rn Is Nothing Then
                    '执行复制
                    If InStr(ws_mod.name, "参照") > 0 Then ws_mod.Cells.NumberFormat = "@"
                    ws_mod.Range(rn.address).value = rn.value
                End If
                is_find = True
                Exit For
            End If
        Next
        If is_find = False Then
            result = result & Chr(10) & "读取外部文件数据失败，模板" & CStr(ws_mod.name) & "未能读取外部文件内容，检查外部文件是否设置好对应工作表。"
        End If
    Next
    
    wb_extend.Close False
    wb_mod.Close True
    
    If err.Number <> 0 Then
        MsgBox "InitExtendModFileMessage读取外部文件数据异常-----" & err.Description
        err.Clear
        InitExtendModFileMessage = False
    Else
        If result <> "" Then
            MsgBox result
            InitExtendModFileMessage = False
        Else
            InitExtendModFileMessage = True
        End If
    End If
End Function
