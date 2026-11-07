Attribute VB_Name = "vbaSync"
Option Explicit

' Requires reference: Microsoft Visual Basic for Applications Extensibility 5.3 (Tools -> References -> check it)
' Batch export/import VBA modules. Export: standard modules, class modules, document modules, forms. Skip by module name.
' д RUNLOG_SHEET_NAME
' 鲻
'
' ================== Export ==================
' exportRoot: empty = ThisWorkbook.Path\VBA_Export\
' charset: used for file read/write (default GBK)
' skipModules: comma-separated module names, e.g. "vbaSync"
Private Const RUNLOG_SHEET_NAME As String = "运行日志"
'  False д12УУ/ Environ
Private Const ENABLE_RUNLOG As Boolean = True
' к12У顢ID/()
Private Const RUNLOG_COL_SEQ As Long = 1
Private Const RUNLOG_COL_TIME As Long = 2
Private Const RUNLOG_COL_USER As Long = 3
Private Const RUNLOG_COL_MODULE As Long = 4
Private Const RUNLOG_COL_OP As Long = 5
Private Const RUNLOG_COL_OBJ As Long = 6
Private Const RUNLOG_COL_BEFORE As Long = 7
Private Const RUNLOG_COL_AFTER As Long = 8
Private Const RUNLOG_COL_RESULT As Long = 9
Private Const RUNLOG_COL_DETAIL As Long = 10
Private Const RUNLOG_COL_ELAPSED As Long = 11
Private Const RUNLOG_COL_PC As Long = 12

Public Sub ExportVBAModules( _
    Optional ByVal exportRoot As String, _
    Optional ByVal charset As String = "GBK", _
    Optional ByVal skipModules As String = "")
    ' Batch export VBA modules (requires VB Extensibility 5.3)

    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim pathModules As String
    Dim pathClasses As String
    Dim pathForms As String
    Dim pathDocuments As String
    Dim targetPath As String
    Dim skipArr As Variant

    On Error GoTo ErrTrust
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0

    If Len(exportRoot) = 0 Then exportRoot = ThisWorkbook.path
    If right$(exportRoot, 1) <> "\" Then exportRoot = exportRoot & "\"
    exportRoot = exportRoot & "VBA_Export\"

    pathModules = exportRoot & "Modules\"
    pathClasses = exportRoot & "Classes\"
    pathForms = exportRoot & "Forms\"
    pathDocuments = exportRoot & "Documents\"

    EnsureFolder exportRoot
    EnsureFolder pathModules
    EnsureFolder pathClasses
    EnsureFolder pathForms
    EnsureFolder pathDocuments

    EnsureRunLogSheet
    Dim t0 As Double: t0 = Timer
    Dim outName As String
    skipArr = BuildSkipList(skipModules)

    RunLog_WriteRow "vba", "", exportRoot, "", "", "", "", ""

    For Each vbComp In vbProj.VBComponents
        If IsSkipped(vbComp.name, skipArr) Then
        Else
            On Error GoTo ExportErr
            outName = ""
            Select Case vbComp.Type
                Case vbext_ct_StdModule
                    targetPath = pathModules & GetExportBasFileName(vbComp.name) & ".bas"
                    outName = GetExportBasFileName(vbComp.name) & ".bas"
                    ExportStdModule vbComp, targetPath, charset
                Case vbext_ct_ClassModule
                    targetPath = pathClasses & GetExportBasFileName(vbComp.name) & ".cls"
                    outName = GetExportBasFileName(vbComp.name) & ".cls"
                    vbComp.Export targetPath
                    RewriteFileAsCharset targetPath, charset
                Case vbext_ct_Document
                    targetPath = pathDocuments & GetExportDocFileName(vbComp.name) & ".cls"
                    outName = GetExportDocFileName(vbComp.name) & ".cls"
                    ExportDocumentModule vbComp, targetPath, charset
                Case vbext_ct_MSForm
                    targetPath = pathForms & vbComp.name & ".frm"
                    outName = vbComp.name & ".frm"
                    vbComp.Export targetPath
            End Select
            RunLog_WriteRow "vba", "", outName, "", "", "", "OK", ""
            On Error GoTo 0
            GoTo NextExport
ExportErr:
            RunLog_WriteRow "vba", "", outName, "", "", "", err.Number & " " & err.Description, ""
            On Error GoTo 0
NextExport:
        End If
    Next vbComp

    RunLog_WriteRow "vba", "", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
    MsgBox "VBA modules exported. Folder:" & vbCrLf & exportRoot & vbCrLf & "", vbInformation
    Exit Sub

ErrTrust:
    If err.Number = 1004 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in File -> Options -> Trust Center -> Trust Center Settings.", vbExclamation
    Else
        MsgBox "Export failed: " & err.Number & " " & err.Description & vbCrLf & "If you see 'User-defined type not defined', add reference: Tools -> References -> Microsoft Visual Basic for Applications Extensibility 5.3.", vbExclamation
    End If
End Sub

' ================== Import ==================
' importRoot: empty = ThisWorkbook.Path\VBA_Export\
' Before import: backup current VBA to VBA_Export\backup\yyyy-mm-dd_hh-nn-ss\, then clear all removable components (std/class/form) to avoid orphan modules after rename.
Public Sub ImportVBAModules( _
    Optional ByVal importRoot As String, _
    Optional ByVal charset As String = "GBK", _
    Optional ByVal skipModules As String = "")
    ' Batch import VBA modules (requires VB Extensibility 5.3)

    Dim vbProj As VBIDE.VBProject
    Dim fso As Object
    Dim folderRoot As Object
    Dim folder As Object
    Dim f As Object
    Dim skipArr As Variant
    Dim backupPath As String

    On Error GoTo ErrTrust
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0

    If Len(importRoot) = 0 Then importRoot = ThisWorkbook.path
    If right$(importRoot, 1) <> "\" Then importRoot = importRoot & "\"
    importRoot = importRoot & "VBA_Export\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(importRoot) Then
        MsgBox "Import folder not found:" & vbCrLf & importRoot, vbExclamation
        Exit Sub
    End If

    EnsureRunLogSheet
    Dim t0 As Double: t0 = Timer

    RunLog_WriteRow "vba", "", importRoot, "", "", "", "", ""

    ' 1) Backup current VBA to export\backup\yyyy-mm-dd_hh-nn-ss\
    backupPath = importRoot & "backup\" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & "\"
    EnsureFolder importRoot & "backup"
    EnsureFolder left(backupPath, Len(backupPath) - 1)
    BackupCurrentProject vbProj, backupPath, charset
    RunLog_WriteRow "vba", "", backupPath, "", "", "", "OK", ""

    ' 2) Clear all removable components (StdModule, ClassModule, MSForm); keep Document (ThisWorkbook, Sheets)
    ClearRemovableComponents vbProj
    RunLog_WriteRow "vba", "", "", "", "", "", "OK", ""

    Set folderRoot = fso.GetFolder(importRoot)
    skipArr = BuildSkipList(skipModules)

    For Each folder In folderRoot.SubFolders
        For Each f In folder.Files
            ImportOneFile vbProj, f, skipArr, charset
        Next f
    Next folder

    For Each f In folderRoot.Files
        ImportOneFile vbProj, f, skipArr, charset
    Next f

    RunLog_WriteRow "vba", "", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
    MsgBox "VBA modules imported. ", vbInformation
    Exit Sub

ErrTrust:
    If err.Number = 1004 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in File -> Options -> Trust Center -> Trust Center Settings.", vbExclamation
    Else
        MsgBox "Import failed: " & err.Number & " " & err.Description & vbCrLf & "If you see 'User-defined type not defined', add reference: Tools -> References -> Microsoft Visual Basic for Applications Extensibility 5.3.", vbExclamation
    End If
End Sub

' Backup entire current VBA project to backupRoot (Modules\, Classes\, Documents\, Forms\). No skip.
' δ塹δ塹 ->  ->  "Microsoft Visual Basic for Applications Extensibility 5.3"
Private Sub BackupCurrentProject(ByVal vbProj As VBIDE.VBProject, ByVal backupRoot As String, ByVal charset As String)
    Dim vbComp As VBIDE.VBComponent
    Dim pathModules As String
    Dim pathClasses As String
    Dim pathForms As String
    Dim pathDocuments As String
    Dim targetPath As String

    pathModules = backupRoot & "Modules\"
    pathClasses = backupRoot & "Classes\"
    pathForms = backupRoot & "Forms\"
    pathDocuments = backupRoot & "Documents\"

    EnsureFolder pathModules
    EnsureFolder pathClasses
    EnsureFolder pathForms
    EnsureFolder pathDocuments

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                targetPath = pathModules & GetExportBasFileName(vbComp.name) & ".bas"
                ExportStdModule vbComp, targetPath, charset
            Case vbext_ct_ClassModule
                targetPath = pathClasses & GetExportBasFileName(vbComp.name) & ".cls"
                vbComp.Export targetPath
                RewriteFileAsCharset targetPath, charset
            Case vbext_ct_Document
                targetPath = pathDocuments & GetExportDocFileName(vbComp.name) & ".cls"
                ExportDocumentModule vbComp, targetPath, charset
            Case vbext_ct_MSForm
                targetPath = pathForms & vbComp.name & ".frm"
                vbComp.Export targetPath
        End Select
    Next vbComp
End Sub

' Remove all StdModule, ClassModule, MSForm. Keep Document (ThisWorkbook, Sheet modules). Keep StdModule "vbaSync" so Import can replace its code only (otherwise sync would leave Excel without vbaSync).
Private Sub ClearRemovableComponents(ByVal vbProj As VBIDE.VBProject)
    Dim i As Long
    Dim vbComp As VBIDE.VBComponent

    For i = vbProj.VBComponents.count To 1 Step -1
        Set vbComp = vbProj.VBComponents(i)
        If vbComp.Type = vbext_ct_ClassModule Or vbComp.Type = vbext_ct_MSForm Then
            vbProj.VBComponents.Remove vbComp
        ElseIf vbComp.Type = vbext_ct_StdModule Then
            If StrComp(vbComp.name, "vbaSync", vbTextCompare) <> 0 Then
                vbProj.VBComponents.Remove vbComp
            End If
        End If
    Next i
End Sub

' ==================  (RunLog) д ==================
' 12У顢ID/()/

' д12У On Error Resume Next
Private Sub EnsureRunLogSheet()
    If Not ENABLE_RUNLOG Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        If Not ws Is Nothing Then ws.name = RUNLOG_SHEET_NAME
    End If
    If ws Is Nothing Then Exit Sub
    With ws
        .Cells(1, RUNLOG_COL_SEQ).Value = "序号"
        .Cells(1, RUNLOG_COL_TIME).Value = "时间"
        .Cells(1, RUNLOG_COL_USER).Value = "用户名"
        .Cells(1, RUNLOG_COL_MODULE).Value = "模块"
        .Cells(1, RUNLOG_COL_OP).Value = "动作"
        .Cells(1, RUNLOG_COL_OBJ).Value = "对象ID/路径"
        .Cells(1, RUNLOG_COL_BEFORE).Value = "前值"
        .Cells(1, RUNLOG_COL_AFTER).Value = "后值"
        .Cells(1, RUNLOG_COL_RESULT).Value = "结果"
        .Cells(1, RUNLOG_COL_DETAIL).Value = "详情"
        .Cells(1, RUNLOG_COL_ELAPSED).Value = "耗时(秒)"
        .Cells(1, RUNLOG_COL_PC).Value = "电脑名"
        .Range(.Cells(1, 1), .Cells(1, RUNLOG_COL_PC)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, RUNLOG_COL_PC)).Interior.color = RGB(220, 230, 241)
    End With
End Sub

Public Sub RunLog_WriteRow(ByVal moduleName As String, ByVal operation As String, ByVal objectId As String, _
    ByVal beforeValue As String, ByVal afterValue As String, ByVal resultText As String, ByVal detailText As String, ByVal elapsedText As String)
    If Not ENABLE_RUNLOG Then Exit Sub
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim seq As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET_NAME)
    If ws Is Nothing Then EnsureRunLogSheet: Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    nextRow = ws.Cells(ws.rows.count, RUNLOG_COL_SEQ).End(xlUp).row + 1
    If nextRow < 2 Then Exit Sub
    seq = nextRow - 1
    ws.Cells(nextRow, RUNLOG_COL_SEQ).Value = seq
    ws.Cells(nextRow, RUNLOG_COL_TIME).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(nextRow, RUNLOG_COL_USER).Value = ""
    ws.Cells(nextRow, RUNLOG_COL_MODULE).Value = moduleName
    ws.Cells(nextRow, RUNLOG_COL_OP).Value = operation
    ws.Cells(nextRow, RUNLOG_COL_OBJ).Value = objectId
    ws.Cells(nextRow, RUNLOG_COL_BEFORE).Value = beforeValue
    ws.Cells(nextRow, RUNLOG_COL_AFTER).Value = afterValue
    ws.Cells(nextRow, RUNLOG_COL_RESULT).Value = resultText
    ws.Cells(nextRow, RUNLOG_COL_DETAIL).Value = detailText
    ws.Cells(nextRow, RUNLOG_COL_ELAPSED).Value = elapsedText
    ws.Cells(nextRow, RUNLOG_COL_PC).Value = ""
End Sub

Public Sub RunLog_Append(ByVal operation As String, Optional ByVal detail As String = "")
    If Not ENABLE_RUNLOG Then Exit Sub
    RunLog_WriteRow "", operation, "", "", "", "", detail, ""
End Sub

' ================== Internal helpers ==================

Private Sub ImportOneFile(ByVal vbProj As VBIDE.VBProject, ByVal f As Object, ByVal skipArr As Variant, ByVal charset As String)
    ' Import one file; on success/fail write one log row each
    Dim ext As String
    Dim moduleName As String
    Dim vbComp As VBIDE.VBComponent

    moduleName = FileBaseName(f.name)
    ext = LCase$(FileExt(f.name))
    If IsSkipped(moduleName, skipArr) And Not (ext = "bas" And StrComp(moduleName, "vbaSync", vbTextCompare) = 0) Then Exit Sub

    On Error GoTo ImportErr

    Select Case ext
        Case "bas"
            If StrComp(moduleName, "vbaSync", vbTextCompare) = 0 Then
                Set vbComp = FindComponent(vbProj, "vbaSync")
                If Not vbComp Is Nothing Then ReplaceCodeFromBasFile vbComp, f.path, charset
            Else
                Set vbComp = FindComponent(vbProj, moduleName)
                If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
                vbProj.VBComponents.Import f.path
            End If
        Case "cls"
            Set vbComp = FindComponent(vbProj, moduleName)
            If Not vbComp Is Nothing Then
                If vbComp.Type = vbext_ct_Document Then
                    ReplaceCodeFromFile vbComp, f.path, charset
                Else
                    vbProj.VBComponents.Remove vbComp
                    vbProj.VBComponents.Import f.path
                End If
            Else
                vbProj.VBComponents.Import f.path
            End If
        Case "frm"
            Set vbComp = FindComponent(vbProj, moduleName)
            If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
            vbProj.VBComponents.Import f.path
    End Select

    RunLog_WriteRow "vba", "", f.name, "", "", "", "OK", ""
    Exit Sub

ImportErr:
    RunLog_WriteRow "vba", "", f.name, "", "", "", err.Number & " " & err.Description, ""
End Sub

Private Function FindComponent(ByVal vbProj As VBIDE.VBProject, ByVal compName As String) As VBIDE.VBComponent
    ' Find component by name
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In vbProj.VBComponents
        If StrComp(vbComp.name, compName, vbTextCompare) = 0 Then
            Set FindComponent = vbComp
            Exit Function
        End If
    Next vbComp
    Set FindComponent = Nothing
End Function

Private Sub ReplaceCodeFromFile(ByVal vbComp As VBIDE.VBComponent, ByVal filePath As String, ByVal charset As String)
    ' Replace code from file (doc/class): insert code only, strip .cls header
    Dim codeText As String
    Dim codeBody As String
    Dim codeMod As VBIDE.CodeModule
    codeText = ReadTextFile(filePath, charset)
    codeBody = StripClsHeader(codeText)
    Set codeMod = vbComp.CodeModule
    With codeMod
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(Trim$(codeBody)) > 0 Then .InsertLines 1, codeBody
    End With
End Sub

' For vbaSync.bas only: replace code only, do not sync Attribute VB_Name
Private Sub ReplaceCodeFromBasFile(ByVal vbComp As VBIDE.VBComponent, ByVal filePath As String, ByVal charset As String)
    Dim codeText As String
    Dim codeBody As String
    Dim codeMod As VBIDE.CodeModule
    codeText = ReadTextFile(filePath, charset)
    codeBody = StripBasAttributeVBName(codeText)
    Set codeMod = vbComp.CodeModule
    With codeMod
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(Trim$(codeBody)) > 0 Then .InsertLines 1, codeBody
    End With
End Sub

' Remove Attribute VB_Name = "xxx" line in .bas content
Private Function StripBasAttributeVBName(ByVal codeText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim result As String
    If Len(codeText) = 0 Then StripBasAttributeVBName = "": Exit Function
    lines = Split(replace(replace(codeText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    result = ""
    For i = 0 To UBound(lines)
        line = lines(i)
        If LCase(Trim(line)) Like "attribute vb_name*" Then
        Else
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & line
        End If
    Next i
    StripBasAttributeVBName = result
End Function

' Remove .cls class file header, keep only code for CodeModule
Private Function StripClsHeader(ByVal codeText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim inBeginBlock As Boolean
    Dim result As String
    If Len(codeText) = 0 Then StripClsHeader = "": Exit Function
    lines = Split(replace(replace(codeText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    inBeginBlock = False
    result = ""
    For i = 0 To UBound(lines)
        line = lines(i)
        If inBeginBlock Then
            If Trim(line) = "END" Then inBeginBlock = False
        ElseIf Trim(line) = "VERSION 1.0 CLASS" Then
        ElseIf Trim(line) = "BEGIN" Then
            inBeginBlock = True
        ElseIf Trim(line) = "END" And Not inBeginBlock Then
        ElseIf LCase(Trim(line)) Like "attribute*" Then
        ElseIf InStr(1, line, "MultiUse = -1", vbTextCompare) > 0 Then
        Else
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & line
        End If
    Next i
    StripClsHeader = result
End Function

Private Function GetExportBasFileName(ByVal compName As String) As String
    ' Export bas/cls file name (same as module name)
    GetExportBasFileName = compName
End Function

Private Function GetExportDocFileName(ByVal compName As String) As String
    ' Export document module file name
    GetExportDocFileName = compName
End Function

' Export document module (Sheet/ThisWorkbook): export then rewrite as code-only with GBK
Private Sub ExportDocumentModule(ByVal vbComp As VBIDE.VBComponent, ByVal targetPath As String, ByVal charset As String)
    Dim content As String
    Dim codeOnly As String
    vbComp.Export targetPath
    content = ReadTextFile(targetPath, charset)
    codeOnly = StripClsHeader(content)
    If Len(Trim$(codeOnly)) = 0 Then
        codeOnly = "' " & vbComp.name & " has no code"
    End If
    WriteTextFile targetPath, codeOnly, charset
End Sub

' Export standard module: rewrite file with correct Attribute VB_Name and GBK
Private Sub ExportStdModule(ByVal vbComp As VBIDE.VBComponent, ByVal targetPath As String, ByVal charset As String)
    Dim content As String
    Dim fixedContent As String
    Dim exportName As String
    vbComp.Export targetPath
    exportName = FileBaseName(Mid(targetPath, InStrRev(targetPath, "\") + 1))
    content = ReadTextFile(targetPath, charset)
    '  Attribute VB_Nameɡ1
    fixedContent = ReplaceFirstAttributeVBName(content, exportName)
    WriteTextFile targetPath, fixedContent, charset
End Sub

' After VBComponent.Export, re-save .cls with specified charset (Export uses system default)
Private Sub RewriteFileAsCharset(ByVal targetPath As String, ByVal charset As String)
    Dim content As String
    content = ReadTextFile(targetPath, "")
    If Len(content) > 0 Then WriteTextFile targetPath, content, charset
End Sub

' Replace first Attribute VB_Name = "xxx" with newName
Private Function ReplaceFirstAttributeVBName(ByVal codeText As String, ByVal newName As String) As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim line As String
    If Len(codeText) = 0 Then ReplaceFirstAttributeVBName = codeText: Exit Function
    lines = Split(replace(replace(codeText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    result = ""
    For i = 0 To UBound(lines)
        line = lines(i)
        If LCase(Trim(line)) Like "attribute vb_name*" And Len(result) = 0 Then
            result = "Attribute VB_Name = """ & newName & """"
        Else
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & line
        End If
    Next i
    ReplaceFirstAttributeVBName = result
End Function

Private Function FileExt(ByVal fileName As String) As String
    ' Get file extension
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then FileExt = Mid$(fileName, pos + 1) Else FileExt = ""
End Function

Private Function FileBaseName(ByVal fileName As String) As String
    ' Get base name without extension
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then FileBaseName = left$(fileName, pos - 1) Else FileBaseName = fileName
End Function

Private Function ExportFileName(ByVal fullPath As String) As String
    ' Get file name from full path (for export log)
    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then ExportFileName = Mid$(fullPath, pos + 1) Else ExportFileName = fullPath
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    ' Create folder if missing
    If Len(folderPath) = 0 Then Exit Sub
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Function BuildSkipList(ByVal skipModules As String) As Variant
    ' Build skip list from comma-separated module names
    If Len(Trim$(skipModules)) = 0 Then
        BuildSkipList = Array()
    Else
        BuildSkipList = Split(skipModules, ",")
    End If
End Function

Private Function IsSkipped(ByVal moduleName As String, ByVal skipArr As Variant) As Boolean
    ' Whether module is in skip list
    Dim i As Long
    Dim nameTrim As String
    On Error GoTo HandleEmpty
    nameTrim = Trim$(moduleName)
    For i = LBound(skipArr) To UBound(skipArr)
        If StrComp(nameTrim, Trim$(CStr(skipArr(i))), vbTextCompare) = 0 Then
            IsSkipped = True
            Exit Function
        End If
    Next i
    Exit Function
HandleEmpty:
    IsSkipped = False
End Function

' Read text file with given charset (GBK for .bas/.cls). Empty = system default.
Private Function ReadTextFile(ByVal path As String, Optional ByVal charset As String = "GBK") As String
    Dim fso As Object
    Dim stream As Object
    If Len(Trim$(charset)) = 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        ReadTextFile = fso.OpenTextFile(path, 1, False).ReadAll
        Exit Function
    End If
    On Error GoTo UseFso
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.charset = charset
    stream.Open
    stream.LoadFromFile path
    ReadTextFile = stream.ReadText(-1)
    stream.Close
    Exit Function
UseFso:
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReadTextFile = fso.OpenTextFile(path, 1, False).ReadAll
End Function

' Write text file with given charset (GBK for .bas/.cls). Empty = system default.
Private Sub WriteTextFile(ByVal path As String, ByVal content As String, Optional ByVal charset As String = "GBK")
    Dim fso As Object
    Dim stream As Object
    If Len(Trim$(charset)) = 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        With fso.CreateTextFile(path, True, False)
            .Write content
            .Close
        End With
        Exit Sub
    End If
    On Error GoTo UseFso
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.charset = charset
    stream.Open
    stream.WriteText content
    stream.SaveToFile path, 2
    stream.Close
    Exit Sub
UseFso:
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(path, True, False)
        .Write content
        .Close
    End With
End Sub

' ================== Shortcut entries ==================

Public Sub Sync_ExportAll()
    ' One-click export (skip vbaSync module)
    Call ExportVBAModules("", "GBK", "vbaSync")
End Sub

Public Sub Sync_ImportAll()
    ' One-click import (skip vbaSync module)
    Call ImportVBAModules("", "GBK", "vbaSync")
End Sub







