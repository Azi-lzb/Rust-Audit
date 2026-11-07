Attribute VB_Name = "VbaBaselineBatch"
Option Explicit

' Expression V2 VBA baseline helper.
'
' Folder contract:
'   <case>_current.csv
'   <case>_previous.csv
'
' Results are written to a VBA基线结果 child folder.  This module deliberately
' uses the existing comparison macro: it does not duplicate comparison rules.

Public Sub VbaBaseline_SelectFolderAndRun()
    Dim folderPath As String

    folderPath = VbaBaseline_PickFolder()
    If Len(folderPath) = 0 Then Exit Sub

    VbaBaseline_RunFolder folderPath
End Sub

Public Sub VbaBaseline_AddButton()
    Dim buttonShape As Shape

    On Error Resume Next
    ActiveSheet.Shapes("VbaBaselineBatchButton").Delete
    On Error GoTo 0

    Set buttonShape = ActiveSheet.Shapes.AddShape(5, 20, 20, 190, 32)
    buttonShape.name = "VbaBaselineBatchButton"
    buttonShape.TextFrame.Characters.text = "VBA baseline batch"
    buttonShape.OnAction = "'" & ThisWorkbook.name & "'!VbaBaseline_SelectFolderAndRun"
End Sub

Public Sub VbaBaseline_RunFolder(ByVal folderPath As String)
    Dim fs As Object
    Dim folder As Object
    Dim file As Object
    Dim currentPath As String
    Dim previousPath As String
    Dim resultFolder As String
    Dim outputPath As String
    Dim caseName As String
    Dim compared As Long
    Dim skipped As Long

    On Error GoTo Failed

    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(folderPath) Then
        err.Raise vbObjectError + 2101, , "Folder does not exist: " & folderPath
    End If

    Set folder = fs.GetFolder(folderPath)
    resultFolder = fs.BuildPath(folderPath, "VBA_baseline_results")
    If Not fs.FolderExists(resultFolder) Then fs.CreateFolder resultFolder

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ShowDoCompare loads the existing modeless default UserForm.  The original
    ' public macro reads tb1/tb2 from that default instance.
    Application.Run "ShowDoCompare"

    For Each file In folder.Files
        currentPath = CStr(file.path)
        If LCase$(right$(currentPath, 12)) = "_current.csv" Then
            previousPath = left$(currentPath, Len(currentPath) - 12) & "_previous.csv"
            If fs.FileExists(previousPath) Then
                caseName = fs.GetBaseName(currentPath)
                caseName = left$(caseName, Len(caseName) - Len("_current"))
                outputPath = fs.BuildPath(resultFolder, caseName & "_vba_compare.xls")

                VbaBaseline_SetInputPaths currentPath, previousPath
                VbaBaseline_RunExistingCompare outputPath
                compared = compared + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next file

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "VBA baseline batch finished. Compared: " & compared & _
           "; skipped (no previous pair): " & skipped & _
           "; output: " & resultFolder, vbInformation
    Exit Sub

Failed:
    MsgBox "VBA baseline batch stopped: " & err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function VbaBaseline_PickFolder() As String
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .title = "Select folder containing *_current.csv and *_previous.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then VbaBaseline_PickFolder = CStr(.SelectedItems(1))
    End With
End Function

Private Sub VbaBaseline_SetInputPaths(ByVal currentPath As String, ByVal previousPath As String)
    Dim formObject As Object

    Set formObject = VbaBaseline_FindCompareForm()
    If formObject Is Nothing Then
        err.Raise vbObjectError + 2102, , "Comparison form was not loaded."
    End If

    formObject.Controls("tb1").Value = currentPath
    formObject.Controls("tb2").Value = previousPath
End Sub

Private Function VbaBaseline_FindCompareForm() As Object
    Dim item As Object

    For Each item In VBA.UserForms
        On Error Resume Next
        If Not item.Controls("tb1") Is Nothing And Not item.Controls("tb2") Is Nothing Then
            Set VbaBaseline_FindCompareForm = item
            Exit Function
        End If
        err.Clear
        On Error GoTo 0
    Next item
End Function

Private Sub VbaBaseline_RunExistingCompare(ByVal outputPath As String)
    ' The legacy entry displays a date-confirmation prompt and a completion
    ' prompt. Queue the confirmations so every pair still runs through the
    ' original VBA procedure without copying its business logic.
    Application.SendKeys "{ENTER}{ENTER}", False
    Application.Run VbaBaseline_CompareMacroName(), outputPath
End Sub

Private Function VbaBaseline_CompareMacroName() As String
    ' "导数比较控制�?" expressed as Unicode code points, keeping this exported
    ' module ASCII-only so vbaSync can import it under legacy GBK environments.
    VbaBaseline_CompareMacroName = ChrW(&H5BFC) & ChrW(&H6570) & _
        ChrW(&H6BD4) & ChrW(&H8F83) & ChrW(&H63A7) & ChrW(&H5236) & _
        ChrW(&H53F0) & "2"
End Function
