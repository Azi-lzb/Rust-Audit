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
    Dim currentFolderPath As String
    Dim previousFolderPath As String

    currentFolderPath = VbaBaseline_PickFolder("Select current-data folder")
    If Len(currentFolderPath) = 0 Then Exit Sub

    previousFolderPath = VbaBaseline_PickFolder("Select previous-data folder")
    If Len(previousFolderPath) = 0 Then Exit Sub

    VbaBaseline_RunFolders currentFolderPath, previousFolderPath
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
    ' Compatibility helper for a root folder that contains current/previous.
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(fs.BuildPath(folderPath, "current")) And _
       fs.FolderExists(fs.BuildPath(folderPath, "previous")) Then
        VbaBaseline_RunFolders fs.BuildPath(folderPath, "current"), _
            fs.BuildPath(folderPath, "previous")
    Else
        MsgBox "Use VbaBaseline_SelectFolderAndRun to select current and previous folders.", vbExclamation
    End If
End Sub

Public Sub VbaBaseline_RunFolders(ByVal currentFolderPath As String, _
                                  ByVal previousFolderPath As String)
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
    If Not fs.FolderExists(currentFolderPath) Or Not fs.FolderExists(previousFolderPath) Then
        err.Raise vbObjectError + 2101, , "Current or previous folder does not exist."
    End If

    Set folder = fs.GetFolder(currentFolderPath)
    resultFolder = fs.BuildPath(currentFolderPath, "VBA_baseline_results")
    If Not fs.FolderExists(resultFolder) Then fs.CreateFolder resultFolder

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each file In folder.Files
        currentPath = CStr(file.path)
        If LCase$(right$(currentPath, 12)) = "_current.csv" Then
            previousPath = fs.BuildPath(previousFolderPath, _
                fs.GetFileName(left$(currentPath, Len(currentPath) - 12) & "_previous.csv"))
            If fs.FileExists(previousPath) Then
                caseName = fs.GetBaseName(currentPath)
                caseName = left$(caseName, Len(caseName) - Len("_current"))
                outputPath = fs.BuildPath(resultFolder, caseName & "_vba_compare.xls")

                VbaBaseline_RunExistingCompare outputPath, currentPath, previousPath
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

Private Function VbaBaseline_PickFolder(ByVal promptTitle As String) As String
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .title = promptTitle
        .AllowMultiSelect = False
        If .Show = -1 Then VbaBaseline_PickFolder = CStr(.SelectedItems(1))
    End With
End Function

Private Sub VbaBaseline_RunExistingCompare(ByVal outputPath As String, _
                                           ByVal currentPath As String, _
                                           ByVal previousPath As String)
    Application.Run "VbaBaseline_RunNoPrompt", outputPath, currentPath, previousPath
End Sub
