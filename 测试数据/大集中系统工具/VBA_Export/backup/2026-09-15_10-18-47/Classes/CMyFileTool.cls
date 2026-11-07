VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyFileTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   1.关闭某些提示
'
Private Sub CloseSomething()

End Sub

'
'   2.开启某些提示
'
Private Sub OpenSomething()

End Sub

'
'   读取文本文件
'
Function GetFileList(folder As String, Optional is_folder As Boolean = False) As Variant
    Dim objFso As Object
    Dim objFolder As Object
    Dim strInfo As String
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFso.GetFolder(folder)
    strInfo = strEnumFile(objFolder, is_folder)
    If Len(strInfo) > 1 Then
        If right(strInfo, 1) = "|" Then
            strInfo = left(strInfo, Len(strInfo) - 1)
        End If
    End If
    Set objFso = Nothing
    Set objFolder = Nothing
    GetFileList = Split(strInfo, "|")
End Function
Function strEnumFile(ByVal objFolder As Object, Optional is_folder As Boolean = False) As String
    Dim objSubFolder As Object
    Dim objFile As Object
    Dim strList As String
    If is_folder = True Then strList = objFolder.path & "|"
    If objFolder.Files.count <> 0 Then
        For Each objFile In objFolder.Files
            strList = strList & objFile.path & "|"
        Next objFile
    End If
    If objFolder.SubFolders.count <> 0 Then
        For Each objSubFolder In objFolder.SubFolders
            strList = strList & strEnumFile(objSubFolder)
        Next objSubFolder
    End If
    strEnumFile = strList
    Set objSubFolder = Nothing
    Set objFile = Nothing
End Function


'
'   重命名文件
'
Public Function RenameFile(oldName As String, newName As String) As Boolean
    On Error GoTo errHandle
    Name oldName As newName
    
    RenameFile = True
Exit Function
errHandle:
    Call gloDe.InsertNewDebug("重命名", oldName, "错误", err.Description)
    'Debug.Print Err.Description
    err.Clear
    RenameFile = False
End Function

'
'   删除文件
'
Public Function KillFile(name As String) As Boolean
    On Error GoTo errHandle
    Kill name
    
    KillFile = True
Exit Function
errHandle:
    Debug.Print err.Description
    err.Clear
    KillFile = False
End Function

Public Function GetOutputFolder(Optional title As String = "请选择输出文件夹") As String
    Call OpenSomething
    
    Dim savePath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If .Show = -1 Then
        'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
            savePath = CStr(.SelectedItems(1))
        Else
            savePath = ""
        End If
    End With
    
    GetOutputFolder = savePath
    
    Call CloseSomething
End Function

Public Function GetOutputFile() As String
    
    Call OpenSomething
    Dim fileName As String
    
    fileName = Application.GetSaveAsFilename(InitialFileName:="", fileFilter:="Excel files(*.xls),*.xls,All files (*.*),*.*", title:="选择输出文件")
        
    If fileName <> "False" Then
        GetOutputFile = fileName
    End If
    
    Call CloseSomething
End Function


'
'   4.判断是否存在指定文件
'
Function IsFileExist(strPath As String) As Boolean
    IsFileExist = (Dir(strPath, vbNormal + vbHidden + vbReadOnly) <> "")
End Function

'
'   5.判断是否存在指定文件夹
'
Function IsFolderExist(strPath As String) As Boolean
    IsFolderExist = (Dir(strPath, vbDirectory + vbHidden) <> "")
End Function

'
'   7.判断文件是否被打开
'
Function IsOpenedFile(strFile As String) As Boolean
    On Error Resume Next
    Open strFile For Binary Access Read Write Lock Read Write As #1
    Close #1
    If err.Number <> 0 Then
        IsOpenedFile = True
        err.Clear
    End If
End Function

'
'   6.获取不含后缀的文件名
'
Public Function getShortFileName(fileName As String) As String
    Dim a As Variant
    Dim b As Variant
    
    a = Split(fileName, "\")
    
    b = UBound(a)
    getShortFileName = left(a(b), InStrRev(a(b), ".") - 1)
End Function

'
'   6.获取含后缀的文件名
'
Public Function getShortFileName2(fileName As String) As String
    Dim a As Variant
    Dim b As Variant
    
    a = Split(fileName, "\")
    
    b = UBound(a)
    getShortFileName2 = left(a(b), Len(a(b)))
End Function
