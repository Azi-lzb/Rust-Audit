VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMySelCmdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'类模块CMySelCmdOne的代码
Public WithEvents cmd As MSForms.CommandButton
Attribute cmd.VB_VarHelpID = -1
Public WithEvents frmT As MSForms.UserForm
Attribute frmT.VB_VarHelpID = -1


'
'   打开选择文件框,多选模式
'
Private Sub cmd_Click()
    Dim fileName As Variant, str As Variant, strs As String
    
    If glo_fileFilter = "" Then
        glo_fileFilter = "Excel 文件 (*.xls; *.csv), *.xls; *.csv,所有文件(*.*),*.*"
    End If
    
    fileName = Application.GetOpenFilename(fileFilter:=glo_fileFilter, title:=cmd.Caption, MultiSelect:=True)
    
    If VarType(fileName) = vbBoolean Then
        frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value = ""
    ElseIf VarType(fileName) = vbArray Or VarType(fileName) = 8200 Or VarType(fileName) = 8204 Then
        '获取控件窗体句柄
        strs = frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value
        For Each str In fileName
            If strs = "" Then
                strs = CStr(str)
            Else
                strs = strs + "," + CStr(str)
            End If
        Next
        frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value = strs
    Else
        
    End If
End Sub




