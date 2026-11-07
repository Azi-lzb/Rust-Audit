VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMySelCmda"
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
'   打开选择文件框,单选模式
'
Private Sub cmd_Click()
    Dim fileName As Variant
    
    If glo_fileFilter = "" Then
        glo_fileFilter = "Excel 文件 (*.xls; *.csv), *.xls; *.csv,所有文件(*.*),*.*"
    End If
    
    fileName = Application.GetOpenFilename(fileFilter:=glo_fileFilter, title:=cmd.Caption, MultiSelect:=False)
    
    If VarType(fileName) = vbBoolean Then
        frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value = ""
    ElseIf VarType(fileName) = vbString Then
        '获取控件窗体句柄
        frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value = fileName
    Else
        frmT.Controls("tb" & replace(cmd.name, "cmd", "")).value = ""
        MsgBox "出现不兼容"
    End If
End Sub



