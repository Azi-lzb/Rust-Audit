VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 系统导数转EXCLE窗体 
   Caption         =   "系统导数转EXCLE窗体"
   ClientHeight    =   4215
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7210
   OleObjectBlob   =   "系统导数转EXCLE窗体.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "系统导数转EXCLE窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Cmda As CMySelCmda
Private Cmdb As CMySelCmdb


Private Sub Userform_Initialize()
    Dim units As Variant, u As Variant
    
    耗时标签.Visible = False
    
    Set Cmda = New CMySelCmda
    Set Cmda.cmd = Me.Controls("cmd1")
    Set Cmda.frmT = Me
    
    Set Cmdb = New CMySelCmdb
    Set Cmdb.cmd = Me.Controls("cmd2")
    Set Cmdb.frmT = Me
    
    Me.chkHideRow.Value = True
    
    units = Array("元", "十元", "百元", "千元", "万元", "十万元", "百万元", "千万元", "亿元")
    For Each u In units
        说明文件单位.AddItem u
    Next
    说明文件单位 = gloCft.GetConfig("系统导数比较转EXCEL说明文件单位历史选项")
    
'    Select Case Application.Version * 1    '设置连接字符串,根据版本创建连接
'    Case Is <= 11
'        Me.chkDelEmptyWs.Value = False
'        Me.chkDelEmptyWs.Visible = False
'    Case Is >= 12
'        Me.chkDelEmptyWs.Value = True
'    End Select
    Me.chkDelEmptyWs.Value = True
    Me.chkDeleteFm.Value = True
End Sub

Private Sub ReloadBtn_Click()
    On Error GoTo errHandle
    
    Dim savePath As String, r As Variant
    savePath = gloFt.GetOutputFolder
    耗时标签.Visible = True
    
    If Me.tb1.Value = "" Or Me.tb2.Value = "" Or savePath = "" Then
        MsgBox "漏选文件"
    Else
        r = MsgBox("请确认模板文件为" + Me.tb1.Value + Chr(10) + "比较文件为" + Me.tb2.Value, vbYesNo, Application.Caption)
        If r = vbYes Then
            gloT = Timer
            Call gloCft.CloseSomething
            Call 系统导数转表控制台(savePath)
            耗时标签.Visible = False
            Unload Me
            Call gloCft.OpenSomething
            Application.Calculation = xlCalculationAutomatic
            Debug.Print "重载文件耗时:" & CStr(Timer - gloT) & "s"
        End If
    End If
Exit Sub

errHandle:
    MsgBox "执行比较_Click 出错" + CStr(err.Description)
    Unload Me
    Application.Calculation = xlCalculationAutomatic
    Call gloCft.OpenSomething
End Sub

Private Sub 说明文件单位_Change()
    Call gloCft.SetConfig("系统导数比较转EXCEL说明文件单位历史选项", CStr(系统导数转EXCLE窗体.说明文件单位))
End Sub
