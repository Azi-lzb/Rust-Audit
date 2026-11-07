VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 导数比较窗体 
   Caption         =   "导数比较"
   ClientHeight    =   7590
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   12640
   OleObjectBlob   =   "导数比较窗体.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "导数比较窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrCmd(1 To 2) As CMySelCmdb
Private selectIndex As Boolean


Private Sub AllSelectCheck_Click()
    Dim mycnt As Control, ss As String
    
    If selectIndex = True Then
        selectIndex = False
    Else
        selectIndex = True
    End If
    
    With Me
        For Each mycnt In .CheckFrame.Controls
            If TypeName(mycnt) = "CheckBox" Then
                mycnt.Value = selectIndex
                '如果是真，则更新
                If selectIndex = True Then
                    Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 1)
                Else
                    Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 0)
                End If
            End If
        Next
    End With
End Sub


'
'   设置是否启用下面的单选框
'
Private Sub CB_TotalCheck_Click()
    Call ChangeTotalCheckStatus
End Sub

'
'   启用总阀值选项
'
Private Sub ChangeTotalCheckStatus()
    Me.SB_Sel1.Enabled = Not Me.SB_Sel1.Enabled
    Me.SB_Sel2.Enabled = Not Me.SB_Sel2.Enabled
    Me.TextBox1.Enabled = Not Me.TextBox1.Enabled
    Me.TextBox2.Enabled = Not Me.TextBox2.Enabled
    Me.TextBox3.Enabled = Not Me.TextBox3.Enabled
End Sub

Private Sub Init_jzd_CheckStatus()
    Me.CB_集中度.Value = CBool(val(gloCft.GetConfig("系统导数比较-" & Me.CB_集中度.name & "-历史选项")))
    
    If Me.CB_集中度.Value = True Then
        Me.OB_集中度_或模式.Enabled = True
        Me.OB_集中度_与模式.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.TextBox6.Enabled = True
        Me.TextBox7.Enabled = True
    Else
        Me.OB_集中度_或模式.Enabled = False
        Me.OB_集中度_与模式.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox7.Enabled = False
    End If
    
    Me.OB_集中度_或模式.Value = CBool(ycV(gloCft.GetConfig("系统导数比较-" & Me.OB_集中度_或模式.name & "-历史选项")))
    Me.OB_集中度_与模式.Value = CBool(ycV(gloCft.GetConfig("系统导数比较-" & Me.OB_集中度_与模式.name & "-历史选项")))
    
    Me.TextBox4.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox4.name & "-历史选项"))
    Me.TextBox5.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox5.name & "-历史选项"))
    Me.TextBox6.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox6.name & "-历史选项"))
    Me.TextBox7.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox7.name & "-历史选项"))
    
    If CStr(Me.TextBox4.Value) = "" Then
        Me.TextBox4.Value = 20000
        Call gloCft.SetConfig("系统导数比较-" & Me.TextBox4.name & "-历史选项", 20000)
    End If
    If CStr(Me.TextBox5.Value) = "" Then
        Me.TextBox5.Value = 30
        Call gloCft.SetConfig("系统导数比较-" & Me.TextBox5.name & "-历史选项", 30)
    End If
    If CStr(Me.TextBox6.Value) = "" Then
        Me.TextBox6.Value = 100
        Call gloCft.SetConfig("系统导数比较-" & Me.TextBox6.name & "-历史选项", 100)
    End If
    If CStr(Me.TextBox7.Value) = "" Then
        Me.TextBox6.Value = 50
        Call gloCft.SetConfig("系统导数比较-" & Me.TextBox6.name & "-历史选项", 50)
    End If
End Sub

'
'   总阈值和集中度选项初始化
'
Private Sub InitTotalCheckStatus()
    Me.CB_TotalCheck.Value = CBool(val(gloCft.GetConfig("系统导数比较-" & Me.CB_TotalCheck.name & "-历史选项")))
    
    If Me.CB_TotalCheck.Value = True Then
        Me.SB_Sel1.Enabled = True
        Me.SB_Sel2.Enabled = True
        Me.TextBox1.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
    Else
        Me.SB_Sel1.Enabled = False
        Me.SB_Sel2.Enabled = False
        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
    End If
    
    Me.SB_Sel1.Value = CBool(ycV(gloCft.GetConfig("系统导数比较-" & Me.SB_Sel1.name & "-历史选项")))
    Me.SB_Sel2.Value = CBool(ycV(gloCft.GetConfig("系统导数比较-" & Me.SB_Sel2.name & "-历史选项")))
    
    Me.TextBox1.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox1.name & "-历史选项"))
    Me.TextBox2.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox2.name & "-历史选项"))
    Me.TextBox3.Value = ycV(gloCft.GetConfig("系统导数比较-" & Me.TextBox3.name & "-历史选项"))
End Sub

Private Sub Save_jzd_CheckStatus()
    If Me.CB_集中度.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.CB_集中度.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.CB_集中度.name & "-历史选项", 0)
    End If
    
    '保存其他选项
    If Me.OB_集中度_或模式.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.OB_集中度_或模式.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.OB_集中度_或模式.name & "-历史选项", 0)
    End If
    
    If Me.OB_集中度_与模式.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.OB_集中度_与模式.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.OB_集中度_与模式.name & "-历史选项", 0)
    End If
    
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox4.name & "-历史选项", CStr(Me.TextBox4.Value))
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox5.name & "-历史选项", CStr(Me.TextBox5.Value))
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox6.name & "-历史选项", CStr(Me.TextBox6.Value))
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox7.name & "-历史选项", CStr(Me.TextBox7.Value))
End Sub

Private Sub SaveTotalCheckStatus()
    If Me.CB_TotalCheck.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.CB_TotalCheck.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.CB_TotalCheck.name & "-历史选项", 0)
    End If
    
    '保存其他选项
    If Me.SB_Sel1.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.SB_Sel1.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.SB_Sel1.name & "-历史选项", 0)
    End If
    
    If Me.SB_Sel2.Value = True Then
        Call gloCft.SetConfig("系统导数比较-" & Me.SB_Sel2.name & "-历史选项", 1)
    Else
        Call gloCft.SetConfig("系统导数比较-" & Me.SB_Sel2.name & "-历史选项", 0)
    End If
    
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox1.name & "-历史选项", CStr(Me.TextBox1.Value))
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox2.name & "-历史选项", CStr(Me.TextBox2.Value))
    Call gloCft.SetConfig("系统导数比较-" & Me.TextBox3.name & "-历史选项", CStr(Me.TextBox3.Value))
End Sub



Private Sub CB_集中度_Click()
    Call Change_jzd_CheckStatus
End Sub

'
'   调整集中度选项
'
Private Sub Change_jzd_CheckStatus()
    Me.OB_集中度_或模式.Enabled = Not Me.OB_集中度_或模式.Enabled
    Me.OB_集中度_与模式.Enabled = Not Me.OB_集中度_与模式.Enabled
    Me.TextBox4.Enabled = Not Me.TextBox4.Enabled
    Me.TextBox5.Enabled = Not Me.TextBox5.Enabled
    Me.TextBox6.Enabled = Not Me.TextBox6.Enabled
    Me.TextBox7.Enabled = Not Me.TextBox7.Enabled
End Sub

Private Sub CommandButton1_Click()
    If CheckFaZhi = False Then Exit Sub
    '执行选项存储
    Dim mycnt As Control
    For Each mycnt In Me.CheckFrame.Controls
        If TypeName(mycnt) = "CheckBox" Then
            If mycnt.Value = True Then
                Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 1)
            Else
                Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 0)
            End If
        End If
    Next
    
    
    Call SaveTotalCheckStatus
    Call Save_jzd_CheckStatus
    Dim savePath As String, r As Variant
    savePath = gloFt.GetOutputFile
    耗时标签.Visible = True
    
    If Me.tb1.Value = "" Or savePath = "" Then
        MsgBox "漏选文件"
    Else
        Call 跨期数值核对(savePath)
        耗时标签.Visible = False
        Unload Me
    End If
End Sub

Private Sub Userform_Initialize()
    Dim units As Variant, u As Variant
    Dim cmdt As CMySelCmdb
    
    耗时标签.Visible = False
    
    TB_分母.Value = "00434400000"
    TB_分母.Enabled = True
    
    Dim i As Integer
    For i = 1 To 2 Step 1
        Set cmdt = New CMySelCmdb
        Set cmdt.cmd = Me.Controls("cmd" & CStr(i))
        Set cmdt.frmT = Me
        Set arrCmd(i) = cmdt
    Next i
    
    Set cmdt = Nothing
    源数据单位.Enabled = True
    units = Array("元", "十元", "百元", "千元", "万元", "十万元", "百万元", "千万元", "亿元")
    For Each u In units
        源数据单位.AddItem u
        目标单位.AddItem u
        说明文件单位.AddItem u
    Next
        源数据单位 = gloCft.GetConfig("源数据单位历史选项")
        目标单位 = gloCft.GetConfig("目标单位历史选项")
        说明文件单位 = gloCft.GetConfig("说明文件单位历史选项")
    '遍历
    selectIndex = True
    
    If glo_iOrgMode = 2 Then
        源数据单位.Value = "元"
        源数据单位.Enabled = False
    End If
    
    Dim x As Integer, w2 As Integer, h As Integer, fontSize As Integer, ws As Worksheet, vNum As Integer
    h = 25
    fontSize = 10
    x = 10
    w2 = (导数比较窗体.CheckFrame.width - 20) / 2 + 10
    '生成复选框，设置勾选状态，以及布局
    i = 0
    vNum = 0
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.name, "特殊指标") Or InStr(ws.name, "复杂校验") > 0 Then
            With 导数比较窗体.CheckFrame.Controls.add(bstrprogid:="forms.CheckBox.1", name:="校验表" & i)
                .Caption = CStr(i + 1) & " " & ws.name
                .name = ws.name
                .width = w2 - 20
                .Height = h
                .Font.Size = fontSize
                .Value = CBool(val(gloCft.GetConfig("系统导数比较-" & ws.name & "-历史选项")))
                If err.Number <> 0 Then
                    .Value = True
                    err.Clear
                End If
                If vNum Mod 2 = 0 Then
                    .left = 20
                    .Top = x
                Else
                    .left = w2
                    .Top = x
                    x = x + h
                End If
                i = i + 1
                If InStr(ws.name, "自定义") > 0 Then
                    vNum = vNum + 1
                Else
                    .Visible = False
                End If
            End With
        End If
    Next
    
    Call InitTotalCheckStatus
    Call Init_jzd_CheckStatus
    导数比较窗体.CheckFrame.ScrollHeight = x + h
End Sub

Private Function Check() As Boolean
    
End Function

Private Sub 目标单位_Change()
    Call gloCft.SetConfig("目标单位历史选项", CStr(导数比较窗体.目标单位))  '源数据单位
End Sub

Private Sub 说明文件单位_Change()
    Call gloCft.SetConfig("说明文件单位历史选项", CStr(导数比较窗体.说明文件单位))
End Sub

Private Sub 源数据单位_Change()
     Call gloCft.SetConfig("源数据单位历史选项", CStr(导数比较窗体.源数据单位))  '源数据单位
End Sub

'
'   校验伐值框是否正确
'
Private Function CheckFaZhi() As Boolean
    Dim t1 As Variant, t2 As Variant, t3 As Variant
    On Error GoTo errHandle
    If Me.CB_TotalCheck = True Then
        
        t1 = Me.TextBox1.Value
        t2 = Me.TextBox2.Value
        t3 = Me.TextBox3.Value
        
        
        t1 = CDbl(t1)
        t2 = CDbl(t2)
        t3 = CDbl(t3)
    
    End If
    
    If Me.CB_集中度 = True Then
        t1 = Me.TextBox4.Value
        t2 = Me.TextBox5.Value
        t3 = Me.TextBox6.Value
        
        t1 = CDbl(t1)
        t2 = CDbl(t2)
        t3 = CDbl(t3)
    
    End If
    CheckFaZhi = True
Exit Function

errHandle:
    CheckFaZhi = False
    MsgBox "阈值框应该为数值或为空"
End Function

Private Sub 执行比较_Click()
    If CheckFaZhi = False Then Exit Sub
    '执行选项存储
    Dim mycnt As Control
    For Each mycnt In Me.CheckFrame.Controls
        If TypeName(mycnt) = "CheckBox" Then
            If mycnt.Value = True Then
                Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 1)
            Else
                Call gloCft.SetConfig("系统导数比较-" & mycnt.name & "-历史选项", 0)
            End If
        End If
    Next
    
    
    Call SaveTotalCheckStatus
    Call Save_jzd_CheckStatus
    Dim savePath As String, r As Variant
    savePath = gloFt.GetOutputFile
    耗时标签.Visible = True
    
    If Me.tb1.Value = "" Or Me.tb2.Value = "" Or savePath = "" Then
        MsgBox "漏选文件"
    Else
        Call 导数比较控制台2(savePath)
        耗时标签.Visible = False
        Unload Me
    End If
End Sub
