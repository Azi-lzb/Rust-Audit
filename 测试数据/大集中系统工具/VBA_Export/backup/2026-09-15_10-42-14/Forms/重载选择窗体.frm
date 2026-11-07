VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 重载选择窗体 
   Caption         =   "选择重载机构"
   ClientHeight    =   5025
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11070
   OleObjectBlob   =   "重载选择窗体.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "重载选择窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private selectIndex As Boolean

Private Sub Userform_Initialize()
    selectIndex = True
    
    Dim x As Integer, w2 As Integer, h As Integer, fontSize As Integer
    h = 25
    fontSize = 10
    x = 10
    w2 = (重载选择窗体.ReloadCheckboxFrame.width - 20) / 2 + 10
    '生成复选框，设置勾选状态，以及布局
    Dim el As Variant, i As Integer
    i = 0
    For Each el In dOrgs.items
        With 重载选择窗体.ReloadCheckboxFrame.Controls.add(bstrprogid:="forms.CheckBox.1", name:="选择" & i)
            .Caption = CStr(i + 1) & " " & el.clsOrgName & " " & el.clsAreaName
            .name = el.clsOrgCode & el.clsAreaCode
            .width = w2 - 20
            .Height = h
            .Font.Size = fontSize
            .Value = True
            If i Mod 2 = 0 Then
                .left = 20
                .Top = x
            Else
                .left = w2
                .Top = x
                x = x + h
            End If
        End With
        i = i + 1
    Next
    
    重载选择窗体.ReloadCheckboxFrame.ScrollHeight = x + h
End Sub



Private Sub AllSelBtn_Click()
    Dim mycnt As Control, ss As String
    
    If selectIndex = True Then
        selectIndex = False
    Else
        selectIndex = True
    End If
    
    With Me
        For Each mycnt In .Controls
            If TypeName(mycnt) = "CheckBox" Then
                mycnt.Value = selectIndex
            End If
        Next
    End With
End Sub

Private Sub DoSelectBtn_Click()
    Dim mycnt As Control, ss As String
    With Me
        For Each mycnt In .Controls
            If TypeName(mycnt) = "CheckBox" Then
                If mycnt.Value = False Then
                    '剔除掉字典中的值
                    dOrgs.Remove (mycnt.name)
                End If
            End If
        Next
  End With
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'CloseMode参数表明事件发生的原因
    If CloseMode = vbFormControlMenu Then
        'MsgBox "你准备退出当前窗口了！"
        '这里写点了关闭之后的代码     写Cancel = True 代表不允许关闭！！
        dOrgs.RemoveAll
        Set dOrgs = Nothing
    End If
End Sub
