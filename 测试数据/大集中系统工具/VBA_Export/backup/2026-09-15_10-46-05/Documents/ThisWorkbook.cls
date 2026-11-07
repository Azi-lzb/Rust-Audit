Option Explicit

Private Sub Workbook_Deactivate()
    Application.CommandBars("Worksheet menu bar").Reset
End Sub

Private Sub Workbook_Open()
    Call 修改右键菜单
End Sub

Private Sub 修改右键菜单()
    Dim ToolMenu As Object
    Dim NewMenu As Object
    With Application.CommandBars("Worksheet menu bar")
        .Reset
        '利用ID属性查找帮助菜单
        Set ToolMenu = .FindControl(id:=.Controls("工具(&T)").id)
        If ToolMenu Is Nothing Then
            '如果该菜单不存在,则将新菜单添加到末尾
            Set NewMenu = .Controls.add(Type:=msoControlPopup)
        Else
            '将新菜单添加到帮助菜单之前
            Set NewMenu = .Controls.add(Type:=msoControlPopup, _
                before:=ToolMenu.index)
        End If
        With NewMenu
            .Caption = "大集中工具(&S)"
            '添加第一个菜单项
            With .Controls.add(Type:=msoControlButton)
                .Caption = "导数比较-指标查询版"
                .FaceId = 88
                .OnAction = "ShowDoCompare"
            End With
            With .Controls.add(Type:=msoControlButton)
                .Caption = "导数比较-IJ文件版"
                .OnAction = "ShowDoCompare2"
            End With
            With .Controls.add(Type:=msoControlButton)
                .Caption = "系统导数转EXCEL"
                .FaceId = 88
                .OnAction = "ShowDoReload"
            End With
            With .Controls.add(Type:=msoControlButton)
                .Caption = "台账拆分"
                .OnAction = "SplitShts_jz"
            End With
            With .Controls.add(Type:=msoControlButton)
                .Caption = "全部校验更新"
                .OnAction = "refresh_check"
            End With
        End With
    End With
    Set ToolMenu = Nothing
    Set NewMenu = Nothing
End Sub

Private Sub Workbook_Activate()
    Call 修改右键菜单
End Sub


