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
            .Caption = "报表工具(&S)"
            '添加第1个菜单项
            With .Controls.add(Type:=msoControlPopup)
                .Caption = "1. 文件选择"
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.1 选择模板文件"
                    .OnAction = "SelHzModFile"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.1_2 选择模板文件（模板列表）"
                    .OnAction = "SelHzModFile_form_list"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.1_3 模板装载"
                    .OnAction = "LoadModFilesSelect"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.2 选择外部文件"
                    .OnAction = "SelExtendFile"
                End With
                

                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.3 选择源文件"
                    .OnAction = "SelHzSrcFiles"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.4 批量选择源文件"
                    .OnAction = "SelHzSrcFilesInFolder"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "1.5 批量选择源文件(含文件夹)"
                    .OnAction = "SelHzSrcFilesInFolder2"
                End With

            End With
            '添加第2个菜单项
            With .Controls.add(Type:=msoControlPopup)
                .Caption = "2. 数据预处理"

                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.1 源数据校验"
                    .OnAction = "CheckFileAvail"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.2 源数据校验(反向)"
                    .OnAction = "CheckFileAvail_Rev"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.3 导出校验日记"
                    .OnAction = "export_log_to_file"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.4 清除模板数据"
                    .OnAction = "pl_clear_mod_data"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.5 清除源数据校验区域"
                    .OnAction = "ClearSrcFilesCheckArea"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.6 数据校验说明接收"
                    .OnAction = "Accept_CheckFileAvail"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.7 提取设置字段"
                    .OnAction = "PickFileSetting"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.8 反向设置字段"
                    .OnAction = "PickSrvFileSetting"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.9 执行文件重命名"
                    .OnAction = "DoRename"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.10 批量转为xls"
                    .OnAction = "ChangeToXls"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.11 批量转为xlsx"
                    .OnAction = "ChangeToXlsx"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.12 批量清除宏"
                    .OnAction = "DoClearMacroCode"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "2.13 批量更新上期数"
                    .OnAction = "RePreData"
                End With
            End With
            
            '添加第3个菜单项
            With .Controls.add(Type:=msoControlPopup)
                .Caption = "3. 汇总数据(&T)"
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "3.1 简单汇总"
                    .OnAction = "EasyHz"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "3.2 源数据汇总(全行)"
                    .OnAction = "EasyStandHz"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "3.3 源数据汇总(单指标)"
                    .OnAction = "StandHz"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "3.4 台账汇总"
                    .OnAction = "RecordHz"
                End With
            End With
            '添加第4个菜单项
            With .Controls.add(Type:=msoControlPopup)
                .Caption = "4. 其余功能"
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "4.1 工作簿归集"
                    .OnAction = "GuijiWb"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "4.2 工作簿拆分"
                    .OnAction = "ChaifenWb"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "4.3 台账拆分"
                    .OnAction = "SplitShts2"
                End With
            End With
            
            '添加第5个菜单项
            With .Controls.add(Type:=msoControlPopup)
                .Caption = "5. 基础数据处理"
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "5.1 核查表据校验"
                    .OnAction = "CheckFileAvail_jcsj"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "5.2 逻辑校验结果汇总"
                    .OnAction = "RecordHz_jcsj_check_hz"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "5.3 校验比对情况汇总"
                    .OnAction = "RecordHz_jcsj_check_cmp"
                End With
                With .Controls.add(Type:=msoControlButton)
                    .Caption = "5.4 报表跨期比较-单sheet导出"
                    .OnAction = "基础数据系统报表跨期比较过程"
                End With
                
            End With
        End With
    End With
    Set ToolMenu = Nothing
    Set NewMenu = Nothing
End Sub

Private Sub Workbook_Activate()
    Call 修改右键菜单
End Sub
