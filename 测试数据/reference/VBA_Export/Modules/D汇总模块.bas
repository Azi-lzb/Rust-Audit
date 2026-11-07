Attribute VB_Name = "D汇总模块"
Option Explicit

Private clsmymenu As CMySysMenu
Private bErrAvailable As Boolean   '用于存储工作簿是否正确
Private srcfiles As Variant '源数据二维数组
Private dsrcfiles As Object '源数据字典
Private Const WsHzSys As String = "执行面板" '执行面板
Private glodHzMod As Object   '存储汇总模板
Private gloHzModFileName As String  '当前汇总模板文件
Private dHzArr As Object    '存储汇总数据
Private dHzArrCount As Object   '存储汇总数据当前行数
Private StandHzCurRow As Long
Private dCheckRow As Object '用于判断是否要新增行列
Private savePath As String  '存储路径
Private dHzSysCol As Object '用于存储列名称与列的对应
Private dStandOrg As Object '当前标准化的机构

Private sfqxtg As Integer   '强制校验同通过标志
Private qzhz As Integer     '强制汇总标志
Private hzqjy As Integer    '汇总去校验标志
Private jyqgsh As Integer   '校验去格式化标志
Private log_des As Integer  '日志文件的输出标志 0 是“输入日记“，1 是“输出到外部
Private qclz As Integer '去除0值标志
Private jy_link As Integer ' 校验结果是否增加超链接
Private jcsh_job_filter As Integer
Private is_check_data As Integer   '是否检查数据区域
Private is_get_all_col_data As Integer '是否提取数据区域中所有列的数据
Private bnfEasyHz As Boolean '标记是否不是第一张表
Private is_rev_export_file As Boolean  '标记反向校验是否另存文件
Private glo_str_relust As String

Private files_sc As String
Private d_ce_mod As Object  ' 老系统核查表
Private d_sc_mod As Object  ' 新系统测试核查表
Private d_predata As Object

Private is_jcsj_check As Boolean



Private Function clear_job_name(job_name As String) As String
    Dim r As String
    r = replace(job_name, "金融基础数据采集（", "")
    r = replace(r, "）", "")
    r = replace(r, "金融基础数据-", "")
    r = replace(r, "五篇大文章-", "")
    clear_job_name = r
End Function
    

'
' 基础数据 逻辑校验结果汇总及比较，还需要增加本期、历史期数的核对
'
Public Sub RecordHz_jcsj_check_hz()
    Call gloCft.CloseSomething
    
    'On Error GoTo errHandle
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles = False Then Exit Sub
    
    Dim fnOut As String
    fnOut = gloFt.GetOutputFile
    If fnOut = "" Then
        Exit Sub
    End If
    
    '判断是否有列指标
    Dim wsdes As Worksheet, wbDes As Workbook, arrSplit As Variant, arrSplit2 As Variant, hzNum As Long, bsplit As Boolean, i As Long, j As Integer
    Dim wbsrc As Workbook, wsSrc As Worksheet, curRow As Long, arrSrc As Variant, srcItem As Variant, z As Long, short_name As String, check_type As String
    Dim d_data As Object, d_data_temp As Object, d_temp As Object, arr_info As Variant, hzi As Long, d_data_cy As Object, d_data_history As Object
    Dim y As Long, key As String, d_job As Object
    
    Set d_data = Nothing
    Set d_data_temp = Nothing
    Set d_temp = Nothing
    Set d_data_cy = Nothing
    Set d_data_history = Nothing
    Set d_job = Nothing
    
    Set d_job = CreateObject("scripting.dictionary")
    Set d_data = CreateObject("scripting.dictionary")
    Set d_data_cy = CreateObject("scripting.dictionary")
    Set d_data_history = CreateObject("scripting.dictionary")
    
    Set wbDes = gloEt.GetWb(fnOut)
    If wbDes Is Nothing Then
        Set wbDes = Workbooks.add
        wbDes.SaveAs fnOut
    End If
    
    
    hzNum = 0
    hzi = 1
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        ' 判断是否 校验对比信息
        short_name = gloFt.getShortFileName(CStr(srcfiles(i, 2)))
        check_type = ""
        If InStr(short_name, "校验结果及报送说明") > 0 Then
            check_type = "本地"
            If InStr(short_name, "_校验结果及报送说明_人行侧") > 0 Then
                check_type = ""
            End If
        Else
            check_type = ""
        End If
        arr_info = Split(short_name, "_")  ' 获取了基本信息
        
        
        
        ' 判断是否是指定类型的汇总
        If check_type <> "" And UBound(arr_info) >= 2 Then
            d_job(clear_job_name(CStr(arr_info(0)) & CStr(arr_info(1)))) = 1
            Set wbsrc = gloEt.GetWb(CStr(srcfiles(i, 2)))
            If wbsrc Is Nothing Then
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F,文件不存在")
                'Call gloDe.InsertNewDebug("台账汇总", CStr(srcfiles(i, 2)), "-", "不参与汇总", "文件不存在。")
                'Call gloDe.AddColorDebug(15)
                GoTo NextHzSrcFiles
            End If
            
            For Each wsSrc In wbsrc.Worksheets
                If wsSrc.name = "任务说明" Or wsSrc.name = "业务说明" Or wsSrc.name = "本地校验结果" Then
                    If InStr(wsSrc.Range("B2").value, "id") > 0 Or InStr(wsSrc.Range("B2").value, "批次") > 0 Then
                        wsSrc.Range("B2").EntireColumn.Delete
                    End If
                    arrSrc = gloEt.GetUsedRange(wsSrc).value
                    
                    If UBound(arrSrc) > 2 Then
                        For y = 3 To UBound(arrSrc)
                            Set d_temp = CreateObject("scripting.dictionary")
                            
                            If wsSrc.name <> "任务说明" Then
                                d_temp("机构名称") = CStr(arr_info(0))
                                d_temp("任务名称") = clear_job_name(CStr(arr_info(1)))
                                d_temp("数据日期") = replace(CStr(arr_info(2)), "校验结果及报送说明", "")
                            End If
                            
                            
                            For j = 1 To UBound(arrSrc, 2)
                                d_temp(CStr(arrSrc(2, j))) = CStr(arrSrc(y, j))
                            Next
                            
                            
                            If Not d_data.exists(wsSrc.name) Then
                                Set d_data(wsSrc.name) = CreateObject("scripting.dictionary")
                            End If
                            
                            If wsSrc.name = "本地校验结果" Then
                                key = d_temp("机构名称") & "|" & d_temp("任务名称") & "|" & d_temp("规则编号") & "|" & d_temp("校验字段")
                                If d_data(wsSrc.name).exists(key) Then
                                    d_data(wsSrc.name)(key)("分类触发条数") = ycV(d_data(wsSrc.name)(key)("分类触发条数")) + ycV(d_temp("分类触发条数"))
                                Else
                                    Set d_data(wsSrc.name)(key) = d_temp
                                End If
                                
                            Else
                                Set d_data(wsSrc.name)(hzi) = d_temp
                            End If
 
                            
                            hzi = hzi + 1 ' 有多个文件，索引会重复使用
                            Set d_temp = Nothing
                        Next
                    End If
                ElseIf wsSrc.name = "在线抽验结果" Then
                    If InStr(wsSrc.Range("B2").value, "id") > 0 Or InStr(wsSrc.Range("B2").value, "批次") > 0 Then
                        wsSrc.Range("B2").EntireColumn.Delete
                    End If
                    arrSrc = gloEt.GetUsedRange(wsSrc).value
                    If UBound(arrSrc) > 2 Then
                        For y = 3 To UBound(arrSrc)
                            Set d_temp = CreateObject("scripting.dictionary")
                            
                            d_temp("机构名称") = CStr(arr_info(0))
                            d_temp("任务名称") = clear_job_name(CStr(arr_info(1)))
                            d_temp("数据日期") = replace(CStr(arr_info(2)), "校验结果及报送说明", "")
                            
                            For j = 1 To UBound(arrSrc, 2)
                                d_temp(CStr(arrSrc(2, j))) = CStr(arrSrc(y, j))
                            Next
                            
                            If Not d_temp.exists("说明") Then
                                d_temp("说明") = "在线抽验，无说明字段。"
                            End If
                            
                            key = d_temp("机构名称") & "|" & d_temp("任务名称") & "|" & d_temp("规则编号") & "|" & d_temp("校验字段")
                            Set d_data_cy(key) = d_temp
                            Set d_temp = Nothing
                        Next
                    End If
                    
                End If
            Next
                
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")

            hzNum = hzNum + 1
    
            wbsrc.Close False
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F,非校验对比信息文件或文件命名不规范")
        End If
        
NextHzSrcFiles:
        Call gloCft.updateScreen_msg(index:=i, step:=5, msg:="")
    Next
    
    Dim arr_temp As Variant, k As Variant, v As Variant, temp_k As Variant, temp_v As Variant
    Dim l_lscz As Long, l_lswtg As Long, l_cycz As Long, l_cyts As Long, l_cysm As Long
    l_lscz = 0: l_lswtg = 0: l_cycz = 0: l_cyts = 0: l_cysm = 0
    
    k = d_data.keys
    v = d_data.items
    
    Do While wbDes.Worksheets.count > 1
        wbDes.Worksheets(1).Delete
    Loop
    
    ' 读取历史校验参照表
    arr_temp = gloEt.GetUsedRange(gloEt.FindWs(ThisWorkbook, "历史校验结果", 0, True)).value
    For i = 2 To UBound(arr_temp)
        If d_job.exists(CStr(arr_temp(i, 1)) & clear_job_name(CStr(arr_temp(i, 2)))) Then
            Set d_temp = CreateObject("scripting.dictionary")
            For j = 1 To UBound(arr_temp, 2)
                d_temp(CStr(arr_temp(1, j))) = CStr(arr_temp(i, j))
            Next
            d_temp("任务名称") = clear_job_name(CStr(d_temp("任务名称")))
            key = d_temp("机构名称") & "|" & d_temp("任务名称") & "|" & d_temp("规则编号") & "|" & d_temp("校验字段")
            Set d_data_history(key) = d_temp
            Set d_temp = Nothing
        End If
    Next
    
    For i = 0 To UBound(v)
        Set d_data_temp = v(i)
        If v(i).count > 0 Then
            ' 再调整一下
            
            ' 如果是本地校验，那就跟历史、在线结果比较以下
            If CStr(k(i)) = "本地校验结果" Then
                temp_k = d_data_cy.keys
                temp_v = d_data_cy.items
                
                For y = 0 To UBound(temp_k)
                    key = CStr(temp_k(y))
                    If d_data_temp.exists(key) Then
                        Set d_temp = d_data_temp(key)
                        d_temp("在线抽验条数") = temp_v(y)("触发条数")
                        d_temp("抽验条数是否一致") = "是"
                        d_temp("抽验说明") = temp_v(y)("说明")
                        d_temp("抽验说明是否一致") = "是"
                        If CLng(ycV(d_temp("分类触发条数")) - ycV(temp_v(y)("触发条数"))) <> 0 Then
                            d_temp("抽验条数是否一致") = "否"
                            l_cyts = l_cyts + 1
                        End If
                        If Trim(CStr(d_temp("分类说明"))) <> Trim(CStr(temp_v(y)("说明"))) Then
                            d_temp("抽验说明是否一致") = "否"
                            l_cysm = l_cysm + 1
                        End If
                        Set d_data_temp(key) = d_temp
                    Else
                        Set d_temp = temp_v(y)
                        d_temp("分类触发条数") = 0
                        d_temp("分类说明") = ""
                        d_temp("示例数据") = CStr(temp_v(y)("示例数据行号"))
                        d_temp("在线抽验条数") = temp_v(y)("触发条数")
                        d_temp("抽验条数是否一致") = "否"
                        d_temp("抽验说明") = temp_v(y)("说明")
                        d_temp("抽验说明是否一致") = "否"
                        Set d_data_temp(key) = d_temp
                        l_cycz = l_cycz + 1
                    End If
                    
                    Call gloCft.updateScreen_msg(index:=y, step:=2000, msg:="")
                Next
                ' 历史数据说明 历史是否通过    历史触发条数    历史判断
                temp_k = d_data_history.keys
                temp_v = d_data_history.items
                For y = 0 To UBound(temp_k)
                    key = CStr(temp_k(y))
                    If d_data_temp.exists(key) Then
                        Set d_temp = d_data_temp(key)
                        d_temp("历史是否通过") = temp_v(y)("历史是否通过")
                        If d_temp("历史是否通过") = "是" Then
                        
                        Else
                            l_lswtg = l_lswtg + 1
                        End If
                        d_temp("历史触发条数") = temp_v(y)("历史触发条数")
                        d_temp("历史分类说明") = temp_v(y)("历史分类说明")
                        d_temp("历史判断") = temp_v(y)("历史判断")
                        d_temp("是否历史出现但本期未出现") = "否"
                        
                        Set d_data_temp(key) = d_temp
                    Else
                        Set d_temp = temp_v(y)
                        d_temp("分类触发条数") = 0
                        d_temp("分类说明") = ""
                        d_temp("示例数据") = temp_v(y)("历史示例数据")
                        d_temp("是否历史出现但本期未出现") = "是"
                        Set d_data_temp(key) = d_temp
                        l_lscz = l_lscz + 1
                    End If
                Next
            End If
                
            arr_temp = dictsTolists(Split(gloCft.GetConfig(CStr(k(i))), ","), d_data_temp, True)
            
            If i = 0 And wbDes.Worksheets.count = 1 Then
                Set wsdes = wbDes.Worksheets(1)
                wsdes.UsedRange.Clear
            Else
                ' 输出到工作表
                Set wsdes = wbDes.Worksheets.add
            End If
             
            Call gloEt.CancelFilter(wsdes)
            wsdes.name = CStr(k(i))
            wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).NumberFormatLocal = "@"
            'wsdes.Range("F:K").NumberFormatLocal = "0"
            wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).value = arr_temp
            wsdes.Cells.WrapText = False
            
            Erase arr_temp
            Set d_data_temp = Nothing
        End If
    Next
    
    Set d_data = Nothing
    
    Set d_data_cy = Nothing
    Set d_data_history = Nothing
    key = "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总。"
    If l_lscz + l_lswtg + l_cycz + l_cyts + l_cysm > 0 Then key = key & "其中:"
    If l_lscz > 0 Then key = key & Chr(13) & "有" & CStr(l_lscz) & "条校验，历史存在，本期无，是否存在漏校验。"
    If l_lswtg > 0 Then key = key & Chr(13) & "有" & CStr(l_lswtg) & "条校验，历史判断为未通过，请核实。"
    If l_cycz > 0 Then key = key & Chr(13) & "有" & CStr(l_lscz) & "条校验，抽验存在，本地校验无，需核对。"
    If l_cyts > 0 Then key = key & Chr(13) & "有" & CStr(l_cyts) & "条校验，抽验条数与在线条数不一致，请核实。"
    If l_cysm > 0 Then key = key & Chr(13) & "有" & CStr(l_cysm) & "条校验，抽验说明与在线不一致，请核实。"

    MsgBox key
    Call clsmymenu.RecoverSysMenu(srcfiles)
    
    wbDes.Windows(1).FreezePanes = False
    wbDes.Windows(1).SplitRow = 1
    wbDes.Windows(1).FreezePanes = True
    wbDes.Save
    Call EraseSomething
    Call gloCft.OpenSomething
Exit Sub
errhandle:
    MsgBox "逻辑校验信息汇总失败 -- " & err.Description
    err.Clear
    Set d_data = Nothing
    Set d_data_temp = Nothing
    Set d_data_cy = Nothing
    Set d_data_history = Nothing
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub


'
'  按指定格式转换成数组
'
Private Function dictsTolists(titles As Variant, ByRef d_data As Object, Optional is_dict_title As Boolean = False) As Variant
    On Error Resume Next
    Dim c As Long, arr As Variant, i As Long, j As Long, v As Variant, v2 As Variant, d_title As Object, k2 As Variant
    c = 1 + d_data.count
    
    ReDim arr(1 To c, 1 To UBound(titles) + 1) As Variant
    Set d_title = CreateObject("scripting.dictionary")
    ' 标题
    For j = 0 To UBound(titles)
        arr(1, j + 1) = CStr(titles(j))
        d_title(CStr(titles(j))) = j + 1
    Next
    
    If is_dict_title = False Then
        v = d_data.items
        For i = 0 To UBound(v)
            v2 = v(i).items
            For j = 0 To UBound(v2)
                arr(i + 2, j + 1) = CStr(v2(j))
            Next
        Next
    Else
        v = d_data.items
        For i = 0 To UBound(v)
            v2 = v(i).items
            k2 = v(i).keys
            For j = 0 To UBound(k2)
                If d_title.exists(CStr(k2(j))) Then
                    arr(i + 2, d_title(CStr(k2(j)))) = CStr(v2(j))
                End If
            Next
        Next
    End If
    
    dictsTolists = arr
End Function
    

'
' 基础数据 校验对比情况汇总，有离线校验对比信息 在线校验对比信息
'
Public Sub RecordHz_jcsj_check_cmp()
    Call gloCft.CloseSomething
    
    On Error GoTo errhandle
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles = False Then Exit Sub
    
    Dim fnOut As String
    fnOut = gloFt.GetOutputFile
    If fnOut = "" Then
        Exit Sub
    End If
    
    '判断是否有列指标
    Dim wsdes As Worksheet, wbDes As Workbook, arrSplit As Variant, arrSplit2 As Variant, hzNum As Long, bsplit As Boolean, i As Long, j As Integer
    Dim wbsrc As Workbook, wsSrc As Worksheet, curRow As Long, arrSrc As Variant, srcItem As Variant, z As Long, short_name As String, check_type As String
    Dim d_data_lx As Object, d_data_zx As Object, d_temp As Object, arr_info As Variant, hzi As Long
    Dim y As Long
    
    Set d_data_lx = Nothing
    Set d_data_zx = Nothing
    Set d_temp = Nothing
    
    Set d_data_lx = CreateObject("scripting.dictionary")
    Set d_data_zx = CreateObject("scripting.dictionary")
    
    Set wbDes = gloEt.GetWb(fnOut)
    If wbDes Is Nothing Then
        Set wbDes = Workbooks.add
        wbDes.SaveAs fnOut
    End If
    
    
    hzNum = 0
    hzi = 1
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        ' 判断是否 校验对比信息
        short_name = gloFt.getShortFileName(CStr(srcfiles(i, 2)))
        check_type = ""
        If InStr(short_name, "_在线校验比对信息") > 0 Then
            check_type = "在线校验比对信息"
        ElseIf InStr(short_name, "_离线校验比对信息") > 0 Then
            check_type = "离线校验比对信息"
        Else
            check_type = ""
        End If
        arr_info = Split(short_name, "_")  ' 获取了基本信息
        
        ' 判断是否是指定类型的汇总
        If check_type <> "" And UBound(arr_info) < 5 Then
            Set wbsrc = gloEt.GetWb(CStr(srcfiles(i, 2)))
            If wbsrc Is Nothing Then
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F,文件不存在")
                GoTo NextHzSrcFiles
            End If
            
            For Each wsSrc In wbsrc.Worksheets
                If wsSrc.name = "离线校验比对信息" Or wsSrc.name = "校验结果比对信息" Then
                    arrSrc = gloEt.GetUsedRange(wsSrc).value
                    
                    If UBound(arrSrc) > 1 Then
                        For y = 2 To UBound(arrSrc)
                            Set d_temp = CreateObject("scripting.dictionary")
                            d_temp("机构名称") = CStr(arr_info(0))
                            d_temp("任务名称") = CStr(arr_info(1))
                            d_temp("数据日期") = CStr(arr_info(3))
                            
                            
                            For j = 1 To UBound(arrSrc, 2)
                                d_temp(CStr(arrSrc(1, j))) = CStr(arrSrc(y, j))
                            Next
                            
                            If check_type = "在线校验比对信息" Then
                                Set d_data_zx(hzi) = d_temp
                            Else
                                Set d_data_lx(hzi) = d_temp
                            End If
                            hzi = hzi + 1 ' 有多个文件，索引会重复使用
                            Set d_temp = Nothing
                        Next
                        
                    End If
                    
                End If
            Next
                
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")

            hzNum = hzNum + 1
    
            wbsrc.Close False
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F,非校验对比信息文件或文件命名不规范")
        End If
        
NextHzSrcFiles:
        Call gloCft.updateScreen(i)
    Next
    
    Do While wbDes.Worksheets.count > 1
        wbDes.Worksheets(1).Delete
    Loop
           
    Dim arr_temp As Variant
    ' 最后，将数组字典转换到指定工作表
    If d_data_lx.count > 0 Then
        
        arr_temp = dictsTolists(Split(gloCft.GetConfig("离线校验对比信息标题"), ","), d_data_lx, True)
        ' 输出到工作表
        Set wsdes = wbDes.Worksheets.add
        Call gloEt.CancelFilter(wsdes)
        wsdes.name = "离线校验对比信息汇总"
        wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).NumberFormatLocal = "@"
        wsdes.Range("F:K").NumberFormatLocal = "0"
        wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).value = arr_temp
        Erase arr_temp
        
    End If
    Set d_data_lx = Nothing
    
    If d_data_zx.count > 0 Then
        arr_temp = dictsTolists(Split(gloCft.GetConfig("在线校验对比信息标题"), ","), d_data_zx, True)
        ' 输出到工作表
        Set wsdes = wbDes.Worksheets.add
        Call gloEt.CancelFilter(wsdes)
        wsdes.name = "在线校验对比信息汇总"
        wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).NumberFormatLocal = "@"
        wsdes.Range("F:G").NumberFormatLocal = "0"
        wsdes.Range("a1").Resize(UBound(arr_temp), UBound(arr_temp, 2)).value = arr_temp
        Erase arr_temp
        
    End If
    Set d_data_zx = Nothing
    

    MsgBox "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总"
    Call clsmymenu.RecoverSysMenu(srcfiles)
    
    wbDes.Windows(1).FreezePanes = False
    wbDes.Windows(1).SplitRow = 1
    wbDes.Windows(1).FreezePanes = True
    wbDes.Save
    Call EraseSomething
    Call gloCft.OpenSomething
Exit Sub
errhandle:
    MsgBox "校验信息汇总失败 -- " & err.Description
    err.Clear
    Set d_data_lx = Nothing
    Set d_data_zx = Nothing
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub


Private Function clear_str(s As String) As String
    Dim rs As String
    
    rs = Trim(s)
    
    rs = replace(rs, "，", ",")
    
    clear_str = rs
End Function

'
'
'
Private Sub isInmodCheck()
    '
    Dim arr As Variant, ws_cz As Worksheet, d1 As Object, i As Long, key2 As String

    Set ws_cz = gloEt.FindWs(ThisWorkbook, "临时取值模板", 0, True)
    arr = gloEt.GetUsedRange(ws_cz).value
    
    Set d_ce_mod = Nothing
    Set d_sc_mod = Nothing
    
    Set d_ce_mod = CreateObject("scripting.dictionary")
    Set d_sc_mod = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr)
        If CStr(arr(i, 8)) = "" Then
            '测试系统
            If d_ce_mod.exists(CStr(arr(i, 1))) = False Then
                Set d_ce_mod(CStr(arr(i, 1))) = CreateObject("scripting.dictionary")
            End If
            Set d1 = d_ce_mod(CStr(arr(i, 1)))
            
            ' 需要把对应哪一个的标记上
            key2 = clear_str(CStr(arr(i, 1)) & "," & CStr(arr(i, 4)) & "," & CStr(arr(i, 5)) & "," & CStr(arr(i, 1)) & "," & CStr(arr(i, 6)) & "," & CStr(arr(i, 7)))
            If d1.exists(CStr(arr(i, 4))) = False Then
                d1(CStr(arr(i, 4))) = key2
            Else
                d1(CStr(arr(i, 4))) = CStr(d1(CStr(arr(i, 4)))) & "," & key2
            End If
            Set d_ce_mod(CStr(arr(i, 1))) = d1
            Set d1 = Nothing
            
            '生产系统，工作表名有可能不同，所以把工作表名加进去
            If d_sc_mod.exists(CStr(arr(i, 1))) = False Then
                Set d_sc_mod(CStr(arr(i, 1))) = CreateObject("scripting.dictionary")
            End If
            Set d1 = d_sc_mod(CStr(arr(i, 1)))
            key2 = clear_str(CStr(arr(i, 1)) & "," & CStr(arr(i, 6)) & "," & CStr(arr(i, 7)))
            If d1.exists(CStr(arr(i, 6))) = False Then
                d1(CStr(arr(i, 6))) = key2
            Else
                d1(CStr(arr(i, 6))) = CStr(d1(CStr(arr(i, 6)))) & "," & key2
            End If
            Set d_sc_mod(CStr(arr(i, 1))) = d1
            Set d1 = Nothing
        End If
    
    Next
    
End Sub

Private Function is_mod(fn As String, wsn As String, t As Object) As String
    Dim k As Variant, i As Long, t2 As Object, k2 As Variant, j As Long, job_name As String
    k = t.keys
    
    For i = 0 To UBound(k)
        If InStr(fn, CStr(k(i))) > 0 Then
            Set t2 = t(CStr(k(i)))
            k2 = t2.keys
            job_name = CStr(k(i))
            For j = 0 To UBound(k2)
                If CStr(wsn) = CStr(k2(j)) Then
                    is_mod = t2(CStr(k2(j)))
                    Exit Function
                End If
            Next
        End If
    Next
    
    is_mod = ""
End Function

'
'    初始化双边数据
'
Private Function RePreData_Init_PreData() As String
    'On Error GoTo errHandel
    Dim fns As Variant, fn As String
    Dim wb_src As Workbook, ws_src As Worksheet, i As Long, j As Long, ws_t As Worksheet, arr As Variant, y As Long
    Dim d_data As Object, d_t As Object, org_name As String, target_code As String, data_value As Variant, date_str As String, job_name As String
    Dim mod_str As String, mod_arr As Variant, t_arr As Variant, g As Long, t_key As String
    
    
    '是否当期为空的时候不读取上期数v
    '构造数据字典
    Set d_predata = Nothing
    Set d_predata = CreateObject("scripting.dictionary")
    
    fn = CStr(files_sc)
    fns = Split(fn, ",")
    For y = 0 To UBound(fns)
        fn = fns(y)
        Set wb_src = gloEt.GetWb(fn)
        For Each ws_src In wb_src.Worksheets
            mod_str = is_mod(gloFt.getShortFileName(fn), ws_src.name, d_sc_mod)
            If mod_str <> "" Then
                arr = gloEt.GetUsedRange(ws_src).value
                
                t_arr = Split(gloFt.getShortFileName(fn), "_")
                org_name = t_arr(0)
                date_str = t_arr(2)
                mod_arr = Split(mod_str, ",")
                job_name = mod_arr(0)
                For g = 0 To UBound(mod_arr) Step 4
                    '读取数据
                    
                    t_key = org_name & "," & mod_arr(g) & "," & mod_arr(g + 1) & "," & mod_arr(g + 2) & "," & mod_arr(g + 3)
                    d_predata(t_key) = arr(mod_arr(g + 2), mod_arr(g + 3))
                Next
            End If
            '读取数据，机构，日期，任务，指标，数值1，数值2
        Next
        wb_src.Close False
        If y Mod 5 = 0 Then Call gloCft.ShowStatusMessage_Up("基础数据上期数读取，正在读取第：" & CStr(y + 1) & "家机构/" & CStr(UBound(fns) + 1) & "家")
    Next

End Function

Private Function DoRePreData(wb_name As String) As String
    On Error GoTo errhandle
    Dim i As Long, wb As Workbook, ws_src As Worksheet, mod_str As String, mod_arr As Variant, t_key As String, t_arr As Variant
    Dim org_name As String, date_str As String, g As Long
    Dim err_ex As String
    
    err_ex = ""
    Set wb = gloEt.GetWb(wb_name)
    If wb Is Nothing Then
        DoRePreData = "无法打开文件"
        Exit Function
    End If
    For Each ws_src In wb.Worksheets
        mod_str = is_mod(gloFt.getShortFileName(wb.fullName), ws_src.name, d_ce_mod)
        If mod_str <> "" Then
            
            t_arr = Split(gloFt.getShortFileName(wb.fullName), "_")
            org_name = t_arr(0)
            date_str = t_arr(2)
            mod_arr = Split(mod_str, ",")
            For g = 0 To UBound(mod_arr) Step 8
                '读取数据
                
                t_key = org_name & "," & mod_arr(g) & "," & mod_arr(g + 5) & "," & mod_arr(g + 6) & "," & mod_arr(g + 7)
                If d_predata.exists(t_key) Then
                    ws_src.Cells(ycV(mod_arr(g + 2)), ycV(mod_arr(g + 3))) = d_predata(t_key)
                Else
                    err_ex = err_ex & mod_arr(g) & "," & mod_arr(g + 1) & ",(" & mod_arr(g) & "," & mod_arr(g + 2) & ")无对对应上期数;"
                End If
            Next
        End If
        '读取数据，机构，日期，任务，指标，数值1，数值2
    Next
    If err_ex = "" Then
        DoRePreData = ""
    Else
        DoRePreData = err_ex
    End If
    
    wb.Close True
Exit Function
errhandle:
    If Not wb Is Nothing Then wb.Close False
    DoRePreData = "读取上期数失败 -- " & err.Description
    err.Clear
End Function

Public Sub RePreData()
    Call isInmodCheck
    '
    files_sc = gloFt.SelectFiles("选择上期核查表")
    'files_sc = "C:\Users\caoya\Desktop\2025-12-05 基础数据修改\生产系统\惠州农村商业银行股份有限公司_金融基础数据采集（单位贷款）_2025-09-30_数据核对表.xlsx"
    If files_sc = "" Then
        MsgBox "上期系统核查表未选择"
        Exit Sub
    End If
    
    
    Call gloCft.CloseSomething
    Application.EnableEvents = False
    
    If IsContinue("是否执行接收上期数") = False Then Exit Sub
    
    Call RePreData_Init_PreData
    
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SetMenuTitle
    If GetSrcFiles = False Then Exit Sub
    
    Dim i As Long, t As String, sum As Long, srcItem As Variant
    sum = 0
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        t = DoRePreData(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        If t = "" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            sum = sum + 1
        ElseIf t = "-" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "-")
            sum = sum + 1
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, t) '不输出到批注了
        End If
        
        If i Mod 5 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理第：" & CStr(i - 5) & "个/" & CStr(dsrcfiles.count) & "个")
    Next
    Call MsgBox("读取上期数结束,共执行" & CStr(dsrcfiles.count) & "个文件，成功读取" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    
    Call EraseSomething
    
    Call gloCft.OpenSomething
    Application.EnableEvents = True
    Call gloCft.ShowStatusMessage_Up("")
End Sub

'
'   清理宏代码
'
Public Sub DoClearMacroCode()
    Call gloCft.CloseSomething
    Application.EnableEvents = False
    'Application.AutomationSecurity = 2 '恢复选择
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行宏清除") = False Then Exit Sub
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SetMenuTitle
    If GetSrcFiles = False Then Exit Sub
    
    Dim i As Long, t As String, sum As Long, srcItem As Variant
    sum = 0
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        t = ClearMacroCode(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        If t = "" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            sum = sum + 1
        ElseIf t = "-" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "-")
            sum = sum + 1
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, t) '不输出到批注了
        End If
    Next
    Call MsgBox("清除宏结束,共执行" & CStr(dsrcfiles.count) & "个文件，成功清除" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    
    Call EraseSomething
    
    Call gloCft.OpenSomething
    Application.EnableEvents = True
End Sub

'
'   批量清除模板宏
'
Public Sub pl_clear_mod_data()
    'Call gloCft.CloseSomething
    '声明挪到前面好了
    Dim i As Long, t As String, sum As Long, srcItem As Variant, fill_value As Variant
    
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行模板数据清除") = False Then Exit Sub
    '选择数据区域填充内容
    fill_value = ycV(Application.InputBox("请输入数据区域填充内容", "选取提示", 0))
    If fill_value = 0 Then fill_value = ""
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.SetMenuTitle
    If GetSrcFiles = False Then Exit Sub
    
    sum = 0
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        ThisWorkbook.Worksheets("执行面板").Cells(1, 2).value = CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名")))
        ThisWorkbook.Worksheets("执行面板").Cells(1, 3).value = gloFt.getShortFileName2(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        
        t = clear_mod_data(fill_value)
        If t = "" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            sum = sum + 1
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, t) '不输出到批注了
        End If
    Next
    MsgBox ("清除结束，共用时" & CStr(Timer - gloT) & "秒，共执行" & CStr(dsrcfiles.count) & "个文件，成功清除" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    
    Call EraseSomething
    
    'Call gloCft.OpenSomething
End Sub

Public Function clear_mod_data(fill_value As Variant) As String
    On Error Resume Next
    Call gloCft.CloseSomething
    
    '声明挪到前面好了
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, arr_address As String
    

    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    If clsmymenu.InitModFileMessage = False Then clear_mod_data = "InitModFileMessage失败": Exit Function '已经包含建立标题行
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then clear_mod_data = "IsAvailableModTableSetting失败": Exit Function
    Call clsmymenu.SetMenuTitle
    
    Set wbMod = clsmymenu.GetMyHzModWb
    
    '清理外部文件,数据
    For Each v In clsmymenu.GetMyExtendMod.keys
        wbMod.Worksheets(CStr(v)).Cells.ClearContents
        If InStr(wbMod.Worksheets(CStr(v)).name, "参照表") > 0 Then wbMod.Worksheets(CStr(v)).Cells.NumberFormat = "@"
        wbMod.Worksheets(CStr(v)).Activate
        wbMod.Worksheets(CStr(v)).Range("c7").Select
    Next

    v = glodHzMod.items '获取子集
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        
        '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
        Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
        If clshzT.NumAreaAddress <> "" Then wsMod.Range(clshzT.NumAreaAddress).value = fill_value
        '清理文本区域
        If clshzT.TextAreaAddress <> "" Then wsMod.Range(clshzT.TextAreaAddress).value = ""
        wsMod.Activate
        wsMod.Range("c7").Select
        wsMod.Range("c7").Activate
    Next j
    wbMod.Close True
    Call gloCft.OpenSomething
    If err.Number <> 0 Then
        clear_mod_data = CStr(err.Description)
        err.Clear
    Else
        clear_mod_data = ""
    End If
    
End Function


Private Function ClearMacroCode(wb_name As String) As String
    On Error GoTo errhandle
    Dim i As Long, wb As Workbook
    
    Set wb = gloEt.GetWb(wb_name)
    'Application.AutomationSecurity = 3 '禁用宏
    'Application.AutomationSecurity = 2 '禁用宏
    If wb Is Nothing Then
        ClearMacroCode = "无法打开文件"
        Exit Function
    End If
    If CheckMacro(wb) = False Then
        ClearMacroCode = "-"
        wb.Close False
        Exit Function
    End If
    Dim vbc As Object
    '遍历wbk工作簿的每一个模块
    For Each vbc In wb.VBProject.VBComponents
        '如果是Excel对象的模块,则清除其中的代码,否则删除整个模块
        If vbc.Type = 100 Then
            Debug.Print vbc.name
            Debug.Print vbc.CodeModule.CountOfLines
            vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
            
'            Debug.Print vbc.CodeModule.Lines
'            Debug.Print vbc.CodeModule.VBE.VBProjects.add
        Else
            wb.VBProject.VBComponents.Remove vbc
        End If
    Next vbc
    ClearMacroCode = ""
    wb.Close True
Exit Function
errhandle:
    If Not wb Is Nothing Then wb.Close False
    ClearMacroCode = "清除宏失败 -- " & err.Description
    err.Clear
End Function

'
'  读取相关函数名
'
Private Function ReadMacroCodeText(CodeModule As Object, i As Long) As String
    With CodeModule
        If .ProcOfLine(i, 3) <> "" Then
            ReadMacroCodeText = .ProcOfLine(i, 3)
        ElseIf .ProcOfLine(i, 1) <> "" Then
            ReadMacroCodeText = .ProcOfLine(i, 1)
        ElseIf .ProcOfLine(i, 2) <> "" Then
            ReadMacroCodeText = .ProcOfLine(i, 2)
        ElseIf .ProcOfLine(i, 0) <> "" Then
            ReadMacroCodeText = .ProcOfLine(i, 0)
        Else
            ReadMacroCodeText = ""
        End If
    End With
End Function

Private Function CheckMacro(wbk As Workbook) As Boolean
    '参数wbk为要检查的工作簿
    Dim i As Long
    Dim vbc As Object
    '遍历wbk工作簿的每一个模块
    For Each vbc In wbk.VBProject.VBComponents
        '逐行检查模块中是否有代码
        With vbc.CodeModule
            '.DeleteLines 1, .CountOfLines
            For i = 1 To .CountOfLines
                If .ProcOfLine(i, 3) <> "" Or _
                   .ProcOfLine(i, 1) <> "" Or _
                   .ProcOfLine(i, 2) <> "" Or _
                   .ProcOfLine(i, 0) <> "" Then
                    CheckMacro = True
                    Exit Function
                End If
            Next i
        End With
    Next vbc
    CheckMacro = False
End Function


'给新的工作簿增加宏代码
Private Sub AddMacroCode_Check(wbDes As Workbook)
    Dim MyCodes As String, fullName As String
    
    On Error GoTo errhandle
    
    '插入启动菜单
    MyCodes = gloCft.GetConfig("校验文件菜单宏")
    With wbDes.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines 2, MyCodes
    End With
    
    MyCodes = gloCft.GetConfig("校验文件工作宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work"
        .CodeModule.InsertLines 2, MyCodes
    End With
Exit Sub

errhandle:
    err.Clear
    MsgBox "说明文件功能受限。需要在EXCEL菜单 工具-宏－安全性－可靠发行商　中勾选【信任对于VB项目的访问】"
End Sub

Public Sub Accept_CheckFileAvail()
    gloCft.CloseSomething
    
    On Error Resume Next
    '选择源数据
    Dim fileName As Variant, f As Variant, wbsrc As Workbook, wsSrc As Worksheet, wsdes As Worksheet
    '机构代码，机构名称，指标代码，指标名称，本期，上期，增减，原因
    Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer, c5 As Integer, c8 As Integer
    Dim arr As Variant, d As Variant, i As Long, t As String, k As Integer
    c1 = 2: c2 = 3: c3 = 4: c4 = 5: c5 = 7: c8 = 16
    fileName = Application.GetOpenFilename(fileFilter:="所有文件 (*.*), *.*", title:="选择源数据", MultiSelect:=True)
    If VarType(fileName) = vbArray Or VarType(fileName) = 8200 Or VarType(fileName) = 8204 Then
        Set d = CreateObject("Scripting.Dictionary")
        For Each f In fileName
            '遍历每个工作簿
            Set wbsrc = gloEt.GetWb(CStr(f))
            Set wsSrc = wbsrc.Worksheets(1)
            
            arr = wsSrc.UsedRange.value
            If VarType(arr) = vbString Or VarType(arr) = vbEmpty Then
                Exit For
            End If
            For i = 2 To UBound(arr) Step 1
                If CStr(arr(i, c8)) <> "" Then
                    '这里机构类是否变成小写呢
                    t = CStr(arr(i, c1)) & CStr(arr(i, c2)) & CStr(arr(i, c3)) & CStr(arr(i, c4)) & CStr(arr(i, c5))
                    t = replace(t, " ", "")
                    t = replace(t, "　", "")
                    d(t) = CStr(arr(i, c8))
                End If
            Next i

            wbsrc.Close False
        Next
    
        Set wsdes = gloEt.FindWs(ThisWorkbook, "输出日记", 0)
        arr = wsdes.Range("A1").CurrentRegion.value
        '获取完源数据后，遍历
        For i = 2 To UBound(arr) Step 1
            If CStr(arr(i, c8)) = "" Then
                t = CStr(arr(i, c1)) & CStr(arr(i, c2)) & CStr(arr(i, c3)) & CStr(arr(i, c4)) & CStr(arr(i, c5))
                t = replace(t, " ", "")
                t = replace(t, "　", "")
                
                If d.exists(t) Then arr(i, c8) = CStr(d(t))
            End If
        Next i
        
        wsdes.Range("A1").CurrentRegion.value = arr
    End If
    gloCft.OpenSomething
End Sub


Private Sub EraseSomething()
    Set dStandOrg = Nothing
    Set clsmymenu = Nothing
    Set dsrcfiles = Nothing
    Set glodHzMod = Nothing
    Set dHzArr = Nothing
    Set dCheckRow = Nothing
End Sub

Private Sub SetMenuTitleValue(srcfiles As Variant, title As String, i As Long, value As String)
    srcfiles(i, clsmymenu.GetMenuTitleCol(title)) = value
End Sub

'
' 台账汇总，将源数据所有的表汇总在一起
'
Public Sub RecordHz()
    Call gloCft.CloseSomething
    
    On Error GoTo errhandle
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles = False Then Exit Sub
    
    Dim fnOut As String
    fnOut = gloFt.GetOutputFile
    If fnOut = "" Then
        Exit Sub
    End If
    
    '初始化面板
    
    
    '判断是否有列指标
    Dim wsdes As Worksheet, wbDes As Workbook, arrSplit As Variant, arrSplit2 As Variant, hzNum As Long, bsplit As Boolean, i As Long, j As Integer
    Dim wbsrc As Workbook, wsSrc As Worksheet, curRow As Long, arrSrc As Variant, srcItem As Variant, z As Long
    Set wbDes = gloEt.GetWb(fnOut)
    If wbDes Is Nothing Then
        Set wbDes = Workbooks.add
        wbDes.SaveAs fnOut
    End If
    Set wsdes = wbDes.Worksheets(1)
    Call gloEt.CancelFilter(wsdes)
    
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    
    arrSplit = Split(clsmymenu.GetConfig("设置不参与的工作表"), ";")
    arrSplit2 = Split(clsmymenu.GetConfig("设置参与的工作表"), ";")
    hzNum = 0
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        Set wbsrc = gloEt.GetWb(CStr(srcfiles(i, 2)))
         
        If wbsrc Is Nothing Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.InsertNewDebug("台账汇总", CStr(srcfiles(i, 2)), "-", "不参与汇总", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        
        For Each wsSrc In wbsrc.Worksheets
            If UBound(arrSplit2) = -1 Then
                bsplit = True
                For j = 0 To UBound(arrSplit) Step 1
                    If InStr(CStr(arrSplit(j)), wsSrc.name) > 0 Then
                        bsplit = False
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit(j)) & "*" Then
                        bsplit = False
                        Exit For
                    End If
                Next j
            Else
                bsplit = False
                For j = 0 To UBound(arrSplit2) Step 1
                    If InStr(CStr(arrSplit2(j)), wsSrc.name) > 0 Then
                        bsplit = True
                        Exit For
                    ElseIf wsSrc.name Like "*" & CStr(arrSplit2(j)) & "*" Then
                        bsplit = True
                        Exit For
                    End If
                Next j
            End If
            
            If UBound(arrSplit2) = -1 And UBound(arrSplit) = -1 Then bsplit = True
            
            '如果没有被筛选，那么则进行汇总
            If bsplit = True Then
                If gloEt.IsBlankSheet(wsdes) Then
                    curRow = 1
                Else
                    curRow = gloEt.GetCurRows(wsdes) + 1
                End If
                
                If gloEt.IsBlankSheet(wsSrc) Then
                    Call gloDe.InsertNewDebug("台账汇总", wbsrc.name, wsSrc.name, "不参与汇总", "原因，空表")
                Else
                    '非空表
                    If gloEt.IsBlankSheet(wsdes) Then
                        curRow = 1
                    Else
                        curRow = gloEt.GetCurRows(wsdes) + 1
                    End If
                    
                    With gloEt.GetUsedRange(wsSrc)
                        wsdes.Range("C" & CStr(curRow)).Resize(.rows.count, .columns.count).ClearFormats
                        .Copy
                        wsdes.Range("C" & CStr(curRow)).Resize(.rows.count, .columns.count).PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
                        wsdes.Range("C" & CStr(curRow)).Resize(.rows.count, .columns.count).PasteSpecial Paste:=xlPasteValues, operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    
                        '补充A\B列
                        wsdes.Range(wsdes.Cells(curRow, 1), wsdes.Cells(curRow + .rows.count - 1, 1)).value = gloFt.getShortFileName(wbsrc.name)
                        wsdes.Range(wsdes.Cells(curRow, 2), wsdes.Cells(curRow + .rows.count - 1, 2)).value = wsSrc.name
                    End With
                    Application.CutCopyMode = False
                    
'                    '再用数组来进行赋值
'                    arrSrc = gloEt.GetUsedRange(wsSrc).value
'                    If VarType(arrSrc) = vbString Then
'                        wsDes.Range("C" & CStr(curRow)).value = arrSrc
'                    Else
'                        If UBound(arrSrc, 2) = 1 Then
'                            For z = 1 To UBound(arrSrc) Step 1
'                                wsDes.Range("C" & CStr(curRow + z - 1)).value = arrSrc(z, 1)
'                            Next z
'                        Else
'                            wsDes.Range("C" & CStr(curRow)).Resize(UBound(arrSrc), UBound(arrSrc, 2)).value = arrSrc
'                        End If
'                    End If
'
'                    '补充A\B列
'                    wsDes.Range(wsDes.Cells(curRow, 1), wsDes.Cells(curRow + UBound(arrSrc) - 1, 1)).value = gloFt.getShortFileName(wbSrc.name)
'                    wsDes.Range(wsDes.Cells(curRow, 2), wsDes.Cells(curRow + UBound(arrSrc) - 1, 2)).value = wsSrc.name
'                    Erase arrSrc
                    'Call gloDe.InsertNewDebug("台账汇总", wbSrc.name, wsSrc.name, "汇总成功")
                End If
            Else
                Call gloDe.InsertNewDebug("台账汇总", wbsrc.name, wsSrc.name, "不参与汇总")
            End If
            
        Next
        Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
        'Call gloDe.InsertNewDebug("台账汇总", wbSrc.name, "-", "汇总成功")
        'Call gloDe.AddColorDebug(15)
        hzNum = hzNum + 1

        wbsrc.Close False
NextHzSrcFiles:
        Call gloCft.updateScreen(i)
    Next

    Call gloDe.InsertNewDebug("台账汇总", wsdes.Parent.name, "-", "汇总结束", "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)

    wbDes.Save
    Call EraseSomething
    Call gloCft.OpenSomething
Exit Sub
errhandle:
    MsgBox "台账汇总失败 -- " & err.Description
    err.Clear
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub

'
'   批量清除源数据中的校验
'
Public Sub ClearSrcFilesCheckArea()
    Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, rAvilab As Integer, srcItem As Variant
    v = glodHzMod.items '获取子集
    
    Set wbMod = clsmymenu.GetMyHzModWb
    
    'Set wbOut = Workbooks.add
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")

    
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, 2)))
        rAvilab = 1
        If wbDes Is Nothing Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "清除结果", i, "F")
            Call gloDe.InsertNewDebug("清除校验", CStr(srcfiles(i, 2)), "-", "清除失败", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextSrcFiles
        End If


        For j = 0 To glodHzMod.count - 1 Step 1
            Set clshzT = v(j)
            
            '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
            Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
            Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
            
            If Not wsdes Is Nothing Then
                '清除校验区域
                If clshzT.CheckAreaAdderss <> "" Then wsdes.Range(clshzT.CheckAreaAdderss).Clear
                If clshzT.ExCheckAreaAddress <> "" Then wsdes.Range(clshzT.ExCheckAreaAddress).Clear
            Else
                rAvilab = 0
                Call gloDe.InsertNewDebug("清除校验", wbDes.name, "-", "清除不全", "源文件中找不到" & wsMod.name)
            End If
        Next j
        
        If rAvilab = 0 Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "清除结果", i, "F")
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "清除结果", i, "T")
        End If

        wbDes.Close True
NextSrcFiles:
    Next
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Call EraseSomething
    wbMod.Close False
    Call gloCft.OpenSomething
End Sub

'
'   源数据（逐指标）
'
Public Sub StandHz()
    Call gloCft.CloseSomething
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    
    If clsmymenu.IsHzRowColumnToData = False Then Exit Sub
    Call InitDicStandOrg
    '检查输出的工作簿是否合适，通过标题判断,如果存在相应工作簿，则需要标题一致
    Dim wbout As Workbook
    Call InitStandHzCol
    Set wbout = StandHzGetCheckOutputFile
    If wbout Is Nothing Then Exit Sub
    
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, hzNum As Long, rAvilab As Integer
    v = glodHzMod.items '获取子集
    
    
    hzqjy = val(clsmymenu.GetConfig("汇总去校验"))
    qzhz = val(clsmymenu.GetConfig("强制汇总"))
    qclz = val(clsmymenu.GetConfig("剔除0值"))
    
    Set wbMod = clsmymenu.GetMyHzModWb
    'Set wbOut = Workbooks.add
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    
    '初始化汇总数据
    If InitStandHz = False Then Exit Sub
    
    hzNum = 0
    For i = SYSSTARTROW To UBound(srcfiles) Step 1
        If Not dsrcfiles.exists(CStr(srcfiles(i, 2))) Then GoTo NextHzSrcFiles
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, 2)))
         
        If wbDes Is Nothing Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.InsertNewDebug("汇总校验", CStr(srcfiles(i, 2)), "-", "不参与汇总", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        rAvilab = IsAvailableFile(wbMod, wbDes)
        '判断工作簿是否通过校验
        If rAvilab = 0 Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        Else
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                
                '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                
                If Not wsdes Is Nothing Then
                    '汇总行列
                    Call StandHzByRowColumn(clshzT, wsdes)
                    'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, wsDes.name, "汇总成功")
                End If
            Next j
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "T")
            If rAvilab = 2 Then Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "汇总成功")
            'Call gloDe.AddColorDebug(15)
            hzNum = hzNum + 1
        End If
        wbDes.Close True
NextHzSrcFiles:
        Call gloCft.updateScreen(i)
    Next i
    Call StandHzRecover(wbout)
    Call gloDe.InsertNewDebug("汇总校验", wbMod.name, "-", "汇总结束", "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Set dHzArr = Nothing    '存储汇总数据
    wbMod.Close False
    Call EraseSomething
    
    Call gloCft.OpenSomething
End Sub

Private Function StandHzGetCheckOutputFile() As Workbook
    savePath = gloFt.GetOutputFile
    If savePath = "" Then
        Set StandHzGetCheckOutputFile = Nothing
        Exit Function
    End If
    
    Dim v As Variant, clshzT As CMyHzModTable, i As Integer, ws As Worksheet, wb As Workbook, j As Integer
    Dim vCol As Variant, kCol As Variant, iCol As Integer
    Dim preRow As Long
    '先判断是否存在,不存在则新建返回
    v = glodHzMod.items
    If UBound(v) = -1 Then
        Set StandHzGetCheckOutputFile = Nothing
        Exit Function
    End If
    Set clshzT = v(0)
    
    If gloFt.IsFileExist(savePath) = False Then
        Set wb = Workbooks.add
        Do While wb.Worksheets.count > 1
            wb.Worksheets(1).Delete
        Loop
        wb.Worksheets(1).name = clshzT.fileName & "-1"
        'wb.SaveAs savePath
        Call gloEt.save_as_excel(wb, savePath)
        Set StandHzGetCheckOutputFile = wb
        Exit Function
    End If
    
    Set wb = gloEt.GetWb(savePath)
    
    j = 1
    Do While j > 0
        Set ws = gloEt.FindWs(wb, clshzT.fileName & "-" & CStr(j), 0)
        If ws Is Nothing Then Exit Do
        j = j + 1
    Loop
    
    If j > 1 Then j = j - 1
    
    Set ws = gloEt.FindWs(wb, clshzT.fileName & "-" & CStr(j), 0, True)
    With ws
        'ws.Cells.Clear
        preRow = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row
        If preRow = 1 And .Cells(1, 1).value = "" Then
           
        Else
            '判断跟原设置标题是否一致
            kCol = dHzSysCol.keys
            vCol = dHzSysCol.items
            
            For iCol = 0 To UBound(kCol) Step 1
                If CStr(.Cells(1, CInt(vCol(iCol))) <> CStr(kCol(iCol))) Then
                    Exit For
                End If
                If iCol > UBound(kCol) Then
                    MsgBox "选择的文件与输出的表结构不一致，是否选错输出表。"
                    Set StandHzGetCheckOutputFile = Nothing
                End If
            Next iCol
'            If CStr(.Cells(1, 1)) <> "模板名" Or CStr(.Cells(1, 2)) <> "日期" Or CStr(.Cells(1, 3)) <> "冗余日期" Or _
'            CStr(.Cells(1, 4)) <> "机构名" Or CStr(.Cells(1, 5)) <> "文件名" Or CStr(.Cells(1, 6)) <> "填报单位" Or _
'            CStr(.Cells(1, 7)) <> "填报人" Or CStr(.Cells(1, 8)) <> "冗余列" Or CStr(.Cells(1, 9)) <> "联系方式" Or _
'            CStr(.Cells(1, 10)) <> "表格名" Or CStr(.Cells(1, 11)) <> "行序号" Or CStr(.Cells(1, 12)) <> "行名称" Or _
'            CStr(.Cells(1, 13)) <> "列名称" Or CStr(.Cells(1, 14)) <> "值" Then
'                MsgBox "选择的文件与输出的表结构不一致，是否选错输出表。"
'                Set StandHzGetCheckOutputFile = Nothing
'                Exit Function
'            End If
        End If
    End With
        
    Set StandHzGetCheckOutputFile = wb
End Function

'
'   获取对应列名称
'
Private Function StandHzGetColName(clshzT As CMyHzModTable, r As Integer, j As Integer) As String
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        StandHzGetColName = ""
        Exit Function
    End If
    
    Dim i As Integer, k1 As Variant, d As Object
    k1 = clshzT.dColumnAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i
    
    Set d = clshzT.dColumnAreaParse(k1(i))
    
    If d.exists(j) Then
        StandHzGetColName = CStr(d(j))
    Else
        StandHzGetColName = ""
    End If
End Function

'
'   获取改行数据所有列
'
Private Function StandHzGetColsIndex(clshzT As CMyHzModTable, r As Integer, j As Integer) As Variant
    On Error Resume Next
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        StandHzGetColsIndex = Split("", "|")
        Exit Function
    End If
    
    Dim arr As Variant
    Dim i As Integer, k1 As Variant, d As Object
    k1 = clshzT.dColumnAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i

    StandHzGetColsIndex = clshzT.dColumnAreaParse(CLng(ycV(k1(i)))).keys
    
    If err.Number <> 0 Then
        StandHzGetColsIndex = Split("", "|")
        err.Clear
    End If
End Function

'
'   获取对应列的行序号
'
Private Function StandHzGetCol_RowIndex(clshzT As CMyHzModTable, r As Integer, j As Integer) As Long
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        StandHzGetCol_RowIndex = 0
        Exit Function
    End If
    
    Dim i As Integer, k1 As Variant, d As Object
    k1 = clshzT.dColumnAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i

    StandHzGetCol_RowIndex = CLng(ycV(k1(i)))
End Function

'
'   获取对应的行名称
'
Private Function StandHzGetRowName(clshzT As CMyHzModTable, r As Integer, j As Integer) As String
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        StandHzGetRowName = ""
        Exit Function
    End If
    
    Dim i As Integer, k1 As Variant, d As Object
    k1 = clshzT.dRowAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If j > k1(i) And j <= k1(i + 1) Then
            Exit For
        End If
    Next i
    
    Set d = clshzT.dRowAreaParse(k1(i))
    
    If d.exists(r) Then
        StandHzGetRowName = CStr(d(r))
    Else
        StandHzGetRowName = ""
    End If
End Function

Private Sub InitEasyStandHzCol()
    Set dHzSysCol = CreateObject("scripting.dictionary")
    '文件名，日期，工作表名，填报单位，填报人，联系方式，行序号，行名称，列1，列2··
'    arr2(1, 1) = "文件名": arr2(1, 2) = "日期": arr2(1, 3) = "工作表名": arr2(1, 4) = "填报单位": arr2(1, 5) = "填报人": arr2(1, 6) = "联系方式"
'    arr2(1, 7) = "行序号": arr2(1, 8) = "行名称"
    dHzSysCol("文件名") = 1: dHzSysCol("日期") = 2: dHzSysCol("工作表名") = 3: dHzSysCol("设置单位") = 4
    dHzSysCol("填报单位") = 5: dHzSysCol("填报人") = 6: dHzSysCol("联系方式") = 7: dHzSysCol("行序号") = 8
    dHzSysCol("行名称") = 9
End Sub

Private Sub InitStandHzCol()
    Set dHzSysCol = CreateObject("scripting.dictionary")
'    arr2(1, 1) = "模板名": arr2(1, 2) = "日期": arr2(1, 3) = "冗余日期": arr2(1, 4) = "机构名": arr2(1, 5) = "文件名"
'    arr2(1, 6) = "填报单位": arr2(1, 7) = "填报人": arr2(1, 8) = "冗余列": arr2(1, 9) = "联系方式": arr2(1, 10) = "表格名"
'    arr2(1, 11) = "行序号": arr2(1, 12) = "行名称": arr2(1, 13) = "列名称": arr2(1, 14) = "值"
    dHzSysCol("业务类") = 1: dHzSysCol("日期") = 2: dHzSysCol("冗余列1") = 3: dHzSysCol("设置单位") = 5
    dHzSysCol("文件名") = 4: dHzSysCol("填报单位") = 6: dHzSysCol("填报人") = 7: dHzSysCol("冗余列2") = 8
    dHzSysCol("联系方式") = 9: dHzSysCol("表格名") = 10: dHzSysCol("行序号") = 11: dHzSysCol("行名称") = 12
    dHzSysCol("列名称") = 13: dHzSysCol("值") = 14
End Sub

''
'   计算简单标准汇总输出大小
'
Private Function InitStandHz() As Boolean
    '工作表名 + 数组
    '初始化汇总数据
    Dim clshzT As CMyHzModTable, v As Variant, j As Integer, i As Integer, k1 As Variant, k2 As Variant, v2 As Variant, i2 As Integer, col As Integer
    v = glodHzMod.items '获取子集
    
    Dim arr As Variant, rCells As Long, c As Long, arr2 As Variant, srcCount As Integer, vCol As Variant, kCol As Variant, iCol As Integer
    srcCount = dsrcfiles.count
    
    Set dHzArr = CreateObject("scripting.dictionary")
    
    rCells = 0
    '获取总行数
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        
        '获取文字区域
        If clshzT.TextAreaAddress <> "" Then
            arr = Split(clshzT.TextAreaParse, ",")
            rCells = rCells + CLng((UBound(arr) + 1) / 2)
        End If
        
        '获取公式区域
        If clshzT.FormatAreaAddress <> "" Then
            arr = Split(clshzT.FormatAreaParse, ",")
            rCells = rCells + CLng((UBound(arr) + 1) / 2)
        End If
        
        '获取数字区域
        If clshzT.NumAreaAddress <> "" Then
            arr = Split(clshzT.NumAreaParse, ",")
            rCells = rCells + CLng((UBound(arr) + 1) / 2)
        End If
    Next j
    
    If rCells = 0 Then
        InitStandHz = False
    Else
        ReDim arr2(1 To rCells * srcCount + 1, 1 To 14)
        
        kCol = dHzSysCol.keys
        vCol = dHzSysCol.items
        For iCol = 0 To UBound(kCol) Step 1
            arr2(1, CInt(vCol(iCol))) = CStr(kCol(iCol))
        Next iCol
        '2,4,6,7,9为不参与应对的列
'        arr2(1, 1) = "模板名": arr2(1, 2) = "日期": arr2(1, 3) = "冗余日期": arr2(1, 4) = "机构名": arr2(1, 5) = "文件名"
'        arr2(1, 6) = "填报单位": arr2(1, 7) = "填报人": arr2(1, 8) = "冗余列": arr2(1, 9) = "联系方式": arr2(1, 10) = "表格名"
'        arr2(1, 11) = "行序号": arr2(1, 12) = "行名称": arr2(1, 13) = "列名称": arr2(1, 14) = "值"
        
        dHzArr(clshzT.fileName) = arr2
        StandHzCurRow = 1  '2
        InitStandHz = True
    End If
End Function

'
'   源数据（全行）汇总，化作流水
'
Public Sub EasyStandHz()
    Call gloCft.CloseSomething
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行汇总") = False Then Exit Sub
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    
    If clsmymenu.IsHzRowColumnToData = False Then Exit Sub
    Call InitDicStandOrg
    '检查输出的工作簿是否合适，通过标题判断,如果存在相应工作簿，则需要标题一致
    Dim wbout As Workbook
    Call InitEasyStandHzCol
    Set wbout = GetEasyStandHzCheckOutputFile
    If wbout Is Nothing Then Exit Sub
    
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, hzNum As Long, rAvilab As Integer, srcItem As Variant
    v = glodHzMod.items '获取子集
    
    
    hzqjy = val(clsmymenu.GetConfig("汇总去校验"))
    qzhz = val(clsmymenu.GetConfig("强制汇总"))
    qclz = val(clsmymenu.GetConfig("剔除0值"))
    
    'Set dHzSysCol = CreateObject("scripting.dictionary")
    
    Set wbMod = clsmymenu.GetMyHzModWb
    'Set wbOut = Workbooks.add
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    '初始化汇总数据
    If InitEasyStandHz = False Then Exit Sub
    
    hzNum = 0
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
         
        If wbDes Is Nothing Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.InsertNewDebug("汇总校验", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "不参与汇总", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        rAvilab = IsAvailableFile(wbMod, wbDes)
        '判断工作簿是否通过校验
        If rAvilab = 0 Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        Else
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                
                '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                
                If Not wsdes Is Nothing Then
                    '汇总行列
                    Call EasyStandHzByRowColumn(clshzT, wsdes)
                    'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, wsDes.name, "汇总成功")   '单张工作表汇总成功不希纳是
                End If
            Next j
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "T")
            If rAvilab = 2 Then Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "汇总成功")  '成功的不显示
            'Call gloDe.AddColorDebug(15)
            hzNum = hzNum + 1
        End If
        wbDes.Close True
NextHzSrcFiles:
        Call gloCft.updateScreen(i)
    Next
    Call EasyStandHzRecover(wbout)
    Call gloDe.InsertNewDebug("汇总校验", wbMod.name, "-", "汇总结束", "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.Close False
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub

'
'   根据行列进行汇总
'
Private Sub StandHzByRowColumn(clshzT As CMyHzModTable, wsdes As Worksheet)
    If clshzT.NumAreaAddress = "" And clshzT.TextAreaAddress = "" And clshzT.FormatAreaAddress = "" Then Exit Sub
    
    '获取数字区域
    Dim col As Integer, row As Long, curArr As Variant, arr As Variant, arrSrcData As Variant, i As Integer
    Set dCheckRow = Nothing
    arrSrcData = wsdes.Range(clshzT.AllRangeAddress).value
    curArr = dHzArr(clshzT.fileName)
    
    Dim c1 As Variant, c2 As Variant, c3 As Variant, c4 As Variant, c5 As Variant, c6 As Variant, c7 As Variant, c8 As Variant, c9 As Variant, c10 As Variant
    Dim szdw As Variant
    c4 = "": c8 = "--"
    c1 = clshzT.fileName    '模板名
    c5 = left(wsdes.Parent.name, InStrRev(wsdes.Parent.name, ".") - 1)  '文件名
    If clshzT.TableDate = "" Then
        c2 = ""
    Else
        c2 = arrSrcData(Split(clshzT.TableDate, ",")(0), Split(clshzT.TableDate, ",")(1))
    End If
    'c2 = IIf(clshzT.TableDate = "", "", arrSrcData(Split(clshzT.TableDate, ",")(0), Split(clshzT.TableDate, ",")(1)))
    c2 = ChangeToCDate(c2)
    c3 = "--"
    c10 = clshzT.TableName
    If clshzT.TableFillerOrg <> "" Then
        c6 = arrSrcData(Split(clshzT.TableFillerOrg, ",")(0), Split(clshzT.TableFillerOrg, ",")(1))
        c6 = GetClearOrgName(CStr(c6))
    End If
    If clshzT.TableFiller <> "" Then
        c7 = arrSrcData(Split(clshzT.TableFiller, ",")(0), Split(clshzT.TableFiller, ",")(1))
        c7 = replace(c7, "填报人", "")
        c7 = replace(c7, "填表人", "")
        c7 = replace(c7, "：", "")
        c7 = replace(c7, ":", "")
    End If
    'c7 = CStr(ycDateToStr(c2)) & " " & CStr(c7)
    If clshzT.TableFillerContact <> "" Then
        c9 = arrSrcData(Split(clshzT.TableFillerContact, ",")(0), Split(clshzT.TableFillerContact, ",")(1))
        c9 = replace(c9, "联系方式：", "")
        c9 = replace(c9, ":", "")
    End If
    If clshzT.TableFillerOrgSetting <> "" Then
        c4 = arrSrcData(Split(clshzT.TableFillerOrgSetting, ",")(0), Split(clshzT.TableFillerOrgSetting, ",")(1))
    End If
    
    If CStr(c4) = "" Then c4 = GetMyOrgName(CStr(c6))
    If CStr(c4) = "" Then c4 = ycGetOrgName(CStr(c5))
    If CStr(c4) = "" Then c4 = c6
    If CStr(c4) = "" Then c4 = CStr(c5)
    
    '获取设置日期,没有的话，则用表格日期
    If clshzT.TableDateSetting <> "" Then
        If CStr(arrSrcData(Split(clshzT.TableDateSetting, ",")(0), Split(clshzT.TableDateSetting, ",")(1))) <> "" Then
            c2 = arrSrcData(Split(clshzT.TableDateSetting, ",")(0), Split(clshzT.TableDateSetting, ",")(1))
            c2 = ChangeToCDate(c2)
        End If
    End If

    If clshzT.NumAreaAddress <> "" Then
        arr = Split(clshzT.NumAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            '不为错误项目和空值
            If VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> vbError Then
                If val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> 0 Then
                    StandHzCurRow = StandHzCurRow + 1
                    curArr(StandHzCurRow, CInt(dHzSysCol("业务类"))) = c1: curArr(StandHzCurRow, CInt(dHzSysCol("日期"))) = c2: curArr(StandHzCurRow, CInt(dHzSysCol("冗余列1"))) = c3
                    curArr(StandHzCurRow, CInt(dHzSysCol("设置单位"))) = CStr(c4): curArr(StandHzCurRow, CInt(dHzSysCol("文件名"))) = c5: curArr(StandHzCurRow, CInt(dHzSysCol("填报单位"))) = CStr(c6)
                    curArr(StandHzCurRow, CInt(dHzSysCol("填报人"))) = CStr(c7): curArr(StandHzCurRow, CInt(dHzSysCol("冗余列2"))) = c8: curArr(StandHzCurRow, CInt(dHzSysCol("联系方式"))) = c9
                    curArr(StandHzCurRow, CInt(dHzSysCol("表格名"))) = c10
                    curArr(StandHzCurRow, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
                    curArr(StandHzCurRow, CInt(dHzSysCol("行名称"))) = StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                    curArr(StandHzCurRow, CInt(dHzSysCol("列名称"))) = StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                    curArr(StandHzCurRow, CInt(dHzSysCol("值"))) = val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
                End If
            End If
        Next i
    End If
    
    If clshzT.TextAreaAddress <> "" Then
        arr = Split(clshzT.TextAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            If CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> "" Then
                StandHzCurRow = StandHzCurRow + 1
                curArr(StandHzCurRow, CInt(dHzSysCol("业务类"))) = c1: curArr(StandHzCurRow, CInt(dHzSysCol("日期"))) = c2: curArr(StandHzCurRow, CInt(dHzSysCol("冗余列1"))) = c3
                curArr(StandHzCurRow, CInt(dHzSysCol("设置单位"))) = CStr(c4): curArr(StandHzCurRow, CInt(dHzSysCol("文件名"))) = c5: curArr(StandHzCurRow, CInt(dHzSysCol("填报单位"))) = c6
                curArr(StandHzCurRow, CInt(dHzSysCol("填报人"))) = CStr(c7): curArr(StandHzCurRow, CInt(dHzSysCol("冗余列2"))) = c8: curArr(StandHzCurRow, CInt(dHzSysCol("联系方式"))) = c9
                curArr(StandHzCurRow, CInt(dHzSysCol("表格名"))) = c10
                curArr(StandHzCurRow, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
                curArr(StandHzCurRow, CInt(dHzSysCol("行名称"))) = StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                curArr(StandHzCurRow, CInt(dHzSysCol("列名称"))) = StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                curArr(StandHzCurRow, CInt(dHzSysCol("值"))) = CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
                'StandHzCurRow = StandHzCurRow + 1
            End If
        Next i
    End If
    
    If clshzT.FormatAreaAddress <> "" Then
        arr = Split(clshzT.FormatAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            If VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> vbError Then
                If VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) = vbString Then
                    If CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> "" Then
                        StandHzCurRow = StandHzCurRow + 1
                        curArr(StandHzCurRow, CInt(dHzSysCol("业务类"))) = c1: curArr(StandHzCurRow, CInt(dHzSysCol("日期"))) = c2: curArr(StandHzCurRow, CInt(dHzSysCol("冗余列1"))) = c3
                        curArr(StandHzCurRow, CInt(dHzSysCol("设置单位"))) = CStr(c4): curArr(StandHzCurRow, CInt(dHzSysCol("文件名"))) = c5: curArr(StandHzCurRow, CInt(dHzSysCol("填报单位"))) = c6
                        curArr(StandHzCurRow, CInt(dHzSysCol("填报人"))) = CStr(c7): curArr(StandHzCurRow, CInt(dHzSysCol("冗余列2"))) = c8: curArr(StandHzCurRow, CInt(dHzSysCol("联系方式"))) = c9
                        curArr(StandHzCurRow, CInt(dHzSysCol("表格名"))) = c10
                        curArr(StandHzCurRow, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
                        curArr(StandHzCurRow, CInt(dHzSysCol("行名称"))) = StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                        curArr(StandHzCurRow, CInt(dHzSysCol("列名称"))) = StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                        curArr(StandHzCurRow, CInt(dHzSysCol("值"))) = CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
                        'StandHzCurRow = StandHzCurRow + 1
                    End If
                ElseIf val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) <> 0 Then
                    StandHzCurRow = StandHzCurRow + 1
                    curArr(StandHzCurRow, CInt(dHzSysCol("业务类"))) = c1: curArr(StandHzCurRow, CInt(dHzSysCol("日期"))) = c2: curArr(StandHzCurRow, CInt(dHzSysCol("冗余列1"))) = c3
                    curArr(StandHzCurRow, CInt(dHzSysCol("设置单位"))) = CStr(c4): curArr(StandHzCurRow, CInt(dHzSysCol("文件名"))) = c5: curArr(StandHzCurRow, CInt(dHzSysCol("填报单位"))) = c6
                    curArr(StandHzCurRow, CInt(dHzSysCol("填报人"))) = CStr(c7): curArr(StandHzCurRow, CInt(dHzSysCol("冗余列2"))) = c8: curArr(StandHzCurRow, CInt(dHzSysCol("联系方式"))) = c9
                    curArr(StandHzCurRow, CInt(dHzSysCol("表格名"))) = c10
                    curArr(StandHzCurRow, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
                    curArr(StandHzCurRow, CInt(dHzSysCol("行名称"))) = StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                    curArr(StandHzCurRow, CInt(dHzSysCol("列名称"))) = StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                    curArr(StandHzCurRow, CInt(dHzSysCol("值"))) = val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
                    'StandHzCurRow = StandHzCurRow + 1
                End If
            End If
        Next i
    End If
    
    dHzArr(clshzT.fileName) = curArr
    Erase curArr
    Set dCheckRow = Nothing
    Erase arrSrcData
    Erase arr
End Sub

'
'   根据行列进行汇总
'
Private Sub EasyStandHzByRowColumn(clshzT As CMyHzModTable, wsdes As Worksheet)
    If clshzT.NumAreaAddress = "" And clshzT.TextAreaAddress = "" And clshzT.FormatAreaAddress = "" Then Exit Sub
    
    '获取数字区域
    Dim col As Integer, row As Long, curArr As Variant, arr As Variant, arrSrcData As Variant, i As Integer
    Set dCheckRow = Nothing
    arrSrcData = wsdes.Range(clshzT.AllRangeAddress).value
    curArr = dHzArr(clshzT.wsName)
    
    Dim c1 As Variant, c2 As Variant, c3 As Variant, c4 As Variant, c5 As Variant, c6 As Variant, szdw As Variant
    c1 = left(wsdes.Parent.name, InStrRev(wsdes.Parent.name, ".") - 1)
    If clshzT.TableDate = "" Then
        c2 = ""
    Else
        c2 = arrSrcData(Split(clshzT.TableDate, ",")(0), Split(clshzT.TableDate, ",")(1))
    End If
    If CStr(c2) <> "" Then c2 = ChangeToCDate(c2)
    c3 = clshzT.TableName
    If clshzT.TableFillerOrg <> "" Then
        c4 = arrSrcData(Split(clshzT.TableFillerOrg, ",")(0), Split(clshzT.TableFillerOrg, ",")(1))
        c4 = GetClearOrgName(CStr(c4))
    End If
    If clshzT.TableFiller <> "" Then
        c5 = arrSrcData(Split(clshzT.TableFiller, ",")(0), Split(clshzT.TableFiller, ",")(1))
        c5 = replace(c5, "填报人", "")
        c5 = replace(c5, "填表人", "")
        c5 = replace(c5, ":", "")
        c5 = CStr(replace(c5, "：", ""))
    End If
    c5 = CStr(ycDateToStr(c2)) & " " & CStr(c5)
    If clshzT.TableFillerContact <> "" Then
        c6 = arrSrcData(Split(clshzT.TableFillerContact, ",")(0), Split(clshzT.TableFillerContact, ",")(1))
        c6 = replace(c6, "联系方式", "")
        c6 = replace(c6, ":", "")
        c6 = CStr(replace(c6, "：", ""))
    End If
    
    '获取设置单位
    If clshzT.TableFillerOrgSetting <> "" Then
        szdw = arrSrcData(Split(clshzT.TableFillerOrgSetting, ",")(0), Split(clshzT.TableFillerOrgSetting, ",")(1))
    End If
    '如果设置单位为空，则使用机构填报的标准化的，不然就使用文件名称取末尾,最后就用机构填报的
    If szdw = "" Then szdw = GetMyOrgName(CStr(c4))
    If szdw = "" Then szdw = ycGetOrgName(CStr(c1))
    If szdw = "" Then szdw = CStr(c4)
    If szdw = "" Then szdw = CStr(c1)
    
    '获取设置日期,没有的话，则用表格日期
    If clshzT.TableDateSetting <> "" Then
        If CStr(arrSrcData(Split(clshzT.TableDateSetting, ",")(0), Split(clshzT.TableDateSetting, ",")(1))) <> "" Then
            c2 = arrSrcData(Split(clshzT.TableDateSetting, ",")(0), Split(clshzT.TableDateSetting, ",")(1))
            c2 = ChangeToCDate(c2)
        End If
    End If
        
     '文件名，日期，工作表名，填报单位，填报人，联系方式，行序号，行名称，列1，列2··
    If clshzT.NumAreaAddress <> "" Then
        arr = Split(clshzT.NumAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            col = GetDataToArrIndex(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            row = GetEasyStandHzRow(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            
            If VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) = vbError Then
                curArr(row, col) = 0
            Else
                curArr(row, col) = val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
            End If
            'curArr(row, col) = val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
            If curArr(row, col) = 0 Then curArr(row, col) = ""

            curArr(row, CInt(dHzSysCol("文件名"))) = c1: curArr(row, CInt(dHzSysCol("日期"))) = c2: curArr(row, CInt(dHzSysCol("工作表名"))) = c3
            curArr(row, CInt(dHzSysCol("设置单位"))) = szdw
            curArr(row, CInt(dHzSysCol("填报单位"))) = c4: curArr(row, CInt(dHzSysCol("填报人"))) = c5: curArr(row, CInt(dHzSysCol("联系方式"))) = c6
            curArr(row, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
            curArr(row, CInt(dHzSysCol("行名称"))) = GetDataRowName(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            Next i
    End If
    
    If clshzT.TextAreaAddress <> "" Then
        arr = Split(clshzT.TextAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            'curRow = curRow + 1
            col = GetDataToArrIndex(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            row = GetEasyStandHzRow(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            curArr(row, col) = CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
            curArr(row, CInt(dHzSysCol("文件名"))) = c1: curArr(row, CInt(dHzSysCol("日期"))) = c2: curArr(row, CInt(dHzSysCol("工作表名"))) = c3
            curArr(row, CInt(dHzSysCol("设置单位"))) = szdw
            curArr(row, CInt(dHzSysCol("填报单位"))) = c4: curArr(row, CInt(dHzSysCol("填报人"))) = c5: curArr(row, CInt(dHzSysCol("联系方式"))) = c6
            curArr(row, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
            curArr(row, CInt(dHzSysCol("行名称"))) = GetDataRowName(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
        Next i
    End If
    
    If clshzT.FormatAreaAddress <> "" Then
        arr = Split(clshzT.FormatAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            'curRow = curRow + 1
            col = GetDataToArrIndex(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            row = GetEasyStandHzRow(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
            curArr(row, col) = arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))
            If VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) = vbError Then
                curArr(row, col) = ""
            ElseIf VarType(arrSrcData(CInt(arr(i)), CInt(arr(i + 1)))) = vbString Then
                curArr(row, col) = CStr(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
            Else
                curArr(row, col) = val(arrSrcData(CInt(arr(i)), CInt(arr(i + 1))))
                If curArr(row, col) = 0 Then curArr(row, col) = ""
            End If
            
            curArr(row, CInt(dHzSysCol("文件名"))) = c1: curArr(row, CInt(dHzSysCol("日期"))) = c2: curArr(row, CInt(dHzSysCol("工作表名"))) = c3
            curArr(row, CInt(dHzSysCol("设置单位"))) = szdw
            curArr(row, CInt(dHzSysCol("填报单位"))) = c4: curArr(row, CInt(dHzSysCol("填报人"))) = c5: curArr(row, CInt(dHzSysCol("联系方式"))) = c6
            curArr(row, CInt(dHzSysCol("行序号"))) = CInt(arr(i))
            curArr(row, CInt(dHzSysCol("行名称"))) = GetDataRowName(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
        Next i
    End If
    
    dHzArr(clshzT.wsName) = curArr
    Erase curArr
    Set dCheckRow = Nothing
    Erase arrSrcData
    Erase arr
End Sub

Private Function GetEasyStandHzRow(clshzT As CMyHzModTable, r As Integer, j As Integer) As Long
    If dCheckRow Is Nothing Then
        Set dCheckRow = CreateObject("scripting.dictionary")
    End If
    
    '判断是否存在行，判断是否存在列
    If dCheckRow.exists(r) Then
'        If Not dCheckRow(r).exists(j) Then
'            '不存在，新增
'            dHzArrCount(clshzT.wsName) = dHzArrCount(clshzT.wsName) + 1
'            dCheckRow(r)(j) = dHzArrCount(clshzT.wsName)
'        End If
    Else
        '不存在，新增
        'Set dCheckRow(r) = CreateObject("scripting.dictionary")
        dHzArrCount(clshzT.wsName) = dHzArrCount(clshzT.wsName) + 1
        dCheckRow(r) = dHzArrCount(clshzT.wsName)
    End If
    GetEasyStandHzRow = dCheckRow(r)
End Function

Private Function GetDataToArrIndex(clshzT As CMyHzModTable, r As Integer, j As Integer) As Integer
    Dim g As Integer, i As Integer, i2 As Integer, k1 As Variant, v1 As Variant, k2 As Variant
    
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        GetDataToArrIndex = dHzSysCol.count + 1
        Exit Function
    End If
    
    k1 = clshzT.dColumnAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i
    k2 = clshzT.dColumnAreaParse(k1(i)).keys
    

    For i2 = 0 To UBound(k2) - 1 Step 1
        If j >= k2(i2) And j < k2(i2 + 1) Then
            Exit For
        End If
    Next i2

    GetDataToArrIndex = (i + 1) * (i2 + 1) + dHzSysCol.count
End Function

'
'   获取对应的行名称
'
Private Function GetDataRowName(clshzT As CMyHzModTable, r As Integer, j As Integer) As String
    Dim g As Integer, i As Integer, i2 As Integer, k1 As Variant, v1 As Variant, k2 As Variant
    
    If clshzT.RowAreaAddress = "" Or clshzT.ColumnAreaAddress = "" Then
        GetDataRowName = ""
        Exit Function
    End If
    
    k1 = clshzT.dRowAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If j > k1(i) And j <= k1(i + 1) Then
            Exit For
        End If
    Next i

    GetDataRowName = clshzT.dRowAreaParse(k1(i))(r)
End Function

''
'   计算简单标准汇总输出大小
'
Private Function InitEasyStandHz() As Boolean
    '工作表名 + 数组
    '初始化汇总数据
    Dim clshzT As CMyHzModTable, v As Variant, j As Integer, i As Integer, k1 As Variant, k2 As Variant, v2 As Variant, i2 As Integer, col As Integer
    v = glodHzMod.items '获取子集
    
    Dim dr As Object, dc As Object, arr As Variant, r As Long, c As Long, arr2 As Variant, srcCount As Integer
    Dim isHasDate As Boolean, vCol As Variant, kCol As Variant, iCol As Integer
    isHasDate = False
    srcCount = dsrcfiles.count
    
    Set dHzArr = CreateObject("scripting.dictionary")
    Set dHzArrCount = CreateObject("scripting.dictionary")
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        
        Set dr = CreateObject("scripting.dictionary")
        Set dc = CreateObject("scripting.dictionary")
        
        '获取文字区域
        If clshzT.TextAreaAddress <> "" Then
            arr = Split(clshzT.TextAreaParse, ",")
            
            For i = 0 To UBound(arr) Step 2
                dr(arr(i)) = 0
                dc(arr(i + 1)) = 0
            Next i
        End If
        
        '获取公式区域
        If clshzT.FormatAreaAddress <> "" Then
            arr = Split(clshzT.FormatAreaParse, ",")
            
            For i = 0 To UBound(arr) Step 2
                dr(arr(i)) = 0
                dc(arr(i + 1)) = 0
            Next i
        End If
        
        '获取数字区域
        If clshzT.NumAreaAddress <> "" Then
            arr = Split(clshzT.NumAreaParse, ",")
            
            For i = 0 To UBound(arr) Step 2
                dr(arr(i)) = 0
                dc(arr(i + 1)) = 0
            Next i
        End If
        
        If dr.count > 0 Then isHasDate = True
        '获取单元格大小
        r = IIf(dr.count > 0, dr.count * srcCount, 1 * srcCount)
        c = IIf(dc.count > 0, dc.count + dHzSysCol.count, dHzSysCol.count + 1)
        '获取实际有多少列
        Dim col_count As Long
        col_count = 0
        If clshzT.ColumnAreaAddress <> "" Then
            k1 = clshzT.dColumnAreaParse.items
            For i = 0 To UBound(k1) Step 1
                k2 = k1(i).items
                col_count = col_count + UBound(k2) + 1
            Next i
        End If
        If col_count + dHzSysCol.count > c Then c = col_count + dHzSysCol.count
        
        ReDim arr2(1 To r + 1, 1 To c)
        '遍历第一行，存放相应标题
        '文件名，日期，工作表名，填报单位，填报人，联系方式，行序号，行名称，列1，列2··
'        arr2(1, 1) = "文件名": arr2(1, 2) = "日期": arr2(1, 3) = "工作表名": arr2(1, 4) = "填报单位": arr2(1, 5) = "填报人": arr2(1, 6) = "联系方式"
'        arr2(1, 7) = "行序号": arr2(1, 8) = "行名称"
        
        kCol = dHzSysCol.keys
        vCol = dHzSysCol.items
        For iCol = 0 To UBound(kCol) Step 1
            arr2(1, CInt(vCol(iCol))) = CStr(kCol(iCol))
        Next iCol
        
        col = dHzSysCol.count + 1
        If clshzT.ColumnAreaAddress <> "" Then
            k1 = clshzT.dColumnAreaParse.items
            For i = 0 To UBound(k1) Step 1
                k2 = k1(i).items
                For i2 = 0 To UBound(k2) Step 1
                    arr2(1, col) = k2(i2)
                    col = col + 1
                Next i2
            Next i
        End If
        
        dHzArr(clshzT.wsName) = arr2
        dHzArrCount(clshzT.wsName) = 1
    Next j
    If isHasDate = False Then MsgBox "模板表未设置汇总区域。"
    InitEasyStandHz = isHasDate
End Function

'
'   输出
'
Private Function StandHzRecover(wb As Workbook) As Boolean
    '遍历
    On Error Resume Next
    Dim k As Variant, v As Variant, i As Integer, ws As Worksheet, preRow As Long, h As Integer
    Dim dch As Object, nArr As Variant, j As Long, tstr As String, cfRns As Range
    k = dHzArr.keys
    v = dHzArr.items
    
    For i = 0 To UBound(k) Step 1
        If UBound(v(i)) > 1 Then
            Set ws = EasyStandHzRecoverFindOutWs(CStr(k(i)), wb, 1)
            With ws
            
                preRow = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row
                If preRow = 1 And .Cells(1, 1).value = "" Then
                    .Range("A1").Resize(UBound(v(i)), UBound(v(i), 2)) = v(i)
                Else
                    .Range("A" & CStr(preRow + 1)).Resize(UBound(v(i)), UBound(v(i), 2)) = v(i)
                    .Range("A" & CStr(preRow + 1)).EntireRow.Delete
                End If
                
                .Cells.Interior.colorindex = 0
                '查重
                nArr = .UsedRange.value
                Set dch = CreateObject("scripting.dictionary")
                Set cfRns = Nothing
                For j = 2 To UBound(nArr) Step 1
                    ' 2-是日期，现在5是设置单位，是否需要改成 4- 文件名
                    tstr = CStr(nArr(j, 1)) & "," & CStr(nArr(j, 2)) & "," & CStr(nArr(j, 4)) & "," & CStr(nArr(j, 8)) & "," & _
                    CStr(nArr(j, 10)) & "," & CStr(nArr(j, 11)) & "," & CStr(nArr(j, 12)) & "," & CStr(nArr(j, 13))
                    If dch.exists(tstr) Then
                        '加入集合
                        If cfRns Is Nothing Then
                            Set cfRns = .Range(.Cells(j, 1), .Cells(j, 14))
                        Else
                            Set cfRns = Application.Union(cfRns, .Range(.Cells(j, 1), .Cells(j, 14)))
                        End If
                    Else
                        dch(tstr) = 1
                    End If
                Next j
                If Not cfRns Is Nothing Then cfRns.Interior.colorindex = 15
                Set cfRns = Nothing
                Set dch = Nothing
                Erase nArr
            End With
        End If
    Next i
    If err.Number <> 0 Then
        MsgBox "StandHzRecover:汇总输出出现异常" & err.Description
        err.Clear
    End If
    StandHzRecover = True
End Function

'
'   输出
'
Private Function EasyStandHzRecover(wb As Workbook) As Boolean
    '遍历
    On Error Resume Next
    Dim k As Variant, v As Variant, i As Integer, ws As Worksheet, preRow As Long, h As Integer
    Dim dch As Object, nArr As Variant, j As Long, tstr As String, cfRns As Range, col As Integer, lbz As Boolean
    Dim ws_fz As Worksheet, arr_fz As Variant, is_has_lz As Boolean
    k = dHzArr.keys
    v = dHzArr.items
    
    For i = 0 To UBound(k) Step 1
        If UBound(v(i)) > 1 Then
            Set ws = EasyStandHzRecoverFindOutWs(CStr(k(i)), wb)
            With ws
                'ws.Cells.Clear
                preRow = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row
                If preRow = 1 And .Cells(1, 1).value = "" Then
                    .Range("A1").Resize(UBound(v(i)), UBound(v(i), 2)) = v(i)
                Else
                    '如果有数据，判断标题是否一致,已经前期筛选了
'                    For h = 1 To UBound(v(i), 2) Step 1
'                        If CStr(v(1, h)) <> .Cells(1, h).Value Then
'                            EasyStandHzRecover = False
'                            Exit Function
'                        End If
'                    Next h
                    .Range("A" & CStr(preRow + 1)).Resize(UBound(v(i)), UBound(v(i), 2)) = v(i)
                    .Range("A" & CStr(preRow + 1)).EntireRow.Delete
                End If
                
                .Cells.Interior.colorindex = 0
                '查重
                nArr = gloEt.GetUsedRange(ws).value  '获取全部范围
                Set dch = CreateObject("scripting.dictionary")
                Set cfRns = Nothing
                For j = 2 To UBound(nArr) Step 1
'                    dHzSysCol("文件名") = 1: dHzSysCol("日期") = 2: dHzSysCol("工作表名") = 3: dHzSysCol("设置单位") = 4
'                    dHzSysCol("填报单位") = 5: dHzSysCol("填报人") = 6: dHzSysCol("联系方式") = 7: dHzSysCol("行序号") = 8
'                    dHzSysCol("行名称") = 9
                    tstr = CStr(nArr(j, CInt(dHzSysCol("文件名")))) & "," & CStr(nArr(j, CInt(dHzSysCol("日期")))) & "," & _
                        CStr(nArr(j, CInt(dHzSysCol("工作表名")))) & "," & CStr(nArr(j, CInt(dHzSysCol("行序号")))) & "," & _
                        CStr(nArr(j, CInt(dHzSysCol("行名称"))))
                    If dch.exists(tstr) Then
                        '加入集合
                        If cfRns Is Nothing Then
                            Set cfRns = .Range(.Cells(j, 1), .Cells(j, dHzSysCol.count))
                        Else
                            Set cfRns = Application.Union(cfRns, .Range(.Cells(j, 1), .Cells(j, dHzSysCol.count)))
                        End If
                    Else
                        dch(tstr) = 1
                    End If
                Next j
                If Not cfRns Is Nothing Then cfRns.Interior.colorindex = 15
                '去除0值
                Set cfRns = Nothing
                If qclz = 1 Then
                    is_has_lz = False '判断是否存在0行
                    Set ws_fz = gloEt.FindWs(ThisWorkbook, "yc辅助区", 0, True)
                    ws_fz.Cells.Clear
                    ReDim arr_fz(1 To UBound(nArr), 1 To 2) As Variant  '需要删除的标记1
                    For j = 2 To UBound(nArr) Step 1
                        lbz = True
                        'col 当前是第9列为数值，如果加入设置单位，就第10列开始
                        For col = dHzSysCol.count + 1 To UBound(nArr, 2) Step 1
                            If CStr(nArr(j, col)) <> "" Then
                                lbz = False
                                Exit For
                            End If
                        Next col
                        
                        If lbz = True Then
                            is_has_lz = True
                            arr_fz(j, 1) = 1
'                            If cfRns Is Nothing Then
'                                Set cfRns = .Cells(j, 1)
'                            Else
'                                Set cfRns = Application.Union(cfRns, .Cells(j, 1))
'                            End If
                        End If
                        
                    Next j
                    'If Not cfRns Is Nothing Then cfRns.EntireRow.Delete
                    If is_has_lz = True Then
                        '说明存在零行
                        ws_fz.Range("A1").Resize(UBound(arr_fz), UBound(arr_fz, 2)).value = arr_fz
                        '定位
                        Set cfRns = ws_fz.Range("A1").Resize(UBound(arr_fz), UBound(arr_fz, 2)).SpecialCells(2)
                        .Range(cfRns.EntireRow.address).Delete
                        ws_fz.Cells.Clear
                    End If
                End If
                Set cfRns = Nothing
                Set dch = Nothing
                Erase nArr
            End With
        End If
    Next i
    If err.Number <> 0 Then
        MsgBox "EasyStandHzRecover:汇总输出出现异常" & err.Description
        err.Clear
    End If
    EasyStandHzRecover = True
End Function

Private Function GetEasyStandHzCheckOutputFile() As Workbook
    savePath = gloFt.GetOutputFile
    If savePath = "" Then
        Set GetEasyStandHzCheckOutputFile = Nothing
        Exit Function
    End If
    
    Dim v As Variant, clshzT As CMyHzModTable, i As Integer, ws As Worksheet, wb As Workbook, v2 As Variant, v3 As Variant, j As Integer
    Dim preRow As Long
    '先判断是否存在,不存在则新建返回
    If gloFt.IsFileExist(savePath) = False Then
        Set wb = Workbooks.add
        Do While wb.Worksheets.count > 1
            wb.Worksheets(1).Delete
        Loop
        'wb.SaveAs savePath
        Call gloEt.save_as_excel(wb, savePath)
        Set GetEasyStandHzCheckOutputFile = wb
        Exit Function
    End If
    
    Set wb = gloEt.GetWb(savePath)
    v = glodHzMod.items
    For i = 0 To UBound(v) Step 1
        Set clshzT = v(i)
        Set ws = EasyStandHzRecoverFindOutWs2(clshzT.wsName, wb)
        j = 10
        With ws
            'ws.Cells.Clear
            preRow = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row
            If preRow = 1 And .Cells(1, 1).value = "" Then
               
            Else
                For Each v2 In clshzT.dColumnAreaParse.items
                    For Each v3 In v2.items
                        If CStr(v3) <> .Cells(1, j).value Then
                            MsgBox "选择的文件与输出的表结构不一致，是否选错输出表。"
                            Set GetEasyStandHzCheckOutputFile = Nothing
                            Exit Function
                        End If
                        j = j + 1
                    Next
                Next
            End If
        End With
    Next i
    Set GetEasyStandHzCheckOutputFile = wb
End Function

'
'   查找相应的输出工作簿
'
Private Function EasyStandHzRecoverFindOutWs(gname As String, wb As Workbook, Optional index As Integer = 0) As Worksheet
    Dim j As Integer, ws As Worksheet, i As Integer, c As Long, name As String, v As Variant, clshzT As CMyHzModTable
    
    If index = 0 Then
        name = gname
        v = glodHzMod.items
        For i = 0 To UBound(v) Step 1
            Set clshzT = v(i)
            
            If clshzT.wsName = gname Then
                name = clshzT.TableName
                Exit For
            End If
        Next i
        
    Else
        name = gname
    End If
    
    j = 1
    c = UBound(dHzArr(gname))
    Do While j > 0
        Set ws = gloEt.FindWs(wb, name & "-" & CStr(j), 0)
        If ws Is Nothing Then
            Exit Do
        End If
        j = j + 1
    Loop
    
    If ws Is Nothing And j = 1 Then
        '需要新增
        Set ws = wb.Worksheets.add
        ws.name = name & "-" & CStr(j)
    Else
        Set ws = wb.Worksheets(name & "-" & CStr(j - 1))
        '判断是否需要新增
        If ws.Cells.rows.count - ws.Range("A" & CStr(ws.Cells.rows.count)).End(xlUp).row - 2000 < c Then
            Set ws = wb.Worksheets.add
            ws.name = name & "-" & CStr(j)
        End If
    End If
    
    Set EasyStandHzRecoverFindOutWs = ws
End Function

'
'   查找相应的输出工作簿
'
Private Function EasyStandHzRecoverFindOutWs2(gname As String, wb As Workbook, Optional index As Integer = 0) As Worksheet
    Dim j As Integer, ws As Worksheet, i As Integer, v As Variant, clshzT As CMyHzModTable, name As String
    
    If index = 0 Then
        name = gname
        v = glodHzMod.items
        For i = 0 To UBound(v) Step 1
            Set clshzT = v(i)
            
            If clshzT.wsName = gname Then
                name = clshzT.TableName
                Exit For
            End If
        Next i
        
    Else
        name = gname
    End If
    
    j = 1
    Do While j > 0
        Set ws = gloEt.FindWs(wb, name & "-" & CStr(j), 0)
        If ws Is Nothing Then
            Exit Do
        End If
        j = j + 1
    Loop
    
    If ws Is Nothing And j = 1 Then
        '需要新增
        Set ws = wb.Worksheets.add
        ws.name = name & "-" & CStr(j)
    Else
        Set ws = wb.Worksheets(name & "-" & CStr(j - 1))
    End If
    
    Set EasyStandHzRecoverFindOutWs2 = ws
End Function

'
'   简单数据汇总
'
Public Sub EasyHz()
    Call gloCft.CloseSomething
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行汇总") = False Then Exit Sub
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    
    savePath = gloFt.GetOutputFile
    If savePath = "" Then Exit Sub
    '获取保存文件名
    
    hzqjy = val(clsmymenu.GetConfig("汇总去校验"))
    qzhz = val(clsmymenu.GetConfig("强制汇总"))
    
    
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, hzNum As Long, rAvilab As Integer, srcItem As Variant
    v = glodHzMod.items '获取子集
    
    Set wbMod = clsmymenu.GetMyHzModWb
    
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    '初始化汇总数据
    Set dHzArr = CreateObject("scripting.dictionary")
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        dHzArr(clshzT.WsIndex) = wbMod.Worksheets(clshzT.WsIndex).Range(clshzT.AllRangeAddress).value
    Next j
    
    hzNum = 0
    bnfEasyHz = False
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
         
        If wbDes Is Nothing Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.InsertNewDebug("汇总校验", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "不参与汇总", "文件不存在。")
            Call gloDe.AddColorDebug(15)
            GoTo NextHzSrcFiles
        End If
        rAvilab = IsAvailableFile(wbMod, wbDes)
        '判断工作簿是否通过校验
        If rAvilab = 0 Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        Else
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                
                '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                
                If Not wsdes Is Nothing Then
                    '汇总数字区域
                    Call EasyHzNumArea(clshzT, wsdes)
                    '汇总文字区域
                    Call EasyHzTextArea(clshzT, wsdes)
                    'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, wsDes.name, "汇总成功")
                End If
            Next j
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "T")
            If rAvilab = 2 Then Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "汇总成功")
            'Call gloDe.AddColorDebug(15)
            hzNum = hzNum + 1
            bnfEasyHz = True
        End If
        wbDes.Close True
NextHzSrcFiles:
        Call gloCft.updateScreen(i)
    Next
    Call EasyHzRecover(wbMod)
    Call gloDe.InsertNewDebug("汇总校验", wbMod.name, "-", "汇总结束", "共汇总" & CStr(hzNum) & "个文件，" & "有" & CStr(dsrcfiles.count - hzNum) & "个文件未参与汇总")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.SaveAs savePath
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub

'
'   还原简单汇总数据
'
Private Sub EasyHzRecover(wbMod As Workbook)
    Dim ws As Worksheet, j As Integer, arrHz As Variant, rns As Range, clshzT As CMyHzModTable, rnT As Range, arrVal As Variant
    Dim arr As Variant, v As Variant, i As Long
    
    v = glodHzMod.items
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
        If clshzT.NumAreaParse = "" Then Exit Sub
        
        Set ws = wbMod.Worksheets(clshzT.WsIndex)
        arrVal = dHzArr(clshzT.WsIndex)
        
        '数字区域
        arr = Split(clshzT.NumAreaAddress, ",")
        
        For i = 0 To UBound(arr) Step 1
            Set rns = ws.Range(CStr(arr(i)))
            
            For Each rnT In rns
                rnT.value = arrVal(rnT.row, rnT.Column)
            Next
        Next i
        
        '文字区域
        arr = Split(clshzT.TextAreaAddress, ",")
        
        For i = 0 To UBound(arr) Step 1
            Set rns = ws.Range(CStr(arr(i)))
            
            For Each rnT In rns
                rnT.value = arrVal(rnT.row, rnT.Column)
            Next
        Next i
        
        '校验区域
        If clshzT.CheckAreaAdderss <> "" And CInt(hzqjy) = 1 Then ws.Range(clshzT.CheckAreaAdderss).Clear
        If clshzT.ExCheckAreaAddress <> "" And CInt(hzqjy) = 1 Then ws.Range(clshzT.ExCheckAreaAddress).Clear
    Next j
End Sub

'
'   简单汇总数字区域
'
Private Sub EasyHzNumArea(clshzT As CMyHzModTable, wsdes As Worksheet)
    Dim arrArea As Variant, arrHz As Variant, arrSrc As Variant, i As Long
    If clshzT.NumAreaParse = "" Then Exit Sub
    
    arrArea = Split(clshzT.NumAreaParse, ",")
    arrHz = dHzArr(clshzT.WsIndex)
    arrSrc = wsdes.Range(clshzT.AllRangeAddress).value
    
    For i = 0 To UBound(arrArea) Step 2
        If VarType(arrSrc(arrArea(i), arrArea(i + 1))) = vbError Then
        
        Else
            '不是第一张表就相加
            If bnfEasyHz = True Then
                arrHz(arrArea(i), arrArea(i + 1)) = ycV(arrHz(arrArea(i), arrArea(i + 1))) + ycV(arrSrc(arrArea(i), arrArea(i + 1)))
            Else
                '是第一张表，就赋值
                If ycV(arrHz(arrArea(i), arrArea(i + 1))) <> 0 Then
                    Call gloDe.InsertNewDebug("汇总校验", "模板文件", "表" & CStr(clshzT.WsIndex), _
                            "模板含数据", "剔除模板 单元格(" & CStr(arrArea(i)) & "," & CStr(arrArea(i + 1)) & ")数据")
                    Call gloDe.AddColorDebug(3)
                End If
                arrHz(arrArea(i), arrArea(i + 1)) = ycV(arrSrc(arrArea(i), arrArea(i + 1)))
            End If
            If ycV(arrHz(arrArea(i), arrArea(i + 1))) = 0 Then arrHz(arrArea(i), arrArea(i + 1)) = ""
        End If
    Next i
    dHzArr(clshzT.WsIndex) = arrHz
End Sub

'
'   简单汇总文字区域
'
Private Sub EasyHzTextArea(clshzT As CMyHzModTable, wsdes As Worksheet)
    Dim arrArea As Variant, arrHz As Variant, arrSrc As Variant, i As Long
    If clshzT.TextAreaParse = "" Then Exit Sub
    
    arrArea = Split(clshzT.TextAreaParse, ",")
    arrHz = dHzArr(clshzT.WsIndex)
    arrSrc = wsdes.Range(clshzT.AllRangeAddress).value
    
    For i = 0 To UBound(arrArea) Step 2
        arrHz(arrArea(i), arrArea(i + 1)) = CStr(arrHz(arrArea(i), arrArea(i + 1))) + left(wsdes.Parent.name, InStrRev(wsdes.Parent.name, ".") - 1) + ":" + CStr((arrSrc(arrArea(i), arrArea(i + 1)))) + "。" + Chr(10)
    Next i
    dHzArr(clshzT.WsIndex) = arrHz
End Sub

'
'   判断汇总文件是否有效
'
Function IsAvailableFile(wbMod As Workbook, wbDes As Workbook) As Integer
    '根据A1的批注，判断文件是否审核通过
    Dim clshzT As CMyHzModTable, v As Variant, wsdes As Worksheet, cm As String, str As String, j As Integer
    v = glodHzMod.items '获取子集
    str = ""
    For j = 0 To glodHzMod.count - 1 Step 1
        Set clshzT = v(j)
            
        '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
        Set wsdes = FindTheHzTable(clshzT, wbMod.Worksheets(clshzT.WsIndex), wbDes)
        
        If wsdes Is Nothing Then
            str = clshzT.TableName & ":表格样式不一致或源文件中无相应报表。"
        Else
            cm = gloEt.GetCommentStr(wsdes.Cells(1, 1))
            If InStr(cm, "校验成功") > 0 Then
            
            ElseIf InStr(cm, "校验失败") > 0 Then
                str = str + wsdes.name + ":校验失败。"
            Else
                str = str + wsdes.name + ":未校验，请校验。"
            End If
        End If
    Next j
    
    If str <> "" Then
        If qzhz = 1 Then
            '若不是找不到表，又是强制汇总，则允许通过
            If InStr(str, "表格样式不一致或源文件中无相应报表") < 1 Then
                IsAvailableFile = 2
                Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "强制汇总", str)
            Else
                Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "不参与汇总", str)
            End If
        Else
            Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "不参与汇总", str)
            IsAvailableFile = 0
        End If
    Else
        IsAvailableFile = 1
    End If
End Function

'
'   初始化标准化机构字典
'
Private Sub InitDicStandOrg()
    Dim ws As Worksheet, arr As Variant, i As Long, t As String, t2 As String
    Set ws = ThisWorkbook.Worksheets("机构名称标准化")
    ws.Range("A1:A6536").NumberFormatLocal = "@"

    Set dStandOrg = Nothing
    Set dStandOrg = CreateObject("scripting.dictionary")
    
    arr = ws.Range("A1").CurrentRegion.value

    For i = 2 To UBound(arr) Step 1
        t = CStr(arr(i, 1))
        t2 = replace(replace(CStr(arr(i, 2)), "'", ""), " ", "")

        '这里就确保空的东西不会进来
        If t <> "" Then
            dStandOrg(t) = t2
        End If
    Next i
End Sub

'
'   反向校验，这个校验就不覆盖原来的原始文件
'
Public Sub CheckFileAvail_Rev()
    Call gloCft.CloseSomething
    
    '声明挪到前面好了
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, dset As Object, setK As Variant, setV As Variant, srcItem As Variant
    Dim kbase As Variant, savename As String
    
    
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行校验") = False Then Exit Sub
    
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    '插入日记
    Call glo_log.clear_log
    Call glo_log.create_log_title("模块", "工作簿", "工作表", "错误类型", "行位置", "行名称", "列位置", "列名称", "批注", "错误区域", "校验描述", "提取值", "列3", "列4", "列5", "说明")
    'Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "误差", "描述")
    '选择输出的文件夹
    
    is_rev_export_file = CBool(ycV(clsmymenu.GetConfig("反向校验输出文件")))
    Dim nFolder As String
    If is_rev_export_file = True Then
        nFolder = gloFt.GetOutputFolder("请选择转化保存的文件夹")
        If nFolder = "" Then Exit Sub
        If right(nFolder, 1) <> "\" Then nFolder = nFolder & "\"
    End If
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub   '已经包含建立标题行
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    If clsmymenu.InitExtendModFileMessage = False Then Exit Sub '判断是否存在外部文件，如果存在的话，则需要选择外部文件
   
    Call clsmymenu.SetMenuTitle
    
    jyqgsh = val(clsmymenu.GetConfig("校验去格式化"))
    sfqxtg = val(clsmymenu.GetConfig("强行通过"))
    is_check_data = CInt(ycV(clsmymenu.GetConfig("提取条件格式校验"))) ' CInt(ycV(ThisWorkbook.Worksheets("执行面板").Range("F1").value))
    is_get_all_col_data = CInt(ycV(clsmymenu.GetConfig("提取所有数据列"))) 'CInt(ycV(ThisWorkbook.Worksheets("执行面板").Range("M1").value))
    
    Set dset = clsmymenu.GetMenuTitleDic
    
    v = glodHzMod.items '获取子集
    
    Set wbMod = clsmymenu.GetMyHzModWb
    
    
    Call InitDicStandOrg
    
    
    
    '遍历items
    gloT = Timer
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        glo_str_relust = ""
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        
        If wbDes Is Nothing Then
            glo_str_relust = glo_str_relust & "校验失败,文件不存在。" & Chr(10)
            'Call gloDe.InsertNewDebug("汇总校验", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "总结", "校验失败,文件不存在。")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call gloDe.AddColorDebug(15)
        Else
            bErrAvailable = True    '工作簿校验通过标志
            '判断目标文件工作表数量与模板文件是否一致
            If glodHzMod.count <> wbDes.Worksheets.count Then
                '表格数量不一致,软性
                glo_str_relust = glo_str_relust & "（软性）目标文件与模板文件表格数量不一致。" & Chr(10)
                'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "软性", "目标文件与模板文件表格数量不一致")
            End If
            
            '再做一个表格样式检查
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                
                '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                
                If wsdes Is Nothing Then
                    'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "校验失败", "表格样式不一致或源文件中无相应报表。")
                    glo_str_relust = glo_str_relust & "表格样式不一致或源文件中无相应报表(" & CStr(clshzT.wsName) & ")。" & Chr(10)
                    'Call InsertCheckResulComment(wsDes, "校验失败")
                    bErrAvailable = False
                End If
            Next j
            
            If bErrAvailable = False Then GoTo nextfilecheck
            
            '确定有样式以后，才进入校验
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                
                '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                
                If Not wsdes Is Nothing Then
                    '进行数据区域复制
                    If clshzT.NumAreaAddress <> "" Then wsMod.Range(clshzT.NumAreaAddress).value = wsdes.Range(clshzT.NumAreaAddress).value
                    If clshzT.TextAreaAddress <> "" Then wsMod.Range(clshzT.TextAreaAddress).value = wsdes.Range(clshzT.TextAreaAddress).value
                    '基础设置进行复制
                    For Each kbase In clshzT.dbase_address.keys
                        wsMod.Range(CStr(kbase)).NumberFormatLocal = "@"  '设置为文本再复制
                        wsMod.Range(CStr(kbase)).value = wsdes.Range(CStr(kbase)).value
                    Next
                    '切换了，这里wsDes可以不用了
                    Set wsdes = wsMod
                    '判断机构名称是否满足标准化
                    If CheckStandOrgName(clshzT, wsdes) = False Then
                        Call clsmymenu.SetMenuTitleValue(srcfiles, "标准化", i, "F")
                        '名称标准化失败
                    Else
                        Call clsmymenu.SetMenuTitleValue(srcfiles, "标准化", i, "T")
                    End If
                    '判断数字区域是否合法性
                    If CheckNumberAreaLegality(clshzT, wsdes, True, wbDes.name) = False Then bErrAvailable = False
                    '替换相关指标 - 比如 汇率，日期
                    Call ReplaceSetValue(dset, clshzT.dSetAreaParse, wsdes, i)  '替换设置的指标
                    '复制 校验区域、额外校验区域 到目标文件，检查正确性
                    If CheckTheCheckArea(clshzT, wsMod, wsdes, True, wbDes.name) = False Then bErrAvailable = False
                    '检查数据区域有无标记的颜色,这里涉及到的一个情况就是，要不断读单元格数据了
                    If is_check_data <> 0 Then Call CheckNumber(clshzT, wsMod, wsdes, True, wbDes.name)
                Else
                    'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "校验失败", "表格样式不一致或源文件中无相应报表。")
                    glo_str_relust = glo_str_relust & "表格样式不一致或源文件中无相应报表。" & Chr(10)
                    'Call InsertCheckResulComment(wsDes, "校验失败")
                    bErrAvailable = False
                End If
            Next j
nextfilecheck:
            '输出工作簿校验结果
            If bErrAvailable Then
                'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "总结", "校验成功。")
                Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "T")
            Else
                'Call gloDe.InsertNewDebug("汇总校验", wbDes.name, "-", "总结", "校验失败。")
                Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            End If
            If glo_str_relust <> "" Then
                Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, glo_str_relust) '不输出到批注了
            Else
                Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, "")
            End If
            'Call gloDe.AddDeHypeLinks(2, wbDes.fullName, "汇总校验")
            'Call gloDe.AddColorDebug(15)
            '判断是否输出文件
            savename = nFolder & wbDes.name
            
            wbDes.Close False
            If is_rev_export_file = True Then
                wbMod.SaveAs savename, FileFormat:=xlOpenXMLWorkbook
                Call clsmymenu.SetMenuTitleValue(srcfiles, "新文件名", i, savename)
                '增加超链接
                Call clsmymenu.addhypelink("新文件名", i, savename)
            End If
        End If
        Call gloCft.updateScreen(i)
    Next
    
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.Close False
    Call EraseSomething
    
    '这里新建一个文件
    If glo_log.get_log_count > 1 Then
        Call glo_log.export_log_to_ws(ThisWorkbook.Worksheets("输出日记"), "a")
        MsgBox "执行结束，共用时。" & CStr(Timer - gloT) & "秒，执行总体结果看执行面板，详细校验结果看输出日记。"
    Else
        Call glo_log.export_log_to_ws(ThisWorkbook.Worksheets("输出日记"), "a")
        MsgBox "执行结束，共用时。" & CStr(Timer - gloT) & "秒，无校验结果。"
    End If
    Call gloCft.OpenSomething
End Sub

'
'   数据校验
'
Public Sub CheckFileAvail_jcsj(Optional is_jcsj As Boolean = False)
    Call CheckFileAvail(True)
End Sub

'
'   如果是基础数据，获取任务名称
'
Private Function is_jcsj_job_name(fn As String, mod_name) As Boolean
    On Error Resume Next
    
    Dim arr As Variant, name As String
    
    arr = Split(fn, "_", 3)
    
    If UBound(arr) > 1 Then
        name = clear_job_name(CStr(arr(1)))
    Else
        name = ""
    End If
    
    If name = "" Then
        is_jcsj_job_name = False
    Else
        If InStr(mod_name, name) > 0 Then
            is_jcsj_job_name = True
        Else
            is_jcsj_job_name = False
        End If
    End If
End Function

'
'   数据校验
'
Public Sub CheckFileAvail(Optional is_jcsj As Boolean = False)
    Call gloCft.CloseSomething
    
    '声明挪到前面好了
    Dim i As Long, j As Long, wbDes As Workbook, wbMod As Workbook, wsdes As Worksheet, wsMod As Worksheet
    Dim clshzT As CMyHzModTable, v As Variant, dset As Object, setK As Variant, setV As Variant, srcItem As Variant, jscj_job_name As String
    
    
    
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行校验") = False Then Exit Sub
    
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    '插入日记
    Call glo_log.clear_log
    Call glo_log.create_log_title("模块", "工作簿", "工作表", "校验结果类型", "行位置", "行名称", "列位置", "列名称", "批注(规则描述)", "校验区域类型", "校验字段", "校验值", "详细描述", "列4", "列5", "说明")
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub   '已经包含建立标题行
    Set glodHzMod = clsmymenu.GetMyHzMod
    If clsmymenu.IsAvailableModTableSetting = False Then Exit Sub
    If clsmymenu.InitExtendModFileMessage = False Then Exit Sub '判断是否存在外部文件，如果存在的话，则需要选择外部文件
    Call clsmymenu.SetMenuTitle
    
    
    jcsh_job_filter = CInt(ycV(clsmymenu.GetConfig("基础数据校验是否任务名称过滤")))
    jyqgsh = val(clsmymenu.GetConfig("校验去格式化"))
    sfqxtg = val(clsmymenu.GetConfig("强行通过"))
    log_des = CInt(ycV(clsmymenu.GetConfig("校验结果是否导出")))
    jy_link = CInt(ycV(clsmymenu.GetConfig("校验结果是否增加超链接")))
    is_check_data = CInt(ycV(clsmymenu.GetConfig("提取条件格式校验"))) ' CInt(ycV(ThisWorkbook.Worksheets("执行面板").Range("F1").value))
    is_get_all_col_data = CInt(ycV(clsmymenu.GetConfig("提取所有数据列"))) 'CInt(ycV(ThisWorkbook.Worksheets("执行面板").Range("M1").value))
    
    Set dset = clsmymenu.GetMenuTitleDic
    
    v = glodHzMod.items '获取子集
    
    Set wbMod = clsmymenu.GetMyHzModWb
    
    
    Call InitDicStandOrg
    
    '遍历items
    gloT = Timer
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        glo_str_relust = ""
        
        If is_jcsj = True And jcsh_job_filter <> 0 And is_jcsj_job_name(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件"))), wbMod.name) = False Then
            glo_str_relust = "过滤启用，任务名与模板不一致。"
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, glo_str_relust)
        Else
        
            Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
            
            If wbDes Is Nothing Then
                glo_str_relust = glo_str_relust & "校验失败,文件不存在。" & Chr(10)
                Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
                
            Else
                bErrAvailable = True    '工作簿校验通过标志
                '判断目标文件工作表数量与模板文件是否一致
                If glodHzMod.count <> wbDes.Worksheets.count Then
                    '表格数量不一致,软性
                    glo_str_relust = glo_str_relust & "（软性）目标文件与模板文件表格数量不一致。" & Chr(10)
                End If
                
                '再做一个表格样式检查
                For j = 0 To glodHzMod.count - 1 Step 1
                    Set clshzT = v(j)
                    
                    '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                    Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                    Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                    
                    If wsdes Is Nothing Then
                        glo_str_relust = glo_str_relust & "表格样式不一致或源文件中无相应报表(" & CStr(clshzT.wsName) & ")。" & Chr(10)
                        bErrAvailable = False
                    End If
                Next j
                
                If bErrAvailable = False Then GoTo nextfilecheck
                
                '确定有样式以后，才进入校验
                For j = 0 To glodHzMod.count - 1 Step 1
                    Set clshzT = v(j)
                    
                    '遍历模板信息，通过 工作表名称 进行匹配,还需要匹配表格一致性。
                    Set wsMod = wbMod.Worksheets(clshzT.WsIndex)
                    Set wsdes = FindTheHzTable(clshzT, wsMod, wbDes)
                    
                    If Not wsdes Is Nothing Then
                        '判断机构名称是否满足标准化
                        If CheckStandOrgName(clshzT, wsdes) = False Then
                            Call clsmymenu.SetMenuTitleValue(srcfiles, "标准化", i, "F")
                            '名称标准化失败
                        Else
                            Call clsmymenu.SetMenuTitleValue(srcfiles, "标准化", i, "T")
                        End If
                        '判断数字区域是否合法性
                        If CheckNumberAreaLegality(clshzT, wsdes) = False Then bErrAvailable = False
                        '替换相关指标 - 比如 汇率，日期
                        Call ReplaceSetValue(dset, clshzT.dSetAreaParse, wsdes, i)
                        '复制 校验区域、额外校验区域 到目标文件，检查正确性
                        If CheckTheCheckArea(clshzT, wsMod, wsdes) = False Then bErrAvailable = False
                        '检查数据区域有无标记的颜色,这里涉及到的一个情况就是，要不断读单元格数据了
                        If is_check_data <> 0 Then Call CheckNumber(clshzT, wsMod, wsdes)
                    Else
                        glo_str_relust = glo_str_relust & "表格样式不一致或源文件中无相应报表。" & Chr(10)
                        bErrAvailable = False
                    End If
                Next j
nextfilecheck:
                '输出工作簿校验结果
                If bErrAvailable Then
                    Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "T")
                Else
                    Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果", i, "F")
                End If
                If glo_str_relust <> "" Then
                    Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, glo_str_relust) '不输出到批注了
                Else
                    Call clsmymenu.SetMenuTitleValue(srcfiles, "校验结果汇总", i, "")
                End If
    
                Call gloEt.save_as_excel(wbDes, wbDes.fullName)
                wbDes.Close False
            End If
        End If
        
        If i Mod 5 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理第：" & CStr(i - 5) & "个/" & CStr(dsrcfiles.count) & "个")
    Next
    
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.Close False
    Call EraseSomething
    
    
    ' 这里判断是否需要输出到外部日记，总得来说，是判断是否是基础数据的校验，如果是的话，就重新处理下
    If glo_log.get_log_count > 1 Then
        If is_jcsj = False Then
            If log_des = 0 Then
                Call glo_log.export_log_to_ws(ThisWorkbook.Worksheets("输出日记"), "a")
            Else
                Call glo_log.export_log_to_nwb(GetMyHzModName)
            End If
        Else
            Call glo_log.export_log_to_nwb_jcsj(GetMyHzModName, jy_link)
        End If
        MsgBox "执行结束，共用时。" & CStr(Timer - gloT) & "秒，执行总体结果看执行面板，详细校验结果看输出日记。"
    Else
        MsgBox "执行结束，共用时。" & CStr(Timer - gloT) & "秒，无校验结果。"
    End If
    Call gloCft.OpenSomething
    
    Call gloCft.ShowStatusMessage_Up("")
End Sub

Private Function GetMyHzModName() As String
    Dim name As String
    name = ThisWorkbook.Worksheets("执行面板").Cells(1, 2).value
    name = gloFt.getShortFileName(name)
    GetMyHzModName = name
    
End Function

Public Sub export_log_to_file()
    Dim savePath As String
    savePath = gloFt.GetOutputFile  '获取保存校验文件结果，每次都要选择，太麻烦了
    If savePath = "" Then Exit Sub
    
    '选择输出所有还是，输出刚刚的校验结果
    Dim bsel As Boolean
    bsel = IsContinue("是否仅输出上一次校验结果，选否则输出所有校验日记")
    
    Dim wb_jiaoyan As Workbook, ws As Worksheet, rn As Range
    Set wb_jiaoyan = Workbooks.add
    wb_jiaoyan.Worksheets(1).name = "校验结果"
    Set ws = ThisWorkbook.Worksheets("输出日记")
    If bsel = True Then
        Set rn = ws.Range("A1").CurrentRegion
    Else
        Set rn = gloEt.GetUsedRange(ws)
    End If
    
    If Not rn Is Nothing Then
        wb_jiaoyan.Worksheets(1).Range(rn.address).value = rn.value
        wb_jiaoyan.Worksheets(1).Cells.EntireRow.WrapText = False
    End If
    '插入宏
    Call AddMacroCode_Check(wb_jiaoyan)
    wb_jiaoyan.SaveAs savePath
End Sub


'
'   判断数字区域是否均为数字
'
Private Function CheckNumberAreaLegality(clshzT As CMyHzModTable, wsdes As Worksheet, Optional is_rev As Boolean = False, Optional fileName As String) As Boolean
    Dim arrNum As Variant, i As Long, arrDes As Variant, str As String, t As Variant
    Dim str2 As String, fn As String
    Dim strComment As String, col_rowIndex As Long
    If clshzT.NumAreaParse = "" Then
        CheckNumberAreaLegality = True
        Exit Function
    End If
    fn = wsdes.Parent.name
    If is_rev = True Then
        fn = fileName
        If fn = "" Then fn = wsdes.Parent.name
    End If
    str = ""
    str2 = ""
    arrNum = Split(clshzT.NumAreaParse, ",")
    arrDes = wsdes.Range(clshzT.AllRangeAddress).value
    For i = 0 To UBound(arrNum) Step 2
        '判断数据类型为文本，但是val后又是数值的那些需要提示
        t = arrDes(CLng(arrNum(i)), CLng(arrNum(i + 1)))
        If VarType(t) = vbEmpty Or VarType(t) = vbSingle Or _
            VarType(t) = vbLong Or VarType(t) = vbInteger Or _
            VarType(t) = vbDouble Or VarType(t) = vbDecimal Then
            '说明是数值，忽略
        ElseIf VarType(t) = vbError Then
            str2 = str2 & "cells(" & CStr(arrNum(i)) & "," & CStr(arrNum(i + 1)) & ")=存在错误项"
            
            strComment = gloEt.GetCommentStr(wsdes.Cells(CLng(arrNum(i)), CLng(arrNum(i + 1))))
            '这里判断有没有批注，没有的话，寻找列标题标题
            col_rowIndex = StandHzGetCol_RowIndex(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1)))
            If strComment = "" Then
                If col_rowIndex <> 0 Then
                    strComment = gloEt.GetCommentStr(wsdes.Cells(col_rowIndex, CLng(arrNum(i + 1))))
                End If
            End If
            strComment = replace(strComment, "Administrator:", "")
            strComment = replace(strComment, Chr(10), "")
            
            str2 = "数据区域中存在错误项目。" & str2

            Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "软性错误", _
                                                                    ycS(arrNum(i)), StandHzGetRowName(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1))), _
                                                                    ycS(arrNum(i + 1)), StandHzGetColName(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1))), _
                                                                    strComment, _
                                                                    "数字区域", "", CStr(arrDes(arrNum(i), arrNum(i + 1))), "cells(" & CStr(arrNum(i)) & "," & CStr(arrNum(i + 1)) & ")=存在错误项")
        Else
            '这里做一个判断，如果转换以后跟转换前一样
            If val(t) <> 0 Then
                If CStr(val(t)) <> CStr(t) Then
                    str = "cells(" & CStr(arrNum(i)) & "," & CStr(arrNum(i + 1)) & ")=" & CStr(t) & "->" & val(t) & "。"
                    
                    strComment = gloEt.GetCommentStr(wsdes.Cells(CLng(arrNum(i)), CLng(arrNum(i + 1))))
                    '这里判断有没有批注，没有的话，寻找列标题标题
                    col_rowIndex = StandHzGetCol_RowIndex(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1)))
                    If strComment = "" Then
                        If col_rowIndex <> 0 Then
                            strComment = gloEt.GetCommentStr(wsdes.Cells(col_rowIndex, CLng(arrNum(i + 1))))
                        End If
                    End If
                    strComment = replace(strComment, "Administrator:", "")
                    strComment = replace(strComment, Chr(10), "")
                    
                    str = "存在非0值文本数值转换，且转换前后不一致。" & str

                    Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "软性错误", _
                                                                            ycS(arrNum(i)), StandHzGetRowName(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1))), _
                                                                            ycS(arrNum(i + 1)), StandHzGetColName(clshzT, CLng(arrNum(i)), CLng(arrNum(i + 1))), _
                                                                            strComment, _
                                                                            "数字区域", "", CStr(arrDes(arrNum(i), arrNum(i + 1))), "存在非0值文本数值转换，且转换前后不一致。")
                End If
            End If
        End If
    Next i
    
    If str <> "" Then
        
        
    End If
    If str2 <> "" Then
        
        
        CheckNumberAreaLegality = True
    Else
        CheckNumberAreaLegality = True
    End If
    CheckNumberAreaLegality = True
End Function

'
'   获取源数据，并清空非校验区域
'
Private Function GetSrcFiles(Optional srclear As Integer = 4) As Boolean
    Call clsmymenu.ClearMenuTitle(srclear)
    srcfiles = clsmymenu.GetSysMenuAreaValue
    Set dsrcfiles = clsmymenu.GetSysMenuSrcs
    
    If dsrcfiles.count = 0 Then
        GetSrcFiles = False
    Else
        GetSrcFiles = True
    End If
End Function

Private Function CheckNumber(clshzT As CMyHzModTable, wsMod As Worksheet, wsdes As Worksheet, Optional is_rev As Boolean = False, Optional fileName As String) As Boolean
    Dim i As Integer, arr As Variant, data As Variant, str As String, arr_temp As Variant, index As Integer
    Dim strComment As String, col_rowIndex As Long, arr_col As Variant, col_index As Long, alldata As String, modCmt As String
    If clshzT.NumAreaParse = "" Then
        CheckNumber = True
        Exit Function
    End If
    Dim fn As String
    fn = wsdes.Parent.fullName
    If is_rev = True Then
        fn = fileName
        If fn = "" Then fn = wsdes.Parent.fullName
    End If
    data = wsdes.Range(clshzT.AllRangeAddress).value
    Dim conds As Object, cond As Object
    With wsdes
        arr = Split(clshzT.NumAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            index = ActiveCondition(.Cells(ycV(arr(i)), ycV(arr(i + 1))), wsdes.Parent) '.Interior.ColorIndex
            If index > 0 Then
                strComment = gloEt.GetCommentStr(.Cells(CLng(arr(i)), CLng(arr(i + 1))))
                '这里判断有没有批注，没有的话，寻找列标题标题
                col_rowIndex = StandHzGetCol_RowIndex(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                If strComment = "" Then
                    If col_rowIndex <> 0 Then
                        strComment = gloEt.GetCommentStr(.Cells(col_rowIndex, CLng(arr(i + 1))))
                    End If
                End If
                strComment = replace(strComment, "Administrator:", "")
                strComment = replace(strComment, Chr(10), "")
                
                ' 补充原始模板中的批注
                modCmt = clshzT.GetCmt(CLng(arr(i)), CLng(arr(i + 1)))
                If modCmt <> strComment Then
                    strComment = strComment & ";" & modCmt
                End If
                
                '构造出所有的列数据
                
                alldata = CStr(data(arr(i), arr(i + 1)))
                alldata = ""
                If is_get_all_col_data = 1 Then
                    arr_col = clshzT.GetDataColArr(CLng(arr(i)), CLng(arr(i + 1)))
                    '列名 列号
                    For col_index = 0 To UBound(arr_col) Step 2
                        If alldata = "" Then
                            alldata = CStr(arr_col(col_index)) & "：" & _
                                CStr(Round(ycV(CStr(data(CLng(arr(i)), CLng(arr_col(col_index + 1))))), 2))
                        Else
                            alldata = alldata & "," & Chr(10) & CStr(arr_col(col_index)) & "：" & _
                                CStr(Round(ycV(CStr(data(CLng(arr(i)), CLng(arr_col(col_index + 1))))), 2))
                        End If
                    Next
                End If
                
                If InStr(strComment, "不显示") <= 0 Then
                    ' 环比不显示了
                    If InStr(StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), "环比") > 0 Then

                    ElseIf InStr(StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), "跨期") > 0 Then
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "条件格式提示", _
                                                                    ycS(arr(i)), StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), _
                                                                    ycS(arr(i + 1)), StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), _
                                                                    strComment, _
                                                                    "数字区域", "跨期校验过大", Format(ycV(data(arr(i), arr(i + 1))), "0.00"), alldata)
                    Else
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "条件格式提示", _
                                                                    ycS(arr(i)), StandHzGetRowName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), _
                                                                    ycS(arr(i + 1)), StandHzGetColName(clshzT, CLng(arr(i)), CLng(arr(i + 1))), _
                                                                    strComment, _
                                                                    "数字区域", "", Format(ycV(data(arr(i), arr(i + 1))), "0.00"), alldata)
                    End If
                End If
            End If
        Next
        
    End With
    CheckNumber = True
End Function


'
'   复制校验关系，及检查校验
'
Private Function CheckTheCheckArea(clshzT As CMyHzModTable, wsMod As Worksheet, wsdes As Worksheet, Optional is_rev As Boolean = False, Optional fileName As String) As Boolean
    '复制校验关系，先判断有没有合并单元格，有的话，先去掉好了
    On Error Resume Next
    Dim i As Integer, arr As Variant, data As Variant, str As String, arr_temp As Variant
    Dim fn As String
    fn = wsdes.Parent.fullName
    If is_rev = True Then
        fn = fileName
        If fn = "" Then fn = wsdes.Parent.fullName
    End If
    
    '如果不是反向的话，则需要把校验复制到原始数据工作表
    If is_rev = False Then
        If clshzT.CheckAreaAdderss <> "" Then
            arr = Split(clshzT.CheckAreaAdderss, ",")
            For i = 0 To UBound(arr) Step 1
                wsdes.Range(CStr(arr(i))).UnMerge
                'Call gloEt.CopyAndPaste(wsMod, wsDes, CStr(arr(i)), False)
                Call gloEt.CopyAndPaste(wsMod, wsdes, CStr(arr(i)), True, True)  '改为只复制公式，就不把本地跨工作表了
                '是否需要增加一次计算，暂不增加先
                If clshzT.CheckAreaAdderss <> "" Or clshzT.ExCheckAreaAddress <> "" Then wsdes.Calculate
                If jyqgsh = 1 Then wsdes.Range(CStr(arr(i))).value = wsdes.Range(CStr(arr(i))).value
            Next i
        End If
        
        If clshzT.ExCheckAreaAddress <> "" Then
            arr = Split(clshzT.ExCheckAreaAddress, ",")
            For i = 0 To UBound(arr) Step 1
                wsdes.Range(CStr(arr(i))).UnMerge
                Call gloEt.CopyAndPaste(wsMod, wsdes, CStr(arr(i)), False)
                If clshzT.ExCheckAreaAddress <> "" Then wsdes.Calculate
                wsdes.Range(CStr(arr(i))).value = wsdes.Range(CStr(arr(i))).value
            Next i
        End If
    End If
    
    
    '判断错误。
    Dim errNum1 As Long, errNum2 As Long, haserrorFm As Boolean
    errNum1 = 0
    errNum2 = 0
    Dim col_rowIndex As Long, col_row_name As String, row_name As String
    haserrorFm = False
    With wsdes
        data = wsdes.Range(clshzT.AllRangeAddress).value
        arr = Split(clshzT.CheckAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            If VarType(data(arr(i), arr(i + 1))) = vbError Or ycS(data(arr(i), arr(i + 1))) = "#N/A" Then
                Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "公式异常", ycS(arr(i)), , ycS(arr(i + 1)), , , "校验公式区域中，存在计算错误项目。")
                haserrorFm = True
            Else
                str = data(arr(i), arr(i + 1))
                col_row_name = ""
                row_name = "" '校验区域不考虑取行列名称了 GetDataRowName(clshzT, CInt(arr(i)), CInt(arr(i + 1)))
                ' col_rowIndex = StandHzGetCol_RowIndex(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                ' If col_rowIndex > 0 Then col_row_name = CStr(data(col_rowIndex, arr(i + 1)))
                If InStr(str, "错") > 0 Or InStr(str, "失败") > 0 Or InStr(str, "硬性") > 0 Or InStr(str, "软性") > 0 Then
                    arr_temp = CheckTheCheckArea_GetExplanArr(str)
                    If InStr(str, "软") > 0 Then
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "软性", ycS(arr(i)), row_name, ycS(arr(i + 1)), col_row_name, _
                                                            gloEt.GetCommentStr(.Cells(CLng(arr(i)), CLng(arr(i + 1)))), "本地校验区域", arr_temp(2), arr_temp(3), arr_temp(4), arr_temp(5))
                        
                    
                        errNum1 = errNum1 + 1
                    Else
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "硬性", ycS(arr(i)), row_name, ycS(arr(i + 1)), col_row_name, _
                                                            gloEt.GetCommentStr(.Cells(CLng(arr(i)), CLng(arr(i + 1)))), "本地校验区域", arr_temp(2), arr_temp(3), arr_temp(4), arr_temp(5))
                        errNum2 = errNum2 + 1
                    End If
                End If
            End If
        Next i
        
        arr = Split(clshzT.ExCheckAreaParse, ",")
        For i = 0 To UBound(arr) Step 2
            If VarType(data(arr(i), arr(i + 1))) = vbError Or ycS(data(arr(i), arr(i + 1))) = "#N/A" Then
                haserrorFm = True
                Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "公式异常", ycS(arr(i)), , ycS(arr(i + 1)), , , "校验公式中，存在计算错误项目。")
            Else
                str = data(arr(i), arr(i + 1))
                col_row_name = ""
                row_name = ""
                'col_rowIndex = StandHzGetCol_RowIndex(clshzT, CLng(arr(i)), CLng(arr(i + 1)))
                'If col_rowIndex > 0 Then col_row_name = CStr(data(col_rowIndex, arr(i + 1)))
                
                If InStr(str, "错") > 0 Or InStr(str, "失败") > 0 Or InStr(str, "硬性") > 0 Or InStr(str, "软性") > 0 Then
                    arr_temp = CheckTheCheckArea_GetExplanArr(str)
                    If InStr(str, "软") > 0 Then
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "软性", ycS(arr(i)), row_name, ycS(arr(i + 1)), col_row_name, _
                                                            gloEt.GetCommentStr(.Cells(CLng(arr(i)), CLng(arr(i + 1)))), "联表校验区域", arr_temp(2), arr_temp(3), arr_temp(4), arr_temp(5))
                        errNum1 = errNum1 + 1
                    Else
                        Call glo_log.add_log_stand("汇总校验", fn, wsdes.name, "硬性", ycS(arr(i)), row_name, ycS(arr(i + 1)), col_row_name, _
                                                            gloEt.GetCommentStr(.Cells(CLng(arr(i)), CLng(arr(i + 1)))), "联表校验区域", arr_temp(2), arr_temp(3), arr_temp(4), arr_temp(5))
                        errNum2 = errNum2 + 1
                    End If
                End If
            End If
        Next i
    End With
    
    If haserrorFm = True Then
        'Call glo_log.add_log_stand("汇总校验", fn, wsDes.name, "公式异常", ycS(arr(i)), , ycS(arr(i + 1)), col_row_name, , "校验公式中，存在计算错误项目。")
    End If
    
    If errNum1 + errNum2 > 0 Then
        glo_str_relust = glo_str_relust & wsdes.name & ":" & Chr(10)
    End If
    If errNum1 > 0 Then
        glo_str_relust = glo_str_relust & "  软性错误有" & CStr(errNum1) & "个。" & Chr(10)
    End If
    If errNum2 > 0 Then
        glo_str_relust = glo_str_relust & "  硬性错误有" & CStr(errNum2) & "个。" & Chr(10)
    End If
    If errNum2 = 0 Then
        Call InsertCheckResulComment(wsdes, "校验成功")
        CheckTheCheckArea = True
    Else
        Call InsertCheckResulComment(wsdes, "校验失败")
        CheckTheCheckArea = False
    End If
End Function

'
' 获取格式的误差
'
Private Function CheckTheCheckArea_GetExplanArr(str As String) As Variant
    On Error Resume Next
    Dim arr As Variant, arrR(1 To 6) As Variant
    Dim i As Integer
    
    For i = 1 To 5
        arrR(i) = ""
    Next
    
    arr = Split(str, "|", 6)
    For i = 0 To UBound(arr)
        arrR(i + 1) = ycS(arr(i))
    Next
    
    
    CheckTheCheckArea_GetExplanArr = arrR
End Function

'
' 获取格式的误差
'
Private Function CheckTheCheckArea_GetExplanValue(str As String) As Variant
    Dim v As Variant, t As String, f As Double
    
    Dim l1 As Integer, l2 As Integer
    str = replace(str, "【", "[")
    str = replace(str, "】", "]")
    l1 = InStr(str, "[")
    l2 = InStr(str, "]")
    
    If l1 > 0 And l2 > 0 And l2 - l1 + 1 > 0 Then
       t = Mid(str, l1 + 1, l2 - l1 - 1)
       f = ycV(t)
       
       If f = 0 Then
        v = t
       Else
        v = f
       End If
    Else
        v = ""
    End If
    
    
    CheckTheCheckArea_GetExplanValue = v
End Function

Private Sub InsertCheckResulComment(ws_t As Worksheet, str As String)
 
    If InStr(str, "校验成功") > 0 Or sfqxtg = 1 Then
        If UBound(Split(gloEt.GetCommentStr(ws_t.Range("A1")), "-")) > -1 Then
            Call gloEt.InsertComment(ws_t, (Split(gloEt.GetCommentStr(ws_t.Range("A1")), "-")(0)) & "-校验成功")
        Else
            Call gloEt.InsertComment(ws_t, "-校验成功")
        End If
    Else
        If UBound(Split(gloEt.GetCommentStr(ws_t.Range("A1")), "-")) > -1 Then
            Call gloEt.InsertComment(ws_t, (Split(gloEt.GetCommentStr(ws_t.Range("A1")), "-")(0)) & "-校验失败")
        Else
            Call gloEt.InsertComment(ws_t, "-校验失败")
        End If
    End If
End Sub

'
'   填报机构填报的名称是否满足标准化
'
Private Function CheckStandOrgName(clshzT As CMyHzModTable, wsdes As Worksheet) As Boolean
    '如果没有设置填报机构名称的话，则不需要取处理
    If clshzT.TableFillerOrg = "" Then
        CheckStandOrgName = False
        '这里添加好日记
        'Call gloDe.InsertNewDebug("校验", wsDes.Parent.name, wsDes.name, "名称标准化", "未设置机构名称标识批注。")
        Exit Function
    End If
    
    Dim orgName As String, arrNum As Variant, oldName As String
    arrNum = Split(clshzT.TableFillerOrg, ",")
    oldName = CStr(wsdes.Cells(CLng(arrNum(0)), CLng(arrNum(1))).value)
    oldName = GetClearOrgName(oldName)  '获取填报单位
    
    If oldName = "" Then
        CheckStandOrgName = False
        'Call gloDe.InsertNewDebug("校验", wsDes.Parent.name, wsDes.name, "名称标准化", "该机构未填报机构名称。")
    Else
        orgName = GetMyOrgName(oldName)
        If orgName = "" Then
            'Call gloDe.InsertNewDebug("校验", wsDes.Parent.name, wsDes.name, "名称标准化", "新增" & oldName & "，请标准化。")
            CheckStandOrgName = False
        Else
            CheckStandOrgName = True
        End If
    End If
End Function

'
'   执行批量转换成xlsx
'
Public Sub ChangeToXlsx()
    Call gloCft.CloseSomething
    Application.EnableEvents = False
    '选择转化后输出的文件夹
    Dim nFolder As String
    nFolder = gloFt.GetOutputFolder("请选择转化保存的文件夹")
    
    If nFolder = "" Then Exit Sub
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles(5) = False Then Exit Sub
    
    Dim i As Long, t As String, sum As Long, srcItem As Variant
    Call gloDe.CreateDebugTitle("模块", "源文件名", "类型", "描述")

    sum = 0
    
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        '这里的话，执行转换，别显示转换后的而结果。
        
        t = gloEt.XlsToXlsx(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), nFolder)
        
        If t <> "" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名")) = t
            srcfiles(i, clsmymenu.GetMenuTitleCol("源文件")) = gloFt.getShortFileName2(CStr(t))
            sum = sum + 1
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        End If

    Next
    Call gloDe.InsertNewDebug("转换xlsx", "-", "结束", _
        "共执行" & CStr(dsrcfiles.count) & "个文件，成功转化" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Call EraseSomething
    Call gloCft.OpenSomething
    Application.EnableEvents = True
End Sub

'
'   执行批量转换成xls
'
Public Sub ChangeToXls()
    Call gloCft.CloseSomething
    
    '选择转化后输出的文件夹
    Dim nFolder As String
    nFolder = gloFt.GetOutputFolder("请选择转化保存的文件夹")
    
    If nFolder = "" Then Exit Sub
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles(5) = False Then Exit Sub
    
    Dim i As Long, t As String, sum As Long, srcItem As Variant
    Call gloDe.CreateDebugTitle("模块", "源文件名", "类型", "描述")

    sum = 0
    
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        '这里的话，执行转换，别显示转换后的而结果。
        
        t = gloEt.XlsxToXls(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), nFolder)
        
        If t <> "" Then
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名")) = t
            srcfiles(i, clsmymenu.GetMenuTitleCol("源文件")) = gloFt.getShortFileName2(CStr(t))
            sum = sum + 1
        Else
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        End If

    Next
    Call gloDe.InsertNewDebug("转换xls", "-", "结束", _
        "共执行" & CStr(dsrcfiles.count) & "个文件，成功转化" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub



'
'   文件重命名模块
'

Public Sub DoRename()
    Call gloCft.CloseSomething
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行重命名") = False Then Exit Sub
    
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    Call clsmymenu.ReloadMenuTitleDic
    If GetSrcFiles(5) = False Then Exit Sub
    
    Dim i As Long, t As String, sum As Long, srcItem As Variant
    Call gloDe.CreateDebugTitle("模块", "源文件名", "类型", "描述")

    sum = 0
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        If CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("新文件名"))) = "" Then
            '如果没填，则不重命名
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
        Else
            If InStr(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("新文件名"))), "\") = 0 Then
                t = left(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), InStrRev(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "\") - 1) & _
                                        "\" & CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("新文件名")))
            End If
            If gloFt.RenameFile(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), t) = True Then
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
                srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名")) = t
                srcfiles(i, clsmymenu.GetMenuTitleCol("源文件")) = gloFt.getShortFileName2(CStr(t))
                sum = sum + 1
            Else
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            End If
        End If
    Next
    Call gloDe.InsertNewDebug("重命名", "-", "结束", _
        "共执行" & CStr(dsrcfiles.count) & "个文件，成功重命名" & CStr(sum) & "个，失败" & CStr(dsrcfiles.count - sum) & "个。")
    Call clsmymenu.RecoverSysMenu(srcfiles)
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub

'
'   反向设置
'
Public Sub PickSrvFileSetting()
    Call gloCft.CloseSomething
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行设置") = False Then Exit Sub
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    
    Dim i As Long, j As Long, wbDes As Workbook, wsdes As Worksheet, wbMod As Workbook
    Dim clshzT As CMyHzModTable, v As Variant, dset As Object, setK As Variant, setV As Variant, srcItem As Variant
    v = glodHzMod.items '获取子集
    Set wbMod = clsmymenu.GetMyHzModWb
    
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    
    Set dset = clsmymenu.GetMenuTitleDic

    
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        If wbDes Is Nothing Then
            Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "总结", "设置失败,文件不存在。")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.AddColorDebug(15)
        Else
            bErrAvailable = True
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                Set wsdes = FindTheHzTable(clshzT, wbMod.Worksheets(clshzT.WsIndex), wbDes)
                'Set wsDes = gloEt.FindWs(wbDes, clshzT.wsName, 0)
                If wsdes Is Nothing Then
                    bErrAvailable = False
                    Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), clshzT.wsName, "设置失败", "不存在相应表。")
                Else
                    '替换相关指标 - 比如 汇率，日期
                    Call ReplaceSetValue(dset, clshzT.dSetAreaParse, wsdes, i)
                End If
            Next j
            If bErrAvailable = False Then
                Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "总结", "设置失败。")
                Call gloDe.AddColorDebug(15)
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Else
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            End If
            wbDes.Close True
        End If
    Next
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.Close False
    Call EraseSomething
    Call gloCft.OpenSomething
    
End Sub

'
'   替换设置的指标
'
Private Sub ReplaceSetValue(dset As Object, dmodSet As Object, wsdes As Worksheet, index As Long)
    If dset Is Nothing Or dmodSet Is Nothing Or dmodSet.count < 1 Or wsdes Is Nothing Then Exit Sub
    
    Dim i As Integer, k As Variant, col As Integer
    
    k = dmodSet.keys
    
    For i = 0 To UBound(k) Step 1
        If dset.exists(CStr(k(i))) Then
            col = dset(CStr(k(i)))
            If CStr(srcfiles(index, col)) <> "" Then
                wsdes.Range(CStr(dmodSet(k(i)))).value = srcfiles(index, col)
            End If
        Else
            '不存在，说明设置被取消了
        End If
    Next i
End Sub

'
'   提取设置
'
Public Sub PickFileSetting()
    Call gloCft.CloseSomething
    '每次执行，直接重新初始化模板信息啦
    If IsContinue("是否执行提取") = False Then Exit Sub
    Set clsmymenu = New CMySysMenu
    Call clsmymenu.SetSysWorkSheet("执行面板")
    
    If GetSrcFiles = False Then Exit Sub
    If clsmymenu.InitModFileMessage = False Then Exit Sub
    Set glodHzMod = clsmymenu.GetMyHzMod
    
    Dim i As Long, j As Long, wbDes As Workbook, wsdes As Worksheet, wbMod As Workbook
    Dim clshzT As CMyHzModTable, v As Variant, dset As Object, setK As Variant, setV As Variant, srcItem As Variant
    v = glodHzMod.items '获取子集
    
    Set wbMod = clsmymenu.GetMyHzModWb
    '插入日记
    Call gloDe.CreateDebugTitle("模块", "工作簿", "工作表", "错误类型", "描述")
    
    Set dset = clsmymenu.GetMenuTitleDic
    
    '遍历items
    For Each srcItem In dsrcfiles.items
        i = CLng(srcItem)
        Set wbDes = gloEt.GetWb(CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))))
        If wbDes Is Nothing Then
            Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "总结", "提取失败,文件不存在。")
            Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Call gloDe.AddColorDebug(15)
        Else
            bErrAvailable = True
            For j = 0 To glodHzMod.count - 1 Step 1
                Set clshzT = v(j)
                '查找不当
                Set wsdes = FindTheHzTable(clshzT, wbMod.Worksheets(clshzT.WsIndex), wbDes)
                'Set wsDes = gloEt.FindWs(wbDes, clshzT.wsName, 0)
                If wsdes Is Nothing Then
                    bErrAvailable = False
                    Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), clshzT.wsName, "提取失败", "不存在相应表。")
                Else
                    '获取相关指标 - 比如 汇率，日期
                    Call PickSetValue(dset, clshzT.dSetAreaParse, wsdes, i)
                End If
            Next j
            If bErrAvailable = False Then
                Call gloDe.InsertNewDebug("设置信息", CStr(srcfiles(i, clsmymenu.GetMenuTitleCol("源文件名"))), "-", "总结", "提取失败。")
                Call gloDe.AddColorDebug(15)
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "F")
            Else
                Call clsmymenu.SetMenuTitleValue(srcfiles, "执行结果", i, "T")
            End If
            wbDes.Close False
        End If
    Next
    
    Call clsmymenu.RecoverSysMenu(srcfiles)
    wbMod.Close False
    Call EraseSomething
    Call gloCft.OpenSomething
End Sub

Private Sub PickSetValue(dset As Object, dmodSet As Object, wsdes As Worksheet, index As Long)
    If dset Is Nothing Or dmodSet Is Nothing Or dmodSet.count < 1 Or wsdes Is Nothing Then Exit Sub
    
    Dim i As Integer, k As Variant, col As Integer
    
    k = dmodSet.keys
    
    For i = 0 To UBound(k) Step 1
        If dset.exists(CStr(k(i))) Then
            col = dset(CStr(k(i)))
            srcfiles(index, col) = wsdes.Range(CStr(dmodSet(k(i)))).value
        Else
            '不存在，说明设置被取消了
        End If
    Next i
End Sub

'
'   获取对应的工作表，并判断一致性
'
Public Function FindTheHzTable(clshzT As CMyHzModTable, wsMod As Worksheet, wbDes As Workbook) As Worksheet
    Dim i As Integer, ws As Worksheet, arrMod As Variant, arrDes As Variant, arrRows As Variant, arrCols As Variant
    Dim tRows As Long, tCols As Long
    
    Set ws = Nothing
    
    arrRows = Split(clshzT.RowAreaParse, ",")
    tRows = UBound(arrRows)
    arrCols = Split(clshzT.ColumnAreaParse, ",")
    tCols = UBound(arrCols)
    arrMod = wsMod.Range(clshzT.AllRangeAddress).value
    
    For i = 1 To wbDes.Worksheets.count Step 1
         With wbDes.Worksheets(i)
            If InStr(wsMod.name, .name) > 0 Or (.name Like "*" & FindTheHzTableDealName(wsMod.name) & "*") Then
            'If .name = wsMod.name Then
                Set ws = wbDes.Worksheets(i)
                
                If clshzT.CheckTabelUniqueAddress = "" And tRows = -1 And tCols = -1 Then
                    Call gloDe.InsertNewDebug("汇总校验", .Parent.name, .name, "匹配工作表", "当前模板未设置表格一致性字段，仅匹配表格名。")
                    Exit For
                End If
                
                '判断行、列一致性
                arrDes = .Range(clshzT.AllRangeAddress).value
                
                If clshzT.CheckTabelUniqueAddress <> "" Then
                    If CStr(wsMod.Range(clshzT.CheckTabelUniqueAddress).value) <> CStr(.Range(clshzT.CheckTabelUniqueAddress)) Then
                        Set ws = Nothing
                    End If
                End If
                
                '判断最后一行是否一致
                If Not ws Is Nothing And tRows > -1 And CStr(arrDes(arrRows(tRows - 1), arrRows(tRows))) <> CStr(arrMod(arrRows(tRows - 1), arrRows(tRows))) Then
                    Set ws = Nothing
                End If
                '判断最后一列是否一致
                If Not ws Is Nothing And tCols > -1 And CStr(arrDes(arrCols(tCols - 1), arrCols(tCols))) <> CStr(arrMod(arrCols(tCols - 1), arrCols(tCols))) Then
                    Set ws = Nothing
                End If
                
                If Not ws Is Nothing Then
                    Exit For
                End If
            End If
         End With
    Next i
    
    Set FindTheHzTable = ws
End Function

Private Function FindTheHzTableDealName(name As String) As String
    Dim s As String
    s = name
    
    s = replace(s, "X", "")
    s = replace(s, "#", "")
    s = replace(s, "x", "")
    
    FindTheHzTableDealName = s
End Function

'
'   获取清除好了的机构名称
'
Private Function GetClearOrgName(s As String) As String
    Dim r As String
    r = replace(s, "：", "")
    r = replace(r, ":", "")
    r = replace(r, "填报单位", "")
    r = replace(r, "填表机构", "")
    r = replace(r, "填报机构", "")
    r = replace(r, "填表单位", "")
    r = replace(r, " ", "")
    GetClearOrgName = r
End Function

'
'   获取标准化的机构名称
'
Private Function GetMyOrgName(orgName As String) As String
    On Error Resume Next
    Dim r As String, ws As Worksheet, curRow As Long
    Set ws = ThisWorkbook.Worksheets("机构名称标准化")
    curRow = gloEt.GetCurRows(ws)
    r = ""
    If orgName <> "" Then
        If dStandOrg.exists(orgName) Then
            r = CStr(dStandOrg(orgName))
        Else
            '追加进入机构名称标准化表
            ws.Cells(curRow + 1, 1) = orgName
            dStandOrg(orgName) = ""
        End If
    End If
    GetMyOrgName = r
End Function
