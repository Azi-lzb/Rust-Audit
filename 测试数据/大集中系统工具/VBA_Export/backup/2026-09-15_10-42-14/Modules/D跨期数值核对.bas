Attribute VB_Name = "D跨期数值核对"
'
' 进行数值比较，为什么要独立，因为可以看到差值有多大
' 是否区分本期上期数据，需要进行同期数据才可以核对吗。
'
'
' 输出说明统一到一起好了
'

Private d_data As Object    '大集中数据
Private arr_data() As Variant     '最终输出的数组
Private dUnitRef As Object '单位转换参照表  字典
Private dFres As Object '数据频度
Private dKuaQiCheck As Object
Private d_data_kuaqi As Object  '用于存储合成数据

Private d_org_ref As Object '机构类参照
Private d_area_ref As Object '地区参照

Private curPaths As String   '当期文件名
Private prePaths As String   '上期文件名
Private SrcUnit As String   '源数据单位
Private DesUnit As String   '目标数据单位

Private savePath As String  '最终存放路径

Private wbDes As Workbook
Private wsCur As Worksheet  ' 最终输出的文件
Private unitC As Double '源数据单位转换目标数据单位

Private d_calue_guocheng As Object  '记录过程

Private Sub EraseSomthing()

    Set d_data = Nothing    '大集中数据
    Erase arr_data     '最终输出的数组
    Set dUnitRef = Nothing '单位转换参照表  字典
    Set dFres = Nothing '数据频度
    Set dKuaQiCheck = Nothing
    Set d_data_kuaqi = Nothing  '用于存储合成数据
    
    Set d_org_ref = Nothing '机构类参照
    Set d_area_ref = Nothing '地区参照
End Sub

Sub 跨期数值核对(sp As String)
    
    Dim iOrgMode As Integer
    Call gloDe.CreateDebugTitle("跨期数值核对开始")
    gloT = Timer
    Call gloCft.CloseSomething
    savePath = sp  '最终存放路径
    '读取当期数据，并进行数值转换，不需要筛选频度
    
    SrcUnit = 导数比较窗体.源数据单位  '源数据单位
    DesUnit = 导数比较窗体.目标单位  '目标数据单位
    
    'SrcUnit = "万元"
    'DesUnit = "万元"
    
    iOrgMode = glo_iOrgMode
    unitC = GetUnitVal(SrcUnit) / GetUnitVal(DesUnit)
    
    
    If iOrgMode = 1 Then
        curPaths = 导数比较窗体.tb1.Value   '当期文件名
        prePaths = 导数比较窗体.tb2.Value  '上期文件名
    ElseIf iOrgMode = 2 Then
        '机构版的话，重新处理一个当期文件路径及上期文件路径给他
        Call InitChangeBjToCsv
        curPaths = ChangeBjToCsv(CStr(导数比较窗体.tb1.Value), "当期数")
        prePaths = ChangeBjToCsv(CStr(导数比较窗体.tb2.Value), "上期数")
        Call EndChangeBjToCsv

        If curPaths = "" Or prePaths = "" Then
            '
            MsgBox "IJ文件转CSV异常"
            Exit Sub
        End If
    Else
        MsgBox "未知模式"
        Exit Sub
    End If
    
    '初始化参照、数据
    Call InitSomething
    '计算指标
    Call CaluData
    '还原数组
    Call ReduceData
    '追加宏代码
    'Call AddMacroCode
    '结束
    Call EndCompare
    
    Call EraseSomthing
    Call gloDe.InsertNewDebug("跨期数值核对结束")
    Call gloCft.OpenSomething
    MsgBox "执行结束，共运行" & CStr(Timer - gloT) & "秒。"
End Sub

Private Sub EndCompare()
    Dim fm$, iRowCur&, i%, ws As Worksheet
    
    Call gloEt.ShowData(wsCur)
    
    With wsCur
        iRowCur = .Range("A1").End(xlDown).row
    
        
        .Range("A2:Z" & CStr(iRowCur)).Sort Key1:=.Range("C2"), Key2:=.Range("E2"), Key3:=.Range("F2"), Header:=xlNo
        
        .name = "对比结果"
    End With
    
    '删除多余表
    For Each ws In wbDes.Worksheets
        If InStr(ws.name, "sheet") > 0 Or InStr(ws.name, "Sheet") > 0 Then ws.Delete
    Next

    '删除完了，保存
    wbDes.SaveAs savePath
    ThisWorkbook.Activate
    wbDes.Activate
End Sub

Private Sub ReduceData()
    Dim i As Long, target_count As Long, i2 As Long
    Dim v As Variant, k As Variant, v2 As Variant
    
    target_count = 0
    v = d_data_kuaqi.items
    k = d_data_kuaqi.keys
    For i = 0 To UBound(v)
        target_count = target_count + v(i).count
    Next
    
    If target_count = 0 Then
        MsgBox "无结果显示"
        Exit Sub
    End If
    
    ReDim arr_data(1 To target_count + 1, 1 To 30) As Variant
    For i = 0 To UBound(v)
        target_count = target_count + v(i).count
    Next
    
    '标题设置
    Dim cur_row As Long, org_uni_code As String, arr_col As Variant, d_col As Object
    arr_col = Split(gloCft.GetConfig("大集中工具-数值核对-标题"), ",")
    Set d_col = CreateObject("scripting.dictionary")
    For i = 0 To UBound(arr_col)
        arr_data(1, i + 1) = arr_col(i)
        d_col(arr_col(i)) = i + 1
    Next
    arr_data(1, d_col("差异")) = arr_data(1, d_col("差异")) & "(" & DesUnit & ")"
    cur_row = 2
    For i = 0 To UBound(v)
        org_uni_code = k(i)
        v2 = v(i).items
        For i2 = 0 To UBound(v2)
            '最终输出
            ' 数据日期    机构类代码  机构类名称  地区代码    地区名称    校验编码    校验名称    左值    右值    差异    差异绝对值  差异幅度    是否说明    说明内容    计算公式
            If ycV(v2(i2)("left_data")) <> 0 Then arr_data(cur_row, d_col("左值")) = v2(i2)("left_data")
            If ycV(v2(i2)("right_data")) <> 0 Then arr_data(cur_row, d_col("右值")) = v2(i2)("right_data")
            If ycV(v2(i2)("差异")) <> 0 Then arr_data(cur_row, d_col("差异")) = v2(i2)("差异")
            If ycV(v2(i2)("差异绝对值")) <> 0 Then arr_data(cur_row, d_col("差异绝对值")) = v2(i2)("差异绝对值")
            If ycV(v2(i2)("差异幅度")) <> 0 Then arr_data(cur_row, d_col("差异幅度")) = v2(i2)("差异幅度")
            arr_data(cur_row, d_col("是否说明")) = v2(i2)("最终提示")
            
            arr_data(cur_row, d_col("校验类型")) = "数值核对"
            If ycV(v2(i2)("是否逻辑校验")) <> 0 Then arr_data(cur_row, d_col("校验类型")) = "逻辑校验"
            arr_data(cur_row, d_col("校验编码")) = v2(i2)("校验编码")
            arr_data(cur_row, d_col("校验名称")) = v2(i2)("校验名称")
            
            arr_data(cur_row, d_col("机构类代码")) = left(org_uni_code, 4)
            arr_data(cur_row, d_col("地区代码")) = right(org_uni_code, 7)
            
            
            
            arr_data(cur_row, d_col("计算公式")) = "left:" & CStr(v2(i2)("left")) & ";" & Chr(10) & "right:" & CStr(v2(i2)("right")) & ";" & Chr(10) & "校验公式:" & CStr(v2(i2)("校验公式")) & ";" & Chr(10) _
                                            & "left:" & CStr(v2(i2)("left_gc")) & ";" & Chr(10) & "right:" & CStr(v2(i2)("right_gc")) & ";"
            
            If d_org_ref.exists(left(org_uni_code, 4)) Then arr_data(cur_row, d_col("机构类名称")) = d_org_ref(left(org_uni_code, 4))
            If d_area_ref.exists(right(org_uni_code, 7)) Then arr_data(cur_row, d_col("地区名称")) = d_area_ref(right(org_uni_code, 7))
            cur_row = cur_row + 1
        Next
    Next
    '
    '   设置单元格格式
    '
    wsCur.Cells.WrapText = False
    wsCur.Cells.NumberFormatLocal = "@"
    wsCur.Cells.Font.Size = 10
    wsCur.Columns(d_col("左值")).NumberFormatLocal = "0.00"
    wsCur.Columns(d_col("右值")).NumberFormatLocal = "0.00"
    wsCur.Columns(d_col("差异绝对值")).NumberFormatLocal = "0.00"
    wsCur.Columns(d_col("差异幅度")).NumberFormatLocal = "0.00"
    wsCur.Columns(d_col("差异")).NumberFormatLocal = "0.00"
    wsCur.Range("A1").Resize(UBound(arr_data), UBound(arr_data, 2)).Value = arr_data
    wsCur.Cells.WrapText = False
End Sub

'
' 生成校验数据
'
Private Sub CaluData()
    '遍历校验指标，生成汇总数据
    Set d_data_kuaqi = Nothing
    Set d_data_kuaqi = CreateObject("scripting.dictionary")
    
    Dim arr_orgs As Variant, i As Long, j As Long, arr_rules As Variant, d_rule As Object, val_r As Variant
    Dim rule_code As String, reg As Object, left_data As Variant, right_data As Variant
    Dim d1 As Object, d_rule_keys As Variant, k As Long, t_calue As String
    Dim d_org_data As Object
    
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    arr_orgs = d_data.keys
    Set d_calue_guocheng = CreateObject("scripting.dictionary")
    '遍历机构数据
    For i = 0 To UBound(arr_orgs)
        'Debug.Print CStr(arr_orgs(i))
        
        Set d_org_data = Nothing
        Set d_org_data = CreateObject("scripting.dictionary")
        
        arr_rules = dKuaQiCheck.items
        '形成每个机构，每个校验的数据
        For j = 0 To UBound(arr_rules)
            d_calue_guocheng("left") = ""
            d_calue_guocheng("right") = ""
            t_calue = "" '保存计算结果
            '计算前提条件，满足前提条件才进行,形成 机构->数据(校验信息、left_data,right_data,差异，差异绝对值,，差异幅度，是否说明
            Set d_rule = arr_rules(j)
            If d_rule Is Nothing Then GoTo nextRule
            '判断机构类型和地区是否符合
            If d_rule("机构类代码") <> "" Then
                With reg
                    .Global = True: .ignorecase = False: .Pattern = d_rule("机构类代码")
                    If .test(left(CStr(arr_orgs(i)), 4)) = False Then GoTo nextRule
                End With
            End If
            If d_rule("地区代码") <> "" Then
                With reg
                    .Global = True: .ignorecase = False: .Pattern = d_rule("地区代码")
                    If .test(right(CStr(arr_orgs(i)), 7)) = False Then GoTo nextRule
                End With
            End If
            
            If CStr(d_rule("前提条件")) <> "" Then
                val_r = calu_rule(CStr(arr_orgs(i)), d_rule("前提条件"))
                If IsError(val_r) Then
                    t_calue = t_calue & "[前提条件]计算有误;"
                    Call gloDe.InsertNewDebug("跨期数值核对", left(CStr(arr_orgs(i)), 4), right(CStr(arr_orgs(i)), 7), "前提条件计算有误", d_rule("校验描述"), d_rule("前提条件"))
                    GoTo nextRule
                End If
                If CBool(val_r) = False Then GoTo nextRule  '不成立则跳过
            End If
            
            '计算左值
            left_data = calu_rule(CStr(arr_orgs(i)), d_rule("left"), "left")
            '重新思考，异常情况下如何处理，直接显示在结果里面吧
            If IsError(left_data) Then
                t_calue = t_calue & "[left_data]计算有误;"
                Call gloDe.InsertNewDebug("跨期数值核对", left(CStr(arr_orgs(i)), 4), right(CStr(arr_orgs(i)), 7), "left_data 计算有误", d_rule("校验描述"), d_rule("left"))
                GoTo nextRule
            End If
            left_data = ycV(left_data)
            
            '计算右值
            right_data = calu_rule(CStr(arr_orgs(i)), d_rule("right"), "right")
            If IsError(right_data) Then
                t_calue = t_calue & "[right_data]计算有误;"
                Call gloDe.InsertNewDebug("跨期数值核对", left(CStr(arr_orgs(i)), 4), right(CStr(arr_orgs(i)), 7), "right_data 计算有误", d_rule("校验描述"), d_rule("right"))
                GoTo nextRule
            End If
            right_data = ycV(right_data)
            
            '如果都没有数据的话，就没有什么好输出的
            If left_data <> 0 Or right_data <> 0 Then
                '计算公式
                val_r = calu_last_rule(left_data, right_data, d_rule("校验公式"))
                If IsError(val_r) Then
                    t_calue = t_calue & "[校验公式]计算有误;"
                    Call gloDe.InsertNewDebug("跨期数值核对", left(CStr(arr_orgs(i)), 4), right(CStr(arr_orgs(i)), 7), "校验公式 计算有误", d_rule("校验描述"), d_rule("校验公式"))
                Else
                    If CStr(d_rule("校验公式")) <> "" And CStr(d_rule("无误是否提示")) <> "" Then
                        ' 如果有校验公示，校验无误，且不显示
                        If (CStr(d_rule("取反标识")) <> "" And CBool(val_r) = True) Or (CStr(d_rule("取反标识")) = "" And CBool(val_r) = False) Then
                            GoTo nextRule
                        End If
                    End If
                End If
                
                Set d1 = Nothing
                Set d1 = CreateObject("scripting.dictionary")
                '先将校验复制一遍，冗余就冗余了
                d_rule_keys = d_rule.keys
                For k = 0 To UBound(d_rule_keys)
                    d1(CStr(d_rule_keys(k))) = d_rule(CStr(d_rule_keys(k)))
                Next
                d1("left_data") = left_data
                d1("right_data") = right_data
                d1("差异") = left_data - right_data
                d1("差异绝对值") = Abs(left_data - right_data)
                If left_data = 0 And right_data = 0 Then
                
                Else
                    d1("差异幅度") = Abs(left_data - right_data) / Application.WorksheetFunction.max(Abs(left_data), Abs(right_data)) * 100 '以大的为基数
                End If
                
                d1("计算结果") = val_r
                d1("left_gc") = CStr(d_calue_guocheng("left"))
                d1("right_gc") = CStr(d_calue_guocheng("right"))
                '最终显示的提示,如果是数值核对，则需要判断左右是否存在
                If (ycV(d_rule("是否逻辑校验")) = 0) And ((left_data = 0 And right_data <> 0) Or (left_data <> 0 And right_data = 0)) Then
                    t_calue = "一边有，一边无;"
                End If
                If IsError(val_r) Then
                    
                Else
                    If CStr(d_rule("校验公式")) <> "" Then
                        If (CStr(d_rule("取反标识")) <> "" And CBool(val_r) = False) Or (CStr(d_rule("取反标识")) = "" And CBool(val_r) = True) Then
                            t_calue = t_calue & d_rule("校验提示")
                        End If
                    End If
                End If
                d1("最终提示") = t_calue
            
                Set d_org_data(j) = d1
                Set d_data_kuaqi(CStr(arr_orgs(i))) = d_org_data
            End If
nextRule:
        Next
    Next
    Set d_calue_guocheng = Nothing
End Sub

Private Function calu_last_rule(left_data As Variant, rihgt_data As Variant, rule_code As String) As Variant
    On Error Resume Next
    If rule_code = "" Then
        calu_last_rule = True
        Exit Function
    End If
    
    ' 替换
    Dim tRule As String
    tRule = rule_code
    tRule = replace(tRule, "left", CStr(ycV(left_data)))
    tRule = replace(tRule, "right", CStr(ycV(rihgt_data)))
    
    calu_last_rule = Evaluate(tRule)  '这里是否需要对结果进行判断
    If err.Number <> 0 Then
        calu_last_rule = Evaluate("=vlook()")
    End If
    
End Function

'
'   计算校验结果
'
Private Function calu_rule(org_code As String, rule_code As String, Optional mode As String = "left") As Variant
    On Error Resume Next
    
    Dim reg As Object, mc As Object, m As Object, tStr As String, tRule As String, v_data As Double, result As Variant
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    tRule = rule_code
    ' 替换上期
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
    End With
    Set mc = reg.Execute(tRule)
    For Each m In mc
        tStr = calu_rule_change_target(CStr(m.Value))
        v_data = get_org_target_data(org_code, tStr, "pre") '前期对0赋值为0.01元，这里不需要修改
        tRule = replace(tRule, CStr(m.Value), v_data)
    Next
    
    ' 替换本期
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    Set mc = reg.Execute(tRule)
    For Each m In mc
        tStr = calu_rule_change_target(CStr(m.Value))
        v_data = get_org_target_data(org_code, tStr, "cur") '前期对0赋值为0.01元，这里不需要修改
        tRule = replace(tRule, CStr(m.Value), v_data)
    Next
    result = Evaluate(tRule)  '这里是否需要对结果进行判断
    If mode = "left" Then
        d_calue_guocheng("left") = tRule
    ElseIf mode = "right" Then
        d_calue_guocheng("right") = tRule
    Else
    
    End If
    
    If err.Number <> 0 Then
        result = Evaluate("=vlook()")
    End If
    calu_rule = result
End Function

Private Function calu_rule_change_target(rule_code As String) As String
    Dim arr As Variant, t As String, uni_code As String
    t = replace(rule_code, " ", "")
    t = replace(t, "[", ""): t = replace(t, "]", "")
    t = replace(t, "{", ""): t = replace(t, "}", "")
    arr = Split(t, ",")
    
    uni_code = CStr(arr(0)) & CStr(arr(1)) & CStr(arr(2)) & CStr(arr(3)) & CStr(arr(4))
    calu_rule_change_target = uni_code
End Function

'
'   根据机构代码、唯一标识获取数据
'
Private Function get_org_target_data(org_code As String, uni_code As String, Optional mode As String = "cur") As Double
    Dim d As Object
    If d_data.exists(org_code) Then
        Set d = d_data(org_code)
        If d.exists(uni_code) Then
            Set d = d(uni_code)
            
            If mode = "cur" Then
                get_org_target_data = d("cur")
                Exit Function
            Else
                get_org_target_data = d("pre")
                Exit Function
            End If
        End If
    End If
    
    get_org_target_data = 0
End Function



'
'  读取跨期校验
'
Sub InitKuaQiCheck()
    '
    Dim arr As Variant, i As Long, j As Long, d As Object, d_col As Object, ws As Worksheet
    Dim left As String, right As String, pre As String
    Set ws = ThisWorkbook.Worksheets("k复杂校验-跨期核对")
    ws.Cells.Interior.colorindex = 0
    arr = gloEt.GetUsedRange(ws).Value
    Set dKuaQiCheck = Nothing
    Set dKuaQiCheck = CreateObject("scripting.dictionary")
    Set d_col = Nothing
    Set d_col = CreateObject("scripting.dictionary")
    
    For j = 1 To UBound(arr, 2)
        d_col(CStr(arr(1, j))) = j
    Next
    
    For i = 2 To UBound(arr)
        ' 非禁用
        If CStr(arr(i, d_col("是否禁用"))) = "" Then
            left = InitComplexRef_ClearRule(CStr(arr(i, d_col("left"))))
            right = InitComplexRef_ClearRule(CStr(arr(i, d_col("right"))))
            pre = CStr(arr(i, d_col("前提条件")))
            If pre <> "" Then pre = InitComplexRef_ClearRule(pre)
            
            '判断是否涉及上下期
            If InStr(left, "{") > 0 And InStr(left, "}") > 0 And prePaths = "" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                GoTo nextCheck
            End If
            If InStr(right, "{") > 0 And InStr(right, "}") > 0 And prePaths = "" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                GoTo nextCheck
            End If
            If InStr(pre, "{") > 0 And InStr(pre, "}") > 0 And prePaths = "" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                GoTo nextCheck
            End If
            
            '检查校验规则是否完整，正确
            If left = "" Or right = "" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 15
                GoTo nextCheck
            End If
            
            ' 判断是否满足频度 和 规则是否填报完整
            If InitComplexRef_CheckFres_KuaQi(left, 4, 4) = False Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                GoTo nextCheck
            End If
            
            If InitComplexRef_CheckFres_KuaQi(right, 4, 4) = False Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                GoTo nextCheck
            End If
            
            If pre <> "" Then
                If InitComplexRef_CheckFres_KuaQi(pre, 4, 4) = False Then
                    ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 37
                    GoTo nextCheck
                End If
            End If
            
            '到这里就是正确的
            Set d = CreateObject("scripting.dictionary")
            For j = 1 To UBound(arr, 2)
                d(CStr(arr(1, j))) = CStr(arr(i, j))
            Next
            
            Set dKuaQiCheck(i) = d
            
        Else
            ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))).Interior.colorindex = 15
        End If
nextCheck:
    Next
End Sub

Private Function InitComplexRef_CheckFres_KuaQi(rule As String, splitCount As Integer, fre_index As Integer) As Boolean
    Dim reg As Object, mc As Object, m As Object, arr As Variant, i As Integer
    On Error Resume Next
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        arr = Split(replace(replace(CStr(m.Value), "{", ""), "}", ""), ",")
        If UBound(arr) <> 4 And UBound(arr) <> 0 And UBound(arr) <> -1 Then
            InitComplexRef_CheckFres_KuaQi = False
            Exit Function
        End If
        If Not dFres.exists(CStr(arr(fre_index - 1)) & CStr(arr(fre_index))) Then
            InitComplexRef_CheckFres_KuaQi = False
            Exit Function
        End If
        For i = 0 To UBound(arr)
            If CStr(arr(i)) = "" Then
                InitComplexRef_CheckFres_KuaQi = False
                Exit Function
            End If
        Next
    Next
    
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        arr = Split(replace(replace(CStr(m.Value), "[", ""), "]", ""), ",")
        If UBound(arr) <> 4 And UBound(arr) <> 0 And UBound(arr) <> -1 Then
            InitComplexRef_CheckFres_KuaQi = False
            Exit Function
        End If
        If Not dFres.exists(CStr(arr(fre_index - 1)) & CStr(arr(fre_index))) Then
            InitComplexRef_CheckFres_KuaQi = False
            Exit Function
        End If
        For i = 0 To UBound(arr)
            If CStr(arr(i)) = "" Then
                InitComplexRef_CheckFres_KuaQi = False
                Exit Function
            End If
        Next
    Next
    InitComplexRef_CheckFres_KuaQi = True
End Function


Sub InitCurData()
    Dim arr_file As Variant, arr_src As Variant, d1 As Object, d2 As Object
    Dim wsSrc As Worksheet
    Dim org_code As String, area_code As String, data_time As String, target_code As String, uni_code As String
    Dim i As Long, j As Long
    
    arr_file = Split(curPaths, ",")
    '遍历所有源数据
    For j = 0 To UBound(arr_file) Step 1
        Set wsSrc = gloEt.GetWb(CStr(arr_file(j))).Worksheets(1)
        'Set wsSrc = ThisWorkbook.Worksheets("test")
        arr_src = gloEt.GetUsedRange(wsSrc)
        arr_src = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(UBound(arr_src), UBound(arr_src, 2) + 10)).Value  '加大列范围，包括上期数据
        
        For i = 2 To UBound(arr_src)
            org_code = replace(CStr(arr_src(i, 3)), "'", "")
            area_code = replace(CStr(arr_src(i, 5)), "'", "")
            target_code = replace(CStr(arr_src(i, 8)), "'", "")
            data_time = ycDateToStr(arr_src(i, 2), "yyyy-mm-dd")
            dFres(CStr(arr_src(i, 12)) & CStr(arr_src(i, 13))) = 1  '频度 批次
            d_org_ref(org_code) = CStr(arr_src(i, 4))
            d_area_ref(area_code) = CStr(arr_src(i, 6))
            ' 判断是否存在
            If Not d_data.exists(org_code & area_code) Then
                Set d_data(org_code & area_code) = CreateObject("scripting.dictionary")
            End If
            Set d1 = d_data(org_code & area_code)
            Set d2 = CreateObject("scripting.dictionary")

            '增加指标数据,指标代码、数据属性、币种、频度、批次,5个
            uni_code = target_code & CStr(arr_src(i, 10)) & CStr(arr_src(i, 11)) & CStr(arr_src(i, 12)) & CStr(arr_src(i, 13))
            If (Not dUnitRef.exists(target_code)) And (IsAbnormalData(target_code, arr_src(i, 14), CStr(arr_file(j))) = False) Then
                ' 进行单位转换
                d2("cur") = ycV(arr_src(i, 14)) * unitC
            Else
                d2("cur") = ycV(arr_src(i, 14))
            End If
            If d2("cur") = 0 Then
                d2("cur") = 0.01 * GetUnitVal("元") / GetUnitVal(DesUnit)
            End If
            d2("pre") = 0
            Set d1(uni_code) = d2
            Set d_data(org_code & area_code) = d1
        Next

        ' 关闭源数据工作簿
        wsSrc.Parent.Close False
    Next

End Sub

Private Sub InitPreData()
    Dim arr_file As Variant, arr_src As Variant, d1 As Object, d2 As Object
    Dim wsSrc As Worksheet
    Dim org_code As String, area_code As String, data_time As String, target_code As String, uni_code As String
    Dim i As Long, j As Long
    
    arr_file = Split(prePaths, ",")
    
    '遍历所有源数据
    For j = 0 To UBound(arr_file) Step 1
        Set wsSrc = gloEt.GetWb(CStr(arr_file(j))).Worksheets(1)
        arr_src = gloEt.GetUsedRange(wsSrc)
        arr_src = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(UBound(arr_src), UBound(arr_src, 2) + 10)).Value  '加大列范围，包括上期数据
        
        For i = 2 To UBound(arr_src)
            org_code = replace(CStr(arr_src(i, 3)), "'", "")
            area_code = replace(CStr(arr_src(i, 5)), "'", "")
            target_code = replace(CStr(arr_src(i, 8)), "'", "")
            data_time = ycDateToStr(arr_src(i, 2), "yyyy-mm-dd")
            dFres(CStr(arr_src(i, 12)) & CStr(arr_src(i, 13))) = 1  '频度 批次
            d_org_ref(org_code) = CStr(arr_src(i, 4))
            d_area_ref(area_code) = CStr(arr_src(i, 6))
            ' 判断是否存在
            If Not d_data.exists(org_code & area_code) Then
                Set d_data(org_code & area_code) = CreateObject("scripting.dictionary")
            End If
            Set d1 = d_data(org_code & area_code)
            'Set d2 = CreateObject("scripting.dictionary")

            '增加指标数据,指标代码、数据属性、币种、频度、批次,5个
            uni_code = target_code & CStr(arr_src(i, 10)) & CStr(arr_src(i, 11)) & CStr(arr_src(i, 12)) & CStr(arr_src(i, 13))
            If Not d1.exists(uni_code) Then
                Set d1(uni_code) = CreateObject("scripting.dictionary")
                d1(uni_code)("cur") = 0
                d1(uni_code)("pre") = 0
            End If
            Set d2 = d1(uni_code)
            If (Not dUnitRef.exists(target_code)) And (IsAbnormalData(target_code, arr_src(i, 14), CStr(arr_file(j))) = False) Then
                ' 进行单位转换
                d2("pre") = ycV(arr_src(i, 14)) * unitC
            Else
                d2("pre") = arr_src(i, 14)
            End If
            If d2("pre") = 0 Then
                d2("pre") = 0.01 * GetUnitVal("元") / GetUnitVal(DesUnit)
            End If
            Set d1(uni_code) = d2
            Set d_data(org_code & area_code) = d1
        Next

        ' 关闭源数据工作簿
        wsSrc.Parent.Close False
    Next
End Sub

Private Sub InitUnitRef(wsDes As Worksheet)
    Dim arrUnitRef() As Variant, i%, ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("单位不需要转换指标参照表")
    Call gloEt.ShowData(ws)
    Set dUnitRef = CreateObject("scripting.dictionary")
    
    With ws
        arrUnitRef = .Range("A1:C" & CStr(.Range("B1").End(xlDown).row))
    End With
    For i = 2 To UBound(arrUnitRef) Step 1
        dUnitRef(replace(CStr(arrUnitRef(i, 2)), "'", "")) = 0
    Next i
    
    wsDes.Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)).NumberFormatLocal = "@"
    wsDes.Range("A1").Resize(UBound(arrUnitRef), UBound(arrUnitRef, 2)) = arrUnitRef
    wsDes.name = ws.name
    wsDes.Visible = xlSheetVeryHidden
End Sub

Private Function InitSomething()
    ' 建立新的工作簿
    Set wbDes = Workbooks.add
    With wbDes
        Do While wbDes.Worksheets.count > 1
            wbDes.Worksheets(1).Delete
        Loop
    End With
    Set wsCur = wbDes.Worksheets(1)
    wsCur.name = "对比结果"
    
    '复制【单位不需要转换指标参照表】表，及初始化
    Call InitUnitRef(wbDes.Worksheets.add)
    '复制【机构参照表】
    Call InitOrgRef(wbDes.Worksheets.add)
    
    Set dFres = Nothing
    Set dFres = CreateObject("scripting.dictionary")
    Set d_data = Nothing
    Set d_data = CreateObject("scripting.dictionary")
    Set d_org_ref = Nothing '机构类参照
    Set d_org_ref = CreateObject("scripting.dictionary")
    Set d_area_ref = Nothing '地区参照
    Set d_area_ref = CreateObject("scripting.dictionary")

    '读取数据
    Call InitCurData
    
    If prePaths <> "" Then Call InitPreData
    ' 初始化跨期校验
    Call InitKuaQiCheck
End Function
