Attribute VB_Name = "基础数据季报表跨期比较"
Option Explicit

Private Type 警戒值                  '定义一个新数据类型记录
    序号 As Integer
    变幅 As Double
    备注文字 As String
    备注颜色 As Integer
    是否整行填充 As Integer
End Type

Private Type SpecialTaget
    text As String
    rmbF As Double
    usdF As Double
    dataType As String   '数据属性，用于判断是发生额还是余额
    explain As String
    TableName As String
    detailed As String
    bf As Double '绝对值变幅
End Type

Private ComplexThd As Double
Private Const ComplexTaget As Integer = 5
Private Type ComplexTaget
    table As String
    ruleEx As String
    rule As String
    qfbz As Boolean
    thdV As Double
End Type

Private Const OtherTaget As Integer = 5
Private Type OtherTaget
    code As String
    name As String
    qfbz As Boolean
End Type

Private Const 是否不调试 As Boolean = True   '用于去公式格式话的
Private Const dt As Boolean = True  '测量时间
Private Const dbl_Epsilon As Double = 2.22044604925031E-13

Private Const 辅助列 As Integer = 30 '辅助列 用 30列好了。

Private dDataCol As Object  '列字典

Private col_社会信用代码 As Integer
Private col_数据日期 As Integer
Private col_机构名称 As Integer
Private col_地区 As Integer
Private col_机构类别 As Integer
Private col_是否重点机构 As Integer
Private col_表单号 As Integer
Private col_指标编码 As Integer
Private col_是否重点指标 As Integer
Private col_指标信息 As Integer
Private col_指标名称 As Integer
Private col_数据属性 As Integer
Private col_频度 As Integer
Private col_当期值 As Integer
Private col_币种列 As Integer

Private 上期值列 As Integer
Private 增减额列 As Integer
Private 环比列 As Integer
Private 备注列 As Integer
Private 是否说明列 As Integer
Private 计算过程列 As Integer
Private 说明内容列 As Integer


Private savePath As String  '最终存放路径

Private CurDate As Date
Private PreDate As Date

Private CurPaths As Variant
Private PrePaths As Variant

Private SforbidCol As Integer    '禁用标记列
Private CforbidCol As Integer    '禁用标记列

' 输出工作簿对象
Private wbDes As Workbook   '输出的工作簿
Private wsCur As Worksheet  '输出的工作表

Private dUnitRef As Object '单位转换参照表  字典
Private dSpcRef As Object  '特殊指标表  字典
Private dComplexRef As Object   '复杂指标表 字典
Private dCurDataRef As Object  '当期表  字典
Private dPreDataRef As Object  '上期表  字典
Private d_data As Object '总数据 字典
Private arrCur() As Variant '当期表数组

Private SrcUnit As String   '源数据单位
Private DesUnit As String   '目标数据单位
Private unitC As Double '数值转换
Private 分析文件数据单位 As String
Private Const 特殊指标数据单位 As String = "亿元"

Private dComplexChecked As Object   '用于存储已校验过的规则

Private dtarget_def As Object '指标参照
Private dorg_def As Object '机构参照
Private dorg_name_def As Object '机构参照 按机构名称
Private darr_target As Object '存放指标代码 指标代码、指标名称、行位置、列位置


Private arrAlert() As 警戒值    '警戒值数组
Private Complex_Soft_Color_Index As Integer '软性校验颜色

Private Const COMPLEX_TARGET_INDEX As Integer = 3  '复杂校验指标代码的索引位置
Private Const COMPLEX_FRE_INDEX As Integer = 5
Private Const COMPLEXT_SPLIT_COUNT As Integer = 4  '复杂校验指标的分隔符数量

Private IsFres As Boolean
Private IsCheckTol As Boolean   '是否启用总伐值校验选项
Private CheckTolModle As Integer    '总阀值校验模式 1、2、0
Private bCheckTol_Value As Double   '绝对值伐值
Private bCheckTol_Value2 As Double   '绝对值伐值
Private bCheckTol_Value3 As Double   '绝对值伐值
Private bCheckTol_Range As Double   '绝对值幅度伐值

Private arrComplexeKeys As Variant
Private IsHzFTol As Boolean   '是否启用分行业汇总
Private IsOnlyOutputHzF As Boolean   '是否仅输出汇总
Private Const FSE_PIANCHA As Double = 0.00001
Private Const STR_HZ_EZ As String = "0_汇总_"

Private bBaiBeiyuzhi As Double
Private d_cur_orgs As Object '当期机构
Private d_not_orgs As Object '当期没有的机构
Private msg_not_orgs As String

Public Sub 基础数据与大集中核对()
    Dim arr As Variant, d_djz As Object, d_js As Object, d_orgs As Object

    Dim ws As Worksheet
    Dim k As Variant, v As Variant
    Dim i As Long, j As Long
    
    Set ws = gloEt.FindWs(ThisWorkbook, "大集中报表查询", 0, True)
    arr = gloEt.GetUsedRange(ws).value
    Set d_djz = CreateObject("scripting.dictionary")
    Set d_js = CreateObject("scripting.dictionary")
    Set d_orgs = CreateObject("scripting.dictionary")
    Set d_data = Nothing
    Set d_data = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr)
        For j = 9 To UBound(arr, 2)
            d_djz(CStr(arr(i, 8)) & CStr(arr(1, j))) = ycV(arr(i, j))
        Next
    Next
    
    For i = 2 To UBound(arrCur)
        d_orgs(CStr(arrCur(i, 3))) = 1
        If CStr(arrCur(i, 6)) = "是" Then
            d_js(CStr(arrCur(i, 3)) & CStr(arrCur(i, 4))) = ycV(arrCur(i, 7))
        End If
    Next
    
    Dim org_name As String, org_code As String, d_t As Object, key As String, max_v As Double
    Dim target_code As String
    
    Set ws = gloEt.FindWs(ThisWorkbook, "指标参照", 0, True)
    arr = gloEt.GetUsedRange(ws).value
    k = d_orgs.keys
    For i = 2 To UBound(arr)
        If CStr(arr(i, 11)) = "是" Then
            target_code = CStr(arr(i, 1))
            For j = 0 To UBound(k)
                org_name = k(j)
                org_code = Find_Org_Pro(org_name, "机构代码")
                
                key = org_name & target_code
                
                Set d_t = CreateObject("scripting.dictionary")
                d_t("数据属性") = CStr(Find_Target_Pro(target_code, "数据属性"))
                
                d_t("基础数据值") = ""
                If d_js.exists(key) Then
                    d_t("基础数据值") = d_js(key)
                    If Round(ycV(d_t("基础数据值")), 6) = 0 Then d_t("基础数据值") = ""
                End If
                
                key = Find_Org_Pro(org_code, "报表项目") & CStr(Find_Target_Pro(target_code, "大集中报表查询指标名称"))
                d_t("大集中值") = ""
                If d_djz.exists(key) Then
                    d_t("大集中值") = d_djz(key)
                    If CStr(Find_Target_Pro(target_code, "是否不转换单位")) = "" Then
                        d_t("大集中值") = d_djz(key) * unitC
                    End If
                    If Round(ycV(d_t("大集中值")), 6) = 0 Then d_t("大集中值") = ""
                End If
                
                d_t("地区") = Find_Org_Pro(org_code, "归属行")
                d_t("机构名称") = org_name
                d_t("指标编码") = target_code
                d_t("指标名称") = CStr(Find_Target_Pro(target_code, "指标名称"))
                d_t("是否与大集中核对") = "是"
                
                
                If CStr(d_t("大集中值")) = "" And CStr(d_t("基础数据值")) = "" Then
                
                ElseIf Round(ycV(d_t("大集中值")), 6) = 0 And Round(ycV(d_t("基础数据值"))) = 0 Then
                
                Else
                    d_t("差异绝对值") = Abs(ycV(d_t("基础数据值")) - ycV(d_t("大集中值")))
                    max_v = Abs(ycV(d_t("基础数据值")))
                    If max_v < Abs(ycV(d_t("大集中值"))) Then max_v = Abs(ycV(d_t("大集中值")))
                    If max_v <> 0 Then
                        d_t("差异幅度") = ycV(d_t("差异绝对值")) / max_v * 100
                    Else
                        d_t("是否说明") = "除数为0;"
                    End If
                    
                    If CStr(d_t("大集中值")) = "" Or CStr(d_t("基础数据值")) = "" Then
                        d_t("是否说明") = d_t("是否说明") & "一边有，一边无，谨慎核实;"
                    ElseIf d_t("差异绝对值") > 0.01 Then
                        d_t("是否说明") = d_t("是否说明") & "差异超过100元;"
                    End If
                    Set d_data(org_name & target_code) = d_t
                End If
                
                
                
                
                
                Set d_t = Nothing
            Next
        End If
    Next
    
    If d_data.count = 0 Then
        Exit Sub
    End If
    
    Call gloCft.ShowStatusMessage_Up("双向比较数据读取完毕，转换为数组。")
    
    ReDim arrCur(1 To 1 + d_data.count, 1 To 12) As Variant

    arrCur(1, 1) = "地区": arrCur(1, 2) = "数据属性": arrCur(1, 3) = "机构名称": arrCur(1, 4) = "指标编码": arrCur(1, 5) = "指标名称"
    arrCur(1, 6) = "是否与大集中核对": arrCur(1, 7) = "基础数据值": arrCur(1, 8) = "大集中值": arrCur(1, 9) = "差异绝对值": arrCur(1, 10) = "差异幅度"
    arrCur(1, 11) = "变动原因/变动说明": arrCur(1, 12) = "是否说明"
    
    Set dDataCol = CreateObject("scripting.dictionary")
    For i = 1 To UBound(arrCur, 2)
        dDataCol(CStr(arrCur(1, i))) = i
    Next
    
    
    
    i = 2 '
    For Each v In d_data.items
        For Each k In v.keys
            arrCur(i, dDataCol(k)) = v(k)  ' 填充对应指标的值
        Next
        i = i + 1
        
        If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行数组转换，执行到第" & CStr(i) & "行。")
    Next
    Dim wb As Workbook

    Set wsCur = wsCur.Parent.Worksheets.add
    wsCur.name = "大集中比较"
    Call OutputExcel(wsCur, arrCur)
    
    Set d_data = Nothing
    Set d_djz = Nothing
    Set d_js = Nothing
    Set d_orgs = Nothing
    Erase arr
End Sub


Public Sub 基础数据系统报表跨期比较过程()
    Call EraseSomthing
    Call InitMBcz
    CurPaths = gloFt.SelectFiles("选择当期数文件")
    If CurPaths = "" Then
        MsgBox "当期数文件未选择"
        Exit Sub
    End If
    
    PrePaths = gloFt.SelectFiles("选择上期数文件")
    If PrePaths = "" Then
        MsgBox "上期数文件未选择"
        Exit Sub
    End If
    
    savePath = gloFt.GetOutputFile()
    If savePath = "" Then
        MsgBox "输出文件未选择"
        Exit Sub
    End If
    
    If InitSomething_Date() <> "" Then
        MsgBox "数据日期未设置正确。"
        Exit Sub
    End If

    On Error GoTo errhandle
    Dim rstr As String
    Call 警戒区间校验
    
    Complex_Soft_Color_Index = gloCft.GetConfig("复杂校验软性颜色")
    
    
    CforbidCol = 5
    Call gloCft.CloseSomething
    
    Call gloCft.ShowStatusMessage_Up("文件选择结束，准备初始化数据")
    gloT = Timer
    
    rstr = InitSomething()
    
    If rstr <> "" Then
        MsgBox rstr
        Call EraseSomthing
        Exit Sub
    End If
    
    Call gloCft.ShowStatusMessage_Up("数据初始化结束，进行校验")
    Call DoCompare
    
    Dim col As Integer, j As Integer, new_rows As Long, new_arr() As Variant, i As Double
    col = 15
    If UBound(arrCur) > 0 Then
        col = UBound(arrCur, 2)
        
        '是否仅保留有提示的
        If CStr(glo_dcz("是否仅保留有提示数据")) = "是" Then
            '
            new_rows = 0
            For i = 2 To UBound(arrCur)
                If CStr(arrCur(i, 是否说明列)) <> "" Then
                    new_rows = new_rows + 1
                End If
            Next
            ReDim new_arr(1 To new_rows + 1, 1 To col) As Variant
            For j = 1 To col
                new_arr(1, j) = arrCur(1, j)
            Next
            new_rows = 2
            For i = 2 To UBound(arrCur)
                If CStr(arrCur(i, 是否说明列)) <> "" Then
                    For j = 1 To col
                        new_arr(new_rows, j) = arrCur(i, j)
                    Next
                    new_rows = new_rows + 1
                End If
                If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
            Next
            arrCur = new_arr
            Erase new_arr
        End If
        
        Call OutputExcel(wsCur, arrCur)
        
        ' 这里再增加一个，与大集中比较
        Call 基础数据与大集中核对
        
        Call AddMacroCode
        
        
        
        wbDes.SaveAs savePath
        If d_not_orgs.count > 0 Then
            Dim k As Variant
            For Each k In d_not_orgs.keys
                Call gloDe.InsertNewDebug("基础数据跨期比较", "-", "-", "初始化-机构本期数据缺失", CStr(k))
            Next
            MsgBox "运行结束，共耗时:" & CStr(Timer - gloT) & "秒。其中：存在" & d_not_orgs.count & "家机构本期无数，具体在【输出日志】查看。"
        Else
            MsgBox "运行结束，共耗时:" & CStr(Timer - gloT) & "秒。"
        End If
        
    Else
        wbDes.Close False
        MsgBox "运行结束，无相关结果。共耗时:" & CStr(Timer - gloT) & "秒。"
    End If
    
    Call EraseSomthing
    Call gloCft.OpenSomething
    Call gloCft.ShowStatusMessage_Up("")
Exit Sub

errhandle:
    
    MsgBox "手工表跨期校验出现错误" & err.Description
    Call gloCft.ShowStatusMessage_Up("")
    Call gloCft.OpenSomething
    Call EraseSomthing
End Sub

Private Sub EraseSomthing()
    Set dSpcRef = Nothing
    Set dCurDataRef = Nothing
    Set dPreDataRef = Nothing
    Set dUnitRef = Nothing
    Set dComplexRef = Nothing
    Set dComplexChecked = Nothing
    Set d_cur_orgs = Nothing
    Set d_not_orgs = Nothing
    Erase arrCur
    Erase arrAlert
End Sub


'
'
'
Private Sub DoCompare()
    'On Error GoTo CompareErr
    Dim rnDes As Range
    Dim iRowCur&, i&, j&, k&, str$, nr As Long, strNr As String
    Dim data_type As String
    
    iRowCur = UBound(arrCur)    '获取当期数据行数
    
    IsCheckTol = False
    If CStr(glo_dcz("是否启用总阈值校验")) = "是" Then IsCheckTol = True
    
    CheckTolModle = 1
    If InStr(CStr(glo_dcz("是否启用总阈值校验")), "或") > 0 Then CheckTolModle = 2
    
    bCheckTol_Value = 5000
    bCheckTol_Value = ycV(glo_dcz("绝对值变动阈值（万元）"))
    bCheckTol_Value2 = 100
    bCheckTol_Value2 = ycV(glo_dcz("绝对值变动阈值（只/个）"))
    bCheckTol_Range = 30
    bCheckTol_Range = ycV(glo_dcz("绝对值变幅阈值"))
    bCheckTol_Value3 = 10
    bCheckTol_Value3 = ycV(glo_dcz("绝对值变动阈值（%）"))
    
    bBaiBeiyuzhi = GetUnitVal("万元") / GetUnitVal(DesUnit) * ycV(glo_dcz("100倍变动绝对值阈值(万元)"))
    
    
    '遍历所有数据，处理特殊指标
    For i = 2 To UBound(arrCur) Step 1

        ' 特殊指标校验
        If dSpcRef.exists(replace(arrCur(i, col_指标编码), "'", "")) Then
            If CStr(arrCur(i, 是否说明列)) = "" Then
                arrCur(i, 是否说明列) = DoCompare_AddMyExplain(replace(arrCur(i, col_指标编码), "'", ""), i)
            Else
                arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & "；" & DoCompare_AddMyExplain(replace(arrCur(i, col_指标编码), "'", ""), i)
            End If
        End If
        
        ' 缩小100倍及以上 ,增幅100倍以上
        If CStr(arrCur(i, 备注列)) <> "" Then
            If InStr(CStr(arrCur(i, 备注列)), "缩小100倍及以上") > 0 Or InStr(CStr(arrCur(i, 备注列)), "增幅100倍以上") > 0 Then
                If Abs(ycV(arrCur(i, 增减额列))) > bBaiBeiyuzhi Then
                    If CStr(arrCur(i, 是否说明列)) <> "" Then
                        arrCur(i, 是否说明列) = CStr(arrCur(i, 是否说明列)) & ";" & CStr(arrCur(i, 备注列)) & "。"
                    Else
                        arrCur(i, 是否说明列) = CStr(arrCur(i, 备注列)) & "。"
                    End If
                End If
            End If
        End If
        
        If dComplexRef.exists(replace(arrCur(i, col_指标编码), "'", "")) Then
            '直接挪到里面进行操作吧
            'arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & " " & DoCompare_AddMyComplexExplain(replace(arrCur(i, 11), "'", ""), i)
            Call DoCompare_AddMyComplexExplain(replace(arrCur(i, col_指标编码), "'", ""), i)
        End If
        
        ' 增加总阈值校验
        data_type = CStr(arrCur(i, col_数据属性))
        If data_type <> "文字" Then
            Call Docompare_CheckTol(Abs(ycV(arrCur(i, 增减额列))), Abs(ycV(arrCur(i, 环比列))), i, data_type)
        End If
        If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在校验第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
    Next i
    
    
        
    
Exit Sub

CompareErr:
    'Resume
    Call EraseSomthing
    str = "DoCompare:执行有误," & err.Description
    err.Clear
    err.Raise Number:=1002, Description:=str
End Sub


'
'   做整体比较
'   已经是绝对值了
'
Private Sub Docompare_CheckTol(绝对值变化 As Double, 绝对值变幅 As Double, i As Long, data_type As String)
    Dim str As String, strSt As String, u As Integer, cv As Double
    Dim unit_str As String
    
    unit_str = CStr(DesUnit)
    If data_type = "余额" Or data_type = "累发" Then
        cv = bCheckTol_Value
    ElseIf data_type = "个数" Then
        cv = bCheckTol_Value2
        unit_str = "只/个"
    ElseIf data_type = "百分数" Then
        cv = bCheckTol_Value3
        unit_str = "%"
    Else
        cv = bCheckTol_Value
    End If
    str = ""
    
    If bCheckTol_Value <> 0 And bCheckTol_Range <> 0 Then
        If CheckTolModle = 1 Then
            If 绝对值变化 >= cv And 绝对值变幅 >= bCheckTol_Range Then
                str = "总伐值校验：绝对值变化超过" & CStr(cv) & CStr(unit_str) & ",且绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
            End If
        ElseIf CheckTolModle = 2 Then
            If 绝对值变化 >= cv Or 绝对值变幅 >= bCheckTol_Range Then
                str = "总伐值校验：绝对值变化超过" & CStr(cv) & CStr(unit_str) & ",或绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
            End If
        Else
        
        End If
    ElseIf bCheckTol_Value <> 0 Then
        If 绝对值变化 >= cv Then
            str = "总伐值校验：绝对值变化超过" & CStr(cv) & unit_str
        End If
    ElseIf bCheckTol_Range <> 0 Then
        If 绝对值变幅 >= bCheckTol_Range Then
            str = "总伐值校验：绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
        End If
    Else
    
    End If
    
    arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & str
End Sub


Private Function InitSomething() As String
    
    '初始化输出工作簿
    Set wbDes = Workbooks.add
    With wbDes
        Do While wbDes.Worksheets.count > 1
            wbDes.Worksheets(1).Delete
        Loop
    End With
    Set wsCur = wbDes.Worksheets(1)
    wsCur.name = "说明"
    With wbDes.Worksheets.add
        .name = "宏代码"
        .Visible = xlSheetHidden
        .Cells(1, 1).value = gloCft.GetConfig("说明文件菜单模块宏代码")
        .Cells(2, 1).value = gloCft.GetConfig("说明文件工作模块宏代码")
        .Cells(3, 1).value = gloCft.GetConfig("输出说明通用宏")
    End With
    
    
    ' 对输出工作簿的初始化结束
    
    
    
    SrcUnit = "元"   '源数据单位
    DesUnit = "万元"   '目标数据单位
    unitC = GetUnitVal(SrcUnit) / GetUnitVal(DesUnit)  '数值转换
    分析文件数据单位 = "万元"
    
    Call Init_Data_def
    
    Dim rstr As String
    rstr = ""
    'rstr = rstr & InitSomething_Date()
    rstr = rstr & InitSomething_Data_gddj()
    rstr = rstr & InitSpecialRef()
    rstr = rstr & InitComplexRef()
    
    
    InitSomething = rstr
End Function

Private Function InitSomething_Date() As String
    On Error GoTo errhandle
    
    CurDate = CDate(glo_dcz("当期数日期"))
    PreDate = CDate(glo_dcz("上期数日期"))
    
    InitSomething_Date = ""
Exit Function
errhandle:
    InitSomething_Date = "初始化日期有误:" & CStr(err.Description) & ";"
End Function

Private Sub InitSomething_Data_Col()
    On Error Resume Next
    col_社会信用代码 = Int(ycV(dDataCol("社会信用代码")))
    col_数据日期 = Int(ycV(dDataCol("数据日期")))
    col_机构名称 = Int(ycV(dDataCol("机构名称"))) '机构名称
    col_地区 = Int(ycV(dDataCol("地区"))) ' 地区
    col_机构类别 = Int(ycV(dDataCol("机构类别")))
    
    'col_是否重点机构 = Int(ycV(dDataCol("社会信用代码")))
    'col_表单号 =
    
    col_指标编码 = Int(ycV(dDataCol("指标编码"))) '指标编码
    col_是否重点指标 = Int(ycV(dDataCol("是否与大集中核对"))) '是否重点关注指标
    col_指标名称 = Int(ycV(dDataCol("指标名称"))) '指标名称
    col_数据属性 = Int(ycV(dDataCol("数据属性"))) '数据属性
    col_频度 = Int(ycV(dDataCol("频度"))) '
    col_当期值 = Int(ycV(dDataCol("当期数"))) '当期数
    col_币种列 = Int(ycV(dDataCol("币种"))) 'col_币种列
    
    上期值列 = Int(ycV(dDataCol("上期数"))) '上期数
    增减额列 = Int(ycV(dDataCol("变动绝对值"))) '变动绝对值
    环比列 = Int(ycV(dDataCol("环比变动"))) '环比变动
    备注列 = Int(ycV(dDataCol("备注")))
    是否说明列 = Int(ycV(dDataCol("是否说明"))) '备注
    计算过程列 = Int(ycV(dDataCol("计算过程"))) '计算过程
    说明内容列 = Int(ycV(dDataCol("变动原因/变动说明"))) '变动原因/变动说明
    '是否说明列 = Int(ycV(dDataCol("是否说明列")))
End Sub


'用来，取分母数据
Private Function get_org_data(d As Variant, org_name As String, keys As String, k As String) As Double
    Dim arr_k As Variant, i As Integer, rel As Double
    arr_k = Split(keys, "+")
    rel = 0
    For i = 0 To UBound(arr_k)
        rel = rel + dgv(d, org_name & CStr(arr_k(i)), k)
    Next
    get_org_data = rel
End Function

Private Function dgv(d As Variant, key As Variant, k As String) As Double
    On Error Resume Next
    Dim dt As Object
    If d.exists(key) Then
        Set dt = d(key)
        dgv = ycV(dt(k))
        
        If err.Number <> 0 Then
            dgv = 0
            err.Clear
        End If
    Else
        dgv = 0
    End If
End Function



'
'    初始化基础数据报表
'
Private Function InitSomething_Data_gddj() As String
    On Error GoTo errHandel
    Dim fns As Variant, fn As String
    Dim wb_src As Workbook, ws_src As Worksheet, i As Long, j As Long, ws_t As Worksheet, arr As Variant, y As Long
    Dim d_data As Object, d_t As Object, org_name As String, org_code As String, target_code As String, data_value As Variant, data_type As String
    Dim ws_err As Worksheet, arr_err As Variant, key As String
    Dim org_type As String, data_err_explain As String
    Dim d_jg As Object, fz As Double, fm As Double, fm2 As Double, isdpre As Boolean, hz_count As Long
    Dim arr_fn As Variant
    Set d_jg = CreateObject("scripting.dictionary")
    IsHzFTol = False
    If CStr(glo_dcz("是否启用长表跨期汇总")) = "是" Then IsHzFTol = True
    IsOnlyOutputHzF = False
    If CStr(glo_dcz("是否仅输出汇总数据")) = "是" Then IsOnlyOutputHzF = True
    hz_count = 0
    isdpre = True
    If CStr(glo_dcz("是否当期为空的时候不读取上期数")) = "是" Then isdpre = False
    '是否当期为空的时候不读取上期数
    
    If IsHzFTol = True Then
        ' 初始化 加权平均参照 表
        Set ws_t = ThisWorkbook.Worksheets("加权平均参照")
        arr = gloEt.GetUsedRange(ws_t).value
        For i = 2 To UBound(arr)
            d_jg(CStr(arr(i, 1))) = CStr(arr(i, 4))
        Next
    End If
    '构造数据字典
    Set d_data = CreateObject("scripting.dictionary")
    Set dCurDataRef = CreateObject("scripting.dictionary")  '当期表  字典 不用复杂校验的话，是不用的
    Set dPreDataRef = CreateObject("scripting.dictionary")  '上期表  字典
    
    Set d_cur_orgs = CreateObject("scripting.dictionary")
    Set d_not_orgs = CreateObject("scripting.dictionary")
    
    ' 读取本期数据
    fn = CStr(CurPaths)
    fns = Split(fn, ",")
    For y = 0 To UBound(fns)
        fn = fns(y)
        Set wb_src = gloEt.GetWb(fn)
        ' 通过文件可以判断是否是相关文件 91441300055343051Q#2025-12-31#01#20201#惠州农村商业银行股份有限公司
        arr_fn = Split(gloFt.getShortFileName(fn), "#")
        If UBound(arr_fn) = 4 Then
            ' 判断名称是否在指定的
            
            For Each ws_src In wb_src.Worksheets
                If InStr(ws_src.name, "2020") <= 0 Then
                    
                Else
                    org_name = CStr(arr_fn(4))
                    org_code = CStr(arr_fn(0))
                    If Not d_cur_orgs.exists(org_name) Then
                        d_cur_orgs(org_name) = 1
                    End If
                    '读取
                    arr = gloEt.GetUsedRange(ws_src).value
                    
                    If CurDate <> ycD(CStr(arr_fn(1))) Then
                        InitSomething_Data_gddj = "初始化原始数据有误:文件当期日期为:" & CStr(arr_fn(1)) & ",与初始设置不一致;"
                        Exit Function
                    End If
                    
                    For i = 4 To UBound(arr)
                        target_code = CStr(arr(i, 1))
                        data_type = Find_Target_Pro(target_code, "数据属性")
                        data_value = CStr(arr(i, 3))

                        If (data_type = "文字" And data_value <> "") Or (data_type <> "文字" And ycV(data_value) <> 0) Then
                           Set d_t = CreateObject("scripting.dictionary")
                           d_t("数据属性") = data_type
                           d_t("当期数") = data_value
                           d_t("机构类别") = Find_Org_Pro(org_code, "机构类别")
                           If data_type = "文字" Then d_t("当期数") = "'" & replace(CStr(data_value), "'", "")
                           If data_type <> "文字" And CStr(Find_Target_Pro(target_code, "是否不转换单位")) = "" Then
                                d_t("当期数") = ycV(data_value) * unitC
                           End If
                           d_t("上期数") = ""
                           d_t("机构名称") = org_name
                           d_t("指标编码") = target_code
                           d_t("指标名称") = CStr(arr(i, 2))
                           d_t("变动原因/变动说明") = ""
                           Set d_data(org_name & target_code) = d_t
                           Set d_t = Nothing
                           
                        End If
                    Next
                        
                    Set ws_src = Nothing
                    wb_src.Close False
                    Set wb_src = Nothing
                    Erase arr
                End If
            Next
        End If
        If y Mod 5 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理当期数第：" & CStr(y + 1) & "个/" & CStr(UBound(fns) + 1) & "个")
    Next
    
    ' 读取上期数据
    fn = CStr(PrePaths)
    fns = Split(fn, ",")
     For y = 0 To UBound(fns)
        fn = fns(y)
        Set wb_src = gloEt.GetWb(fn)
        ' 通过文件可以判断是否是相关文件 91441300055343051Q#2025-12-31#01#20201#惠州农村商业银行股份有限公司
        arr_fn = Split(gloFt.getShortFileName(fn), "#")
        If UBound(arr_fn) = 4 Then
            ' 判断名称是否在指定的
            
            For Each ws_src In wb_src.Worksheets
                If InStr(ws_src.name, "2020") <= 0 Then
                    
                Else
                    org_name = CStr(arr_fn(4))
                    org_code = CStr(arr_fn(0))
                    If Not d_cur_orgs.exists(org_name) Then
                        d_cur_orgs(org_name) = 1
                    End If
                    '读取
                    arr = gloEt.GetUsedRange(ws_src).value
                    If PreDate <> ycD(CStr(arr_fn(1))) Then
                        InitSomething_Data_gddj = "初始化原始数据有误:文件上期日期为:" & CStr(arr_fn(1)) & ",与初始设置不一致;"
                        Exit Function
                    End If
            
                    If d_cur_orgs.exists(org_name) Or isdpre = True Then
                        
                        For i = 4 To UBound(arr)
                            target_code = CStr(arr(i, 1))
                            data_type = Find_Target_Pro(target_code, "数据属性")
                            data_value = CStr(arr(i, 3))
                            '存在则补充上期数，不存在则新增一个
                            
                            If (data_type = "文字" And data_value <> "") Or (data_type <> "文字" And ycV(data_value) <> 0) Then
                                If d_data.exists(org_name & target_code) Then
                                    Set d_t = d_data(org_name & target_code)
                                    d_t("上期数") = data_value
                                    If data_type = "文字" Then d_t("上期数") = "'" & replace(CStr(data_value), "'", "")
                                    If data_type <> "文字" And CStr(Find_Target_Pro(target_code, "是否不转换单位")) = "" Then
                                         d_t("上期数") = ycV(data_value) * unitC
                                    End If
                                    Set d_data(org_name & target_code) = d_t
                                Else
                                    Set d_t = CreateObject("scripting.dictionary")
                                    d_t("数据属性") = data_type
                                    d_t("上期数") = data_value
                                    If data_type = "文字" Then d_t("上期数") = "'" & replace(CStr(data_value), "'", "")
                                    If data_type <> "文字" And CStr(Find_Target_Pro(target_code, "是否不转换单位")) = "" Then
                                         d_t("上期数") = ycV(data_value) * unitC
                                    End If
                                    d_t("当期数") = ""
                                    d_t("机构名称") = org_name
                                    d_t("指标编码") = target_code
                                    If CStr(d_t("机构类别")) = "" Then
                                        d_t("机构类别") = Find_Org_Pro(org_code, "机构类别")
                                    End If
                                    d_t("指标名称") = CStr(arr(i, 2))
                                    Set d_data(org_name & target_code) = d_t
                                End If
                                Set d_t = Nothing
                                
                            End If
                        Next
                    Else
                        d_not_orgs(org_name) = 1
                        msg_not_orgs = msg_not_orgs & org_name & ";"
                    End If
                        
                    Set ws_src = Nothing
                    wb_src.Close False
                    Set wb_src = Nothing
                    Erase arr
                End If
            Next
        End If
        If y Mod 5 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理上期数第：" & CStr(y + 1) & "个/" & CStr(UBound(fns) + 1) & "个")
    Next
    
    If d_data.count = 0 Then
        InitSomething_Data_gddj = "无相关数据."
        Exit Function
    End If
    
    Call gloCft.ShowStatusMessage_Up("数据读取完毕，转换为数组。")
    If IsOnlyOutputHzF = False Or IsHzFTol = False Then
        ReDim arrCur(1 To 1 + d_data.count, 1 To 20) As Variant
    Else
        ReDim arrCur(1 To 1 + hz_count, 1 To 20) As Variant
    End If
    
    arrCur(1, 1) = "地区": arrCur(1, 2) = "数据属性": arrCur(1, 3) = "机构名称": arrCur(1, 4) = "指标编码": arrCur(1, 5) = "指标名称"
    arrCur(1, 6) = "是否与大集中核对": arrCur(1, 7) = "当期数": arrCur(1, 8) = "上期数": arrCur(1, 9) = "变动绝对值": arrCur(1, 10) = "环比变动"
    arrCur(1, 11) = "变动原因/变动说明": arrCur(1, 12) = "是否说明": arrCur(1, 13) = "计算过程": arrCur(1, 14) = "备注": arrCur(1, 15) = "机构类别"
    arrCur(1, 16) = "承接行": arrCur(1, 17) = "社会信用代码": arrCur(1, 18) = "数据日期": arrCur(1, 19) = "币种": arrCur(1, 20) = "频度"
    
    Set dDataCol = CreateObject("scripting.dictionary")
    For i = 1 To UBound(arrCur, 2)
        dDataCol(CStr(arrCur(1, i))) = i
    Next
    
    
    Dim v As Variant, k As Variant, k0 As Variant
    
    If IsOnlyOutputHzF = False Or IsHzFTol = False Then
        i = 2 '
        For Each v In d_data.items
            For Each k In v.keys
                arrCur(i, dDataCol(k)) = v(k)  ' 填充对应指标的值
            Next
            i = i + 1
            
            If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行数组转换，执行到第" & CStr(i) & "行。")
        Next
    Else
        i = 2 '
        j = 2
        For Each k0 In d_data.keys
            If InStr(k0, STR_HZ_EZ) > 0 Then
                Set v = d_data(k0)
                For Each k In v.keys
                    arrCur(j, dDataCol(k)) = v(k)  ' 填充对应指标的值
                    
                Next
                j = j + 1
            End If
            i = i + 1
            
            If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行数组转换，执行到第" & CStr(i) & "行。")
        Next
    End If
    
    Set d_data = Nothing
    
    Call InitSomething_Data_Col
    
    '构造数据字典
    Set dCurDataRef = CreateObject("scripting.dictionary")  '当期表  字典 不用复杂校验的话，是不用的
    Set dPreDataRef = CreateObject("scripting.dictionary")  '上期表  字典
    For i = 2 To UBound(arrCur)
        '补充机构类别 频度
        arrCur(i, col_数据日期) = CurDate
        arrCur(i, col_频度) = "季"
        arrCur(i, col_币种列) = "人民币"
        arrCur(i, col_是否重点指标) = Find_Target_Pro(CStr(arrCur(i, col_指标编码)), "是否与大集中核对")

        arrCur(i, col_地区) = Find_Org_Pro(right(CStr(arrCur(i, col_机构名称)), 18), "归属行")
        arrCur(i, Int(ycV(dDataCol("承接行")))) = Find_Org_Pro(right(CStr(arrCur(i, col_机构名称)), 18), "承接行")
        
        If arrCur(i, col_数据属性) <> "文字" Then
            If ycV(arrCur(i, col_当期值)) <> 0 And ycV(arrCur(i, 上期值列)) = 0 Then
                arrCur(i, 备注列) = "本期有，上期无"
            ElseIf ycV(arrCur(i, col_当期值)) = 0 And ycV(arrCur(i, 上期值列)) <> 0 Then
                arrCur(i, 备注列) = "本期无，上期有"
            ElseIf CStr(arrCur(i, col_当期值)) = CStr(arrCur(i, 上期值列)) Then
                
            Else
                arrCur(i, 增减额列) = ycV(arrCur(i, col_当期值)) - ycV(arrCur(i, 上期值列))
                arrCur(i, 环比列) = ycV(arrCur(i, 增减额列)) / ycV(arrCur(i, 上期值列)) * 100
                arrCur(i, 备注列) = AddbfMark(ycV(arrCur(i, 环比列)))
            End If
        Else
            If CStr(arrCur(i, col_当期值)) <> "" And CStr(arrCur(i, 上期值列)) = "" Then
                arrCur(i, 备注列) = "本期有，上期无"
            ElseIf CStr(arrCur(i, col_当期值)) = "" And CStr(arrCur(i, 上期值列)) <> "" Then
                arrCur(i, 备注列) = "本期无，上期有"
            ElseIf replace(CStr(arrCur(i, col_当期值)), "'", "") = replace(CStr(arrCur(i, 上期值列)), "'", "") Then
                
            Else
                arrCur(i, 增减额列) = "文字变动"
            End If
        End If
        If arrCur(i, col_机构类别) = "" Then
            arrCur(i, col_机构类别) = Find_Org_Pro(CStr(arrCur(i, col_机构名称)), "机构类别")
        End If
        
        If CStr(arrCur(i, col_当期值)) <> "" Then
            dCurDataRef(CStr(arrCur(i, col_机构名称)) & CStr(arrCur(i, col_指标编码))) = ycV(arrCur(i, col_当期值))
            If arrCur(i, col_数据属性) = "文字" Then dCurDataRef(CStr(arrCur(i, col_机构名称)) & CStr(arrCur(i, col_指标编码))) = CStr(arrCur(i, col_当期值))
        End If
        
        If CStr(arrCur(i, 上期值列)) <> "" Then
            dPreDataRef(CStr(arrCur(i, col_机构名称)) & CStr(arrCur(i, col_指标编码))) = ycV(arrCur(i, 上期值列))
            If arrCur(i, col_数据属性) = "文字" Then dPreDataRef(CStr(arrCur(i, col_机构名称)) & CStr(arrCur(i, col_指标编码))) = CStr(arrCur(i, 上期值列))
        End If
        
        If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在处理跨期计算，第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
    Next
    InitSomething_Data_gddj = ""
Exit Function

errHandel:
    InitSomething_Data_gddj = "初始化原始数据有误:" & err.Description & ";"
End Function



Sub test()
    Dim ws As Worksheet, s As String
    Set ws = ThisWorkbook.Worksheets("执行面板")
    s = ws.Range("A8").value
    Debug.Print Len(s)
    Debug.Print CStr(right(s, 18))
    Debug.Print Len("潮州市湘桥区晟达小额贷款股份有限公司")
End Sub

'
'   初始化特殊指标表字典
'
Private Function InitSpecialRef() As String
    Dim ws As Worksheet, i%, j%, k%, arrSpcRef() As Variant, arrT() As String, code As String, str As String
    
    Dim wsFre As String
    
    Set dSpcRef = CreateObject("scripting.dictionary")
    
    On Error GoTo errhandle
    '指标代码，指标名，备注，人民币伐值，美元合计伐值，数据属性，说明，表单
    Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer, c5 As Integer, c6 As Integer, c7 As Integer, c8 As Integer, c9 As Integer
    c1 = 1: c2 = 3: c3 = 4: c4 = 5: c5 = 6: c6 = 7: c7 = 8: c8 = 12: c9 = 13
    

    '若选了，则
    Set ws = ThisWorkbook.Worksheets("特殊指标-自定义")
    
    If InStr(ws.name, "特殊指标") > 0 Then
        'Call gloEt.ShowData(ws)
        arrSpcRef = ws.Range("A1").CurrentRegion.value   'ws.UsedRange.Value
        For i = 2 To UBound(arrSpcRef) Step 1
            If CStr(arrSpcRef(i, 10)) = "" Or VarType(arrSpcRef(i, 10)) = vbEmpty Then
                arrT = Split(arrSpcRef(i, c1), "/")
                For j = 0 To UBound(arrT) Step 1
                    code = replace(arrT(j), "'", "")
                    str = arrSpcRef(i, c2) & "," & _
                        CStr(val(arrSpcRef(i, c3))) & "," & _
                        CStr(val(arrSpcRef(i, c4))) & "," & _
                        CStr(arrSpcRef(i, c5)) & "," & _
                        replace(CStr(arrSpcRef(i, c6)), "'", "") & "," & _
                        CStr(arrSpcRef(i, c7)) & "," & CStr(arrSpcRef(i, c8)) & "," & CStr(arrSpcRef(i, c9))
                        
                        
                    '若出现重复，说明还有另一种校验,按 key(指标代码) = val(备注,人民币伐值,美元伐值,数据属性,说明,表单排列)
                    If dSpcRef.exists(code) Then
                        dSpcRef(code) = CStr(dSpcRef(code)) & "," & str
                    Else
                        dSpcRef(code) = str
                    End If
                    
                    '如果是这个，arrSpcRef(i, c6) 和 arrT(j) 位置换一下，再记录一次
                    If CStr(arrSpcRef(i, c2)) = "指标需要对应存在。" Then
                        code = replace(CStr(arrSpcRef(i, c6)), "'", "")
                        str = arrSpcRef(i, c2) & "," & _
                            CStr(val(arrSpcRef(i, c3))) & "," & _
                            CStr(val(arrSpcRef(i, c4))) & "," & _
                            CStr(arrSpcRef(i, c5)) & "," & _
                            replace(CStr(arrT(j)), "'", "") & "," & _
                            CStr(arrSpcRef(i, c7)) & "," & CStr(arrSpcRef(i, c8)) & "," & CStr(arrSpcRef(i, c9))
                            
                        '若出现重复，说明还有另一种校验,按 key(指标代码) = val(备注,人民币伐值,美元伐值,数据属性,说明,表单排列)
                        If dSpcRef.exists(code) Then
                            dSpcRef(code) = CStr(dSpcRef(code)) & "," & str
                        Else
                            dSpcRef(code) = str
                        End If
                    End If
                Next j
            Else
                '禁用指标
            End If  '对于禁用的跳过
        Next i
    End If
    
    InitSpecialRef = ""
Exit Function

errhandle:
    Set dSpcRef = Nothing
    Set dCurDataRef = Nothing
    Set dPreDataRef = Nothing
    Erase arrAlert
    InitSpecialRef = "特殊指标表初始化失败," & err.Description & ";"
    err.Clear
End Function

'
'   初始化指标参照和机构参照
'
Private Sub Init_Data_def()
    Set dtarget_def = Nothing '指标参照
    Set dorg_def = Nothing '机构参照
    Set dorg_name_def = Nothing

    Set dtarget_def = CreateObject("scripting.dictionary") '指标参照
    Set dorg_def = CreateObject("scripting.dictionary") '机构参照
    Set dorg_name_def = CreateObject("scripting.dictionary")
    
    Dim arr As Variant, i As Long, j As Integer, d_temp As Object
    
    '指标参照
    arr = gloEt.GetUsedRange(ThisWorkbook.Worksheets("指标参照")).value
    For i = 2 To UBound(arr)
        '不需要过滤，因为前期已经清除了空值
        If Int(ycV(arr(i, 10))) = 0 Then
            Set d_temp = CreateObject("scripting.dictionary")
            For j = 1 To UBound(arr, 2)
                d_temp(CStr(arr(1, j))) = CStr(arr(i, j))
            Next
            Set dtarget_def(CStr(arr(i, 1))) = d_temp
        End If
    Next
    
    '机构参照
    arr = gloEt.GetUsedRange(ThisWorkbook.Worksheets("机构参照")).value
    For i = 2 To UBound(arr)
        Set d_temp = CreateObject("scripting.dictionary")
        For j = 1 To UBound(arr, 2)
            d_temp(CStr(arr(1, j))) = CStr(arr(i, j))
        Next
        Set dorg_def(CStr(arr(i, 5))) = d_temp
    Next
    
    '机构参照-按名称
    arr = gloEt.GetUsedRange(ThisWorkbook.Worksheets("机构参照")).value
    For i = 2 To UBound(arr)
        Set d_temp = CreateObject("scripting.dictionary")
        For j = 1 To UBound(arr, 2)
            d_temp(CStr(arr(1, j))) = CStr(arr(i, j))
        Next
        Set dorg_name_def(CStr(arr(i, 1))) = d_temp
    Next
    
    Erase arr
    
    
End Sub


'


'
'   添加是否说明标志
'   strSt - 指标
'   index - 数组对应的行
'   arrCur - 为所有数据
'
Private Function DoCompare_AddMyExplain(strSt As String, index As Long) As String
    Dim str As String, strRel As String, arrSt() As SpecialTaget, st As SpecialTaget, arr As Variant, u As Double, i As Integer
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, IsRuan As Boolean
    Dim dqStep As Integer, bz As String
    dqStep = 8  '现在分8列,增加了幅度
    
    curD = ycV(arrCur(index, col_当期值))   '当期值
    preD = ycV(arrCur(index, 上期值列))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, col_币种列))
    dType = CStr(arrCur(index, col_数据属性)) '数据属性
    bz = CStr(arrCur(index, 备注列))  '备注
    
    arr = Split(dSpcRef(strSt), ",")
    ReDim arrSt(Int(UBound(arr) / dqStep))
    
    For i = 0 To Int(UBound(arr) / dqStep) Step 1
        arrSt(i).text = arr(dqStep * i + 0)
        arrSt(i).rmbF = Abs(CDbl(ycV(arr(dqStep * i + 1))))
        arrSt(i).usdF = Abs(CDbl(ycV(arr(dqStep * i + 2))))
        arrSt(i).dataType = CStr(arr(dqStep * i + 3))
        arrSt(i).explain = CStr(arr(dqStep * i + 4))
        arrSt(i).TableName = CStr(arr(dqStep * i + 5))
        arrSt(i).detailed = CStr(arr(dqStep * i + 6))
        arrSt(i).bf = Abs(CDbl(ycV(arr(dqStep * i + 7))))
        
        If arrSt(i).detailed <> "" Then arrSt(i).detailed = "-" & arrSt(i).detailed
    Next i
    
    
    u = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    
    If ycV(Find_Target_Pro(strSt, "是否不转换单位")) = 1 Then u = 1
    
    strRel = ""
    IsRuan = True
    For i = 0 To UBound(arrSt) Step 1
        st = arrSt(i)
        '判断设置的指标是否是全数据属性,或者跟传入的属性一致
        If st.dataType = "" Or st.dataType = "0" Or InStr(st.dataType, dType) > 0 Then
            str = ""
            If st.text = "指标不应有数，需说明。" Then
                'If curD <> 0 Then
                '改成不为空，是防止有些有数据，但是导出值为0
                If CStr(arrCur(index, col_当期值)) <> "" Then
                    str = "指标有数需说明。" & st.detailed
                ElseIf preD <> 0 Then
                    'str = st.text & "(上期有数)"
                End If
            ElseIf st.text = "指标有变动需说明。" Then
                str = 指标有变动需说明(st, index, u)
            ElseIf st.text = "指标不应为负数。" Then
                If CStr(arrCur(index, col_当期值)) <> "" Then
                    If ycV(arrCur(index, col_当期值)) < 0 Then
                        str = "指标不应为负数。" & st.detailed
                    End If
                End If
            ElseIf st.text = "当年累计指标比上期不应减少。" Then
                '首先判断是否为负数
                If curD < 0 Then
                    str = "发生额指标不应为负数，需说明。" & st.detailed
                End If
                '判断时间，首先本期日期是大于上期的，下来判断年份是否一致就好
                If Year(CurDate) = Year(PreDate) Then
                    '同年份的话，则必须满足
                    If valH < 0 Then
                        If CStr(st.rmbF) <> "" Then
                            If Abs(valH) > ycV(st.rmbF) / u Then
                                str = str & " " & "当年累计指标比上期不应减少,阈值超过" & Format(ycV(st.rmbF) / u, "0.000") & CStr(DesUnit) & "，请核实。" & st.detailed
                            End If
                        Else
                            str = str & " " & "当年累计指标比上期不应减少，需说明。" & st.detailed
                        End If
                    End If
                Else
                    '若不同年份,若月份>=上期，如2018.8,2017.8，则小于0异常
                    If Month(CurDate) >= Month(PreDate) Then
                        If valH < 0 Then
                            If CStr(st.rmbF) <> "" Then
                                If Abs(valH) > ycV(st.rmbF) * u Then
                                    str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据比" & CStr(PreDate) & "的小,阈值超过" & Format(ycV(st.rmbF) / u, "0.000") & CStr(DesUnit) & "，请核实。" & st.detailed
                                End If
                            Else
                                str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据比" & CStr(PreDate) & "的小，请核实。" & st.detailed
                            End If
                            
                        End If
                    Else    '若月份<上期，如2018.1,2017.12，则>=0异常
                        If valH >= 0 Then
                            str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据 >= " & CStr(PreDate) & "的，请核实。" & st.detailed
                        End If
                    End If
                End If
            ElseIf st.text = "历史累计指标比上期不应减少。" Then
                If curD < 0 Then
                    str = "发生额指标不应为负数，需说明。" & st.detailed
                End If
                If valH < 0 Then
                    str = str & " " & "历史累计指标比上期不应减少，需说明。" & st.detailed
                End If
            ElseIf st.text = "指标变幅异常需说明。" Then
                If dPro = "美元合计" Then
                    If Abs(valBf) >= Abs(st.usdF) Then
                        str = "指标变幅超过" & Format(Abs(st.usdF), "0.00%") & "%，需说明。" & st.detailed
                    End If
                Else
                    If Abs(valBf) >= Abs(st.rmbF) Then
                        str = "指标变幅超过" & Format(Abs(st.rmbF), "0.00%") & "%，需说明。" & st.detailed
                    End If
                End If
            '-------------------------------------   开始每个校验一个过程      ----------------------
            ElseIf st.text = "指标比上期增加需说明。" Then
                str = 指标比上期增加需说明(st, index, u)
                
            ElseIf st.text = "指标比上期减少需说明。" Then
                str = 指标比上期减少需说明(st, index, u)
                
            ElseIf st.text = "指标相除值核查。" Then
                str = 指标相除值核查(st, index, u, strSt)
            
            ElseIf st.text = "指标相除值太小核查。" Then
                str = 指标相除值太小核查(st, index, u, strSt)
            
            ElseIf st.text = "指标需要对应存在。" Then
                str = 指标需要对应存在(st, index, u, strSt)
            
            ElseIf st.text = "指标相减应大于核查。" Then
                str = 指标相减应大于核查(st, index, u, strSt)
                
            ElseIf st.text = "指标应相等核查。" Then
                str = 指标应相等核查(st, index, u, strSt)
            
            ElseIf st.text = "指标应为整数。" Then
                If CStr(Abs(curD)) <> CStr(CLng(Abs(curD))) Then
                    str = "指标应该为整数，出现小数，需说明。"
                End If
            ElseIf st.text = "指标应能被某数整除。" Then
                str = 校验指标能被某数整除(st, index)
                
            ElseIf st.text = "指标本期余额应小于（上期余额+本期发生额）。" Then
                str = 校验本期上期余额及本期发生额(st, index)
                     
            ElseIf st.text = "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。" Then
                str = 校验余额及当年累计发生额(st, index)
                
            ElseIf st.text = "指标新增需说明。" Then
                If curD > 0 And preD = 0 Then
                    If Abs(valH) * u >= Abs(st.rmbF) Then
                        str = "指标新增超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                    End If
                End If
                
            ElseIf st.text = "指标结清需说明。" Then
                str = 指标结清需说明(st, index, u)

            Else
                str = "其他需要说明指标"
            End If
            
            If str <> "" And InStr(str, "软性") <= 0 Then IsRuan = False
            
            If str <> "" Then
                If strRel = "" Then
                    strRel = strRel & str
                Else
                    strRel = strRel & "    ||    " & str
                End If
            End If
            'strRel = strRel & str
        End If
    Next i
    
    If strRel <> "" And IsRuan = True Then
        wsCur.Cells(index, 是否说明列).Interior.colorindex = Complex_Soft_Color_Index
    End If
    DoCompare_AddMyExplain = strRel
End Function

Function 指标结清需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, bz As String
    
    curD = ycV(arrCur(index, col_当期值))   '当期值
    preD = ycV(arrCur(index, 上期值列))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, col_币种列))
    dType = CStr(arrCur(index, col_数据属性)) '数据属性
    bz = CStr(arrCur(index, 备注列))  '备注
    
    str = ""
    
    If curD = 0 And preD > 0 Then
        If Abs(valH) * u >= Abs(st.rmbF) Then
            str = "指标结清超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
        End If
    End If
    
    指标结清需说明 = str
End Function

Function 指标比上期减少需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, bz As String
    
    curD = ycV(arrCur(index, col_当期值))   '当期值
    preD = ycV(arrCur(index, 上期值列))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, col_币种列))  '币种
    dType = CStr(arrCur(index, col_数据属性)) '数据属性
    bz = CStr(arrCur(index, 备注列))  '备注
    
    str = ""
    '仅绝对值
    If valH * u < 0 Then
        If dPro = "美元合计" Then
            If st.usdF = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标比上期减少超过，需说明。" & st.detailed
                End If
            Else
                If Abs(valH) * u >= Abs(st.usdF) Then
                    str = "指标比上期减少超过" & CStr(Abs(st.usdF)) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
                End If
            End If
        Else
            ' 变动需说明
            If st.rmbF <> 0 And st.bf = 0 Then
                If Abs(valH) * u >= Abs(st.rmbF) Then
                    str = "指标比上期减少超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                End If
            ' 相当于执行变幅异常
            ElseIf st.rmbF = 0 And st.bf <> 0 Then
                If bz = "本期有，上期无" Then
                    str = "指标比上期减少超过" & Format(Abs(st.bf), "0.00%") & "%，需说明(指标新增)。" & st.detailed
                Else
                    If Abs(valBf) >= Abs(st.bf) Then
                        str = "指标比上期减少超过" & Format(Abs(st.bf), "0.00%") & "%，需说明。" & st.detailed
                    End If
                End If
            ' 有变动就需要说明
            ElseIf st.rmbF = 0 And st.bf = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标比上期减少超过，需说明。" & st.detailed
                End If
            ' 都有的情况
            ElseIf st.rmbF <> 0 And st.bf <> 0 Then
                '回到单纯变动
                If bz = "本期有，上期无" Then
                    If Abs(valH) * u >= Abs(st.rmbF) Then
                        str = "指标比上期减少超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                    End If
                Else
                    If Abs(valH) * u >= Abs(st.rmbF) And Abs(valBf) >= Abs(st.bf) Then
                        str = "指标比上期减少超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & ",且变幅超过" & Format(Abs(st.bf), "0.00%") & "%，需说明。" & st.detailed
                    End If
                End If
            Else
            
            End If
        End If
    End If
    
    指标比上期减少需说明 = str
End Function

'
'   指标比上期增加需说明。
'
Function 指标比上期增加需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, bz As String
    
    curD = ycV(arrCur(index, col_当期值))   '当期值
    preD = ycV(arrCur(index, 上期值列))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, col_币种列))  '币种
    dType = CStr(arrCur(index, col_数据属性)) '数据属性
    bz = CStr(arrCur(index, 备注列))  '备注
    
    str = ""
    '仅绝对值
    If valH * u > 0 Then
        If dPro = "美元合计" Then
            If st.usdF = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标比上期增加超过，需说明。" & st.detailed
                End If
            Else
                If Abs(valH) * u >= Abs(st.usdF) Then
                    str = "指标比上期增加超过" & CStr(Abs(st.usdF)) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
                End If
            End If
        Else
            ' 变动需说明
            If st.rmbF <> 0 And st.bf = 0 Then
                If Abs(valH) * u >= Abs(st.rmbF) Then
                    str = "指标比上期增加超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                End If
            ' 相当于执行变幅异常
            ElseIf st.rmbF = 0 And st.bf <> 0 Then
                If bz = "本期有，上期无" Then
                    str = "指标比上期增加超过" & Format(Abs(st.bf), "0.00%") & "%，需说明(指标新增)。" & st.detailed
                Else
                    If Abs(valBf) >= Abs(st.bf) Then
                        str = "指标比上期增加超过" & Format(Abs(st.bf), "0.00%") & "%，需说明。" & st.detailed
                    End If
                End If
            ' 有变动就需要说明
            ElseIf st.rmbF = 0 And st.bf = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标比上期增加超过，需说明。" & st.detailed
                End If
            ' 都有的情况
            ElseIf st.rmbF <> 0 And st.bf <> 0 Then
                '回到单纯变动
                If bz = "本期有，上期无" Then
                    If Abs(valH) * u >= Abs(st.rmbF) Then
                        str = "指标比上期增加超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                    End If
                Else
                    If Abs(valH) * u >= Abs(st.rmbF) And Abs(valBf) >= Abs(st.bf) Then
                        str = "指标比上期增加超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & ",且变幅超过" & Format(Abs(st.bf), "0.00%") & "%，需说明。" & st.detailed
                    End If
                End If
            Else
            
            End If
        End If
    End If

    指标比上期增加需说明 = str
End Function

Function 指标有变动需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, bz As String
    
    curD = ycV(arrCur(index, col_当期值))   '当期值
    preD = ycV(arrCur(index, 上期值列))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, col_币种列))  '币种
    dType = CStr(arrCur(index, col_数据属性)) '数据属性
    bz = CStr(arrCur(index, 备注列))  '备注
    
    str = ""
    '仅绝对值
    If dType = "文字" Then
        If replace(CStr(arrCur(index, col_当期值)), "'", "") = replace(CStr(arrCur(index, 上期值列)), "'", "") Then
            
        Else
            str = "指标有变动，需说明。" & st.detailed
        End If
    Else
        If dPro = "美元合计" Then
            If st.usdF = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标有变动，需说明。" & st.detailed
                End If
            Else
                If Abs(valH) * u >= Abs(st.usdF) Then
                    str = "指标变动超过" & CStr(Abs(st.usdF)) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
                End If
            End If
        Else
            ' 变动需说明
            If st.rmbF <> 0 And st.bf = 0 Then
                If Abs(valH) * u >= Abs(st.rmbF) Then
                    str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                End If
            ' 相当于执行变幅异常
            ElseIf st.rmbF = 0 And st.bf <> 0 Then
                If bz = "本期有，上期无" Then
                    str = "指标变幅超过" & Format(Abs(st.bf), "0.00") & "%，需说明(指标新增)。" & st.detailed
                Else
                    If Abs(valBf) >= Abs(st.bf) Then
                        str = "指标变幅超过" & Format(Abs(st.bf), "0.00") & "%，需说明。" & st.detailed
                    End If
                End If
            ' 有变动就需要说明
            ElseIf st.rmbF = 0 And st.bf = 0 Then
                If Abs(valH) <> 0 Then
                    str = "指标有变动，需说明。" & st.detailed
                End If
            ' 都有的情况
            ElseIf st.rmbF <> 0 And st.bf <> 0 Then
                '回到单纯变动
                If bz = "本期有，上期无" Then
                    If Abs(valH) * u >= Abs(st.rmbF) Then
                        str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                    End If
                Else
                    If Abs(valH) * u >= Abs(st.rmbF) And Abs(valBf) >= Abs(st.bf) Then
                        str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & ",且变幅超过" & Format(Abs(st.bf), "0.00") & "%，需说明。" & st.detailed
                    End If
                End If
            Else
            
            End If
        End If
    End If
    
    
    指标有变动需说明 = str
End Function


Function 指标应相等核查(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    Dim b As Double
    b = 0
    curD = val(arrCur(index, col_当期值))
    temp = arrCur(index, col_机构名称) & st.explain
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    
    If CStr(st.rmbF) = "" Then
        If Abs(curD - b) > dbl_Epsilon Then
            str = strSt & "不等于" & st.explain & "偏差" & Format(curD - b, "0.0000") & ",请核实。" & st.detailed
        End If
    Else
        If Abs(curD - b) > ycV(st.rmbF) * u Then
            str = strSt & "不等于" & st.explain & "偏差" & Format(curD - b, "0.0000") & ",请核实。" & st.detailed
        End If
    End If
    
    指标应相等核查 = str
End Function

'
'   指标需要对应存在
'
Function 指标需要对应存在(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, dbz As String
    Dim b As Double
    
    str = ""
    b = 0
    curD = val(arrCur(index, col_当期值))
    dbz = CStr(arrCur(index, col_币种列))
    
    If dbz = "美元合计" Then
        If st.usdF <> 0 Then
            dbz = "人民币"
        End If
    End If
    temp = arrCur(index, col_机构名称) & st.explain 'arrCur(index, col_指标编码) &
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    
    '是否需要小于0的呢。
    If curD <> 0 And b = 0 Then
        str = strSt & "有数," & st.explain & "无数,请核实。" & st.detailed
    ElseIf curD = 0 And b <> 0 Then
        str = st.explain & "有数," & strSt & "无数,请核实。" & st.detailed
    End If
    
    指标需要对应存在 = str
End Function


'
'   指标相除值需小于
'
Function 指标相除值太小核查(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    Dim b As Double, u2 As Double, bijiao As Double
    
    str = ""
    b = 0
    dPro = CStr(arrCur(index, col_币种列))  '币种
    curD = val(arrCur(index, col_当期值))
    
    '999 代表那些美元合计/人民币的。
    If st.usdF = 999 Then
        If dPro = "美元合计" Then
            dPro = "人民币"
        End If
    End If
    
    temp = arrCur(index, col_机构名称) & st.explain
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If ycV(Find_Target_Pro(st.explain, "是否不转换单位")) = 1 Then u2 = 1
    
    '是否需要小于0的呢。
    If b = 0 Then
        If curD <> 0 Then
            str = strSt & "有数," & st.explain & "为空，相除值异常,请核实。"
        End If
    Else
        bijiao = Abs(curD * u / (b * u2))
        If dPro = "美元合计" Then
            If bijiao < st.usdF Then
                str = strSt & "/" & st.explain & "=" & CStr(Format(bijiao, "0.0000")) & ",小于" & CStr(st.usdF) & "，需核实。" & st.detailed
            End If
        Else
            If bijiao < st.rmbF Then
                If st.usdF = 999 And CStr(arrCur(index, 11)) = "美元合计" Then
                    str = strSt & "(美元合计)/" & st.explain & "(人民币)=" & CStr(Format(bijiao, "0.0000")) & ",小于" & CStr(st.rmbF) & "，需核实。" & st.detailed
                Else
                    str = strSt & "/" & st.explain & "=" & CStr(Format(bijiao, "0.0000")) & ",小于" & CStr(st.rmbF) & "，需核实。" & st.detailed
                End If
            End If
        End If
    End If
    
    指标相除值太小核查 = str
End Function

'
'   指标相除值核查
'
Function 指标相除值核查(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    Dim b As Double, u2 As Double, bijiao As Double
    
    str = ""
    b = 0
    dPro = CStr(arrCur(index, col_币种列))  '币种
    curD = val(arrCur(index, col_当期值))
    
    '999 代表那些美元合计/人民币的。
    If st.usdF = 999 Then
        If dPro = "美元合计" Then
            dPro = "人民币"
        End If
    End If
    
    temp = arrCur(index, col_机构名称) & st.explain
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If ycV(Find_Target_Pro(st.explain, "是否不转换单位")) = 1 Then u2 = 1
    
    '是否需要小于0的呢。
    If b = 0 Then
        If curD <> 0 Then
            str = strSt & "有数," & st.explain & "为空，相除值异常,请核实。"
        End If
    Else
        bijiao = Abs(curD * u / (b * u2))
        If dPro = "美元合计" Then
            If bijiao > st.usdF Then
                str = strSt & "/" & st.explain & "=" & CStr(Format(bijiao, "0.0000")) & ",超过" & CStr(st.usdF) & "，需说明。" & st.detailed
            End If
        Else
            If bijiao > st.rmbF Then
                If st.usdF = 999 And CStr(arrCur(index, 11)) = "美元合计" Then
                    str = strSt & "（美元合计）/" & st.explain & "（人民币）=" & CStr(Format(bijiao, "0.0000")) & ",超过" & CStr(st.rmbF) & "，需说明。" & st.detailed
                Else
                    str = strSt & "/" & st.explain & "=" & CStr(Format(bijiao, "0.0000")) & ",超过" & CStr(st.rmbF) & "，需说明。" & st.detailed
                End If
            End If
        End If
    End If
    
    指标相除值核查 = str
End Function

'
'   指标相减应大于核查
'
Function 指标相减应大于核查(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    Dim b As Double, u2 As Double, bijiao As Double
    
    str = ""
    b = 0
    dPro = CStr(arrCur(index, col_币种列))  '币种
    curD = val(arrCur(index, col_当期值))
    
    temp = arrCur(index, col_机构名称) & st.explain
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If ycV(Find_Target_Pro(st.explain, "是否不转换单位")) = 1 Then u2 = 1
    
    '是否需要小于0的呢。
    bijiao = (curD * u) - (b * u2)
    If dPro = "美元合计" Then
        If bijiao < Abs(st.usdF) Then
            str = strSt & "-" & st.explain & "小于" & CStr(st.usdF) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
        End If
    Else
        If bijiao < Abs(st.rmbF) Then
            str = strSt & "-" & st.explain & "小于" & CStr(st.rmbF) & 特殊指标数据单位 & "，需说明。" & st.detailed
        End If
    End If
    
    指标相减应大于核查 = str
    
End Function




'
'   指标应能被某数整除。
'
Function 校验指标能被某数整除(st As SpecialTaget, index As Long) As String
    Dim str As String, u As Double, r As Double
    Dim curD As Double
    str = ""
    If ycV(Find_Target_Pro(replace(arrCur(index, col_指标编码), "'", ""), "是否不转换单位")) = 1 Then
        u = 1
    Else
        u = GetUnitVal(DesUnit)
    End If
    
    On Error GoTo errH
    curD = Abs(val(arrCur(index, col_当期值)))   '当期值
    r = (u * curD) / Int(st.rmbF)
    
    If CStr(r) <> CStr(CLng(r)) Then
        str = "指标应能被" & CStr(Format(st.rmbF, "0")) & "整除。" & st.detailed
    End If
    
    校验指标能被某数整除 = str
    
Exit Function

errH:
    str = "除数为0，特殊指标表设置有误。"
    校验指标能被某数整除 = str
End Function

'
'   指标本期余额应小于（上期余额+本期当年累计发生额-上期当年累计发生额）
'
Function 校验余额及当年累计发生额(st As SpecialTaget, index As Long) As String
    
End Function

'
'   获取三个数值的最大值
'
Function GetMaxOfThree(a As Double, b As Double, c As Double) As Double
    Dim max As Double
    max = a
    If max > b Then
    
    Else
    
    End If
End Function


'
'   指标本期余额应小于（上期余额+本期发生额）。
'
Function 校验本期上期余额及本期发生额(st As SpecialTaget, index As Long) As String
    
End Function

Private Function 发生额与余额误差判断(v As Double, max As Double) As String
    Dim str As String, l1 As Double, l2 As Double, l As Double, t As Double
    str = ""
    l = 10000
    
    If Round(v, 2) >= 0 Then
        发生额与余额误差判断 = ""
        Exit Function
    End If
    
    t = Abs(v)
    If Round(t, 2) < 150 Then
        str = "精度误差：上期余额+本期发生额>本期余额，偏差150元内。"
    Else
        If max < l Then
            If Round(t, 2) < max Then
                str = "精度误差：上期余额+本期发生额>本期余额，超过150元，未超过1万元。"
            Else
                str = "上期余额+本期发生额>本期余额，超过该业务量万分之一。"
            End If
        Else
            If Round(t, 2) < l Then
                str = "精度误差：上期余额+本期发生额>本期余额，超过150元，未超过" & CStr(l) & "元。"
            Else
                str = "上期余额+本期发生额>本期余额，超过" & CStr(l) & "元。"
            End If
        End If
    End If
    发生额与余额误差判断 = str
End Function


'给新的工作簿增加宏代码
Private Sub AddMacroCode()
    Dim MyCodes As String, fullName As String
    
    On Error GoTo errhandle
    
    '插入启动菜单
    MyCodes = gloCft.GetConfig("说明文件菜单模块宏代码")
    With wbDes.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines 2, MyCodes
    End With
    '插入执行命令
    
    MyCodes = gloCft.GetConfig("输出说明通用宏")
    With wbDes.VBProject.VBComponents.add(1)
        .name = "common"
        .CodeModule.InsertLines 2, MyCodes
    End With
    
'    MyCodes = gloCft.GetConfig("输出说明单元格处理宏")
'    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
'        .name = "DealCell"
'        .CodeModule.InsertLines 2, MyCodes
'    End With
    
    MyCodes = "public Const unit as Double = " + CStr(GetUnitVal(DesUnit) / GetUnitVal(分析文件数据单位)) + Chr(10) + Chr(13) _
            + "public Const 分析文件数据单位 as String = " + Chr(34) + CStr(分析文件数据单位) + Chr(34) + Chr(10) + Chr(13) _
            + "public Const 说明文件名 as String = " + Chr(34) + CStr(Year(CurDate)) + "." + CStr(Month(CurDate)) + " 说明文件.xls" + Chr(34) + Chr(10) + Chr(13) _
            + "public const 当期数据日期 as Date = #" & CStr(CurDate) & "#" + Chr(10) + Chr(13) _
            + "public const 上期数据日期 as Date = #" & CStr(PreDate) & "#" + Chr(10) + Chr(13)
    MyCodes = MyCodes + gloCft.GetConfig("说明文件工作模块宏代码")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work"
        .CodeModule.InsertLines 2, MyCodes
    End With
Exit Sub

errhandle:
    err.Clear
    MsgBox "说明文件功能受限。需要在EXCEL菜单 工具-宏－安全性－可靠发行商　中勾选【信任对于VB项目的访问】"
End Sub


'
'   初始化复杂指标表
'
Private Function InitComplexRef() As String
    Dim ws As Worksheet, i%, j%, k%, arr As Variant, arrT As Variant, code As String, str As String
    Dim reg As Object, mc As Object, m As Object, rule As String, rn As Range, errRuleNum As Long, rightRule As Boolean, ruleNum As Long, jyNum As Long
    Dim dtemp As Object, thdvalue As Double, targetRuleNum As Long, notfreNum As Long, rn2 As Range
    Set dComplexRef = Nothing
    Set dComplexRef = CreateObject("scripting.dictionary")
    Set dComplexChecked = Nothing
    Set dComplexChecked = CreateObject("scripting.dictionary")
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    dComplexRef.RemoveAll
    ComplexThd = 0  '误差项
    '单位转换
    ComplexThd = ycV(ThisWorkbook.Worksheets("更新日志").Range("C1").value)
    'ComplexThd = ComplexThd * (GetUnitVal("元") / GetUnitVal(DesUnit))
    
    On Error GoTo errhandle
    '校验表单，校验描述，校验规则，取反标识，禁用标识
    
    CforbidCol = 5

    Set ws = ThisWorkbook.Worksheets("复杂校验-自定义")
    'Call gloEt.ShowData(ws)
    errRuleNum = 0
    ruleNum = 0
    targetRuleNum = 0
    jyNum = 0
    notfreNum = 0
    Set rn = Nothing
    Set rn2 = Nothing '用来标记不读取的多频
    ws.Cells.Interior.colorindex = 0
    arr = ws.Range("A1").CurrentRegion.value   'ws.UsedRange.Value
    

    For i = 2 To UBound(arr) Step 1
        '判断是否禁用 - 空则标识启用，非空则表示禁用
        rightRule = False
        If (CStr(arr(i, CforbidCol)) = "" Or VarType(arr(i, CforbidCol)) = vbEmpty) And CStr(arr(i, 3)) <> "" Then
            rightRule = True
            '清理校验规则
            rule = InitComplexRef_ClearRule(CStr(arr(i, 3)))
            If rule = "" Then
                rightRule = False
                errRuleNum = errRuleNum + 1
                If rn Is Nothing Then
                    Set rn = ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2)))
                Else
                    Set rn = Application.Union(rn, ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))))
                End If
                GoTo nextRule
            Else
                If InitComplexRef_CheckTaget(rule) = False Then
                    rightRule = False
                    errRuleNum = errRuleNum + 1
                    If rn Is Nothing Then
                        Set rn = ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2)))
                    Else
                        Set rn = Application.Union(rn, ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))))
                    End If
                    GoTo nextRule
                Else
                    '可以开始读取当期指标了
                    With reg
                        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
                    End With
                    Set mc = reg.Execute(rule)
                    For Each m In mc
                        '0业务类，1机构类代码，2地区代码，3指标代码，4数据属性，5币种，6频度，7批次
                        arrT = Split(CStr(m.value), ",")
                        code = replace(CStr(arrT(COMPLEX_TARGET_INDEX)), "]", "")  '此处读入指标代码
                        If dComplexRef.exists(code) Then
                            If VarType(dComplexRef(code)) = 0 Then
                                Set dtemp = CreateObject("scripting.dictionary")
                            Else
                                Set dtemp = dComplexRef(code)
                            End If
                        Else
                            Set dtemp = CreateObject("scripting.dictionary")
                        End If
                        ' 3 是校验规则，就不重复读取了吧
                        '判断是h是否有阈值
                        If Not dtemp.exists(rule) Then
                            If CStr(arr(i, 6)) <> "" Then
                                'thdvalue = ycV(CStr(arr(i, 6))) * (GetUnitVal("万元") / GetUnitVal(DesUnit))
                                thdvalue = ycV(CStr(arr(i, 6))) * GetUnitVal("万元")    '转化为元
                            Else
                                thdvalue = ComplexThd
                            End If
                            dtemp(rule) = CStr(arr(i, 1)) & "|" & CStr(arr(i, 2)) & "|" & CStr(m.value) & "|" & CStr(arr(i, 4)) & "|" & CStr(thdvalue)
                            Set dComplexRef(code) = dtemp
                            targetRuleNum = targetRuleNum + 1
                        Else
                            '如果有重复，可能是不同币种、属性、频度，追加上去。
                            If InStr(CStr(dtemp(rule)), CStr(m.value)) > 0 Then
                            
                            Else
                                '也就是，如果超过了0，1，2，3，4。5个数组，说明还有其他属性的该指标校验
                                dtemp(rule) = CStr(dtemp(rule)) & "|" & CStr(m.value)
                            End If
                        End If
                    Next
                End If
            End If
        Else
            rightRule = False
            jyNum = jyNum + 1
        End If
nextRule:
        If rightRule = False Then

        Else
            ruleNum = ruleNum + 1
        End If
    Next i
    
    If rn Is Nothing Then
        Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "读取复杂校验规则" & CStr(UBound(arr) - 1) _
            & "条，实际有效载入" & CStr(ruleNum) & "条，形成指标校验" & CStr(targetRuleNum) & "条，其中，未读取禁用校验" & CStr(jyNum) & "条；未满足频度" & CStr(notfreNum) & "条。")
    Else
        rn.Interior.colorindex = 15
        Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "读取复杂校验规则" & CStr(UBound(arr) - 1) & _
            "条，实际有效载入" & CStr(ruleNum) & "条，形成指标校验" & CStr(targetRuleNum) & "条，其中，未读取禁用校验" & CStr(jyNum) & "条；未满足频度" & CStr(notfreNum) & "条。,未读取的错误校验规则" & CStr(errRuleNum) & "条。")
    End If
    
    If Not rn2 Is Nothing Then rn2.Interior.colorindex = 37
    InitComplexRef = ""
    
Exit Function

errhandle:
    Call EraseSomthing
    InitComplexRef = "复杂指标表初始化失败," & err.Description & ";"
    err.Clear
End Function

Private Function InitComplexRef_CheckTaget(rule As String) As Boolean
    Dim reg As Object, mc As Object, splitCount As Integer, m As Object, arr As Variant
    On Error Resume Next
    splitCount = COMPLEXT_SPLIT_COUNT  '这里修改指标分隔符
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        If Len(m.value) - Len(replace(m.value, ",", "")) <> splitCount Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
        arr = Split(CStr(m.value), ",")
        If CStr(arr(COMPLEX_TARGET_INDEX)) = "" Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
    Next
    
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        If Len(m.value) - Len(replace(m.value, ",", "")) <> splitCount Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
        arr = Split(CStr(m.value), ",")
        '指标代码一定得有
        If CStr(arr(COMPLEX_TARGET_INDEX)) = "" Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
    Next
    InitComplexRef_CheckTaget = True
End Function

'
'   初始化校验规则,检查括号是否匹配
'
Private Function InitComplexRef_ClearRule(rule As String) As String
    Dim clearRule As String
    
    clearRule = replace(rule, "【", "[")
    clearRule = replace(clearRule, "】", "]")
    clearRule = replace(clearRule, "（", "(")
    clearRule = replace(clearRule, "）", ")")
    clearRule = replace(clearRule, "｛", "{")
    clearRule = replace(clearRule, "｝", "}")
    clearRule = replace(clearRule, "，", ",")
    clearRule = replace(clearRule, " ", " ")
    clearRule = replace(clearRule, "　", " ")
    
    
    If Len(replace(clearRule, "[", "")) <> Len(replace(clearRule, "]", "")) Or _
        Len(replace(clearRule, "(", "")) <> Len(replace(clearRule, ")", "")) Or _
        Len(replace(clearRule, "{", "")) <> Len(replace(clearRule, "}", "")) Then
        InitComplexRef_ClearRule = ""
    Else
        If Len(replace(clearRule, "[", "")) = Len(clearRule) Then
            InitComplexRef_ClearRule = ""
        Else
            InitComplexRef_ClearRule = clearRule
        End If
    End If
End Function



'
'   添加复杂校验描述
'
'
Private Function DoCompare_AddMyComplexExplain(strSt As String, index As Long) As String
    Dim str As String, drules As Object, i As Long, arrRules As Variant, arrT As Variant, tRule As String, oldRule As String, roundRule As String
    Dim ct As ComplexTaget, arrCt As Variant, reg As Object, mc As Object, m As Object, tstr As String, jsStr As String, g As Long
    Dim result As Variant, relStr, isText As Boolean
    str = ""
    relStr = ""
    On Error Resume Next
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    Set drules = dComplexRef(strSt)
    arrRules = drules.keys
    
    For i = 0 To UBound(arrRules) Step 1
        oldRule = CStr(arrRules(i))
        tRule = CStr(arrRules(i))
        arrCt = Split(drules(tRule), "|")
        ct.table = CStr(arrCt(0))
        ct.ruleEx = CStr(arrCt(1))
        ct.rule = CStr(arrCt(2))    '用这个判断是否是此指标
        ct.qfbz = CBool(ycV(arrCt(3)))  '空和字符串 代表否
        ct.thdV = ycV(arrCt(4))
        If UBound(arrCt) >= 5 Then
            '说明还有别的，先继续混合，然后输出真正属于的指标
            For g = 5 To UBound(arrCt) Step 1
                ct.rule = CStr(ct.rule) & "|" & CStr(arrCt(g))
            Next g
        End If
        
        'Debug.Assert (ct.ruleEx <> "漏填18.10单户授信小于100万（含）的境内小微型企业贷款户数相应的贷款或贴现")
        'Debug.Assert (ct.ruleEx <> "漏填3370D 18.10单户授信小于100万（含）的境内小微型企业贷款户数")
        
        If DoCompare_AddMyComplexExplain_IsThisTargetRule(ct.rule, index) Then
            isText = False
            If DoCompare_AddMyComplexExplain_IsNotChecked(oldRule, index) Then
                '替换上期
                With reg
                    .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
                End With
                Set mc = reg.Execute(tRule)
                For Each m In mc
                    tstr = DoCompare_AddMyComplexExplain_ChangeTarge(CStr(m.value), index)
                    If dPreDataRef.exists(tstr) Then
                        '判断是否是文本
                        
                        If Find_Target_Pro(ycMid(CStr(m.value), ",", 3, 0, ",", 4, 0), "数据属性") = "文字" Then
                            tRule = replace(tRule, CStr(m.value), replace(Chr(34) & dPreDataRef(tstr) & Chr(34), "'", ""))
                            isText = True
                        Else
                            '存在则替换相应的值
                            If ycV(dPreDataRef(tstr)) = 0 Then
                                '存在，又为0的话，当做有一个很小的数。
                                '这里因为后期会用round添上，故改为用0.01元
                                'tRule = replace(tRule, CStr(m.value), ycV(dbl_Epsilon))
                                tRule = replace(tRule, CStr(m.value), ycV(0.01))
                            Else
                                tRule = replace(tRule, CStr(m.value), DoCompare_AddMyComplexExplain_ChangeData(CStr(m.value), ycV(dPreDataRef(tstr))))
                            End If
                        End If
                    Else
                        If Find_Target_Pro(ycMid(CStr(m.value), ",", 3, 0, ",", 4, 0), "数据属性") = "文字" Then
                            tRule = replace(tRule, CStr(m.value), Chr(34) & "" & Chr(34))
                            isText = True
                        Else '否则用0替换
                            tRule = replace(tRule, CStr(m.value), 0)
                        End If
                    End If
                    '添加规则
                    Call DoCompare_AddMyComplexExplain_AddHadChecked(ct.ruleEx, oldRule, tstr)
                Next
                
                With reg
                    .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
                End With
                Set mc = reg.Execute(tRule)
                For Each m In mc
                    tstr = DoCompare_AddMyComplexExplain_ChangeTarge(CStr(m.value), index)
                    If dCurDataRef.exists(tstr) Then
                        If InStr(ct.rule, "C101-R007C001") > 0 Then
                            Debug.Print ct.rule
                        End If
                        '判断是否是文本
                        If Find_Target_Pro(ycMid(CStr(m.value), ",", 3, 0, ",", 4, 0), "数据属性") = "文字" Then
                            tRule = replace(tRule, CStr(m.value), replace(Chr(34) & dCurDataRef(tstr) & Chr(34), "'", ""))
                            isText = True
                        Else
                            '存在则替换相应的值
                            If ycV(dCurDataRef(tstr)) = 0 Then
                                '存在，又为0的话，当做有一个很小的数。
                                '这里因为后期会用round添上，故改为用0.01元
                                'tRule = replace(tRule, CStr(m.value), ycV(dbl_Epsilon))
                                tRule = replace(tRule, CStr(m.value), ycV(0.01))
                            Else
                                '存在，又为0的话，当做有一个很小的数。
                                tRule = replace(tRule, CStr(m.value), DoCompare_AddMyComplexExplain_ChangeData(CStr(m.value), ycV(dCurDataRef(tstr))))
                            End If
                            'tRule = replace(tRule, CStr(m.value), ycV(dCurDataRef(tStr)))
                        End If
                    Else
                        If Find_Target_Pro(ycMid(CStr(m.value), ",", 3, 0, ",", 4, 0), "数据属性") = "文字" Then
                            tRule = replace(tRule, CStr(m.value), Chr(34) & "" & Chr(34))
                            isText = True
                        Else '否则用0替换
                            tRule = replace(tRule, CStr(m.value), 0)
                        End If
                    End If
                    Call DoCompare_AddMyComplexExplain_AddHadChecked(ct.ruleEx, oldRule, tstr)
                Next
                
                tRule = replace(tRule, "Thd", CStr(ct.thdV))
                tRule = replace(tRule, "thd", CStr(ct.thdV))
                '这里有一个当前结果
                
                '执行结果
                'result = Evaluate(tRule)
                'Debug.Assert (strSt <> "12D26")
                '如果不含文本的
                If InStr(oldRule, "C410") > 0 Then
                    Debug.Print ""
                End If
                roundRule = tRule
                If isText = False Then roundRule = DoCompare_AddMyComplexExplain_RoundRule(tRule)
                result = Evaluate(roundRule)
                jsStr = "计算过程：【" & oldRule & "】->【" & tRule & "】"
                
                '兜底回退
                If IsError(result) Or err.Number <> 0 Then
                    result = Evaluate(tRule)
                    jsStr = "回退计算过程：【" & oldRule & "】->【" & tRule & "】"
                    Call gloDe.InsertNewDebug("系统导数比较", CStr(arrCur(index, 4)), strSt, "校验回退", "执行[" & ct.ruleEx & "]出现回退")
                End If
                
                
                'Debug.Print jsStr
                
                str = ""
                If err.Number <> 0 Then
                    str = " 计算【" & ct.ruleEx & "】校验时发生错误" & "--" & err.Description & "。" & jsStr
                    jsStr = " 计算【" & ct.ruleEx & "】校验时发生错误" & "--" & err.Description & "。" & jsStr
                    err.Clear
                Else
                    If IsError(result) Then
                        '条件设置错误
                        str = " 计算【" & ct.ruleEx & "】校验时发生错误" & "--" & "。" & jsStr
                        jsStr = " 计算【" & ct.ruleEx & "】校验时发生错误" & "--" & "。" & jsStr
                    ElseIf VarType(result) = vbBoolean Then
                        If ct.qfbz = False Then
                            '不取反，表示满足则输出
                            If result = True Then
                                str = "【" & ct.ruleEx & "】。" & jsStr
                            End If
                        Else
                            If result = False Then
                                str = "【" & ct.ruleEx & "】。" & jsStr
                            End If
                        End If
                    Else
                        str = " 计算【" & ct.ruleEx & "】" & "校验结果并不是逻辑值:" & CStr(result) & "--" & "。" & jsStr
                        jsStr = " 计算【" & ct.ruleEx & "】" & "校验结果并不是逻辑值:" & CStr(result) & "--" & "。" & jsStr
                    End If
                End If
                If str <> "" Then
                    If InStr(ct.ruleEx, "软性") > 0 Then
                        If VarType(arrCur(index, 是否说明列)) = vbEmpty Or CStr(arrCur(index, 是否说明列)) = "" Then
                            wsCur.Cells(index, 是否说明列).Interior.colorindex = Complex_Soft_Color_Index
                        End If
                    Else
                        wsCur.Cells(index, 是否说明列).Interior.colorindex = 0
                    End If

                    If VarType(arrCur(index, 是否说明列)) = vbEmpty Or CStr(arrCur(index, 是否说明列)) = "" Then
                        arrCur(index, 是否说明列) = ct.ruleEx
                    Else
                        arrCur(index, 是否说明列) = CStr(arrCur(index, 是否说明列)) & "    ||    " & ct.ruleEx
                    End If
                    
                    If VarType(arrCur(index, 计算过程列)) = vbEmpty Or CStr(arrCur(index, 计算过程列)) = "" Then
                        arrCur(index, 计算过程列) = replace(replace(jsStr, "左中", "["), "右中", "]")
                    Else
                        arrCur(index, 计算过程列) = CStr(arrCur(index, 计算过程列)) & "    ||    " & replace(replace(jsStr, "左中", "["), "右中", "]")
                    End If

    
                End If
            End If  '判断是否执行过
        End If '若是本指标的校验才执行
    Next i
    '获取校验规则
    'DoCompare_AddMyComplexExplain = relStr
End Function

Sub testround()
    Dim rule As String
    'and(abs(round(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600,2))>round(.01,2),round(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600,2)<0
    'and(or(3.4-3.2>9,4-3.000>4-3),(3+(3.4-3.2))/2-2>4)
    'abs(11097336000-327600-0-11097008400-0)<=.01
    arrComplexeKeys = Split(gloCft.GetConfig("系统导数比较-复杂校验-公式精度处理关键字数组"), ",")
    'rule = "and(abs(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600)>.01,28449922100-15634092600-0-474520000-750127600-9831475200-1759706600<0)"
    'rule = "and(abs(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600)>.01,28449922100-15634092600-0-474520000-750127600-9831475200-1759706600>0)"
    '(28449922100-15634092600-0)<(474520000+750127600+9831475200+1759706600)
    'rule = "(28449922100-15634092600-0)<(474520000+750127600+9831475200+1759706600)"
    rule = "0<1"
    rule = DoCompare_AddMyComplexExplain_RoundRule(rule)
    Debug.Print (rule)
End Sub

'
'   将两边数值加个round(x,2)
'
Public Function DoCompare_AddMyComplexExplain_RoundRule(rule As String) As String
    On Error GoTo errhandle
    Dim strbk As String, rel As String
    strbk = rule
    
    '去掉空格就方便了
    rule = LCase(replace(rule, " ", ""))
    
    Dim i As Long, strKey As String
    '1 - 存放关键字，2 - 存放 "("
    Dim arr1 As Variant, n1 As Long, arr2 As Variant, n2 As Long, s As Long
    Dim lkh As Integer  '0-非关键字括号,1 - 关键字括号,
    Dim relust As String    '存放位置结果
    Dim zd As Boolean

    relust = ""
    i = 1
    n1 = 0
    n2 = 0
    s = 0
    lkh = 0
    ReDim arr1(1 To (Len(rule) - Len(replace(rule, "(", "")) + 1)) As Variant
    ReDim arr2(1 To (Len(replace(rule, "(", "")) + 1)) As Variant
    
    Do While i <= Len(rule)
        lkh = -1
        zd = False
        strKey = ""
        '遍历rule
        strKey = DoCompare_AddMyComplexExplain_RoundRule_GetKeyType(rule, i)
        If strKey <> "" Then
            '如果是关键字，则读取，并快进
            n1 = n1 + 1
            arr1(n1) = i
            i = i + Len(strKey) - 1
        ElseIf Mid(rule, i, 1) = "(" Then
            '如果是关键字，则读取，并快进
            n2 = n2 + 1
            arr2(n2) = i
        ElseIf DoCompare_AddMyComplexExplain_RoundRule_IsNumber(Mid(rule, i, 1)) Then
            s = s + 1
        Else
            If Mid(rule, i, 1) = ")" Then
                If n1 = 0 And n2 = 0 Then
                    GoTo errhandle
                ElseIf n1 = 0 And n2 > 0 Then
                    n2 = n2 - 1
                    s = s + 2
                    lkh = 0
                ElseIf n1 > 0 And n2 = 0 Then
                    '相当于中断结清
                    n1 = n1 - 1
                    zd = True
                    lkh = 1
                Else
                    If CLng(arr1(n1)) > CLng(arr2(n2)) Then
                        n1 = n1 - 1
                        zd = True
                    Else
                        n2 = n2 - 1
                        s = s + 2
                    End If
                End If
            Else
                '到这里就是中断
                zd = True
            End If
        End If
        
        If zd = True Or i = Len(rule) Then
            '结清
            If s > 0 Then
                
                If (i = Len(rule) And DoCompare_AddMyComplexExplain_RoundRule_IsNumber(Mid(rule, i, 1))) Or _
                    (i = Len(rule) And Mid(rule, i, 1) = ")" And lkh = 0) Then
                    If relust = "" Then
                        relust = CStr(i - s + 1) & "," & CStr(s)
                    Else
                        relust = relust & "," & CStr(i - s + 1) & "," & CStr(s)
                    End If
                Else
                    If relust = "" Then
                        relust = CStr(i - s) & "," & CStr(s)
                    Else
                        relust = relust & "," & CStr(i - s) & "," & CStr(s)
                    End If
                End If
                s = 0
            End If
        End If
        '下一个循环
        i = i + 1
nextI:

    Loop
    'Debug.Print relust
    '拆分出来了，下来就赋值
    Dim arr As Variant, j As Long
    arr = Split(relust, ",")
    i = 1
    j = 0
    rel = ""
    Do While j <= UBound(arr) Or i <= Len(rule)
        If j > UBound(arr) Then
            'j > 的话，说明剩下的都是
            rel = rel & Mid(rule, i, Len(rule) - i + 1)
            i = Len(rule)
            Exit Do
        ElseIf CLng(arr(j)) > i Then
            rel = rel & Mid(rule, i, CLng(arr(j)) - i)
            i = CLng(arr(j))
        Else
            rel = rel & "round(" & Mid(rule, CLng(arr(j)), CLng(arr(j + 1))) & ",2)"
            i = CLng(arr(j)) + CLng(arr(j + 1))
            j = j + 2
        End If
        If i > Len(rule) Then Exit Do
    Loop
    DoCompare_AddMyComplexExplain_RoundRule = rel
Exit Function

errhandle:
    DoCompare_AddMyComplexExplain_RoundRule = strbk
End Function

'
'   获取关键字
'
Private Function DoCompare_AddMyComplexExplain_RoundRule_GetKeyType(rule As String, i As Long) As String
    Dim str As String, j As Long
    
    For j = 0 To UBound(arrComplexeKeys) Step 1
        str = CStr(arrComplexeKeys(j))
        
        If Len(rule) - i + 1 >= Len(str) And Mid(rule, i, Len(str)) = str Then
            DoCompare_AddMyComplexExplain_RoundRule_GetKeyType = str
            Exit Function
        End If
    Next j
    
    DoCompare_AddMyComplexExplain_RoundRule_GetKeyType = ""
End Function

'
'   数值计算
'
Private Function DoCompare_AddMyComplexExplain_RoundRule_IsNumber(str As String) As Boolean
    If str = "0" Or str = "1" Or str = "2" Or str = "3" Or str = "4" Or str = "5" Or str = "6" _
        Or str = "7" Or str = "8" Or str = "9" Or str = "." Or str = "+" Or str = "-" Or str = "*" Or str = "/" Then
        DoCompare_AddMyComplexExplain_RoundRule_IsNumber = True
    Else
        DoCompare_AddMyComplexExplain_RoundRule_IsNumber = False
    End If
End Function


'
'   添加校验到读取的里面去,只添加软性的
'
Private Sub DoCompare_AddMyComplexExplain_AddHadChecked(ruleEx As String, rule As String, target As String)
    If InStr(ruleEx, "软性") > 0 Then
        Dim d As Object
        Set d = Nothing
        If dComplexChecked.exists(rule) Then
            Set d = dComplexChecked(rule)
            If VarType(d) = vbEmpty Then Set d = CreateObject("scripting.dictionary")
        Else
            Set d = CreateObject("scripting.dictionary")
        End If
        d(target) = 1
        Set dComplexChecked(rule) = d
    End If
End Sub

'
'   判断校验规则是否需要执行
'
Private Function DoCompare_AddMyComplexExplain_IsThisTargetRule(ByRef cTrule As String, i As Long) As Boolean
    On Error Resume Next
    
    Dim result As Boolean, arr As Variant, t As String, reg As Object, arrCt As Variant, j As Long
    Dim c1 As String, c2 As String, c3 As String, c4 As String, c5 As String
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    
    result = False
    arrCt = Split(cTrule, "|")
    
    For j = 0 To UBound(arrCt) Step 1
         t = replace(CStr(arrCt(j)), "[", "")
         t = replace(t, "]", "")
         arr = Split(t, ",")   '[机构名称，地区，机构类别，指标代码，频度]
         '3 - 机构名称，7-产品品种，11-指标代码
          'arr 0 -，1，2
         c1 = CStr(arrCur(i, col_机构名称)): c2 = CStr(arrCur(i, col_地区)): c3 = CStr(arrCur(i, col_机构类别)): c4 = CStr(arrCur(i, col_指标编码)): c5 = CStr(arrCur(i, col_频度))
         
        
         '不需要检测指标代码，是因为，就是因为指标代码是他才会进入到这里
         'If CStr(arr(0)) <> "" And CStr(arr(0)) <> c1 Then GoTo NoThisTarget
         If CStr(arr(0)) <> "" Then
             With reg
                 .Global = True: .ignorecase = False: .Pattern = replace(replace(CStr(arr(0)), "左中", "["), "右中", "]")
                 If .test(c1) = False Then GoTo NoThisTarget
             End With
         End If
         
         If CStr(arr(1)) <> "" And CStr(arr(1)) <> c2 Then GoTo NoThisTarget
         'If CStr(arr(2)) <> "" And CStr(arr(2)) <> c3 Then GoTo NoThisTarget
         If CStr(arr(4)) <> "" And CStr(arr(4)) <> c5 Then GoTo NoThisTarget
         
         '构建正则匹配，若设置了产品品种，是否匹配
         If CStr(arr(2)) <> "" Then
             With reg
                 .Global = True: .ignorecase = False: .Pattern = replace(replace(CStr(arr(2)), "左中", "["), "右中", "]")
                 If .test(c3) = False Then GoTo NoThisTarget
             End With
         End If
         
         '构建正则匹配，若设置了地区代码，是否匹配
'         If CStr(arr(2)) <> "" Then
'             With reg
'                 .Global = True: .ignorecase = False: .Pattern = replace(replace(CStr(arr(2)), "左中", "["), "右中", "]")
'                 If .test(c3) = False Then GoTo NoThisTarget
'             End With
'         End If
        '能执行到这里，说明满足
        result = True
        cTrule = CStr(arrCt(j))
NoThisTarget:
    Next j
    
    'Debug.Print cTrule
    Set reg = Nothing
    DoCompare_AddMyComplexExplain_IsThisTargetRule = result
End Function

'
'   判断是否是此指标的校验,true 没检测过，false 检测过
'
Private Function DoCompare_AddMyComplexExplain_IsNotChecked(cTrule As String, i As Long) As Boolean
    On Error Resume Next
    
    If dComplexChecked.exists(cTrule) Then
        Dim str As String, d As Object

        str = CStr(arrCur(i, col_机构名称)) & CStr(arrCur(i, col_指标编码))    '值的唯一识别方式
        str = replace(str, "'", "")
        Set d = dComplexChecked(cTrule)
        
        If d.exists(str) Then
            DoCompare_AddMyComplexExplain_IsNotChecked = False
        Else
            DoCompare_AddMyComplexExplain_IsNotChecked = True
        End If
    Else
        DoCompare_AddMyComplexExplain_IsNotChecked = True
    End If
End Function

'
'   进行单位转换成元,targerRule是对应的指标，val是当前值
'
Private Function DoCompare_AddMyComplexExplain_ChangeData(targeRule As String, val As Double) As Double
    Dim arr As Variant, t As String, str As String, v As Double
    t = replace(targeRule, " ", "")
    t = replace(t, "[", ""): t = replace(t, "]", "")
    t = replace(t, "{", ""): t = replace(t, "}", "")
    arr = Split(t, ",")
    
    t = CStr(arr(COMPLEX_TARGET_INDEX))
    '判断是否需要转换
    
    Dim unit As Double
    unit = GetUnitVal(DesUnit)
    
    If is_target_change_data_unit(t) Then
        v = val * unit
    Else
        v = val
    End If
    
    DoCompare_AddMyComplexExplain_ChangeData = v
End Function

'
'   判断指标是否需要进行数值转换
'
Private Function is_target_change_data_unit(code As String) As Boolean
    On Error Resume Next
    Dim s As String
    
    s = Find_Target_Pro(code, "数据属性")
    If s = "文字" Or s = "百分数" Or s = "个数" Then
        is_target_change_data_unit = False
        Exit Function
    End If
    
    s = Find_Target_Pro(code, "是否不转换单位")
    If CBool(ycV(s)) = True Then
        is_target_change_data_unit = False
        Exit Function
    End If
    
    is_target_change_data_unit = True
End Function


'
'   返回对应的指标
'
Private Function DoCompare_AddMyComplexExplain_ChangeTarge(targeRule As String, i As Long) As String
    '获取本指标的一些基本属性，以应对不存在时进行替换
    Dim c1 As String, c2 As String, c3 As String, c4 As String, c5 As String
    
    On Error Resume Next
    
    c1 = CStr(arrCur(i, col_机构名称)): c2 = CStr(arrCur(i, col_地区)): c3 = CStr(arrCur(i, col_机构类别)): c4 = CStr(arrCur(i, col_指标编码)): c5 = CStr(arrCur(i, col_频度))
    
    Dim arr As Variant, t As String, str As String
    t = replace(targeRule, " ", "")
    t = replace(t, "[", ""): t = replace(t, "]", "")
    t = replace(t, "{", ""): t = replace(t, "}", "")
    arr = Split(t, ",")
    '1 - 业务类，3-机构类代码，5-地区代码，8-指标代码，10 - 数据属性，11币种，12频度，13批次
     'arr 0 -业务类，1机构类代码，2 地区代码，3指标代码，4数据属性，5币种，6频度，7批次
     
    '如果校验那是空置的，则需要填充
    If CStr(arr(0)) <> "" Then c1 = CStr(arr(0))  '这里得用金融机构编码了
    If CStr(arr(1)) <> "" And InStr(CStr(arr(1)), "^") <= 0 Then c2 = CStr(arr(1))
    If CStr(arr(2)) <> "" And InStr(CStr(arr(2)), "^") <= 0 Then c3 = CStr(arr(2))
    If CStr(arr(3)) <> "" Then c4 = CStr(arr(3))
    If CStr(arr(4)) <> "" Then c5 = CStr(arr(4))
    
    str = CStr(arrCur(i, col_机构名称)) & c4 '这里两个就能识别唯一了，所以就不用弄这么多
    
    DoCompare_AddMyComplexExplain_ChangeTarge = str
End Function







Function GetFreStep(fre As String) As Integer
    Dim step As Integer
    step = 1
    
    If fre = "月" Then
        step = 1
    ElseIf fre = "季" Then
        step = 3
    ElseIf fre = "年" Then
        step = 12
    ElseIf fre = "半年" Then
        step = 6
    Else
        step = 1
    End If
    
    GetFreStep = step
End Function


'  获取某个机构的某个参照属性
'
Private Function Find_Org_Pro(code As String, pro_name As String) As String
    Dim d As Object
    If dorg_def.exists(code) Then
        '
        Set d = dorg_def(code)
        If d.exists(pro_name) Then
            Find_Org_Pro = CStr(d(pro_name))
            Exit Function
        Else
            '属性参照被修改
            Find_Org_Pro = ""
        End If
    Else
        If dorg_name_def.exists(code) Then
            '
            Set d = dorg_name_def(code)
            If d.exists(pro_name) Then
                Find_Org_Pro = CStr(d(pro_name))
                Exit Function
            Else
                '属性参照被修改
                Find_Org_Pro = ""
            End If
        Else
            '指标参照不齐全
            Find_Org_Pro = ""
        End If
    End If
End Function


'
'  获取某个指标的某个参照属性
'
Private Function Find_Target_Pro(code As String, pro_name As String) As String
    Dim d As Object
    If dtarget_def.exists(code) Then
        '
        Set d = dtarget_def(code)
        If d.exists(pro_name) Then
            Find_Target_Pro = CStr(d(pro_name))
            If pro_name = "数据属性" And CStr(d(pro_name)) = "" Then
                Find_Target_Pro = "余额"
            End If
            Exit Function
        Else
            '属性参照被修改
            Find_Target_Pro = ""
        End If
    Else
        '指标参照不齐全
        Find_Target_Pro = ""
    End If
End Function

'
'   警戒区间校验
'
Private Sub 警戒区间校验()
    Dim i As Integer, wsAlert As Worksheet, row As Integer
    Dim 警戒值行数1, 警戒值行数2 As Integer
    Dim 警戒值是否正确
    
    Set wsAlert = ThisWorkbook.Worksheets("警戒区间")
    

    '警戒值设定逻辑检查
    With wsAlert
        If .Range("A1").value <> "默认警戒区间" Then
            err.Raise Number:=1001, Description:="警戒区间区域被破坏，请正确设置"
        ElseIf .Range("B1").End(xlDown).row <> .Range("A1").End(xlDown).row Then
            err.Raise Number:=1001, Description:="警戒区间数值上标、下标设置不全！"
        Else
            row = .Range("A1").End(xlDown).row
            ReDim arrAlert(row - 2)
            '检测变幅是否从小到大
            For i = 2 To row Step 1
                If i = 2 Then
                    arrAlert(i - 2).序号 = .Cells(i, 1).Value2
                    arrAlert(i - 2).变幅 = .Cells(i, 2).Value2
                    arrAlert(i - 2).备注文字 = .Cells(i, 3).Value2
                    arrAlert(i - 2).备注颜色 = .Cells(i, 3).Interior.colorindex
                    arrAlert(i - 2).是否整行填充 = CInt(val(.Cells(i, 4).Value2))
                Else
                    If .Cells(i, 2).Value2 > .Cells(i, 3).Value2 Then
                        err.Raise Number:=1001, Description:="警戒区间阀值顺序不对，应该从小到大"
                    Else
                        arrAlert(i - 2).序号 = .Cells(i, 1).Value2
                        arrAlert(i - 2).变幅 = .Cells(i, 2).Value2
                        arrAlert(i - 2).备注文字 = .Cells(i, 3).Value2
                        arrAlert(i - 2).备注颜色 = .Cells(i, 3).Interior.colorindex
                        arrAlert(i - 2).是否整行填充 = CInt(val(.Cells(i, 4).Value2))
                    End If
                End If
            Next i
        End If
    End With
End Sub

'
'   添加备注 mark 返回备注内容
Private Function AddbfMark(h As Double) As String
    If Round(h, 6) = 0 Then
        AddbfMark = ""
        Exit Function
    End If
    Dim i%, c%, str$
    c = UBound(arrAlert)
    
    str = ""
    For i = c To LBound(arrAlert) Step -1
        If Round(h - ycV(arrAlert(i).变幅) * 100, 10) >= 0 Then
            str = arrAlert(i).备注文字
            If Round(h - ycV(-0.99) * 100, 10) = 0 Then
                str = "缩小100倍及以上"
                Exit For
            End If
            Exit For
        End If
    Next i
    If i < LBound(arrAlert) Then
        str = "近缩小100倍及以上"
    End If
    
    AddbfMark = str
End Function



