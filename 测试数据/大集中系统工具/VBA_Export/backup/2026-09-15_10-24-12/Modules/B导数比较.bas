Attribute VB_Name = "B导数比较"
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
    bf As Double
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
Private Const DT As Boolean = True  '测量时间
Private Const dbl_Epsilon As Double = 2.22044604925031E-13


Private savePath As String  '最终存放路径
Private curPath As String   '当期文件名
Private prePath As String   '上期文件名
Private curPaths As String   '当期文件名
Private prePaths As String   '上期文件名

Private SrcUnit As String   '源数据单位
Private DesUnit As String   '目标数据单位
Private 分析文件数据单位 As String
Private Const 特殊指标数据单位 As String = "亿元"
Private CurDate As Date
Private PreDate As Date
Private dPreDate As Object  '日期字典
Private dFres As Object '频度字典
Private SforbidCol As Integer    '禁用标记列
Private CforbidCol As Integer    '禁用标记列

Private wbDes As Workbook   '输出的工作簿
Private wsCur As Worksheet  '输出的工作表
Private dUnitRef As Object '单位转换参照表  字典
Private dSpcRef As Object  '特殊指标表  字典
Private dComplexRef As Object   '复杂指标表 字典
Private dCurDataRef As Object  '当期表  字典
Private dPreDataRef As Object  '上期表  字典

Private dhzDataRef As Object    '集中度 分母
Private is_check_jzd As Boolean
Private d_not_jzd As Object

Private dComplexChecked As Object   '用于存储已校验过的规则
Private dOtherTarget As Object  '用于存储其他标记项目
Private OtherR1 As Integer
Private OtherR2 As Integer
Private OtherColorIndex As Integer
Private OtherFontColorIndex As Integer
Private dtargetIndex_ref As Object
Private dtargetNewName_ref As Object

Private arrCur() As Variant '当期表数组
Private arrPre() As Variant '上期表数组
Private arrCurRows As Long
Private arrPreRows As Long
Private arrAlert() As 警戒值    '警戒值数组
Private Complex_Soft_Color_Index As Integer '软性校验颜色

Private dtarget_tips As Object '指标提示
Private dtarget_accu As Object '发生额指标
    
Private mdcfg As Object      '配置字典

Private Const 辅助列 As Integer = 30 '辅助列 用 30列好了。
Private Const 上期值列 As Integer = 15
Private Const 增减额列 As Integer = 16
Private Const 环比列 As Integer = 17
Private Const 备注列 As Integer = 18
Private Const 是否说明列 As Integer = 19
Private Const 说明内容列 As Integer = 20
Private Const 硬性过程列 As Integer = 21
Private Const 软性描述列 As Integer = 22
Private Const 软性过程列 As Integer = 23
Private Const UnitRefWbsName = "单位不需要转换指标参照表"
Private Const COMPLEX_TARGET_INDEX = 3  '复杂校验指标代码的索引位置
Private Const COMPLEX_FRE_INDEX = 7
Private Const COMPLEXT_SPLIT_COUNT = 7   '复杂校验指标的分隔符数量
Private IsFres As Boolean
Private IsCheckTol As Boolean   '是否启用总伐值校验选项
Private CheckTolModle As Integer    '总阀值校验模式 1、2、0
Private bCheckTol_Value As Double   '绝对值伐值
Private bCheckTol_Value2 As Double   '绝对值伐值
Private bCheckTol_Range As Double   '绝对值幅度伐值
Private arrComplexeKeys As Variant
Private iOrgMode As Integer '1 - 人行版本，2 - 机构版本
Private Const FSE_PIANCHA As Double = 0.00001


'
' 全部替换掉
'
Sub refresh_check()
    Dim r As Variant
    r = MsgBox("执行校验更新会替换掉原文件中，目标文件中有的所有校验工作表，执行前，最好保存一副当前校验程序副本，以防止需要找回校验。", vbYesNo, Application.Caption)
    If r = vbYes Then
        
    Else
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False                              '关闭屏幕更新
    Application.CutCopyMode = False
    
    On Error Resume Next
    '选择源数据
    Dim fileName As Variant, f As Variant, wbSrc As Workbook, wsSrc As Worksheet, wsDes As Worksheet
    Dim arr As Variant, i_gx As Integer, msg As String
    
    i_gx = 0
    fileName = Application.GetOpenFilename(fileFilter:="所有文件 (*.*), *.*", title:="选择源数据", MultiSelect:=True)
    If VarType(fileName) = vbArray Or VarType(fileName) = 8200 Or VarType(fileName) = 8204 Then
        For Each f In fileName
            '遍历每个工作簿
            Set wbSrc = gloEt.GetWb(CStr(f))
            For Each wsSrc In wbSrc.Worksheets
                If gloEt.IsBlankSheet(wsSrc) = False Then
                    If InStr(wsSrc.name, "特殊指标-") > 0 Or InStr(wsSrc.name, "复杂校验-") > 0 Then
                        Set wsDes = gloEt.FindWs(ThisWorkbook, wsSrc.name, 0, True)
                        wsDes.Cells.ClearComments
                        wsDes.Cells.ClearContents
                        '全复制，包括批注
                        Call gloEt.CopyAndPaste(wsSrc, wsDes, gloEt.GetUsedRange(wsSrc).Address, False)
                        i_gx = i_gx + 1
                        If msg = "" Then
                            msg = wsDes.name
                        Else
                            msg = msg & "," & wsDes.name
                        End If
                    End If
                End If
            Next
            wbSrc.Close False
        Next
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True                              '关闭屏幕更新
    Application.CutCopyMode = True
    
    
    If err.Number <> 0 Then
        MsgBox "校验更新结束，共更新" & CStr(i_gx) & "个工作表，具体为：" & msg & "。存在异常情况,异常提示为:" & CStr(err.Description)
        err.Clear
    Else
        MsgBox "校验更新结束，共更新" & CStr(i_gx) & "个工作表，具体为：" & msg
    End If
End Sub


Sub 导数比较控制台2(sp As String)
    Dim r As Variant, preDateStr As String, pdpc As Variant, i As Integer
    
    On Error GoTo errHandle
    Call gloDe.CreateDebugTitle("系统导数比较开始")
    gloT = Timer
    Call gloCft.CloseSomething
    savePath = sp  '最终存放路径
    iOrgMode = glo_iOrgMode
    
    '
    '   如果要弄机构版本的话，简单无缝对接就是在这里处理好
    '
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
    

    SrcUnit = 导数比较窗体.源数据单位  '源数据单位
    DesUnit = 导数比较窗体.目标单位  '目标数据单位
    
    IsCheckTol = CBool(导数比较窗体.CB_TotalCheck.Value)
    is_check_jzd = CBool(导数比较窗体.CB_集中度.Value)
    
    If IsCheckTol = True Then
        Call gloDe.InsertNewDebug("系统导数比较", "-", "-", "初始化", "启用", "读取总伐值校验")
        If CBool(导数比较窗体.SB_Sel1.Value) = True Then
            CheckTolModle = 1
        ElseIf CBool(导数比较窗体.SB_Sel2.Value) = True Then
            CheckTolModle = 2
        Else
            CheckTolModle = 0
        End If
        '还需要转换单位！
        bCheckTol_Value = Abs(ycV(导数比较窗体.TextBox1.Value) * GetUnitVal("万元") / GetUnitVal(DesUnit))
        bCheckTol_Value2 = Abs(ycV(导数比较窗体.TextBox3.Value))
        bCheckTol_Range = Abs(ycV(导数比较窗体.TextBox2.Value))
    End If
    
    分析文件数据单位 = 导数比较窗体.说明文件单位
    Complex_Soft_Color_Index = gloCft.GetConfig("复杂校验软性颜色")
    
    Call 警戒区间校验
    SforbidCol = glo_SforbidCol
    CforbidCol = glo_CforbidCol
    Call InitSomething
    pdpc = dPreDate.keys
    For i = 0 To UBound(pdpc) Step 1
        If i = 0 Then
            preDateStr = CStr(pdpc(i)) & ":" & CStr(dPreDate(CStr(pdpc(i))))
        Else
            preDateStr = preDateStr & ";" & CStr(pdpc(i)) & ":" & CStr(dPreDate(CStr(pdpc(i))))
        End If
    Next i
    r = MsgBox("请确认：本期数据为" + CStr(CurDate) + "，上期数据为" + preDateStr, vbYesNo, Application.Caption)
    If r = vbYes Then
        Call DoCompare
        '追加宏代码
        Call AddMacroCode
        '结束
        Call EndCompare
    Else
        Call EraseSomthing
    End If
    Call gloDe.InsertNewDebug("系统导数比较结束")
    Call gloCft.OpenSomething
    MsgBox "执行结束，共运行" & CStr(Timer - gloT) & "秒。"
    
Exit Sub

errHandle:
    MsgBox "系统导数比较出错:" + CStr(err.Description)
    Call gloDe.InsertNewDebug("系统导数比较结束")
    Call gloCft.OpenSomething
End Sub

'
'   获取单位保留位数
'
Private Function GetRoundNum(str As String) As Integer
    Dim unit As Long
    unit = 2
    
    Select Case str
        Case "元"
            unit = 2
        Case "十元"
            unit = 3
        Case "百元"
            unit = 4
        Case "千元"
            unit = 5
        Case "万元"
            unit = 6
        Case "十万元"
            unit = 7
        Case "百万元"
            unit = 8
        Case "千万"
            unit = 9
        Case "亿元"
            unit = 10
    End Select
    
    GetRoundNum = unit
End Function

Private Sub DoCompare_Init_jzd()
    Dim d_temp As Object, i As Long, is_has As Boolean
    Set dhzDataRef = Nothing
    Set dhzDataRef = CreateObject("scripting.dictionary")
    is_has = False
    Dim code_fm As String
    code_fm = CStr(导数比较窗体.TB_分母.Value)
    Dim v_yz1 As Double, v_yz2 As Double, v_bf As Double, v_yz As Double
    v_yz1 = ycV(导数比较窗体.TextBox4.Value) * (GetUnitVal("万元") / GetUnitVal(DesUnit))
    v_yz2 = ycV(导数比较窗体.TextBox6.Value)
    v_bf = ycV(导数比较窗体.TextBox5.Value)
    Dim mode_and As Boolean
    If CBool(导数比较窗体.OB_集中度_或模式) = True Then
        mode_and = False
    Else
        mode_and = True
    End If
    
    Dim org_code As String, target_code As String, is_recode As Boolean, area_code As String
    For i = 2 To UBound(arrCur)
        is_recode = False
        org_code = replace(CStr(arrCur(i, 3)), ",", "")
        target_code = replace(CStr(arrCur(i, 8)), ",", "")
        area_code = replace(CStr(arrCur(i, 5)), ",", "")
        If org_code & area_code = code_fm Then
            v_yz = v_yz1
            If dUnitRef.exists(target_code) Then v_yz = v_yz2
            If CStr(arrCur(i, 11)) = "美元合计" Then v_yz = v_yz / 6.4   '汇率按照6.4
            
            If CStr(arrCur(i, 18)) = "本期无，上期有" Or CStr(arrCur(i, 18)) = "本期有，上期无" Then
                is_recode = True
            Else
                If mode_and = True Then
                    'dhzDataRef(指标代码) 行号，当期值，上期值，变动值，环比，备注
                    If Abs(ycV(arrCur(i, 增减额列))) > v_yz And Abs(ycV(arrCur(i, 环比列)) * 100) > v_bf Then is_recode = True
                Else
                    If Abs(ycV(arrCur(i, 增减额列))) > v_yz Or Abs(ycV(arrCur(i, 环比列)) * 100) > v_bf Then is_recode = True
                End If
            End If
            
            Set d_temp = CreateObject("scripting.dictionary")
            d_temp("行号") = i
            d_temp("当期值") = ycV(arrCur(i, 14))
            d_temp("上期值") = ycV(arrCur(i, 上期值列))
            d_temp("增减额") = ycV(arrCur(i, 增减额列))
            d_temp("环比") = ycV(arrCur(i, 环比列))
            d_temp("备注") = ycV(arrCur(i, 备注列))
            If is_recode = True Then
                d_temp("超出") = 1
                is_has = True
            Else
                d_temp("超出") = 0
            End If
            Set dhzDataRef(target_code & CStr(arrCur(i, 10)) & CStr(arrCur(i, 11)) & CStr(arrCur(i, 12)) & CStr(arrCur(i, 13))) = d_temp
        End If
    Next
    If is_has = False Then
        is_check_jzd = False   '如果没有后续也不用计算了
        Call gloDe.InsertNewDebug("系统导数比较，分母指标均未超过阈值，集中度校验不启用。")
    End If
End Sub

Private Sub DoCompare()
    Dim rnDes As Range
    Dim iRowCur&, i&, j&, k&, str$, nr As Long, strNr As String
    Dim bBaiBeiyuzhi As Double
    
    bBaiBeiyuzhi = GetUnitVal("万元") / GetUnitVal(DesUnit) * ycV(gloCft.GetConfig("100倍变动绝对值阈值(万元)"))
    
    On Error GoTo CompareErr
    
    'Set wsCur = wbDes.Worksheets(1) '获取输出的工作表
    iRowCur = UBound(arrCur)    '获取当期数据行数
    
    '初始化输出标题
    arrCur(1, 上期值列) = "上期值": arrCur(1, 增减额列) = "增减额(" & DesUnit & ")": arrCur(1, 环比列) = "环比"
    arrCur(1, 备注列) = "备注": arrCur(1, 是否说明列) = "是否说明": arrCur(1, 说明内容列) = "说明内容"
    arrCur(1, 硬性过程列) = "计算过程"  ': arrCur(1, 软性描述列) = "软性描述": arrCur(1, 软性过程列) = "软性过程"
    
    With wsCur
        For i = 2 To UBound(arrCur) Step 1
            str = arrCur(i, 1) & arrCur(i, 3) & arrCur(i, 5) & arrCur(i, 8) & arrCur(i, 10) & arrCur(i, 11) & arrCur(i, 12) & arrCur(i, 13)
            str = replace(str, "'", "")
            If dPreDataRef.exists(str) Then '如果存在，则填充上期值
                arrCur(i, 上期值列) = dPreDataRef(str)
                
                '若相减的绝对值比伐值要大，就说明不为0
                'If ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列)) <> 0 And Abs(ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列))) > dbl_Epsilon Then
                If Abs(ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列))) > dbl_Epsilon Then
                    'arrCur(i, 增减额列) = ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列))
                    arrCur(i, 增减额列) = Application.WorksheetFunction.Round(ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列)), GetRoundNum(DesUnit))
                End If
                
                If Abs(ycV(arrCur(i, 14)) - ycV(arrCur(i, 上期值列))) > dbl_Epsilon Then
                    If ycV(arrCur(i, 上期值列)) <> 0 Then
                        arrCur(i, 环比列) = arrCur(i, 增减额列) / ycV(arrCur(i, 上期值列))
                    Else
                        arrCur(i, 环比列) = "除数为0"
                    End If
                End If
                '添加备注
                arrCur(i, 备注列) = DoCompare_AddMark(wsCur, i, val(arrCur(i, 环比列)))
            Else    '不存在，本期有，上期无。红色 3
                arrCur(i, 上期值列) = ""
                arrCur(i, 增减额列) = arrCur(i, 14)
                arrCur(i, 备注列) = "本期有，上期无"
                .Range(.Cells(i, 1), .Cells(i, 备注列)).Interior.colorindex = 3
            End If
            
            If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在校验第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
        Next i
        
        '计算有多少本期无，上期有指标数量，并存为字符串
        Dim colnr As New Collection '定义s变量为集合对象
        
        For i = 2 To UBound(arrPre) Step 1
            str = arrPre(i, 1) & arrPre(i, 3) & arrPre(i, 5) & arrPre(i, 8) & arrPre(i, 10) & arrPre(i, 11) & arrPre(i, 12) & arrPre(i, 13)
            str = replace(str, "'", "")
            If Not dCurDataRef.exists(str) Then
                colnr.add (i)
            End If
        Next i
        
        '若存在本期无，上期有指标，从字符串中，取回指标
        If colnr.count > 0 Then
            arrCur = Application.WorksheetFunction.Transpose(arrCur)
            ReDim Preserve arrCur(1 To UBound(arrCur), 1 To (colnr.count + iRowCur))
            arrCur = Application.WorksheetFunction.Transpose(arrCur)
            
            For k = 1 To colnr.count Step 1
                '如果不存在，则添加到当期表
                i = CLng(colnr.item(k))
                arrPre(i, 2) = arrCur(2, 2)
                arrPre(i, 上期值列) = arrPre(i, 14)
                arrPre(i, 14) = ""
                arrPre(i, 增减额列) = 0 - ycV(arrPre(i, 上期值列))
                arrPre(i, 环比列) = -1
                arrPre(i, 备注列) = "本期无，上期有"
                iRowCur = iRowCur + 1
                .Range(.Cells(iRowCur, 1), .Cells(iRowCur, 备注列)).Interior.colorindex = 6

                '.Range(.Cells(iRowCur, 1), .Cells(iRowCur, 说明内容列)) = Application.Index(arrPre, i, 0)
                For j = 1 To 说明内容列 Step 1
                    arrCur(iRowCur, j) = arrPre(i, j)
                Next j
            Next k
        End If
        Set colnr = Nothing
        
        '先复制一次来排序,这样来保证校验顺序一致
        .Cells.ClearContents
    
        '格式设置
        .Range(.Cells(2, 2), .Cells(UBound(arrCur), 2)).NumberFormatLocal = "yyyy/m/d"
        .Range(.Cells(2, 3), .Cells(UBound(arrCur), 8)).NumberFormatLocal = "@"
        .Range(.Cells(2, 14), .Cells(UBound(arrCur), 增减额列)).NumberFormatLocal = "0.00"
        .Range(.Cells(2, 环比列), .Cells(UBound(arrCur), 环比列)).NumberFormatLocal = "0.00%"
        '.Range(.Cells(2, 环比列), .Cells(UBound(arrCur), 13)).NumberFormatLocal = "0"
        .Range("A1").Resize(UBound(arrCur), UBound(arrCur, 2)).Value = arrCur
        If UBound(arrCur) > 2 Then
            .Range("A2:Z" & CStr(UBound(arrCur)) - 1).Sort Key1:=.Range("C2"), Key2:=.Range("E2"), Key3:=.Range("H2"), Header:=xlNo
        End If
        
        '这里导致没有了
        arrCur = .Range("A1").Resize(UBound(arrCur), UBound(arrCur, 2))
        
        '初始化集中度数据
        If is_check_jzd = True Then Call DoCompare_Init_jzd
            
        '遍历所有数据，处理特殊指标
        For i = 2 To UBound(arrCur) Step 1
            '添加处理顺序码，有则替换，没有则使用指标代码
            If dtargetIndex_ref.exists(replace(arrCur(i, 8), "'", "")) Then
                arrCur(i, 7) = dtargetIndex_ref(replace(arrCur(i, 8), "'", ""))
            Else
                arrCur(i, 7) = arrCur(i, 8)
            End If
            
            ' 如果有得话，就提还
            If dtargetNewName_ref.exists(replace(arrCur(i, 8), "'", "")) Then
                arrCur(i, 9) = dtargetNewName_ref(replace(arrCur(i, 8), "'", ""))
            End If
            
            PreDate = CDate(dPreDate(CStr(arrCur(i, 12))))
            
            If DoCompare_is_Accu_Target(replace(arrCur(i, 8), "'", "") & CStr(arrCur(i, 10)) & CStr(arrCur(i, 12))) = True Then
                arrCur(i, 是否说明列) = 当年累计指标比上期不应减少(i)
            End If
            
            If dSpcRef.exists(replace(arrCur(i, 8), "'", "")) Then
                arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & DoCompare_AddMyExplain(replace(arrCur(i, 8), "'", ""), i)
            End If
            If dComplexRef.exists(replace(arrCur(i, 8), "'", "")) Then
                '直接挪到里面进行操作吧
                'arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & " " & DoCompare_AddMyComplexExplain(replace(arrCur(i, 8), "'", ""), i)
                Call DoCompare_AddMyComplexExplain(replace(arrCur(i, 8), "'", ""), i)
            End If
            
            ' 缩小100倍及以上 ,增幅100倍以上
            If CStr(arrCur(i, 备注列)) <> "" Then
                If InStr(CStr(arrCur(i, 备注列)), "100倍以上") > 0 Then
                    If Abs(ycV(arrCur(i, 增减额列))) > bBaiBeiyuzhi Then
                        If CStr(arrCur(i, 是否说明列)) <> "" Then
                            arrCur(i, 是否说明列) = CStr(arrCur(i, 是否说明列)) & ";绝对值变动超过" & CStr(Round(bBaiBeiyuzhi, 2)) & DesUnit & "且，" & CStr(arrCur(i, 备注列)) & "。"
                        Else
                            arrCur(i, 是否说明列) = CStr(arrCur(i, 备注列)) & "。"
                        End If
                    End If
                End If
            End If
            
            
            '指标名称包含其他的，需要标红
            Call Docompare_MarkOtherTarget(replace(arrCur(i, 8), "'", ""), CStr(arrCur(i, 9)), i)
            ' 总伐值校验
            If IsCheckTol = True Then Call Docompare_CheckTol(Abs(ycV(arrCur(i, 增减额列))), Abs(ycV(arrCur(i, 环比列))) * 100, i)
            If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在校验第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
        Next i
        '增加集中度计算
        If is_check_jzd = True Then Call Docompare_CheckJzd
        
        .Cells.ClearContents
        .Range("A1").Resize(UBound(arrCur), UBound(arrCur, 2)).Value = arrCur
        
        Call EraseSomthing
    End With
    Call gloCft.ShowStatusMessage_Up("")
Exit Sub

CompareErr:
    'Resume
    Call EraseSomthing
    str = "DoCompare:执行有误," & err.Description
    err.Clear
    Call gloCft.ShowStatusMessage_Up("")
    err.Raise Number:=1002, Description:=str
End Sub

'
'   这里已经在循环了,要么就放出去
'
Private Sub Docompare_CheckJzd()
    Dim str As String
    Dim code_fm As String
    code_fm = CStr(导数比较窗体.TB_分母.Value)
    
    Dim v_bz As Double
    v_bz = ycV(导数比较窗体.TextBox7.Value)
    
    Dim org_code As String, target_code As String, d_temp As Object, i As Long, area_code As String
    Dim v1 As Double, v2 As Double
    
    For i = 2 To UBound(arrCur)
        str = ""
        org_code = replace(CStr(arrCur(i, 3)), ",", "")
        target_code = replace(CStr(arrCur(i, 8)), ",", "")
        area_code = replace(CStr(arrCur(i, 5)), ",", "")
        If d_not_jzd.exists(target_code) Then GoTo nextTarget   '不参与的直接跳过
        If org_code & area_code <> code_fm Then
            '计算比重
            If dhzDataRef.exists(target_code & CStr(arrCur(i, 10)) & CStr(arrCur(i, 11)) & CStr(arrCur(i, 12)) & CStr(arrCur(i, 13))) Then
                Set d_temp = dhzDataRef(target_code & CStr(arrCur(i, 10)) & CStr(arrCur(i, 11)) & CStr(arrCur(i, 12)) & CStr(arrCur(i, 13)))
                If ycV(d_temp("超出")) = 1 Then
                    If ycV(d_temp("增减额")) <> 0 Then
                        v1 = ycV(arrCur(i, 16))
                        v2 = ycV(d_temp("增减额"))
                        If Abs(v1 / v2) * 100 > v_bz Then
                            '集中度超出
                            str = "母项变动较大，集中度校验超过" & CStr(v_bz) & "%。"
                            arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & str
                            ' 判断分母是否标识
                            If Not d_temp Is Nothing Then
                                If InStr(arrCur(d_temp("行号"), 是否说明列), str) <= 0 Then
                                    arrCur(d_temp("行号"), 是否说明列) = arrCur(d_temp("行号"), 是否说明列) & str
                                End If
                            End If
                            
                            '计算过程
                            str = "集中度计算过程：" & CStr(Round(v1, 2)) & "/" & CStr(Round(v2, 2))
                            str = str & " = " & CStr(Round(v1 / v2 * 100, 1)) & "% > " & CStr(v_bz) & "%。"
                            arrCur(i, 硬性过程列) = CStr(arrCur(i, 硬性过程列)) & str
                            
                            
                        End If
                    End If
                End If
            Else
                'arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & "集中度校验异常，分母无此指标数值。"
            End If
            
        End If
        If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在校验第：" & CStr(i) & "行/" & CStr(UBound(arrCur)) & "行")
nextTarget:
    Next
End Sub

'
'   做整体比较
'   已经是绝对值了
'
Private Sub Docompare_CheckTol(绝对值变化 As Double, 绝对值变幅 As Double, i As Long)
    Dim str As String, strSt As String, u As Integer, cv As Double
    strSt = replace(CStr(arrCur(i, 8)), "'", "")
    If dUnitRef.exists(strSt) Then
        cv = bCheckTol_Value2
    Else
        cv = bCheckTol_Value
        If CStr(arrCur(i, 11)) = "美元合计" Then
            cv = cv / 6.4   '汇率按照6.4
        End If
    End If
    str = ""

    If bCheckTol_Value <> 0 And bCheckTol_Range <> 0 Then
        If CheckTolModle = 1 Then
            If 绝对值变化 >= cv And 绝对值变幅 >= bCheckTol_Range Then
                str = "总伐值校验：绝对值变化超过" & CStr(cv) & CStr(DesUnit) & "/个 且绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
            End If
        ElseIf CheckTolModle = 2 Then
            If 绝对值变化 >= cv Or 绝对值变幅 >= bCheckTol_Range Then
                str = "总伐值校验：绝对值变化超过" & CStr(cv) & CStr(DesUnit) & "/个 或绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
            End If
        Else
        
        End If
    ElseIf bCheckTol_Value <> 0 Then
        If 绝对值变化 >= cv Then
            str = "总伐值校验：绝对值变化超过" & CStr(cv) & CStr(DesUnit) & "/个"
        End If
    ElseIf bCheckTol_Range <> 0 Then
        If 绝对值变幅 >= bCheckTol_Range Then
            str = "总伐值校验：绝对值变幅超过" & CStr(bCheckTol_Range) & "%"
        End If
    Else
    
    End If
    
    arrCur(i, 是否说明列) = arrCur(i, 是否说明列) & str
End Sub

Private Sub Docompare_MarkOtherTarget(code As String, name As String, i As Long)
    If OtherColorIndex = 0 Then Exit Sub
    Dim t As Boolean

    t = False
    With wsCur
        If dOtherTarget.exists(code) Then
            If CDbl(dOtherTarget(code)) = 0 Then
                t = True
            End If
        Else
            If InStr(name, "其他") Then
                t = True
            End If
        End If
    
        If t = True Then
'            If .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Interior.ColorIndex = -4142 Or _
'                .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Interior.ColorIndex = 0 Then
'                .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Interior.ColorIndex = OtherColorIndex
'            End If
'
'            If .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Font.ColorIndex = 0 Or _
'                .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Font.ColorIndex = -4142 Or _
'                .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Font.ColorIndex = -4105 Then
'                .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Font.ColorIndex = OtherFontColorIndex
'            End If
            .Range(.Cells(i, OtherR1), .Cells(i, OtherR2)).Interior.colorindex = OtherColorIndex
        End If
    End With
End Sub

Private Sub EraseSomthing()
    Set dSpcRef = Nothing
    Set dCurDataRef = Nothing
    Set dPreDataRef = Nothing
    Set dUnitRef = Nothing
    Set dComplexRef = Nothing
    Set dComplexChecked = Nothing
    Set dOtherTarget = Nothing
    Set dFres = Nothing
    Set dhzDataRef = Nothing
    Set dtargetIndex_ref = Nothing
    Set dtargetNewName_ref = Nothing
    Erase arrPre
    Erase arrCur
    Erase arrAlert
End Sub

Private Sub InitSomthing_FullWs(wsSrc As Worksheet, wsTemp As Worksheet)
    Dim r1 As Long, r2 As Long, c2 As Integer
    Call gloEt.ShowData(wsTemp)
    Call gloEt.ShowData(wsSrc)
    r2 = gloEt.GetCurRows(wsTemp)
    c2 = wsTemp.Range("A1").End(xlToRight).Column
    If r2 < 2 Then Exit Sub
    If c2 = 256 Then c2 = 14
    r1 = gloEt.GetCurRows(wsSrc)
    
    If gloEt.IsBlankSheet(wsSrc) Then
        wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(r2, c2)).Copy wsSrc.Cells(1, 1)
    Else
        wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(r2, c2)).Copy wsSrc.Cells(r1 + 1, 1)  'wsSrc.Range(wsSrc.Cells(r1 + 1, 1), wsSrc.Cells(r1 + 1 + r2, c2))
    End If
End Sub

Private Sub InitLevleRef(wsDes As Worksheet)
   Dim arr() As Variant, ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("指标参照")
    Call gloEt.ShowData(ws)
    With ws
        arr = .Range("A1:F" & CStr(.Range("B1").End(xlDown).row))
    End With
    
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)).NumberFormatLocal = "@"
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)) = arr
    wsDes.name = ws.name
    wsDes.Visible = xlSheetVeryHidden
End Sub


'
'   初始化源数据及输出工作簿
'   进行单位转换处理
'
Private Sub InitSomething()
    Dim wsSrc As Worksheet, wsDes As Worksheet, unitC As Double, arr As Variant, i As Integer, wsTemp As Worksheet
    Dim isCurDate As Boolean, isPreDate As Boolean
    '新建工作簿
    Set wbDes = Workbooks.add
    With wbDes
        Do While wbDes.Worksheets.count > 1
            wbDes.Worksheets(1).Delete
        Loop
    End With
    Set wsCur = wbDes.Worksheets(1)
    wsCur.name = "草稿"
    '复制【单位不需要转换指标参照表】表
    Call InitUnitRef(wbDes.Worksheets.add)
    '复制【机构参照表】
    Call InitOrgRef(wbDes.Worksheets.add)
    '复制【指标提示】
    Call InitTargetTips(wbDes.Worksheets.add)
    Call InitLevleRef(wbDes.Worksheets.add)
    
    
    'Call InitComplexRef
    Call Init_not_jzd_target
    
    Call InitdOtherTargetRef
    
    Call InitdtargetIndex_ref
    
    Call Init_Target_Tips  'dtarget_tips
    
    '初始化当年累计发生额指标
    Call Init_Target_Accu
    
    unitC = GetUnitVal(SrcUnit) / GetUnitVal(DesUnit)
    '初始化当期数据
    arr = Split(curPaths, ",")
    For i = 0 To UBound(arr) Step 1
        If i = 0 Then
            Set wsSrc = gloEt.GetWb(CStr(arr(i))).Worksheets(1)
            Call CheckDateUnieque(wsSrc)
        Else
            Set wsTemp = gloEt.GetWb(CStr(arr(i))).Worksheets(1)
            Call CheckDateUnieque(wsTemp)
            '追加
            Call InitSomthing_FullWs(wsSrc, wsTemp)
            wsTemp.Parent.Close SaveChanges:=False
        End If
    Next i
    curPath = CStr(arr(0))
    
    Set dFres = Nothing
    Set dFres = CreateObject("scripting.dictionary")
    dFres("") = 1
    
    isCurDate = InitCurData(unitC, wsSrc)
    'If isCurDate = False Then err.Raise Number:=1001, Description:="当期无数，执行退出。"
     '获取上期数据表
    arr = Split(prePaths, ",")
    For i = 0 To UBound(arr) Step 1
        If i = 0 Then
            Set wsSrc = gloEt.GetWb(CStr(arr(i))).Worksheets(1)
            Call CheckDateUnieque(wsSrc)
        Else
            Set wsTemp = gloEt.GetWb(CStr(arr(i))).Worksheets(1)
            Call CheckDateUnieque(wsTemp)
            '追加
            Call InitSomthing_FullWs(wsSrc, wsTemp)
            wsTemp.Parent.Close SaveChanges:=False
        End If
    Next i
    prePath = CStr(arr(0))
    isPreDate = InitPreData(unitC, wsSrc)
    
    '当期日期小于上期日期
    If CurDate < PreDate Then
        Call EraseSomthing
        err.Raise Number:=1001, Description:="本期数据日期" & CStr(CurDate) & " 小于等于上期数据日期" & CStr(PreDate)
    End If
    
    '获取特殊指标字典
    Call InitSpecialRef
    
    '这里才初始化复杂校验
    arrComplexeKeys = Split(gloCft.GetConfig("系统导数比较-复杂校验-公式精度处理关键字数组"), ",")
    Call InitComplexRef
    
    With wbDes.Worksheets.add
        .name = "宏代码"
        .Visible = xlSheetHidden
        .Cells(1, 1).Value = gloCft.GetConfig("说明文件菜单模块宏代码")
        .Cells(2, 1).Value = gloCft.GetConfig("输出说明单元格处理宏")
        .Cells(3, 1).Value = gloCft.GetConfig("输出说明word处理宏")
        .Cells(1, 2).Value = gloCft.GetConfig("说明文件菜单模块宏代码_日报")
        .Cells(2, 2).Value = gloCft.GetConfig("work_ri_shuoming")
    End With
End Sub

Private Sub InitdtargetIndex_ref()
    On Error Resume Next
    Set dtargetIndex_ref = Nothing
    Set dtargetNewName_ref = Nothing
    Set dtargetIndex_ref = CreateObject("scripting.dictionary")
    Set dtargetNewName_ref = CreateObject("scripting.dictionary")
    
    Dim i As Long, arr As Variant
    arr = ThisWorkbook.Worksheets("指标顺序参照").Range("A1").CurrentRegion.Value
    
    For i = 2 To UBound(arr)
        dtargetIndex_ref(replace(CStr(arr(i, 2)), "'", "")) = CStr(arr(i, 1)) & "_" & Format(ycV(arr(i, 4)), "0000")
        If CStr(arr(i, 5)) <> "" Then
            dtargetNewName_ref(replace(CStr(arr(i, 2)), "'", "")) = CStr(arr(i, 5))
        End If
    Next
End Sub

Private Sub Init_Target_Accu()
    On Error Resume Next
    Set dtarget_accu = Nothing
    Set dtarget_accu = CreateObject("scripting.dictionary")
    
    Dim i As Long, arr As Variant
    arr = ThisWorkbook.Worksheets("当年累计指标").Range("A1").CurrentRegion.Value
    
    For i = 2 To UBound(arr)
        dtarget_accu(CStr(arr(i, 1)) & CStr(arr(i, 3)) & CStr(arr(i, 4))) = 1
    Next
End Sub

Private Sub Init_Target_Tips()
    On Error Resume Next
    Set dtarget_tips = Nothing
    Set dtarget_tips = CreateObject("scripting.dictionary")
    
    Dim i As Long, arr As Variant
    arr = ThisWorkbook.Worksheets("指标提示").Range("A1").CurrentRegion.Value
    For i = 2 To UBound(arr)
        If CStr(arr(i, 5)) = "" Then
            dtarget_tips(CStr(arr(i, 1))) = CStr(arr(i, 4))
        End If
    Next
End Sub

Public Function IsAbnormalData(str As String, v As Variant, name As String) As Boolean
    Dim sfdu As Boolean
    sfdu = False
    
    If VarType(v) = vbString Then
        If CStr(v) = "" Then
            Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "空字符串。")
            sfdu = True
            Exit Function
        ElseIf val(v) = 0 Then
            Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "纯文本型字符串。")
            sfdu = True
        Else
            If CStr(val(v)) = CStr(v) Then
                Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "异常，有纯数值文本，转换数值型读入。")
            Else
                Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "异常，有非纯数值文本，转换数值型读入。")
                sfdu = True
            End If
        End If
    ElseIf VarType(v) = vbEmpty Then
        Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "有空单元格。")
        sfdu = True
    ElseIf VarType(v) = vbError Then
        Call gloDe.InsertNewDebug("系统导数比较", "-", name, "读取数据错误", "读取指标" & str & "有错误项目。")
        sfdu = True
    Else

    End If
    
    IsAbnormalData = sfdu
End Function


'
'   初始化当期数据
'
Private Function InitCurData(unitC As Double, wsSrc As Worksheet) As Boolean
    Dim i As Long, str$, tDate As Date
    
    Call gloEt.ShowData(wsSrc)
    
    
    arrCur = wsSrc.Range("A1:W" & CStr(wsSrc.Range("A1").CurrentRegion.rows.count))
    
    Set dCurDataRef = Nothing
    Set dCurDataRef = CreateObject("scripting.dictionary")
    
    If UBound(arrCur) < 2 Then
        CurDate = CDate(Date)
        InitCurData = False
        Exit Function
    End If
    
    Call CheckDateUnieque(wsSrc)
    
    '这里决定的
    arrCur = wsSrc.Range("A1:W" & CStr(wsSrc.Range("A1").End(xlDown).row))
    CurDate = CDate(arrCur(2, 2))

    wsSrc.Parent.Close SaveChanges:=False
    
    
    
    For i = 2 To UBound(arrCur) Step 1
        If Not dFres.exists(CStr(arrCur(i, 12))) Then dFres(CStr(arrCur(i, 12))) = 1
        If Not dFres.exists(CStr(arrCur(i, 12)) & CStr(arrCur(i, 13))) Then dFres(CStr(arrCur(i, 12)) & CStr(arrCur(i, 13))) = 1
        '生成字典
        str = arrCur(i, 1) & arrCur(i, 3) & arrCur(i, 5) & arrCur(i, 8) & arrCur(i, 10) & arrCur(i, 11) & arrCur(i, 12) & arrCur(i, 13)
        str = replace(str, "'", "")
        
        '如果不存在，则需要进行单位转换
        If Not dUnitRef.exists(replace(arrCur(i, 8), "'", "")) Then
            '正常项目才进行转换
            If IsAbnormalData(str, arrCur(i, 14), curPath) = False Then
                arrCur(i, 14) = val(arrCur(i, 14)) * unitC
            End If
        End If
        dCurDataRef(str) = arrCur(i, 14)
    Next i
    InitCurData = True
End Function



'
'   初始化上期数据
'
Private Function InitPreData(unitC As Double, wsSrc As Worksheet) As Boolean
    Dim i As Long, str$, tDate As Date
    
    Call gloEt.ShowData(wsSrc)
    
    'arrPre = wsSrc.Range("A1:W" & CStr(wsSrc.Range("A1").End(xlDown).row))
    
    arrPre = wsSrc.Range("A1:W" & CStr(wsSrc.Range("A1").CurrentRegion.rows.count))
    Set dPreDate = Nothing
    Set dPreDate = CreateObject("scripting.dictionary")
    
    If UBound(arrPre) < 2 Then
        PreDate = CDate(Date)
        InitPreData = False
        Exit Function
    End If
    
    IsFres = False
    
    arrPre = wsSrc.Range("A1:W" & CStr(wsSrc.Range("A1").End(xlDown).row))
    '读取上期频度日期
    For i = 2 To UBound(arrPre) Step 1
        '读取频度
        If Not dFres.exists(CStr(arrPre(i, 12))) Then dFres(CStr(arrPre(i, 12))) = 1
        If Not dFres.exists(CStr(arrPre(i, 12)) & CStr(arrPre(i, 13))) Then dFres(CStr(arrPre(i, 12)) & CStr(arrPre(i, 13))) = 1
        
        If VarType(ChangeToCDate(arrPre(i, 2))) = vbDate Then
            PreDate = CDate(arrPre(i, 2))
        Else
            PreDate = Date
            Call gloDe.InsertNewDebug("系统导数比较", "-", "-", "初始化", "上期数据有误", "上期文件数据日期格式有误")
        End If
        
        '当期日期小于上期日期
        If CurDate < PreDate Then
            Call EraseSomthing
            wsSrc.Parent.Close SaveChanges:=False
            err.Raise Number:=1001, Description:="本期数据日期" & CStr(CurDate) & " 小于等于上期数据日期" & CStr(PreDate)
        End If
        
        If dPreDate.exists(CStr(arrPre(i, 12))) Then
            tDate = dPreDate(CStr(arrPre(i, 12)))
            If VarType(tDate) = vbEmpty Or CLng(tDate) = 0 Then
                dPreDate(CStr(arrPre(i, 12))) = PreDate
            Else
                If tDate <> PreDate Then
                    Call EraseSomthing
                    wsSrc.Parent.Close SaveChanges:=False
                    err.Raise Number:=1001, Description:="上期文件数据日期，同频度数据，日期不相同。"
                End If
            End If
        Else
            dPreDate(CStr(arrPre(i, 12))) = PreDate
        End If
    Next i
    
    
    wsSrc.Parent.Close SaveChanges:=False

    Set dPreDataRef = Nothing
    Set dPreDataRef = CreateObject("scripting.dictionary")

    For i = 2 To UBound(arrPre) Step 1
        '生成字典
        str = arrPre(i, 1) & arrPre(i, 3) & arrPre(i, 5) & arrPre(i, 8) & arrPre(i, 10) & arrPre(i, 11) & arrPre(i, 12) & arrPre(i, 13)
        str = replace(str, "'", "")
        
        '不存在则进行单位转换
        If Not dUnitRef.exists(replace(arrPre(i, 8), "'", "")) Then
            '正常项目才进行转换
            If IsAbnormalData(str, arrPre(i, 14), prePath) = False Then
                arrPre(i, 14) = ycV(arrPre(i, 14)) * unitC
            End If
        End If
        dPreDataRef(str) = arrPre(i, 14)
    Next i
    
    If dFres.count > 1 Then IsFres = True
    
    InitPreData = True
End Function

'
'   初始化单位不需要转换指标参照表字典
'
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

Private Sub InitTargetTips(wsDes As Worksheet)
    Dim arr() As Variant, ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("指标提示")
    Call gloEt.ShowData(ws)
    With ws
        arr = .Range("A1").CurrentRegion.Value
    End With
    
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)).NumberFormatLocal = "@"
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)) = arr
    wsDes.name = ws.name
    wsDes.Visible = xlSheetVeryHidden
End Sub

'
' 复制机构参照表
'
Public Sub InitOrgRef(wsDes As Worksheet)
    Dim arr() As Variant, ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("机构参照表")
    Call gloEt.ShowData(ws)
    With ws
        arr = .Range("A1:F" & CStr(.Range("B1").End(xlDown).row))
    End With
    
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)).NumberFormatLocal = "@"
    wsDes.Range("A1").Resize(UBound(arr), UBound(arr, 2)) = arr
    wsDes.name = ws.name
    wsDes.Visible = xlSheetVeryHidden
End Sub

'
'   初始化特殊指标表字典
'
Private Sub InitSpecialRef()
    Dim ws As Worksheet, i%, j%, k%, arrSpcRef() As Variant, arrT() As String, code As String, str As String
    
    Dim wsFre As String
    
    Set dSpcRef = CreateObject("scripting.dictionary")
    
    On Error GoTo errHandle
    '指标代码，指标名，备注，人民币伐值，美元合计伐值，数据属性，说明，表单
    Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer, c5 As Integer, c6 As Integer, c7 As Integer, c8 As Integer
    c1 = 1: c2 = 3: c3 = 4: c4 = 5: c5 = 6: c6 = 7: c7 = 8: c8 = 12
    
    '遍历所有特殊指标表
    Dim mycnt As Control
    For Each mycnt In 导数比较窗体.CheckFrame.Controls
        If TypeName(mycnt) = "CheckBox" And mycnt.Value = True Then
            '若选了，则
            Set ws = ThisWorkbook.Worksheets(mycnt.name)
            
            If InStr(ws.name, "特殊指标") > 0 Then
                '获取报表频度
                wsFre = replace(replace(ws.name, "特殊指标-", ""), " ", "")
                If Month(CurDate) = 1 And Day(CurDate) = 1 And InStr(ws.name, "结转数") > 0 Then
                    Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "读取结转数校验")
                ElseIf Month(CurDate) = 1 And Day(CurDate) = 1 And InStr(ws.name, "结转数") <= 0 Then
                    '这里要跳出
                    Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "结转期，不读取该校验")
                    GoTo nextWs
                ElseIf Day(CurDate) <> 1 And InStr(ws.name, "结转数") > 0 Then
                    Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "非结转期，不读取该校验")
                    GoTo nextWs
                Else
                    If dFres.exists(wsFre) Then
                        Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "读取内置校验")
                    Else
                        If InStr(ws.name, "自定义") > 0 Then
                            Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "读取自定义校验")
                        Else
                            Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "频度不一致，不读取该校验")
                        End If
                    End If
                End If
                
                'Call gloEt.ShowData(ws)
                arrSpcRef = ws.Range("A1").CurrentRegion.Value   'ws.UsedRange.Value
                For i = 2 To UBound(arrSpcRef) Step 1
                    If CStr(arrSpcRef(i, SforbidCol)) = "" Or VarType(arrSpcRef(i, SforbidCol)) = vbEmpty Then
                        arrT = Split(arrSpcRef(i, c1), "/")
                        For j = 0 To UBound(arrT) Step 1
                            code = replace(arrT(j), "'", "")
                            str = arrSpcRef(i, c2) & "," & _
                                CStr(val(arrSpcRef(i, c3))) & "," & _
                                CStr(val(arrSpcRef(i, c4))) & "," & _
                                CStr(arrSpcRef(i, c5)) & "," & _
                                replace(CStr(arrSpcRef(i, c6)), "'", "") & "," & _
                                CStr(arrSpcRef(i, c7)) & "," & CStr(arrSpcRef(i, c8)) & "," & CStr(arrSpcRef(i, 13))
                                
                                
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
                                    CStr(arrSpcRef(i, c7)) & "," & CStr(arrSpcRef(i, c8)) & "," & CStr(arrSpcRef(i, 13))
                                    
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
        End If
nextWs:

    Next
    
Exit Sub

errHandle:
    Set dSpcRef = Nothing
    Set dCurDataRef = Nothing
    Set dPreDataRef = Nothing
    Set dUnitRef = Nothing
    Erase arrAlert
    str = "特殊指标表初始化失败," & err.Description
    err.Clear
    err.Raise Number:=1002, Description:=str
End Sub

'
'   判断频度是否满足
'
Private Function InitComplexRef_CheckFres(rule As String) As Boolean
    Dim reg As Object, mc As Object, splitCount As Integer, m As Object, arr As Variant
    On Error Resume Next
    splitCount = COMPLEXT_SPLIT_COUNT  '这里修改指标分隔符
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        arr = Split(replace(replace(CStr(m.Value), "{", ""), "}", ""), ",")
        If Not dFres.exists(CStr(arr(COMPLEX_FRE_INDEX - 1)) & CStr(arr(COMPLEX_FRE_INDEX))) Then
            InitComplexRef_CheckFres = False
            Exit Function
        End If
    Next
    
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    Set mc = reg.Execute(rule)
    For Each m In mc
        arr = Split(replace(replace(CStr(m.Value), "[", ""), "]", ""), ",")
        If Not dFres.exists(CStr(arr(COMPLEX_FRE_INDEX - 1)) & CStr(arr(COMPLEX_FRE_INDEX))) Then
            InitComplexRef_CheckFres = False
            Exit Function
        End If
    Next
    InitComplexRef_CheckFres = True
End Function

'
'   初始化复杂指标表
'
Private Sub InitComplexRef()
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
    ComplexThd = ycV(ThisWorkbook.Worksheets("更新日志").Range("C1").Value)
    'ComplexThd = ComplexThd * (GetUnitVal("元") / GetUnitVal(DesUnit))
    
    On Error GoTo errHandle
    '校验表单，校验描述，校验规则，取反标识，禁用标识
    
    
    '遍历所有特殊指标表
    Dim mycnt As Control
    For Each mycnt In 导数比较窗体.CheckFrame.Controls
        If TypeName(mycnt) = "CheckBox" And mycnt.Value = True And InStr(mycnt.name, "复杂校验") > 0 Then
            '若选了，则
            Set ws = ThisWorkbook.Worksheets(mycnt.name)
            'Call gloEt.ShowData(ws)
            errRuleNum = 0
            ruleNum = 0
            targetRuleNum = 0
            jyNum = 0
            notfreNum = 0
            Set rn = Nothing
            Set rn2 = Nothing '用来标记不读取的多频
            ws.Cells.Interior.colorindex = 0
            arr = ws.Range("A1").CurrentRegion.Value   'ws.UsedRange.Value
            
            '判断是否结转数
            If Month(CurDate) = 1 And Day(CurDate) = 1 And InStr(ws.name, "结转数") > 0 Then
                
            ElseIf Month(CurDate) = 1 And Day(CurDate) = 1 And InStr(ws.name, "结转数") <= 0 Then
                '这里要跳出
                Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "本期为结转数，不读取此校验")
                GoTo nextRuleWs
            ElseIf Day(CurDate) <> 1 And InStr(ws.name, "结转数") > 0 Then
                Call gloDe.InsertNewDebug("系统导数比较", "-", ws.name, "初始化", "非结转数，不读取此校验")
                GoTo nextRuleWs
            Else
                
            End If
            
            'Debug.Assert (ws.name <> "复杂校验-月报多频")
 
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
                        '是否满足频度要求
                        If InitComplexRef_CheckFres(rule) = False Then
                            rightRule = False
                            notfreNum = notfreNum + 1
                            
                            If rn2 Is Nothing Then
                                Set rn2 = ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2)))
                            Else
                                Set rn2 = Application.Union(rn2, ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))))
                            End If
                            GoTo nextRule
                        End If
                        
                        'Debug.Assert (ws.name <> "复杂校验-单频")
                        '可以匹配里面的指标，判断指标描述指标采用splitcount个逗号
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
                                arrT = Split(CStr(m.Value), ",")
                                code = CStr(arrT(COMPLEX_TARGET_INDEX))    '此处读入指标代码
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
                                '
                                If Not dtemp.exists(rule) Then
                                    If CStr(arr(i, 6)) <> "" Then
                                        'thdvalue = ycV(CStr(arr(i, 6))) * (GetUnitVal("万元") / GetUnitVal(DesUnit))
                                        thdvalue = ycV(CStr(arr(i, 6))) * GetUnitVal("万元")    '转化为元
                                    Else
                                        thdvalue = ComplexThd
                                    End If
                                    dtemp(rule) = CStr(arr(i, 1)) & "|" & CStr(arr(i, 2)) & "|" & CStr(m.Value) & "|" & CStr(arr(i, 4)) & "|" & CStr(thdvalue)
                                    Set dComplexRef(code) = dtemp
                                    targetRuleNum = targetRuleNum + 1
                                Else
                                    '如果有重复，可能是不同币种、属性、频度，追加上去。
                                    If InStr(CStr(dtemp(rule)), CStr(m.Value)) > 0 Then
                                    
                                    Else
                                        '也就是，如果超过了0，1，2，3，4。5个数组，说明还有其他属性的该指标校验
                                        dtemp(rule) = CStr(dtemp(rule)) & "|" & CStr(m.Value)
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
        End If
nextRuleWs:
    Next
    
Exit Sub

errHandle:
    Call EraseSomthing
    str = "复杂指标表初始化失败," & err.Description
    err.Clear
    err.Raise Number:=1002, Description:=str
End Sub

Private Sub Init_not_jzd_target()
    On Error Resume Next
    Dim ws As Worksheet, arr As Variant, i As Long
    
    Set d_not_jzd = Nothing
    Set d_not_jzd = CreateObject("scripting.dictionary")
    
    Set ws = gloEt.FindWs(ThisWorkbook, "不参与集中度核查指标", 0, False)
    If Not ws Is Nothing Then
        arr = gloEt.GetUsedRange(ws).Value
        For i = 2 To UBound(arr)
            d_not_jzd(CStr(arr(i, 1))) = 1
        Next
    End If
End Sub

'
'   初始化复杂指标表
'
Private Sub InitdOtherTargetRef()
    Dim ws As Worksheet, i As Long, j%, k%, arr As Variant, arrT As Variant, code As String, str As String

    On Error GoTo errHandle
    
    Set dOtherTarget = Nothing
    Set dOtherTarget = CreateObject("scripting.dictionary")

    Set ws = ThisWorkbook.Worksheets("其他项标记")
    arr = ws.UsedRange.Value
    
    
    For i = 2 To UBound(arr) Step 1
        If CStr(arr(i, 4)) = "" Or VarType(arr(i, 4)) = vbEmpty Then
            dOtherTarget(replace(CStr(arr(i, 1)), "'", "")) = ycV(CStr(arr(i, 3)))
        End If
    Next i
    
    OtherR1 = 8
    OtherR2 = 9
    
    OtherR1 = CInt(gloCft.GetConfig("其他项标记起始列"))
    OtherR2 = CInt(gloCft.GetConfig("其他项标记末尾列"))
    OtherColorIndex = CInt(gloCft.GetConfig("其他项标记颜色"))
    OtherFontColorIndex = CInt(gloCft.GetConfig("其他项标记字体颜色"))
Exit Sub

errHandle:
    Call EraseSomthing
    str = "其他指标表初始化失败," & err.Description
    err.Clear
    err.Raise Number:=1002, Description:=str
End Sub

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
        If Len(m.Value) - Len(replace(m.Value, ",", "")) <> splitCount Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
        arr = Split(CStr(m.Value), ",")
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
        If Len(m.Value) - Len(replace(m.Value, ",", "")) <> splitCount Then
            InitComplexRef_CheckTaget = False
            Exit Function
        End If
        arr = Split(CStr(m.Value), ",")
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
Public Function InitComplexRef_ClearRule(rule As String) As String
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
            If InStr(clearRule, "[") <= 0 And InStr(clearRule, "]") <= 0 And InStr(clearRule, "{") <= 0 And InStr(clearRule, "}") <= 0 Then
                InitComplexRef_ClearRule = clearRule
            Else
                InitComplexRef_ClearRule = ""
            End If
        Else
            InitComplexRef_ClearRule = clearRule
        End If
    End If
End Function

'
'   添加备注 mark 返回备注内容
'
Private Function DoCompare_AddMark(ws As Worksheet, row As Long, h As Double) As String
    If Round(h, 6) = 0 Then
        DoCompare_AddMark = ""
        Exit Function
    End If
    Dim i%, c%, str$
    c = UBound(arrAlert)
    
    str = ""
    For i = c To LBound(arrAlert) Step -1
        If Round(h - arrAlert(i).变幅, 10) >= 0 Then
            str = arrAlert(i).备注文字
            If arrAlert(i).备注颜色 <> 0 Then
                With ws
                    If arrAlert(i).是否整行填充 = 0 Then
                        .Cells(row, 备注列).Interior.colorindex = arrAlert(i).备注颜色  '调整只显示那一列
                    Else
                        .Range(.Cells(row, 1), .Cells(row, 备注列)).Interior.colorindex = arrAlert(i).备注颜色
                    End If
                End With
            End If
            Exit For
        End If
    Next i
    If i < LBound(arrAlert) Then
        str = "缩小100倍以上"
        With ws
            .Range(.Cells(row, 1), .Cells(row, 备注列)).Interior.colorindex = 10
        End With
    End If
    
    DoCompare_AddMark = str
End Function

'
'   添加复杂校验描述
'
'
Private Function DoCompare_AddMyComplexExplain(strSt As String, index As Long) As String
    Dim str As String, drules As Object, i As Long, arrRules As Variant, arrT As Variant, tRule As String, oldRule As String, roundRule As String
    Dim ct As ComplexTaget, arrCt As Variant, reg As Object, mc As Object, m As Object, tStr As String, jsStr As String, g As Long, w_y As Worksheet
    Dim result As Variant, relStr
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
            If DoCompare_AddMyComplexExplain_IsNotChecked(oldRule, index) Then
                '替换上期
                With reg
                    .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
                End With
                Set mc = reg.Execute(tRule)
                For Each m In mc
                    'If DoCompare_is_Accu_Target(CStr(m.Value), index) = False Then
                    tStr = DoCompare_AddMyComplexExplain_ChangeTarge(CStr(m.Value), index, True)
                    If dPreDataRef.exists(tStr) Then
                        '存在则替换相应的值
                        If ycV(dPreDataRef(tStr)) = 0 Then
                            '存在，又为0的话，当做有一个很小的数。
                            '这里因为后期会用round添上，故改为用0.01元
                            'tRule = replace(tRule, CStr(m.value), ycV(dbl_Epsilon))
                            tRule = replace(tRule, CStr(m.Value), ycV(0.01))
                        Else
                            tRule = replace(tRule, CStr(m.Value), DoCompare_AddMyComplexExplain_ChangeData(CStr(m.Value), ycV(dPreDataRef(tStr))))
                        End If
                    Else
                        '否则用0替换
                        tRule = replace(tRule, CStr(m.Value), 0)
                    End If
                    '添加规则
                    Call DoCompare_AddMyComplexExplain_AddHadChecked(ct.ruleEx, oldRule, tStr)
                Next
                
                With reg
                    .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
                End With
                Set mc = reg.Execute(tRule)
                For Each m In mc
                    tStr = DoCompare_AddMyComplexExplain_ChangeTarge(CStr(m.Value), index)
                    If dCurDataRef.exists(tStr) Then
                        '存在则替换相应的值
                        If ycV(dCurDataRef(tStr)) = 0 Then
                            '存在，又为0的话，当做有一个很小的数。
                            '这里因为后期会用round添上，故改为用0.01元
                            'tRule = replace(tRule, CStr(m.value), ycV(dbl_Epsilon))
                            tRule = replace(tRule, CStr(m.Value), ycV(0.01))
                        Else
                            '存在，又为0的话，当做有一个很小的数。
                            tRule = replace(tRule, CStr(m.Value), DoCompare_AddMyComplexExplain_ChangeData(CStr(m.Value), ycV(dCurDataRef(tStr))))
                        End If
                        'tRule = replace(tRule, CStr(m.value), ycV(dCurDataRef(tStr)))
                    Else
                        '否则用0替换
                        tRule = replace(tRule, CStr(m.Value), 0)
                    End If
                    Call DoCompare_AddMyComplexExplain_AddHadChecked(ct.ruleEx, oldRule, tStr)
                Next
                
                tRule = replace(tRule, "Thd", CStr(ct.thdV))
                tRule = replace(tRule, "thd", CStr(ct.thdV))
                '这里有一个当前结果
                
                '执行结果
                'result = Evaluate(tRule)
                'Debug.Assert (strSt <> "12D26")
                roundRule = DoCompare_AddMyComplexExplain_RoundRule(tRule)
                If Len(roundRule) < 250 Then
                    result = Evaluate(roundRule)
                Else
                    Set w_y = gloEt.FindWs(ThisWorkbook, "yc辅助区", 0, True)
                    With w_y.Range("A1")
                        .ClearContents
                        .NumberFormat = "General"
                        .Formula = "=" & roundRule
                        .Calculate
                        result = .Value
                    End With
                End If
                jsStr = "计算过程：【" & oldRule & "】->【" & tRule & "】"
                
                '兜底回退
                If IsError(result) Or err.Number <> 0 Then
                    If Len(roundRule) < 250 Then
                        result = Evaluate(tRule)
                    Else
                        Set w_y = gloEt.FindWs(ThisWorkbook, "yc辅助区", 0, True)
                        With w_y.Range("A1")
                            .ClearContents
                            .NumberFormat = "General"
                            .Formula = "=" & tRule
                            .Calculate
                            result = .Value
                        End With
                    End If
                    jsStr = "回退计算过程：【" & oldRule & "】->【" & tRule & "】"
                    Call gloDe.InsertNewDebug("系统导数比较", CStr(arrCur(index, 4)), strSt, "校验回退", "执行[" & ct.ruleEx & "]出现回退" & jsStr)
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
                    
                    If VarType(arrCur(index, 硬性过程列)) = vbEmpty Or CStr(arrCur(index, 硬性过程列)) = "" Then
                        arrCur(index, 硬性过程列) = replace(replace(jsStr, "左中", "["), "右中", "]")
                    Else
                        arrCur(index, 硬性过程列) = CStr(arrCur(index, 硬性过程列)) & "    ||    " & replace(replace(jsStr, "左中", "["), "右中", "]")
                    End If

    
                End If
            End If  '判断是否执行过
        End If '若是本指标的校验才执行
    Next i
    '获取校验规则
    'DoCompare_AddMyComplexExplain = relStr
End Function

'
'   核实是否是当年累计发生额指标
'
Private Function DoCompare_is_Accu_Target(strSt As String) As Boolean
    On Error Resume Next
    Dim is_Accu_Target As Boolean
    is_Accu_Target = False
    
    If (Year(CurDate) = Year(PreDate)) And (Month(PreDate) <> 1 And Day(PreDate) <> 1) Then
        '当一年度就可以比较,且上期值不是1月1日，上期值是1月1日的，也返回找不到
    Else
        '不同年度，判断是否是发生额指标，如果是，则返回True
        If dtarget_accu.exists(strSt) Then
            is_Accu_Target = True
        Else
        
        End If
    End If
    
    If err.Number <> 0 Then
        Debug.Print "DoCompare_is_Accu_Target 执行当年累计指标判断有误。"
        err.Clear
        DoCompare_is_Accu_Target = False
    Else
        DoCompare_is_Accu_Target = is_Accu_Target
    End If
    
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
    rule = "and((2460676800-2476576200>0),(0-0<=0))"
    Debug.Print (rule)
    rule = DoCompare_AddMyComplexExplain_RoundRule(rule)
    Debug.Print rule
    
End Sub

'
'   将两边数值加个round(x,2)
'
Public Function DoCompare_AddMyComplexExplain_RoundRule(rule As String) As String
    On Error GoTo errHandle
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
                    GoTo errHandle
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

errHandle:
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
    Dim c1 As String, c10 As String, c11 As String, c12 As String, c13 As String, c3 As String, c5 As String
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    
    result = False
    arrCt = Split(cTrule, "|")
    
    For j = 0 To UBound(arrCt) Step 1
         t = replace(CStr(arrCt(j)), "[", "")
         t = replace(t, "]", "")
         arr = Split(t, ",")
         '1 - 业务类，3-机构类代码，5-地区代码，8-指标代码，10 - 数据属性，11币种，12频度，13批次
          'arr 0 -业务类，1机构类代码，2 地区代码，3指标代码，4数据属性，5币种，6频度，7批次
         c1 = CStr(arrCur(i, 1)): c10 = CStr(arrCur(i, 10)): c11 = CStr(arrCur(i, 11)): c12 = CStr(arrCur(i, 12)): c13 = CStr(arrCur(i, 13))
         c3 = replace(CStr(arrCur(i, 3)), "'", ""): c5 = replace(CStr(arrCur(i, 5)), "'", "")
         
        
         '不需要检测指标代码，是因为，就是因为指标代码是他才会进入到这里
         If CStr(arr(0)) <> "" And CStr(arr(0)) <> c1 Then GoTo NoThisTarget
         If CStr(arr(4)) <> "" And CStr(arr(4)) <> c10 Then GoTo NoThisTarget
         If CStr(arr(5)) <> "" And CStr(arr(5)) <> c11 Then GoTo NoThisTarget
         If CStr(arr(6)) <> "" And CStr(arr(6)) <> c12 Then GoTo NoThisTarget
         If CStr(arr(7)) <> "" And CStr(arr(7)) <> c13 Then GoTo NoThisTarget
         
         '构建正则匹配，若设置了机构代码，是否匹配
         If CStr(arr(1)) <> "" Then
             With reg
                 .Global = True: .ignorecase = False: .Pattern = replace(replace(CStr(arr(1)), "左中", "["), "右中", "]")
                 If .test(c3) = False Then GoTo NoThisTarget
             End With
         End If
         
         '构建正则匹配，若设置了地区代码，是否匹配
         If CStr(arr(2)) <> "" Then
             With reg
                 .Global = True: .ignorecase = False: .Pattern = replace(replace(CStr(arr(2)), "左中", "["), "右中", "]")
                 If .test(c5) = False Then GoTo NoThisTarget
             End With
         End If
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
        Dim c1 As String, c3 As String, c5 As String, c8 As String, c10 As String, c11 As String, c12 As String, c13 As String, str As String, d As Object
        c1 = CStr(arrCur(i, 1)): c3 = replace(CStr(arrCur(i, 3)), "'", ""): c5 = replace(CStr(arrCur(i, 5)), "'", ""): c8 = replace(CStr(arrCur(i, 8)), "'", "")
        c10 = CStr(arrCur(i, 10)): c11 = CStr(arrCur(i, 11)): c12 = CStr(arrCur(i, 12)): c13 = CStr(arrCur(i, 13))
        str = c1 & c3 & c5 & c8 & c10 & c11 & c12 & c13
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
    
    If Not dUnitRef.exists(t) Then
        v = val * unit
    Else
        v = val
    End If
    
    DoCompare_AddMyComplexExplain_ChangeData = v
End Function

'
'   返回对应的指标
'
Private Function DoCompare_AddMyComplexExplain_ChangeTarge(targeRule As String, i As Long, Optional is_accu As Boolean = False) As String
    '获取本指标的一些基本属性，以应对不存在时进行替换
    Dim c1 As String, c3 As String, c5 As String, c8 As String, c10 As String, c11 As String, c12 As String, c13 As String
    
    On Error Resume Next
    
    c1 = CStr(arrCur(i, 1)): c3 = replace(CStr(arrCur(i, 3)), "'", ""): c5 = replace(CStr(arrCur(i, 5)), "'", ""): c8 = replace(CStr(arrCur(i, 8)), "'", "")
    c10 = CStr(arrCur(i, 10)): c11 = CStr(arrCur(i, 11)): c12 = CStr(arrCur(i, 12)): c13 = CStr(arrCur(i, 13))
    
    Dim arr As Variant, t As String, str As String
    t = replace(targeRule, " ", "")
    t = replace(t, "[", ""): t = replace(t, "]", "")
    t = replace(t, "{", ""): t = replace(t, "}", "")
    arr = Split(t, ",")
    '1 - 业务类，3-机构类代码，5-地区代码，8-指标代码，10 - 数据属性，11币种，12频度，13批次
     'arr 0 -业务类，1机构类代码，2 地区代码，3指标代码，4数据属性，5币种，6频度，7批次
     
    If CStr(arr(0)) <> "" Then c1 = CStr(arr(0))
    If CStr(arr(1)) <> "" And InStr(CStr(arr(1)), "^") <= 0 Then c3 = CStr(arr(1))
    If CStr(arr(2)) <> "" And InStr(CStr(arr(2)), "^") <= 0 Then c5 = CStr(arr(2))
    If CStr(arr(3)) <> "" Then c8 = CStr(arr(3))
    If CStr(arr(4)) <> "" Then c10 = CStr(arr(4))
    If CStr(arr(5)) <> "" Then c11 = CStr(arr(5))
    If CStr(arr(6)) <> "" Then c12 = CStr(arr(6))
    If CStr(arr(7)) <> "" Then c13 = CStr(arr(7))
    
    'If CStr(arr(1)) <> "" Then c3 = CStr(arr(4))
    'If CStr(arr(2)) <> "" Then c5 = CStr(arr(5))   '暂不支持跨机构地区的校验
    
    'str = arrCur(i, 1) & arrCur(i, 3) & arrCur(i, 5) & arrCur(i, 8) & arrCur(i, 10) & arrCur(i, 11) & arrCur(i, 12) & arrCur(i, 13)
    str = c1 & c3 & c5 & c8 & c10 & c11 & c12 & c13
    
    '如果是当年累计发生额指标，则判断是否同年份
    If is_accu = True Then
        If DoCompare_is_Accu_Target(c8 & c10 & c12) = True Then
            'Debug.Print c8 & c10 & c12
            str = "DoCompare_is_Accu_Target Not Find"
        End If
    End If
    
    DoCompare_AddMyComplexExplain_ChangeTarge = str
End Function





'
'   添加是否说明标志
'   strSt - 指标
'   index - 数组对应的行
'   arrCur - 为所有数据
'
Private Function DoCompare_AddMyExplain(strSt As String, index As Long) As String
    Dim str As String, strRel As String, arrSt() As SpecialTaget, st As SpecialTaget, arr As Variant, u As Double, i As Integer
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, IsRuan As Boolean
    Dim dqStep As Integer
    dqStep = 8  '现在分7列，增加一个变幅取值
    
    curD = ycV(arrCur(index, 14))   '当期值
    preD = ycV(arrCur(index, 15))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, 11))  '币种
    dType = CStr(arrCur(index, 10)) '数据属性
    
    arr = Split(dSpcRef(strSt), ",")
    ReDim arrSt(Int(UBound(arr) / dqStep))
    
    For i = 0 To Int(UBound(arr) / dqStep) Step 1
        arrSt(i).text = arr(dqStep * i + 0)
        arrSt(i).rmbF = Abs(CDbl(arr(dqStep * i + 1)))
        arrSt(i).usdF = Abs(CDbl(arr(dqStep * i + 2)))
        arrSt(i).dataType = CStr(arr(dqStep * i + 3))
        arrSt(i).explain = CStr(arr(dqStep * i + 4))
        arrSt(i).TableName = CStr(arr(dqStep * i + 5))
        arrSt(i).detailed = CStr(arr(dqStep * i + 6))
        arrSt(i).bf = Abs(ycV(arr(dqStep * i + 7)))
        If arrSt(i).detailed <> "" Then arrSt(i).detailed = "-" & arrSt(i).detailed
    Next i
    
    
    u = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    
    If dUnitRef.exists(strSt) Then u = 1
    
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
                If CStr(arrCur(index, 14)) <> "" Then
                    str = st.text & st.detailed
                ElseIf preD <> 0 Then
                    'str = st.text & "(上期有数)"
                End If
            ElseIf st.text = "指标有变动需说明。" Then
                If Round(st.bf, 6) = 0 Then
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
                        If st.rmbF = 0 Then
                            If Abs(valH) <> 0 Then
                                str = "指标有变动，需说明。" & st.detailed
                            End If
                        Else
                            If Abs(valH) * u >= Abs(st.rmbF) Then
                                str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
                            End If
                        End If
                    End If
                Else
                    If dPro = "美元合计" Then
                        If st.usdF = 0 Then
                            If Abs(valH) <> 0 And Abs(valBf) > st.bf Then
                                str = "指标变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                            End If
                        Else
                            If Abs(valH) * u >= Abs(st.usdF) And Abs(valBf) > st.bf Then
                                str = "指标变动超过" & CStr(Abs(st.usdF)) & replace(特殊指标数据单位, "元", "美元") & "," & "且变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                            End If
                        End If
                    Else
                        If st.rmbF = 0 Then
                            If Abs(valH) <> 0 And Abs(valBf) > st.bf Then
                                str = "指标变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                            End If
                        Else
                            If Abs(valH) * u >= Abs(st.rmbF) And Abs(valBf) > st.bf Then
                                str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "," & "且变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                            End If
                        End If
                    End If
                End If
            ElseIf st.text = "指标变动或变幅超过阈值需说明。" Then
                If dPro = "美元合计" Then
                    If st.usdF = 0 Then
                        If Abs(valH) <> 0 Or Abs(valBf) > st.bf Then
                            str = "指标变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                        End If
                    Else
                        If Abs(valH) * u >= Abs(st.usdF) Or Abs(valBf) > st.bf Then
                            str = "指标变动超过" & CStr(Abs(st.usdF)) & replace(特殊指标数据单位, "元", "美元") & "," & "或变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                        End If
                    End If
                Else
                    If st.rmbF = 0 Then
                        If Abs(valH) <> 0 Or Abs(valBf) > st.bf Then
                            str = "指标变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                        End If
                    Else
                        If Abs(valH) * u >= Abs(st.rmbF) Or Abs(valBf) > st.bf Then
                            str = "指标变动超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "," & "或变幅超过" & Format(st.bf, "0.00%") & "，需说明。" & st.detailed
                        End If
                    End If
                End If
            '在前面加上所有的当年累计发生额指标的校验，这样的话，就不需要管理了
            ElseIf st.text = "当年累计指标比上期不应减少。" Then
                '首先判断是否为负数
                If curD < 0 Then
                    str = "发生额指标不应为负数，需说明。" & st.detailed
                End If
                If Month(CurDate) = 1 And Day(CurDate) = 1 Then
                    '结转数不进行判断
                Else
                    '判断时间，首先本期日期是大于上期的，下来判断年份是否一致就好
                    If Year(CurDate) = Year(PreDate) Then
                        '同年份的话，则必须满足
                        If valH < 0 Then
                            str = str & " " & "当年累计指标比上期不应减少，需说明。" & st.detailed
                        End If
                    Else
                        '若不同年份,若月份>=上期，如2018.8,2017.8，则小于0异常
                        If Month(CurDate) >= Month(PreDate) Then
                            If valH < 0 Then
                                str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据比" & CStr(PreDate) & "的小，请核实。" & st.detailed
                            End If
                        Else    '若月份<上期，如2018.1,2017.12，则>=0异常
                            If valH >= 0 Then
                                str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据 >= " & CStr(PreDate) & "的，请核实。" & st.detailed
                            End If
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
                
            ElseIf st.text = "指标为负数需核实。" Then
                str = 指标为负数需核实(st, index)
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
    
    ' 直接重复再判断一次
    If InStr(strRel, "发生额指标不应为负数，需说明。") <= 0 And InStr(strRel, "当年累计指标比上期不应减少，需说明。") <= 0 Then
        str = ""
        If (dtarget_accu.exists(replace(arrCur(index, 8), "'", "") & CStr(arrCur(index, 10)) & CStr(arrCur(index, 12))) = True) Then
            '首先判断是否为负数
            If curD < 0 Then
                str = "发生额指标不应为负数，需说明。"
            End If
            '判断时间，首先本期日期是大于上期的，下来判断年份是否一致就好
            If Month(CurDate) = 1 And Day(CurDate) = 1 Then
                '结转数不进行判断
            Else
                If Year(CurDate) = Year(PreDate) Then
                    '同年份的话，则必须满足
                    If valH < 0 Then
                        str = str & " " & "当年累计指标比上期不应减少，需说明。"
                    End If
                Else
                    '若不同年份,若月份>=上期，如2018.8,2017.8，则小于0异常
                    If Month(CurDate) >= Month(PreDate) Then
                        If valH < 0 Then
                            str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据比" & CStr(PreDate) & "的小，请核实。"
                        End If
                    Else    '若月份<上期，如2018.1,2017.12，则>=0异常
                        If valH >= 0 Then
                            str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据 >= " & CStr(PreDate) & "的，请核实。"
                        End If
                    End If
                End If
            End If
        End If
        
        If str <> "" Then
            If strRel = "" Then
                strRel = strRel & str
            Else
                strRel = strRel & "    ||    " & str
            End If
        End If
    End If
    
    If strRel <> "" And IsRuan = True Then
        wsCur.Cells(index, 是否说明列).Interior.colorindex = Complex_Soft_Color_Index
    End If
    DoCompare_AddMyExplain = strRel
End Function

Function 指标结清需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String, bz As String
    
    curD = ycV(arrCur(index, 14))   '当期值
    preD = ycV(arrCur(index, 15))   '上期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    valBf = ycV(arrCur(index, 环比列))  '环比
    dPro = CStr(arrCur(index, 11))  '币种
    dType = CStr(arrCur(index, 10)) '数据属性
    bz = CStr(arrCur(index, 18))  '备注
    
    str = ""
    
    If curD = 0 And preD > 0 Then
        If Abs(valH) * u >= Abs(st.rmbF) Then
            str = "指标结清超过" & CStr(Abs(st.rmbF)) & 特殊指标数据单位 & "，需说明。" & st.detailed
        End If
    End If
    
    指标结清需说明 = str
End Function

Function 当年累计指标比上期不应减少(index) As String
    Dim str As String
    Dim curD As Double, valH As Double
    
    curD = ycV(arrCur(index, 14))   '当期值
    valH = ycV(arrCur(index, 增减额列))  '比上期
    
    '首先判断是否为负数
    If curD < 0 Then
        str = "发生额指标不应为负数，需说明。"
    End If
    
    If Month(CurDate) = 1 And Day(CurDate) = 1 Then
                '结转数不进行判断
        Else
        '判断时间，首先本期日期是大于上期的，下来判断年份是否一致就好
        If Year(CurDate) = Year(PreDate) Then
            '同年份的话，则必须满足
            If valH < 0 Then
                str = str & " " & "当年累计指标比上期不应减少，需说明。"
            End If
        Else
            '若不同年份,若月份>=上期，如2018.8,2017.8，则小于0异常
            If Month(CurDate) >= Month(PreDate) Then
                If valH < 0 Then
                    str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据比" & CStr(PreDate) & "的小，请核实。"
                End If
            Else    '若月份<上期，如2018.1,2017.12，则>=0异常
                If valH >= 0 Then
                    str = str & " " & "当年累计指标比上期不应减少：" & CStr(CurDate) & "数据 >= " & CStr(PreDate) & "的，请核实。"
                End If
            End If
        End If
    End If
    
    当年累计指标比上期不应减少 = str
End Function

Function 指标应相等核查(st As SpecialTaget, index As Long, u As Double, strSt As String) As String
    Dim str As String, temp As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    Dim b As Double
    b = 0
    curD = val(arrCur(index, 14))
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & st.explain & _
        arrCur(index, 10) & arrCur(index, 11) & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    
    If Abs(curD - b) > dbl_Epsilon Then
        str = strSt & "不等于" & st.explain & "偏差" & Format(curD - b, "0.0000") & ",请核实。" & st.detailed
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
    curD = val(arrCur(index, 14))
    dbz = CStr(arrCur(index, 11))
    
    If dbz = "美元合计" Then
        If st.usdF <> 0 Then
            dbz = "人民币"
        End If
    End If
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & st.explain & _
        arrCur(index, 10) & dbz & arrCur(index, 12) & arrCur(index, 13)
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
    dPro = CStr(arrCur(index, 11))  '币种
    curD = val(arrCur(index, 14))
    
    '999 代表那些美元合计/人民币的。
    If st.usdF = 999 Then
        If dPro = "美元合计" Then
            dPro = "人民币"
        End If
    End If
    
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & st.explain & _
        arrCur(index, 10) & dPro & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If dUnitRef.exists(st.explain) Then u2 = 1
    
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
    dPro = CStr(arrCur(index, 11))  '币种
    curD = val(arrCur(index, 14))
    
    '999 代表那些美元合计/人民币的。
    If st.usdF = 999 Then
        If dPro = "美元合计" Then
            dPro = "人民币"
        End If
    End If
    
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & st.explain & _
        arrCur(index, 10) & dPro & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If dUnitRef.exists(st.explain) Then u2 = 1
    
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
    dPro = CStr(arrCur(index, 11))  '币种
    curD = val(arrCur(index, 14))
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & st.explain & _
        arrCur(index, 10) & arrCur(index, 11) & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    If dCurDataRef.exists(temp) Then b = val(dCurDataRef(temp))
    u2 = GetUnitVal(DesUnit) / GetUnitVal(特殊指标数据单位)
    If dUnitRef.exists(st.explain) Then u2 = 1
    
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

Function 指标比上期减少需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    
    valH = val(arrCur(index, 增减额列))  '比上期
    dPro = CStr(arrCur(index, 11))  '币种
    str = ""
    If dPro = "美元合计" Then
        If valH * u < 0 And Abs(valH * u) > st.usdF Then
            str = "指标比上期减少超过" & CStr(st.usdF) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
        End If
    Else
        If valH * u < 0 And Abs(valH * u) > st.rmbF Then
            str = "指标比上期减少超过" & CStr(st.rmbF) & 特殊指标数据单位 & "，需说明。" & st.detailed
        End If
    End If
    指标比上期减少需说明 = str
End Function

'
' 指标为负数需核实
'
Function 指标为负数需核实(st As SpecialTaget, index As Long) As String
    Dim str As String
    Dim curD As Double
    str = ""
    
    curD = ycV(arrCur(index, 14))   '当期值
    
    If curD < 0 Then
        str = "指标为负数需核实。"
    End If
    
    指标为负数需核实 = str
End Function

'
'   指标比上期增加需说明。
'
Function 指标比上期增加需说明(st As SpecialTaget, index As Long, u As Double) As String
    Dim str As String
    Dim curD As Double, preD As Double, valH As Double, valBf As Double, dPro As String, dType As String
    
    valH = val(arrCur(index, 增减额列))  '比上期
    dPro = CStr(arrCur(index, 11))  '币种
    str = ""
    If dPro = "美元合计" Then
        If valH * u > 0 And Abs(valH * u) > st.usdF Then
            str = "指标比上期增加超过" & CStr(st.usdF) & replace(特殊指标数据单位, "元", "美元") & "，需说明。" & st.detailed
        End If
    Else
        If valH * u > 0 And Abs(valH * u) > st.rmbF Then
            str = "指标比上期增加超过" & CStr(st.rmbF) & 特殊指标数据单位 & "，需说明。" & st.detailed
        End If
    End If
    指标比上期增加需说明 = str
End Function


'
'   指标应能被某数整除。
'
Function 校验指标能被某数整除(st As SpecialTaget, index As Long) As String
    Dim str As String, u As Double, r As Double
    Dim curD As Double
    str = ""
    If dUnitRef.exists(replace(arrCur(index, 8), "'", "")) Then
        u = 1
    Else
        u = GetUnitVal(DesUnit)
    End If
    
    On Error GoTo errH
    curD = Abs(val(arrCur(index, 14)))   '当期值
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


'
'   指标本期余额应小于（上期余额+本期当年累计发生额-上期当年累计发生额）
'
Function 校验余额及当年累计发生额(st As SpecialTaget, index As Long) As String
    Dim str As String, temp As String, dp2 As String, step As Integer
    Dim curD As Double, preD As Double, curFsD As Double, preFsD As Double
    str = ""
    curFsD = 0: curD = 0: preD = 0: preFsD = 0
    
    If CStr(arrCur(index, 10)) = "余额" Then
        dp2 = "发生额"
    Else
        dp2 = "余额"
    End If
    
    step = GetFreStep(CStr(arrCur(index, 12)))
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & arrCur(index, 8) & _
            dp2 & arrCur(index, 11) & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    '加上日期判断,1-同年份，2-上年末，3不连续
    If Year(CurDate) = Year(PreDate) Or ((Year(CurDate) - Year(PreDate) = 1) And (Month(PreDate) = 12 And Day(PreDate) = 31)) Then
        If dp2 = "发生额" Then
    
            If dCurDataRef.exists(temp) Then curFsD = val(dCurDataRef(temp))
            If dPreDataRef.exists(temp) Then preFsD = val(dPreDataRef(temp))
        
            curD = val(arrCur(index, 14))   '当期值
            preD = val(arrCur(index, 15))   '上期值
        Else    '余额
            '当期余额
            If dCurDataRef.exists(temp) Then curD = val(dCurDataRef(temp))
            '上期余额
            If dPreDataRef.exists(temp) Then preD = val(dPreDataRef(temp))
            curFsD = val(arrCur(index, 14)) '发生额
            preFsD = val(arrCur(index, 15))  '发生额-上期
        End If
        
        If (Month(PreDate) = 12 And Day(PreDate) = 31) Or (Month(PreDate) = 1 And Day(PreDate) = 1) Then preFsD = 0    '如果上月是年末即，跨年度了。不考虑上期发生额。有可能会把数据日期改成1月1日
        
        Dim fz As Double, lfz As Double, maxP As Double
        
        maxP = curD
        If maxP < preD Then maxP = preD
        If maxP < curFsD Then maxP = curFsD
        If maxP < preFsD Then maxP = preFsD
        
        '增加一个校准过程
        If Not dUnitRef.exists(replace(CStr(arrCur(index, 8)), "'", "")) Then
            preD = preD * GetUnitVal(DesUnit)
            curFsD = curFsD * GetUnitVal(DesUnit)
            curD = curD * GetUnitVal(DesUnit)
            preFsD = preFsD * GetUnitVal(DesUnit)
            maxP = maxP * GetUnitVal(DesUnit)
        End If
            
        
        If Round(preD + curFsD - preFsD - curD, 2) >= 0 Then
            str = ""
        Else
            If Round(curFsD - preFsD, 10) = 0 Then
                str = "本期无发生额，余额增加了，请核实"
            Else
                str = 发生额与余额误差判断(Round(preD + curFsD - preFsD - curD, 2), maxP)
            End If
        End If
    Else
        str = "不进行余额与当年累计发生额校验，因" & Format(CurDate, "yyyy-m-d") & "与" & Format(PreDate, "yyyy-m-d") & "年份不连续。"
    End If
    
    
    校验余额及当年累计发生额 = str
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
    Dim str As String, temp As String, dp2 As String, step As Integer
    Dim curD As Double, preD As Double, curFsD As Double
    str = ""
    curFsD = 0: curD = 0: preD = 0
    
    If CStr(arrCur(index, 10)) = "余额" Then
        dp2 = "发生额"
    Else
        dp2 = "余额"
    End If
    
    step = GetFreStep(CStr(arrCur(index, 12)))
    temp = arrCur(index, 1) & arrCur(index, 3) & arrCur(index, 5) & arrCur(index, 8) & _
            dp2 & arrCur(index, 11) & arrCur(index, 12) & arrCur(index, 13)
    temp = replace(temp, "'", "")
    
    '加上日期判断,1-连续，2-不连续，但是1月与年初
    If Format(DateSerial(Year(CurDate), Month(CurDate) - step + 1, 0), "yyyy-m-d") = Format(PreDate, "yyyy-m-d") Or _
      (Format(DateSerial(Year(PreDate), 1, 1), "yyyy-m-d") = Format(PreDate, "yyyy-m-d") And Format(DateSerial(Year(CurDate), Month(CurDate) - step + 1, 1), "yyyy-m-d") = Format(PreDate, "yyyy-m-d")) Then
        If dp2 = "发生额" Then
    
            If dCurDataRef.exists(temp) Then curFsD = val(dCurDataRef(temp))
        
            curD = val(arrCur(index, 14))   '当期值
            preD = val(arrCur(index, 15))   '上期值
        Else    '余额
            '当期余额
            If dCurDataRef.exists(temp) Then curD = val(dCurDataRef(temp))
            '上期余额
            If dPreDataRef.exists(temp) Then preD = val(dPreDataRef(temp))
            curFsD = val(arrCur(index, 14)) '发生额
        End If
        
        '计算阀值
        Dim maxP As Double
        
        maxP = curD
        If maxP < preD Then maxP = preD
        If maxP < curFsD Then maxP = curFsD
        
        '增加一个校准过程
        If Not dUnitRef.exists(replace(CStr(arrCur(index, 8)), "'", "")) Then
            preD = preD * GetUnitVal(DesUnit)
            curFsD = curFsD * GetUnitVal(DesUnit)
            curD = curD * GetUnitVal(DesUnit)
            maxP = maxP * GetUnitVal(DesUnit)
        End If
            
        
        If Round(preD + curFsD - curD, 2) >= 0 Then
            str = ""
        Else
            str = 发生额与余额误差判断(Round(preD + curFsD - curD, 2), maxP)
        End If
    Else
        str = "不进行余额与本期发生额校验，因" & Format(CurDate, "yyyy-m-d") & "与" & Format(PreDate, "yyyy-m-d") & "不连续。"
    End If
    
    If curFsD < 0 Then
        str = "发生额指标不应为负数。" & str
    End If
    
    校验本期上期余额及本期发生额 = str
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
        str = "精度误差：上期余额+本期发生额<本期余额，偏差150元内。"
    Else
        If max < l Then
            If Round(t, 2) < max Then
                str = "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过1万元。"
            Else
                str = "上期余额+本期发生额>本期余额，超过该业务量万分之一。"
            End If
        Else
            If Round(t, 2) < l Then
                str = "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过" & CStr(l) & "元。"
            Else
                str = "上期余额+本期发生额>本期余额，超过" & CStr(l) & "元。"
            End If
        End If
    End If
    发生额与余额误差判断 = str
End Function

'结束
Private Sub EndCompare()
    Dim fm$, iRowCur&, i%, ws As Worksheet
    
    Call gloEt.ShowData(wsCur)
    
    With wsCur
        iRowCur = .Range("A1").End(xlDown).row
    
        '格式设置
        .Range(.Cells(2, 2), .Cells(iRowCur, 2)).NumberFormatLocal = "yyyy/m/d"
        .Range(.Cells(2, 3), .Cells(iRowCur, 8)).NumberFormatLocal = "@"
        .Range(.Cells(2, 14), .Cells(iRowCur, 增减额列)).NumberFormatLocal = "0.00"
        .Range(.Cells(2, 环比列), .Cells(iRowCur, 环比列)).NumberFormatLocal = "0.00%"
        
        Call SetColumnWidth(wsCur)
        .Range("A2:Z" & CStr(iRowCur)).Sort Key1:=.Range("C2"), Key2:=.Range("E2"), Key3:=.Range("G2"), Header:=xlNo
        
        .name = CStr(Year(.Cells(2, 2).Value2)) + "年" + CStr(Month(.Cells(2, 2).Value2)) + "月对比结果"
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

'给新的工作簿增加宏代码
Private Sub AddMacroCode()
    Dim MyCodes As String, fullName As String
    
    On Error GoTo errHandle
    
    '插入启动菜单
    MyCodes = gloCft.GetConfig("输出说明菜单宏")
    With wbDes.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines 2, MyCodes
    End With
    '插入执行命令
    MyCodes = "public Const unit as Double = " + CStr(GetUnitVal(DesUnit) / GetUnitVal(分析文件数据单位)) + Chr(10) + Chr(13) _
            + "public Const 分析文件数据单位 as String = " + Chr(34) + CStr(分析文件数据单位) + Chr(34) + Chr(10) + Chr(13) _
            + "public Const 说明文件名 as String = " + Chr(34) + CStr(Year(CurDate)) + "." + CStr(Month(CurDate)) + " 说明文件.xls" + Chr(34) + Chr(10) + Chr(13) _
            + "public const 当期数据日期 as Date = #" & CStr(CurDate) & "#" + Chr(10) + Chr(13) _
            + "public const 上期数据日期 as Date = #" & CStr(PreDate) & "#" + Chr(10) + Chr(13)
    MyCodes = MyCodes + gloCft.GetConfig("输出说明通用宏")
    With wbDes.VBProject.VBComponents.add(1)
        .name = "common"
        .CodeModule.InsertLines 2, MyCodes
    End With
    
    MyCodes = gloCft.GetConfig("输出说明单元格处理宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "DealCell"
        .CodeModule.InsertLines 2, MyCodes
    End With
    
    MyCodes = gloCft.GetConfig("输出说明工作宏")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work"
        .CodeModule.InsertLines 2, MyCodes
    End With
    MyCodes = gloCft.GetConfig("输出说明工作宏_日报")
    With wbDes.VBProject.VBComponents.add(1)  'vbext_ct_StdModule
        .name = "work_day"
        .CodeModule.InsertLines 2, MyCodes
    End With
Exit Sub

errHandle:
    err.Clear
    MsgBox "说明文件功能受限。需要在EXCEL菜单 工具-宏－安全性－可靠发行商　中勾选【信任对于VB项目的访问】"
End Sub


'
'   检查源数据日期的一致性
'
Private Sub CheckDateUnieque(ws As Worksheet)
    Dim row As Long
    With ws
        row = gloEt.GetCurRows(ws)
        If Application.WorksheetFunction.CountIf(.Range(.Cells(2, 2), .Cells(row, 2)), .Cells(2, 2).Value2) <> row - 1 Then
            err.Raise Number:=1001, Description:="CheckDateUnieque:" & ws.name & "数据日期不一致。"
        End If
    End With
End Sub

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
        If .Range("A1").Value <> "默认警戒区间" Then
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

Private Sub SetColumnWidth(ws As Worksheet)
    With ws
        .Columns(1).ColumnWidth = gloCft.GetConfig("业务类列宽")
        .Columns(2).ColumnWidth = gloCft.GetConfig("数据日期列宽")
        .Columns(3).ColumnWidth = gloCft.GetConfig("机构类代码列宽")
        .Columns(4).ColumnWidth = gloCft.GetConfig("机构类名称列宽")
        .Columns(5).ColumnWidth = gloCft.GetConfig("地区代码列宽")
        .Columns(6).ColumnWidth = gloCft.GetConfig("地区名称列宽")
        .Columns(7).ColumnWidth = gloCft.GetConfig("指标顺序码列宽")
        .Columns(8).ColumnWidth = gloCft.GetConfig("指标代码列宽")
        .Columns(9).ColumnWidth = gloCft.GetConfig("指标名称列宽")
        .Columns(10).ColumnWidth = gloCft.GetConfig("数据属性列宽")
        .Columns(11).ColumnWidth = gloCft.GetConfig("币种列宽")
        .Columns(12).ColumnWidth = gloCft.GetConfig("频度列宽")
        .Columns(13).ColumnWidth = gloCft.GetConfig("批次列宽")
        .Columns(14).ColumnWidth = gloCft.GetConfig("数据值列宽")
        .Columns(15).ColumnWidth = gloCft.GetConfig("上期值列宽")
        .Columns(16).ColumnWidth = gloCft.GetConfig("比上期列宽")
        .Columns(17).ColumnWidth = gloCft.GetConfig("环比列宽")
        .Columns(18).ColumnWidth = gloCft.GetConfig("备注列宽")
        .Columns(19).ColumnWidth = gloCft.GetConfig("说明列宽") + 3
        .Columns(20).ColumnWidth = gloCft.GetConfig("说明列宽") + 3
        .Columns(21).ColumnWidth = gloCft.GetConfig("说明列宽")
        '.columns(22).ColumnWidth = gloCft.GetConfig("说明列宽") + 8
        '.columns(23).ColumnWidth = gloCft.GetConfig("说明列宽")
    End With
End Sub


