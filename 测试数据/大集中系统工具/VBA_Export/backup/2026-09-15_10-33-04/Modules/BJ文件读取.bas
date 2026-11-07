Attribute VB_Name = "BJ文件读取"
Option Explicit

Private Const STRSPLIT As String = "|"
Private dTargetRef As Object   '指标字典
Private dOrgRef As Object  '机构类代码
Private dAreaRef As Object '地区代码
Private dFreRef As Object  '频度字典
Private dBusinessRef As Object '业务类字典
Private dDataProRef As Object   '数据属性
Private dCurrencyRef As Object  '币种
Private dUnitRef As Object





'
'   初始化基本的映射
'
Public Function InitChangeBjToCsv()
    Dim ws As Worksheet, arr As Variant, i As Long
    
    '读取基本映射
    Set dFreRef = InitChangeBjToCsv_GetDic("IJ文件转换-频度代码", "IJ文件转换-频度名称")
    Set dBusinessRef = InitChangeBjToCsv_GetDic("IJ文件转换-业务类代码", "IJ文件转换-业务类名称")
    Set dDataProRef = InitChangeBjToCsv_GetDic("IJ文件转换-数据属性代码", "IJ文件转换-数据属性名称")
    Set dCurrencyRef = InitChangeBjToCsv_GetDic("IJ文件转换-币种代码", "IJ文件转换-币种名称")
    
    Set ws = ThisWorkbook.Worksheets("IJ文件-指标参照")
    arr = ws.Range("A1").CurrentRegion.Value
    Set dTargetRef = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr) Step 1
        If Not dTargetRef.exists(replace(CStr(arr(i, 1)), "'", "")) Then
            dTargetRef(replace(CStr(arr(i, 1)), "'", "")) = CStr(arr(i, 2))
        End If
    Next i
    
    Set ws = ThisWorkbook.Worksheets("IJ文件-机构类代码参照")
    arr = ws.Range("A1").CurrentRegion.Value
    Set dOrgRef = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr) Step 1
        If Not dOrgRef.exists(replace(CStr(arr(i, 1)), "'", "")) Then
            dOrgRef(replace(CStr(arr(i, 1)), "'", "")) = CStr(arr(i, 2))
        End If
    Next i
    
    Set ws = ThisWorkbook.Worksheets("IJ文件-地区代码参照")
    arr = ws.Range("A1").CurrentRegion.Value
    Set dAreaRef = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr) Step 1
        If Not dAreaRef.exists(replace(CStr(arr(i, 1)), "'", "")) Then
            dAreaRef(replace(CStr(arr(i, 1)), "'", "")) = CStr(arr(i, 2))
        End If
    Next i
    
    Set ws = ThisWorkbook.Worksheets("单位不需要转换指标参照表")
    arr = ws.Range("A1").CurrentRegion.Value
    Set dUnitRef = CreateObject("scripting.dictionary")
    
    For i = 2 To UBound(arr) Step 1
        dUnitRef(replace(CStr(arr(i, 2)), "'", "")) = 0
    Next i
End Function

Private Function InitChangeBjToCsv_GetDic(s1 As String, s2 As String) As Object
    Dim arr1 As Variant, arr2 As Variant, d As Object, i As Long
    
    '读取频度
    arr1 = Split(gloCft.GetConfig(s1), ",")
    arr2 = Split(gloCft.GetConfig(s2), ",")
    Set d = CreateObject("scripting.dictionary")
    For i = 0 To UBound(arr1) Step 1
        If Not d.exists(CStr(arr1(i))) Then
            d(CStr(arr1(i))) = CStr(arr2(i))
        End If
    Next i
    
    Set InitChangeBjToCsv_GetDic = d
End Function

'
'   清楚
'
Public Function EndChangeBjToCsv()
    Set dTargetRef = Nothing
    Set dOrgRef = Nothing
    Set dAreaRef = Nothing
    Set dFreRef = Nothing
    Set dBusinessRef = Nothing
    Set dDataProRef = Nothing
    Set dCurrencyRef = Nothing
    Set dUnitRef = Nothing
End Function

Sub testChange()
    Call InitChangeBjToCsv
    Call ChangeBjToCsv("C:\Documents and Settings\admin\桌面\2019.07\03 集中系统源数据\中小法人2批\127752000_20190731_YB2\BI7020441323120190731421.IDX", "当期数")
    Call EndChangeBjToCsv
End Sub

Public Function ChangeBjToCsv(Paths As String, mode As String) As String
    On Error GoTo errorHandle
    Dim csvPath As String, dLoadPath As Object, arrPaths As Variant, i As Long, strTmp As String, arrTmp As Variant, j As Long, arrTmp2 As Variant
    Dim fso As Object, f As Object, sfn As String, strfnIdx As String, unit As Double
    Dim c2 As Date, c3 As String, c4 As String, c5 As String, c6 As String, c12 As String, c13 As Integer
    Dim c1 As String, c10 As String, c11 As String
    Dim c8 As String, c9 As String
    Dim wbDes As Workbook, wsDes As Worksheet, nowCur As Long
    Dim n As Long
    
    Dim dLoadIDX As Object, arrData As Variant, dataNum As Long
    Dim arr2 As Variant, k As Long, hf As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dLoadPath = Nothing
    Set dLoadPath = CreateObject("scripting.dictionary")
    sfn = ""
    arrPaths = Split(Paths, ",")
    
    
    
    For i = 0 To UBound(arrPaths) Step 1
        '确定保存文件名
        If sfn = "" Then
            '创建csv文件
            sfn = left(CStr(arrPaths(i)), InStrRev(CStr(arrPaths(i)), "\")) & mode & "Temp" & Format(Now(), "yyyymmddhhmmss") & ".csv"
            Set f = fso.CreateTextFile(sfn, True)
            f.Close
            Set wbDes = gloEt.GetWb(sfn)
            Set wsDes = wbDes.Worksheets(1)
            With wsDes
                .Cells(1, 1).Value = "业务类": .Cells(1, 2).Value = "数据日期": .Cells(1, 3).Value = "机构类代码": .Cells(1, 4).Value = "机构类名称"
                .Cells(1, 5).Value = "地区代码": .Cells(1, 6).Value = "地区名称": .Cells(1, 7).Value = "指标顺序码": .Cells(1, 8).Value = "指标代码"
                .Cells(1, 9).Value = "指标名称": .Cells(1, 10).Value = "数据属性": .Cells(1, 11).Value = "币种": .Cells(1, 12).Value = "频度"
                .Cells(1, 13).Value = "批次": .Cells(1, 14).Value = "数据值"
            End With
        End If
        
        If dLoadPath.exists(CStr(arrPaths(i))) Then
            Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), "-", "转化IJ文件-重复读取", "已读取过，跳过")
            GoTo NextFileIDX
        Else
            dLoadPath(CStr(arrPaths(i))) = 1
        End If
        
        If Not gloFt.IsFileExist(GetDatFileName(CStr(arrPaths(i)))) Then
            '不存在，下一个
            Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), "-", "转化IJ文件失败", "未找到此IJ文件对应DAT文件，请确定在同一目录下")
            GoTo NextFileIDX
        End If
        
        '读取DAT文件，获取总数据量
        dataNum = GetDataNum(GetDatFileName(CStr(arrPaths(i))))
        If dataNum = 0 Then
            Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), "-", "转化IJ文件-空数据", "对应DAT文件数据未空，不读取")
            GoTo NextFileIDX
        End If
        ReDim arrData(1 To dataNum, 1 To 14)
        n = 0
        strfnIdx = right(CStr(arrPaths(i)), Len(CStr(arrPaths(i))) - InStrRev(CStr(arrPaths(i)), "\"))
        '解析数据日期
        c2 = GetDateFromIDX(strfnIdx)
        
        '解读频度
        If dFreRef.exists(Mid(strfnIdx, 22, 1)) Then
            c12 = CStr(dFreRef(Mid(strfnIdx, 22, 1)))
        Else
            c12 = Mid(strfnIdx, 22, 1)
            Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), Mid(strfnIdx, 22, 1), "转化IJ文件-频度映射为空", "基本配置中频度无此频度代码")
        End If
        
        '解读批次
        c13 = CInt(Mid(strfnIdx, 23, 1))

        '读取IDX文件,获取 关键字段 - 业务类，数据属性，币种，单位。多余属性表单名
'        If Left(strfnIdx, 2) = "AI" Then
'            Set dLoadIDX = LoadIDX_AI(CStr(arrPaths(i)))
'        Else
'            Set dLoadIDX = LoadIDX_BI(CStr(arrPaths(i)))
'        End If
        Set dLoadIDX = LoadIDX(CStr(arrPaths(i)))
        
        '读取DAT文件
        Set f = fso.OpenTextFile(GetDatFileName(CStr(arrPaths(i))))
        'Do Until f.atendofstream
            'strTmp = f.readline()
        strTmp = f.ReadAll
        hf = GetHf(strTmp)
        arr2 = Split(strTmp, hf)
        
        For k = 0 To UBound(arr2) Step 1
            strTmp = CStr(arr2(k))
            arrTmp = Split(strTmp, STRSPLIT)
            '进入
            For j = 0 To UBound(arrTmp) Step 3
                If dLoadIDX.exists(CStr(arrTmp(j))) Then
                    arrTmp2 = Split(CStr(dLoadIDX(CStr(arrTmp(j)))), ",")
                    
                    '不为0才进行读入
                    If ycV(arrTmp(j + 2)) <> 0 Then
                        '解析
                        n = n + 1
                        arrData(n, 1) = arrTmp2(1): arrData(n, 2) = c2: arrData(n, 3) = arrTmp2(3): arrData(n, 4) = arrTmp2(4)
                        arrData(n, 5) = arrTmp2(5): arrData(n, 6) = arrTmp2(6): arrData(n, 7) = "": arrData(n, 8) = "'" & CStr(arrTmp(j + 1))
                        
                        If dTargetRef.exists(CStr(arrTmp(j + 1))) Then
                            arrData(n, 9) = dTargetRef(CStr(arrTmp(j + 1)))
                        Else
                            arrData(n, 9) = ""
                            Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), CStr(arrTmp(j + 1)), "转化IJ文件-未定义指标代码", "未定义该指标代码")
                        End If
                        arrData(n, 10) = arrTmp2(10): arrData(n, 11) = arrTmp2(11): arrData(n, 12) = c12: arrData(n, 13) = c13
                    
                        If dUnitRef.exists(CStr(arrTmp(j + 1))) Then
                             arrData(n, 14) = ycV(arrTmp(j + 2))
                        Else
                            arrData(n, 14) = ycV(arrTmp(j + 2)) * (10 ^ (ycV(arrTmp2(12)) - 1))
                        End If
                    End If
                Else
                     Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), CStr(arrTmp(j)), "转化IJ文件-DAT关键字对应无效", "IDX找不到相应关键字")
                End If
            Next j
        Next k
        'Loop
        
        '输出到目标工作表
        nowCur = gloEt.GetCurRows(wsDes)
        If n > 0 Then
            wsDes.Cells(nowCur + 1, 1).Resize(n, 14).Value = arrData
        End If
        Erase arrData
        Set dLoadIDX = Nothing
        f.Close
        Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), "-", "转化IJ文件-转换结束", "")
        Call gloDe.AddColorDebug(15)
NextFileIDX:
    Next i
    
    If wbDes Is Nothing Then
        ChangeBjToCsv = ""
    Else
        wbDes.Save
        ChangeBjToCsv = wbDes.fullName
    End If
    
    Set fso = Nothing
    Set dLoadPath = Nothing
Exit Function

errorHandle:
    Call gloDe.InsertNewDebug("系统导数比较", CStr(arrPaths(i)), "-", "转化IJ文件-中断", err.Description)
    ChangeBjToCsv = ""
End Function

'
'   要在这里找出换行符
'
Private Function GetHf(strTmp As String) As String
    On Error Resume Next
    Dim hf As String, ia As Integer, i As Long, SubStr As String
    hf = ""
    If InStr(strTmp, Chr(13) & Chr(10)) > 0 Then
        hf = Chr(13) & Chr(10)
    ElseIf InStr(strTmp, Chr(10) & Chr(13)) > 0 Then
        hf = Chr(10) & Chr(13)
    ElseIf InStr(strTmp, Chr(13)) > 0 Then
        hf = Chr(13)
    ElseIf InStr(strTmp, Chr(10)) > 0 Then
        hf = Chr(10)
    Else
        For i = 1 To Len(strTmp)
            SubStr = Mid(strTmp, i, 1)
            ia = Asc(SubStr)
            If SubStr = "|" Or SubStr = "." Or (ia >= 48 And ia <= 57) Or (ia >= 65 And ia <= 90) Or (ia >= 97 And ia <= 122) Then
                If hf <> "" Then
                    Exit For
                End If
            Else
                hf = hf & SubStr
            End If
        Next i
    End If
    
    If hf = "" Then hf = Chr(10)
    
    'Debug.Print "换行符是:" & hf
    GetHf = hf
End Function

'
'   BI
'
Private Function LoadIDX(fn As String) As Object
    Dim i As Long, strTmp As String, arrTmp As Variant
    Dim fso As Object, f As Object, sfn As String
    Dim c2 As Date, c3 As String, c4 As String, c5 As String, c6 As String, c12 As String, c13 As Integer
    Dim c1 As String, c10 As String, c11 As String
    Dim dLoadIDX As Object, arr2 As Variant, k As Long, hf As String
    
    Dim strfnIdx As String, idxStep As Integer
    strfnIdx = right(fn, Len(fn) - InStrRev(fn, "\"))
    If left(strfnIdx, 2) = "AI" Then
        idxStep = 9
    Else
        idxStep = 10
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '读取IDX文件,获取 关键字段 - 业务类，数据属性，币种，单位。多余属性表单名
    Set dLoadIDX = Nothing
    Set dLoadIDX = CreateObject("scripting.dictionary")
    Set f = fso.OpenTextFile(fn, 1)

    strTmp = f.ReadAll
    hf = GetHf(strTmp)
    arr2 = Split(strTmp, hf)
    
    For k = 0 To UBound(arr2) Step 1
        strTmp = CStr(arr2(k))
        
        arrTmp = Split(strTmp, STRSPLIT)
        '进入
        For i = 0 To UBound(arrTmp) Step idxStep
            If Not dLoadIDX.exists(CStr(arrTmp(i))) Then
                '解析机构类
                c3 = "'" & CStr(arrTmp(i + 2))
                If dOrgRef.exists(replace(c3, "'", "")) Then
                    c4 = CStr(dOrgRef(replace(c3, "'", "")))
                Else
                    c4 = c3
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c3, "转化IJ文件-机构类映射为空", "IJ文件-机构类代码参照缺少该映射")
                End If
                
                '解析地区代码
                c5 = "'" & CStr(arrTmp(i + 3))
                If dAreaRef.exists(replace(c5, "'", "")) Then
                    c6 = CStr(dAreaRef(replace(c5, "'", "")))
                Else
                    c6 = c5
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c5, "转化IJ文件-地区映射为空", "IJ文件-地区代码参照缺少该映射")
                End If
                
                '业务类,通过表单名称来获取的
                c1 = GetBusinessName(CStr(arrTmp(i + 1)))
                
                '数据属性
                If dDataProRef.exists(CStr(arrTmp(i + 4))) Then
                    c10 = dDataProRef(CStr(arrTmp(i + 4)))
                Else
                    c10 = CStr(arrTmp(i + 4))
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 4)), "转化IJ文件-数据属性映射为空", "基本配置中频度无此数据属性代码")
                End If
                
                '币种
                If dCurrencyRef.exists(CStr(arrTmp(i + 5))) Then
                    c11 = dCurrencyRef(CStr(arrTmp(i + 5)))
                Else
                    c11 = CStr(arrTmp(i + 5))
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 5)), "转化IJ文件-币种映射为空", "基本配置中频度无此币种代码")
                End If
                
                dLoadIDX(CStr(arrTmp(0))) = "," & c1 & "," & "," & c3 & "," & c4 & "," & c5 & "," & c6 & "," & "," & "," & "," & c10 & "," & c11 & "," & CStr(arrTmp(i + 6))
            End If
        Next i
    Next k
    
    f.Close
    
    Set LoadIDX = dLoadIDX
End Function

'
'   AI
'
Private Function LoadIDX_AI(fn As String) As Object
    Dim i As Long, strTmp As String, arrTmp As Variant
    Dim fso As Object, f As Object, sfn As String, strfnIdx As String
    Dim c2 As Date, c3 As String, c4 As String, c5 As String, c6 As String, c12 As String, c13 As Integer
    Dim c1 As String, c10 As String, c11 As String
    Dim dLoadIDX As Object, arr2 As Variant, k As Long, hf As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '读取IDX文件,获取 关键字段 - 业务类，数据属性，币种，单位。多余属性表单名
    Set dLoadIDX = Nothing
    Set dLoadIDX = CreateObject("scripting.dictionary")
    Set f = fso.OpenTextFile(fn, 1)

    strTmp = f.ReadAll
    hf = GetHf(strTmp)
    arr2 = Split(strTmp, hf)
    
    For k = 0 To UBound(arr2) Step 1
        strTmp = CStr(arr2(k))
        
        arrTmp = Split(strTmp, STRSPLIT)
        '进入
        For i = 0 To UBound(arrTmp) Step 9
            If Not dLoadIDX.exists(CStr(arrTmp(i))) Then
                '解析机构类
                c3 = "'" & CStr(arrTmp(i + 2))
                If dOrgRef.exists(replace(c3, "'", "")) Then
                    c4 = CStr(dOrgRef(replace(c3, "'", "")))
                Else
                    c4 = ""
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c3, "转化IJ文件-机构类映射为空", "IJ文件-机构类代码参照缺少该映射")
                End If
                
                '解析地区代码
                c5 = "'" & CStr(arrTmp(i + 3))
                If dAreaRef.exists(replace(c5, "'", "")) Then
                    c6 = CStr(dAreaRef(replace(c5, "'", "")))
                Else
                    c6 = ""
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c5, "转化IJ文件-地区映射为空", "IJ文件-地区代码参照缺少该映射")
                End If
                
                '业务类,通过表单名称来获取的
                c1 = GetBusinessName(CStr(arrTmp(i + 1)))
                
                '数据属性
                If dDataProRef.exists(CStr(arrTmp(i + 4))) Then
                    c10 = dDataProRef(CStr(arrTmp(i + 4)))
                Else
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 4)), "转化IJ文件-数据属性映射为空", "基本配置中频度无此数据属性代码")
                End If
                
                '币种
                If dCurrencyRef.exists(CStr(arrTmp(i + 5))) Then
                    c11 = dCurrencyRef(CStr(arrTmp(i + 5)))
                Else
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 5)), "转化IJ文件-币种映射为空", "基本配置中频度无此币种代码")
                End If
                
                dLoadIDX(CStr(arrTmp(0))) = "," & c1 & "," & "," & c3 & "," & c4 & "," & c5 & "," & c6 & "," & "," & "," & "," & c10 & "," & c11 & "," & CStr(arrTmp(i + 6))
            End If
        Next i
    Next k
    
    f.Close
    
    Set LoadIDX_AI = dLoadIDX
End Function


'
'   留底，逐行读取，当前因为换行符的问题，弃用
'
Private Function LoadIDX_bk(fn As String) As Object
    Dim i As Long, strTmp As String, arrTmp As Variant
    Dim fso As Object, f As Object, sfn As String, strfnIdx As String
    Dim c2 As Date, c3 As String, c4 As String, c5 As String, c6 As String, c12 As String, c13 As Integer
    Dim c1 As String, c10 As String, c11 As String
    Dim dLoadIDX As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '读取IDX文件,获取 关键字段 - 业务类，数据属性，币种，单位。多余属性表单名
    Set dLoadIDX = Nothing
    Set dLoadIDX = CreateObject("scripting.dictionary")
    Set f = fso.OpenTextFile(fn, 1)
    Do Until f.atendofstream
        strTmp = f.readline()
        arrTmp = Split(strTmp, STRSPLIT)
        '进入
        For i = 0 To UBound(arrTmp) Step 10
            If Not dLoadIDX.exists(CStr(arrTmp(i))) Then
                '解析机构类
                c3 = "'" & CStr(arrTmp(i + 2))
                If dOrgRef.exists(replace(c3, "'", "")) Then
                    c4 = CStr(dOrgRef(replace(c3, "'", "")))
                Else
                    c4 = ""
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c3, "转化IJ文件-机构类映射为空", "IJ文件-机构类代码参照缺少该映射")
                End If
                
                '解析地区代码
                c5 = "'" & CStr(arrTmp(i + 3))
                If dAreaRef.exists(replace(c5, "'", "")) Then
                    c6 = CStr(dAreaRef(replace(c5, "'", "")))
                Else
                    c6 = ""
                    Call gloDe.InsertNewDebug("系统导数比较", fn, c5, "转化IJ文件-地区映射为空", "IJ文件-地区代码参照缺少该映射")
                End If
                
                '业务类,通过表单名称来获取的
                c1 = GetBusinessName(CStr(arrTmp(i + 1)))
                
                '数据属性
                If dDataProRef.exists(CStr(arrTmp(i + 4))) Then
                    c10 = dDataProRef(CStr(arrTmp(i + 4)))
                Else
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 4)), "转化IJ文件-数据属性映射为空", "基本配置中频度无此数据属性代码")
                End If
                
                '币种
                If dCurrencyRef.exists(CStr(arrTmp(i + 5))) Then
                    c11 = dCurrencyRef(CStr(arrTmp(i + 5)))
                Else
                    Call gloDe.InsertNewDebug("系统导数比较", fn, CStr(arrTmp(i + 5)), "转化IJ文件-币种映射为空", "基本配置中频度无此币种代码")
                End If
                
                dLoadIDX(CStr(arrTmp(0))) = "," & c1 & "," & "," & c3 & "," & c4 & "," & c5 & "," & c6 & "," & "," & "," & "," & c10 & "," & c11 & "," & CStr(arrTmp(i + 6))
            End If
        Next i
    Loop
    f.Close
    
    Set LoadIDX_bk = dLoadIDX
End Function

Private Function GetDatFileName(fn As String) As String
    Dim strfnIdx As String, path As String
    
    path = left(fn, InStrRev(fn, "\"))
    strfnIdx = right(fn, Len(fn) - InStrRev(fn, "\"))
    
    'BI70204413a0120190814711.IDX
    'BJ70204413a0120190814711.DAT
    GetDatFileName = path & left(strfnIdx, 1) & "J" & Mid(strfnIdx, 3, InStrRev(strfnIdx, ".") - 2) & "DAT"
End Function

Private Function GetBusinessName(tName As String) As String
    If InStr(tName, "A1") Then
        GetBusinessName = "人民币"
    ElseIf InStr(tName, "A2") Then
        GetBusinessName = "外币"
    ElseIf InStr(tName, "A3") Then
        GetBusinessName = "本外币"
    Else
        GetBusinessName = "其他"
    End If
End Function

Private Function GetDataNum(fn As String) As Long
    Dim i As Long, strTmp As String, arr As Variant, n As Long, arr2 As Variant, j As Variant
    Dim hf As String
    Dim fso As Object, f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fn, 1)
    n = 0
    
    Do Until f.atendofstream
        'strTmp = f.readline()
        strTmp = f.ReadAll
        hf = GetHf(strTmp)
        
        arr2 = Split(strTmp, hf)
        
        For j = 0 To UBound(arr2) Step 1
            strTmp = CStr(arr2(j))
            If strTmp <> "" Then
                arr = Split(strTmp, STRSPLIT)
                n = n + (UBound(arr) + 1) / 3
            End If
        Next j
        
    Loop
    
    f.Close

    GetDataNum = n
End Function

'
'   获取日期
'
Private Function GetDateFromIDX(strfnIdx As String) As Date
    'BI70204413a0120190814711.IDX
    Dim y As Integer, m As Integer, d As Integer
    
    y = Int(Mid(strfnIdx, 14, 4))
    m = Int(Mid(strfnIdx, 18, 2))
    d = Int(Mid(strfnIdx, 20, 2))
    
    GetDateFromIDX = DateSerial(y, m, d)
End Function


'
'   异常符号："
'"
'
Private Sub ClearFile(fn As String)
    
End Sub
