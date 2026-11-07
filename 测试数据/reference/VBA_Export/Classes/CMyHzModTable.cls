VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyHzModTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableSettingAvailableDescription As String
Public IsRowColumnToData As String  '是否满足行列与文字数字区域对应

Private AreaAddressUnique As String '验证批注每个只设置了一次

Public AllRangeAddress As String

Public NumAreaAddress As String
Public NumAreaParse As String

Public TextAreaAddress As String
Public TextAreaParse As String

Public CheckAreaAdderss As String
Public CheckAreaParse As String

Public CheckTabelUniqueAddress As String

Public FormatAreaAddress As String
Public FormatAreaParse As String

Public ExCheckAreaAddress As String
Public ExCheckAreaParse As String

Public RowAreaAddress As String
Public RowAreaParse As String
Public dRowAreaParse As Object

Public ColumnAreaAddress As String
Public ColumnAreaParse As String
Public dColumnAreaParse As Object
Public dColumnLevel As Object   ' 储存最

Public SetAreaAddress As String
Public dSetAreaParse As Object

Public dCmt As Object

Public fileName As String
Public wsName As String
Public WsIndex As Integer
Public TableName As String
Public TableDate As String
Public TableDateSetting As String   '保留机构填报的日期
Public TableFillerOrg As String
Public TableFillerOrgSetting As String
Public TableFiller As String
Public TableFillerContact As String
Public dbase_address As Object


'类初始化
Private Sub Class_Initialize()
   'ReDim arr_base_address(1 To 10) As String
End Sub

'类销毁
Private Sub Class_Terminate()
    Set dSetAreaParse = Nothing
    Set dRowAreaParse = Nothing
    Set dColumnAreaParse = Nothing
    Set dbase_address = Nothing
    Set dCmt = Nothing
    Set dColumnLevel = Nothing
End Sub

'
'   提供一个获取批注的功能
'
Public Function GetCmt(i As Variant, j As Variant) As String
    Dim strComment As String
    If dCmt.exists("(" & CStr(CLng(i)) & "," & CStr(ycV(j)) & ")") Then

        GetCmt = CStr(dCmt("(" & CStr(CLng(i)) & "," & CStr(ycV(j)) & ")"))
    Else
        GetCmt = ""
    End If
End Function

Private Sub SortDRowAndCol()
    Dim k As Variant, i As Integer
    
    If RowAreaAddress <> "" Then
        Set dRowAreaParse = SortDicByNumKey(dRowAreaParse)
        k = dRowAreaParse.keys
        For i = 0 To UBound(k) Step 1
            Set dRowAreaParse(k(i)) = SortDicByNumKey(dRowAreaParse(k(i)))
        Next i
    End If
    
    If ColumnAreaAddress <> "" Then
        Set dColumnAreaParse = SortDicByNumKey(dColumnAreaParse)
        k = dColumnAreaParse.keys
        For i = 0 To UBound(k) Step 1
            Set dColumnAreaParse(k(i)) = SortDicByNumKey(dColumnAreaParse(k(i)))
        Next i
    End If
End Sub

'
'   根据字典的keys排序
'
Private Function SortDicByNumKey(d As Object) As Object
    Dim dOut As Object, arr() As Variant, k As Variant, t As Variant, i As Integer, j As Integer, z As Integer
    k = d.keys
    
    ReDim arr(UBound(k))
    '将排序结果存入数组
    For i = 0 To UBound(k) Step 1
        If i = 0 Then
            arr(i) = k(i)
        Else
            For j = 0 To i - 1 Step 1
                If arr(j) > k(i) Then
                    For z = i To j + 1 Step -1
                        arr(z) = arr(z - 1)
                    Next z
                    arr(j) = k(i)
                    Exit For
                End If
            Next j
            If j = i Then arr(i) = k(i)
        End If
    Next i
    
    Set dOut = CreateObject("scripting.dictionary")
    For i = 0 To UBound(arr) Step 1
        dOut.add arr(i), d(arr(i))
    Next i
    
    Set SortDicByNumKey = dOut
End Function


'
'   输出行区域的值
'
Public Sub testdRowParse()
    Dim k1 As Variant, k2 As Variant
    
    For Each k1 In dRowAreaParse.keys
        For Each k2 In dRowAreaParse(k1).keys
            Debug.Print "列" & CStr(k1) & "行" & CStr(k2) & ":" & dRowAreaParse(k1)(k2)
        Next
    Next
End Sub

'
'   形成新的列分类
'
Private Sub ParseColumnsToLevel()
    Set dColumnLevel = CreateObject("scripting.dictionary")
    
    Dim k1 As Variant, k2 As Variant
    Dim pos As Long
    Dim part1 As String, part2 As String, s As String, d As Object, part0 As String
    Dim part_1_0 As String, part_2_0 As String
    Dim t As String, si As Long, j As Long, sei As Long
    ' 要保留一个列数组，和列名称数组，这样的话，才能一起遍历
    t = ""
    si = 0
    For Each k1 In dColumnAreaParse.keys
        part_1_0 = ""
        part_2_0 = ""
        Set d = CreateObject("scripting.dictionary")
        Set dColumnLevel(k1) = CreateObject("scripting.dictionary")
        For Each k2 In dColumnAreaParse(k1).keys
            sei = CLng(k2)
            If si = 0 Then
                si = CLng(k2)
            End If
            s = CStr(dColumnAreaParse(k1)(k2))
            pos = InStrRev(s, ",")
        
            If pos > 0 Then
                part1 = left(s, pos - 1)
                part2 = Mid(s, pos + 1)
            Else
                part1 = s
                part2 = ""  ' 没有逗号的情况
            End If
            
            
            ' 1.没有合并项的情况,连续为控
            If (part_2_0 = "" And part2 = "") Or (part_1_0 = "") Or (part_1_0 = part1) Then
                If t = "" Then
                    t = s & "|" & CStr(k2)
                Else
                    t = t & "|" & s & "|" & CStr(k2)
                End If
            Else
                
                For j = si To CLng(k2) - 1
                    dColumnLevel(k1)(j) = Split(t, "|")
                Next
                Debug.Print "行" & CStr(k1) & "列" & CStr(k2) & ":" & t
                t = s & "|" & CStr(k2)
                si = CLng(k2)
            End If
            
            part_1_0 = part1 ' 看是不是连续的
            part_2_0 = part2
            
        Next
        
        ' 结束以后，还要收尾一次
        If si < sei Then
            For j = si To sei
                dColumnLevel(k1)(j) = Split(t, "|")
            Next
        End If
    Next
    
    
End Sub
    



'
'   输出列区域的值
'
Public Sub testdColParse()
    Dim k1 As Variant, k2 As Variant
    
    ' 找到一个
    For Each k1 In dColumnAreaParse.keys
        For Each k2 In dColumnAreaParse(k1).keys
            
            Debug.Print "行" & CStr(k1) & "列" & CStr(k2) & ":" & ycRight(CStr(dColumnAreaParse(k1)(k2)), ",", 1, False)
        Next
    Next
End Sub

'
'   判断表格设置区域是否合理
'
Private Function GetAvailableSettingDescription() As String
    '判断5个区域单元格数量。
    Dim d As Object, str As String, i As Long, t As Variant, tKey As String, strRel As String
    Set d = CreateObject("scripting.dictionary")
    
    strRel = ""
'    If CheckTabelUniqueAddress = "" And ColumnAreaAddress = "" And RowAreaAddress = "" Then
'        str = wsName & ":" & "表格一致性未设置，请设置表格行区域或列区域或表格一致性批注。"
'        strRel = IIf(strRel = "", str, strRel & Chr(10) & str)
'    End If
    
    If RowAreaAddress = "" Then
        str = wsName & ":" & "表格行区域未设置，请设置表格行区域。"
        strRel = IIf(strRel = "", str, strRel & Chr(10) & str)
    End If
    
    If ColumnAreaAddress = "" Then
        str = wsName & ":" & "表格列区域未设置，请设置表格列区域。"
        strRel = IIf(strRel = "", str, strRel & Chr(10) & str)
    End If
    
    str = ""
    '数字区域
    If NumAreaAddress <> "" Then
        t = Split(NumAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            tKey = "(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")"
            If d.exists(tKey) Then
                '添加进str
                str = IIf(str = "", wsName & ":" & tKey, str & "," & tKey)
            Else
                d(tKey) = 0
            End If
        Next i
    End If
    
    '文字区域
    If TextAreaAddress <> "" Then
        t = Split(TextAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            tKey = "(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")"
            If d.exists(tKey) Then
                '添加进str
                str = IIf(str = "", wsName & ":" & tKey, str & "," & tKey)
            Else
                d(tKey) = 0
            End If
        Next i
    End If
    
    '公式区域
    If FormatAreaAddress <> "" Then
        t = Split(FormatAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            tKey = "(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")"
            If d.exists(tKey) Then
                '添加进str
                str = IIf(str = "", wsName & ":" & tKey, str & "," & tKey)
            Else
                d(tKey) = 0
            End If
        Next i
    End If
    
    '行区域
    If RowAreaAddress <> "" Then
        t = Split(RowAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            tKey = "(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")"
            If d.exists(tKey) Then
                '添加进str
                str = IIf(str = "", wsName & ":" & tKey, str & "," & tKey)
            Else
                d(tKey) = 0
            End If
        Next i
    End If
    
    '列区域
    If ColumnAreaAddress <> "" Then
        t = Split(ColumnAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            tKey = "(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")"
            If d.exists(tKey) Then
                '添加进str
                str = IIf(str = "", wsName & ":" & tKey, str & "," & tKey)
            Else
                d(tKey) = 0
            End If
        Next i
    End If
    
'    '本地校验区域,联表校验区域不去比较了。

    
    If str <> "" Then
        str = str & "等区域出现重复，请合理设置表格区域。"
        strRel = IIf(strRel = "", str, strRel & Chr(10) & str)
    End If
    
    strRel = IIf(strRel = "", AreaAddressUnique, strRel & Chr(10) & AreaAddressUnique)
    'Debug.Print str
    TableSettingAvailableDescription = strRel
End Function


Private Sub Init_dCmt(ws As Worksheet)
    On Error Resume Next
    Dim dr As Object, dc As Object, t As Variant, i As Long, strComment As String
    
    Set dCmt = CreateObject("scripting.dictionary")
    '数字区域
    If NumAreaAddress <> "" Then
        t = Split(NumAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            ' 判断是否存在批注
            strComment = gloEt.GetCommentStr(ws.Cells(ycV(t(i)), ycV(t(i + 1))))
            strComment = replace(strComment, "Administrator:", "")
            strComment = replace(strComment, Chr(10), "")
            dCmt("(" & CStr(t(i)) & "," & CStr(t(i + 1)) & ")") = strComment
        Next i
    End If
End Sub

Private Sub GetRowColumnToDataIndex()
    Dim dr As Object, dc As Object, t As Variant, i As Long
    Set dr = CreateObject("scripting.dictionary")
    Set dc = CreateObject("scripting.dictionary")
    IsRowColumnToData = True
    '行区域
    If RowAreaAddress <> "" Then
        t = Split(RowAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            dr(t(i)) = 1
        Next i
    End If
    
    '列区域
    If ColumnAreaAddress <> "" Then
        t = Split(ColumnAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            dc(t(i + 1)) = 1
        Next i
    End If
    
    '数字区域
    If NumAreaAddress <> "" Then
        t = Split(NumAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            If Not dr.exists(t(i)) Or Not dc.exists(t(i + 1)) Then
                IsRowColumnToData = False
                Exit Sub
            End If
        Next i
    End If
    
    '文字区域
    If TextAreaAddress <> "" Then
        t = Split(TextAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            If Not dr.exists(t(i)) Or Not dc.exists(t(i + 1)) Then
                IsRowColumnToData = False
                Exit Sub
            End If
        Next i
    End If
    
    '公式区域
    If FormatAreaAddress <> "" Then
        t = Split(FormatAreaParse, ",")
        For i = 0 To UBound(t) Step 2
            If Not dr.exists(t(i)) Or Not dc.exists(t(i + 1)) Then
                IsRowColumnToData = False
                Exit Sub
            End If
        Next i
    End If
End Sub

'
'   判断目标表格与表格的一致性
'
Public Function IsUniqueTable(wsMod As Worksheet, wsdes As Worksheet) As Boolean
    
    '判断有没有行列，若有，直接用行列
    
    '若无，判断有没有【表格一致性】批注
    
    IsUniqueTable = True
End Function

'
'   初始化表格信息
'
Public Sub InitTableInfo(ws As Worksheet)
    '获取表格信息：名称、日期、填报人、联系方式
    Call GetTableBase(ws)
    
    '获取设置区域
    Call GetSetArea(ws)
    
    '获取数字区域，解析数字区域
    NumAreaAddress = GetAreabyKey2(ws, "数字区域")
    NumAreaParse = ParseArea(ws, NumAreaAddress)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "数字区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "数字区域"))
    
    '获取文字区域，解析文字区域
    TextAreaAddress = GetAreabyKey2(ws, "文字区域")
    TextAreaParse = ParseArea(ws, TextAreaAddress)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "文字区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "文字区域"))
    
    '获取校验区域，解析校验区域
    CheckAreaAdderss = GetAreabyKey2(ws, "本地校验区域")
    CheckAreaParse = ParseArea(ws, CheckAreaAdderss)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "本地校验区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "本地校验区域"))
    
    '获取校验区域，解析校验区域
    FormatAreaAddress = GetAreabyKey2(ws, "公式区域")
    FormatAreaParse = ParseArea(ws, FormatAreaAddress)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "公式区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "公式区域"))
    
    '获取额外校验区域，解析额外校验区域，额外校验区域需要去格式化
    ExCheckAreaAddress = GetAreabyKey2(ws, "联表校验区域")
    ExCheckAreaParse = ParseArea(ws, ExCheckAreaAddress)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "联表校验区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "联表校验区域"))
    
    '获取行区域，解析行区域
    RowAreaAddress = GetAreabyKey2(ws, "行区域")
    RowAreaParse = ParseArea(ws, RowAreaAddress)
    Set dRowAreaParse = ParseRowsArea(ws, RowAreaAddress)
    'AreaAddressUnique = IIf(AreaAddressUnique = "", CheckAreaKey(ws, "行区域"), AreaAddressUnique & Chr(10) & CheckAreaKey(ws, "行区域"))
    
    '获取列区域，解析列区域
    ColumnAreaAddress = GetAreabyKey2(ws, "列区域")
    ColumnAreaParse = ParseArea(ws, ColumnAreaAddress)
    Set dColumnAreaParse = ParseColumnsArea(ws, ColumnAreaAddress)

    ' 增加一个，近获取相关合并单元格的列
    Call ParseColumnsToLevel

    Call Init_dCmt(ws)
    
    '排序行列字典
    Call SortDRowAndCol
    
    '获取表格一致性校验
    Call GetCheckTabelUniqueAddress(ws)
    
    Call GetAvailableSettingDescription
    
    Call GetRowColumnToDataIndex
    
    
    'Call testdRowParse
    'Call testdColParse
    ' 增加一个获取某一行、某一列的数据区域，他共有多少列
    'Dim t As Variant
    't = GetDataColArr(10, 7)
    'Debug.Print t
End Sub

'
'   获取对应的行名称
'
Public Function GetDataRowName(r As Integer, j As Integer) As String
    Dim g As Integer, i As Integer, i2 As Integer, k1 As Variant, v1 As Variant, k2 As Variant
    
    If RowAreaAddress = "" Or ColumnAreaAddress = "" Then
        GetDataRowName = ""
        Exit Function
    End If
    
    k1 = dRowAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If j > k1(i) And j <= k1(i + 1) Then
            Exit For
        End If
    Next i

    GetDataRowName = dRowAreaParse(k1(i))(r)
End Function

'
'   获取对应列名称
'
Private Function GetDataColName(r As Integer, j As Integer) As String
    If RowAreaAddress = "" Or ColumnAreaAddress = "" Then
        GetDataColName = ""
        Exit Function
    End If
    
    Dim i As Integer, k1 As Variant, d As Object
    k1 = dColumnAreaParse.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i
    
    Set d = dColumnAreaParse(k1(i))
    
    If d.exists(j) Then
        GetDataColName = CStr(d(j))
    Else
        GetDataColName = ""
    End If
End Function

'
'   获取对应列名称
'
Public Function GetDataColArr(r As Integer, j As Integer) As Variant
    If RowAreaAddress = "" Or ColumnAreaAddress = "" Then
        GetDataColArr = Split("", ",")
        Exit Function
    End If
    
    Dim i As Integer, k1 As Variant, d As Object
    k1 = dColumnLevel.keys

    For i = 0 To UBound(k1) - 1 Step 1
        If r > k1(i) And r <= k1(i + 1) Then
            Exit For
        End If
    Next i
    
    Set d = dColumnLevel(k1(i))
    
    If d.exists(j) Then
        GetDataColArr = d(j)
    Else
        GetDataColArr = Split("", "|")
    End If
End Function

Private Sub GetCheckTabelUniqueAddress(ws As Worksheet)
    Dim cmt As Comment, str As String
    str = ""
    For Each cmt In ws.Comments
        If InStr(cmt.text, "表格一致性校验") > 0 Then
            str = cmt.Parent.address
            Exit For
        End If
    Next cmt
    If str = "" Then
        '输出日记
        If ColumnAreaAddress = "" And RowAreaAddress = "" Then
            'Call gloDe.InsertNewDebug("模板初始化", ws.Parent.Name, ws.Name, "模板风险", "模板未设置一致性校验标志，存在识别错表的风险。")
        End If
    End If
    CheckTabelUniqueAddress = str
End Sub

Private Sub GetSetArea(ws As Worksheet)
    Dim cmt As Comment, str As String
    str = ""
    Set dSetAreaParse = CreateObject("scripting.dictionary")
    For Each cmt In ws.Comments
        If InStr(cmt.text, "设置") > 0 Then
            If str <> "" Then str = str + ","
            str = str + cmt.Parent.address
            dSetAreaParse(CStr(ws.name + "-" + cmt.Parent.address + "-" + cmt.text)) = cmt.Parent.address
        End If
    Next cmt
    SetAreaAddress = str
End Sub

Private Function ParseRowsArea(ws As Worksheet, address As String) As Object
    Dim k As Integer, obj As Object, rn As Range, arr As Variant, obj2 As Object
    If address = "" Then Exit Function
    
    Set rn = ws.Range(address)
    If rn.count < 1 Then Exit Function
    Set obj2 = CreateObject("scripting.dictionary")
    
    arr = Split(address, ",")
    
    Dim col1 As Integer, col2 As Integer, row1 As Integer, row2 As Integer
    Dim i As Integer, j As Integer, str As String, rnT As Range, strRel As String
    For k = 0 To UBound(arr) Step 1
        Set rn = ws.Range(CStr(arr(k)))
        
        col1 = rn(1).Column: col2 = rn(rn.count).Column
        row1 = rn(1).row: row2 = rn(rn.count).row
        
        If Not obj2.exists(col2) Then
            Set obj = CreateObject("scripting.dictionary")
        Else
            Set obj = obj2(col2)
        End If
        
        For i = row1 To row2 Step 1
            strRel = ""
            For j = col2 To col1 Step -1
                str = ""
                Set rnT = ws.Cells(i, j)
                If rnT.MergeCells = True Then
                    '存在合并单元格
                    str = CStr(rnT.MergeArea(1).value)
                    j = j - rnT.MergeArea.columns.count + 1
                Else
                    '不存在
                    str = CStr(rnT.value)
                End If
                
                strRel = IIf(strRel = "", str, IIf(str = "", strRel, str & "," & strRel))
            Next j
            'Debug.Print "第" & CStr(i) & "行:" & strRel
            obj(i) = strRel
        Next i
        Set obj2(col2) = obj
        Set obj = Nothing
    Next k

    Set ParseRowsArea = obj2
End Function

'
'   解析列
'
Private Function ParseColumnsArea(ws As Worksheet, address As String) As Object
    Dim k As Integer, obj As Object, rn As Range, arr As Variant, obj2 As Object
    If address = "" Then Exit Function
    
    Set rn = ws.Range(address)
    If rn.count < 1 Then Exit Function
    
    Set obj2 = CreateObject("scripting.dictionary")
    
    arr = Split(address, ",")
    
    Dim col1 As Integer, col2 As Integer, row1 As Integer, row2 As Integer
    Dim i As Integer, j As Integer, str As String, rnT As Range, strRel As String
    For k = 0 To UBound(arr) Step 1
        Set rn = ws.Range(CStr(arr(k)))
        
        col1 = rn(1).Column: col2 = rn(rn.count).Column
        row1 = rn(1).row: row2 = rn(rn.count).row
        
        If Not obj2.exists(row2) Then
            Set obj = CreateObject("scripting.dictionary")
        Else
            Set obj = obj2(row2)
        End If
        
        For j = col1 To col2 Step 1
            strRel = ""
            For i = row2 To row1 Step -1
                str = ""
                Set rnT = ws.Cells(i, j)
                If rnT.MergeCells = True Then
                    '存在合并单元格
                    str = CStr(rnT.MergeArea(1).value)
                    i = i - rnT.MergeArea.rows.count + 1
                Else
                    '不存在
                    str = CStr(rnT.value)
                End If
                
                strRel = IIf(strRel = "", str, IIf(str = "", strRel, str & "," & strRel))
            Next i
            'Debug.Print "第" & CStr(j) & "列:" & strRel
            obj(j) = strRel
        Next j
        Set obj2(row2) = obj
        Set obj = Nothing
    Next k

    Set ParseColumnsArea = obj2
End Function


Private Function ParseArea(ws As Worksheet, address As String) As String
    Dim i As Long, str As String, rn As Range, j As Long, t As Variant
    If address = "" Then Exit Function
    
    't = Split(address, ",")
    str = ""
    
    For Each rn In ws.Range(address)
        str = IIf(str = "", CStr(rn.row) + "," + CStr(rn.Column), str & "," & CStr(rn.row) + "," + CStr(rn.Column))
    Next
    
'    For j = 0 To UBound(t) Step 1
'        Set rn = ws.Range(CStr(t(j)))
'        For i = 1 To rn.count Step 1
'            If str <> "" Then str = str + ","
'           str = str + CStr(rn(i).row) + "," + CStr(rn(i).Column)
'        Next i
'    Next j
    
    ParseArea = str
End Function

Private Function CheckAreaKey(ws As Worksheet, key As String) As String
    Dim c As Integer, c2 As Integer, str As String, cmt As Comment, j As String, t As String
    
    str = ""
    j = 1
    Do While j > 0
        t = key & CStr(j)
        For Each cmt In ws.Comments
            If InStr(cmt.text, t) > 0 Then
                c = c + 1
            End If
        Next cmt
        
        t = key & "#" & CStr(j)
        For Each cmt In ws.Comments
            If InStr(cmt.text, t) > 0 Then
                c2 = c2 + 1
            End If
        Next cmt
        
        If c > 1 Then str = IIf(str = "", t & "设置存在" & CStr(c) & "个。", str & Chr(10) & t & "设置存在" & CStr(c) & "个。")
        If c2 > 1 Then str = IIf(str = "", t & "设置存在" & CStr(c2) & "个。", str & Chr(10) & t & "设置存在" & CStr(c2) & "个。")
        If c2 > c Then str = IIf(str = "", t & "设置不对称。", str & Chr(10) & t & "设置不对称。")
        j = j + 1
        
        If c2 = 0 And c = 0 Then Exit Do
        c = 0
        c2 = 0
    Loop
    
    
    CheckAreaKey = str
End Function

'
'  根据关键字，获取相应区域
'
Private Function GetAreabyKey2(ws As Worksheet, key As String) As String
    Dim i As Integer, j As Integer, str As String, s1 As String, s2 As String
    Dim key_index As Integer '存放关键字末尾位置
    Dim key_str As String '当前得到的关键字
    Dim left_top As String, left_top_address As String
    Dim right_bottom As String, right_bottom_address As String
    Dim str_err As String
    '用来计算重复
    Dim d_ltkey As Object, d_rbkey As Object, d As Object, k As Variant, v As Variant
    Set d_ltkey = CreateObject("scripting.dictionary")
    Set d_rbkey = CreateObject("scripting.dictionary")
    
    str = ""
    '遍历所有的批注,找到所有左上角
    Dim cmt As Comment, cmt2 As Comment
    
    
    
    For Each cmt In ws.Comments
        ' 如果包含关键字段，假设XXX 行区域12 XXX
        left_top = Get_Key_Str(cmt.text, key, True)
        left_top_address = cmt.Parent.address
        If left_top <> "" Then
            '不为空，说明是左上角,继续寻找右上角
            If Not d_ltkey.exists(left_top) Then
                Set d = CreateObject("scripting.dictionary")
                d("address") = left_top_address
                d("count") = 1
                Set d_ltkey(left_top) = d
            Else
                Set d = d_ltkey(left_top)
                d("count") = ycV(d("count")) + 1
                d("address") = CStr(d("address")) & "," & left_top_address
                Set d_ltkey(left_top) = d
            End If
        End If
    Next
    Set cmt = Nothing
    '遍历所有右下角
    For Each cmt In ws.Comments
        ' 如果包含关键字段，假设XXX 行区域12 XXX
        right_bottom = Get_Key_Str(cmt.text, key & "#", True)
        right_bottom_address = cmt.Parent.address
        If right_bottom <> "" Then
            '不为空，说明是左上角,继续寻找右上角
            If Not d_rbkey.exists(right_bottom) Then
                Set d = CreateObject("scripting.dictionary")
                d("address") = right_bottom_address
                d("count") = 1
                Set d_rbkey(right_bottom) = d
            Else
                Set d = d_rbkey(right_bottom)
                d("count") = ycV(d("count")) + 1
                d("address") = CStr(d("address")) & "," & right_bottom_address
                Set d_rbkey(left_top) = d
            End If
        End If
    Next
    Set cmt = Nothing
    '遍历字典，先找到设置数量重复的
    For Each k In d_ltkey.keys
        Set d = d_ltkey(k)
        If ycV(d("count")) > 1 Then
            str_err = ws.name & ":" & CStr(k) & "设置数量超过1个(" & CStr(ycV(d("count"))) & "个);"
            AreaAddressUnique = IIf(AreaAddressUnique = "", str_err, AreaAddressUnique & Chr(10) & str_err)
        End If
    Next
    For Each k In d_rbkey.keys
        Set d = d_rbkey(k)
        If ycV(d("count")) > 1 Then
            str_err = ws.name & ":" & CStr(k) & "设置数量超过1个(" & CStr(ycV(d("count"))) & "个);"
            AreaAddressUnique = IIf(AreaAddressUnique = "", str_err, AreaAddressUnique & Chr(10) & str_err)
        End If
    Next
    
    '寻找对称
    For Each k In d_ltkey.keys
        right_bottom = key & "#" & replace(CStr(k), key, "")
        If Not d_rbkey.exists(right_bottom) Then
            str_err = ws.name & ":" & CStr(k) & "未能找到右下角(" & right_bottom & ");"
            AreaAddressUnique = IIf(AreaAddressUnique = "", str_err, AreaAddressUnique & Chr(10) & str_err)
        End If
    Next
    For Each k In d_rbkey.keys
        left_top = replace(CStr(k), "#", "")
        If Not d_ltkey.exists(left_top) Then
            str_err = ws.name & ":" & CStr(k) & "未能找到左上角(" & left_top & ");"
            AreaAddressUnique = IIf(AreaAddressUnique = "", str_err, AreaAddressUnique & Chr(10) & str_err)
        End If
    Next
    
    '最后形成地址,最后是模板正确的化，就会一一对应，所以只需要遍历一个
    For Each k In d_ltkey.keys
        Set d = d_ltkey(k)
        left_top_address = CStr(d("address"))
        right_bottom = key & "#" & replace(CStr(k), key, "")
        If d_rbkey.exists(right_bottom) Then
            Set d = d_rbkey(right_bottom)
            right_bottom_address = CStr(d("address"))
            If str <> "" Then str = str & ","
            str = str & left_top_address & ":" & right_bottom_address
        End If
    Next
    GetAreabyKey2 = str
End Function

'
'  判断关键字在文件中是否独立存在,其实就看分隔符有多少
'
Private Function Get_Key_Str(text As String, key As String, Optional is_start As Boolean = True) As String
    '对text进行预处理
    Dim str As String
    '先判断存不存在，不存在直接跳过
    If InStr(text, key) <= 0 Then
        Get_Key_Str = ""
        Exit Function
    End If
    str = replace(text, vbLf, " ")
    str = replace(str, vbCr, " ")
    str = replace(str, vbCrLf, " ")
    str = replace(str, "|", " ")
    '中文转英文
    str = replace(str, "【", "[")
    str = replace(str, "】", "]")
    str = replace(str, "（", "(")
    str = replace(str, "）", ")")
    str = replace(str, "｛", "{")
    str = replace(str, "｝", "}")
    str = replace(str, "，", ",")
    str = replace(str, "　", " ")
    
    str = replace(str, "[", " ")
    str = replace(str, "]", " ")
    str = replace(str, "(", " ")
    str = replace(str, ")", " ")
    str = replace(str, "{", " ")
    str = replace(str, "}", " ")
    str = replace(str, ",", " ")
    
    '下一步，进行空格分割
    Dim arr As Variant, i As Long
    arr = Split(str, " ")
    For i = 0 To UBound(arr)
        '如果包含关键字，又不包含#号的，如果有两个，则说明有问题
        If is_start = True Then
            '找到左上角命名
            If InStr(CStr(arr(i)), key) > 0 And InStr(CStr(arr(i)), key & "#") <= 0 Then
                Get_Key_Str = CStr(arr(i))
                Exit Function
            End If
        Else
            If CStr(arr(i)) = key Then
                Get_Key_Str = CStr(arr(i))
                Exit Function
            End If
        End If
    Next
    '没有找到，说明就不是这个单元格
    Get_Key_Str = ""
End Function


'
'  根据关键字，获取相应区域
'
Private Function GetAreabyKey(ws As Worksheet, key As String) As String
    Dim i As Integer, j As Integer, str As String, s1 As String, s2 As String

    Do While j > 0
        s1 = gloEt.FindComment(ws, key & CStr(j), 1)
        
        If s1 <> "" Then
            If j > 1 Then str = str + ","
            
            str = str + s1
            s2 = gloEt.FindComment(ws, key & "#" & CStr(j), 1)
            If s2 <> "" Then
                str = str + ":" + s2
            End If
        Else
            Exit Do
        End If
        j = j + 1
    Loop
        
    GetAreabyKey = str
End Function


'
'   获取表格基本信息
'
Sub GetTableBase(ws As Worksheet)
    Dim str As String, arr As Variant, rn As Variant
    Set dbase_address = Nothing
    Set dbase_address = CreateObject("scripting.dictionary")
    With ws
        On Error Resume Next
        '文件名称
        fileName = left(.Parent.fullName, InStrRev(.Parent.fullName, ".") - 1)
        arr = Split(fileName, "\")
        fileName = arr(UBound(arr))
        wsName = ws.name
        WsIndex = ws.index
        '表格名称
        TableName = ws.name
        
        Set rn = .UsedRange
        AllRangeAddress = .Range(.Cells(1, 1), .Cells(rn(rn.count).row, rn(rn.count).Column)).address
        
        str = "表格名称"
        If gloEt.FindComment(ws, str) <> "" Then
            dbase_address(CStr(gloEt.FindComment(ws, str))) = 1
            TableName = CStr(.Range(CStr(gloEt.FindComment(ws, str))).value)
            If TableName = "" Then TableName = ws.name
        End If
        
        Dim cmt As Comment
        For Each cmt In ws.Comments
            If InStr(cmt.text, "表格日期") > 0 Or InStr(cmt.text, "数据日期") > 0 Then
                TableDate = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
                dbase_address(CStr(cmt.Parent.address)) = 1
                If InStr(cmt.text, "设置日期") > 0 Then
                    TableDateSetting = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
                End If
            ElseIf InStr(cmt.text, "填报人") > 0 Then
                dbase_address(CStr(cmt.Parent.address)) = 1
                TableFiller = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
            ElseIf InStr(cmt.text, "填报单位") > 0 Then
                dbase_address(CStr(cmt.Parent.address)) = 1
                TableFillerOrg = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
            ElseIf InStr(cmt.text, "联系方式") > 0 Then
                dbase_address(CStr(cmt.Parent.address)) = 1
                TableFillerContact = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
            ElseIf InStr(cmt.text, "设置单位") > 0 Then
                TableFillerOrgSetting = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
            ElseIf InStr(cmt.text, "设置日期") > 0 Then
                TableDateSetting = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
                If InStr(cmt.text, "表格日期") > 0 Then
                    TableDate = CStr(cmt.Parent.row) & "," & CStr(cmt.Parent.Column)
                End If
            End If
        Next cmt
    End With
End Sub
