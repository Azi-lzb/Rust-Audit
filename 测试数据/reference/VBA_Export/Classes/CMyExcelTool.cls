VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyExcelTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'类初始化
Private Sub Class_Initialize()

End Sub

'类销毁
Private Sub Class_Terminate()

End Sub




Public Function GetUsedRange(ws As Worksheet) As Range
    Dim rn As Range
    Call ReduceWorkSheet(ws)
    If IsBlankSheet(ws) Then
        Set rn = Nothing
    Else
        Set rn = ws.Range(ws.Range("A1"), ws.UsedRange)
    End If
    Set GetUsedRange = rn
End Function

'
' 对工作表进行减肥
'
Sub ReduceWorkSheet(ws As Worksheet, Optional IsDeleleEmptyWs As Boolean = False)
    Dim RowNum As Long, ColNum As Long, c As Long, cell As Range, j As Long, k As Long, sht As Worksheet, a As Shape
    Dim rnLastRow As Range, rnLastColumn As Range, wsfz As Worksheet
    On Error GoTo err
  
    Dim abOldStatus(1 To OLD_SATUS_NUM) As Boolean
    Call gloCft.ChangeSomethingStatus(abOldStatus, False)

    With ws
        '1.删除空表
        If WorksheetFunction.CountA(.Cells) = 0 Then
            If IsDeleleEmptyWs = True Then .Delete
            GoTo line '如果是空表则删除
        End If
        
        Set wsfz = FindWs(ThisWorkbook, "yc辅助区", 0, True)
        wsfz.Cells.Clear    '清空
        
        '2.删除UsedRange中的空白区域格式信息,需要先取消隐藏，再复原隐藏。使用定位
        RowNum = .Range(.Range("A1"), .UsedRange).rows.count
        If RowNum > wsfz.Cells.rows.count Then RowNum = wsfz.Cells.rows.count
        ColNum = .Range(.Range("A1"), .UsedRange).columns.count
        If ColNum > wsfz.Cells.columns.count Then ColNum = wsfz.Cells.columns.count
        Set rnLastRow = .Range(.Cells(RowNum, 1), .Cells(RowNum, ColNum))
        Set rnLastColumn = .Range(.Cells(1, ColNum), .Cells(RowNum, ColNum))
        
        If Application.WorksheetFunction.CountA(rnLastRow) > 0 And _
            Application.WorksheetFunction.CountA(rnLastColumn) > 0 Then
            '说明不需要删除
            GoTo line
        End If
        
        '需要判断两边的范围,取小的那个
        wsfz.Range(.Range(.Cells(1, 1), .Cells(RowNum, ColNum)).address).value = .Range(.Cells(1, 1), .Cells(RowNum, ColNum)).value
        
        '通过辅助区获取删除范围
        If Application.WorksheetFunction.CountA(rnLastRow) = 0 Then
            j = 1
            For Each cell In wsfz.Range(rnLastRow.address)
                If j < cell.End(xlUp).row Then j = cell.End(xlUp).row
            Next
            .Range(.Cells(j + 1, 1), .Cells(.rows.count, 1)).EntireRow.Delete
        End If
        
        If Application.WorksheetFunction.CountA(rnLastColumn) = 0 Then
            k = 1
            For Each cell In wsfz.Range(rnLastColumn.address)
                If k < cell.End(xlToLeft).Column Then k = cell.End(xlToLeft).Column
            Next
            .Range(.Cells(1, k + 1), .Cells(1, .columns.count)).EntireColumn.Delete Shift:=xlToLeft
        End If
        wsfz.Cells.Clear    '清空
    
        '3.删除隐藏的图片
        'For Each a In .Shapes
        '    If a.Visible = msoFalse Then a.Delete
        'Next
    End With
line:
    
    Call gloCft.RestoreSomethingStatus(abOldStatus)
Exit Sub

err:
    Debug.Print "ReduceWorkSheet出现异常" & err.Description
    err.Clear
End Sub

Sub ReduceWorkbook(wb As Workbook, Optional IsDeleleEmptyWs As Boolean = False)
    On Error GoTo err
    
    Dim sht As Worksheet
  
    Dim abOldStatus(1 To OLD_SATUS_NUM) As Boolean
    Call gloCft.ChangeSomethingStatus(abOldStatus, False)

    For Each sht In wb.Worksheets '仅对工作表进行操作，忽略图表、宏表
        Call ReduceWorkSheet(sht)
    Next sht
    
    Call gloCft.RestoreSomethingStatus(abOldStatus)
err:
End Sub

'
'   复制报表打印格式
'
Public Sub CopyPrintFormat(wsdes As Worksheet, wsSrc As Worksheet)
    With wsdes
        '页面
        .PageSetup.Orientation = wsSrc.PageSetup.Orientation    '方向
        .PageSetup.Zoom = wsSrc.PageSetup.Zoom  '缩放
        .PageSetup.PaperSize = wsSrc.PageSetup.PaperSize    '纸张大小
        .PageSetup.FirstPageNumber = wsSrc.PageSetup.FirstPageNumber    '起始页码
        
        '页边距
        .PageSetup.LeftMargin = wsSrc.PageSetup.LeftMargin
        .PageSetup.RightMargin = wsSrc.PageSetup.RightMargin
        .PageSetup.TopMargin = wsSrc.PageSetup.TopMargin
        .PageSetup.BottomMargin = wsSrc.PageSetup.BottomMargin
        .PageSetup.HeaderMargin = wsSrc.PageSetup.HeaderMargin
        .PageSetup.FooterMargin = wsSrc.PageSetup.FooterMargin
        .PageSetup.CenterHorizontally = wsSrc.PageSetup.CenterHorizontally
        .PageSetup.CenterVertically = wsSrc.PageSetup.CenterVertically
        
        '页眉/页脚
        .PageSetup.LeftHeader = wsSrc.PageSetup.LeftHeader
        .PageSetup.CenterHeader = wsSrc.PageSetup.CenterHeader
        .PageSetup.RightHeader = wsSrc.PageSetup.RightHeader
        .PageSetup.LeftFooter = wsSrc.PageSetup.LeftFooter
        .PageSetup.CenterFooter = wsSrc.PageSetup.CenterFooter
        .PageSetup.RightFooter = wsSrc.PageSetup.RightFooter
        .PageSetup.OddAndEvenPagesHeaderFooter = wsSrc.PageSetup.OddAndEvenPagesHeaderFooter    '奇偶页不同
        .PageSetup.DifferentFirstPageHeaderFooter = wsSrc.PageSetup.DifferentFirstPageHeaderFooter  '首页不同
        .PageSetup.ScaleWithDocHeaderFooter = wsSrc.PageSetup.ScaleWithDocHeaderFooter  '随文档自动缩放
        .PageSetup.AlignMarginsHeaderFooter = wsSrc.PageSetup.AlignMarginsHeaderFooter  '与页边距对齐
        
        '工作表
        .PageSetup.PrintArea = wsSrc.PageSetup.PrintArea
        .PageSetup.PrintTitleRows = wsSrc.PageSetup.PrintTitleRows
        .PageSetup.PrintTitleColumns = wsSrc.PageSetup.PrintTitleColumns
        .PageSetup.PrintGridlines = wsSrc.PageSetup.PrintGridlines
        .PageSetup.PrintComments = wsSrc.PageSetup.PrintComments
        .PageSetup.BlackAndWhite = wsSrc.PageSetup.BlackAndWhite
        .PageSetup.Draft = wsSrc.PageSetup.Draft
        .PageSetup.PrintErrors = wsSrc.PageSetup.PrintErrors
        .PageSetup.order = wsSrc.PageSetup.order
    End With
End Sub

Public Function GetCurRows2(ws As Worksheet, Optional colName As String = "A") As Long
    Dim row As Long, i As Long
    On Error Resume Next
    With ws
        row = .Range(colName & CStr(.Cells.rows.count)).End(xlUp).row
        If err.Number <> 0 Then
          row = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row
          colName = "A"
        End If
        
        i = 2
        If CStr(.Range(colName & CStr(row + 1))) <> "" Then
            '说明有隐藏的
            Do While i > 1
                If CStr(.Range(colName & CStr(row + i))) <> "" Then
                    i = i + 1
                Else
                    GetCurRows2 = row + i - 1
                    Exit Do
                End If
            Loop
        Else
            GetCurRows2 = row
        End If
    End With
End Function


''
'   获取的当前行
'
Public Function GetCurRows(ws As Worksheet, Optional colName As String = "A") As Long
    Dim row As Long, i As Long, rn As Range, arr As Variant
    On Error Resume Next
    Set rn = GetUsedRange(ws)
    If rn Is Nothing Then
        row = 1
    Else
        arr = rn.value
        row = UBound(arr)
    End If
    
    GetCurRows = row
End Function

'
'   返回提出标题行，工作表的所有区域
'
Public Function GetDeleteTileArea(ws As Worksheet, lngTitleCount As Long, Optional mode As Integer = 1) As String
    Dim str As String, arr As Variant, rn As Range, r1 As Long, r2 As Long, c1 As Integer, c2 As Integer
    
    If IsBlankSheet(ws) Then
        GetDeleteTileArea = ""
        Exit Function
    End If
    
    If mode = 1 Then
        arr = ws.UsedRange.value
        Set rn = ws.UsedRange
    Else
        arr = ws.Range("A1").CurrentRegion.value
        Set rn = ws.Range("A1").CurrentRegion
    End If
    r1 = rn(1).row
    c1 = rn(1).Column
    r2 = rn(rn.count).row
    c2 = rn(rn.count).Column
    
    If r2 <= lngTitleCount Then
        GetDeleteTileArea = ""
        Exit Function
    End If
    
    str = ws.Range(ws.Cells(lngTitleCount + 1, c1), ws.Cells(r2, c2)).address
    
    GetDeleteTileArea = str
End Function


'
'   获取工作簿,若没打开，则打开相应excel文件,若不存在，则返回Nothing
'
Public Function GetWb(path As String) As Workbook
    If path = "" Then
        Set GetWb = Nothing
        Exit Function
    End If
    'Call OpenSomething
    Dim wb As Workbook, r As Boolean
    
    '如果是.xlsx文件，则转为.xls
    'path = XlsxToXls(path)
    r = True
    For Each wb In Workbooks
        If wb.fullName = path Then   '查询是否开启的文件名
            Set GetWb = wb  'Workbooks(path)
            r = False
            Exit Function
        End If
    Next
    
    On Error GoTo errhandle
    Dim fso As Object, fhz As String
    'fhz = Right(path, Len(path) - InStrRev(path, "."))
    
    If r = True Then
        'Set GetWb = Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False)
        If ycV(Application.Version) < 11 Then
            Set GetWb = ycWbOpen09(path)
        Else
            Set GetWb = ycWbOpen10(path)
        End If
    End If
    'Call CloseSomething
Exit Function

errhandle:
    Set GetWb = Nothing
    Set fso = Nothing
    'Call CloseSomething
    Debug.Print "程序打开" & path & "失败,文件路径错误或兼容性异常。"
End Function

Private Function ycWbOpen09(path As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False)
    If err.Number <> 0 Then
        Set ycWbOpen09 = Nothing
        'Debug.Print "ycWbOpen09" & Err.Description
        err.Clear
    Else
        Set ycWbOpen09 = wb
    End If
End Function

Private Function ycWbOpen10(path As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False)
    If err.Number <> 0 Then
        Set wb = Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False, corruptload:=xlExtractData)
        If err.Number <> 0 Then
            Set ycWbOpen10 = ycWbOpen09(path)
            'Debug.Print "ycWbOpen10" & Err.Description
            err.Clear
        Else
            Set ycWbOpen10 = wb
        End If
    Else
        Set ycWbOpen10 = wb
    End If
End Function

'
'   将xlsx转为xls
'
Public Function XlsxToXls(path As String, forlder As String) As String

    On Error GoTo errhandle
    Dim Hz As String, newPath As String, wb As Workbook, name As String
    
    name = gloFt.getShortFileName2(path)
    
    Hz = right(path, Len(path) - InStrRev(path, "."))
    If right(forlder, 1) <> "\" Then forlder = forlder & "\"
    If Hz = "xlsx" Then
        newPath = forlder + left(name, InStrRev(name, ".")) + "xls"
        Set wb = Workbooks.Open(path, UpdateLinks:=False)
        wb.SaveAs newPath, FileFormat:=xlExcel8
        wb.Close False
    Else
        newPath = forlder + name
        Set wb = Workbooks.Open(path, UpdateLinks:=False)
        wb.SaveAs newPath, FileFormat:=xlExcel8
        wb.Close False
    End If
    
    XlsxToXls = newPath
    
Exit Function

errhandle:
    XlsxToXls = ""
End Function

Public Sub save_as_excel(ByRef wb As Workbook, path As String)
    If right(path, 5) = ".xlsx" Then
        wb.SaveAs path, FileFormat:=51  'xlOpenXMLWorkbook 51
    ElseIf right(path, 4) = ".xls" Then
        wb.SaveAs path, FileFormat:=56  'xlExcel8
    ElseIf right(path, 4) = ".csv" Then
        wb.SaveAs path, FileFormat:=6   'xlCSV 6,xlCSVUTF8  62
    Else
        wb.SaveAs path
    End If
End Sub


'
'   将xls转为xlsx
'
Public Function XlsToXlsx(path As String, forlder As String) As String

    On Error GoTo errhandle
    Dim Hz As String, newPath As String, wb As Workbook, name As String
    
    name = gloFt.getShortFileName2(path)
    
    Hz = right(path, Len(path) - InStrRev(path, "."))
    If right(forlder, 1) <> "\" Then forlder = forlder & "\"
    If Hz = "xls" Then
        newPath = forlder + left(name, InStrRev(name, ".")) + "xlsx"
        Set wb = Workbooks.Open(path, UpdateLinks:=False)
        wb.SaveAs newPath, FileFormat:=xlOpenXMLWorkbook
        wb.Close False
    Else
        newPath = forlder + name
        Set wb = Workbooks.Open(path, UpdateLinks:=False)
        wb.SaveAs newPath, FileFormat:=xlOpenXMLWorkbook
        wb.Close False
    End If
    
    XlsToXlsx = newPath
    
Exit Function

errhandle:
    XlsToXlsx = ""
End Function

'单元格C值与Excel表格行值字母转换函数
Public Function CtoExC(i As Integer) As String
    If i < 1 Then
        CtoExC = Null
        Exit Function
    End If
        
    If i <= 26 Then
        CtoExC = Chr(65 + i - 1)
    Else
        CtoExC = Chr(65 + (i - (i - 1) Mod 26) / 26 - 1) + Chr(65 + (i - 1) Mod 26)
    End If
End Function

'
'   返回单元格所在位置
'

'
'   查找工作表，可模糊可精确，若不存在可增加
'   返回查找到的工作表
'   m = 0 精准，m = 0 包含
'
'
Public Function FindWs(wb As Workbook, key As String, m As Integer, Optional add As Boolean = False) As Worksheet
    On Error GoTo errhandle
    Dim ws As Worksheet, sum As Integer, i As Integer, r
    r = 0
    
    For i = 1 To wb.Worksheets.count Step 1
        Set ws = wb.Worksheets(i)
        If m = 0 Then
            If ws.name = key Then GoTo MyFind
        Else
            If InStr(ws.name, key) > 0 Then GoTo MyFind
        End If
    Next i
    
    '如果执行到这里，说明没有找到
    If add Then
        '创建一个
        Set ws = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = key
        GoTo MyFind
    End If
    
    Set FindWs = Nothing
Exit Function

MyFind:
    Set FindWs = ws
Exit Function

errhandle:
    err.Clear
    err.Raise Number:=1062, Description:="创建失败，工作簿已满或其他."
    Set FindWs = Nothing
End Function

'
'   单独保存某个工作表
'
Public Function SaveSheetAsWorkbook(fn As String, ws As Worksheet) As Boolean
    On Error GoTo errhandle
    
    ws.Copy
    ActiveWorkbook.Close SaveChanges:=True, _
              fileName:=fn
    SaveSheetAsWorkbook = True
Exit Function

errhandle:
    Call gloDe.InsertNewDebug("程序保存" & fn & "失败 -- " & err.Description)
    err.Clear
    SaveSheetAsWorkbook = False
End Function

'
'   在工作表范围中中，搜索指定单元格，可偏移，默认不偏移
'   m xlWhole,xlPart
'
'
Public Function FindRng(rng As Range, key As String, m As Long, Optional ofr As Integer = 0, Optional ofc As Integer = 0) As Range
    Dim c As Range
    With rng
        Set c = .find(key, LookIn:=xlValues, LookAt:=m)
    End With
    
    If Not c Is Nothing Then
        Set FindRng = c.Offset(ofr, ofc)
    Else
        Set FindRng = Nothing
    End If
End Function

'
'   复制相应区域到另一个工作表
'   wsSrc 源工作表，wsDec 目的工作表,addr 复制区域
'   返回复制区域
'   默认无格式粘贴
'
Function CopyAndPaste(wsSrc As Worksheet, wsdes As Worksheet, Optional addr As String = "1:65536", Optional isValue As Boolean = True, Optional isFormal As Boolean = False) As String
    Dim abOldStatus(1 To OLD_SATUS_NUM) As Boolean
    Call gloCft.ChangeSomethingStatus(abOldStatus, False)
        
    If addr = "1:65536" Then
        addr = GetUsedRange(wsSrc).address
    End If
    
    wsdes.Range(addr).Clear
    wsSrc.Range(addr).Copy
    'wsDes.Range(addr).value = wsSrc.Range(addr).value
    wsdes.Range(addr).PasteSpecial Paste:=xlPasteFormats, operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    If isValue Then
        If isFormal Then
            wsdes.Range(addr).Formula = wsSrc.Range(addr).Formula
            wsdes.Range(addr).PasteSpecial Paste:=xlPasteComments, operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
        Else
            wsdes.Range(addr).value = wsSrc.Range(addr).value
        End If
    Else
        wsdes.Range(addr).PasteSpecial Paste:=xlPasteAll, operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    End If
    
    Application.CutCopyMode = False
    CopyAndPaste = addr
    Call gloCft.RestoreSomethingStatus(abOldStatus)
End Function

'
'   1.给某个表格的单元格添加批注,默认为A1
'   wbs 制定表格，str 批注内容，rnStr 可选单元格
'
Public Sub InsertComment(ws As Worksheet, str As String, Optional rnStr As String = "A1")
    On Error Resume Next
    With ws.Range(rnStr)
        If .Comment Is Nothing Then
            .AddComment text:=str
        Else
            If str = "" Then
                .ClearComments
            Else
                .Comment.text str
            End If
        End If
    End With
End Sub

Public Function GetCommentStr(rn As Range) As String
    If Not rn.Comment Is Nothing Then
         GetCommentStr = rn.Comment.text
    Else
        GetCommentStr = ""
    End If
End Function

'
'   查找某个批注,0 - 精确,1 - 模糊
'
Public Function FindComment(ws As Worksheet, str As String, Optional m As Integer = 0) As String
    Dim cmt As Comment
    For Each cmt In ws.Comments
        If m = 0 Then
            If cmt.text = str Then
                FindComment = cmt.Parent.address
                Exit Function
            End If
        Else
            If InStr(cmt.text, str) > 0 Then
                FindComment = cmt.Parent.address
                Exit Function
            End If
        End If
        
    Next cmt
    FindComment = ""
End Function

Public Sub ShowData(wbs As Worksheet)
    On Error Resume Next
    With wbs
        If .AutoFilterMode Then
            If .FilterMode = True Then
                .ShowAllData
            End If
            .UsedRange.EntireColumn.AutoFit
            .Range("A1").AutoFilter
        Else
            .UsedRange.EntireColumn.AutoFit
            .Range("A1").AutoFilter
        End If
    End With
End Sub

'
'   取消筛选
'
Public Sub CancelFilter(wbs As Worksheet)
    With wbs
        If .AutoFilterMode Then
            If .FilterMode = True Then
                .ShowAllData
            End If
        End If
    End With
End Sub



'
'   复制筛选结果到某区域
'
Public Sub CopyFilterResult(filter_wbs As Worksheet, des_range As Range)
    With filter_wbs
        If .FilterMode Then '如果处于筛选模式
            des_range.Clear
            .AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy des_range
        End If
    End With
    Application.CutCopyMode = False
End Sub

'
'   筛选
'

'
'   获取筛选结果条数
'   .AutoFilter.Range.SpecialCells(xlCellTypeVisible) -> 筛选以后的分区，可能是多个组合，如 $A$1:$C$7,$A$15:$C$17
'   0 则表示无相应结果
'
Public Function GetFilterRecordCount(filter_wbs As Worksheet) As Long
    Dim x As Range, i As Integer
    Dim rdCnt As Long
    
    rdCnt = 0
    With filter_wbs
        If .FilterMode Then
            Set x = .AutoFilter.Range.SpecialCells(xlCellTypeVisible)
            '遍历所有分区
            For i = 1 To x.Areas.count
                'Range.rows.count -> 获取的是该区域的行数
                rdCnt = rdCnt + x.Areas(i).rows.count
            Next
            Set x = Nothing
            rdCnt = rdCnt - 1
        End If
    End With
    GetFilterRecordCount = rdCnt
End Function


'
'   判断筛选出来结果是否为空
'
Public Function FilterIsEmpty(filter_wbs As Worksheet) As Boolean
    Dim LastRow As Long, FirstColumn As Long
    With filter_wbs.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
        If .Areas.count = 1 And .rows.count = 1 Then
            FilterIsEmpty = True
        Else
            FilterIsEmpty = False
        End If
    End With
End Function

'
'   查找是Key的所有单元格,以range返回
'
Public Function FindAllCellsByKey(rn As Range, key As String) As Range
    Dim rng As Range, rng1 As Range, rng2 As Range
    
    Set rng = rn.find(key, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        Set rng1 = rng
        Set rng2 = rng
        
        Do
            Set rng2 = rn.FindNext(rng2)
            If rng2.address = rng1.address Then
                '相等，就退出 loop
                Exit Do
            Else
                Set rng = Union(rng, rng2)
            End If
        Loop
    End If
    
    Set FindAllCellsByKey = rng
End Function

'
'   获取单元格区域非空单元格个数
'
Public Function GetNotEmptyCellCount(rn As Range) As Double
    Dim iTotal As Double

    '获取区域中非空单元格个数
    iTotal = WorksheetFunction.CountIf(rn, "<>""""") - WorksheetFunction.CountIf(rn, "")
    GetNotEmptyCellCount = iTotal
End Function

'
'   获取某个区域不重复值，以字符串形式反馈
'
'
Public Function GetUniqueValueOfRanges(rns As Range) As String
    Dim rn As Range, d
    Dim str As String
    str = ""
    
    Set d = CreateObject("scripting.dictionary")
    For Each rn In rns
        If rn <> "" And Not d.exists(rn.value) Then
            If str = "" Then
                str = rn.value
            Else
                str = str + "," + rn.value
            End If
            d(rn.value) = rn.value
        End If
    Next
    d.RemoveAll
    GetUniqueValueOfRanges = str
End Function

'   判断工作表是否为空
'
Public Function IsBlankSheet(wbs As Worksheet) As Boolean
   If Application.WorksheetFunction.CountA(wbs.UsedRange.Cells) = 0 Then
       IsBlankSheet = True
   Else
       IsBlankSheet = False
   End If
End Function

'
'   替换工作博中的所有相关字符
'   loa xlPart or xlWhole
'
Public Function ReplaceSomething(wb As Workbook, find As String, replace, loa As Long)
    Dim i As Integer
    Dim rn As Range
    For i = 1 To wb.Worksheets.count Step 1
        Call wb.Worksheets(i).UsedRange.replace(find, replace, LookAt:=loa, searchOrder:=xlByRows)
    Next i
End Function

Function GetTitleColumn(wbs As Worksheet, key As String) As Integer
    Dim rn As Range, c As Range, r As Integer
    Set rn = wbs.rows(1)
    r = 0
    With rn
        Set c = .find(key, LookIn:=xlValues, LookAt:=xlWhole)
        If Not c Is Nothing Then
             r = c.Column
        End If
    End With
    
    GetTitleColumn = r
End Function

Function GetTitleRangesStr(wbs As Worksheet, key As String) As String
    Dim rn As Range, c As Range, r As String
    Set rn = wbs.rows(1)
    r = "A:A"
    With rn
        Set c = .find(key, LookIn:=xlValues, LookAt:=xlWhole)
        If Not c Is Nothing Then
             r = c.EntireColumn.address
        End If
    End With
    
    GetTitleRangesStr = r
End Function

Sub SetColumnWidth(ws As Worksheet, title As String, width As Integer)
    
End Sub

Sub DeleteColumn(ws As Worksheet, key As String)
    Dim r As Integer
    
    r = GetTitleColumn(ws, key)
    
    If r > 0 Then ws.columns(r).Delete
End Sub

