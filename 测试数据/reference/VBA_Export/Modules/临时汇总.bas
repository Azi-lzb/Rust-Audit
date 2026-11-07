Attribute VB_Name = "临时汇总"


Sub linshihz()
    Dim arr_src_file As Variant
    Dim wbDes As Workbook, wsdes As Worksheet
    
    Set wbDes = gloEt.GetWb("C:\Users\hzpbc.DESKTOP-JMSN70K\Desktop\2022-05-09 关于开展金融机构金融业综合统计标准执行情况评估\金融机构金融业综合统计标准执行情况统计表.xls")
    arr_src_file = gloEt.GetUsedRange(ThisWorkbook.Worksheets("执行面板"))
    
    Dim i As Long, wbsrc As Workbook, curRow As Long, arrTemp As Variant, j As Long, wsSrc As Worksheet, nRow As Long
    
    For Each wsdes In wbDes.Worksheets
        arrTemp = gloEt.GetUsedRange(wsdes)
        If wsdes.name = "总体情况" Then
            If UBound(arrTemp) >= 9 Then wsdes.Range("9:" & CStr(UBound(arrTemp))).Delete
        Else
            If UBound(arrTemp) >= 4 Then wsdes.Range("4:" & CStr(UBound(arrTemp))).Delete
        End If
    Next
    
    '打开一个文件
    For i = 6 To UBound(arr_src_file)
        Set wbsrc = gloEt.GetWb(CStr(arr_src_file(i, 2)))
        '遍历汇总表
        For Each wsdes In wbDes.Worksheets
            arrTemp = gloEt.GetUsedRange(wsdes)
            curRow = UBound(arrTemp) + 1
            If wsdes.name = "总体情况" Then
                Set wsSrc = wbsrc.Worksheets("A")
                nRow = 7
                wsdes.Cells(curRow, 1).value = ycLeft(wbsrc.name, "-", 1, False)
                For j = 2 To 61
                    If wsdes.Cells(8, j).value = "评估选项" Then
                        nRow = nRow + 1
                        wsdes.Cells(curRow, j).value = left(Trim(wsSrc.Cells(nRow, 5)), 1)
                    ElseIf wsdes.Cells(8, j).value = "次数" Then
                          nRow = nRow + 1
                          wsdes.Cells(curRow, j).value = wsSrc.Cells(nRow, 5)
                    Else
                        wsdes.Cells(curRow, j).value = wsSrc.Cells(nRow, 6)
                    End If
                Next
            Else
                Set wsSrc = wbsrc.Worksheets(wsdes.name)
                
                If Application.WorksheetFunction.CountA(wsSrc.Range("D3:D11")) <> 0 Then
                    wsdes.Cells(curRow, 1).value = ycLeft(wbsrc.name, "-", 1, False)
                    For j = 2 To 3
                        wsdes.Cells(curRow, j).value = left(Trim(wsSrc.Cells(j + 1, 4)), 1)
                    Next
                    For j = 4 To 6
                        If Application.WorksheetFunction.sum(wsSrc.Range("D5:D7")) < 2 Then
                            wsdes.Cells(curRow, j).value = wsSrc.Cells(j + 1, 4) * 100
                        Else
                            wsdes.Cells(curRow, j).value = wsSrc.Cells(j + 1, 4)
                        End If
                    Next
                    
                    For j = 7 To 10
                        If Application.WorksheetFunction.sum(wsSrc.Range("D8:D11")) < 2 Then
                            wsdes.Cells(curRow, j).value = wsSrc.Cells(j + 1, 4) * 100
                        Else
                            wsdes.Cells(curRow, j).value = wsSrc.Cells(j + 1, 4)
                        End If
                    Next
                    
                    If Application.WorksheetFunction.sum(wsdes.Range(wsdes.Cells(curRow, 4), wsdes.Cells(curRow, 6))) <> 100 Then
                        Debug.Print wbsrc.name & " " & wsSrc.name & "2 不平"
                    End If
                    
                    If Application.WorksheetFunction.sum(wsdes.Range(wsdes.Cells(curRow, 7), wsdes.Cells(curRow, 10))) <> 100 Then
                        Debug.Print wbsrc.name & " " & wsSrc.name & " 3 不平"
                    End If
                End If
            End If
        Next
        wbsrc.Close False
    Next
End Sub
