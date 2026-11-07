Attribute VB_Name = "测试模块"
Sub test()
    '迁移跨期
    Dim row As Integer
    Dim reg As Object, mc As Object, m As Object, arr As Variant, i As Integer
    Dim str As String, rule As String
    Dim ws As Worksheet, rn As Range, wsDes As Worksheet, out_row As Long
    Set ws = ThisWorkbook.Worksheets("test")
    
    Set wsDes = ThisWorkbook.Worksheets("test2")
    out_row = 2
    
    For row = 2 To 92
        rule = ws.Cells(row, 3).Value
        
        Set reg = CreateObject("VBSCRIPT.REGEXP")
        With reg
            .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
        End With
        Set mc = reg.Execute(rule)
        wsDes.Range(wsDes.Cells(out_row, 1), wsDes.Cells(out_row, 5)).Value = "-"
        wsDes.Cells(out_row, 6).Value = ws.Cells(row, 2).Value
        wsDes.Cells(out_row, 7).Value = ws.Cells(row, 3).Value
        out_row = out_row + 1
        For Each m In mc
            str = ""
            arr = Split(replace(replace(CStr(m.Value), "{", ""), "}", ""), ",")
            '合并 [业务类，机构类代码，地区代码，指标代码，数据属性，币种，频度，批次]>0
            str = "{" & arr(3) & "," & arr(4) & "," & arr(5) & "," & arr(6) & "," & arr(7) & "}"
             ws.Cells(row, 1).End(xlToRight).Offset(0, 1).Value = str
             wsDes.Cells(out_row, 1).Value = arr(3)
             wsDes.Cells(out_row, 2).Value = arr(4)
             wsDes.Cells(out_row, 3).Value = arr(5)
             wsDes.Cells(out_row, 4).Value = arr(6)
             wsDes.Cells(out_row, 5).Value = arr(7)
             out_row = out_row + 1
        Next
        
        With reg
            .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
        End With
        Set mc = reg.Execute(rule)
        For Each m In mc
            str = ""
            arr = Split(replace(replace(CStr(m.Value), "[", ""), "]", ""), ",")
            '合并 [业务类，机构类代码，地区代码，指标代码，数据属性，币种，频度，批次]>0
            str = "[" & arr(3) & "," & arr(4) & "," & arr(5) & "," & arr(6) & "," & arr(7) & "]"
             ws.Cells(row, 1).End(xlToRight).Offset(0, 1).Value = str
             wsDes.Cells(out_row, 1).Value = arr(3)
             wsDes.Cells(out_row, 2).Value = arr(4)
             wsDes.Cells(out_row, 3).Value = arr(5)
             wsDes.Cells(out_row, 4).Value = arr(6)
             wsDes.Cells(out_row, 5).Value = arr(7)
             out_row = out_row + 1
             wsDes.Cells(out_row, 1).Value = "-"
        Next
        out_row = out_row + 1
    Next
End Sub
