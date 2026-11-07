Attribute VB_Name = "temp"
Sub test()
    'Ôö¼ÓÅú×¢
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet7")
    
    Dim arr As Variant
    arr = gloEt.GetUsedRange(ws)
    
    ws.Range("X1:X" & CStr(UBound(arr))).ClearContents
    
    Dim i As Long
    For i = 2 To UBound(arr)
        If CStr(arr(i, 23)) <> "" Then
            ws.Cells(i, "X").Value = gloEt.GetCommentStr(ws.Cells(i, 23))
        End If
    Next
    
End Sub

Sub test2()
    Dim reg As Object
    Dim s As String, p As String
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    
    p = "^f": s = "f012"
    With reg
        .Global = True: .ignorecase = False: .Pattern = p
    End With
    
    
    Debug.Print "reg:" & p & ",str:" & s & ":" & reg.test(s)
    
    
    p = "^(?![79fgij])": s = "6000"
    p = "^[^79fgij]": s = "ck0a"
    With reg
        .Global = True: .ignorecase = False: .Pattern = p
    End With
    
    
    Debug.Print "reg:" & p & ",str:" & s & ":" & reg.test(s)
End Sub
