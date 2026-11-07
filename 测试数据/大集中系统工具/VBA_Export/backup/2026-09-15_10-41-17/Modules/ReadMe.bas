Attribute VB_Name = "ReadMe"
Option Explicit

'
'
'   bugÖÆÔì»ú
'
'   0752-2380885 cyc
'
'

Sub test()
    Dim str As String
    str = "a|bc|dd|ee|ff"
    
    Dim arr As Variant
    arr = Split(str, "|", 3)
    Debug.Print UBound(arr) - LBound(arr) + 1
End Sub

