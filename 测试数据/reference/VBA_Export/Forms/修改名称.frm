VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 修改名称 
   Caption         =   "修改名称"
   ClientHeight    =   3210
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5370
   OleObjectBlob   =   "修改名称.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "修改名称"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrCmd(1 To 1) As CMySelCmdb

Private Sub ChangeNameCmd_Click()
    Call gloCft.CloseSomething
    
    Dim strfiles As String, arrfns() As String, i As Integer, wb As Workbook, t As String
    
    strfiles = tb1.value
    arrfns = Split(strfiles, ",")
    For i = 0 To UBound(arrfns) Step 1
        Set wb = gloEt.GetWb(arrfns(i))
        With wb.Worksheets(1)
            t = wb.path & "\" & .Cells(1, 1).value & .Cells(2, 1).value & _
            .Cells(3, 1).value & right(arrfns(i), Len(arrfns(i)) - InStrRev(arrfns(i), ".") + 1)
        End With
        wb.Close SaveChanges:=False
        Name arrfns(i) As t
    Next i
    
    Call gloCft.OpenSomething
End Sub

Private Sub Userform_Initialize()
    Dim cmdt As CMySelCmdb
    
    
    Dim i As Integer
    For i = 1 To 1 Step 1
        Set cmdt = New CMySelCmdb
        Set cmdt.cmd = Me.Controls("cmd" & CStr(i))
        Set cmdt.frmT = Me
        Set arrCmd(i) = cmdt
    Next i
End Sub

