VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyDebugTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAXCOUNT As Integer = 15
Private Const MAXROWS As Long = 30000
Private Const INSERTROW As Integer = 2  '在第二行插入
Private Const STEPROW As Integer = 15
Private wsDe As Worksheet

'类初始化
Private Sub Class_Initialize()
    
End Sub

'类销毁
Private Sub Class_Terminate()
    
End Sub


'
'   创建日记标题
'
'
Public Sub CreateDebugTitle(ParamArray strs() As Variant)
 
    Dim i As Integer
    Application.CutCopyMode = False
    
    Set wsDe = gloEt.FindWs(ThisWorkbook, "输出日记", 1, True)
    wsDe.Visible = xlSheetVisible
    
    With wsDe
        If .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row > MAXROWS Then .Cells.Clear
        
        .rows(1).Resize(STEPROW).Insert
        .rows(1).Interior.colorindex = 22
        .Cells(1, 1) = "日期"
    
        For i = 0 To UBound(strs) Step 1
            .Cells(1, i + 2) = strs(i)
        Next i
    End With
End Sub


'
'   添加一条记录，不限参数，位置调用的时候自己弄好
'
Public Sub InsertNewDebug(ParamArray strs() As Variant)
    Dim i As Integer, delRow As Long
    Application.CutCopyMode = False

    Set wsDe = gloEt.FindWs(ThisWorkbook, "输出日记", 1, True)
    
    With wsDe
        If .Range("A" & CStr(.Cells.rows.count)).End(xlUp).row > MAXROWS Then
            '导数两次的删掉
            delRow = .Range("A" & CStr(.Cells.rows.count)).End(xlUp).End(xlUp).End(xlUp).End(xlUp).row
            .Range(CStr(delRow) & ":" & CStr(.Cells.rows.count)).Delete
        End If
    
        .rows(INSERTROW).Insert
        .rows(INSERTROW).Interior.colorindex = 0
        .Cells(INSERTROW, 1) = Format(Date, "yy-m-d") & " " & Format(Time, "hh:mm:ss")
        For i = 0 To UBound(strs) Step 1
            .Cells(INSERTROW, i + 2) = strs(i)
        Next i
    End With
End Sub

'
'   给上一条记录添加颜色
'
Public Sub AddColorDebug(color As Integer, Optional c As Integer = 0)
    Set wsDe = gloEt.FindWs(ThisWorkbook, "输出日记", 1, True)
    
    With wsDe
        If c = 0 Then
            .rows(2).Interior.colorindex = color
        Else
            .Cells(2, c).Interior.colorindex = color
        End If
    End With
End Sub

Public Sub AddDeHypeLinks(col As Integer, path As String, showName As String)
    Set wsDe = gloEt.FindWs(ThisWorkbook, "输出日记", 1, True)
    
    With wsDe
       .Cells(2, col).Hyperlinks.add anchor:=.Cells(2, col), address:=path, SubAddress:="", ScreenTip:="", TextToDisplay:=showName
    End With
End Sub

Public Sub SetDeFoucs()
    Set wsDe = gloEt.FindWs(ThisWorkbook, "输出日记", 1, True)
    
    On Error GoTo FoucsErrHandle
    ThisWorkbook.Activate
    wsDe.Activate
    wsDe.Range("A1").Activate
Exit Sub

FoucsErrHandle:
    Call InsertNewDebug("设置焦点", "日志工作表焦点转移功能，不影响其他使用。" & err.Description)
    err.Clear
End Sub

