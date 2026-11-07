VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyConfigTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WS_CONFIG_NAME As String = "基本配置"

'Application.Calculation = xlCalculationManual    '手动计算
'Application.Calculation = xlCalculationAutomatic '自动计算

'
'  激活状态栏
'
Public Sub ShowStatusMessage_Up(msg As String)
    Dim t1 As Boolean, t2 As Boolean, t3 As Boolean
    t1 = Application.ScreenUpdating
    t2 = Application.AskToUpdateLinks
    t3 = Application.DisplayAlerts
    
    Application.ScreenUpdating = True  '屏幕刷新
    Application.AskToUpdateLinks = True    '询问是否刷新
    Application.DisplayAlerts = True   '是否警告
    
    Application.StatusBar = msg
    DoEvents
    Application.ScreenUpdating = t1  '屏幕刷新
    Application.AskToUpdateLinks = t2    '询问是否刷新
    Application.DisplayAlerts = t3   '是否警告

End Sub


Public Sub ChangeSomethingStatus(ByRef abOldStatus() As Boolean, newSatus As Boolean)
    abOldStatus(1) = Application.ScreenUpdating
    abOldStatus(2) = Application.AskToUpdateLinks
    abOldStatus(3) = Application.DisplayAlerts
    Application.ScreenUpdating = newSatus  '屏幕刷新
    Application.AskToUpdateLinks = newSatus    '询问是否刷新
    Application.DisplayAlerts = newSatus   '是否警告
End Sub

Public Sub RestoreSomethingStatus(ByRef abOldStatus() As Boolean)
    Application.ScreenUpdating = abOldStatus(1)
    Application.AskToUpdateLinks = abOldStatus(2)
    Application.DisplayAlerts = abOldStatus(3)
End Sub

'
'   1.关闭某些提示
'
Public Sub CloseSomething()
    Application.ScreenUpdating = False  '屏幕刷新
    Application.AskToUpdateLinks = False    '询问是否刷新
    Application.DisplayAlerts = False   '是否警告
End Sub



'
'   2.开启某些提示
'
Public Sub OpenSomething()
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
End Sub


'
'   根据关键字,获取一些配置对应的值
'
Public Function GetConfig(key As String) As String
    Dim wsConfig As Worksheet
    
    Set wsConfig = ThisWorkbook.Worksheets(WS_CONFIG_NAME)
    
    Dim rn As Range, c As Range
    Set rn = wsConfig.UsedRange
    With rn
        Set c = .find(key, LookIn:=xlValues, LookAt:=xlWhole)
        If c Is Nothing Then
            '看看会不会是数组参数
            Set c = .find("【" & key, LookIn:=xlValues, LookAt:=xlWhole)
            If c Is Nothing Then
                err.Raise Number:=1001, Description:="不存在" & key & "相关配置"
            Else
                Dim count As Integer, i As Integer, str As String
                count = CInt(wsConfig.Cells(c.row, c.Column + 1).Value)
                For i = 1 To count Step 1
                    If i = 1 Then
                        str = wsConfig.Cells(c.row, c.Column + 1 + i).Value
                    Else
                        str = str & "," & wsConfig.Cells(c.row, c.Column + 1 + i).Value
                    End If
                Next i
                GetConfig = str
            End If
        Else
            GetConfig = wsConfig.Cells(c.row, c.Column + 1).Value
        End If
    End With
End Function

Public Sub SetConfig(key As String, val As String)
    Dim wsConfig As Worksheet, row As Long
    
    Set wsConfig = ThisWorkbook.Worksheets(WS_CONFIG_NAME)
    
    Dim rn As Range, c As Range
    Set rn = wsConfig.UsedRange
    With rn
        Set c = .find(key, LookIn:=xlValues, LookAt:=xlWhole)
        If c Is Nothing Then
            With wsConfig
                row = .Range("A65536").End(xlUp).row + 1
                .Cells(row, 1).Value = key
                .Cells(row, 2).Value = val
            End With
        Else
            c.Offset(0, 1).Value = val
        End If
    End With
End Sub


