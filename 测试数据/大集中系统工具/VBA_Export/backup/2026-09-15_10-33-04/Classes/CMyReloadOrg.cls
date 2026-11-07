VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyReloadOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsOrgName As String
Public clsOrgCode As String
Public clsAreaName As String
Public clsAreaCode As String
Public clsDate As Date
Public clsDicExportData As Object

'类初始化
Private Sub Class_Initialize()
    Set clsDicExportData = CreateObject("scripting.dictionary")
End Sub

'类销毁
Private Sub Class_Terminate()
    Set clsDicExportData = Nothing
End Sub

Public Sub clsAddExportData(arr() As Variant, id As Long)
    
End Sub


Public Function clsIsThisOrg(arr() As Variant, id As Long) As Boolean
    
End Function

