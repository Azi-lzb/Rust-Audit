VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyDictList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_title As Variant
Private m_datas As Object

'类初始化
Private Sub Class_Initialize()
   'ReDim arr_base_address(1 To 10) As String
End Sub

'类销毁
Private Sub Class_Terminate()
    Set m_datas = Nothing
End Sub
