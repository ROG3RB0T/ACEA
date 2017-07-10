VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_ACEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cod_Socio As String
Dim Nombres As String
Dim Apellidos As String
Dim MontoAporte As Double
Dim MontoAhorro As Double


Private Sub Class_Initialize()
    MontoAporte = 0
    MontoAhorro = 0
End Sub
Public Function AddSocio( _
    New_Cod_Socio As String, _
    New_Nombres As String, _
    New_Apellidos As String, _
    New_MontoAporte As dou, _
    New_Ahorro As String) As Boolean
    

                            


End Function

