VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub LeerINI()
On Error Resume Next
    'If FileExists(App.Path & "\Preferencias\config.ini") = True Then
        varCambios = False
        varCatRe = ReadINI("Categoría", "Reemplazar", "1")
    'Else
        'varCambios = False
        'varVerCat = ""
    'End If
End Sub

Sub GuardarINI()
On Error Resume Next
    WriteINI "Categoría", "Reemplazar", CStr(varCatRe)
    WriteINI "Categoría", "Mostrar", varVerCat
End Sub

Public Function ReadINI(strSection As String, strSetting As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strSetting, strDefault, strReturn, lngSize, App.Path & "\Preferencias\config.ini")
    If strReturn = "" Then
        ReadINI = strDefault
        WriteINI strSection, strSetting, strDefault
    Else
        ReadINI = LeftOf(strReturn, Chr(0))
    End If
End Function

Public Function WriteINI(strSection As String, strLValue As String, strRValue As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSection, strLValue, strRValue, App.Path & "\Preferencias\config.ini")
End Function

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function
