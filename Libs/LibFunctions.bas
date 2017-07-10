Attribute VB_Name = "LibFunctions"
'---------------------------------------------------------------------------------------
' Modulo   : ACEA.LibFunctions
'---------------------------------------------------------------------------------------
' Creado   : 11/10/2006/23:20
' Autor    : ROGER. All Right reserved
' Propósito: Contenedor de las librerias utilizadas en el Programa.
' Observ.  : Variables Públicas en Mayúsculas, Privadas a nivel de modulo
'            la primera en mayúscula y lo demas en minúscula. Locales de procedimiento
'            en minúsculas
'---------------------------------------------------------------------------------------

'Las siguientes declaraciones API son para el funcionamiento de los
'procedimientos para archivos INI

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
    As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal _
    nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
    lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias _
    "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize _
    As Long, ByVal lpFileName As String) As Long

'Variables publicas
Public INIFILE As String
Public DBFILE As String
Public CN As Connection
Public FX As New Funciones
Public USUARIOACTIVO As String
Public VAR_COD_SOCIO As String
Public VAR_CODSERV As String 'Codigo de Servicio de descuento
Public DBPWD As String
Public FLAG As Boolean
Public VAR_CAP As Double


Enum TipoTransaccion
    NoTransac = 0
    Agregarnuevo = 1
    EditarExistente
    EliminarRegistro
    CancelarTransac
End Enum

Enum Conexion 'Tipo de datos para identificar el tipo de procedimiento a utilizar para la conexion
        activar = 1
        Desactivar
End Enum
Sub main()
If App.PrevInstance = True Then
    MsgBox "La aplicación ya esta en ejecución"
Else
    LOADPARAMETROSSYS
    FrmMain.Show
End If
End Sub
Public Sub LOADPARAMETROSSYS()
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.LibFunctions.LOADPARAMETROSSYS
' Tipo       : Sub
' Autor      : ROGER
' Descripción: Procedimiento de tipo público para cargar los parametros necesarios para
'               la aplicación
'---------------------------------------------------------------------------------------

On Error Resume Next

    'Dim Base64_Function As New Funciones
     INIFILE = App.Path & "\Libs\Lib.dll"
        
    DBFILE = FX.LeeINI(INIFILE, "DATABASE", "PATCH")
'    DBFILE = App.Path & x.INITool1.GetFromINI("DATABASE", "PATCH", App.Path & IniFile)
    DBPWD = Trim(FX.LeeINI(INIFILE, "DATABASE", "PWD"))
    'DBPWD = x.INITool1.GetFromINI("DATABASE", "PWD", App.Path & IniFile)
    VAR_CAP = FX.LeeINI(INIFILE, "CAP", "Porcentaje")
    If IsEmpty(VAR_CAP) Then VAR_CAP = 0.02
    If Not IsEmpty(DBPWD) Then
        'DBPWD = Base64_Function.Base64Decode(DBPWD)
        DBPWD = FX.Base64Decode(DBPWD)
    End If

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & _
        "] en ACEA.LibFunctions.LoadParametros." & vbCrLf & _
        "Si el problema persiste contacte con su Administrador de Sistemas."
End If
End Sub




