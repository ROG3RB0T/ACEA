VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo   : ACEA.Funciones
'---------------------------------------------------------------------------------------
' Creado   : 13/10/2006/22:41
' Autor    : ROGER. All Right reserved
' Propósito: Se almacenan todas las funciones que se utilizan para la aplicacion
'---------------------------------------------------------------------------------------
Enum TipoQuery
    DBQuery = 0
    DbTable
    StringQry
End Enum

Public Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    If Err = 1 Then
        MsgBox "Base64Decode Bad Base64 string."
    End If
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)
    
      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Public Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim sOut, i
  
  'For each group of 3 bytes
  For i = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, i, 1)) + _
      &H100 * MyASC(Mid(inData, i + 1, 1)) + MyASC(Mid(inData, i + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Public Sub CmdFillCombos(NameComboQuery As String, ObjCombo As ComboBox, Parametros, _
                        Optional LoadFirst As Boolean = True, _
                        Optional LoadLabel As Boolean = True, _
                        Optional Loadindex As Boolean = False)
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.Funciones.CmdFillCombos
' Tipo       : Sub
' Autor      : ROGER
' Descripción: Función que permite llenar un combobox con los datos que se encuentran
'               en una tabla en la Base de Datos
'---------------------------------------------------------------------------------------
' Declaración de variables y objetos necesarios
    Dim RsCmb As Recordset
'On Error Resume Next

    Set RsCmb = New Recordset
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
       
    RsCmb.CursorType = 1
    
    'Configuración del objeto command
    Set Cmd.ActiveConnection = CN
    Cmd.CommandText = NameComboQuery
    Cmd.CommandType = adCmdStoredProc
    
    If Parametros = "" Then 'Si la consulta no necesita parametros
        Set RsCmb = Cmd.Execute
    Else 'De lo contrario
        Set RsCmb = Cmd.Execute(, Parametros)
    End If
    'Limpiamos el contro combobox
    ObjCombo.Clear
    Do While Not RsCmb.EOF
        'Comeinza el barrido de la informacion
        If LoadLabel Then
            ObjCombo.AddItem (RsCmb(1) & " [" & RsCmb(0).Value & "]") 'Agregamos los items encontrados en el Query
        Else
            ObjCombo.AddItem (RsCmb(1)) 'Agregamos los items encontrados en el Query
        End If
        If Loadindex Then
            ObjCombo.ItemData(ObjCombo.NewIndex) = CDbl(RsCmb.Fields.item(0).DefinedSize)
        End If
        RsCmb.MoveNext
        
    Loop
    
    If LoadFirst Then
        RsCmb.MoveFirst
        'Dejamos como selección predeterminada el primer registro
        ObjCombo.ListIndex = 0
    End If
    
    RsCmb.Close
    Set RsCmb = Nothing
    Set Cmd = Nothing
    

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.Funciones.CmdFillCombos." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If

End Sub

Public Function CmdTransacciones(NombreProcedimiento As String, Parametros) As Boolean
On Error Resume Next
        Dim Cmd As ADODB.Command
        Set Cmd = New ADODB.Command
                
                
                Cmd.ActiveConnection = CN
                Cmd.CommandText = NombreProcedimiento
                Cmd.CommandType = adCmdStoredProc
                Cmd.Execute , Parametros
                
                
            If Err.Number <> 0 Then
                    MsgBox "Error en CMDTRANSACCIONES -  " & NombreProcedimiento & ". Error: [" & Err.Number & "] " & Err.Description
                    CmdTransacciones = False
            Else
                    CmdTransacciones = True
            End If
        Set Cmd = Nothing
End Function

Public Function ConnectDb(Tipo As Conexion) As Boolean
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.Funciones.ConnectDb
' Tipo       : Function
' Autor      : ROGER
' Descripción: Permite establecer la conexion a la base de datos
'---------------------------------------------------------------------------------------
On Error Resume Next
    STRCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBFILE & _
                ";Jet OLEDB:Database Password = " & DBPWD
    Select Case Tipo
            Case 1
                Set CN = New ADODB.Connection
                CN.CursorLocation = adUseClient
                CN.Open (STRCONN)
            Case 2
                If CN.State = adStateOpen Then
                    CN.Close
                    Set CN = Nothing
                    ConnectDb = False
                End If
    End Select
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.Funciones.ConnectDb." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If

End Function

 
Public Sub LoadListView(ByRef RST As ADODB.Recordset, ByRef ListView As ListView, Optional XSorted As Boolean = True)
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.Funciones.Recordset_a_ListView
' Tipo       : Sub
' Autor      : ROGER
' Descripción: Carga de manera dinámica datos de un recordset a un control
'               Listview
'---------------------------------------------------------------------------------------
Dim i As Integer, k As Integer, Alineacion() As Integer, iColumna As Integer, vt() As Long
Dim st() As Single, w() As Long, t As Long, LVitem As ListItem
    
On Error Resume Next
    
    ListView.ListItems.Clear
 
    If RST.State = adStateOpen Then
       
       If Not (RST.BOF And RST.EOF) Then
          ListView.View = lvwReport
            iColumna = RST.Fields.Count - 1
            ReDim w(0 To iColumna)
            ReDim st(0 To iColumna)
            ReDim vt(0 To iColumna)
            ReDim Alineacion(0 To iColumna)
 
            For i = 0 To iColumna
                If RST(i).DefinedSize > 9 Then
                    w(i) = RST(i).DefinedSize
                Else
                    w(i) = 10
                End If
                t = t + w(i)
            Next
            For i = 0 To iColumna
                st(i) = w(i) / t
                vt(i) = RST.Fields(i).Type
                Select Case RST.Fields(i).Type
                        Case adCurrency, adDate, adDecimal, adDouble, _
                            adInteger, adNumeric, adSingle, adSmallInt, adVarNumeric
                            
                            Alineacion(i) = lvwColumnRight
                        Case adBSTR, adChar, adVarChar, adVariant, 202, 203
                            
                            Alineacion(i) = lvwColumnLeft
                            'If i = 0 Then Alineacion(i) = lvwColumnRight
                        Case adBoolean
                            Alineacion(i) = lvwColumnCenter
                        Case Else
                            Alineacion(i) = -1
                End Select
            Next
                    
            If ListView.ColumnHeaders.Count = 0 Then
            For i = 0 To iColumna
                If i = 0 Then Alineacion(i) = lvwColumnLeft
               ListView.ColumnHeaders.Add , , RST.Fields(i).Name, ListView.Width * st(i), Alineacion(i)
            Next
            
            ElseIf ListView.ColumnHeaders.Count > 0 Then
                            
            Else
                For i = 0 To iColumna
                    ListView.ColumnHeaders(i + 1).Width = ListView.Width * st(i)
                Next
            End If
            RST.MoveFirst
 
            While Not RST.EOF
            
               If vt(0) = adBoolean Then
                If ListView.Checkboxes = True Then
                  If RST.Fields(0).Value = vbFalse Then
                     Set LVitem = ListView.ListItems.Add()
                     LVitem.Checked = False
                  Else
                     Set LVitem = ListView.ListItems.Add()
                     LVitem.Checked = True
                  End If
                Else
                  If RST.Fields(0).Value = vbFalse Then
                     Set LVitem = ListView.ListItems.Add(, , "NO")
                  Else
                     Set LVitem = ListView.ListItems.Add(, , "SI")
                  End If
                End If
               ElseIf vt(0) = adDouble Then 'cuando sea numerico currency se le asigna formato
                     Set LVitem = ListView.ListItems.Add(, , Format(RST.Fields(0).Value, "$#.##"))
               ElseIf vt(0) = adDate Then ' cuando sea de tipo fecha se cambia el formato
                    Set LVitem = ListView.ListItems.Add(, , Format(RST.Fields(0).Value, "dd/MMM/YYYY"))
               Else
                     Set LVitem = ListView.ListItems.Add(, , RST.Fields(0).Value)
               End If
 
               If iColumna > 0 Then
 
               For k = 1 To iColumna
                   If vt(k) = adBoolean Then
                      If RST.Fields(k).Value = vbFalse Then
                         LVitem.ListSubItems.Add , , "NO"
                       Else
                         LVitem.ListSubItems.Add , , "SI"
                       End If
                   Else
                       If IsNull(RST.Fields(k).Value) Then
                          LVitem.ListSubItems.Add , , ""
                       ElseIf vt(k) = adDouble Then
                        LVitem.ListSubItems.Add , , Format(RST.Fields(k).Value, "$##0.00")
                       ElseIf vt(k) = adDate Then
                        LVitem.ListSubItems.Add , , Format(RST.Fields(k).Value, "dd/MMM/YYYY")
                       Else
                          LVitem.ListSubItems.Add , , RST.Fields(k).Value
                       End If
                   End If
               Next
               End If
               
               RST.MoveNext
            Wend
            
            ListView.Sorted = XSorted
            'If XSorted Then ListView.SortKey = 0
            
        End If
        End If
 
       
        Set RST = Nothing
    
    

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.Funciones.LoadListView." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
    
    Set RST = Nothing
End If

End Sub

Sub LoadRstFromDB(Storedproc As String, ByRef rs As ADODB.Recordset, ByRef Parametros, Optional Tipo As TipoQuery = 0)
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.Funciones.LoadRstFromDB
' Tipo       : Sub
' Autor      : ROGER
' Descripción:
'---------------------------------------------------------------------------------------

    Dim Cmd As ADODB.Command
On Error Resume Next

    Set Cmd = New ADODB.Command


    Cmd.ActiveConnection = CN
    Cmd.CommandText = Storedproc
    If Tipo = DbTable Then
        Cmd.CommandType = adCmdTable
    ElseIf Tipo = DBQuery Then
        Cmd.CommandType = adCmdStoredProc
    ElseIf Tipo = StringQry Then
        Cmd.CommandType = adCmdText
    End If
    
    'If Parametros = "" Then 'Si la consulta no necesita parametros
    '    Set Rs = Cmd.Execute(, Parametros)
    'Else 'De lo contrario
        
        '//17/09/2006
        'se elimina la evaluaciòn de la variable parametros
        'ya que no pasaba los datos provenientes de una variable tipo array
        'y se estableciò al prinicipio del procedimiento como valor por referencia(ByRef)
        'De esta manera aunque la variable ByRef de Parametros no contenga valores siempre
        'se ejecuta el comando y se pasa el resultado al Recordset
            
            Set rs = Cmd.Execute(, Parametros)
            
    'End If
    Set Cmd = Nothing
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.Funciones.LoadRstFromDB." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
End Sub
Private Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function
Public Function LeeINI(Archivo As String, Seccion As String, Clave As String)
    Dim iRetLen As Integer
    Dim sRet As String
    sRet = Space(255)
    iRetLen = GetPrivateProfileString(Seccion, Clave, "", sRet, Len(sRet), Archivo)
    sRet = Left(sRet, iRetLen)
    LeeINI = sRet
End Function

Public Sub GrabaINI(Archivo As String, Seccion As String, Clave As String, Text As String)
    WritePrivateProfileString Seccion, Clave, Text, Archivo
End Sub
Public Sub GetSectionsIni(File As String, ByRef Lv As ListView)
    Dim Secciones As String, NumCar As Long, i As Long, j As Long
    Secciones = Space$(2000)
    NumCar = GetPrivateProfileSectionNames(Secciones, 2000, File)
    i = InStr(Secciones, Chr$(0))
    j = 1
    While i > 0 And i < NumCar
        Lv.ListItems.Add , , Mid(Secciones, j, i - j)
        j = i + 1
        i = InStr(j, Secciones, Chr$(0))
    Wend
End Sub
Public Sub LockTxMaxText(obj As TextBox, Optional MaxText As Integer = 0)
    Dim NumChars As Integer
    If MaxText > 0 Then
        NumChars = Left(obj.Text, MaxText)
        obj.Text = NumChars
    End If
End Sub


Public Sub CargarPermisos(usuario As String)
On Error Resume Next
        Dim Strqry As String
        Dim Rspermiso As New Recordset
        FX.ConnectDb activar
        Strqry = "SELECT * From PermisosDb Where IdUsuario = '" & usuario & "'"
        
        Rspermiso.CursorType = 1
    
        Rspermiso.Open Strqry, CN
        If Rspermiso.RecordCount > 0 Then
            Do While Not Rspermiso.EOF
                Select Case Rspermiso("CaptionMnu").Value
                        Case Is = "MnuTools"
                            FrmMain.MnuTools(Rspermiso("Index").Value).Enabled = Rspermiso("Enable").Value
                        Case Is = "MnuSocios"
                            FrmMain.MnuSocios(Rspermiso("Index").Value).Enabled = Rspermiso("Enable").Value
                        Case Is = "MnuDtos"
                            FrmMain.MnuDtos(Rspermiso("Index").Value).Enabled = Rspermiso("Enable").Value
                        Case Is = "MnuPtamos"
                            FrmMain.MnuPtamos(Rspermiso("Index").Value).Enabled = Rspermiso("Enable").Value
                End Select
                Rspermiso.MoveNext
            Loop
                FrmMain.MnuchgPWD.Enabled = True
        Else
                MsgBox "No se pudieron cargar los permisos de usuario, Contacte a su Administrador de Sistema", vbInformation, "Mensaje"
                End
        End If
        'Rs.Close
End Sub
Public Function Login(usuario As String, pwd As String) As Boolean
On Error Resume Next
    Dim encodepwd As String
    Dim Rsload As New Recordset
    FX.ConnectDb activar
    encodepwd = FX.Base64Encode(pwd)
    Rsload.CursorType = 1
    
    LoadRstFromDB "Qryloaduser", Rsload, usuario
    
    If Rsload.RecordCount > 0 Then
            If encodepwd = Rsload("password") Then
                Login = True
                USUARIOACTIVO = usuario
            Else
                Login = False
            End If
    End If
If Err.Number <> 0 Then
        Login = False
End If
End Function
Function Exportar_Excel(Path_Libro As String, _
                        ListView As ListView, _
                        Optional Progressbar As Progressbar) As Boolean

If ListView.ListItems.Count > 0 Then
'CREAR EL OBJETO (INSTANCIAR)CON EL OBJETO APLICACION (obj_Excel)
Dim obj_Excel As Object
Dim obj_Libro As Object
Dim IrowIndex As Integer
Dim icolindex As Integer

Progressbar.Visible = True
Progressbar.Value = 0
    ' Nueva referencia a Excel y nuevo referencia al Libro
    Set obj_Excel = CreateObject("Excel.Application")
    'Set obj_Libro = obj_Excel.Workbooks.Add
    
    'Variables para las columnas y filas
    Dim Col As Integer, Fila As Integer
    
    With obj_Excel
        .Visible = False
        'crea el libro
        Set obj_Libro = .Workbooks.Add
        
    End With

    With obj_Libro
    
        'Asignamos El valor Maximo del Progress teniendo _
         como dato la cantidad de items en el ListView
        If Not Progressbar Is Nothing Then
            Progressbar.Max = ListView.ListItems.Count
        End If
    
        'Referencia a la hoja con índice 1
        With .Sheets(1)
            IrowIndex = 1
            For icolindex = 1 To ListView.ColumnHeaders.Count
                With obj_Excel.Cells(IrowIndex, icolindex)
                    .Value = ListView.ColumnHeaders.item(icolindex).Text
                    With .Font
                        .Name = "Times New Roman"      'Make the headers stand out
                        .Size = 12
                        .Bold = True
                    End With
                End With
            Next icolindex
            'Recorremos la cantidad de items del ListView
            For Fila = 1 To ListView.ListItems.Count
                Col = 1
                'Asignamos EL Item actual en la celda
                .Cells(Fila + 1, Col) = ListView.ListItems.item(Fila)
                
                'Asignamos EL SubitemItem actual en la celda
                 For Col = 1 To ListView.ColumnHeaders.Count - 1
                     .Cells(Fila + 1, Col + 1) = _
                      ListView.ListItems(Fila).SubItems(Col)
                 Next

                 If Not Progressbar Is Nothing Then
                     'Aumentamos en 1 la propiedad value
                     Progressbar.Value = Progressbar.Value + 1
                 End If
            Next
        End With
    End With

    'Destruimos las variables de objeto
    MsgBox "Se ha finalizado la exportación a Excel"
    obj_Excel.Visible = True
    Set obj_Libro = Nothing
    Set obj_Excel = Nothing
    'Ok
    Exportar_Excel = True
    
    If Not Progressbar Is Nothing Then
       Progressbar.Value = 0
       Progressbar.Visible = False
    End If
End If
Exit Function

errSub:

Exportar_Excel = False

MsgBox Err.Description, vbCritical

On Error Resume Next
Set obj_Libro = Nothing
Set obj_Excel = Nothing

Progressbar.Value = 0
End Function
Sub Autocompletar(ListView As ListView, TBox As TextBox)
    
    ' variable para usar con el método FindItem que _
      permite buscar en el LV
    Dim item As ListItem
    Dim seleccion As Integer
    
    ' busca en el item, la cadena escrita en el textbox, si coincide _
      devuelve una referencia al item
    Set item = ListView.FindItem(TBox.Text, 0, , 1)
        
        ' verifica que el item no sea un valor nothing
        If Not item Is Nothing Then
            ' Muestra la selección pormas que no tenga el foco
            ListView.HideSelection = False
            
            ' desplaza la lista
            item.EnsureVisible
            
            ' selecciona el item
            item.Selected = True

            If Not FLAG Then
                ' Almacena la posición de la selección en el textbox
                seleccion = TBox.SelStart
                
                ' Asigna el texto completo del item encontrado
                TBox.Text = CStr(item)
                    If Not TBox.Text = vbNullString Then
                        ' posición de la selección
                        TBox.SelStart = seleccion
                        ' selecciona el texto
                        TBox.SelLength = Len(TBox.Text) - seleccion
                    End If
            End If
        Else
            ' Oculta la selección ya que no hay coincidencia
            ListView.HideSelection = True
        End If

End Sub

