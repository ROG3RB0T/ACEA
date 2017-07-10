VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form FrmMttoUsuario 
   BackColor       =   &H00D8E9EC&
   Caption         =   "MANTENIMIENTO USUARIOS"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   9105
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11456
      BorderStyle     =   6
      BorderWidth     =   3
      BorderType      =   12
      Caption         =   ""
      BorderMarginTop =   3
      BackColor       =   16250871
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   12632256
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Usuario"
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8535
         Begin Proyecto1.ButtonOffice Adduser 
            Height          =   495
            Left            =   3600
            TabIndex        =   10
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BackColor       =   12230304
            Caption         =   "Agregar Nuevo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox TxPwdUser 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   4920
            MaxLength       =   8
            PasswordChar    =   "?"
            TabIndex        =   8
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox TxNombreUser 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            MaxLength       =   140
            TabIndex        =   7
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox TxIdUser 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            MaxLength       =   7
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin LabelDegradado.LabelDegrade LabelDegrade5 
            Height          =   285
            Left            =   3480
            TabIndex        =   3
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Nombre Usuario"
            BackColor       =   255
            BorderColor     =   9655840
            Transparente    =   0   'False
            ShadowDepth     =   0
            ShadowStyle     =   0
            ShadowColorStart=   0
            DegradadoOrientacion=   2
            DegradadoColorStart=   13993792
            DegradadoColorEnd=   12218153
         End
         Begin LabelDegradado.LabelDegrade LabelDegrade6 
            Height          =   285
            Left            =   3480
            TabIndex        =   4
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "ID Usuario"
            BackColor       =   255
            BorderColor     =   9655840
            Transparente    =   0   'False
            ShadowDepth     =   0
            ShadowStyle     =   0
            ShadowColorStart=   0
            DegradadoOrientacion=   2
            DegradadoColorStart=   13993792
            DegradadoColorEnd=   12218153
         End
         Begin LabelDegradado.LabelDegrade LabelDegrade1 
            Height          =   285
            Left            =   3480
            TabIndex        =   5
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Password"
            BackColor       =   255
            BorderColor     =   9655840
            Transparente    =   0   'False
            ShadowDepth     =   0
            ShadowStyle     =   0
            ShadowColorStart=   0
            DegradadoOrientacion=   2
            DegradadoColorStart=   13993792
            DegradadoColorEnd=   12218153
         End
         Begin Proyecto1.ButtonOffice EditUser 
            Height          =   495
            Left            =   5040
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BackColor       =   12230304
            Caption         =   "Editar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
         Begin Proyecto1.ButtonOffice DeleteUser 
            Height          =   495
            Left            =   6480
            TabIndex        =   12
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BackColor       =   12230304
            Caption         =   "Eliminar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicOpacity      =   0
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmMttoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UsuarioSel As String
Dim Transac As String
Sub CargarUsuarios()
    Dim Strqry As String
    Set Rs = New ADODB.Recordset
    Rs.CursorType = 1
    
    Strqry = "SELECT Usuarios.*" & _
            " FROM Usuarios " & _
            " WHERE Usuarios.IdUsuario <> '" & USUARIOACTIVO & "' AND Usuarios.IdUsuario <> 'ADMIN'"
    Rs.Open Strqry, CN
    List1.Clear
    
    Do While Not Rs.EOF
        List1.AddItem (Rs(0))
        Rs.MoveNext
    Loop
    Rs.Close
End Sub
Sub CargarPermisos(usuario As String)
    Dim RsUser As Recordset
    Dim RsPermisos As Recordset
    Set RsUser = New Recordset
    Set RsPermisos = New Recordset
    RsUser.CursorType = 1
    RsPermisos.CursorType = 1
    
    FX.LoadRstFromDB "QryUsuarios", RsUser, usuario
    
    TxIdUser.Text = RsUser("Idusuario").Value
    TxNombreUser.Text = RsUser("nombre").Value
    TxPwdUser.Text = FX.Base64Decode(RsUser("password").Value)
    
    FX.LoadRstFromDB "QryPermisosUsuarios", RsPermisos, usuario
    If RsPermisos.RecordCount > 0 Then
        ListView1.ListItems.Clear
        FX.LoadListView RsPermisos, ListView1
        
    End If
    
    
End Sub

Private Sub Adduser_Click()
    If Adduser.Caption = "Agregar Nuevo" Then
            EditUser.Caption = "Cancelar"
            Adduser.Caption = "Aceptar"
            TxIdUser.Locked = False
            TxNombreUser.Locked = False
            TxPwdUser.Locked = False
            ListView1.Enabled = False
            TxNombreUser.Text = ""
            TxIdUser.Text = ""
            TxPwdUser.Text = ""
            ListView1.ListItems.Clear
            TxNombreUser.SetFocus
            Transac = "Add"
            List1.Enabled = False
            DeleteUser.Enabled = False
    Else
        If Transac = "Add" Then
            Dim Txok As Boolean
            Dim Params As Variant
            Dim pwduser As String
            Txok = False
            If TxNombreUser.Text < "" Then
                MsgBox "Debe ingresar un nombre de usuario"
            ElseIf TxIdUser.Text < "" Then
                MsgBox "Debe ingresar un id de usuario"
            ElseIf TxPwdUser.Text < "" Then
                MsgBox "El password del usuario no puede ser vacio"
            Else
               pwduser = FX.Base64Encode(TxPwdUser.Text)
               Params = Array(UCase(TxIdUser.Text), TxNombreUser.Text, pwduser)
               Txok = FX.CmdTransacciones("AddUser", Params)
                If Txok Then
                    Txok = False
                    Txok = FX.CmdTransacciones("QryAddPermisosUser", TxIdUser.Text)
                    If Txok Then
                        MsgBox "Se ha adicionado el usuario satisfactoriamente"
                    End If
                End If
            End If
            EditUser.Caption = "Editar"
            Adduser.Caption = "Agregar Nuevo"
            TxIdUser.Locked = True
            TxNombreUser.Locked = True
            TxPwdUser.Locked = True
            ListView1.Enabled = False
            TxNombreUser.Text = ""
            TxIdUser.Text = ""
            TxPwdUser.Text = ""
            CargarUsuarios
        ElseIf Transac = "Edit" Then
            
            Txok = False
            pwduser = FX.Base64Encode(TxPwdUser.Text)
            Params = Array(TxNombreUser.Text, pwduser, TxIdUser.Text)
            
            Txok = FX.CmdTransacciones("EditUser", Params)
            If Txok Then
                Txok = False
                For i = 1 To ListView1.ListItems.Count
                    Params = Array(ListView1.ListItems(i).Checked, ListView1.ListItems(i).SubItems(1), Val(ListView1.ListItems(i).SubItems(2)), TxIdUser.Text)
                    Txok = FX.CmdTransacciones("EditPermisoUser", Params)
                    If Not Txok Then Exit For
                Next
                If Txok Then
                    MsgBox "Se ha modificado el usuario"
                    EditUser.Caption = "Editar"
                    Adduser.Caption = "Agregar Nuevo"
                    TxIdUser.Locked = True
                    TxNombreUser.Locked = True
                    TxPwdUser.Locked = True
                    ListView1.Enabled = False
                    TxNombreUser.Text = ""
                    TxIdUser.Text = ""
                    TxPwdUser.Text = ""
                    ListView1.ListItems.Clear
                    CargarUsuarios
                End If
            End If
                
        End If
        List1.Enabled = True
        DeleteUser.Enabled = True
        Transac = ""
    End If

End Sub

Private Sub DeleteUser_Click()
If TxIdUser.Text > "" Then
    If MsgBox("Desea eliminar el usuario Seleccionado?", vbInformation + vbYesNo) = vbYes Then
        Dim Txok As Boolean
        Txok = FX.CmdTransacciones("DeleteUser", TxIdUser.Text)
        If Txok Then
            MsgBox "Se ha eliminado el usuario"
            CargarUsuarios
            ListView1.ListItems.Clear
        End If
    End If
End If
End Sub

Private Sub EditUser_Click()
    If EditUser.Caption = "Editar" Then
            EditUser.Caption = "Cancelar"
            Adduser.Caption = "Aceptar"
            'TxIdUser.Locked = False
            TxNombreUser.Locked = False
            TxPwdUser.Locked = False
            ListView1.Enabled = True
            List1.Enabled = False
            DeleteUser.Enabled = False
            Transac = "Edit"
    Else
            EditUser.Caption = "Editar"
            Adduser.Caption = "Agregar Nuevo"
            TxIdUser.Locked = True
            TxNombreUser.Locked = True
            TxPwdUser.Locked = True
            ListView1.Enabled = False
            TxNombreUser.Text = ""
            TxIdUser.Text = ""
            TxPwdUser.Text = ""
            ListView1.ListItems.Clear
            List1.Enabled = True
            DeleteUser.Enabled = True
            Transac = ""
    End If

End Sub

Private Sub Form_Load()
FX.ConnectDb activar
CargarUsuarios
End Sub

Private Sub List1_DBLClick()
        CargarPermisos (List1.List(List1.ListIndex))
        UsuarioSel = List1.List(List1.ListIndex)
End Sub
