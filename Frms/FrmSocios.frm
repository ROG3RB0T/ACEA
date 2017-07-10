VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form FrmSocios 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Socios"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   2850
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      BorderStyle     =   4
      BorderType      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor1  =   14457180
      GradientColor2  =   15650739
      BackgroundType  =   1
      Begin pRmFrame.RmFrame CmdBtns 
         Height          =   420
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   741
         BorderStyle     =   4
         BorderWidth     =   0
         BorderType      =   12
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   16777215
         GradientColor2  =   14592118
         BackgroundType  =   2
         ShadowColor     =   14457180
         Picture         =   "FrmSocios.frx":0000
         PictureSize     =   99
         PictureWidth    =   16
         PictureHeight   =   28
         Begin Proyecto1.ButtonOffice BtnCancel 
            Height          =   345
            Left            =   500
            TabIndex        =   14
            ToolTipText     =   "Cancelar"
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":0462
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
            State           =   3
         End
         Begin Proyecto1.ButtonOffice Btnsave 
            Height          =   345
            Left            =   20
            TabIndex        =   15
            ToolTipText     =   "Guardar"
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":0BDC
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
            State           =   3
         End
         Begin Proyecto1.ButtonOffice Btnnuevo 
            Height          =   345
            Left            =   980
            TabIndex        =   16
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":1356
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice Btnedit 
            Height          =   345
            Left            =   1460
            TabIndex        =   17
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":1AD0
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice BtnDelete 
            Height          =   345
            Left            =   1940
            TabIndex        =   18
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":224A
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   345
            Left            =   2420
            TabIndex        =   19
            Top             =   50
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   609
            BackColor       =   14592118
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HandPointer     =   -1  'True
            PicNormal       =   "FrmSocios.frx":2C5C
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
      End
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
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
      Begin TxNumOcx.TxNum TxAportacion 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Text            =   "$0.00"
         Value           =   "$0.00"
         BorderColor     =   9655840
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   285
         Left            =   240
         TabIndex        =   7
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
         Text            =   "Codigo Socio"
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
      Begin VB.CheckBox ChkActivo 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Activo"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin TextBoxWinXP.TextboxXP TxNombres 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TextBoxWinXP.TextboxXP TxCod_Socio 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TextBoxWinXP.TextboxXP TxApellidos 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   720
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
         Text            =   "Nombres"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1080
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
         Text            =   "Apellidos"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1560
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
         Text            =   "Aportación"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1920
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
         Text            =   "Ahorros"
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
      Begin TxNumOcx.TxNum TxAhorro 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Text            =   "$0.00"
         Value           =   "$0.00"
         BorderColor     =   9655840
      End
   End
End
Attribute VB_Name = "FrmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cod_Socio As String
Dim Transac As TipoTransaccion
Sub Clean_Tx()
    TxCod_Socio.Text = ""
    TxNombres.Text = ""
    TxApellidos.Text = ""
    TxAportacion.Value = 0
    TxAhorro.Value = 0
    ChkActivo.Value = False
End Sub
Sub CargarSocio()
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    Strsql = "Select * from socios where Cod_Socio = '" & Cod_Socio & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxCod_Socio.Text = rssocio("Cod_Socio").Value
        TxNombres.Text = rssocio("Nombres").Value
        TxApellidos.Text = rssocio("Apellidos").Value
        TxAportacion.Value = rssocio("MontoAportacion").Value
        TxAhorro.Value = rssocio("MontoAhorros").Value
        If rssocio("Estado").Value Then
            ChkActivo.Value = 1
        Else
            ChkActivo.Value = 0
        End If
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub
Sub AddNvoSocio()
'---------------------------------------------------------------------------------------
' Modulo     : ACEA.FrmSocios.AddNvoSocio
' Tipo       : Sub
' Autor      : ROGER
' Descripción:
'---------------------------------------------------------------------------------------
    Dim Transaction As Boolean
    Dim Paramet
    
On Error Resume Next

    Paramet = Array(TxCod_Socio.Text, _
                    TxNombres.Text, _
                    TxApellidos.Text, _
                    Format(TxAhorro.Value, "#.00"), _
                    Format(TxAportacion.Value, "#.00"))
    
    Transaction = FX.CmdTransacciones("QryInsertSocios", Paramet)
    
    If Transaction Then
        MsgBox "Se ha ingresado el nuevo registro", vbInformation
    Else
        MsgBox "No se pudo ingresar el registro", vbInformation
    End If
    

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmSocios.AddNvoSocio." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
            Unload Me
End If
    
End Sub
Sub EditSocio()
    Dim Transaction As Boolean
    Dim Paramet
    
On Error Resume Next

    Paramet = Array(TxNombres.Text, _
                    TxApellidos.Text, _
                    Format(TxAhorro.Value, "#.00"), _
                    Format(TxAportacion.Value, "#.00"), _
                    ChkActivo.Value, _
                    TxCod_Socio.Text)
    
    Transaction = FX.CmdTransacciones("QryEditSocio", Paramet)
    
    If Transaction Then
        MsgBox "Se ha Editado el registro", vbInformation
    Else
        MsgBox "No se pudo Editado el registro", vbInformation
    End If
    

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmSocios.EditSocio." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
            Unload Me
End If
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnDelete_Click()
    If MsgBox("Esta seguro de eliminar este registro", vbYesNo + vbInformation) = vbYes Then
        Dim TranOk As Boolean
        TranOk = FX.CmdTransacciones("DeleteSocios", TxCod_Socio.Text)
        If TranOk Then
            MsgBox "Se ha eliminado del listado el socio " & TxCod_Socio.Text
            Unload Me
        End If
    End If
End Sub

Private Sub BtnEdit_Click()
    BtnCancel.Enabled = True
    BtnSave.Enabled = True
    BtnNuevo.Enabled = False
    BtnEdit.Enabled = False
    BtnDelete.Enabled = False
    RmFrame1.Enabled = True
    Transac = EditarExistente
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnNuevo_Click()
    BtnCancel.Enabled = True
    BtnSave.Enabled = True
    BtnNuevo.Enabled = False
    BtnEdit.Enabled = False
    BtnDelete.Enabled = False
    RmFrame1.Enabled = True
    Clean_Tx
    Transac = Agregarnuevo
End Sub

Private Sub BtnSave_Click()
    
    Select Case Transac
            Case 1 ' Agregarnuevo
                AddNvoSocio
                Unload Me
            Case 2
                EditSocio
    End Select
    BtnCancel.Enabled = False
    BtnSave.Enabled = False
    BtnNuevo.Enabled = True
    BtnEdit.Enabled = True
    BtnDelete.Enabled = True
    RmFrame1.Enabled = False
End Sub

Private Sub Form_Load()
    If Cod_Socio = "" Then
            BtnNuevo_Click
    Else
        CargarSocio
        RmFrame1.Enabled = False
    End If
End Sub

