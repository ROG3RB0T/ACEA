VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDescPlan 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicación de Descuentos a Socios"
   ClientHeight    =   3660
   ClientLeft      =   765
   ClientTop       =   405
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
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
      Begin TxNumOcx.TxNum TxMonto 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   2400
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
      End
      Begin VB.ComboBox CmbDesc 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   120
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
      Begin TextBoxWinXP.TextboxXP TxNombres 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
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
         Locked          =   -1  'True
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
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
         Text            =   "Nombre del Socio"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
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
         Text            =   "Código Servicio"
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
      Begin TextBoxWinXP.TextboxXP TxNomDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1920
         Width           =   4455
         _ExtentX        =   7858
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
         Locked          =   -1  'True
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   120
         TabIndex        =   14
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
         Text            =   "Descripción"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2400
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
         Text            =   "Monto($)"
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
      Begin TextBoxWinXP.TextboxXP TxCod_Socio 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   120
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
      Begin MSComCtl2.DTPicker Fecinicio 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   39010
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade6 
         Height          =   285
         Left            =   120
         TabIndex        =   16
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
         Text            =   "Aplicar en Fecha"
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
      Begin Proyecto1.ButtonOffice BtnSearch 
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14457180
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
         PicNormal       =   "FrmDescPlan.frx":0000
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3165
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   873
      BorderStyle     =   2
      BorderWidth     =   2
      BorderType      =   8192
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
      GradientColor1  =   14522474
      GradientColor2  =   16640213
      BackgroundType  =   1
      ShadowOffsetX   =   10
      ShadowColor     =   0
      Begin pRmFrame.RmFrame RmFrame3 
         Height          =   435
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         BorderStyle     =   1
         BorderWidth     =   0
         BorderType      =   3
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
         GradientColor2  =   14522474
         BackgroundType  =   2
         ShadowOffsetX   =   5
         ShadowOffsetY   =   5
         ShadowColor     =   0
         Picture         =   "FrmDescPlan.frx":077A
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnAdd 
            Height          =   405
            Left            =   45
            TabIndex        =   4
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
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
            PicAlign        =   0
            PicNormal       =   "FrmDescPlan.frx":0BDC
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   560
            TabIndex        =   5
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
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
            PicAlign        =   0
            PicNormal       =   "FrmDescPlan.frx":15EE
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmDescPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnSearch_Click()
FrmSelSocios.Show vbModal, Me
If VAR_COD_SOCIO > "" Then
    CargarSocio VAR_COD_SOCIO
    TxCod_Socio.Text = VAR_COD_SOCIO
End If
End Sub
Sub CargarSocio(Codigo As String)
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    Strsql = "Select Apellidos & ', ' & Nombres AS NomSocio, Estado, Saldo_Cuenta " & _
            "from socios where Cod_Socio = '" & Codigo & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxNombres.Text = rssocio("Nomsocio").Value
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub
Sub CargarTipoServ(IdTipo As String)
    Dim TipoServ As New Recordset
    TipoServ.CursorType = adOpenKeyset
    Dim Strsql As String
    Strsql = "Select * from servicios where Cod_Servicio = '" & IdTipo & "'"
    TipoServ.Open Strsql, CN
    If Not TipoServ.EOF Then
        TxNomDesc.Text = TipoServ("Descripcion").Value
    End If
    TipoServ.Close
    Set TipoServ = Nothing
End Sub
Sub AddDescuento()
    Dim Params
    Dim fecha As String
    Dim Transac As Boolean
    
On Error Resume Next

    fecha = Format(Fecinicio, "mmm/dd/yyyy")
    Params = Array(fecha, CmbDesc.Text, TxCod_Socio.Text, TxMonto.Value)
        
    Transac = FX.CmdTransacciones("AddDescuentoPlan", Params)
    If Transac Then
        Unload Me
    Else
        MsgBox "No fue posible adicionar el registro"
    End If
    
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmDescPlan.AddDescuento." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
        
End Sub

Private Sub BtnAdd_Click()
    AddDescuento
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub CmbDesc_Click()
    CargarTipoServ CmbDesc.Text
End Sub

Private Sub CmbDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub Form_Load()
    FX.CmdFillCombos "QryServicios", CmbDesc, "", False, False
    Fecinicio.Value = FrmPlanilla.Fecinicio.Value
End Sub

Private Sub TxCod_Socio_LostFocus()
    CargarSocio TxCod_Socio.Text
End Sub
