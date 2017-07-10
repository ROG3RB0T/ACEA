VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCtas 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas de Socios"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CcDialog 
      Left            =   5520
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9763
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
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "MOVIMIENTOS DE LA CUENTA"
         BackColor       =   255
         BorderColor     =   0
         BorderSize      =   2
         Transparente    =   0   'False
         ShadowDepth     =   0
         ShadowStyle     =   0
         ShadowColorStart=   0
         Alignment       =   2
         DegradadoOrientacion=   3
         DegradadoColorStart=   10838307
         DegradadoColorEnd=   14988165
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo Tran"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comentario"
            Object.Width           =   5980
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cargo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Abono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   18
         EndProperty
      End
      Begin Proyecto1.ButtonOffice BtnDelSelected 
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   5160
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BackColor       =   -2147483633
         Estilo          =   1
         Caption         =   "Eliminar registro seleccioando"
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
         PicNormal       =   "FrmCtas.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3836
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
      Begin VB.OptionButton OptAportaciones 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Aportaciones"
         Height          =   255
         Left            =   7560
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton OptAhorros 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Ahorros"
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin pRmFrame.RmFrame FrameActivo 
         Height          =   495
         Left            =   6000
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BorderStyle     =   6
         BorderWidth     =   0
         Caption         =   ""
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
         Begin VB.CheckBox ChkActivo 
            BackColor       =   &H00F7F7F7&
            Caption         =   "Activo"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   10935
         Begin Proyecto1.ButtonOffice BtnImprimir 
            Height          =   255
            Left            =   6000
            TabIndex        =   4
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483633
            Estilo          =   1
            Caption         =   "Imprimir"
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
            PicNormal       =   "FrmCtas.frx":059A
            PicSizeH        =   16
            PicSizeW        =   16
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnFiltrar 
            Height          =   255
            Left            =   5040
            TabIndex        =   3
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483633
            Estilo          =   1
            Caption         =   "Filtrar"
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
            PicNormal       =   "FrmCtas.frx":0B34
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin LabelDegradado.LabelDegrade LabelDegrade5 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "Desde Fecha"
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
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "Hasta Fecha"
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
         Begin MSComCtl2.DTPicker Fecinicio 
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   39010
         End
         Begin MSComCtl2.DTPicker Fecfinal 
            Height          =   285
            Left            =   3600
            TabIndex        =   2
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   39010
         End
         Begin Proyecto1.ButtonOffice BtnMovCta 
            Height          =   255
            Left            =   7080
            TabIndex        =   19
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackColor       =   -2147483633
            Estilo          =   1
            Caption         =   "Movmientos de Cta."
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
            PicNormal       =   "FrmCtas.frx":10CE
            PicSizeH        =   16
            PicSizeW        =   16
            State           =   3
         End
         Begin Proyecto1.ButtonOffice ButtonOffice1 
            Height          =   255
            Left            =   9000
            TabIndex        =   23
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackColor       =   -2147483633
            Estilo          =   1
            Caption         =   "ReCalular Saldos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicNormal       =   "FrmCtas.frx":1668
            PicSize         =   5
            PicSizeH        =   16
            PicSizeW        =   16
         End
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   240
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
         Left            =   1680
         TabIndex        =   10
         Top             =   600
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
         Locked          =   -1  'True
      End
      Begin TextBoxWinXP.TextboxXP TxCod_Socio 
         Height          =   285
         Left            =   1680
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
      Begin TextBoxWinXP.TextboxXP TxApellidos 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   960
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
         Locked          =   -1  'True
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   285
         Left            =   240
         TabIndex        =   12
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
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   240
         TabIndex        =   13
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
      Begin Proyecto1.ButtonOffice BtnSearch 
         Height          =   285
         Left            =   3480
         TabIndex        =   20
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
         PicNormal       =   "FrmCtas.frx":2062
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin pRmFrame.RmFrame RmFrame3 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   24
      Top             =   7935
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   847
      BorderStyle     =   2
      BorderWidth     =   0
      BorderType      =   6
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
      GradientColor2  =   16640471
      BackgroundType  =   1
      ShadowOffsetX   =   10
      ShadowColor     =   0
      Begin pRmFrame.RmFrame RmFrame4 
         Height          =   435
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "FrmCtas.frx":27DC
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   400
            Left            =   40
            TabIndex        =   26
            Top             =   8
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
            PicNormal       =   "FrmCtas.frx":2C3E
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   16
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechaI As String
Dim fechaF As String
Dim TipoCta As String
Dim RST As New Recordset
Dim Paramet As Variant


Private Sub BtnDelSelected_Click()
    If MsgBox("Desea Eliminar el registro seleccionado?", vbInformation + vbYesNo) = vbYes Then
        Dim Transac As Boolean
        Transac = FX.CmdTransacciones("DeleteMovCta", ListView1.SelectedItem.ListSubItems(6).Text)
        If Transac Then
            ButtonOffice1_Click
        End If
    End If
End Sub
Sub CargarSocio(Codigo As String)
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    ListView1.ListItems.Clear
    Strsql = "Select * from socios where Cod_Socio = '" & Codigo & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxCod_Socio.Text = rssocio("Cod_Socio").Value
        TxNombres.Text = rssocio("Nombres").Value
        TxApellidos.Text = rssocio("Apellidos").Value
        If rssocio("Estado").Value Then
            ChkActivo.Value = 1
            BtnMovCta.Enabled = True
            Frame1.Enabled = True
        Else
            ChkActivo.Value = 0
            BtnMovCta.Enabled = False
            Frame1.Enabled = False
        End If
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub
Sub CalculoSaldos(Codigo As String, TipoCtaSoc As String)
   Dim i As Integer
    Dim Strqry As String
    Dim Saldo As Double
    Dim Paramets
    Dim RST As Recordset
    Set RST = New Recordset
    RST.CursorType = 1
    
    'RST.LockType = adLockOptimistic
    Strqry = "ReCalculoSaldosCtasSocios"
    'For i = 1 To ListView1.ListItems.Count
            'VAR_COD_SOCIO = ListView1.ListItems(i).Text
            FX.LoadRstFromDB Strqry, RST, Array(Codigo, "AHO"), DBQuery
            'ListView1.Visible = False
            'ListView2.Visible = True
            'FX.LoadListView RST, ListView2
            'MsgBox "Press Ok"
            If RST.RecordCount > 0 Then
                Saldo = 0
                Do While Not RST.EOF
                'ListView1.Visible = True
                    Debug.Print "Saldo: " & Saldo & " "; RST("Cargo").Value & " " & RST("Abono") & " " & _
                                RST("Saldo").Value
                    Saldo = (RST("Abono").Value - RST("Cargo").Value) + Saldo
                    'If RST("Saldo").Value <> Saldo Then
                        Paramets = Array(Saldo, RST("Id_Transac").Value)
                        TxOk = FX.CmdTransacciones("UpdateSaldoenCtaSocios", Paramets)
                        If TxOk Then
                            RST.MoveNext
                        Else
                            Exit Do
                        End If
                        'Paramets = Array(Saldo, VAR_COD_SOCIO)
                        'FX.CmdTransacciones "QryUpdateSaldoSocio", Paramets
                    'End If
                    'RST.MoveNext
                Loop
                Paramets = Array(Saldo, Codigo)
                If TipoCtaSoc = "AHO" Then FX.CmdTransacciones "QryUpdateSaldoSocio", Paramets
            End If
            'ListView1.ListItems(i).Bold = True
        'Bar1.Value = Bar1.Value + i
    'Next
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnFiltrar_Click()
    If OptAhorros.Value = True Then
        TipoCta = "AHO"
    Else
        TipoCta = "APT"
    End If
    fechaI = Format(Fecinicio.Value, "MMM-dd-yy")
    fechaF = Format((Fecfinal.Value + 1), "MMM-dd-yy")
    Paramet = Array(fechaI, fechaF, TipoCta, TxCod_Socio.Text)
    RST.CursorType = 1
    FX.LoadRstFromDB "QryLstEstadoCta", RST, Paramet
    'Do While Not RST.EOF
        ' para verificar como esta leyendo el recordset antes de pasar al listview
    '    Debug.Print RST("Fecha").Value & " " & RST("Cod_Trans").Value & " " & RST("Comentario").Value
    '    RST.MoveNext
    'Loop
    FX.LoadListView RST, ListView1, False
    BtnImprimir.Enabled = True
End Sub

Private Sub BtnImprimir_Click()
    Dim Qry As String
    Dim RptRs As New Recordset
    Ccdialog.ShowPrinter
    
    RptRs.CursorType = 1
    If OptAhorros.Value = True Then
        TipoCta = "AHO"
    Else
        TipoCta = "APT"
    End If
    'If CmbTipoMov.Text = "TODOS" Then
        Paramet = Array(fechaI, fechaF, TipoCta, TxCod_Socio.Text)
        Qry = "QryEstadoCta_Socio"
   ' Else
        'Paramet = Array(fechaI, fechaF, TxCod_Socio.Text, CmbTipoMov.Text)
        'Qry = "QryLstTipoEstadoCta"
    'End If
    FX.LoadRstFromDB Qry, RptRs, Paramet
    With RptEstadoCta
        Set RptEstadoCta.DataSource = Nothing
        Set RptEstadoCta.DataSource = RptRs.DataSource
        With .Sections("Sección4")
        If OptAhorros.Value = True Then
            .Controls.item("EstadoCta").Caption = "Detalle de Cuenta Ahorros de Socios"
        Else
            .Controls.item("EstadoCta").Caption = "Detalle de Cuenta Aportaciones de Socios"
        End If
            
            .Controls.item("LblFecha").Caption = "Del " & Fecinicio.Value & _
                                                " hasta " & Fecfinal.Value
        End With
        With .Sections("Sección2")
             .Controls.item("NumCta").Caption = TxCod_Socio.Text
             .Controls.item("NomSocio").Caption = TxApellidos.Text & ", " & TxNombres.Text
        End With
    .Show vbModal, Me
    End With
End Sub

Private Sub BtnMovCta_Click()
    VAR_COD_SOCIO = ""
    VAR_COD_SOCIO = TxCod_Socio.Text
    FrmMov.Show vbModal, Me
End Sub

Private Sub BtnSave_Click()
    
End Sub

Private Sub BtnSearch_Click()
FrmSelSocios.Show vbModal, Me
If VAR_COD_SOCIO > "" Then
    CargarSocio VAR_COD_SOCIO
    TxCod_Socio.Text = VAR_COD_SOCIO
End If
End Sub

Private Sub ButtonOffice1_Click()
    If OptAhorros.Value = True Then
        CalculoSaldos TxCod_Socio.Text, "AHO"
    Else
        CalculoSaldos TxCod_Socio.Text, "APT"
    End If
'CalculoSaldos TxCod_Socio.Text, "AHO"
BtnFiltrar_Click
End Sub

Private Sub Form_Load()
    FrameActivo.Enabled = False
    FX.ConnectDb activar
    Fecinicio.Value = Date
    Fecfinal.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FX.ConnectDb Desactivar
End Sub

Private Sub ListView1_Click()
   Debug.Print ListView1.ColumnHeaders(1).Text; ListView1.ColumnHeaders(1).Width
   Debug.Print ListView1.ColumnHeaders(2).Text; ListView1.ColumnHeaders(2).Width
   Debug.Print ListView1.ColumnHeaders(3).Text; ListView1.ColumnHeaders(3).Width
   Debug.Print ListView1.ColumnHeaders(4).Text; ListView1.ColumnHeaders(4).Width
   Debug.Print ListView1.ColumnHeaders(5).Text; ListView1.ColumnHeaders(5).Width
   Debug.Print ListView1.ColumnHeaders(6).Text; ListView1.ColumnHeaders(6).Width
End Sub

Private Sub TxCod_Socio_LostFocus()
    CargarSocio TxCod_Socio.Text
End Sub
