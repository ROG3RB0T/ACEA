VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConsultaPtamo 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Consulta de Prestamos"
   ClientHeight    =   7140
   ClientLeft      =   885
   ClientTop       =   420
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CCd 
      Left            =   4800
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11245
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
      Begin VB.TextBox TxInteres 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Historial del Prestamo"
         BackColor       =   255
         BorderColor     =   9655840
         Transparente    =   0   'False
         ShadowDepth     =   0
         ShadowStyle     =   0
         ShadowColorStart=   0
         Alignment       =   2
         DegradadoOrientacion=   2
         DegradadoColorStart=   13993792
         DegradadoColorEnd=   12218153
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2090
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Saldo Anterior"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Interes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Abono Capital"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo Ptamo"
            Object.Width           =   2540
         EndProperty
      End
      Begin TextBoxWinXP.TextboxXP TxNombre 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
      Begin VB.TextBox TxNumPtamo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Número de Prestamo"
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
         Left            =   4440
         TabIndex        =   6
         Top             =   240
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
         PicNormal       =   "FrmConsultaPtamo.frx":0000
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Nombre del Cliente"
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
      Begin TxNumOcx.TxNum TxMontoPtamo 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
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
         Locked          =   -1  'True
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Monto Ptamo."
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
         Left            =   3600
         TabIndex        =   11
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Interés(%)"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade7 
         Height          =   285
         Left            =   6840
         TabIndex        =   12
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
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
         Text            =   "Plazo"
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
      Begin TextBoxWinXP.TextboxXP TxPlazo 
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
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
      Begin LabelDegradado.LabelDegrade LabelDegrade9 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Monto Cuota"
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
      Begin TxNumOcx.TxNum TxCuota 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
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
         Locked          =   -1  'True
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade10 
         Height          =   285
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Fecha Otorgado"
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
      Begin TextBoxWinXP.TextboxXP FecOtorgado 
         Height          =   285
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
      Begin LabelDegradado.LabelDegrade LabelDegrade11 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Estdo del Ptamo"
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
      Begin TextBoxWinXP.TextboxXP TxEstado 
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   6645
      Width           =   10125
      _ExtentX        =   17859
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
         TabIndex        =   3
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "FrmConsultaPtamo.frx":077A
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BarBtnLogin 
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
            PicNormal       =   "FrmConsultaPtamo.frx":0BDC
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnPrint 
            Height          =   405
            Left            =   520
            TabIndex        =   23
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
            PicNormal       =   "FrmConsultaPtamo.frx":15EE
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmConsultaPtamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ConsultaPtamo()
    Dim RsPtamo As New Recordset
    Dim RsHistorial As New Recordset
    RsHistorial.CursorType = 1
    RsPtamo.CursorType = 1
    
    FX.LoadRstFromDB "Qryprestamo", RsPtamo, TxNumPtamo.Text
    
    If RsPtamo.RecordCount > 0 Then
        TxNombre.Text = RsPtamo("Nombre").Value
        FecOtorgado.Text = RsPtamo("FechaOtorgado").Value
        TxMontoPtamo.Value = RsPtamo("MontoPtamo").Value
        TxInteres.Text = RsPtamo("Interes").Value & "%"
        TxPlazo.Text = RsPtamo("Plazo").Value & " " & RsPtamo("FormaPago").Value
        TxCuota.Value = RsPtamo("MontoCuota").Value
        TxEstado.Text = RsPtamo("EstadoPtamo").Value
'        Frame1.Enabled = True
    End If
            
    FX.LoadRstFromDB "QryHistorialPtamo", RsHistorial, TxNumPtamo.Text
    
    If RsHistorial.RecordCount > 0 Then
        FX.LoadListView RsHistorial, ListView1, False
    End If
        
End Sub

Private Sub BarBtnLogin_Click()
 Unload Me
End Sub

Private Sub BtnFiltrar_Click()
    Dim RptPtmo As New Recordset
On Error Resume Next

    RptPtmo.CursorType = 1
    
    FX.LoadRstFromDB "QryCuotasPtamos", RptPtmo, Array(Fecinicio.Value, Fecfinal.Value, TxNumPtamo.Text)
    
    If RptPtmo.RecordCount > 0 Then
        FX.LoadListView RptPtmo, ListView1
    End If

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmConsultaPtamo.BtnFiltrar_Click." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
    
End Sub

Private Sub BtnPrint_Click()
If TxNumPtamo.Text > "" Then
    Dim RsHistorialrpt As New Recordset
    RsHistorialrpt.CursorType = 1
    
        With RptEstadoPtamo
            With .Sections("Sección2")
                .Controls.item("NumPtamo").Caption = TxNumPtamo.Text
                .Controls.item("NombrePtamo").Caption = TxNombre.Text
                .Controls.item("MontoPtamo").Caption = Format(TxMontoPtamo.Value, "#,##0.00")
            End With
            With .Sections("Sección4")
                .Controls.item("lblfecha").Caption = Format(Date, "dd - MMMM - yyyy")
            End With
        End With
        
    FX.LoadRstFromDB "QryHistorialPtamo", RsHistorialrpt, TxNumPtamo.Text
    
    If RsHistorialrpt.RecordCount > 0 Then
        CCD.ShowPrinter
        With RptEstadoPtamo
            Set .DataSource = Nothing
            Set .DataSource = RsHistorialrpt.DataSource
            .Show vbModal, Me
        End With
    End If
End If
End Sub

Private Sub BtnSearch_Click()
    ConsultaPtamo
End Sub

Private Sub Form_Load()
    FX.ConnectDb activar
'    Fecinicio.Value = Date
 '   Fecfinal.Value = Date
End Sub

Private Sub ListView1_Click()
    For i = 1 To ListView1.ColumnHeaders.Count
        Debug.Print ListView1.ColumnHeaders(i).Text; _
                ListView1.ColumnHeaders(i).Width
    Next
End Sub

Private Sub TxNumPtamo_LostFocus()
    ConsultaPtamo
End Sub

Private Sub TxPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub

Private Sub TxNumPtamo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub
