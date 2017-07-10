VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSaldosCta 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Saldos de Cuenta"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Ccdialog 
      Left            =   6000
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4471
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
         Left            =   5400
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton OptAhorros 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Ahorros"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   5535
         Begin LabelDegradado.LabelDegrade LabelDegrade5 
            Height          =   285
            Left            =   120
            TabIndex        =   6
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
            Left            =   120
            TabIndex        =   7
            Top             =   600
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
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   49938433
            CurrentDate     =   39010
         End
         Begin MSComCtl2.DTPicker Fecfinal 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   49938433
            CurrentDate     =   39010
         End
         Begin Proyecto1.ButtonOffice BtnImprimir 
            Height          =   375
            Left            =   3120
            TabIndex        =   13
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BackColor       =   -2147483633
            Estilo          =   1
            Caption         =   "Imprimir Reprote de Saldos"
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
            PicNormal       =   "FrmSaldosCta.frx":0000
            PicSizeH        =   16
            PicSizeW        =   16
         End
      End
      Begin pRmFrame.RmFrame FrameActivo 
         Height          =   495
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
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
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   240
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         PicNormal       =   "FrmSaldosCta.frx":059A
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
End
Attribute VB_Name = "FrmSaldosCta"
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
Private Sub BtnImprimir_Click()
    Dim RptRs As New Recordset
    Dim Rptlast As New Recordset
    Dim Qry As String
    Dim TipoCta As String
    If OptAhorros.Value = True Then
        TipoCta = "AHO"
    Else
        TipoCta = "APT"
    End If
    
    Dim Paramet As Variant
    Ccdialog.ShowPrinter
    fechaI = Format(Fecinicio.Value, "MMMM-dd-yy")
    fechaF = Format((Fecfinal.Value + 1), "MMMM-dd-yy")
    
    RptRs.CursorType = adOpenKeyset
        Paramet = Array(fechaI, fechaF, TxCod_Socio.Text, TipoCta)
        Qry = "SaldosCtasCodSocio"
    FX.LoadRstFromDB Qry, RptRs, Paramet
   ' Do While Not RptRs.EOF
    '        RptRs("Saldo").Value = (RptRs("Abonos").Value - RptRs("Cargos").Value)
    '        RptRs.MoveNext
    'Loop
    With RptSaldosCtas
        Set RptSaldosCtas.DataSource = Nothing
        Set RptSaldosCtas.DataSource = RptRs.DataSource
        With .Sections("Sección4")
            .Controls.Item("LblFecha").Caption = "Del " & Fecinicio.Value & _
                                                " hasta " & Fecfinal.Value
        End With
        With .Sections("Sección2")
             .Controls.Item("NumCta").Caption = TxCod_Socio.Text
            If TipoCta = "AHO" Then
                .Controls.Item("TxCuenta").Caption = "Cuenta de Ahorros"
            Else
                .Controls.Item("TxCuenta").Caption = "Cuenta de Aportaciones"
            End If
             
             .Controls.Item("NomSocio").Caption = TxApellidos.Text & ", " & TxNombres.Text
        End With
        Rptlast.CursorType = 1
        Paramet = Array(fechaI, TxCod_Socio.Text, TipoCta)
        FX.LoadRstFromDB "SelTopLastSaldo", Rptlast, Paramet
                
        With .Sections("Sección3")
            Dim SumSaldo As Double
            Do While Not RptRs.EOF
                SumSaldo = SumSaldo + CDbl(RptRs("Saldo").Value)
                RptRs.MoveNext
            Loop
        If Rptlast.RecordCount > 0 Then
            'SumSaldo = CDbl(RptSaldosCtas.Sections("Sección5").Controls.Item("SumSaldos").Value)
            '.Controls.Item("LastSaldo").Caption = Format(Rptlast("Saldo").Value, "#,##0.00")
            .Controls.Item("LastSaldo").Caption = Format(Rptlast("Saldo").Value, "#,##0.00")
            .Controls.Item("RptSaldo").Caption = Format(SumSaldo, "#,##0.00")
            SumSaldo = SumSaldo + CDbl(.Controls.Item("LastSaldo").Caption)
            .Controls.Item("SaldoCta").Caption = SumSaldo
        Else
            .Controls.Item("LastSaldo").Caption = "0.00"
            SumSaldo = SumSaldo + CDbl(.Controls.Item("LastSaldo").Caption)
            .Controls.Item("RptSaldo").Caption = Format(SumSaldo, "#,##0.00")
            .Controls.Item("SaldoCta").Caption = SumSaldo
        End If
        End With
    .Show vbModal, Me
    End With
End Sub
Sub CargarSocio(Codigo As String)
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    Strsql = "Select * from socios where Cod_Socio = '" & Codigo & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxCod_Socio.Text = rssocio("Cod_Socio").Value
        TxNombres.Text = rssocio("Nombres").Value
        TxApellidos.Text = rssocio("Apellidos").Value
        If rssocio("Estado").Value Then
            ChkActivo.Value = 1
'            BtnMovCta.Enabled = True
            Frame1.Enabled = True
        Else
            ChkActivo.Value = 0
 '           BtnMovCta.Enabled = False
            Frame1.Enabled = False
        End If
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub

Private Sub Form_Load()
    FX.ConnectDb activar
    Fecinicio.Value = Date
    Fecfinal.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FX.ConnectDb Desactivar
End Sub
Private Sub TxCod_Socio_LostFocus()
    CargarSocio TxCod_Socio.Text
End Sub
