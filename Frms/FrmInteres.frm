VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInteres 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CCD 
      Left            =   4440
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Calculo de Intereses de las cuentas de Ahorros de los Socios"
      BackColor       =   255
      ForeColor       =   0
      BorderColor     =   8388608
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      Alignment       =   2
      DegradadoColorStart=   14522474
      DegradadoColorEnd=   16640213
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cap"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NvoSaldo"
         Object.Width           =   2540
      EndProperty
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4350
      Width           =   11550
      _ExtentX        =   20373
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
         Top             =   45
         Width           =   8775
         _ExtentX        =   15478
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
         Picture         =   "FrmInteres.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin TxNumOcx.TxNum Txinteres 
            Height          =   315
            Left            =   4400
            TabIndex        =   13
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   16773091
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Text            =   "0.0000"
            Value           =   "0.0000"
            BorderColor     =   14522474
            Numdec          =   4
            Moneda          =   ""
         End
         Begin MSComCtl2.DTPicker DTFecha 
            Height          =   315
            Left            =   2950
            TabIndex        =   12
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   39210
         End
         Begin Proyecto1.ButtonOffice BtnAplicar 
            Height          =   405
            Left            =   5520
            TabIndex        =   4
            Top             =   15
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   "Imprimir Reporte"
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
            PicNormal       =   "FrmInteres.frx":0462
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   8040
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
            PicNormal       =   "FrmInteres.frx":0E74
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice ButtonOffice1 
            Height          =   405
            Left            =   7080
            TabIndex        =   8
            Top             =   15
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   "Excel"
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
            PicNormal       =   "FrmInteres.frx":140E
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   24
            PicSizeW        =   24
         End
         Begin Proyecto1.ButtonOffice ButtonOffice2 
            Height          =   390
            Left            =   60
            TabIndex        =   10
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            BackColor       =   14522474
            Caption         =   "Leer Saldos"
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
            PicNormal       =   "FrmInteres.frx":1B88
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice ButtonOffice3 
            Height          =   390
            Left            =   1320
            TabIndex        =   11
            Top             =   15
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            BackColor       =   14522474
            Caption         =   "Aplicar en Ctas."
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
            PicNormal       =   "FrmInteres.frx":2302
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         DragMode        =   1  'Automatic
         Height          =   255
         Left            =   8880
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade5 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
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
   Begin MSComCtl2.DTPicker Fecinicio 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   39010
   End
End
Attribute VB_Name = "FrmInteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub ButtonOffice1_Click()
    Dim ret As Boolean
    If ListView1.ListItems.Count > 0 Then ret = FX.Exportar_Excel("", ListView1, ProgressBar1)
End Sub
Private Sub BtnAplicar_Click()
    CCD.ShowPrinter
    Dim Rscap As Recordset
    Set Rscap = New Recordset
    Rscap.CursorType = 1
    
    FX.LoadRstFromDB "QryInteresCtaAHO", Rscap, Txinteres.Value
    
    With RptCap
        Set .DataSource = Nothing
        Set .DataSource = Rscap.DataSource
        With .Sections("Sección4")
            .Controls.item("LblFecha").Caption = "Al mes de " & Format(Date, "MMMM") & " del " & Format(Fecinicio.Value, "yyyy")
        End With
        RptCap.Show vbModal, Me
    End With
End Sub

Private Sub BtnExit_Click()
 Unload Me
End Sub

Private Sub ButtonOffice2_Click()

    Dim rs As Recordset
    Set rs = New Recordset
    rs.CursorType = 1
    
    FX.LoadRstFromDB "QryInteresCtaAHO", rs, Txinteres.Value
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.Visible = True
    Screen.MousePointer = vbHourglass
    Do While Not rs.EOF
        CalculoSaldos rs("Cod_Socio").Value, "AHO"
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    Screen.MousePointer = vbDefault
    MsgBox "Proceso finalizado"
    ProgressBar1.Visible = False
    rs.MoveFirst
    ListView1.SortKey = 1
    FX.LoadListView rs, ListView1
    If ListView1.ListItems.Count > 0 Then ButtonOffice3.Enabled = True
End Sub

Private Sub ButtonOffice3_Click()
    If MsgBox("Esta seguro de aplicar los " & ListView1.ListItems.Count & _
        " registros", vbInformation + vbYesNo) = vbYes Then
        Dim rs As Recordset
        Dim Paramet
        Dim TxOk As Boolean
        Dim Nvosaldo As Double
        Set rs = New Recordset
        rs.CursorType = 1
        
        FX.LoadRstFromDB "QryInteresCtaAHO", rs, Txinteres.Value
        ProgressBar1.Value = 0
        ProgressBar1.Min = 0
        ProgressBar1.Max = rs.RecordCount
        ProgressBar1.Visible = True
        
        Screen.MousePointer = vbHourglass
        Do While Not rs.EOF
            
            Nvosaldo = rs("Saldo_Cuenta").Value + rs("Cap").Value
            Paramet = Array(Format(DTFecha.Value, "MMMM/dd/yyyy"), "CAP", 0, _
                rs("Cap").Value, Nvosaldo, "Capitalización de fecha " & _
                Format(DTFecha.Value, "dd/MMMM/yyyy"), rs("Cod_Socio").Value, "AHO")
                
            TxOk = FX.CmdTransacciones("QryInsertMovCta", Paramet)
            
            If TxOk Then
                CalculoSaldos rs("Cod_Socio").Value, "AHO"
                rs.MoveNext
                ProgressBar1.Value = ProgressBar1.Value + 1
            Else
                Exit Do
                MsgBox "No se pudo procesar la transacción para la cuenta " & _
                    rs("Cod_Socio").Value & ".", vbInformation
            End If
        Loop
        Screen.MousePointer = vbDefault
        MsgBox "Proceso finalizado"
        ProgressBar1.Visible = False
        ButtonOffice3.Enabled = False
    End If
    FX.GrabaINI IniFile, "CAP", "Porcentaje", CStr(Txinteres.Value)
    
End Sub

Private Sub Form_Load()
    FX.ConnectDb activar
    Txinteres.Value = VAR_CAP
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.SortOrder = lvwAscending Then
       ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
End Sub

Private Sub Txinteres_LostFocus()
FX.GrabaINI IniFile, "CAP", "Porcentaje", CStr(Txinteres.Value)
End Sub
