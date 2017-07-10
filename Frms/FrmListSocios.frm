VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmListSocios 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado Socios"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CcDialog 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Height          =   4935
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
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
      Begin VB.OptionButton Opt2 
         Caption         =   "Mostrar todos los Socios"
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   600
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Mostrar solo Socios Activos"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7575
         _ExtentX        =   13361
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
         Text            =   "LISTADO DE SOCIOS"
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
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   5130
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   873
      BorderStyle     =   10
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
      GradientColor2  =   16640213
      BackgroundType  =   1
      ShadowOffsetX   =   10
      ShadowColor     =   0
      Begin Proyecto1.ButtonOffice ButtonOffice1 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   12230304
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
      Begin pRmFrame.RmFrame RmFrame3 
         Height          =   435
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
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
         Picture         =   "FrmListSocios.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnPrintList 
            Height          =   405
            Left            =   645
            TabIndex        =   5
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14324046
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
            PicNormal       =   "FrmListSocios.frx":0462
            PicOpacity      =   0.65
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice BtnNewSocio 
            Height          =   405
            Left            =   120
            TabIndex        =   6
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14324046
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
            PicNormal       =   "FrmListSocios.frx":0BDC
            PicOpacity      =   0.5
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnRefresh 
            Height          =   405
            Left            =   1140
            TabIndex        =   7
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14324046
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
            PicNormal       =   "FrmListSocios.frx":1356
            PicOpacity      =   0.65
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   1635
            TabIndex        =   8
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14324046
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
            PicNormal       =   "FrmListSocios.frx":1AD0
            PicOpacity      =   0.65
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
      End
   End
End
Attribute VB_Name = "FrmListSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst As New Recordset
Dim RptSqlstr As String
Dim LstSqlstr As String
Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnNewSocio_Click()
    FrmSocios.Cod_Socio = ""
    FrmSocios.Show vbModal, Me
End Sub

Private Sub BtnPrintList_Click()
    CcDialog.ShowPrinter
    Rst.Open RptSqlstr, CN
    'FX.LoadRstFromDB "RptListSocios", Rst, ""
    Set RptListSocios.DataSource = Rst.DataSource
    With RptListSocios
        Debug.Print .ReportWidth
        'Debug.Print .Orientation
        Debug.Print .Width
        With .Sections("Sección3").Controls
            .Item("Etiqueta7").Caption = Rst.RecordCount
        End With
        With .Sections("Sección4").Controls
            .Item("Etiqueta2").Caption = "al " & Format(Date, "dd/MMMM/yyyy")
        End With
    End With
    Debug.Print "Rpt: "; RptListSocios.Height
    RptListSocios.Show vbModal, Me
    Rst.Close
End Sub

Private Sub BtnRefresh_Click()
'    Rst.Close
'    FX.LoadRstFromDB "QrySelSocios", Rst, ""
    LstSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Nombres, " & _
                "Socios.Apellidos, Socios.Estado AS Activo " & _
                "FROM Socios"
    If Opt1 Then
        LstSqlstr = LstSqlstr & " Where Estado = true"
    Else
        LstSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Nombres, " & _
                "Socios.Apellidos, Socios.Estado AS Activo " & _
                "FROM Socios"
    End If
    Rst.Open LstSqlstr, CN
    FX.LoadListView Rst, ListView1
End Sub

Private Sub ButtonOffice1_Click()
    Dim Nvosaldo As Double
    Dim SaldoCta As Double
    Dim Paramet
    Dim Txok As Boolean
    Dim TipoCta As String
    Dim RstVal As New Recordset
    RstVal.CursorType = adOpenKeyset
    For i = 1 To ListView1.ListItems.Count
        FX.LoadRstFromDB "QryEvalLastSaldo", RstVal, Array(ListView1.ListItems(i).Text, "AHO")
        SaldoCta = 0
        If RstVal.RecordCount > 0 Then
            SaldoCta = RstVal("Saldo").Value
        End If
        Txok = FX.CmdTransacciones("QryUpdateSaldoSocio", Array(SaldoCta, ListView1.ListItems(i).Text))
        If Txok Then
            RstVal.Close
        End If
    Next
        If Txok Then MsgBox "finalizado"
End Sub

Private Sub Form_Load()
    Rst.CursorType = adOpenKeyset
    LstSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Nombres, " & _
                "Socios.Apellidos, Socios.Estado AS Activo " & _
                "FROM Socios"
    If Opt1 Then
        LstSqlstr = LstSqlstr & " Where Estado = true"
    Else
        LstSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Nombres, " & _
                "Socios.Apellidos, Socios.Estado AS Activo " & _
                "FROM Socios"
    End If
    
    FX.ConnectDb activar
    Rst.Open LstSqlstr, CN
    'FX.LoadRstFromDB "QrySelSocios", Rst, ""
    FX.LoadListView Rst, ListView1
    RptSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Apellidos +', '+ Socios.Nombres AS " & "[Nombre del Socio], Socios.Estado AS Activo From Socios"
                
End Sub

Private Sub Form_Unload(Cancel As Integer)
        FX.ConnectDb Desactivar
End Sub

Private Sub ListView1_Click()
On Error Resume Next
    FrmSocios.Cod_Socio = ListView1.SelectedItem.Text
    If FrmSocios.Cod_Socio > "" Then
        FrmSocios.Show vbModal, Me
    End If
    BtnRefresh_Click
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As _
    MSComctlLib.ColumnHeader)
If ColumnHeader <> "Código" And ColumnHeader <> "Activo" Then
    ListView1.SortKey = ColumnHeader.Index - 1
    RptSqlstr = "SELECT Socios.Cod_Socio AS Código, Socios.Apellidos +', '+ Socios.Nombres AS " & _
                "[Nombre del Socio], Socios.Estado AS Activo From Socios"
                
    If Opt1 Then
        RptSqlstr = RptSqlstr & " Where Estado = true"
    End If
                
    If ListView1.SortOrder = lvwAscending Then
       ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
       If ListView1.SortOrder = lvwAscending Then
            RptSqlstr = RptSqlstr & _
           " ORDER BY " & CStr(ColumnHeader) & " ASC"
           
       Else
            RptSqlstr = RptSqlstr & _
           " ORDER BY " & CStr(ColumnHeader) & " DESC"
       End If
    ListView1.Sorted = True
       Debug.Print RptSqlstr
End If
End Sub
