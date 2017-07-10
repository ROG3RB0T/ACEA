VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPlanillaPtamo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Cuotas de Prestamo"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
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
      Text            =   "Nota: El presente reporte de Cuotas de Prestamos  será aplicado en la planilla"
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre del Socio"
         Object.Width           =   6429
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número de Ptamo."
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Monto Cuota"
         Object.Width           =   2540
      EndProperty
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4650
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   873
      BorderStyle     =   2
      BorderWidth     =   2
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
      Begin pRmFrame.RmFrame RmFrame3 
         Height          =   435
         Left            =   30
         TabIndex        =   3
         Top             =   45
         Width           =   5535
         _ExtentX        =   9763
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
         Picture         =   "FrmPlanillaPtamo.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin MSComCtl2.DTPicker FecAplicar 
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   80
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Format          =   20971521
            CurrentDate     =   39033
         End
         Begin Proyecto1.ButtonOffice BtnAplicar 
            Height          =   405
            Left            =   2880
            TabIndex        =   5
            Top             =   15
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   "Aplicar en Planilla"
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
            PicNormal       =   "FrmPlanillaPtamo.frx":0462
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   4695
            TabIndex        =   6
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
            PicNormal       =   "FrmPlanillaPtamo.frx":0E74
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha App."
            Height          =   270
            Left            =   40
            TabIndex        =   7
            Top             =   105
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmPlanillaPtamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AddDescuento()
    Dim Params
    Dim fecha As String
    Dim Transac As Boolean
    Dim RsA As New Recordset
On Error Resume Next
    RsA.CursorType = 1
    
    FX.LoadRstFromDB "QryMontoPtamos", RsA, ""
    
    If RsA.RecordCount > 0 Then
        Do While Not RsA.EOF
            fecha = Format(FecAplicar.Value, "mmm/dd/yyyy")
            Params = Array(fecha, "PTMO*", RsA("Cod_Socio").Value, RsA("MontoCuota").Value)
                
            Transac = FX.CmdTransacciones("AddDescuentoPlan", Params)
        
            If Transac Then
                RsA.MoveNext
            Else
                MsgBox "No fue posible adicionar el registro"
                Transac = False
                Exit Do
            End If
        Loop
        
        If Transac Then
            MsgBox "Se han procesado " & RsA.RecordCount & " registros", vbInformation
            BtnAplicar.Enabled = False
            ListView1.Enabled = False
        End If
    End If
    
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmDescPlan.AddDescuento." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
        
End Sub

Private Sub BtnAplicar_Click()
    AddDescuento
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As New Recordset
    rs.CursorType = 1
    
    FX.ConnectDb activar
    FX.LoadRstFromDB "QryMontoPtamos", rs, ""
    FX.LoadListView rs, ListView1
    FecAplicar.Value = Date
End Sub
