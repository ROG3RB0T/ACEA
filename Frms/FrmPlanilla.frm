VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPlanilla 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Planilla de Descuentos"
   ClientHeight    =   5145
   ClientLeft      =   2310
   ClientTop       =   420
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10065
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Ccd 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
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
      Begin Proyecto1.ButtonOffice BtnAddDesc 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BackColor       =   -2147483633
         Estilo          =   1
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
         PicNormal       =   "FrmPlanilla.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   80
         TabIndex        =   5
         Top             =   600
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         NumItems        =   0
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
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
         Text            =   "Aplicación de Planilla en Fecha"
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
         Left            =   2640
         TabIndex        =   4
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   16449539
         CurrentDate     =   39010
      End
      Begin Proyecto1.ButtonOffice BtnSearch 
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         BackColor       =   14457180
         Caption         =   "Ver Descuentos"
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
         PicAlign        =   3
         PicNormal       =   "FrmPlanilla.frx":059A
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Proyecto1.ButtonOffice BtnDelSelected 
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BackColor       =   -2147483633
         Estilo          =   1
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
         PicNormal       =   "FrmPlanilla.frx":0D14
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin pRmFrame.RmFrame RmFrame4 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4650
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   873
      BorderStyle     =   8
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         DragMode        =   1  'Automatic
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin pRmFrame.RmFrame RmFrame5 
         Height          =   435
         Left            =   30
         TabIndex        =   2
         Top             =   45
         Width           =   4575
         _ExtentX        =   8070
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
         Picture         =   "FrmPlanilla.frx":12AE
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   3240
            TabIndex        =   9
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
            PicNormal       =   "FrmPlanilla.frx":1710
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnDeletePlanilla 
            Height          =   405
            Left            =   120
            TabIndex        =   10
            Top             =   15
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   "Eliminar Planilla"
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
            PicNormal       =   "FrmPlanilla.frx":1CAA
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice ButtonOffice1 
            Height          =   405
            Left            =   1920
            TabIndex        =   11
            Top             =   15
            Width           =   1290
            _ExtentX        =   2275
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
            PicNormal       =   "FrmPlanilla.frx":2244
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   24
            PicSizeW        =   24
         End
      End
   End
End
Attribute VB_Name = "FrmPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LeerDescuentos(fechainicio As String)
    Dim rs As Recordset
    Set rs = New Recordset
    rs.CursorType = 1
    
    FX.LoadRstFromDB "DescuentosSocios", rs, fechainicio
    
    If rs.RecordCount > 0 Then
'        If CBool(Rs("Conciliado").Value) = True Then
'            ListView1.Enabled = False
'            BtnAplicar.Enabled = False
'            BtnPrint.Enabled = True
'            BtnDeletePlanilla.Enabled = False
'            BtnAddDesc.Enabled = False
'            BtnDelSelected.Enabled = False
'        Else
'            ListView1.Enabled = True
'            BtnAplicar.Enabled = True
'            BtnPrint.Enabled = False
'            BtnDeletePlanilla.Enabled = True
'            BtnAddDesc.Enabled = True
'            BtnDelSelected.Enabled = True
        ListView1.ListItems.Clear
        FX.LoadListView rs, ListView1
    End If
    'Rs.Close
    Set rs = Nothing
End Sub
Sub AplicarPlanilla()
    If MsgBox("Esta seguro de procesar la planilla?", vbInformation + vbYesNo) = vbYes Then
       Dim Transac As Boolean
       Transac = FX.CmdTransacciones("UpdateDescuentos", Format(Fecinicio.Value, "MMM-dd-yyyy"))
            If Transac Then
                MsgBox "Se ha procesado la planilla satisfactoriamente"
                BtnPrint.Enabled = True
                BtnAplicar.Enabled = False
                BtnDeletePlanilla.Enabled = False
                ListView1.Enabled = False
                BtnAddDesc.Enabled = False
                BtnDelSelected.Enabled = False
            Else
                MsgBox "Ocurrió un problema durante el procesamiento, no fue posible procesar la planilla", vbInformation
            End If
    End If
End Sub

Private Sub BtnAddDesc_Click()
    FrmDescPlan.Show vbModal, Me
End Sub

Private Sub BtnAplicar_Click()
    AplicarPlanilla
End Sub

Private Sub BtnDeletePlanilla_Click()
On Error Resume Next
If ListView1.ListItems.Count > 0 Then
    If MsgBox("Esta seguro de eliminar la planilla?", vbInformation + vbYesNo) = vbYes Then
        Dim Transac As Boolean
        Dim Paramet
        Do
        Paramet = Array(Format(Fecinicio.Value, "MMM/dd/yyyy"))
            Transac = FX.CmdTransacciones("DeletePlanillaDto", Paramet)
            If Transac Then
                ListView1.ListItems.Remove ListView1.ListItems(1).Index
            End If
        Loop While Not ListView1.ListItems.Count < 1
        Unload Me
    End If
End If
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmPlanilla.BtnDeletePlanilla_Click." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
End Sub

Private Sub BtnDelSelected_Click()
    If MsgBox("Desea Eliminar el registro seleccionado?", vbInformation + vbYesNo) = vbYes Then
        Dim Transac As Boolean
        Dim Params
        Params = Array(Format(Fecinicio.Value, "MMM/dd/yyyy"), ListView1.SelectedItem.Text)
        Transac = FX.CmdTransacciones("DeleteDescuento", Params)
        If Transac Then
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End If
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnPrint_Click()
        Dim RptDesc As New Recordset
        RptDesc.CursorType = 1
        
        FX.LoadRstFromDB "RptResumenDtos", RptDesc, Format(Fecinicio.Value, "MMM-dd-yyyy")
        
        If RptDesc.RecordCount > 0 Then
            CCD.ShowPrinter
            With RptPlanilla
                Set RptPlanilla.DataSource = Nothing
                Set RptPlanilla.DataSource = RptDesc.DataSource
                With .Sections("Sección4")
                    .Controls.Item("LblFecha").Caption = "Descuentos Correspondientes del mes de " & Format(Fecinicio.Value, "MMMM") & " del " & Format(Fecinicio.Value, "yyyy")
                End With
                .Show vbModal, Me
            End With
            RptDesc.Close
        End If
        
        FX.LoadRstFromDB "RptDetalleDescuentos", RptDesc, Format(Fecinicio.Value, "MMM-dd-yyyy")
        
        If RptDesc.RecordCount > 0 Then
            CCD.ShowPrinter
            With RptDetallePlanilla
                Set .DataSource = Nothing
                Set .DataSource = RptDesc.DataSource
                With .Sections("Sección4")
                    .Controls.Item("LblFecha").Caption = "Descuentos Correspondientes del mes de " & Format(Fecinicio.Value, "MMMM") & " del " & Format(Fecinicio.Value, "yyyy")
                End With
                .Show vbModal, Me
            End With
        End If
        
End Sub

Private Sub BtnSearch_Click()
        LeerDescuentos (Format(Fecinicio.Value, "MMM-dd-yyyy"))
End Sub

Private Sub ButtonOffice1_Click()
    Dim ret As Boolean
    ret = FX.Exportar_Excel("", ListView1, ProgressBar1)
End Sub

Private Sub Form_Load()
    FX.ConnectDb activar
    Fecinicio.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FX.ConnectDb Desactivar
End Sub

Private Sub ListView1_Click()
    For i = 1 To ListView1.ColumnHeaders.Count
        Debug.Print ListView1.ColumnHeaders(i).Text; _
                ListView1.ColumnHeaders(i).Width
    Next
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
End Sub
