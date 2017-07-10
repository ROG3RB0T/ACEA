VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Begin VB.Form FrmSelSocios 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Socios"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre del Socio"
         Object.Width           =   7056
      EndProperty
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
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
      Begin pRmFrame.RmFrame RmFrame1 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   873
         BorderStyle     =   6
         Caption         =   "Búsqueda"
         CaptionAlign    =   1
         CaptionAutoCenter=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionShadow   =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmSelSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim RsSel As New Recordset
    RsSel.CursorType = 1
    FX.LoadRstFromDB "RptListSocios", RsSel, ""
    
    If RsSel.RecordCount > 0 Then
        FX.LoadListView RsSel, ListView1
    Else
        Unload Me
    End If
    ListView1.SortKey = 1 'Ordenar por apellidos(Default)
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
     If ListView1.SortOrder = lvwAscending Then
       ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
End Sub

Private Sub ListView1_DblClick()
    VAR_COD_SOCIO = ""
    VAR_COD_SOCIO = ListView1.SelectedItem.Text
    If VAR_COD_SOCIO > "" Then Unload Me
End Sub
