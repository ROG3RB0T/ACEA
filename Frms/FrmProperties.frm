VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Process"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin Proyecto1.ButtonOffice ButtonOffice1 
      Height          =   1335
      Left            =   8760
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      BackColor       =   12230304
      Caption         =   "Recalcular los saldos"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
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
    Screen.MousePointer = vbHourglass
For i = 1 To ListView1.ListItems.Count
    CalculoSaldos ListView1.ListItems(i).Text, "AHO"
    CalculoSaldos ListView1.ListItems(i).Text, "APT"
    'ListView1.ListItems(i).Bold = True
Next
    Screen.MousePointer = vbDefault
    MsgBox "Proceso finalizado"
'BtnFiltrar_Click
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 18/04/2007 01:02
' Author    : Administrator
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Dim rs As Recordset

    Set rs = New Recordset
    rs.CursorType = 1
    
    FX.LoadRstFromDB "Socios", rs, "", DbTable
    
    If Not rs.EOF Then
        Label1.Caption = "Se encontraron " & rs.RecordCount & " Registros"
        Bar1.Min = 0
        Bar1.Max = rs.RecordCount
        FX.LoadListView rs, ListView1
    End If
    ListView2.Visible = False
End Sub
