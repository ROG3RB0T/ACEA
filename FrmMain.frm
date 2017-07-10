VERSION 5.00
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOCIOS"
   ClientHeight    =   7035
   ClientLeft      =   975
   ClientTop       =   510
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin pRmFrame.RmFrame RmFrame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   5741
      BorderStyle     =   6
      Caption         =   ""
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   13215
      End
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   1320
      Top             =   360
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   20
      BmpCount        =   9
      CheckBorderColor=   6929919
      SelMenuBorder   =   6929919
      SelMenuBackColor=   8379903
      SelMenuForeColor=   0
      SelCheckBackColor=   14731446
      MenuBorderColor =   13603685
      SeparatorColor  =   -2147483632
      MenuBackColor   =   16109232
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   0
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   -2147483644
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   8421504
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Mask:1          =   16711935
      Key:1           =   "#MnuInicio"
      Key:2           =   "#Mnulogout"
      Key:3           =   "#MnuLogin"
      Mask:4          =   16711935
      Key:4           =   "#Sep1"
      Key:5           =   "#MnuForm"
      Mask:6          =   16711935
      Key:6           =   "#MnuCtas"
      Mask:7          =   16711935
      Key:7           =   "#MnuSaldos"
      Key:8           =   "#MnuSalir"
      Mask:9          =   8158332
      Key:9           =   "#MnuPpal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuPpal 
      Caption         =   "Inicio"
      Begin VB.Menu MnuInicio 
         Caption         =   "Ingresar"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuInicio 
         Caption         =   "Desconectarse"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuPpalSocios 
      Caption         =   "Socios"
      Index           =   0
      Begin VB.Menu MnuSocios 
         Caption         =   "Listado Socios"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MnuSocios 
         Caption         =   "Cuentas Socios"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu MnuSocios 
         Caption         =   "Saldos Cuentas"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu MnuSocios 
         Caption         =   "Capitalización de Cuentas"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu MnuSocios 
         Caption         =   "Aplicar pagos en Planilla"
         Enabled         =   0   'False
         Index           =   4
      End
   End
   Begin VB.Menu MnuPpalDtos 
      Caption         =   "Descuentos"
      Begin VB.Menu MnuDtos 
         Caption         =   "Planilla de descuentos"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MnuDtos 
         Caption         =   "Tabla de descuentos"
         Enabled         =   0   'False
         Index           =   1
      End
   End
   Begin VB.Menu MnuPpalPtamos 
      Caption         =   "Préstamos"
      Begin VB.Menu MnuPtamos 
         Caption         =   "Apertura de Préstamos"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MnuPtamos 
         Caption         =   "Consulta Préstamo"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu MnuPtamos 
         Caption         =   "Pago a Préstamo"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu MnuPtamos 
         Caption         =   "Aplicar Préstamos en planilla"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu MnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu MnuTools 
         Caption         =   "Mantenimiento Usuarios"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MnuchgPWD 
         Caption         =   "Cambio de Password"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarBtnLogin_Click()
    FrmLogin.Show vbModal, Me
End Sub

Private Sub ButtonOffice1_Click()
Form1.Show
End Sub

Private Sub Form_Load()
 Label1.Caption = ""
 Label1.Caption = Label1.Caption & vbCrLf & ""
 Label1.Caption = Label1.Caption & vbCrLf & "Sistema Control de Cuentas de Socios"
 Label1.Caption = Label1.Caption & vbCrLf & "Ver. " & App.Major & "." & App.Minor & App.Revision
 Label1.Caption = Label1.Caption & vbCrLf & Date
End Sub

Private Sub Mnuapptmo_Click()
    FrmAperturaPtamo.Show vbModal, Me
End Sub

Private Sub Mnucap_Click()
    FrmInteres.Show vbModal, Me
End Sub

Private Sub MnuConsulta_Click()
    FrmConsultaPtamo.Show vbModal, Me
End Sub

Private Sub MnuCtas_Click()
    FrmCtas.Show vbModal, Me
End Sub

Private Sub MnuForm_Click()
    FrmListSocios.Show vbModal, Me
End Sub
Private Sub MnuMttoUser_Click()
    FrmMttoUsuario.Show vbModal, Me
End Sub

Private Sub MnuchgPWD_Click()
    FrmChangePWD.Show vbModal, Me
End Sub

Private Sub MnuDtos_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmPlanilla.Show vbModal, Me
            
        Case 1
            FrmListServicios.Show vbModal, Me
    End Select
End Sub

Private Sub MnuInicio_Click(Index As Integer)
    Select Case Index
        Case 0 'login
            FrmLogin.Show vbModal, Me
        Case 1
            For i = 0 To MnuSocios.Count - 1
                MnuSocios(i).Enabled = False
            Next
            For i = 0 To MnuDtos.Count - 1
                MnuDtos(i).Enabled = False
            Next
            For i = 0 To MnuPtamos.Count - 1
                MnuPtamos(i).Enabled = False
            Next
            For i = 0 To MnuTools.Count - 1
                MnuTools(i).Enabled = False
            Next
            MnuchgPWD.Enabled = False
    End Select
End Sub

Private Sub MnuPagoPtamo_Click()
    FrmPagoPtamo.Show vbModal, Me
End Sub

Private Sub mnuplanilla_Click()
    FrmPlanilla.Show vbModal, Me
End Sub
Private Sub MnuPtamoPlanilla_Click()
    FrmPlanillaPtamo.Show vbModal, Me
End Sub

Private Sub MnuSaldos_Click()
    FrmSaldosCta.Show vbModal, Me
End Sub

Private Sub MnuPtamos_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmAperturaPtamo.Show vbModal, Me
        Case 1
            FrmConsultaPtamo.Show vbModal, Me
        Case 2
            FrmPagoPtamo.Show vbModal, Me
        Case 3
            FrmPlanillaPtamo.Show vbModal, Me
    End Select
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub MnuSociosPlanilla_Click()
    FrmDescSocios.Show vbModal, Me
End Sub

Private Sub mnutabla_Click()
    FrmListServicios.Show vbModal, Me
End Sub

Private Sub MnuSocios_Click(Index As Integer)
    Select Case Index
        Case 0 'Listado Socios
            FrmListSocios.Show vbModal, Me
        Case 1 'Cuentas Socios
            FrmCtas.Show vbModal, Me
        Case 2 'Saldos Cuentas
            FrmSaldosCta.Show vbModal, Me
        Case 3 'Capitalización
            FrmInteres.Show vbModal, Me
        Case 4 'planilla socios
            FrmDescSocios.Show vbModal, Me
    End Select
End Sub

Private Sub MnuTools_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmMttoUsuario.Show vbModal, Me
    End Select
End Sub
