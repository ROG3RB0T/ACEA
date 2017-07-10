VERSION 5.00
Begin VB.Form FrmChangePWD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Password de usuarios"
   ClientHeight    =   2160
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1276.199
   ScaleMode       =   0  'User
   ScaleWidth      =   4478.772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txpwd1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "?"
      TabIndex        =   1
      Top             =   480
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txpwd2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "?"
      TabIndex        =   3
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nueva Contraseña"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label lblLabels 
      Caption         =   "Confirmación de Contraseña"
      Height          =   390
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   960
      Width           =   1800
   End
End
Attribute VB_Name = "FrmChangePWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TxOk As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
'    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    If txpwd1.Text = txpwd2.Text Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        txpwd2.Text = FX.Base64Encode(txpwd2.Text)
        TxOk = FX.CmdTransacciones("QryChangePWDUser", Array(txpwd2.Text, USUARIOACTIVO))
        If TxOk Then
            MsgBox "Se ha cambiado la contraseña del usuario " & USUARIOACTIVO & " Satisfactoriamente"
            Unload Me
        End If
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo"
        txpwd1.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
    lblLabels(0).Caption = USUARIOACTIVO
    FX.ConnectDb activar
End Sub
