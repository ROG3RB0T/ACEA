VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   3600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Proyecto1.ButtonOffice CmdAceptar 
      Default         =   -1  'True
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BackColor       =   14457180
      Caption         =   "Ingresar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicNormal       =   "FrmLogin.frx":058A
      PicSize         =   5
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin TextBoxWinXP.TextboxXP TxUsuario 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      Text            =   ""
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
   Begin TextBoxWinXP.TextboxXP TxPwd 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      Text            =   ""
      PasswordChar    =   "*"
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
   Begin Proyecto1.ButtonOffice CmdCancel 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BackColor       =   14457180
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicNormal       =   "FrmLogin.frx":0924
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
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
      Text            =   "Usuario"
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
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
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
      Text            =   "Password"
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
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   1482
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
      Picture         =   "FrmLogin.frx":0EBE
      PictureSize     =   99
      PictureWidth    =   300
      PictureHeight   =   56
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   615
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   2535
      End
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade3 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
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
      Text            =   "Base de Datos:"
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
   Begin TextBoxWinXP.TextboxXP DbPatch 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   503
      Text            =   ""
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
   Begin Proyecto1.ButtonOffice BtnDb 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   2040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BackColor       =   14457180
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
    Dim logok As Boolean
    logok = False
    logok = FX.Login(TxUsuario.Text, TxPwd.Text)
    If logok Then
        FX.CargarPermisos TxUsuario.Text
        Unload Me
    Else
        MsgBox "No fue posible validar los datos ingresados, por favor verifique"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Ingrese su Usuario y " & vbCrLf & _
                    "Password para ingresar"
                    
    DbPatch.Text = FX.LeeINI(INIFILE, "DATABASE", "PATCH")
End Sub
Private Sub BtnDb_Click()
    Dialog1.FileName = "*.mdb"
    Dialog1.Filter = "DbSys.mdb"
    Dialog1.ShowOpen
    DbPatch.Text = Dialog1.FileName
    FX.GrabaINI INIFILE, "DATABASE", "PATCH", DbPatch.Text
    'DBFILE = DbPatch.Text
    LOADPARAMETROSSYS
End Sub
