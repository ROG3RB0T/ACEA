VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCapCtas 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Proceso de Capitalización de Cuentas"
   ClientHeight    =   6135
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
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
      Begin VB.TextBox TxNumPtamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "3"
         Top             =   120
         Width           =   855
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   120
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
         TabIndex        =   4
         Top             =   480
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
         TabIndex        =   5
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39010
      End
      Begin MSComCtl2.DTPicker Fecfinal 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39010
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   120
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
         Text            =   "Interés(%)"
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
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   5640
      Width           =   9090
      _ExtentX        =   16034
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
      GradientColor1  =   16109232
      GradientColor2  =   16640213
      BackgroundType  =   1
      ShadowOffsetX   =   10
      ShadowOffsetY   =   100
      ShadowColor     =   0
      Begin pRmFrame.RmFrame RmFrame3 
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         BorderStyle     =   8
         BorderWidth     =   0
         BorderType      =   2048
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
         GradientColor2  =   14592118
         BackgroundType  =   2
         ShadowOffsetX   =   20
         ShadowOffsetY   =   20
         ShadowColor     =   0
         Picture         =   "FrmCapCtas.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BarBtnLogin 
            Height          =   405
            Left            =   45
            TabIndex        =   2
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14592118
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
            PicAlign        =   0
            PicNormal       =   "FrmCapCtas.frx":0462
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmCapCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TxInteres_Change()

End Sub

Private Sub TxInteres_KeyPress(KeyAscii As Integer)

End Sub

