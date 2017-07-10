VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form FrmAperturaPtamo 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Apertura/Consulta de Préstamos"
   ClientHeight    =   4245
   ClientLeft      =   810
   ClientTop       =   420
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   3750
      Width           =   7830
      _ExtentX        =   13811
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
         TabIndex        =   26
         Top             =   50
         Width           =   3135
         _ExtentX        =   5530
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
         Picture         =   "FrmAperturaPtamo.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnOk 
            Height          =   405
            Left            =   480
            TabIndex        =   28
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   ""
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":0462
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnCancel 
            Height          =   405
            Left            =   960
            TabIndex        =   29
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   ""
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":0BDC
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnEdit 
            Height          =   405
            Left            =   1440
            TabIndex        =   30
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   ""
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":1176
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   2400
            TabIndex        =   32
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":1510
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnDelete 
            Height          =   405
            Left            =   1920
            TabIndex        =   31
            Top             =   15
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   714
            BackColor       =   14522474
            Caption         =   ""
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":18AA
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnNuevo 
            Height          =   405
            Left            =   15
            TabIndex        =   27
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
            PicAlign        =   0
            PicNormal       =   "FrmAperturaPtamo.frx":2024
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
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
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.CheckBox ChkActivo 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Activo"
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin TxNumOcx.TxNum TxInteres 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Locked          =   -1  'True
         Text            =   "0.000"
         Value           =   "0.000"
         Numdec          =   3
         Moneda          =   ""
      End
      Begin TxNumOcx.TxNum TxMontoPtamo 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Locked          =   -1  'True
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Número de Prestamo"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Codigo Socio"
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
      Begin TextBoxWinXP.TextboxXP TxNombres 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin TextBoxWinXP.TextboxXP TxCod_Socio 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin TextBoxWinXP.TextboxXP TxApellidos 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Nombres"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Apellidos"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Monto Ptamo."
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
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
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
      Begin LabelDegradado.LabelDegrade LabelDegrade7 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Plazo"
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
      Begin TextBoxWinXP.TextboxXP TxPlazo 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin LabelDegradado.LabelDegrade LabelDegrade9 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "Monto Cuota"
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
      Begin TxNumOcx.TxNum TxCuota 
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Locked          =   -1  'True
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin Proyecto1.ButtonOffice BtnSearchSocio 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14457180
         Caption         =   ""
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
         PicNormal       =   "FrmAperturaPtamo.frx":28FE
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
         State           =   3
      End
      Begin Proyecto1.ButtonOffice BtnSearchPtmo 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14457180
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
         PicNormal       =   "FrmAperturaPtamo.frx":3078
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade8 
         Height          =   285
         Left            =   4080
         TabIndex        =   15
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
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
         Text            =   "Ingresado"
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
      Begin TextBoxWinXP.TextboxXP TxFecOtorgado 
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin LabelDegradado.LabelDegrade LabelDegrade10 
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
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
         Text            =   "Estado"
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
      Begin TextBoxWinXP.TextboxXP TxEstadoPtmo 
         Height          =   285
         Left            =   4920
         TabIndex        =   18
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   9655840
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
      Begin VB.Label Label1 
         Caption         =   "Meses"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   2520
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmAperturaPtamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tipotx As TipoTransaccion
Sub LimpiarTx()
    TxCod_Socio.Text = ""
    TxNombres.Text = ""
    TxApellidos.Text = ""
    TxMontoPtamo.Value = 0
    TxInteres.Value = 0
    TxPlazo.Text = ""
    TxCuota.Value = 0
    TxFecOtorgado.Text = ""
    TxEstadoPtmo.Text = ""
End Sub
Sub lockTx(Optional Bloqueo As Boolean = True)
    TxCod_Socio.Locked = Bloqueo
    'TxNombres.Locked = Bloqueo
    'TxApellidos.Locked = Bloqueo
    TxMontoPtamo.Locked = Bloqueo
    TxInteres.Locked = Bloqueo
    TxPlazo.Locked = Bloqueo
    TxCuota.Locked = Bloqueo
    BtnSearchSocio.Enabled = Not Bloqueo
End Sub
Sub Lockbtns(Optional Bloqueo As Boolean = True)

    BtnNuevo.Enabled = Bloqueo
    BtnOk.Enabled = Not Bloqueo
    BtnCancel.Enabled = Not Bloqueo
    BtnEdit.Enabled = Not Bloqueo
    BtnDelete.Enabled = Not Bloqueo
    
End Sub
Sub CargarSocio(Codigo As String)
On Error Resume Next
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    Strsql = "Select * from socios where Cod_Socio = '" & Codigo & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxCod_Socio.Text = rssocio("Cod_Socio").Value
        TxNombres.Text = rssocio("Nombres").Value
        TxApellidos.Text = rssocio("Apellidos").Value
        If rssocio("Estado").Value Then
            ChkActivo.Value = 1
        Else
            ChkActivo.Value = 0
            BtnOk.Enabled = False
            MsgBox "El socio no esta activo"
        End If
    Else
        TxNombres.Text = ""
        TxApellidos.Text = ""
        MsgBox "No se encuentra el número de socio"
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub
Sub CargarPtamo(NumPtamo As String)
On Error Resume Next
    LimpiarTx
    Dim RsPtamo As Recordset
    Dim Strsql As String
    Set RsPtamo = New Recordset
    
    RsPtamo.CursorType = adOpenKeyset
    Strsql = "Select * from Prestamos Where NumPtamo = '" & NumPtamo & "'"
    RsPtamo.Open Strsql, CN
        If Not RsPtamo.EOF Then
            TxMontoPtamo.Value = RsPtamo("MontoPtamo").Value
            TxInteres.Value = RsPtamo("Interes").Value
            TxPlazo.Text = RsPtamo("Plazo").Value
            TxCuota.Value = RsPtamo("MontoCuota").Value
            TxFecOtorgado.Text = RsPtamo("FechaOtorgado").Value
            TxEstadoPtmo.Text = RsPtamo("EstadoPtamo").Value
            CargarSocio RsPtamo("Cod_Socio").Value
        End If
    RsPtamo.Close
    Set RsPtamo = Nothing
End Sub
Sub AddPtamo()
On Error Resume Next
    Dim Transac As Boolean
    Dim Paramets
    Paramets = Array(Trim(TxNumPtamo.Text), Trim(TxNombres.Text), _
                    Trim(TxApellidos.Text), TxMontoPtamo.Value, _
                    TxPlazo.Text, TxInteres.Value, _
                    TxCod_Socio.Text, "MESES", TxCuota.Value)
    Transac = FX.CmdTransacciones("AddPrestamo", Paramets)
    If Transac Then
        MsgBox "Se ha procesado la información satisfactoriamente"
    End If
    Tipotx = NoTransac
End Sub
Sub EditPtamo()
On Error Resume Next
    Dim Transac As Boolean
    Dim Parameters
    Paramets = Array(Trim(TxNombres.Text), _
                    Trim(TxApellidos.Text), TxMontoPtamo.Value, _
                    TxPlazo.Text, TxInteres.Value, _
                    TxCod_Socio.Text, TxCuota.Value, TxNumPtamo.Text)
    Transac = FX.CmdTransacciones("QryEditPtamo", Paramets)
    If Transac Then
        MsgBox "Se ha procesado la información satisfactoriamente"
    End If
    Tipotx = NoTransac
End Sub

Private Sub BtnDelete_Click()
    If MsgBox("Si elimina el Pretamo " & TxNumPtamo.Text & " , se eliminará todo el historioal" & vbCrLf & _
            "¿Desea Continuar?", vbYesNo) = vbYes Then
        
        If (FX.CmdTransacciones("DeletePtamo", TxNumPtamo.Text)) Then
            MsgBox "Se ha eliminado el Registro"
            LimpiarTx
        End If
    End If
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnOk_Click()

    If Tipotx = Agregarnuevo Then
        AddPtamo
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
    ElseIf Tipotx = EditarExistente Then
        EditPtamo
        BtnEdit.Enabled = True
        BtnDelete.Enabled = True
    End If
    BtnOk.Enabled = False
    BtnCancel.Enabled = False
    BtnNuevo.Enabled = True
    lockTx
End Sub

Private Sub TxCod_Socio_Change()
    If TxCod_Socio.Text > "" And Tipotx = NoTransac Then
        BtnEdit.Enabled = True
        BtnDelete.Enabled = True
    Else
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
    End If
End Sub

Private Sub TxCod_Socio_LostFocus()
    CargarSocio TxCod_Socio.Text
End Sub
Private Sub Form_Load()

    FX.ConnectDb activar
    lockTx True
    Lockbtns True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If CN.State Then
        FX.ConnectDb Desactivar
    End If
End Sub

Private Sub TxNumPtamo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub

Private Sub TxNumPtamo_LostFocus()
    CargarPtamo TxNumPtamo.Text
End Sub

Private Sub TxPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub
Private Sub BtnCancel_Click()
    If Tipotx = EditarExistente Then
        lockTx
        BtnEdit.Enabled = True
        BtnDelete.Enabled = True
    ElseIf Tipotx = Agregarnuevo Then
        LimpiarTx
        lockTx
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
    End If
        BtnOk.Enabled = False
        BtnCancel.Enabled = False
        BtnNuevo.Enabled = True
End Sub

Private Sub BtnEdit_Click()
    lockTx False
    BtnOk.Enabled = True
    BtnCancel.Enabled = True
    BtnNuevo.Enabled = False
    BtnEdit.Enabled = False
    BtnDelete.Enabled = False
    Tipotx = EditarExistente
End Sub

Private Sub BtnNuevo_Click()
    lockTx False
    LimpiarTx
    BtnOk.Enabled = True
    BtnCancel.Enabled = True
    BtnNuevo.Enabled = False
    BtnEdit.Enabled = False
    BtnDelete.Enabled = False
    Tipotx = Agregarnuevo
End Sub
Private Sub BtnSearchPtmo_Click()
    CargarPtamo TxNumPtamo.Text
End Sub
Private Sub BtnSearchSocio_Click()
FrmSelSocios.Show vbModal, Me
If VAR_COD_SOCIO > "" Then
    CargarSocio VAR_COD_SOCIO
    TxCod_Socio.Text = VAR_COD_SOCIO
End If
End Sub
