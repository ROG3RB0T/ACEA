VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPagoPtamo 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago a Prestamos"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5318
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
      Begin MSComCtl2.DTPicker TxToday 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39214
      End
      Begin VB.TextBox TxNumPtamo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxInteres 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin TextBoxWinXP.TextboxXP TxNombre 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   240
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
      Begin Proyecto1.ButtonOffice BtnSearch 
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   240
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
         PicNormal       =   "FrmPagoPtamo.frx":0000
         PicSize         =   5
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   285
         Left            =   240
         TabIndex        =   10
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
         Text            =   "Nombre del Cliente"
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
      Begin TxNumOcx.TxNum TxMontoPtamo 
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   5640
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
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
         Left            =   3600
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
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
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
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
         Left            =   5400
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
         TabIndex        =   16
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
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade10 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   1440
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
         Text            =   "Fecha Aplicación"
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
      Begin TextBoxWinXP.TextboxXP FecOtorgado 
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
      Begin LabelDegradado.LabelDegrade LabelDegrade11 
         Height          =   285
         Left            =   240
         TabIndex        =   19
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
         Text            =   "Estdo del Ptamo"
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
      Begin TextBoxWinXP.TextboxXP TxEstado 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         Text            =   ""
         BorderColor     =   12582912
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
      Begin TxNumOcx.TxNum TxSaldoActual 
         Height          =   285
         Left            =   5880
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   4200
         TabIndex        =   22
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
         Text            =   "Saldo Anterior"
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
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   285
         Left            =   240
         TabIndex        =   23
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
         Text            =   "intereses"
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
      Begin TxNumOcx.TxNum TxIntereses 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
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
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade8 
         Height          =   285
         Left            =   4200
         TabIndex        =   24
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
         Text            =   "Abono a Capital"
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
      Begin TxNumOcx.TxNum TxACapital 
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
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
         Text            =   "$0.00"
         Value           =   "$0.00"
      End
      Begin TxNumOcx.TxNum TxNuevoSaldo 
         Height          =   285
         Left            =   5880
         TabIndex        =   25
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
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
      Begin LabelDegradado.LabelDegrade LabelDegrade12 
         Height          =   285
         Left            =   4200
         TabIndex        =   26
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
         Text            =   "Nuevo Saldo"
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
      TabIndex        =   27
      Top             =   3180
      Width           =   8970
      _ExtentX        =   15822
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
         TabIndex        =   28
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "FrmPagoPtamo.frx":077A
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BarBtnLogin 
            Height          =   405
            Left            =   45
            TabIndex        =   4
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
            PicNormal       =   "FrmPagoPtamo.frx":0BDC
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnPrint 
            Height          =   405
            Left            =   520
            TabIndex        =   29
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
            PicNormal       =   "FrmPagoPtamo.frx":15EE
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmPagoPtamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BarBtnLogin_Click()
    Dim Param
    Dim Transac As Boolean
    If TxCuota.Value < 1 Then
        MsgBox "El Monto de la cuota no puede ser cero"
    ElseIf TxSaldoActual.Value = 0 Then
        MsgBox "El Prestamo ya fue cancelado"
    ElseIf (TxIntereses.Value + TxACapital.Value) <> TxCuota.Value Then
        MsgBox "La suma de los intereses y Abono a Captital no coinciden con el monto de la cuota"
    Else
        Param = Array(Format(TxToday.Value, "mmm/dd/yyyy"), TxNumPtamo.Text, TxSaldoActual.Value, TxCuota. _
                        Value, TxIntereses.Value, TxACapital. _
                        Value, TxNuevoSaldo.Value)
        
        Transac = FX.CmdTransacciones("AddCuotaPtamo", Param)
        
        If Transac Then
            MsgBox "Se ha procesado el pago al prestamo"
            If TxNuevoSaldo.Value = 0 Then
                FX.CmdTransacciones "UpdateStadoPtamo", TxNumPtamo.Text
            End If
            'TxNumPtamo.Text = ""
            TxNombre.Text = ""
            TxEstado.Text = ""
            TxMontoPtamo.Value = 0
            TxInteres.Text = ""
            TxPlazo.Text = ""
            TxCuota.Value = 0
            TxIntereses.Value = 0
            TxSaldoActual.Value = 0
            TxACapital.Value = 0
            TxNuevoSaldo.Value = 0
            TxNumPtamo.SetFocus
            
        End If
        'BarBtnLogin.Enabled = False
        'RmFrame2.Enabled = False
    End If
    
End Sub
Sub ConsultaPtamo()
    Dim RsPtamo As New Recordset
    Dim RsHistorial As New Recordset
    RsHistorial.CursorType = 1
    RsPtamo.CursorType = 1
    
    FX.LoadRstFromDB "Qryprestamo", RsPtamo, TxNumPtamo.Text
    
    If RsPtamo.RecordCount > 0 Then
        TxNombre.Text = RsPtamo("Nombre").Value
        FecOtorgado.Text = RsPtamo("FechaOtorgado").Value
        TxMontoPtamo.Value = RsPtamo("MontoPtamo").Value
        TxInteres.Text = RsPtamo("Interes").Value
        TxPlazo.Text = RsPtamo("Plazo").Value & " " & RsPtamo("FormaPago").Value
        TxCuota.Value = RsPtamo("MontoCuota").Value
        TxEstado.Text = RsPtamo("EstadoPtamo").Value
'        Frame1.Enabled = True
    End If
            
    FX.LoadRstFromDB "QryLastSaldoPtamo", RsHistorial, TxNumPtamo.Text
    If TxEstado.Text = "No Cancelado" Then
        If RsHistorial.RecordCount > 0 Then
            TxSaldoActual.Value = RsHistorial("SaldoActual").Value
            'TxIntereses.Value = ((TxSaldoActual.Value * (RsPtamo("Interes").Value / 100))) / Val(RsPtamo("plazo").Value)
            TxIntereses.Value = 0
            'TxACapital.Value = TxCuota.Value - TxIntereses.Value
            TxACapital.Value = 0
            TxNuevoSaldo.Value = TxSaldoActual.Value - TxACapital.Value
        Else
            TxSaldoActual.Value = TxMontoPtamo.Value
            'TxIntereses.Value = (TxCuota.Value * (CDbl(TxInteres.Text) / 100))
            TxIntereses.Value = 0
            'TxACapital.Value = TxCuota.Value - TxIntereses.Value
            TxACapital.Value = 0
            TxNuevoSaldo.Value = TxSaldoActual.Value - TxACapital.Value
        End If
    Else
        RmFrame2.Enabled = False
        BarBtnLogin.Enabled = False
    End If
    
End Sub
Sub calculopago()
    Dim RsHistorial As New Recordset
    RsHistorial.CursorType = 1
    
    FX.LoadRstFromDB "QryLastSaldoPtamo", RsHistorial, TxNumPtamo.Text
    
    If RsHistorial.RecordCount > 0 Then
        TxSaldoActual.Value = RsHistorial("SaldoActual").Value
        TxIntereses.Value = (TxCuota.Value * (CDbl(TxInteres.Text) / 100))
        TxACapital.Value = TxCuota.Value - TxIntereses.Value
        TxNuevoSaldo.Value = TxSaldoActual.Value - TxACapital.Value
    Else
        TxSaldoActual.Value = TxMontoPtamo.Value
'        TxSaldoActual.Value = RsHistorial("SaldoActual").Value
        TxIntereses.Value = (TxCuota.Value * (CDbl(TxInteres.Text) / 100))
        TxACapital.Value = TxCuota.Value - TxIntereses.Value
        TxNuevoSaldo.Value = TxSaldoActual.Value - TxACapital.Value
    End If
End Sub

Private Sub BtnPrint_Click()
    Unload Me
End Sub

Private Sub BtnSearch_Click()
    ConsultaPtamo
End Sub

Private Sub Form_Load()
    FX.ConnectDb activar
    TxToday.Value = Date
End Sub

Private Sub TxACapital_LostFocus()
If TxNumPtamo.Text > "" Then
    TxNuevoSaldo.Value = TxSaldoActual.Value - TxACapital.Value
End If
End Sub

Private Sub TxCuota_LostFocus()
'If TxNumPtamo.Text > "" Then calculopago
End Sub

Private Sub TxIntereses_LostFocus()
'If TxNumPtamo.Text > "" Then calculopago
End Sub

Private Sub TxNumPtamo_LostFocus()
    ConsultaPtamo
End Sub

Private Sub TxPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub

Private Sub TxNumPtamo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub

