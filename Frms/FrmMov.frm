VERSION 5.00
Object = "{ACC6F197-D72E-4FCC-ACC2-1E6C49D008B9}#5.0#0"; "TxNumOcx.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMov 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos de Cuentas"
   ClientHeight    =   4575
   ClientLeft      =   390
   ClientTop       =   465
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame4 
      Height          =   3855
      Left            =   45
      TabIndex        =   4
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6800
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
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39206
      End
      Begin VB.OptionButton OptAhorros 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Ahorros"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptAportaciones 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Aportaciones"
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxComentario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1560
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2400
         Width           =   5175
      End
      Begin TxNumOcx.TxNum TxMonto 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   2040
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
         Text            =   "0.00"
         Value           =   "0.00"
         BorderColor     =   9655840
         Moneda          =   ""
      End
      Begin TextBoxWinXP.TextboxXP TxDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
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
      Begin VB.ComboBox CmbTipoMov 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade8 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Tipo de Mov."
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
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Descripción"
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
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
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
         Left            =   1560
         TabIndex        =   0
         Top             =   120
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
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade4 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Nombre del Socio"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Afecta"
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
      Begin TextBoxWinXP.TextboxXP TxAfecta 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
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
      Begin LabelDegradado.LabelDegrade LabelDegrade5 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Monto($)"
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Comentario"
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
         Left            =   3360
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Fecha"
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
      Begin MSComCtl2.DTPicker DTFechaCap 
         Height          =   285
         Left            =   4920
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMMM/yyyy"
         Format          =   16449539
         CurrentDate     =   39206
      End
      Begin LabelDegradado.LabelDegrade lblfeCap 
         Height          =   285
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Fecha a capitalizar"
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
      Begin VB.Label LblSaldo 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin pRmFrame.RmFrame RmFrame2 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   15
      Top             =   4095
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   847
      BorderStyle     =   2
      BorderWidth     =   0
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
      GradientColor2  =   16640471
      BackgroundType  =   1
      ShadowOffsetX   =   10
      ShadowColor     =   0
      Begin pRmFrame.RmFrame RmFrame3 
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   5055
         _ExtentX        =   8916
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
         Picture         =   "FrmMov.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnCancel 
            Height          =   395
            Left            =   560
            TabIndex        =   17
            Top             =   8
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   688
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
            PicNormal       =   "FrmMov.frx":0462
            PicOpacity      =   0.65
            PicSize         =   5
            PicSizeH        =   20
            PicSizeW        =   20
         End
         Begin Proyecto1.ButtonOffice BtnSave 
            Height          =   400
            Left            =   40
            TabIndex        =   18
            Top             =   8
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
            PicNormal       =   "FrmMov.frx":0E74
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   16
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnSearch_Click()
FrmSelSocios.Show vbModal, Me
If VAR_COD_SOCIO > "" Then
    CargarSocio VAR_COD_SOCIO
    TxCod_Socio.Text = VAR_COD_SOCIO
End If
End Sub
Function CapPorFecha(Txcod As String, Txtipo As String, Txmes As String)
        Dim TxMontoCap As Double
        Dim Rscap As Recordset
        Dim Params
        Set Rscap = New Recordset
        Rscap.CursorType = 1
        'Txmes = Format(Txmes, "MMM-yyyy")
        'Params = Array(Txcod, Txtipo, Txmes, VAR_CAP)
        Txmes = Format(Txmes, "MMM/yyyy")
        Params = Array(Txcod, Txtipo, Txmes, VAR_CAP)
        FX.LoadRstFromDB "QryCapitalizacionporsocio", Rscap, Params
        If Rscap.RecordCount > 0 Then
            TxMontoCap = Rscap("Capitalizacion").Value
            CapPorFecha = TxMontoCap
        Else
            CapPorFecha = 0
        End If
        Rscap.Close
        Set Rscap = Nothing
End Function

Sub ValidarSaldoCta(NumCta As String)
    Dim TxOk As Boolean
    Dim Nvosaldo As Double
    Dim SaldoCta As Double
    Dim Paramet
    Dim TipoCta As String
    Dim RstVal As New Recordset
    RstVal.CursorType = adOpenKeyset
        If OptAhorros.Value = True Then
            TipoCta = "AHO"
        Else
            TipoCta = "APT"
        End If
    FX.LoadRstFromDB "QryEvalLastSaldo", RstVal, Array(NumCta, TipoCta)
    SaldoCta = 0
    TxComentario.Text = Trim(TxComentario.Text)
    If RstVal.RecordCount > 0 Then
        SaldoCta = RstVal("Saldo").Value
    End If
    If TxAfecta.Text = "CARGO" Then
        If TxMonto.Value > SaldoCta Then
            MsgBox "El Saldo de la Cuenta es menor, no se puede realizar la Transacción", vbInformation
            Exit Sub
        ElseIf CmbTipoMov.Text = "CAN" Then
            If Not SaldoCta > 0 Then
                MsgBox "La cuenta tiene saldo cero"
                Exit Sub
            End If
            TxMonto.Value = SaldoCta
            TxMonto.Locked = False
        End If
    
    ElseIf TxAfecta.Text = "ABONO" Then
        If CmbTipoMov.Text = "CAP" Then
            'Función que devuelva el valor de la capitalización en base al
            'saldo de la fecha seleccionada
            TxMonto.Value = CapPorFecha(NumCta, "AHO", DTFechaCap.Value)
        End If
    
    End If
    If CmbTipoMov.Text <> "AHO" And CmbTipoMov.Text <> "APT" Then
        If OptAhorros.Value = True Then
            TipoCta = "AHO"
        Else
            TipoCta = "APT"
        End If
    Else
        TipoCta = CmbTipoMov.Text
    End If
    Select Case TxAfecta.Text
            Case Is = "ABONO"
                Nvosaldo = TxMonto.Value + SaldoCta
                Paramet = Array(Format(DTFecha.Value, "MMMM/dd/yyyy"), CmbTipoMov.Text, 0, TxMonto.Value, Nvosaldo, TxComentario.Text, NumCta, TipoCta)
            Case Else
                Nvosaldo = SaldoCta - TxMonto.Value
                Paramet = Array(Format(DTFecha.Value, "MMMM/dd/yyyy"), CmbTipoMov.Text, TxMonto.Value, 0, Nvosaldo, TxComentario.Text, NumCta, TipoCta)
    End Select
    
    
    TxOk = FX.CmdTransacciones("QryInsertMovCta", Paramet)
    If TxOk Then
        'Txok = False
        'If TipoCta = "AHO" Then TxOk = FX.CmdTransacciones("QryUpdateSaldoSocio", Array(Nvosaldo, NumCta))
        CalculoSaldos NumCta, TipoCta
        If TxOk Then
            MsgBox "Se ha procesado la transacción para la cuenta " & NumCta, vbInformation
        Else
            MsgBox "No se pudo procesar la transacción para la cuenta " & NumCta & ".", vbInformation
        End If
    Else
            MsgBox "No se pudo procesar la transacción para la cuenta " & NumCta & ".", vbInformation
    End If
    Unload Me
    
End Sub
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
            FX.LoadRstFromDB Strqry, RST, Array(Codigo, TipoCtaSoc), DBQuery
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
Sub CargarSocio(Codigo As String)
    Dim rssocio As Recordset
    Dim Strsql As String
    Set rssocio = New Recordset
    rssocio.CursorType = adOpenKeyset
    Strsql = "Select Apellidos & ', ' & Nombres AS NomSocio, Estado, Saldo_Cuenta " & _
            "from socios where Cod_Socio = '" & Codigo & "'"
    rssocio.Open Strsql, CN
    If Not rssocio.EOF Then
        TxNombres.Text = rssocio("Nomsocio").Value
        LblSaldo.Caption = rssocio("Saldo_Cuenta").Value
    End If
    rssocio.Close
    Set rssocio = Nothing
End Sub
Sub CargarTipomov(IdTipo As String)
    Dim Tipomov As New Recordset
    Tipomov.CursorType = adOpenKeyset
    Dim Strsql As String
    Strsql = "Select * from TipoTransaccion where Cod_Trans = '" & IdTipo & "'"
    Tipomov.Open Strsql, CN
    If Not Tipomov.EOF Then
        TxAfecta.Text = Tipomov("Tipo").Value
        TxDesc.Text = Tipomov("Descripcion").Value
    End If
    Tipomov.Close
    Set Tipomov = Nothing
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnSave_Click()
    ValidarSaldoCta TxCod_Socio.Text
End Sub
Private Sub CmbTipoMov_Click()
Dim RsMonto As Recordset
Set RsMonto = New Recordset

    RsMonto.CursorType = 1

    FX.LoadRstFromDB "QryMontosSocios", RsMonto, TxCod_Socio.Text

   CargarTipomov CmbTipoMov.Text
   If CmbTipoMov.Text = "CAN" Then TxMonto.Locked = True Else TxMonto.Locked = False
    If CmbTipoMov.Text = "AHO" Then
        OptAhorros.Value = True
        If RsMonto.RecordCount > 0 Then
            TxMonto.Value = Val(RsMonto("MontoAhorros").Value)
        End If
        lblfeCap.Visible = False
        DTFechaCap.Visible = False
    ElseIf CmbTipoMov.Text = "APT" Then
        OptAportaciones.Value = True
        If RsMonto.RecordCount > 0 Then
            TxMonto.Value = Val(RsMonto("MontoAportacion").Value)
        End If
        lblfeCap.Visible = False
        DTFechaCap.Visible = False
    ElseIf CmbTipoMov.Text = "CAP" Then
    
        lblfeCap.Visible = True
        DTFechaCap.Visible = True
            
    Else
        lblfeCap.Visible = False
        DTFechaCap.Visible = False
        TxMonto.Value = 0
    End If
    RsMonto.Close
    Set RsMonto = Nothing
End Sub

Private Sub CmbTipoMov_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    If VAR_COD_SOCIO > "" Then
        CargarSocio VAR_COD_SOCIO
        TxCod_Socio.Text = VAR_COD_SOCIO
        VAR_COD_SOCIO = ""
    End If
    FX.CmdFillCombos "CmbTipoMov", CmbTipoMov, "", False, False
    DTFechaCap.Value = Date
End Sub

Private Sub TxCod_Socio_LostFocus()
    CargarSocio TxCod_Socio.Text
End Sub
