VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Object = "{D7B4B7D4-F6C3-4494-BFAD-B02E19333C9E}#1.0#0"; "TextBoxWinXP.ocx"
Object = "{CE212AA6-A6B5-4BE8-9EB2-0A77F9DBB0B3}#2.0#0"; "RmFrame.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form FrmServicios 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Listado de Servicios"
   ClientHeight    =   3270
   ClientLeft      =   675
   ClientTop       =   510
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin pRmFrame.RmFrame RmFrame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
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
      Begin VB.TextBox TxDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1080
         Width           =   6255
      End
      Begin TextBoxWinXP.TextboxXP TxCodServicio 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
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
         Locked          =   -1  'True
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade2 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   240
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
         Text            =   "Cod. Servicio"
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
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
         Text            =   "Descripción del Servicio"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "* Nemónico de 5 caracteres"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
   End
   Begin pRmFrame.RmFrame RmFrame1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2775
      Width           =   7350
      _ExtentX        =   12965
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
         Left            =   45
         TabIndex        =   3
         Top             =   30
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
         Picture         =   "FrmServicios.frx":0000
         PictureSize     =   99
         PictureWidth    =   15
         PictureHeight   =   30
         PictureMarginTop=   -1
         Begin Proyecto1.ButtonOffice BtnSave 
            Height          =   405
            Left            =   45
            TabIndex        =   4
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
            PicNormal       =   "FrmServicios.frx":0462
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnCancel 
            Height          =   405
            Left            =   545
            TabIndex        =   9
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
            PicNormal       =   "FrmServicios.frx":0BDC
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
            State           =   3
         End
         Begin Proyecto1.ButtonOffice BtnEdit 
            Height          =   405
            Left            =   1045
            TabIndex        =   10
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
            PicNormal       =   "FrmServicios.frx":1176
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnExit 
            Height          =   405
            Left            =   2045
            TabIndex        =   11
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
            PicNormal       =   "FrmServicios.frx":1510
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
         Begin Proyecto1.ButtonOffice BtnDelete 
            Height          =   405
            Left            =   1545
            TabIndex        =   12
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
            PicNormal       =   "FrmServicios.frx":18AA
            PicOpacity      =   0.85
            PicSize         =   5
            PicSizeH        =   18
            PicSizeW        =   18
         End
      End
   End
End
Attribute VB_Name = "FrmServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tipotx As TipoTransaccion
Sub LoadServicio(CodServicio As String)
    Dim Rsrv As New Recordset
    Rsrv.CursorType = 1
    
    FX.LoadRstFromDB "QryServiciosbyCodigo", Rsrv, CodServicio
    
    If Rsrv.RecordCount > 0 Then
        TxCodServicio.Text = Rsrv("Cod_Servicio").Value
        TxDescripcion.Text = Rsrv("Descripcion").Value
    End If
    
End Sub
Sub EditServicio(CodServ)
    Dim Params
    
On Error Resume Next

    Params = Array(Trim(TxDescripcion.Text), TxCodServicio.Text)
    
    If (FX.CmdTransacciones("EditServicios", Params)) Then
        MsgBox "Se ha editado el Registro", vbInformation
    Else
        MsgBox "No fue posible editar el Registro", vbInformation
    End If
    

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmServicios.EditServicio." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If

End Sub
Sub AddServicio()
    Dim Params
    

On Error Resume Next
    
    Params = Array(Trim(UCase(Left(TxCodServicio.Text, 5))), Trim(TxDescripcion.Text))
    
    If (FX.CmdTransacciones("QryInsertServicios", Params)) Then
        MsgBox "Se insertado el nuevo registro", vbInformation
    Else
        MsgBox "No fue posible insertar el nuevo registro", vbInformation
    End If

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmServicios.AddServicio." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnDelete_Click()
On Error Resume Next

    If MsgBox("Desea Eliminar el Registro Actual?", vbYesNo) = vbYes Then
        If (FX.CmdTransacciones("DeleteServicio", TxCodServicio.Text)) Then
            MsgBox "Se ha eliminiado el Registro", vbInformation
        Else
            MsgBox "No fue posible eliminar el registro", vbInformation
        End If
    End If
        Unload Me
        FrmListServicios.LoadServicios
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " [" & Err.Description & "] en ACEA.FrmServicios.BtnDelete_Click." _
            & vbCrLf & "Si el problema persiste contacte con su Administrador de Sistemas."
End If
        
End Sub

Private Sub BtnEdit_Click()
    TxDescripcion.Locked = False
    Btnedit.Enabled = False
    BtnSave.Enabled = True
    BtnCancel.Enabled = True
        Tipotx = EditarExistente
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnSave_Click()
    Select Case Tipotx
            Case EditarExistente
                EditServicio TxCodServicio.Text
            Case Agregarnuevo
                AddServicio
    End Select
            FrmListServicios.LoadServicios
            Unload Me
End Sub
Public Sub Clear()
    FrmServicios.TxCodServicio.Locked = False
    TxCodServicio.Text = ""
    FrmServicios.TxDescripcion.Locked = False
    TxDescripcion.Text = ""
    FrmServicios.BtnSave.Enabled = True
    FrmServicios.BtnCancel.Enabled = True
    FrmServicios.Btnedit.Enabled = False
    FrmServicios.Tipotx = Agregarnuevo
End Sub
Private Sub Form_Load()
    If VAR_CODSERV > "" Then
        LoadServicio VAR_CODSERV
    Else
        Clear
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VAR_CODSERV = ""
End Sub

