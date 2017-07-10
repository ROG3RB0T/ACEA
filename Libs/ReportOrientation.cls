VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'// ------------------------------------------------------------
'// CLASS       : ReportOrientation
'// DESCRIPTION : Changes printing orientation by API
'// AUTHOR      : Phil Weber
'// LAST UPDATE : -
'// SOURCE      : VBPJ, MAY 1999,
'//               «The Case of the Disoriented DataReport»
'// ------------------------------------------------------------
Option Explicit

Public Enum Enum_ReportOriention
    OrientUndefined = 0
    OrientPortrait = vbPRORPortrait
    OrientLandscape = vbPRORLandscape
End Enum

Private Type DEVMODE
    dmDeviceName    As String * 32
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer
    dmFields        As Long
    dmOrientation   As Integer
    dmPaperSize     As Integer
    dmPaperLength   As Integer
    dmPaperWidth    As Integer
    dmScale         As Integer
    dmCopies        As Integer
    dmDefaultSource As Integer
    dmPrintQuality  As Integer
    dmColor         As Integer
    dmDuplex        As Integer
    dmYResolution   As Integer
    dmTTOption      As Integer
    dmCollate       As Integer
    dmFormName      As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel    As Integer
    dmPelsWidth     As Long
    dmPelsHeight    As Long
    dmDisplayFlags  As Long
    dmDisplayFrequency As Long
End Type

Private Type PRINTER_DEFAULTS
    pDataType       As String
    pDevMode        As Long
    DesiredAccess   As Long
End Type

Private Type PRINTER_INFO_2
    pServerName     As Long
    pPrinterName    As Long
    pShareName      As Long
    pPortName       As Long
    pDriverName     As Long
    pComment        As Long
    pLocation       As Long
    pDevMode        As Long
    pSepFile        As Long
    pPrintProcessor As Long
    pDataType       As Long
    pParameters     As Long
    pSecurityDescriptor As Long
    Attributes      As Long
    Priority        As Long
    DefaultPriority As Long
    StartTime       As Long
    UntilTime       As Long
    Status          As Long
    cJobs           As Long
    AveragePPM      As Long
End Type

Private Const DM_IN_BUFFER      As Long = 8
Private Const DM_OUT_BUFFER     As Long = 2
Private Const DM_ORIENTATION    As Long = &H1
Private Const NULLPTR           As Long = 0&

Private Const PRINTER_ACCESS_ADMINISTER  As Long = &H4
Private Const PRINTER_ACCESS_USE         As Long = &H8
Private Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Private Const PRINTER_ALL_ACCESS         As Long = (STANDARD_RIGHTS_REQUIRED Or _
                                                   PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long _
        )
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" ( _
        ByVal pPrinterName As String, phPrinter As Long, pDefault As Any _
        ) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" ( _
        ByVal hPrinter As Long _
        ) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" ( _
        ByVal hWnd As Long, ByVal hPrinter As Long, _
        ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, _
        ByVal fMode As Long _
        ) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" ( _
        ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, _
        ByVal cbBuf As Long, pcbNeeded As Long _
        ) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" ( _
        ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long _
        ) As Long

Public Function SetPrinterOrientation( _
       ByVal eOrientation As Enum_ReportOriention _
       ) As Boolean

    Dim bDevMode()      As Byte
    Dim bPrinterInfo2() As Byte
    Dim hPrinter        As Long
    Dim lResult         As Long
    Dim nSize           As Long
    Dim sPrnName        As String
    Dim dm              As DEVMODE
    Dim pd              As PRINTER_DEFAULTS
    Dim pi2             As PRINTER_INFO_2

    ' Get device name of default printer
    sPrnName = Printer.DeviceName
    ' PRINTER_ALL_ACCESS required under NT, because we're going to call SetPrinter
    pd.DesiredAccess = PRINTER_ALL_ACCESS

    ' Get a handle to the printer.
    If OpenPrinter(sPrnName, hPrinter, pd) Then
        ' Get number of bytes requires for PRINTER_INFO_2 structure
        Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
        ' Create a buffer of the required size
        ReDim bPrinterInfo2(1 To nSize) As Byte
        ' Fill buffer with structure
        lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
        ' Copy fixed portion of structure into VB Type variable
        Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))

        ' Get number of bytes requires for DEVMODE structure
        nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
        ' Create a buffer of the required size
        ReDim bDevMode(1 To nSize)

        ' If PRINTER_INFO_2 points to a DEVMODE structure, copy it into our buffer
        If pi2.pDevMode Then
           Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
        Else
           ' Otherwise, call DocumentProperties to get a DEVMODE structure
           Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
        End If

        ' Copy fixed portion of structure into VB Type variable
        Call CopyMemory(dm, bDevMode(1), Len(dm))
        With dm
            ' Set new orientation
            .dmOrientation = eOrientation
            .dmFields = DM_ORIENTATION
        End With
        ' Copy our Type back into buffer
        Call CopyMemory(bDevMode(1), dm, Len(dm))
        ' Set new orientation
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)

        ' Point PRINTER_INFO_2 at our modified DEVMODE
        pi2.pDevMode = VarPtr(bDevMode(1))
        ' Set new orientation system-wide
        lResult = SetPrinter(hPrinter, 2, pi2, 0&)

        ' Clean up and exit
        Call ClosePrinter(hPrinter)
        SetPrinterOrientation = True
    Else
        SetPrinterOrientation = False
    End If
End Function

Public Function GetPrinterOrientation( _
      DeviceName As String, _
      hDC As Long _
    ) As Enum_ReportOriention
    
    Dim hPrinter    As Long
    Dim nSize       As Long
    Dim pDevMode    As DEVMODE
    Dim aDevMode()  As Byte
   
    If OpenPrinter(DeviceName, hPrinter, NULLPTR) Then
       nSize = DocumentProperties( _
               NULLPTR, hPrinter, DeviceName, NULLPTR, NULLPTR, 0 _
               )
       ReDim aDevMode(1 To nSize)
       nSize = DocumentProperties( _
               NULLPTR, hPrinter, DeviceName, aDevMode(1), NULLPTR, DM_OUT_BUFFER _
               )
       Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
       GetPrinterOrientation = pDevMode.dmOrientation
       Call ClosePrinter(hPrinter)
    Else
       GetPrinterOrientation = OrientUndefined
    End If
End Function

