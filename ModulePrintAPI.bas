Attribute VB_Name = "ModulePrintAPI"
' Constantes pour la Plateforme Système
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32s = 0

' Global constants for Win32 API
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40

Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&

Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

'Custom Global Constants
Public Const DLG_PRINT = 0
Public Const DLG_PRINTSETUP = 1

'type definitions:
Type PRINTDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long

        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Type DEVNAMES_TYPE
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Type DEVMODE_TYPE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'API declarations:
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function PrintDialog Lib "comdlg32.dll" _
   Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

Public Declare Sub CopyMemory Lib "Kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function GlobalLock Lib "Kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "Kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "Kernel32" _
   (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "Kernel32" _
   (ByVal hMem As Long) As Long

'*******************************************************************
' Détermine si la plateforme est NT
'*******************************************************************
Public Function PlateformeNT() As Boolean
    Dim VersionInfo As OSVERSIONINFO
    
    VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
    If GetVersionEx(VersionInfo) Then
      PlateformeNT = (VersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
    End If
End Function

'*******************************************************************
' Appel des propriétés imprimantes par API(dll) si la plateforme = NT
'*******************************************************************
Public Sub ShowPrinter(frmOwner As Form, Optional PrintFlags As Integer)

    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE
        
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hwnd

    PrintDlg.flags = PrintFlags

    'Set the current orientation, duplex, papersize, etc... setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    'On initialize avec les valeurs du PRINTER par défaut
    DevMode.dmFields = DM_ORIENTATION Or DM_COPIES Or DM_DUPLEX Or DM_PAPERSIZE Or DM_COLOR Or DM_SCALE Or DM_PRINTQUALITY
    On Error Resume Next
    With Printer
        DevMode.dmOrientation = .Orientation
        DevMode.dmCopies = .Copies
        DevMode.dmDuplex = .Duplex
        DevMode.dmPaperSize = .PaperSize
        DevMode.dmColor = .ColorMode
        DevMode.dmScale = .Zoom
        DevMode.dmPrintQuality = .PrintQuality
    End With
    On Error GoTo 0
    
    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
       GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    With Printer
        DevName.extra = .DriverName & Chr(0) & _
        .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
        GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
            CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
            CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
            InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
               If TronqueChaine(UCase(objPrinter.DeviceName), CCHDEVICENAME - 1) = NewPrinterName Then
                    Set Printer = objPrinter
                    Exit For
               End If
            Next
        End If
        On Error Resume Next

        'Set printer object properties according to selections made
        'by user
        DoEvents
        With Printer
            .Copies = DevMode.dmCopies
            .Duplex = DevMode.dmDuplex
            .Orientation = DevMode.dmOrientation
            .PaperSize = DevMode.dmPaperSize
            .ColorMode = DevMode.dmColor
            .Zoom = DevMode.dmScale
            .PrintQuality = DevMode.dmPrintQuality
        End With
        On Error GoTo 0
    End If

End Sub

Public Function TronqueChaine(ByVal chaine As String, ByVal LgChaine As Integer)
  If Len(chaine) < LgChaine Then LgChaine = Len(chaine)
  TronqueChaine = Left(UCase(chaine), LgChaine)
End Function

