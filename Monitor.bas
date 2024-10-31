'Monitor显示器
' Copyright (c) 2024 CM.Wang
' Freeware. Use at your own risk.

#include once "Monitor.bi"

Using My.Sys.Forms

Private Function EDSEdwFlags(Index As Integer) As DWORD
	Select Case Index
	Case 0
		Function = EDS_RAWMODE
	Case 1
		Function = EDS_ROTATEDMODE
	End Select
End Function

Private Function CDSdwFlags(Index As Integer) As DWORD
	Select Case Index
	Case 0
		Function = 0
	Case 1
		Function = CDS_FULLSCREEN
	Case 2
		Function = CDS_GLOBAL
	Case 3
		Function = CDS_NORESET
	Case 4
		Function = CDS_RESET
	Case 5
		Function = CDS_SET_PRIMARY
	Case 6
		Function = CDS_TEST
	Case 7
		Function = CDS_UPDATEREGISTRY
	Case 8
		Function = CDS_VIDEOPARAMETERS
	Case 9
		Function = CDS_ENABLE_UNSAFE_MODES
	Case 10
		Function = CDS_DISABLE_UNSAFE_MODES
	End Select
End Function

Private Function ScanOrderWStr(tpy As DISPLAYCONFIG_SCANLINE_ORDERING) ByRef As WString
	Select Case tpy
	Case DISPLAYCONFIG_SCANLINE_ORDERING_UNSPECIFIED
		Return "UNSPECIFIED"
	Case DISPLAYCONFIG_SCANLINE_ORDERING_PROGRESSIVE
		Return "PROGRESSIVE"
	Case DISPLAYCONFIG_SCANLINE_ORDERING_INTERLACED
		Return "INTERLACED"
	Case DISPLAYCONFIG_SCANLINE_ORDERING_INTERLACED_UPPERFIELDFIRST
		Return "UPPERFIELDFIRST"
	Case DISPLAYCONFIG_SCANLINE_ORDERING_INTERLACED_LOWERFIELDFIRST
		Return "LOWERFIELDFIRST"
	Case DISPLAYCONFIG_SCANLINE_ORDERING_FORCE_UINT32
		Return "FORCE_UINT32"
	Case Else
		Return "Unknow"
	End Select
End Function

Private Function OutPutWStr(tpy As DISPLAYCONFIG_VIDEO_OUTPUT_TECHNOLOGY) ByRef As WString
	Select Case tpy
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_OTHER
		Return "OTHER"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_HD15
		Return "HD15(VGA)"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_SVIDEO
		Return "SVIDEO"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_COMPOSITE_VIDEO
		Return "COMPOSITE_VIDEO"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_COMPONENT_VIDEO
		Return "COMPONENT_VIDEO"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_DVI
		Return "DVI"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_HDMI
		Return "HDMI"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_LVDS
		Return "LVDS"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_D_JPN
		Return "D_JPN"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_SDI
		Return "SDI"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_DISPLAYPORT_EXTERNAL
		Return "DISPLAYPORT_EXTERNAL"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_DISPLAYPORT_EMBEDDED
		Return "DISPLAYPORT_EMBEDDED"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_UDI_EXTERNAL
		Return "UDI_EXTERNAL"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_UDI_EMBEDDED
		Return "UDI_EMBEDDED"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_SDTVDONGLE
		Return "SDTVDONGLE"
		'Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_MIRACAST
		'	Return "MIRACAST"
		'Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_INDIRECT_WIRED
		'	Return "INDIRECT_WIRED"
		'Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_INDIRECT_VIRTUAL
		'	Return "INDIRECT_VIRTUAL"
		'Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_DISPLAYPORT_USB_TUNNEL
		'	Return "DISPLAYPORT_USB_TUNNEL"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_INTERNAL
		Return "INTERNAL"
	Case DISPLAYCONFIG_OUTPUT_TECHNOLOGY_FORCE_UINT32
		Return "FORCE_UINT32"
	Case Else
		Return "Unknow"
	End Select
End Function

Private Sub RECT2WStr(ret As tagRECT, txt As TextBox Ptr, ByVal h As String = "")
	txt->AddLine h & "Left:     " & ret.left
	txt->AddLine h & "Right:    " & ret.right
	txt->AddLine h & "Top:      " & ret.top
	txt->AddLine h & "Bottom:   " & ret.bottom
End Sub

Private Sub Path2WStr(PathInfo As DISPLAYCONFIG_PATH_INFO, txt As TextBox Ptr)
	txt->AddLine "FLAGS: " & PathInfo.flags
	txt->AddLine "sourceInfo.adapterId.HighPart: " & PathInfo.sourceInfo.adapterId.HighPart
	txt->AddLine "sourceInfo.adapterId.LowPart: " & PathInfo.sourceInfo.adapterId.LowPart
	txt->AddLine "sourceInfo.id: " & PathInfo.sourceInfo.id
	txt->AddLine "sourceInfo.modeInfoIdx: " & PathInfo.sourceInfo.modeInfoIdx
	txt->AddLine "sourceInfo.statusFlags: " & PathInfo.sourceInfo.statusFlags
	txt->AddLine "targetInfo.adapterId.HighPart: " & PathInfo.targetInfo.adapterId.HighPart
	txt->AddLine "targetInfo.adapterId.LowPart: " & PathInfo.targetInfo.adapterId.LowPart
	txt->AddLine "targetInfo.id: " & PathInfo.targetInfo.id
	txt->AddLine "targetInfo.modeInfoIdx: " & PathInfo.targetInfo.modeInfoIdx
	txt->AddLine "targetInfo.outputTechnology: " & PathInfo.targetInfo.outputTechnology
	txt->AddLine "targetInfo.refreshRate.Denominator: " & PathInfo.targetInfo.refreshRate.Denominator
	txt->AddLine "targetInfo.refreshRate.Numerator: " & PathInfo.targetInfo.refreshRate.Numerator
	txt->AddLine "targetInfo.rotation: " & PathInfo.targetInfo.rotation
	txt->AddLine "targetInfo.scaling: " & PathInfo.targetInfo.scaling
	txt->AddLine "targetInfo.scanLineOrdering: " & PathInfo.targetInfo.scanLineOrdering
	txt->AddLine "targetInfo.statusFlags: " & PathInfo.targetInfo.statusFlags
	txt->AddLine "targetInfo.targetAvailable: " & PathInfo.targetInfo.targetAvailable
	txt->AddLine "outputTechnology=" & OutPutWStr(PathInfo.targetInfo.outputTechnology)
End Sub

Private Sub Mode2WStr(ModeInfo As DISPLAYCONFIG_MODE_INFO, txt As TextBox Ptr)
	txt->AddLine "adapterId.HighPart: " & ModeInfo.adapterId.HighPart
	txt->AddLine "adapterId.LowPart: " & ModeInfo.adapterId.LowPart
	txt->AddLine "id: " & ModeInfo.id
	txt->AddLine "infoType: " & ModeInfo.infoType
	txt->AddLine "sourceMode.Height: " & ModeInfo.sourceMode.height
	txt->AddLine "sourceMode.Width: " & ModeInfo.sourceMode.width
	txt->AddLine "sourceMode.PixelFormat: " & ModeInfo.sourceMode.pixelFormat
	txt->AddLine "sourceMode.position.x: " & ModeInfo.sourceMode.position.x
	txt->AddLine "sourceMode.position.y: " & ModeInfo.sourceMode.position.y
	txt->AddLine "targetMode.targetVideoSignalInfo.activeSize.CX : " & ModeInfo.targetMode.targetVideoSignalInfo.activeSize.cx
	txt->AddLine "targetMode.targetVideoSignalInfo.activeSize.CY: " & ModeInfo.targetMode.targetVideoSignalInfo.activeSize.cy
	txt->AddLine "targetMode.targetVideoSignalInfo.hSyncFreq.Denominator : " & ModeInfo.targetMode.targetVideoSignalInfo.hSyncFreq.Denominator
	txt->AddLine "targetMode.targetVideoSignalInfo.hSyncFreq.Numerator: " & ModeInfo.targetMode.targetVideoSignalInfo.hSyncFreq.Numerator
	txt->AddLine "targetMode.targetVideoSignalInfo.pixelRate: " & ModeInfo.targetMode.targetVideoSignalInfo.pixelRate
	txt->AddLine "targetMode.targetVideoSignalInfo.scanLineOrdering: " & ModeInfo.targetMode.targetVideoSignalInfo.scanLineOrdering
	txt->AddLine "targetMode.targetVideoSignalInfo.totalSize.cx: " & ModeInfo.targetMode.targetVideoSignalInfo.totalSize.cx
	txt->AddLine "targetMode.targetVideoSignalInfo.totalSize.yx: " & ModeInfo.targetMode.targetVideoSignalInfo.totalSize.cy
	txt->AddLine "targetMode.targetVideoSignalInfo.vSyncFreq.Denominator: " & ModeInfo.targetMode.targetVideoSignalInfo.vSyncFreq.Denominator
	txt->AddLine "targetMode.targetVideoSignalInfo.vSyncFreq.Numerator: " & ModeInfo.targetMode.targetVideoSignalInfo.vSyncFreq.Numerator
	txt->AddLine "targetMode.targetVideoSignalInfo.videoStandard: " & ModeInfo.targetMode.targetVideoSignalInfo.videoStandard
	txt->AddLine "scanLineOrdering=" & ScanOrderWStr(ModeInfo.targetMode.targetVideoSignalInfo.scanLineOrdering)
End Sub

Private Sub DD2WStr(ddDisplay As DISPLAY_DEVICE, txt As TextBox Ptr)
	txt->AddLine ML("Device ID") & " = " & ddDisplay.DeviceID
	txt->AddLine ML("Device Key") & " = " & ddDisplay.DeviceKey
	txt->AddLine ML("Device Name") & " = " & ddDisplay.DeviceName
	txt->AddLine ML("Device String") & " = " & ddDisplay.DeviceString
	txt->AddLine ML("State Flags") & " = " & ddDisplay.StateFlags
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_ACTIVE , " @ ", "   ") & "DISPLAY_DEVICE_ACTIVE"
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_MIRRORING_DRIVER , " @ ", "   ") & "DISPLAY_DEVICE_MIRRORING_DRIVER"
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_MODESPRUNED , " @ ", "   ") & "DISPLAY_DEVICE_MODESPRUNED"
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE , " @ ", "   ") & "DISPLAY_DEVICE_PRIMARY_DEVICE"
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_REMOVABLE , " @ ", "   ") & "DISPLAY_DEVICE_REMOVABLE"
	txt->AddLine IIf (ddDisplay.StateFlags And DISPLAY_DEVICE_VGA_COMPATIBLE , " @ ", "   ") & "DISPLAY_DEVICE_VGA_COMPATIBLE"
End Sub

Private Function DM2SimpleWStr(dmDevMode As DEVMODE) ByRef As String
	Static a As String
	a = "(" & dmDevMode.dmPelsWidth & " x " & dmDevMode.dmPelsHeight & ") " & dmDevMode.dmBitsPerPel & "Bit, @" & dmDevMode.dmDisplayFrequency & "Hertz, O" & dmDevMode.dmDisplayOrientation
	Return a
End Function

Private Sub DM2WStr(dmDevMode As DEVMODE, txt As TextBox Ptr)
	txt->AddLine IIf(dmDevMode.dmFields And DM_BITSPERPEL , " @ " , "   ")      & "dmBitsPerPel         = " & dmDevMode.dmBitsPerPel
	txt->AddLine IIf(dmDevMode.dmFields And DM_COLLATE , " @ " , "   ")         & "dmCollate            = " & dmDevMode.dmCollate
	txt->AddLine IIf(dmDevMode.dmFields And DM_COLOR , " @ " , "   ")           & "dmColor              = " & dmDevMode.dmColor
	txt->AddLine IIf(dmDevMode.dmFields And DM_COPIES , " @ " , "   ")          & "dmCopies             = " & dmDevMode.dmCopies
	txt->AddLine IIf(dmDevMode.dmFields And DM_DEFAULTSOURCE , " @ " , "   ")   & "dmDefaultSource      = " & dmDevMode.dmDefaultSource
	txt->AddLine "   " & "dmDeviceName         = " & dmDevMode.dmDeviceName
	txt->AddLine IIf(dmDevMode.dmFields And DM_DISPLAYFIXEDOUTPUT , " @ " , "   ")  & "dmDisplayFixedOutput = " & dmDevMode.dmDisplayFixedOutput
	txt->AddLine IIf(dmDevMode.dmFields And DM_DISPLAYFLAGS , " @ " , "   ")        & "dmDisplayFlags       = " & dmDevMode.dmDisplayFlags
	txt->AddLine IIf(dmDevMode.dmFields And DM_DISPLAYFREQUENCY , " @ " , "   ")    & "dmDisplayFrequency   = " & dmDevMode.dmDisplayFrequency
	txt->AddLine IIf(dmDevMode.dmFields And DM_DISPLAYORIENTATION , " @ " , "   ")  & "dmDisplayOrientation = " & dmDevMode.dmDisplayOrientation
	txt->AddLine IIf(dmDevMode.dmFields And DM_DITHERTYPE , " @ " , "   ")          & "dmDitherType         = " & dmDevMode.dmDitherType
	txt->AddLine "   " & "dmDriverExtra        = " & dmDevMode.dmDriverExtra
	txt->AddLine "   " & "dmDriverVersion      = " & dmDevMode.dmDriverVersion
	txt->AddLine IIf(dmDevMode.dmFields And DM_DUPLEX , " @ " , "   ") & "dmDuplex             = " & dmDevMode.dmDuplex
	txt->AddLine "   " & "dmFields             = " & dmDevMode.dmFields ' & ", member to change the display settings.")
	txt->AddLine IIf(dmDevMode.dmFields And DM_FORMNAME , " @ " , "   ")        & "dmFormName           = " & dmDevMode.dmFormName
	txt->AddLine IIf(dmDevMode.dmFields And DM_ICMINTENT , " @ " , "   ")       & "dmICMIntent          = " & dmDevMode.dmICMIntent
	txt->AddLine IIf(dmDevMode.dmFields And DM_ICMMETHOD , " @ " , "   ")       & "dmICMMethod          = " & dmDevMode.dmICMMethod
	txt->AddLine IIf(dmDevMode.dmFields And DM_LOGPIXELS , " @ " , "   ")       & "dmLogPixels          = " & dmDevMode.dmLogPixels
	txt->AddLine IIf(dmDevMode.dmFields And DM_MEDIATYPE , " @ " , "   ")       & "dmMediaType          = " & dmDevMode.dmMediaType
	txt->AddLine IIf(dmDevMode.dmFields And DM_NUP , " @ " , "   ")             & "dmNup                = " & dmDevMode.dmNup
	txt->AddLine IIf(dmDevMode.dmFields And DM_ORIENTATION , " @ " , "   ")     & "dmOrientation        = " & dmDevMode.dmOrientation
	txt->AddLine IIf(dmDevMode.dmFields And DM_PANNINGHEIGHT , " @ " , "   ")   & "dmPanningHeight      = " & dmDevMode.dmPanningHeight
	txt->AddLine IIf(dmDevMode.dmFields And DM_PANNINGWIDTH , " @ " , "   ")    & "dmPanningWidth       = " & dmDevMode.dmPanningWidth
	txt->AddLine IIf(dmDevMode.dmFields And DM_PAPERLENGTH , " @ " , "   ")     & "dmPaperLength        = " & dmDevMode.dmPaperLength
	txt->AddLine IIf(dmDevMode.dmFields And DM_PAPERSIZE , " @ " , "   ")       & "dmPaperSize          = " & dmDevMode.dmPaperSize
	txt->AddLine IIf(dmDevMode.dmFields And DM_PAPERWIDTH , " @ " , "   ")      & "dmPaperWidth         = " & dmDevMode.dmPaperWidth
	txt->AddLine IIf(dmDevMode.dmFields And DM_PELSHEIGHT , " @ " , "   ")      & "dmPelsHeight         = " & dmDevMode.dmPelsHeight
	txt->AddLine IIf(dmDevMode.dmFields And DM_PELSWIDTH , " @ " , "   ")       & "dmPelsWidth          = " & dmDevMode.dmPelsWidth
	txt->AddLine IIf(dmDevMode.dmFields And DM_POSITION , " @ " , "   ")        & "dmPosition.x         = " & dmDevMode.dmPosition.x
	txt->AddLine IIf(dmDevMode.dmFields And DM_POSITION , " @ " , "   ")        & "dmPosition.y         = " & dmDevMode.dmPosition.y
	txt->AddLine IIf(dmDevMode.dmFields And DM_PRINTQUALITY , " @ " , "   ")    & "dmPrintQuality       = " & dmDevMode.dmPrintQuality
	txt->AddLine "   " & "dmReserved1          = " & dmDevMode.dmReserved1
	txt->AddLine "   " & "dmReserved2          = " & dmDevMode.dmReserved2
	txt->AddLine IIf(dmDevMode.dmFields And DM_SCALE , " @ " , "   ")           & "dmScale              = " & dmDevMode.dmScale
	txt->AddLine "   " & "dmSize               = " & dmDevMode.dmSize
	txt->AddLine "   " & "dmSpecVersion        = " & dmDevMode.dmSpecVersion
	txt->AddLine IIf(dmDevMode.dmFields And DM_TTOPTION , " @ " , "   ")        & "dmTTOption           = " & dmDevMode.dmTTOption
	txt->AddLine IIf(dmDevMode.dmFields And DM_YRESOLUTION , " @ " , "   ")     & "dmYResolution        = " & dmDevMode.dmYResolution
End Sub

Private Function QDCrtn2WStr(ByVal rtn As Long) ByRef As WString
	Select Case rtn
	Case ERROR_SUCCESS
		Return ML("The function succeeded.")
	Case ERROR_INVALID_PARAMETER
		Return ML("The combination of parameters and flags specified is invalid.")
	Case ERROR_NOT_SUPPORTED
		Return ML("The system is not running a graphics driver that was written according to the Windows Display Driver Model (WDDM). The function is only supported on a system with a WDDM driver running.")
	Case ERROR_ACCESS_DENIED
		Return ML("The caller does not have access to the console session. This error occurs if the calling process does not have access to the current desktop or is running on a remote session.")
	Case ERROR_GEN_FAILURE
		Return ML("An unspecified error occurred.")
	Case ERROR_BAD_CONFIGURATION
		Return ML("The function could not find a workable solution for the source and target modes that the caller did not specify.")
	Case Else
		Return ML("Unknow status.")
	End Select
End Function

Private Function CDSErtnWstr(ByVal rtn As Long) ByRef As WString
	Select Case rtn
	Case DISP_CHANGE_SUCCESSFUL
		Return ML("The settings change was successful.")
	Case DISP_CHANGE_BADFLAGS
		Return ML("An invalid set of values was used in the dwFlags parameter.")
	Case DISP_CHANGE_BADMODE
		Return ML("The graphics mode is not supported.")
	Case DISP_CHANGE_BADPARAM
		Return ML("An invalid parameter was used. This error can include an invalid value or combination of values.")
	Case DISP_CHANGE_FAILED
		Return ML("The display driver failed the specified graphics mode.")
	Case DISP_CHANGE_NOTUPDATED
		Return ML("ChangeDisplaySettingsEx was unable to write settings to the registry.")
	Case DISP_CHANGE_RESTART
		Return ML("The user must restart the computer for the graphics mode to work.")
	Case Else
		Return ML("Unknow status.")
	End Select
End Function

