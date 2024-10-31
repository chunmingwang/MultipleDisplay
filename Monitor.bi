'Monitor显示器
' Copyright (c) 2024 CM.Wang
' Freeware. Use at your own risk.

'EnumDisplayDevices
'	  https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaydevicesa
'
'EnumDisplaySettingsEx
'	  https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaysettingsexa
'
'ChangeDisplaySettingsEx
'    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-changedisplaysettingsexw
'
'SetDisplayConfig
'	  https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setdisplayconfig
'
'QueryDisplayConfig
'    https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-querydisplayconfig
'
'GetDisplayConfigBufferSizes
'    https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdisplayconfigbuffersizes

Const EDS_ROTATEDMODE = &H00000004
Const QDC_VIRTUAL_MODE_AWARE= &h00000010
Const QDC_INCLUDE_HMD = &h00000020
Const QDC_VIRTUAL_REFRESH_RATE_AWARE= &h00000040

#ifndef __USE_MAKE__
	#include once "Monitor.bas"
#endif
