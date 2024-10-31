'MultipleMonitor多显示器
' Copyright (c) 2024 CM.Wang
' Freeware. Use at your own risk.

'#Region "Form"
	#if defined(__FB_MAIN__) AndAlso Not defined(__MAIN_FILE__)
		#define __MAIN_FILE__
		#ifdef __FB_WIN32__
			#cmdline "frmDisplay.rc"
		#endif
		Const _MAIN_FILE_ = __FILE__
	#endif
	#include once "mff/Form.bi"
	#include once "mff/Panel.bi"
	#include once "mff/TextBox.bi"
	#include once "mff/CommandButton.bi"
	#include once "mff/ComboBoxEdit.bi"
	#include once "mff/CheckBox.bi"
	#include once "mff/ListControl.bi"
	#include once "mff/Splitter.bi"
	
	#include once "win/winuser.bi"
	#include once "Monitor.bi"
	
	Using My.Sys.Forms
	
	Type frmDisplayType Extends Form
		Dim mtrCount As Integer = -1
		Dim mtrMI(Any) As MONITORINFO
		Dim mtrHMtr(Any) As HMONITOR
		Dim mtrHDC(Any) As HDC
		Dim mtrRECT(Any) As tagRECT
		
		Declare Sub EnumDisplayMode(ByVal NameIndex As Integer, ByVal FlagIndex As Integer)
		Declare Sub GetDisplayMode(ByVal NameIndex As Integer, ByVal FlagIndex As Integer, ByVal Index As Integer)
		Declare Static Function MonitorEnumProc(ByVal hMtr As HMONITOR , ByVal hDCMonitor As HDC , ByVal lprcMonitor As LPRECT , ByVal dwData As LPARAM) As WINBOOL
		
		Declare Sub Form_Create(ByRef Sender As Control)
		Declare Sub CommandButton2_Click(ByRef Sender As Control)
		Declare Sub CommandButton5_Click(ByRef Sender As Control)
		Declare Sub CommandButton4_Click(ByRef Sender As Control)
		Declare Sub ComboBoxEdit3_Selected(ByRef Sender As ComboBoxEdit, ItemIndex As Integer)
		Declare Sub CommandButton3_Click(ByRef Sender As Control)
		Declare Sub CommandButton1_Click(ByRef Sender As Control)
		Declare Sub CommandButton6_Click(ByRef Sender As Control)
		Declare Sub ComboBoxEdit4_Selected(ByRef Sender As ComboBoxEdit, ItemIndex As Integer)
		Declare Sub ListControl2_Click(ByRef Sender As Control)
		Declare Sub Form_Close(ByRef Sender As Form, ByRef Action As Integer)
		Declare Constructor
		
		Dim As Panel Panel1
		Dim As TextBox TextBox1
		Dim As CommandButton CommandButton1, CommandButton2, CommandButton3, CommandButton4, CommandButton5, CommandButton6
		Dim As ComboBoxEdit ComboBoxEdit1, ComboBoxEdit2, ComboBoxEdit3, ComboBoxEdit5, ComboBoxEdit4
		Dim As ListControl ListControl1, ListControl2
		Dim As Splitter Splitter1
	End Type
	
	Constructor frmDisplayType
		' frmDisplay
		With This
			.Name = "frmDisplay"
			.StartPosition = FormStartPosition.CenterScreen
			.Designer = @This
			.OnCreate = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @Form_Create)
			#ifdef __FB_64BIT__
				'...instructions for 64bit OSes...
				.Caption = ML("VisualFBEditor Multiple Display64")
			#else
				'...instructions for other OSes
				.Caption = ML("VisualFBEditor Multiple Display32")
			#endif
			.OnClose = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Form, ByRef Action As Integer), @Form_Close)
			.SetBounds 0, 0, 800, 610
		End With
		' Panel1
		With Panel1
			.Name = "Panel1"
			.Text = "Panel1"
			.TabIndex = 0
			.Align = DockStyle.alLeft
			.ExtraMargins.Top = 10
			.ExtraMargins.Right = 0
			.ExtraMargins.Left = 10
			.ExtraMargins.Bottom = 10
			.SetBounds 10, 10, 230, 551
			.Parent = @This
		End With
		' ComboBoxEdit1
		With ComboBoxEdit1
			.Name = "ComboBoxEdit1"
			.Text = "ComboBoxEdit1"
			.TabIndex = 1
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 0, 230, 21
			.Parent = @Panel1
		End With
		' ListControl1
		With ListControl1
			.Name = "ListControl1"
			.Text = "ListControl1"
			.TabIndex = 2
			.SelectionMode = SelectionModes.smMultiExtended
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 22, 230, 95
			.Parent = @Panel1
		End With
		' CommandButton1
		With CommandButton1
			.Name = "CommandButton1"
			.Text = ML("Query Display Config")
			.TabIndex = 3
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 120, 230, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton1_Click)
			.Parent = @Panel1
		End With
		' ComboBoxEdit2
		With ComboBoxEdit2
			.Name = "ComboBoxEdit2"
			.Text = "ComboBoxEdit2"
			.TabIndex = 4
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 160, 110, 21
			.Parent = @Panel1
		End With
		' CommandButton2
		With CommandButton2
			.Name = "CommandButton2"
			.Text = ML("Set Display Config")
			.TabIndex = 5
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 120, 160, 110, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton2_Click)
			.Parent = @Panel1
		End With
		' CommandButton3
		With CommandButton3
			.Name = "CommandButton3"
			.Text = ML("Enum Display Monitors")
			.TabIndex = 6
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 200, 230, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton3_Click)
			.Parent = @Panel1
		End With
		' CommandButton4
		With CommandButton4
			.Name = "CommandButton4"
			.Text = ML("Enum Display Devices")
			.TabIndex = 7
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 240, 230, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton4_Click)
			.Parent = @Panel1
		End With
		' ComboBoxEdit3
		With ComboBoxEdit3
			.Name = "ComboBoxEdit3"
			.Text = "ComboBoxEdit3"
			.TabIndex = 8
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 260, 230, 21
			.Designer = @This
			.OnSelected = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As ComboBoxEdit, ItemIndex As Integer), @ComboBoxEdit3_Selected)
			.Parent = @Panel1
		End With
		' CommandButton5
		With CommandButton5
			.Name = "CommandButton5"
			.Text = ML("Enum Display Settings")
			.TabIndex = 9
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 300, 230, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton5_Click)
			.Parent = @Panel1
		End With
		' ComboBoxEdit4
		With ComboBoxEdit4
			.Name = "ComboBoxEdit4"
			.Text = "ComboBoxEdit4"
			.TabIndex = 10
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.SetBounds 0, 320, 230, 21
			.Designer = @This
			.OnSelected = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As ComboBoxEdit, ItemIndex As Integer), @ComboBoxEdit4_Selected)
			.Parent = @Panel1
		End With
		' ListControl2
		With ListControl2
			.Name = "ListControl2"
			.Text = "ListControl2"
			.TabIndex = 19
			.ControlIndex = 11
			.Anchor.Bottom = AnchorStyle.asAnchor
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.Anchor.Top = AnchorStyle.asAnchor
			.SetBounds 0, 350, 230, 154
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @ListControl2_Click)
			.Parent = @Panel1
		End With
		' ComboBoxEdit5
		With ComboBoxEdit5
			.Name = "ComboBoxEdit5"
			.Text = "ComboBoxEdit5"
			.TabIndex = 12
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.ControlIndex = 3
			.Anchor.Bottom = AnchorStyle.asAnchor
			.SetBounds 0, 510, 230, 21
			.Parent = @Panel1
		End With
		' CommandButton6
		With CommandButton6
			.Name = "CommandButton6"
			.Text = ML("Change Display Settings")
			.TabIndex = 13
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.ControlIndex = 4
			.Anchor.Bottom = AnchorStyle.asAnchor
			.SetBounds 0, 530, 230, 20
			.Designer = @This
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @CommandButton6_Click)
			.Parent = @Panel1
		End With
		' Splitter1
		With Splitter1
			.Name = "Splitter1"
			.Text = "Splitter1"
			.SetBounds 240, 10, 10, 110
			.Designer = @This
			.Parent = @This
		End With
		' TextBox1
		With TextBox1
			.Name = "TextBox1"
			.Text = ML("Hello!") & WChr(55357, 56832)
			.TabIndex = 15
			.Multiline = True
			.ScrollBars = ScrollBarsType.Both
			.Align = DockStyle.alClient
			.Font.Name = "Consolas"
			.ControlIndex = 3
			.ExtraMargins.Bottom = 10
			.ExtraMargins.Right = 10
			.ExtraMargins.Top = 10
			.SetBounds 250, 0, 534, 561
			.Parent = @This
		End With
	End Constructor
	
	Dim Shared frmDisplay As frmDisplayType
	
	#if _MAIN_FILE_ = __FILE__
		App.DarkMode = True
		frmDisplay.MainForm = True
		frmDisplay.Show
		App.Run
	#endif
'#End Region

Private Sub frmDisplayType.GetDisplayMode(ByVal NameIndex As Integer, ByVal FlagIndex As Integer, ByVal Index As Integer)
	If NameIndex < 0 Or FlagIndex < 0 Then Exit Sub
	
	Dim dmDevMode() As DEVMODE
	Dim dmDevModeCur As DEVMODE
	Dim iModeCur As Integer = -1
	Dim iModeNum As Integer = -1
	Dim i As Long
	Dim tmpi As WString Ptr
	Dim tmpc As WString Ptr
	
	Dim dwFlags As DWORD = EDSEdwFlags(FlagIndex)
	
	memset (@dmDevModeCur, 0, SizeOf(DEVMODE))
	dmDevModeCur.dmSize = SizeOf(DEVMODE)
	EnumDisplaySettingsEx(ComboBoxEdit3.Item(NameIndex), Index, @dmDevModeCur, dwFlags)
	
	TextBox1.Clear
	TextBox1.AddLine ComboBoxEdit3.Item(NameIndex)
	TextBox1.AddLine ML("Total DEVMODE number") & ": " & ListControl2.ItemCount
	TextBox1.AddLine ML("Current DEVMODE") & ": " & ListControl2.Item(Index)
	TextBox1.AddLine ML("Current DEVMODE number") & ": " & Index
	DM2WStr(dmDevModeCur, @TextBox1)
	Deallocate(tmpi)
	Deallocate(tmpc)
End Sub

Private Sub frmDisplayType.EnumDisplayMode(ByVal NameIndex As Integer, ByVal FlagIndex As Integer)
	ListControl2.Clear
	If NameIndex < 0 Or FlagIndex < 0 Then Exit Sub
	
	Dim dmDevMode() As DEVMODE
	Dim dmDevModeCur As DEVMODE
	Dim iModeCur As Integer = -1
	Dim iModeNum As Integer = -1
	Dim i As Long
	Dim tmpi As String
	Dim tmpc As String
	
	Dim dwFlags As DWORD = EDSEdwFlags(FlagIndex)
	
	memset (@dmDevModeCur, 0, SizeOf(DEVMODE))
	dmDevModeCur.dmSize = SizeOf(DEVMODE)
	EnumDisplaySettingsEx(ComboBoxEdit3.Item(NameIndex), ENUM_CURRENT_SETTINGS, @dmDevModeCur, dwFlags)
	
	tmpc = DM2SimpleWStr(dmDevModeCur)
	
	Do
		i = iModeNum + 1
		ReDim Preserve dmDevMode(i)
		memset (@dmDevMode(i), 0, SizeOf(DEVMODE))
		dmDevMode(i).dmSize = SizeOf(DEVMODE)
		
		If EnumDisplaySettingsEx(ComboBoxEdit3.Item(NameIndex), i , @dmDevMode(i), dwFlags) Then
			tmpi = DM2SimpleWStr(dmDevMode(i))
			ListControl2.AddItem i & " - " & tmpi
			If tmpc = tmpi Then
				iModeCur = i
			End If
			iModeNum = i
		Else
			Exit Do
		End If
	Loop While (True)
	ListControl2.ItemIndex = iModeCur
	TextBox1.Clear
	TextBox1.AddLine ComboBoxEdit3.Item(NameIndex)
	TextBox1.AddLine ML("Total DEVMODE number") & ": " & iModeNum
	TextBox1.AddLine ML("Current DEVMODE number") & ": " & iModeCur
	DM2WStr(dmDevModeCur, @TextBox1)
End Sub

Private Sub frmDisplayType.Form_Create(ByRef Sender As Control)
	ComboBoxEdit1.AddItem ML("QDC_ALL_PATHS")
	ComboBoxEdit1.AddItem ML("QDC_ONLY_ACTIVE_PATHS")
	ComboBoxEdit1.AddItem ML("QDC_DATABASE_CURRENT")
	ComboBoxEdit1.ItemIndex = 0
	ListControl1.AddItem ML("QDC_VIRTUAL_MODE_AWARE")
	ListControl1.AddItem ML("QDC_INCLUDE_HMD")
	ListControl1.AddItem ML("QDC_VIRTUAL_REFRESH_RATE_AWARE")
	ListControl1.AddItem ML("NONE")
	ListControl1.ItemIndex = 0
	
	ComboBoxEdit4.AddItem ML("EDS RAWMODE")
	ComboBoxEdit4.AddItem ML("EDS ROTATEDMODE")
	ComboBoxEdit4.ItemIndex = 0
	
	ComboBoxEdit2.AddItem ML("Clone") '"SDC_TOPOLOGY_CLONE"
	ComboBoxEdit2.AddItem ML("Extended") '"SDC_TOPOLOGY_EXTEND"
	ComboBoxEdit2.AddItem ML("Internal") '"SDC_TOPOLOGY_INTERNAL"
	ComboBoxEdit2.AddItem ML("External") '"SDC_TOPOLOGY_EXTERNAL"
	ComboBoxEdit2.ItemIndex = 0
	
	ComboBoxEdit5.AddItem "0"
	ComboBoxEdit5.AddItem ML("CDS FULLSCREEN")
	ComboBoxEdit5.AddItem ML("CDS GLOBAL")
	ComboBoxEdit5.AddItem ML("CDS NORESET")
	ComboBoxEdit5.AddItem ML("CDS RESET")
	ComboBoxEdit5.AddItem ML("CDS SET PRIMARY")
	ComboBoxEdit5.AddItem ML("CDS TEST")
	ComboBoxEdit5.AddItem ML("CDS UPDATEREGISTRY")
	ComboBoxEdit5.AddItem ML("CDS VIDEOPARAMETERS")
	ComboBoxEdit5.AddItem ML("CDS ENABLE UNSAFE MODES")
	ComboBoxEdit5.AddItem ML("CDS DISABLE UNSAFE MODES")
	ComboBoxEdit5.ItemIndex = 7
End Sub

Private Function frmDisplayType.MonitorEnumProc(ByVal hMtr As HMONITOR , ByVal hDCMonitor As HDC , ByVal lprcMonitor As LPRECT , ByVal dwData As LPARAM) As WINBOOL
	Dim a As frmDisplayType Ptr = Cast(frmDisplayType Ptr, dwData)
	a->mtrCount += 1
	
	ReDim Preserve a->mtrHMtr(a->mtrCount)
	ReDim Preserve a->mtrHDC(a->mtrCount)
	ReDim Preserve a->mtrRECT(a->mtrCount)
	ReDim Preserve a->mtrMI(a->mtrCount)
	a->mtrHMtr(a->mtrCount) = hMtr
	a->mtrHDC(a->mtrCount) = hDCMonitor
	memcpy(VarPtr(a->mtrRECT(a->mtrCount)), lprcMonitor, SizeOf(tagRECT))
	a->mtrMI(a->mtrCount).cbSize = SizeOf(MONITORINFO)
	GetMonitorInfo(hMtr, @(a->mtrMI(a->mtrCount)))
	Return True
End Function

Private Sub frmDisplayType.CommandButton2_Click(ByRef Sender As Control)
	Dim flag As UINT32
	Select Case ComboBoxEdit2.ItemIndex
	Case 0
		flag = SDC_TOPOLOGY_CLONE
	Case 1
		flag = SDC_TOPOLOGY_EXTEND
	Case 2
		flag = SDC_TOPOLOGY_INTERNAL
	Case 3
		flag = SDC_TOPOLOGY_EXTERNAL
	End Select
	flag = flag Or SDC_APPLY
	
	Dim rtn As Integer = SetDisplayConfig(0, NULL, 0, NULL, flag)
	TextBox1.Clear
	TextBox1.AddLine ComboBoxEdit2.Item(ComboBoxEdit2.ItemIndex)
	TextBox1.AddLine ML("Set Display Config") & " = " & rtn
	TextBox1.AddLine QDCrtn2WStr(rtn)
	ComboBoxEdit3.Clear
	ListControl2.Clear
End Sub

Private Sub frmDisplayType.CommandButton5_Click(ByRef Sender As Control)
	EnumDisplayMode(ComboBoxEdit3.ItemIndex, ComboBoxEdit4.ItemIndex)
End Sub

Private Sub frmDisplayType.CommandButton4_Click(ByRef Sender As Control)
	ListControl2.Clear
	ComboBoxEdit3.Clear
	
	Dim iDevNum As DWORD = 0
	Dim ddDisplay As DISPLAY_DEVICE
	Dim dmDevMode As DEVMODE
	Dim dwFlags As DWORD = EDSEdwFlags(ComboBoxEdit4.ItemIndex)
	TextBox1.Clear
	Do
		memset (@ddDisplay, 0, SizeOf(ddDisplay))
		ddDisplay.cb = SizeOf(ddDisplay)
		memset (@dmDevMode, 0, SizeOf(dmDevMode))
		dmDevMode.dmSize = SizeOf(dmDevMode)
		
		If EnumDisplayDevices(NULL, iDevNum, @ddDisplay, EDD_GET_DEVICE_INTERFACE_NAME) Then
			TextBox1.AddLine "EnumDisplayDevices " & iDevNum + 1 & " =========================================="
			DD2WStr(ddDisplay, @TextBox1)
			TextBox1.AddLine ""
			
			If EnumDisplaySettingsEx(ddDisplay.DeviceName, ENUM_CURRENT_SETTINGS, @dmDevMode, dwFlags) Then
				ComboBoxEdit3.AddItem ddDisplay.DeviceName
				TextBox1.AddLine "EnumDisplaySettingsEx " & ddDisplay.DeviceName & " --------------------------------------"
				TextBox1.AddLine "dmDevMode"
				DM2WStr(dmDevMode, @TextBox1)
				TextBox1.AddLine "---------------------------------------"
				TextBox1.AddLine ""
			End If
			
			iDevNum += 1
		Else
			Exit Do
		End If
	Loop While True
End Sub

Private Sub frmDisplayType.ComboBoxEdit3_Selected(ByRef Sender As ComboBoxEdit, ItemIndex As Integer)
	EnumDisplayMode(ItemIndex, ComboBoxEdit4.ItemIndex)
End Sub

Private Sub frmDisplayType.CommandButton3_Click(ByRef Sender As Control)
	mtrCount = -1
	Erase mtrMI
	Erase mtrHMtr
	Erase mtrHDC
	Erase mtrRECT
	
	TextBox1.Clear
	TextBox1.AddLine "EnumDisplayMonitors: " & EnumDisplayMonitors(NULL , NULL , @MonitorEnumProc, Cast(LPARAM, @This))
	TextBox1.AddLine "Monitor count: " & mtrCount + 1
	Dim i As Integer
	For i = 0 To mtrCount
		TextBox1.AddLine ""
		TextBox1.AddLine "Monitor " & i + 1 & " ================="
		TextBox1.AddLine "EnumDisplayMonitors---------"
		TextBox1.AddLine "hMtr =        " & mtrHMtr(i)
		TextBox1.AddLine "hDCMonitor =  " & mtrHDC(i)
		TextBox1.AddLine "####mrtRECT"
		RECT2WStr(mtrRECT(i), @TextBox1, "    ")
		TextBox1.AddLine "GetMonitorInfo--------------"
		TextBox1.AddLine "####rcMonitor"
		RECT2WStr(mtrMI(i).rcMonitor, @TextBox1, "    ")
		TextBox1.AddLine "####rcWork"
		RECT2WStr(mtrMI(i).rcWork, @TextBox1, "    ")
		TextBox1.AddLine "dwFlags =     " & mtrMI(i).dwFlags & IIf(mtrMI(i).dwFlags = MONITORINFOF_PRIMARY, ", This is the primary display monitor.", ", This is not the primary display monitor.")
	Next
End Sub

Private Sub frmDisplayType.CommandButton1_Click(ByRef Sender As Control)
	Dim rtn As Long
	
	Dim PathArraySize As UINT32 = 0
	Dim ModeArraySize  As UINT32 = 0
	Dim flags As UINT32
	
	Select Case ComboBoxEdit1.ItemIndex
	Case 0
		flags = QDC_ALL_PATHS
	Case 1
		flags = QDC_ONLY_ACTIVE_PATHS
	Case Else
		flags = QDC_DATABASE_CURRENT
	End Select
	
	Dim i As Integer
	For i = 0 To ListControl1.ItemCount - 1
		If ListControl1.Selected(i) Then
			Select Case i
			Case 0
				flags += QDC_VIRTUAL_MODE_AWARE
			Case 1
				flags += QDC_INCLUDE_HMD
			Case 2
				flags += QDC_VIRTUAL_REFRESH_RATE_AWARE
			End Select
		End If
	Next

	rtn = GetDisplayConfigBufferSizes(flags, @PathArraySize, @ModeArraySize)
	TextBox1.Clear
	TextBox1.AddLine "GetDisplayConfigBufferSizes = " & rtn & ", " & flags
	TextBox1.AddLine "PathArraySize = " & PathArraySize
	TextBox1.AddLine "ModeArraySize = " & ModeArraySize
	TextBox1.AddLine QDCrtn2WStr(rtn)
	
	Dim currentTopologyId As DISPLAYCONFIG_TOPOLOGY_ID
	Dim PathArray(PathArraySize-1) As DISPLAYCONFIG_PATH_INFO
	Dim ModeArray(ModeArraySize-1) As DISPLAYCONFIG_MODE_INFO
	
	rtn = QueryDisplayConfig(flags, @PathArraySize, @PathArray(0), @ModeArraySize, @ModeArray(0) , @currentTopologyId)
	TextBox1.AddLine ""
	TextBox1.AddLine "QueryDisplayConfig = " & rtn
	TextBox1.AddLine QDCrtn2WStr(rtn)
	
	Dim tmp As WString Ptr
	
	TextBox1.AddLine "PathArray()===================="
	For i = 0 To PathArraySize-1
		TextBox1.AddLine "Index: " & i & " --------------------"
		Path2WStr(PathArray(i), @TextBox1)
		TextBox1.AddLine *tmp
	Next
	TextBox1.AddLine "ModeArray()===================="
	For i = 0 To ModeArraySize-1
		TextBox1.AddLine "Index: " & i & " --------------------"
		Mode2WStr(ModeArray(i), @TextBox1)
		TextBox1.AddLine *tmp
	Next
	Deallocate(tmp)
End Sub

Private Sub frmDisplayType.CommandButton6_Click(ByRef Sender As Control)
	Dim dmDevMode As DEVMODEW
	memset(@dmDevMode, 0, SizeOf(dmDevMode))
	dmDevMode.dmSize = SizeOf (dmDevMode)
	Dim dwFlags As DWORD = EDSEdwFlags(ComboBoxEdit4.ItemIndex)
	
	Dim brtn As WINBOOL = EnumDisplaySettingsEx(ComboBoxEdit3.Item(ComboBoxEdit3.ItemIndex), ListControl2.ItemIndex , @dmDevMode, dwFlags)
	Dim rtn As Long = ChangeDisplaySettingsEx(ComboBoxEdit3.Item(ComboBoxEdit3.ItemIndex), @dmDevMode , NULL , CDSdwFlags(ComboBoxEdit5.ItemIndex) , NULL)
	TextBox1.Clear
	TextBox1.AddLine "EnumDisplaySettingsEx: " & brtn
	TextBox1.AddLine "ChangeDisplaySettingsEx: " & rtn
	TextBox1.AddLine CDSErtnWstr(rtn)
End Sub

Private Sub frmDisplayType.ListControl2_Click(ByRef Sender As Control)
	GetDisplayMode(ComboBoxEdit3.ItemIndex, ComboBoxEdit4.ItemIndex, ListControl2.ItemIndex)
End Sub

Private Sub frmDisplayType.ComboBoxEdit4_Selected(ByRef Sender As ComboBoxEdit, ItemIndex As Integer)
	EnumDisplayMode(ComboBoxEdit3.ItemIndex, ItemIndex)
End Sub

Private Sub frmDisplayType.Form_Close(ByRef Sender As Form, ByRef Action As Integer)
	Erase mtrMI
	Erase mtrHMtr
	Erase mtrHDC
	Erase mtrRECT
End Sub
