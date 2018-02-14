#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\LOGO.ICO
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=13/09/2016
#AutoIt3Wrapper_Res_Description=Wallpaper Explorer v4.0
#AutoIt3Wrapper_Res_Fileversion=4.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.6.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Tooltip.au3>
Opt("TrayMenuMode", 3)
#include <TrayConstants.au3>
#include "image_get_info.au3"
#include <WinAPI.au3>
#include <GDIPlus.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <MetroButton.au3>
#include <GuiListView.au3>
#include <WinAPITheme.au3>
#include <WinAPIShPath.au3>
#include <WinAPIDlg.au3>
#include <WinAPISys.au3>
#include <ITaskBarList.au3>
#include <_Transitions.au3>
#include <WinAPIFiles.au3>
#include <GUI Wallpaper Explorer Library.au3>

_GDIPlus_Startup()
Global Const $sVersion = "4.0"
Global Const $iBorder = 1
If StringInStr($cmdlineraw, "/SaveImageProc") Then
	_SaveImageProc()
ElseIf StringInStr($cmdlineraw, "/open_location") Then
	Exit ShellExecute("explorer.exe", '/select,"' & _GetPathWallpaper() & '"')
ElseIf StringInStr($cmdlineraw, "/delete") Then
	Local $spath = _GetPathWallpaper()
	FileDelete($spath)
	If FileExists($spath) Then
		_WinAPI_DeleteFile($spath)
		If FileExists($spath) Then MsgBox(64, "Wallpaper Explorer " & $sVersion, "File is being used by another programs!")
	EndIf
	Exit
EndIf



Global $ARR_CHECK[15] = ["Resources\MENU.PNG", "Resources\logo.png", "Resources\logo.ico", "Resources\icon_share.ico", _
		"Resources\icon_picture.ico", "Resources\icon_info.ico", "Resources\icon_folder.ico", "Resources\ICON_PROP.ICO", _
		"Resources\icon_details.ico", "Resources\icon_delete.ico", "Resources\icon_copy.ico", "Resources\HELP_USER.PNG", _
		"Resources\icon_001.ico", "About.exe", "config.exe"]
For $i = 0 To UBound($ARR_CHECK) - 1
	If Not FileExists(@ScriptDir & "\" & $ARR_CHECK[$i]) Then
		Exit MsgBox(16, "Wallpaper Explorer Error Message", _
				'Oops! We can not find "' & StringUpper($ARR_CHECK[$i]) & '" file' & @CRLF & "Program Version: " & FileGetVersion(@ScriptFullPath) & @CRLF & @CRLF & @CRLF & "@errorcode = 0x" & Hex($i + 50, 2) & Hex(@MDAY, 2) & Hex(@MON, 2) & Hex(StringTrimLeft(@YEAR, 2), 2) & @CRLF)
	EndIf
Next


Global Const $INI_FILE = @ScriptDir & "\Wallpaper Explorer.ini"
Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
Global Const $sURL_Folder = "http://googledrive.com/host/0B4S2pTQPGLo6NjZjZFBmOWpaMEk"
Global Const $sURL_Check = $sURL_Folder & "/WallpaperExplorerUpdate.txt"
Global Const $sLocalFileDir = @TempDir & "\CheckForUpdate.ini"
Global $hCheckUpdate = InetGet($sURL_Check, $sLocalFileDir, 1, 1)

If Not FileExists($INI_FILE) Then FileOpen($INI_FILE, 8)


Local $sRegRead = RegRead("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer\command", "")
If @Compiled Then
	If ($sRegRead <> '"' & @ScriptFullPath & '"' Or @error) Then
		$Form1 = GUICreate("Install Context Menu", 580, 450, -1, -1, $WS_SYSMENU)
		GUISetBkColor(0xFFFFFF)
		GUICtrlCreateLabel("Do you want to create context menu of Wallpaper Explorer program in desktop?", 10, 5, 600, 25)
		GUICtrlSetColor(-1, 0x444444)
		GUICtrlSetFont(-1, 12, 400, 0, "Segoe UI")
		$hPic1 = GUICtrlGetHandle(GUICtrlCreatePic("", 0, 40, _WinAPI_GetClientWidth($Form1), _WinAPI_GetClientHeight($Form1) - 80))

		Local $hScanBitmap = _GDIPlus_BitmapCreateFromScan0(_WinAPI_GetClientWidth($Form1), _WinAPI_GetClientHeight($Form1) - 80)
		Local $hContext = _GDIPlus_ImageGetGraphicsContext($hScanBitmap)
		_GDIPlus_GraphicsSetSmoothingMode($hContext, $GDIP_SMOOTHINGMODE_HIGHQUALITY)
		_GDIPlus_GraphicsClear($hContext, 0xFFFFFF)
		Local $hUser_Help = _GDIPlus_BitmapCreateFromFile(@ScriptDir & "\Resources\HELP_USER.PNG")
		_GDIPlus_GraphicsDrawImage($hContext, $hUser_Help, 160, 0)
		Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hScanBitmap)
		_SendMessage($hPic1, 0x0172, $IMAGE_BITMAP, $hBitmap) ;Set bitmap to control

		$hButton1 = GUICtrlCreateButton("OK", 355, 390, 100, 25)
		$hButton2 = GUICtrlCreateButton("Cancel", 465, 390, 100, 25)
		GUISetState(@SW_SHOW)
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $hButton1
					ShellExecuteWait("config.exe", "/start")
					$nMsg = $GUI_EVENT_CLOSE
					ContinueCase
				Case $hButton2
					IniWrite($INI_FILE, "INFO", "DISPLAY", False)
					$nMsg = $GUI_EVENT_CLOSE
					ContinueCase
				Case $GUI_EVENT_CLOSE
					_GDIPlus_GraphicsDispose($hContext)
					_GDIPlus_BitmapDispose($hScanBitmap)
					_GDIPlus_BitmapDispose($hUser_Help)
					_WinAPI_DeleteObject($hBitmap)
					GUIDelete($Form1)

					ExitLoop
			EndSwitch
		WEnd
	EndIf
EndIf

Global $bSpecialOsVersion = False
If StringInStr("WIN_7_ WIN_VISTA_ WIN_XP_ WIN_XPe_", @OSVersion & "_") Then $bSpecialOsVersion = True

Global $sPathWallpaper, $hImageSource = 0, $temp_bitmap, $sImageInfo = ""
Global $__EditingWallpaper = False
Global $__WallpaperEdited = False
Global $__hGUI2_IsShow = True
Global Const $iPanelWidth = 160
Global Const $hMenuIcon = _GDIPlus_ImageResize(_GDIPlus_ImageLoadFromFile(@ScriptDir & "\Resources\MENU.PNG"), 45, 45)

;-------------------------------------------------------------------------------------------------------------------------------------------;
; Create GUI -------------------------------------------------------------------------------------------------------------------------------;
;-------------------------------------------------------------------------------------------------------------------------------------------;
Global Const $hGUI = GUICreate("Wallpaper Explorer v" & $sVersion, 0, 0, -10, 0, $WS_POPUP, $WS_EX_APPWINDOW)
Global Const $hGUI2 = GUICreate("$hGUI2", $iPanelWidth, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $hGUI3 = GUICreate("$hGUI3", _GDIPlus_ImageGetWidth($hMenuIcon), _GDIPlus_ImageGetHeight($hMenuIcon), 12, 12, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hGUI2)
Global Const $hGUI4 = GUICreate("$hGUI4", 250, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $__ImageInfoGUI = GUICreate("Details", 500, 400, 1, 1, $WS_SYSMENU, $WS_EX_MDICHILD, $hGUI)
Global Const $hGUICaption = GUICreate("$hGUICaption", 80, 32, 0, 0, $WS_POPUP, $WS_EX_LAYERED, $hGUI)

GUISetBkColor(0xFCFCFC, $hGUICaption)
GUISetCursor(7)
_WinAPI_SetLayeredWindowAttributes($hGUICaption, 0xFCFCFC, 255)
Local $iW = 40
Local $iH = _WinAPI_GetClientHeight($hGUICaption)
Global $ARR_BTNSYS_HANDLE[2]
$ARR_BTNSYS_HANDLE[0] = GUICtrlCreateLabel(@CRLF & "➖", 0, 0, $iW, $iH, BitOR($SS_NOTIFY, $SS_CENTER))
GUICtrlSetColor(-1, 0xFFFFFF)
GUICtrlSetFont(-1, 9, 900, "", "Segui UI")
GUICtrlSetState(-1, $GUI_ONTOP)
$ARR_BTNSYS_HANDLE[1] = GUICtrlCreateLabel("❌", $iW, 0, $iW, $iH, BitOR($SS_NOTIFY, $SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetColor(-1, 0xFFFFFF)
GUICtrlSetFont(-1, 12, 800, "", "Segui UI")
GUICtrlSetState(-1, $GUI_ONTOP)
Global Const $ARR_BTNSYS_COLOR[3][2] = [[0xFFFFFF, ""], [0xFFFFFF, 0xB8B8B8], [0x000000, 0xB3B3B3]] ; Text color and Background color of Normal, Hover, Click
Global Const $ARR_BTNSYS_POS[2][4] = [[0, 0, $iW, $iH], [$iW, 0, $iW, $iH]]
_Animation_ControlGUI_Init()



GUISwitch($hGUI)
Global Const $hPic = GUICtrlGetHandle(GUICtrlCreatePic('', 0, 0, 1000, 1000, BitOR($SS_CENTERIMAGE, $SS_NOPREFIX, $SS_NOTIFY), $GUI_WS_EX_PARENTDRAG))
GUICtrlSetBkColor(-1, 0xFFFFFF)


; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
GUISwitch($hGUI2)
Global Const $ARR_BUTTON_TEXT[7] = ["1. Open with...", "2. Delete", "3. Image Details", "4. Quick edit", "5. Open location", "6. Properties", "7. About"]
Global Const $ARR_BUTTON_ICON[UBound($ARR_BUTTON_TEXT)] = ["ICON_SHARE.ico", "ICON_DELETE.ICO", "ICON_DETAILS.ico", "icon_picture.ico", "ICON_FOLDER.ico", "ICON_PROP.ICO", "icon_info.ico"]
Global $ARR_HANDLE_BUTTON[UBound($ARR_BUTTON_TEXT)]
For $i = 0 To UBound($ARR_BUTTON_TEXT) - 1
	$ARR_HANDLE_BUTTON[$i] = _GUICtrlMetroButton_Create($hGUI2, $ARR_BUTTON_TEXT[$i], 0, 60 + 50 * $i, $iPanelWidth, 50, @ScriptDir & "\Resources\" & $ARR_BUTTON_ICON[$i])
Next
; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

GUISwitch($hGUI3)
SetBitmap($hGUI3, $hMenuIcon, 255)


GUISwitch($hGUI4)
#include <EditConstants.au3>
#include <SliderConstants.au3>
Global $g_aDefault[11] = [10000, 10000, 10000, 0, 10000, 0, 0, 0, 0, 0, 0]
Global $g_aAdjust = $g_aDefault

Global Const $hGUI4_BUTTONBACK = _GUICtrlMetroButton_Create($hGUI4, "Back", 0, 0, _WinAPI_GetClientWidth($hGUI4), 50, @ScriptDir & "\Resources\icon_001.ico")
Global $g_aidSlider[9][6] = [[0, 0, 'Red', 25, 650, 100], [0, 0, 'Green', 25, 650, 100], [0, 0, 'Blue', 25, 650, 100], [0, 0, 'Black', 0, 400, 10], [0, 0, 'White', 600, 1000, 10], [0, 0, 'Contrast', -100, 100, 1], [0, 0, 'Brightness', -100, 100, 1], [0, 0, 'Colorfulness', -100, 100, 1], [0, 0, 'Tint', -100, 100, 1]]
Global Const $iDeltaHeight = 50
For $i = 0 To 8
	GUICtrlCreateLabel($g_aidSlider[$i][2] & ':', 10, 21 + $i * 31 + $iDeltaHeight, 60, 14, $SS_RIGHT)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	$g_aidSlider[$i][0] = GUICtrlCreateSlider(71, 19 + $i * 31 + $iDeltaHeight, 170, 20, BitOR($TBS_BOTH, $TBS_TOOLTIPS, $TBS_NOTICKS))
	GUICtrlSetLimit(-1, $g_aidSlider[$i][4], $g_aidSlider[$i][3])
	$g_aidSlider[$i][1] = GUICtrlGetHandle(-1)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
Next
GUICtrlCreateLabel('Illuminant:', 10, 310 + $iDeltaHeight, 60, 14, $SS_RIGHT)
GUICtrlSetBkColor(-1, 0xFFFFFF)
#include <ComboConstants.au3>
#include <GuiComboBox.au3>
Global $g_idCombo = GUICtrlCreateCombo('', 77, 306 + $iDeltaHeight, 158, 160, $CBS_DROPDOWNLIST)
Local $aLight = StringSplit('Default|Tungsten lamp|Noon sunlight|NTSC daylight|Normal print|Bond paper print|Standard daylight|Northern daylight|Cool white lamp', '|')
For $i = 1 To $aLight[0]
	_GUICtrlComboBox_AddString(-1, $aLight[$i])
Next
GUICtrlCreateLabel('Filters:', 10, 347 + $iDeltaHeight, 60, 14, $SS_RIGHT)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $g_aidCheck[3]
$g_aidCheck[0] = GUICtrlCreateCheckbox('Negative (invert colors)', 77, 343 + $iDeltaHeight, 131, 21)
GUICtrlSetBkColor(-1, 0xFFFFFF)
$g_aidCheck[1] = GUICtrlCreateCheckbox('Logarithmic', 77, 368 + $iDeltaHeight, 76, 21)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Local $aidButton[2]
$aidButton[0] = GUICtrlCreateButton('Reset', 136 - 100, 448 + $iDeltaHeight, 75, 25)
$aidButton[1] = GUICtrlCreateButton('Save...', 218 - 100, 448 + $iDeltaHeight, 75, 25)
GUICtrlSetState($g_aidCheck[0], $GUI_FOCUS)
_Reset()


GUISwitch($__ImageInfoGUI)
Global Const $hListView = _GUICtrlListView_Create($__ImageInfoGUI, " Property | Information ", 0, 0, _WinAPI_GetClientWidth($__ImageInfoGUI), _WinAPI_GetClientHeight($__ImageInfoGUI), BitOR($LVS_AUTOARRANGE, $LVS_REPORT))
_GUICtrlListView_SetExtendedListViewStyle($hListView, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_HEADERDRAGDROP, $LVS_EX_INFOTIP))
_GUICtrlListView_SetColumnWidth($hListView, 0, 200)
_GUICtrlListView_SetColumnWidth($hListView, 1, _WinAPI_GetClientWidth($__ImageInfoGUI) - 200)
_WinAPI_SetWindowTheme($hListView, "Explorer")

GUIRegisterMsg($WM_EXITSIZEMOVE, "WM_EXITSIZEMOVE")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_HSCROLL, 'WM_HSCROLL')
GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE")
ConsoleWrite(StringFormat("Desktop Handle: %s\n", _WinAPI_GetDesktopWindow()))
GUISetBkColor(0xFFFFFF, $hGUI2)
WinSetTrans($hGUI2, "", 220)
WinSetTrans($hGUI, "", 0)
GUISetBkColor(0x2CDB3A, $hGUI)

GUISetBkColor(0xFFFFFF, $hGUI4)
WinSetTrans($hGUI4, "", 180)

Local $IsFirstTime = True
Global $hGUI4_BUTTONBACK_CLICKED = False
Global $arrhGUICaptionArea[4], $hBackupExtended
Global $tBkRect, $isMaximize = False, $IsMinimize = False
Global $hTimerWallpaperIsChanged

Func WM_ACTIVATE($hWnd, $iCode, $wParam, $lParam)
	If $hWnd = $hGUICaption And $wParam = 0x00000001 Then
		If $IsMinimize Then
			_WinAPI_MoveWindow($hGUI, DllStructGetData($tBkRect, "left"), DllStructGetData($tBkRect, "top"), DllStructGetData($tBkRect, "right") - DllStructGetData($tBkRect, "left"), DllStructGetData($tBkRect, "bottom") - DllStructGetData($tBkRect, "top"))
			$IsMinimize = False
		EndIf
	EndIf
EndFunc   ;==>WM_ACTIVATE

Global $stemp
Global $TaskBar_Info, $TaskBar_Delete
;-------------------------------------------------------------------------------------------------------------------------------------------;
; While loop -------------------------------------------------------------------------------------------------------------------------------;
;-------------------------------------------------------------------------------------------------------------------------------------------;
While True
	$nMsg = GUIGetMsg(1)
;~ 	If $nMsg[0] <> $stemp Then
;~ 		ConsoleWrite(StringFormat("$nMsg[0] = %s\n", $nMsg[0]))
;~ 		$stemp = $nMsg[0]
;~ 	EndIf
	Switch $nMsg[1]
		Case $hGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					ExitLoop
				Case $GUI_EVENT_RESTORE
				Case $GUI_EVENT_MINIMIZE
			EndSwitch
		Case $hGUI3
			Switch $nMsg[0]
				Case $GUI_EVENT_PRIMARYUP
					Local $tPoint = DllStructCreate("int X;int Y")
					DllStructSetData($tPoint, "X", MouseGetPos(0))
					DllStructSetData($tPoint, "Y", MouseGetPos(1))
					If _WinAPI_WindowFromPoint($tPoint) = $hGUI3 Then
						$arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
						_WinAPI_MoveWindow($hGUI2, $arrPos[0] + $iBorder, $arrPos[1] + $iBorder, _WinAPI_GetWindowWidth($hGUI2), _WinAPI_GetClientHeight($hGUI) - $iBorder * 2)
						If $__hGUI2_IsShow Then
							_Animation_FaceOut($hGUI2)
						Else
							_Animation_FaceIn($hGUI2)
						EndIf
						__hGUI2State_Toggle()
					EndIf
			EndSwitch
		Case $hGUI4
			Switch $nMsg[0]
				Case $aidButton[0]
					_Reset()
				Case $aidButton[1]
					_SaveImage()
				Case $g_aidCheck[0]
					If GUICtrlRead($g_aidCheck[0]) = $GUI_CHECKED Then
						$g_aAdjust[9] = BitOR($g_aAdjust[9], $CA_NEGATIVE)
					Else
						$g_aAdjust[9] = BitAND($g_aAdjust[9], BitNOT($CA_NEGATIVE))
					EndIf
					_Update()
				Case $g_aidCheck[1]
					If GUICtrlRead($g_aidCheck[1]) = $GUI_CHECKED Then
						$g_aAdjust[9] = BitOR($g_aAdjust[9], $CA_LOG_FILTER)
					Else
						$g_aAdjust[9] = BitAND($g_aAdjust[9], BitNOT($CA_LOG_FILTER))
					EndIf
					_Update()
				Case $g_idCombo
					$sData = _GUICtrlComboBox_GetCurSel($g_idCombo)
					If $g_aAdjust[10] <> $sData Then
						$g_aAdjust[10] = $sData
						_Update()
					EndIf
			EndSwitch
		Case $__ImageInfoGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $__ImageInfoGUI)
					GUISetState(@SW_ENABLE, $hGUI)
					GUISetState(@SW_ENABLE, $hGUI2)
					GUISetState(@SW_ENABLE, $hGUI3)
					GUISetState(@SW_ENABLE, $hGUI4)
					WinActivate($hGUI)
			EndSwitch
		Case $hGUICaption
			Switch $nMsg[0]
				Case $ARR_BTNSYS_HANDLE[0]
					If Not $isMaximize Then $tBkRect = _WinAPI_GetWindowRect($hGUI)
					_WinAPI_MoveWindow($hGUI, @DesktopWidth + 100, @DesktopHeight + 100, DllStructGetData($tBkRect, "right") - DllStructGetData($tBkRect, "left"), DllStructGetData($tBkRect, "bottom") - DllStructGetData($tBkRect, "top"), False)
					;WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
					$IsMinimize = True
					$iColor = $ARR_BTNSYS_COLOR[0][1]
					If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
					GUICtrlSetBkColor($nMsg[0], $iColor)
				Case $ARR_BTNSYS_HANDLE[1]
					$iColor = $ARR_BTNSYS_COLOR[1][1]
					If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
					GUICtrlSetBkColor($nMsg[0], $iColor)
					If $__EditingWallpaper Then
						If MsgBox(64 + 4, "Wallpaper Explorer Message", "This image has been changed. Do you want to save it?", 0, $hGUI) = 6 Then
							ConsoleWrite(StringFormat("Debug: line=%s, Call _SaveImage()=%s\n", @ScriptLineNumber, True))
							_SaveImage()
						EndIf
					EndIf
					ExitLoop
			EndSwitch
	EndSwitch

	If __WallpaperIsChanged() And Not $__EditingWallpaper Then
		If TimerDiff($hTimerWallpaperIsChanged) >= 3000 Then
			__ResetImage()
			WinSetTitle($__ImageInfoGUI, '', 'Image Details - ' & _WinAPI_PathStripPath($sPathWallpaper))
			__ImageInfoUpdate()
			$hTimerWallpaperIsChanged = TimerInit()
		EndIf
	ElseIf $IsFirstTime Then
		GUISetState(@SW_SHOW, $hGUI)
		For $i = 0 To 255 Step 255 / 10
			WinSetTrans($hGUI, "", $i)
			Sleep(1)
		Next
		GUISetState(@SW_SHOW, $hGUI2)
		GUISetState(@SW_SHOW, $hGUI3)
		GUISetState(@SW_SHOW, $hGUICaption)
		Sleep(10)
		WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
		Sleep(100)
		_Animation_ControlGUI_Start()

		Local $iTaskBar = _ITaskBar_CreateTaskBarObj()
		If $iTaskBar = 0 Then
			ConsoleWrite(StringFormat("Can't create TaskBar Button... \t @error = %s\n", @error))
		EndIf
		$TaskBar_Info = _ITaskBar_CreateTBButton('About Wallpaper Explorer', @ScriptDir & "\Resources\ICON_INFO.ICO")
		$TaskBar_Delete = _ITaskBar_CreateTBButton('Delete wallpaper', @ScriptDir & "\Resources\ICON_DELETE.ICO")
		_ITaskBar_AddTBButtons($hGUI)

		$IsFirstTime = False
	EndIf

	Switch _MouseEvent_hGUICaption()
		Case "WM_MOUSEINSIDE"
			Local $iColor, $Extended = @extended
			$iColor = $ARR_BTNSYS_COLOR[1][0]
			If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
			If $Extended = UBound($ARR_BTNSYS_HANDLE) - 1 Then $iColor = 0xFF3b3b
			GUICtrlSetColor($ARR_BTNSYS_HANDLE[$Extended], $iColor)
			$iColor = $ARR_BTNSYS_COLOR[1][1]
			If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
			GUICtrlSetBkColor($ARR_BTNSYS_HANDLE[$Extended], $iColor)
			WinSetTrans(GUICtrlGetHandle($Extended), "", 200)
		Case "WM_MOUSEOUTSIDE"
			Local $iColor, $Extended = @extended
			$iColor = $ARR_BTNSYS_COLOR[0][0]
			If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
			GUICtrlSetColor($ARR_BTNSYS_HANDLE[$Extended], $iColor)
			$iColor = $ARR_BTNSYS_COLOR[0][1]
			If $iColor = "" Then $iColor = $GUI_BKCOLOR_TRANSPARENT
			GUICtrlSetBkColor($ARR_BTNSYS_HANDLE[$Extended], $iColor)
			WinSetTrans(GUICtrlGetHandle($Extended), "", 255)
	EndSwitch

	If InetGetInfo($hCheckUpdate, 2) Then
		Local $iPresentVersion = FileGetVersion($sFile)
		Local $sNewVersion = IniRead($sLocalFileDir, "Info", "Version", "")
		If StringReplace($sNewVersion, ".", "") > StringReplace($iPresentVersion, ".", "") Then MsgBox(64, "Message", "Hi " & @UserName & ", new update is now available. Go to About -> Check for update to update program. Thank you!", $hGUI)
		InetClose($hCheckUpdate)
		$hCheckUpdate = 0
	EndIf

	If $hGUI4_BUTTONBACK_CLICKED = True Then
		If $__WallpaperEdited Then
			If MsgBox(64 + 4, "Wallpaper Explorer Message", "This image has been changed. Do you want to save it?", 0, $hGUI) = 6 Then
				ConsoleWrite(StringFormat("Debug: line=%s, Call _SaveImage()=%s\n", @ScriptLineNumber, True))
				_SaveImage()
			EndIf
		EndIf

		GUISetState(@SW_HIDE, $hGUI4)
		GUISetState(@SW_SHOW, $hGUI2)
		GUISetState(@SW_SHOW, $hGUI3)
		$__EditingWallpaper = False
		WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
		$hGUI4_BUTTONBACK_CLICKED = False
	EndIf
WEnd

Func _MouseEvent_hGUICaption()
	Local Static $hBTNSYS_PREVIOUSHOVER = -1, $hBkVarReturn = "", $hBkEvent = "", $bLastMessage = False
	Local $iX = DllStructGetData($tMousePoint, "X"), $iY = DllStructGetData($tMousePoint, "Y"), $VarReturn, $sEvent
	If ($iX >= $arrhGUICaptionArea[0] And $iX <= ($arrhGUICaptionArea[0] + $arrhGUICaptionArea[2])) And ($iY >= $arrhGUICaptionArea[1] And $iY <= ($arrhGUICaptionArea[1] + $arrhGUICaptionArea[3])) Then
		$iX -= $arrhGUICaptionArea[0]
		$iY -= $arrhGUICaptionArea[1]
		For $i = 0 To UBound($ARR_BTNSYS_POS) - 1
			If ($iX >= $ARR_BTNSYS_POS[$i][0] And $iX <= ($ARR_BTNSYS_POS[$i][0] + $ARR_BTNSYS_POS[$i][2])) And ($iY >= $ARR_BTNSYS_POS[$i][1] And $iY <= ($ARR_BTNSYS_POS[$i][1] + $ARR_BTNSYS_POS[$i][3])) Then
				$VarReturn = $i
				$sEvent = "WM_MOUSEINSIDE"
				$bLastMessage = False
				ExitLoop
			EndIf
		Next
		If $VarReturn <> $hBTNSYS_PREVIOUSHOVER And $hBTNSYS_PREVIOUSHOVER <> -1 Then
			$VarReturn = $hBTNSYS_PREVIOUSHOVER
			$sEvent = "WM_MOUSEOUTSIDE"
			$hBTNSYS_PREVIOUSHOVER = -1
		EndIf
	Else
		If $hBTNSYS_PREVIOUSHOVER <> -1 And $bLastMessage = False Then
			$VarReturn = $hBTNSYS_PREVIOUSHOVER
			$sEvent = "WM_MOUSEOUTSIDE"
			$hBTNSYS_PREVIOUSHOVER = -1
			$bLastMessage = True
		EndIf
	EndIf

	If $hBkEvent <> $sEvent Or $hBkVarReturn <> $VarReturn Then
		$hBkEvent = $sEvent
		$hBkVarReturn = $VarReturn
		$hBTNSYS_PREVIOUSHOVER = $VarReturn
		Return SetError("", $VarReturn, $sEvent)
	EndIf
	Return SetError("", "", "")
EndFunc   ;==>_MouseEvent_hGUICaption

Func _SaveImageProc()
	Local $tRect = _WinAPI_GetWindowRect(WinGetHandle("Wallpaper Explorer v" & $sVersion))
	Local $Handle = WinWaitActive("WallpaperExplorer: Save Image")
	Local $iBorderFrameW = 0, $iBorderFrameH = 0
	If @OSVersion = "WIN_10" Then
		$iBorderFrameW = _WinAPI_GetSystemMetrics(32)
		$iBorderFrameH = _WinAPI_GetSystemMetrics(33)
	EndIf
	WinMove($Handle, "", DllStructGetData($tRect, "left") + $iBorder - $iBorderFrameW, DllStructGetData($tRect, "top") + $iBorder, _
			DllStructGetData($tRect, "right") - DllStructGetData($tRect, "left") - $iBorder * 2 + $iBorderFrameW * 2, _
			DllStructGetData($tRect, "bottom") - DllStructGetData($tRect, "top") - $iBorder * 2 + $iBorderFrameH, 2)
	Exit
EndFunc   ;==>_SaveImageProc

Func _SaveImage()
	Local $sFolder = StringTrimRight($sPathWallpaper, StringLen(_WinAPI_PathStripPath($sPathWallpaper)) + 1)
	ShellExecute(@ScriptFullPath, "/SaveImageProc", @ScriptDir, "", @SW_HIDE)
	$sData = FileSaveDialog('WallpaperExplorer: Save Image', $sFolder, 'Image Files (*.bmp;*.dib;*.gif;*.jpg;*.tif)|All Files (*.*)', 2 + 16, $sFolder & '\' & StringTrimRight(_WinAPI_PathStripPath($sPathWallpaper), 4) & '_Edited.jpg', $hGUI)
	;ConsoleWrite(StringFormat("Debug: @error = %s\n", @error))
	GUICtrlSetState($aidButton[1], $GUI_FOCUS)

	If StringIsSpace($sData) Then
		Return
	EndIf
;~ 	If FileExists($sData) Then
;~ 		If MsgBox(1, "Wallpaper Explorer Saving Message", '"' & _WinAPI_PathStripPath($sPathWallpaper) & '" is already exist. If you want to replace it, cliking OK button to continue.', 0, $hGUI) = 2 Then Return
;~ 	EndIf

	Local $aAdjust = $g_aAdjust
	If IsArray($aAdjust) Then
		$tAdjust = _WinAPI_CreateColorAdjustment($aAdjust[9], $aAdjust[10], $aAdjust[0], $aAdjust[1], $aAdjust[2], $aAdjust[3], $aAdjust[4], $aAdjust[5], $aAdjust[6], $aAdjust[7], $aAdjust[8])
	EndIf
	Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImageSource)
	Local $hAdjustBitmap = _WinAPI_AdjustBitmap($hBitmap, -1, -1, $HALFTONE, $tAdjust)
	If $hAdjustBitmap Then
		Local $hBitmap_d = _GDIPlus_BitmapCreateFromHBITMAP($hAdjustBitmap)
		If _GDIPlus_ImageSaveToFile($hBitmap_d, $sData) Then
			$sData = 0
		EndIf
		_WinAPI_DeleteObject($hAdjustBitmap)
		_GDIPlus_BitmapDispose($hBitmap_d)
	EndIf
	_WinAPI_DeleteObject($hBitmap)
	If $sData Then
		MsgBox(16, "Wallpaper Explorer Message", "Unable to save image")
	Else
		_Tooltip("Sucessful to save image\nEnjoy! =)) :D")
	EndIf
	$__WallpaperEdited = False
EndFunc   ;==>_SaveImage

GUIDelete($__ImageInfoGUI)
GUIDelete($hGUI4)
GUIDelete($hGUI3)
GUIDelete($hGUI2)
GUIDelete($hGUI)
_GDIPlus_BitmapDispose($hMenuIcon)
_WinAPI_DeleteObject($temp_bitmap)
_GDIPlus_ImageDispose($hImageSource)
_GDIPlus_Shutdown()
GUIRegisterMsg($WM_ENTERSIZEMOVE, "")
GUIRegisterMsg($WM_EXITSIZEMOVE, "")
GUIRegisterMsg($WM_COMMAND, "")
GUIRegisterMsg($WM_HSCROLL, "")

FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
InetClose($hCheckUpdate)
ProcessClose(ProcessExists("about.exe"))

Func __ImageInfoUpdate()
	If Not StringInStr($sImageInfo, "File size=") Then $sImageInfo &= "File size=" & Round(FileGetSize($sPathWallpaper) / 1024 / 1024, 4) & " MB" & @CRLF
	ConsoleWrite(StringFormat("%s\n", $sImageInfo))
	Local $__arrInfo = StringRegExp($sImageInfo, "(?m)^.*$", 3)
	Local $__tempString
	If Not IsArray($__arrInfo) Then Return ConsoleWrite("Not array..." & @CRLF)
	_GUICtrlListView_BeginUpdate($hListView)
	_GUICtrlListView_DeleteAllItems($hListView)
	For $i = 0 To UBound($__arrInfo) - 1
		$__tempString = StringInStr($__arrInfo[$i], "=")
		If $__tempString <> 0 Then
			_GUICtrlListView_InsertItem($hListView, StringLeft($__arrInfo[$i], $__tempString - 1), $i)
			_GUICtrlListView_SetItemText($hListView, $i, StringTrimLeft($__arrInfo[$i], $__tempString), 1)
		EndIf
	Next
	_GUICtrlListView_EndUpdate($hListView)
EndFunc   ;==>__ImageInfoUpdate

Func WM_COMMAND($hWnd, $iMsg, $lParam, $wParam)
	Switch $hWnd
		Case $hGUI2
			Switch $lParam
				Case $WM_LBUTTONUP
					Switch $wParam
						Case $ARR_HANDLE_BUTTON[0]
							_WinAPI_ShellOpenWithDlg($sPathWallpaper, $OAIF_EXEC, $hGUI)
						Case $ARR_HANDLE_BUTTON[1]
							If MsgBox(64 + 1, "Wallpaper Explorer Message", StringFormat("Are you sure to delete your image file?\n - Click on OK button to continue.\n\n\n * Your image will be sent to recycle bin so you can restore one if need."), 0, $hGUI) = 1 Then
								FileRecycle($sPathWallpaper)
								If FileExists($sPathWallpaper) Then
									Local $sText = ""
									$sText = " - We have not permission to delete file in system folder or system disk" & @CRLF & _
											" - This file is used by other processes" & @CRLF & _
											" - Attribute of your disk driver or your file is read-only" & @CRLF & _
											" - Your file has been crashed. "
									MsgBox(16, "Wallpaper Explorer Error Message", "Oops, we are meeting some errors when deleting your wallpaper file. Please checking again some following reasons:" & @CRLF & @CRLF & $sText & @CRLF, 0, $hGUI)
								Else
									Dim $ARR_TEXT[3] = [":) :D :)", "Ooh oh, This previous picture is so great. Regret! :'(", "Have a good day, " & @UserName & "!"]
									_Tooltip(StringFormat("Delete image successfully!\n" & $ARR_TEXT[Random(0, UBound($ARR_TEXT) - 1, 1)]))
								EndIf
							EndIf
						Case $ARR_HANDLE_BUTTON[2]
							GUISetState(@SW_DISABLE, $hGUI)
							GUISetState(@SW_DISABLE, $hGUI2)
							GUISetState(@SW_DISABLE, $hGUI3)
							GUISetState(@SW_DISABLE, $hGUI4)
							__ImageInfoUpdate()

							Local $iW = _WinAPI_GetWindowWidth($__ImageInfoGUI)
							Local $iH = _WinAPI_GetWindowHeight($__ImageInfoGUI)
							_WinAPI_MoveWindow($__ImageInfoGUI, @DesktopWidth / 2 - $iW / 2, @DesktopHeight / 2 - $iH / 2, $iW, $iH)

							WinSetTitle($__ImageInfoGUI, '', 'Image Details - ' & _WinAPI_PathStripPath($sPathWallpaper))
							GUISetState(@SW_SHOWNORMAL, $__ImageInfoGUI)
						Case $ARR_HANDLE_BUTTON[3]
							Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
							_WinAPI_MoveWindow($hGUI4, $arrPos[0] + $iBorder, $arrPos[1] + $iBorder, 250, _WinAPI_GetClientHeight($hGUI) - $iBorder * 2)
							GUISetState(@SW_SHOW, $hGUI4)
							GUISetState(@SW_HIDE, $hGUI3)
							GUISetState(@SW_HIDE, $hGUI2)
							$__EditingWallpaper = True
							_Reset()
						Case $ARR_HANDLE_BUTTON[4]
							ShellExecute("explorer.exe", '/select,"' & $sPathWallpaper & '"')
						Case $ARR_HANDLE_BUTTON[5]
							Local $arrBkPosition = MouseGetPos()
							;_WinAPI_ShowCursor(False)
;~ 							MouseMove(@DesktopWidth / 2 - 180, @DesktopHeight / 2 - 280, 0)
							$oshellApp = ObjCreate("shell.application")
							$oshellApp.namespace(0).parsename($sPathWallpaper).invokeverb("Properties")
;~ 							Sleep(100)
;~ 							MouseMove($arrBkPosition[0], $arrBkPosition[1], 0)
							;_WinAPI_ShowCursor(True)
							;WinSetOnTop($Handle, "", 1)
						Case $ARR_HANDLE_BUTTON[6]
							ShellExecute(@ScriptDir & "\about.exe", "/start")
					EndSwitch
			EndSwitch
		Case $hGUI4
			Switch $lParam
				Case $WM_LBUTTONUP
					Switch $wParam
						Case $hGUI4_BUTTONBACK
							$hGUI4_BUTTONBACK_CLICKED = True
					EndSwitch
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func _Reset()
	$g_aAdjust = $g_aDefault
	For $i = 0 To 8
		GUICtrlSetData($g_aidSlider[$i][0], $g_aAdjust[$i] / $g_aidSlider[$i][5])
		GUICtrlSetData($g_aidSlider[$i][2], $g_aAdjust[$i])
	Next
	_GUICtrlComboBox_SetCurSel($g_idCombo, $g_aAdjust[10])
	If BitAND($g_aAdjust[9], $CA_LOG_FILTER) Then
		GUICtrlSetState($g_aidCheck[1], $GUI_CHECKED)
	Else
		GUICtrlSetState($g_aidCheck[1], $GUI_UNCHECKED)
	EndIf
	If BitAND($g_aAdjust[9], $CA_NEGATIVE) Then
		GUICtrlSetState($g_aidCheck[0], $GUI_CHECKED)
	Else
		GUICtrlSetState($g_aidCheck[0], $GUI_UNCHECKED)
	EndIf
	_Update()
	$__WallpaperEdited = False
EndFunc   ;==>_Reset

Func _Update()
	_SetBitmapAdjust($hPic, $temp_bitmap, $g_aAdjust)
	$__WallpaperEdited = True
EndFunc   ;==>_Update

Func WM_HSCROLL($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam

	Switch $hWnd
		Case $hGUI4
			For $i = 0 To 8
				If $g_aidSlider[$i][1] = $lParam Then
					$g_aAdjust[$i] = $g_aidSlider[$i][5] * GUICtrlRead($g_aidSlider[$i][0])
					GUICtrlSetData($g_aidSlider[$i][2], $g_aAdjust[$i])
					_Update()
					ExitLoop
				EndIf
			Next
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_HSCROLL


Func __hGUI2State_Toggle()
	$__hGUI2_IsShow = Not $__hGUI2_IsShow
	Return (Not $__hGUI2_IsShow) ; previous state
EndFunc   ;==>__hGUI2State_Toggle

Func WM_ENTERSIZEMOVE($hWnd, $iMsg, $lParam, $wParam)
EndFunc   ;==>WM_ENTERSIZEMOVE

Func __ResetImage()
	_GDIPlus_ImageDispose($hImageSource)
	FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
	$sPathWallpaper = _GetPathWallpaper()
	If Not FileExists($sPathWallpaper) Or FileGetSize($sPathWallpaper) = 0 Then Exit MsgBox(16, "Wallpaper Explorer Error Message", "Ooops, file doesn't exist.")
	$sImageInfo = _ImageGetInfo($sPathWallpaper)
	FileCopy($sPathWallpaper, @TempDir & "\", 1)

	$hImageSource = _GDIPlus_ImageLoadFromFile(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
	If $hImageSource = 0 Then
		MsgBox(64, "Wallpaper Explorer Error Message", "Sorry, some errors occur when trying to load your image file." & @CRLF & "Error function: _GDIPlus_ImageLoadFromFile()" & @CRLF & "@ErrorCode = " & @error & @CRLF)
		Exit
	Else
		ConsoleWrite("+> Load image succesfully..." & @TAB)
	EndIf
	Local $iW = _GDIPlus_ImageGetWidth($hImageSource)
	Local $iH = _GDIPlus_ImageGetHeight($hImageSource)
	Local $iTrayHeight = _WinAPI_GetWindowHeight(WinGetHandle("[CLASS:Shell_TrayWnd]"))

	If $iW > @DesktopWidth * 0.75 Or $iH > @DesktopHeight * 0.75 Then
		$iW = @DesktopWidth * 0.75
		$iH = @DesktopHeight * 0.75
	EndIf

	Local Static $iBkW = 0, $iBkH = 0
	If $IsMinimize = False Then
		If $iBkW <> $iW Or $iBkH <> $iH Then
			GUISetState(@SW_HIDE, $hGUICaption)
			_WinAPI_MoveWindow($hGUI, @DesktopWidth / 2 - $iW / 2, (@DesktopHeight - $iTrayHeight) / 2 - $iH / 2, $iW, $iH)
			Local $tRect = _WinAPI_GetClientRect($hGUI)
			Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
			_WinAPI_MoveWindow($hGUI2, $aGUIPos[0] + $iBorder, $aGUIPos[1] + $iBorder, _WinAPI_GetWindowWidth($hGUI2), _
					DllStructGetData($tRect, "bottom") - DllStructGetData($tRect, "top") - $iBorder * 2)

			$iBkW = $iW
			$iBkH = $iH
		EndIf
	EndIf
	WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
	GUISetState(@SW_SHOW, $hGUICaption)
EndFunc   ;==>__ResetImage

;-------------------------------------------------------------------------------------------------------------------------------------------;
; WM_EXITSIZEMOVE---------------------------------------------------------------------------------------------------------------------------;
;-------------------------------------------------------------------------------------------------------------------------------------------;
Func WM_EXITSIZEMOVE($hWnd, $iMsg, $wParam, $lParam)
	If $hWnd = $hGUI Then
		Local $iW = _WinAPI_GetClientWidth($hGUI), $iH = _WinAPI_GetClientHeight($hGUI)
		Local $temp_bitmap_old = $temp_bitmap
		Local $__tempimage = $hImageSource
		Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)

		If $__EditingWallpaper Then $__tempimage = _GDIPlus_BitmapCreateFromHBITMAP(_SendMessage($hPic, $STM_GETIMAGE))
		$temp_bitmap = _GDIPlus_ImageResizeToBitmap($__tempimage, $iW, $iH)

		If $__hGUI2_IsShow Then _WinAPI_MoveWindow($hGUI2, $aGUIPos[0] + $iBorder, $aGUIPos[1] + $iBorder, $iPanelWidth, _WinAPI_GetClientHeight($hGUI) - $iBorder * 2)

		_WinAPI_MoveWindow($hGUI4, $aGUIPos[0] + $iBorder, $aGUIPos[1] + $iBorder, _WinAPI_GetClientWidth($hGUI4), _WinAPI_GetClientHeight($hGUI) - $iBorder * 2)

		_WinAPI_MoveWindow($hPic, $iBorder, $iBorder, $iW - $iBorder * 2, $iH - $iBorder * 2)

		; API Move $hGUICaption ------------------------------------------------------------------------------------------------------------;
		$iW = _WinAPI_GetClientWidth($hGUICaption)
		$arrhGUICaptionArea[0] = $aGUIPos[0] + _WinAPI_GetClientWidth($hGUI) - $iBorder - $iW - 8
		$arrhGUICaptionArea[1] = $aGUIPos[1] + $iBorder + 5
		$arrhGUICaptionArea[2] = $iW
		$arrhGUICaptionArea[3] = _WinAPI_GetClientHeight($hGUICaption)
		_WinAPI_MoveWindow($hGUICaption, $arrhGUICaptionArea[0], $arrhGUICaptionArea[1], $arrhGUICaptionArea[2], $arrhGUICaptionArea[3])

		_SetBitmapAdjust($hPic, $temp_bitmap, 0)
		_WinAPI_DeleteObject($temp_bitmap_old)
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_EXITSIZEMOVE

Func _GDIPlus_ImageResizeToBitmap($hImage, $iW, $iH)
	Local $temp_image = _GDIPlus_ImageResize($hImage, $iW, $iH)
	Local $temp_bitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($temp_image)
	_GDIPlus_ImageDispose($temp_image)
	Return $temp_bitmap
EndFunc   ;==>_GDIPlus_ImageResizeToBitmap

Func __WallpaperIsChanged()
	Local $Now_Wallpaper = _GetPathWallpaper()
	If $Now_Wallpaper <> $sPathWallpaper And StringLen($Now_Wallpaper) > 0 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>__WallpaperIsChanged

Func _Animation_FaceOut($hWnd, $ihGUI4 = False)
	GUISetState(@SW_DISABLE, $hGUI)
	Local $tRect[UBound($ARR_HANDLE_BUTTON)][4]
	Local $temp
	For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
		$temp = _GUICtrlMetroButton_GetRect($hWnd, $ARR_HANDLE_BUTTON[$i])
		$tRect[$i][0] = DllStructGetData($temp, "left")
		$tRect[$i][1] = DllStructGetData($temp, "top")
		$tRect[$i][2] = DllStructGetData($temp, "right") - DllStructGetData($temp, "left")
		$tRect[$i][3] = DllStructGetData($temp, "bottom") - DllStructGetData($temp, "top")

		Assign("iFrameLoop" & $i, "0", 1)
	Next
	Local $iSumFrame = 20, $x, $y, $w, $h
	Local $sBtnCanStart = "0"
	Local $iBtnCount = UBound($ARR_HANDLE_BUTTON) - 1
	While True
		For $i = 0 To $iBtnCount
			If StringInStr($sBtnCanStart, String($i)) Then
				If Eval("iFrameLoop" & $i) <= $iSumFrame Then
					;ConsoleWrite(StringFormat("Debug: line=%s, $iFrameLoop%s=%s, $sBtnCanStart=%s\n", @ScriptLineNumber, $i, Eval("iFrameLoop" & $i), $sBtnCanStart))
					Assign("iX" & $i, -_QuadEaseIn(Eval("iFrameLoop" & $i), 0, $tRect[$i][2], $iSumFrame), 1)
					_WinAPI_MoveWindow($ARR_HANDLE_BUTTON[$i], Eval("iX" & $i), $tRect[$i][1], $tRect[$i][2], $tRect[$i][3], True)
					If Abs(Eval("iX" & $i)) >= $tRect[$i][2] * (1 / $iSumFrame) Then
						If $i + 1 <= $iBtnCount And Not StringInStr($sBtnCanStart, String($i + 1)) Then
							$sBtnCanStart &= String($i + 1)
						EndIf
					EndIf
					;If Eval("iFrameLoop" & $i) = $iSumFrame Then $sBtnCanStart = StringReplace($sBtnCanStart, $i, "")
					Assign("iFrameLoop" & $i, Number(Eval("iFrameLoop" & $i)) + 1, 1)
				EndIf
			EndIf
		Next
		If Eval("iFrameLoop" & $iBtnCount) > $iSumFrame Then ExitLoop
	WEnd


	Local $iOW = $iPanelWidth
	Local $iW = _WinAPI_GetWindowWidth($hWnd), $iH = _WinAPI_GetWindowHeight($hWnd)
	Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
	For $i = 1 To 50
		Sleep(1)
		;If $i >= 30 And $i < 50 Then Sleep(1);Sleep(1 + (($i - 30) * 2))
		_WinAPI_MoveWindow($hWnd, $arrPos[0] + $iBorder, $arrPos[1] + $iBorder, $iPanelWidth - _ExpoEaseOut($i, 0, $iPanelWidth, 50), $iH, False)
	Next
	GUISetState(@SW_ENABLE, $hGUI)
EndFunc   ;==>_Animation_FaceOut

Func _Animation_FaceIn($hWnd)
	GUISetState(@SW_DISABLE, $hGUI)
	Local $tRect[UBound($ARR_HANDLE_BUTTON)][4]
	Local $temp
	For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
		$temp = _GUICtrlMetroButton_GetRect($hWnd, $ARR_HANDLE_BUTTON[$i])
		$tRect[$i][0] = DllStructGetData($temp, "left")
		$tRect[$i][1] = DllStructGetData($temp, "top")
		$tRect[$i][2] = DllStructGetData($temp, "right") - DllStructGetData($temp, "left")
		$tRect[$i][3] = DllStructGetData($temp, "bottom") - DllStructGetData($temp, "top")

		Assign("iFrameLoop" & $i, "0", 1)
	Next
	Local $iSumFrame = 7, $x, $y, $w, $h
	Local $sBtnCanStart = "0"
	Local $iBtnCount = UBound($ARR_HANDLE_BUTTON) - 1
	Local $iOW = $iPanelWidth, $iH = _WinAPI_GetWindowHeight($hWnd)
	Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
	For $i = 1 To 50
		If $i <= 30 Then Sleep(1)
		;If $i >= 30 Then Sleep(1 + (($i - 30) * 2))
		_WinAPI_MoveWindow($hWnd, $arrPos[0] + $iBorder, $arrPos[1] + $iBorder, _ExpoEaseOut($i, 0, $iOW, 50), $iH, True)
	Next
	While True
		For $i = 0 To $iBtnCount
			If StringInStr($sBtnCanStart, String($i)) Then
				If Eval("iFrameLoop" & $i) <= $iSumFrame Then
					;ConsoleWrite(StringFormat("Debug: line=%s, $iFrameLoop%s=%s, $sBtnCanStart=%s\n", @ScriptLineNumber, $i, Eval("iFrameLoop" & $i), $sBtnCanStart))
					Assign("iX" & $i, $tRect[$i][0] + _ExpoEaseOut(Eval("iFrameLoop" & $i), 0, $tRect[$i][2], $iSumFrame), 1)
					;If Eval("iFrameLoop" & $i) >= $iSumFrame * 0.5 Then Sleep(1)
					_WinAPI_MoveWindow($ARR_HANDLE_BUTTON[$i], Eval("iX" & $i), $tRect[$i][1], $tRect[$i][2], $tRect[$i][3], False)
					If Abs(Eval("iX" & $i)) <= $tRect[$i][2] * (1 / 2.6) Then
						If $i + 1 <= $iBtnCount And Not StringInStr($sBtnCanStart, String($i + 1)) Then
							$sBtnCanStart &= String($i + 1)
						EndIf
					EndIf
					;If Eval("iFrameLoop" & $i) = $iSumFrame Then $sBtnCanStart = StringReplace($sBtnCanStart, $i, "")
					Assign("iFrameLoop" & $i, Number(Eval("iFrameLoop" & $i)) + 1, 1)
				EndIf
			EndIf
		Next
		If Eval("iFrameLoop" & $iBtnCount) > $iSumFrame Then ExitLoop
		Sleep(1)
	WEnd
	GUISetState(@SW_ENABLE, $hGUI)
EndFunc   ;==>_Animation_FaceIn

Func _Animation_ControlGUI_Init()
	For $i = 0 To UBound($ARR_BTNSYS_HANDLE) - 1
		GUICtrlSetPos($ARR_BTNSYS_HANDLE[$i], $ARR_BTNSYS_POS[$i][0], -$ARR_BTNSYS_POS[$i][3])
	Next
EndFunc   ;==>_Animation_ControlGUI_Init

Func _Animation_ControlGUI_Start()
	Local $iSumFrame = 13
	Local $iHeight = _WinAPI_GetWindowHeight($hGUICaption)
	Local $iControlCount = UBound($ARR_BTNSYS_HANDLE) - 1
	Local $sBtnCanStart = "0"

	While True
		For $i = 0 To $iControlCount
			If StringInStr($sBtnCanStart, String($i)) Then
				If Eval("iFrameLoop" & $i) <= $iSumFrame Then
					Assign("iY" & $i, -$iHeight + _CubicEaseOut(Eval("iFrameLoop" & $i), 0, $ARR_BTNSYS_POS[$i][3], $iSumFrame), 1)
					GUICtrlSetPos($ARR_BTNSYS_HANDLE[$i], $ARR_BTNSYS_POS[$i][0], Eval("iY" & $i))
					If Abs(Eval("iY" & $i)) <= $ARR_BTNSYS_POS[$i][3] * (16 / 17) Then
						If $i + 1 <= $iControlCount And Not StringInStr($sBtnCanStart, String($i + 1)) Then
							$sBtnCanStart &= String($i + 1)
						EndIf
					EndIf
					Assign("iFrameLoop" & $i, Number(Eval("iFrameLoop" & $i)) + 1, 1)
				EndIf
			EndIf
		Next
		If Eval("iFrameLoop" & $iControlCount) > $iSumFrame Then ExitLoop
		Sleep(1)
	WEnd
EndFunc   ;==>_Animation_ControlGUI_Start


