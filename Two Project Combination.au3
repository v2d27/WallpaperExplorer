#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\LOGO.ICO
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Wallpaper Explorer v3.7.6
#AutoIt3Wrapper_Res_Fileversion=3.7.6.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.6.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#include "image_get_info.au3"
#include <WinAPI.au3>
#include <GDIPlus.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <MetroButton.au3>
#include <GUI Wallpaper Explorer Library.au3>
#include <GuiListView.au3>
#include <WinAPITheme.au3>
#include <WinAPIShPath.au3>
#include <WinAPIDlg.au3>
#include <WinAPISys.au3>
#include <_Transitions.au3>

_GDIPlus_Startup()
Global Const $sVersion = "3.7.6"

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
Global Const $hMenuIcon = _GDIPlus_ImageResize(_GDIPlus_ImageLoadFromFile(@ScriptDir & "\Resources\MENU.PNG"), 45, 45)

Global Const $hGUI = GUICreate("Wallpaper Explorer v" & $sVersion, 0, 0, -200, 0, BitOR($WS_MAXIMIZEBOX, $WS_SIZEBOX, $GUI_SS_DEFAULT_GUI), $WS_EX_APPWINDOW)
;~ Global Const $hGUI = GUICreate("Wallpaper Explorer v" & $sVersion, 0, 0, -200, 0, $WS_POPUP, $WS_EX_APPWINDOW)
Global Const $hGUI2 = GUICreate("", 200, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $hGUI3 = GUICreate("", _GDIPlus_ImageGetWidth($hMenuIcon), _GDIPlus_ImageGetHeight($hMenuIcon), 10, 10, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hGUI2)
Global Const $hGUI4 = GUICreate("", 250, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $__ImageInfoGUI = GUICreate("Details", 600, 500, 1, 1, $WS_SYSMENU, $WS_EX_MDICHILD, $hGUI)

GUISwitch($hGUI)
Global Const $hPic = GUICtrlGetHandle(GUICtrlCreatePic('', 200, 0, 500, 200, BitOR($SS_CENTERIMAGE, $SS_NOPREFIX, $SS_NOTIFY)))
GUICtrlSetBkColor(-1, 0xFFFFFF)


GUISwitch($hGUI2)
Global Const $ARR_BUTTON_TEXT[8] = ["Open with...", "Delete", "Recycle bin", "Image Details", "Quick edit", "Open location", "Properties", "About"]
Global Const $ARR_BUTTON_ICON[8] = ["ICON_SHARE.ico", "ICON_DELETE.ICO", "ICON_DELETE.ICO", "ICON_DETAILS.ico", "icon_picture.ico", "ICON_FOLDER.ico", "ICON_PROP.ICO", "icon_info.ico"]
Global $ARR_HANDLE_BUTTON[8]
For $i = 0 To UBound($ARR_BUTTON_TEXT) - 1
	$ARR_HANDLE_BUTTON[$i] = _GUICtrlMetroButton_Create($hGUI2, $ARR_BUTTON_TEXT[$i], 0, 60 + 50 * $i, 200, 50, @ScriptDir & "\Resources\" & $ARR_BUTTON_ICON[$i])
Next


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
Global Const $hListView = _GUICtrlListView_Create($__ImageInfoGUI, " Property | Value ", 0, 0, _WinAPI_GetClientWidth($__ImageInfoGUI), _WinAPI_GetClientHeight($__ImageInfoGUI))
_GUICtrlListView_SetExtendedListViewStyle($hListView, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
_GUICtrlListView_SetColumnWidth($hListView, 0, 200)
_GUICtrlListView_SetColumnWidth($hListView, 1, _WinAPI_GetClientWidth($__ImageInfoGUI) - 200)
_WinAPI_SetWindowTheme($hListView, "Explorer")


GUIRegisterMsg($WM_ENTERSIZEMOVE, "WM_ENTERSIZEMOVE")
GUIRegisterMsg($WM_EXITSIZEMOVE, "WM_EXITSIZEMOVE")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_HSCROLL, 'WM_HSCROLL')
GUISetBkColor(0xFFFFFF, $hGUI2)
WinSetTrans($hGUI2, "", 200)
GUISetBkColor(0xFFFFFF, $hGUI)
GUISetBkColor(0xFFFFFF, $hGUI4)
WinSetTrans($hGUI4, "", 230)
GUISetState(@SW_SHOW, $hGUI)
Sleep(300)
GUISetState(@SW_SHOW, $hGUI2)
GUISetState(@SW_SHOW, $hGUI3)


While True
	$nMsg = GUIGetMsg(1)
	Switch $nMsg[1]
		Case $hGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					ExitLoop
				Case $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESTORE
					WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
			EndSwitch
		Case $hGUI3
			Switch $nMsg[0]
				Case $GUI_EVENT_PRIMARYUP
					$arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
					_WinAPI_MoveWindow($hGUI2, $arrPos[0], $arrPos[1], _WinAPI_GetWindowWidth($hGUI2), _WinAPI_GetClientHeight($hGUI))
					If $__hGUI2_IsShow Then
						_Animation_FaceOut($hGUI2)
					Else
						_Animation_FaceIn($hGUI2)
					EndIf
					__hGUI2State_Toggle()
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
	EndSwitch

	If __WallpaperIsChanged() And Not $__EditingWallpaper Then
		__ResetImage()
		WinSetTitle($__ImageInfoGUI, '', 'Details of "' & _WinAPI_PathStripPath($sPathWallpaper) & '"')
		__ImageInfoUpdate()
	EndIf

	If InetGetInfo($hCheckUpdate, 2) Then
		Local $iPresentVersion = FileGetVersion($sFile)
		Local $sNewVersion = IniRead($sLocalFileDir, "Info", "Version", "")
		If StringReplace($sNewVersion, ".", "") > StringReplace($iPresentVersion, ".", "") Then MsgBox(64, "Message", "Hi " & @UserName & ", new update is now available. Go to About -> Check for update to update program. Thank you!", $hGUI)
		InetClose($hCheckUpdate)
		$hCheckUpdate = 0
	EndIf
WEnd

Func _SaveImage()
	$sData = FileSaveDialog('Save Image', @WorkingDir, 'Image Files (*.bmp;*.dib;*.gif;*.jpg;*.tif)|All Files (*.*)', 2 + 16, @ScriptDir & '\Wallpaper_' & Random(1, 999999999, 1) & '.jpg', $hGUI)
	GUICtrlSetState($aidButton[1], $GUI_FOCUS)

	If Not $sData Then
		Return
	EndIf
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
		MsgBox(64, "Wallpaper Explorer Message", "Successful to save image")
	EndIf
	$__WallpaperEdited = False
EndFunc   ;==>_SaveImage


_GDIPlus_BitmapDispose($hMenuIcon)
_WinAPI_DeleteObject($temp_bitmap)
_GDIPlus_ImageDispose($hImageSource)
_GDIPlus_Shutdown()
GUIRegisterMsg($WM_ENTERSIZEMOVE, "")
GUIRegisterMsg($WM_EXITSIZEMOVE, "")
GUIRegisterMsg($WM_COMMAND, "")
GUIRegisterMsg($WM_HSCROLL, "")
GUIDelete($__ImageInfoGUI)
GUIDelete($hGUI4)
GUIDelete($hGUI3)
GUIDelete($hGUI2)
GUIDelete($hGUI)
FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
InetClose($hCheckUpdate)
ProcessClose(ProcessExists("about.exe"))



Func _Animation_FaceOut($hWnd, $ihGUI4 = False)
	For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
		_GUICtrlMetroButton_Animation($ARR_HANDLE_BUTTON[$i])
	Next

	Local $iW = _WinAPI_GetWindowWidth($hWnd), $iH = _WinAPI_GetWindowHeight($hWnd)
	Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
	; Parameters ....: $iFrame - Current frame, or point.
	;                   $iStartValue - Beginning point
	;                   $iInterval - Change in value
	;                   $iEndValue - Number of frames
	;_ExpoEaseOut($iFrame, $iStartValue, $iInterval, $iEndValue)
	;return c * ( -Math.pow( 2, -10 * t/d ) + 1 )
	;Return $iInterval * (-(2 ^ (-10 * $iFrame / $iEndValue)) + 1) + $iStartValue
	For $i = 1 To 28
		_WinAPI_MoveWindow($hWnd, $arrPos[0], $arrPos[1], 200 - _CubicEaseOut($i, 0, 200, 28), $iH, False)
		Sleep(1)
	Next
EndFunc   ;==>_Animation_FaceOut

Func _Animation_FaceIn($hWnd)
	Local $iOW = 200, $iH = _WinAPI_GetWindowHeight($hWnd)
	Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)

	For $i = 1 To 28
		_WinAPI_MoveWindow($hWnd, $arrPos[0], $arrPos[1], _CubicEaseOut($i, 0, 200, 28), $iH, True)
		Sleep(1)
	Next
	For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
		_GUICtrlMetroButton_Animation($ARR_HANDLE_BUTTON[$i])
	Next
EndFunc   ;==>_Animation_FaceIn

Func __ImageInfoUpdate()
	Local $__arrInfo = StringRegExp($sImageInfo, "(?m)^.*$", 3)
	Local $__tempString = 0
	If Not IsArray($__arrInfo) Then Return ConsoleWrite("Not array..." & @CRLF)
	_GUICtrlListView_BeginUpdate($hListView)
	_GUICtrlListView_DeleteAllItems($hListView)
	For $i = 0 To UBound($__arrInfo) - 1
		$__tempString = StringSplit($__arrInfo[$i], "=")
		_GUICtrlListView_InsertItem($hListView, $__tempString[1], $i)
		_GUICtrlListView_SetItemText($hListView, $i, $__tempString[2], 1)
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
							If MsgBox(64 + 1, "Wallpaper Explorer Message", "Do your want to delete your wallpaper file? You can't restore it. Still continue?" & @CRLF & @CRLF) = 1 Then
								FileDelete($sPathWallpaper)
								If FileExists($sPathWallpaper) Then
									Local $sText = ""
									$sText = "- We have not permission to delete file in system folder or system disk" & @CRLF & _
											"- This file is used by other processes" & @CRLF & _
											"- Attribute of your disk driver or your file is read-only" & @CRLF & _
											"- Your file has been crashed. "
									MsgBox(16, "Wallpaper Explorer Error Message", "Oops, we are meeting some errors when deleting your wallpaper file. Please checking again some following reasons:" & @CRLF & @CRLF & $sText & @CRLF)
								Else
									MsgBox(64, "Wallpaper Explorer Message", "Delete image successfully!")
								EndIf
							EndIf
						Case $ARR_HANDLE_BUTTON[2]
							FileRecycle($sPathWallpaper)
							If FileExists($sPathWallpaper) Then
								Local $sText = ""
								$sText = "- We have not permission to delete file in system folder or system disk" & @CRLF & _
										"- This file is used by other processes" & @CRLF & _
										"- Attribute of your disk driver or your file is read-only" & @CRLF & _
										"- Your file has been crashed. "
								MsgBox(16, "Wallpaper Explorer Error Message", "Sorry, your file can not be moved to recycle bin. Please checking again some following reasons:" & @CRLF & @CRLF & $sText & @CRLF)
							Else
								MsgBox(64, "Wallpaper Explorer Message", "Image was moved to recycle bin!")
							EndIf
						Case $ARR_HANDLE_BUTTON[3]
							GUISetState(@SW_DISABLE, $hGUI)
							GUISetState(@SW_DISABLE, $hGUI2)
							GUISetState(@SW_DISABLE, $hGUI3)
							GUISetState(@SW_DISABLE, $hGUI4)
							__ImageInfoUpdate()

							Local $iW = _WinAPI_GetWindowWidth($__ImageInfoGUI)
							Local $iH = _WinAPI_GetWindowHeight($__ImageInfoGUI)
							_WinAPI_MoveWindow($__ImageInfoGUI, @DesktopWidth / 2 - $iW / 2, @DesktopHeight / 2 - $iH / 2, $iW, $iH)

							WinSetTitle($__ImageInfoGUI, '', 'Details of "' & _WinAPI_PathStripPath($sPathWallpaper) & '"')
							GUISetState(@SW_SHOWNORMAL, $__ImageInfoGUI)
						Case $ARR_HANDLE_BUTTON[4]
							Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
							_WinAPI_MoveWindow($hGUI4, $arrPos[0], $arrPos[1], 250, _WinAPI_GetClientHeight($hGUI))
							GUISetState(@SW_SHOW, $hGUI4)
							GUISetState(@SW_HIDE, $hGUI3)
							GUISetState(@SW_HIDE, $hGUI2)
							$__EditingWallpaper = True
							_Reset()
						Case $ARR_HANDLE_BUTTON[5]
							ShellExecute("explorer.exe", '/select,"' & $sPathWallpaper & '"')
						Case $ARR_HANDLE_BUTTON[6]
							$oshellApp = ObjCreate("shell.application")
							$oshellApp.namespace(0).parsename($sPathWallpaper).invokeverb("Properties")
						Case $ARR_HANDLE_BUTTON[7]
							ShellExecute(@ScriptDir & "\about.exe", "/start")
					EndSwitch
			EndSwitch
		Case $hGUI4
			Switch $lParam
				Case $WM_LBUTTONUP
					Switch $wParam
						Case $hGUI4_BUTTONBACK
							If $__WallpaperEdited Then
								If MsgBox(64 + 4, "Wallpaper Explorer Message", "This image has been changed. Do you want to save it?") = 6 Then
									_SaveImage()
								EndIf
							EndIf

							GUISetState(@SW_HIDE, $hGUI4)
							GUISetState(@SW_SHOW, $hGUI2)
							GUISetState(@SW_SHOW, $hGUI3)
							$__EditingWallpaper = False
							WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
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
	Switch $hWnd
		Case $hGUI
			Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
			_WinAPI_MoveWindow($hGUI2, $arrPos[0], $arrPos[1], 0, 400)
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_ENTERSIZEMOVE

Func WM_EXITSIZEMOVE($hWnd, $iMsg, $lParam, $wParam)
	Switch $hWnd
		Case $hGUI
			Local $iW = _WinAPI_GetClientWidth($hGUI)
			Local $iH = _WinAPI_GetClientHeight($hGUI)
			Local $temp_bitmap_old = $temp_bitmap
			Local $__tempimage = $hImageSource
			If $__EditingWallpaper Then $__tempimage = _GDIPlus_BitmapCreateFromHBITMAP(_SendMessage($hPic, $STM_GETIMAGE))
			$temp_bitmap = _GDIPlus_ImageResizeToBitmap($__tempimage, $iW, $iH)
			_WinAPI_MoveWindow($hPic, -1, -1, $iW + 2, $iH + 2)
			_SetBitmapAdjust($hPic, $temp_bitmap, 0)
			_WinAPI_DeleteObject($temp_bitmap_old)

			Local $tRect = _WinAPI_GetClientRect($hGUI)
			Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
			If $__hGUI2_IsShow Then _WinAPI_MoveWindow($hGUI2, $aGUIPos[0], $aGUIPos[1], 200, _WinAPI_GetClientHeight($hGUI))
			_WinAPI_MoveWindow($hGUI4, $aGUIPos[0], $aGUIPos[1], _WinAPI_GetClientWidth($hGUI4), _WinAPI_GetClientHeight($hGUI))
	EndSwitch
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

Func __ResetImage()
	_GDIPlus_ImageDispose($hImageSource)
	FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
	$sPathWallpaper = _GetPathWallpaper()
	$sImageInfo = _ImageGetInfo($sPathWallpaper)
	FileCopy($sPathWallpaper, @TempDir & "\", 1)
	$hImageSource = _GDIPlus_ImageLoadFromFile(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
	If $hImageSource = -1 Then Exit ConsoleWrite("! Can't load image file..." & @CRLF)

	Local $iW = _GDIPlus_ImageGetWidth($hImageSource)
	Local $iH = _GDIPlus_ImageGetHeight($hImageSource)
	Local $iTrayHeight = _WinAPI_GetWindowHeight(WinGetHandle("[CLASS:Shell_TrayWnd]"))

	If $iW > @DesktopWidth * 0.75 Or $iH > @DesktopHeight * 0.75 Then
		$iW = @DesktopWidth * 0.75
		$iH = @DesktopHeight * 0.75
	EndIf

	_WinAPI_MoveWindow($hGUI, @DesktopWidth / 2 - $iW / 2, (@DesktopHeight - $iTrayHeight) / 2 - $iH / 2, $iW, $iH)

	Local $tRect = _WinAPI_GetClientRect($hGUI)
	Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
	_WinAPI_MoveWindow($hGUI2, $aGUIPos[0], $aGUIPos[1], _WinAPI_GetWindowWidth($hGUI2), _
			DllStructGetData($tRect, "bottom") - DllStructGetData($tRect, "top"))
	WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
EndFunc   ;==>__ResetImage
