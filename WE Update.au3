#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\LOGO.ICO
#AutoIt3Wrapper_Outfile=WE Updater.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Wallpaper Explorer Updater
#AutoIt3Wrapper_Res_Fileversion=3.7.4.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.4.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

If StringLen($cmdlineraw) = 0 Then Exit

_GDIPlus_Startup()
Global Const $ScriptDir = $cmdlineraw
Global Const $sFile = $ScriptDir & "\WallPaperExplorer.exe"
If Not FileExists($sFile) Then Exit
Global Const $INI_FILE = $ScriptDir & "\Wallpaper Explorer.ini"
Global Const $sURL_Folder = "http://googledrive.com/host/0B4S2pTQPGLo6NjZjZFBmOWpaMEk"
Global Const $sURL_Check = $sURL_Folder & "/WallpaperExplorerUpdate.txt"
Global Const $sLocalFileDir = @TempDir & "\CheckForUpdate.ini"
Global Const $sLocalPackageDir = @TempDir & "\WallpaperExplorerUpdate.exe"
Global Const $STM_SETIMAGE = 0x0172

If StringInStr($cmdlineraw, "/StopRightNow") Then Exit IniWrite($INI_FILE, "UPDATE", "StopNow", "True")

Global $sSize, $sUpdateFileName, $sNewVersion
If UBound(ProcessList("WE Updater.exe")) > 2 Then Exit MsgBox(64, "", "Oops, another process of Wallpaper Explorer Updater is running..!")


Global $iW = 400, $iH = 250
Global Const $hGUI = GUICreate("Wallpaper Explorer Installation", $iW, $iH, -1, -1, $WS_POPUP, $WS_EX_TOOLWINDOW)
Global Const $iPic = GUICtrlCreatePic("", 0, 0, $iW, $iH)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $hHBmp_BG, $hB, $iPerc = 0, $iSleep = 80, $s = 0, $t, $m = 0, $bAnimate = True
Global $arrText[3], $iText = 0, $hTimeInt = TimerInit()
_SetText("Checking", True)
GUIRegisterMsg($WM_TIMER, "PlayAnim")

_UpdateFunction()

GUIRegisterMsg($WM_TIMER, "")
_WinAPI_DeleteObject($hHBmp_BG)
_GDIPlus_Shutdown()
GUIDelete()
Exit

Func _CloseAllProc()
	Local $ARR_CHECK[3] = ["About.exe", "config.exe", "WallPaperExplorer.exe"]
	Local $sArrList
	For $i = 0 To UBound($ARR_CHECK) - 1
		Do
			ProcessClose($ARR_CHECK[$i])
		Until Not ProcessExists($ARR_CHECK[$i])
	Next
EndFunc   ;==>_CloseAllProc

Func _UpdateFunction()
	If FileExists($sLocalPackageDir) Then FileDelete($sLocalPackageDir)

	If InetGet($sURL_Check, $sLocalFileDir, 1, 0) = 0 Then
		_SetText("Error", False)
		MsgBox(16, "Wallpaper Explorer Error Message", "Oops,check your internet connections and try again.", $hGUI)
		Return
	EndIf

	$sNewVersion = IniRead($sLocalFileDir, "Info", "Version", "")
	$sUpdateFileName = IniRead($sLocalFileDir, "Info", "File", "")
	FileDelete($sLocalFileDir)
	Local $bIsWriteToFile = False
	If StringLeft($sUpdateFileName, 5) = "https" Then
		$sUpdateFileName = $sUpdateFileName
	Else
		$sUpdateFileName = $sURL_Folder & "/" & $sUpdateFileName
	EndIf

	$sSize = InetGetSize($sUpdateFileName)

	If Not $sSize = 0 Then
		$iPresentVersion = FileGetVersion($sFile)
		If StringReplace($sNewVersion, ".", "") > StringReplace($iPresentVersion, ".", "") Then
			If MsgBox(64 + 1, "Wallpaper Explorer Message", "Do you want to update Wallpaper Explorer program from version " & $iPresentVersion & _
					" to new version " & $sNewVersion & "?" & @CRLF & "Update package size: " & Round($sSize / (1024 * 1024), 4) & " MB", $hGUI) = 1 Then
				_DownloadUpdateGUI()
				Sleep(1000)
			EndIf
		Else
			MsgBox(64, "Wallpaper Explorer Message", "Thank you! This is the lastest version.", $hGUI)
			_SetText("Thanks!", True)
		EndIf
	Else
		MsgBox(16, "Wallpaper Explorer Error Message", "Sorry, some problems has occurred while downloading file.", $hGUI)
		_SetText("Error!", True)
	EndIf
EndFunc   ;==>_UpdateFunction

Func _DownloadUpdateGUI()
	_CloseAllProc()
	_SetText("Downloading", True)
	DllCall("user32.dll", "int", "SetTimer", "hwnd", $hGUI, "int", 0, "int", $iSleep, "int", 0)
	GUISetState(@SW_SHOW, $hGUI)

	Local Const $hDownload = InetGet($sUpdateFileName, $sLocalPackageDir, 1, 1)
	Local $iPercent = 0
	While True
		$iPercent = Round(InetGetInfo($hDownload, 0) / $sSize, 2)
		_SetText("Downloading" & @CRLF & $iPercent * 100 & "%")
		Sleep(100)
		If InetGetInfo($hDownload, 2) Then
			_SetText("Closing App", False)
			_CloseAllProc()
			_SetText("Installing", True)
			ShellExecuteWait($sLocalPackageDir, $ScriptDir)
			If IniRead(@TempDir & "\WE.ini", "UPDATE", "STATE", "OK") = "OK" Then
				_SetText("Successful!" & @CRLF & ":')", False)
			Else
				_SetText("Failed!" & @CRLF & ":'(", False)
			EndIf
			WinSetOnTop($hGUI, "", 1)
			Sleep(1000)
			FileDelete(@TempDir & "\WE.ini")
			FileDelete($sLocalPackageDir)
			ExitLoop
		ElseIf InetGetInfo($hDownload, 4) Then
			_SetText("Error!", True)
			ExitLoop
		EndIf

		If IniRead($INI_FILE, "UPDATE", "StopNow", "False") = "True" Then
			InetClose($hDownload)
			ExitLoop
		EndIf
	WEnd
EndFunc   ;==>_DownloadUpdateGUI



Func PlayAnim()
	$hHBmp_BG = _GDIPlus_MultiColorLoader($iW, $iH, $arrText[$iText])
	If TimerDiff($hTimeInt) >= 1000 Then
		$iText += 1
		$hTimeInt = TimerInit()
	EndIf
	If $iText >= UBound($arrText) Then $iText = 0
	$hB = GUICtrlSendMsg($iPic, $STM_SETIMAGE, $IMAGE_BITMAP, $hHBmp_BG)
	If $hB Then _WinAPI_DeleteObject($hB)
	_WinAPI_DeleteObject($hHBmp_BG)
EndFunc   ;==>PlayAnim

Func _SetText($sText = "Loading", $bPoint = True)
	If $bPoint = True Then
		$arrText[0] = $sText & "."
		$arrText[1] = $sText & ".."
		$arrText[2] = $sText & "..."
		$bAnimate = True
	Else
		$arrText[0] = $sText
		$arrText[1] = $sText
		$arrText[2] = $sText
		$bAnimate = False
	EndIf
EndFunc   ;==>_SetText

Func _GDIPlus_MultiColorLoader($iW, $iH, $sText = "", $sFont = "Segoe UI", $bHBitmap = True)
	Local Const $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
	Local Const $hGfx = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	_GDIPlus_GraphicsSetSmoothingMode($hGfx, 4 + (@OSBuild > 5999))
	_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 3)
	_GDIPlus_GraphicsSetPixelOffsetMode($hGfx, $GDIP_PIXELOFFSETMODE_HIGHQUALITY)
	_GDIPlus_GraphicsClear($hGfx, 0xFF232323)

	Local $iRadius = ($iW > $iH) ? $iH * 0.6 : $iW * 0.6

	Local Const $hPath = _GDIPlus_PathCreate()
	_GDIPlus_PathAddEllipse($hPath, ($iW - ($iRadius + 24)) / 2, ($iH - ($iRadius + 24)) / 2, $iRadius + 24, $iRadius + 24)

	Local $hBrush = _GDIPlus_PathBrushCreateFromPath($hPath)
	_GDIPlus_PathBrushSetCenterColor($hBrush, 0xFFFFFFFF)
	_GDIPlus_PathBrushSetSurroundColor($hBrush, 0x08101010)
	_GDIPlus_PathBrushSetGammaCorrection($hBrush, True)

	Local $aBlend[4][2] = [[3]]
	$aBlend[1][0] = 0 ;0% center color
	$aBlend[1][1] = 0 ;position = boundary
	$aBlend[2][0] = 0.33 ;70% center color
	$aBlend[2][1] = 0.1 ;10% of distance boundary->center point
	$aBlend[3][0] = 1 ;100% center color
	$aBlend[3][1] = 1 ;center point
	_GDIPlus_PathBrushSetBlend($hBrush, $aBlend)

	Local $aRect = _GDIPlus_PathBrushGetRect($hBrush)
	_GDIPlus_GraphicsFillRect($hGfx, $aRect[0], $aRect[1], $aRect[2], $aRect[3], $hBrush)

	_GDIPlus_PathDispose($hPath)
	_GDIPlus_BrushDispose($hBrush)

	Local Const $hBrush_Black = _GDIPlus_BrushCreateSolid(0xFF161616)
	_GDIPlus_GraphicsFillEllipse($hGfx, ($iW - ($iRadius + 10)) / 2, ($iH - ($iRadius + 10)) / 2, $iRadius + 10, $iRadius + 10, $hBrush_Black)

	Local Const $hBitmap_Gradient = _GDIPlus_BitmapCreateFromScan0($iRadius, $iRadius)
	Local Const $hGfx_Gradient = _GDIPlus_ImageGetGraphicsContext($hBitmap_Gradient)
	_GDIPlus_GraphicsSetSmoothingMode($hGfx_Gradient, 4 + (@OSBuild > 5999))
	Local Const $hMatrix = _GDIPlus_MatrixCreate()
	Local Static $r = 0
	_GDIPlus_MatrixTranslate($hMatrix, $iRadius / 2, $iRadius / 2)
	_GDIPlus_MatrixRotate($hMatrix, $r)
	_GDIPlus_MatrixTranslate($hMatrix, -$iRadius / 2, -$iRadius / 2)
	_GDIPlus_GraphicsSetTransform($hGfx_Gradient, $hMatrix)
	$r += 10
	Local Const $hBrush_Gradient = _GDIPlus_LineBrushCreate($iRadius, $iRadius / 2, $iRadius, $iRadius, 0xFF000000, 0xFF33CAFD, 1)
	_GDIPlus_LineBrushSetGammaCorrection($hBrush_Gradient)
	_GDIPlus_GraphicsFillEllipse($hGfx_Gradient, 0, 0, $iRadius, $iRadius, $hBrush_Gradient)
	_GDIPlus_GraphicsFillEllipse($hGfx_Gradient, 4, 4, $iRadius - 8, $iRadius - 8, $hBrush_Black)
	_GDIPlus_GraphicsDrawImageRect($hGfx, $hBitmap_Gradient, ($iW - $iRadius) / 2, ($iH - $iRadius) / 2, $iRadius, $iRadius)
	_GDIPlus_BrushDispose($hBrush_Gradient)
	_GDIPlus_BrushDispose($hBrush_Black)
	_GDIPlus_GraphicsDispose($hGfx_Gradient)
	_GDIPlus_BitmapDispose($hBitmap_Gradient)
	_GDIPlus_MatrixDispose($hMatrix)

	Local Const $hFormat = _GDIPlus_StringFormatCreate()
	Local Const $hFamily = _GDIPlus_FontFamilyCreate($sFont)
	Local Const $hFont = _GDIPlus_FontCreate($hFamily, $iRadius / 10)
	_GDIPlus_StringFormatSetAlign($hFormat, 1)
	_GDIPlus_StringFormatSetLineAlign($hFormat, 1)
	Local $tLayout = _GDIPlus_RectFCreate(0, 0, $iW, $iH)
	Local Static $iColor = 0x00, $iDir = 13
	Local $hBrush_txt = _GDIPlus_BrushCreateSolid(0xFF000000 + 0x010000 * $iColor + 0x0100 * $iColor + $iColor)
	_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hFont, $tLayout, $hFormat, $hBrush_txt)
	If $bAnimate Then
		$iColor += $iDir
		If $iColor > 0xFF Then
			$iColor = 0xFF
			$iDir *= -1
		ElseIf $iColor < 0x16 Then
			$iDir *= -1
			$iColor = 0x16
		EndIf
	Else
		$iColor = 0xFF
	EndIf
	_GDIPlus_BrushDispose($hBrush_txt)
	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_StringFormatDispose($hFormat)
	_GDIPlus_GraphicsDispose($hGfx)

	If $bHBitmap Then
		Local $hHBITMAP = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hBitmap)
		_GDIPlus_BitmapDispose($hBitmap)
		Return $hHBITMAP
	EndIf
	Return $hBitmap
EndFunc   ;==>_GDIPlus_MultiColorLoader

