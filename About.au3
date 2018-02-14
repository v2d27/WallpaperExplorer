#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\icon_info.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=About Wallpaper Explorer v3.7.6
#AutoIt3Wrapper_Res_Fileversion=3.7.6.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.6.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#NoTrayIcon
#include <GUIConstantsEx.au3>
#include <GuiRichEdit.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <GDIPlus.au3>
#include <StaticConstants.au3>
#include <WinAPISys.au3>
#include <_GUIResourcePic.au3>

Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
Global Const $sVersion = StringTrimRight(FileGetVersion($sFile), 2)
Global Const $sLoading = @ScriptDir & "\Resources\Loading.gif"
Global Const $INI_FILE = @ScriptDir & "\Wallpaper Explorer.ini"
Global Const $TEMP_UPDATE_EXE = @TempDir & "\WE Updater.exe"

If WinExists("About Wallpaper Explorer v" & $sVersion) Then Exit WinActivate("About Wallpaper Explorer v" & $sVersion)
If StringInStr($CmdLineRaw, "/start") Then
	Example()
Else
	Exit MsgBox(16, "Wallpaper Explorer Error Message", "Oops, Wrong parameters...!")
EndIf

Func Example()
	FileInstall(".\WE Updater.exe", $TEMP_UPDATE_EXE)
	Local Const $__sText = "" & _
			"                                                            Product name" & @CRLF & _
			"                                                            Product version" & @CRLF & _
			"                                                            Date publish" & @CRLF & _
			"                                                            Author" & @CRLF & _
			"                                                            Email" & @CRLF & @CRLF & @CRLF & _
			"When we use the computer, we often choose desktop background in slide show mode, so we can" & @CRLF & _
			"view many pictures stored in our hard disk drivers and work more comfortably. Some of these" & @CRLF & _
			"picture make us do like whereas others do not. Sometimes, we want to delete them immediately" & @CRLF & _
			"from desktop without finding them in their folders to save our time. Unfortunately, none of " & @CRLF & _
			"windows versions supports this feature. And now, with Wallpaper Explorer, you can do it only by" & @CRLF & _
			"one mouse click. I think this program is so useful for you. Let try it..." & @CRLF & _
			"I express my heartfelt thanks to these web-pages below, who help me to write this program: " & @CRLF & _
			"        www.stackoverflow.com*" & @CRLF & _
			"        www.autoitscript.com*" & @CRLF & _
			"        www.msdn.microsoft.com" & @CRLF & _
			"        www.codeproject.com" & @CRLF
	Local $arrFileTime = FileGetTime($sFile, 0, 0)
	Local Const $__sTextInfo = ": Wallpaper Explorer" & @CRLF & _
			": " & $sVersion & @CRLF & _
			": " & StringFormat("%s/%s/%s  %s:%s", $arrFileTime[2], $arrFileTime[1], $arrFileTime[0], $arrFileTime[3], $arrFileTime[4]) & @CRLF & _
			": Vu Van Duc" & @CRLF & _
			": ducduc08@gmail.com"




	Local $__AboutGUI = GUICreate("About Wallpaper Explorer v" & $sVersion, 640, 400)
	GUISetBkColor(0xFFFFFF, $__AboutGUI)
	Local $iW = _WinAPI_GetClientWidth($__AboutGUI)
	Local $iH = _WinAPI_GetClientHeight($__AboutGUI) - 45

	Local $__hPic = GUICtrlGetHandle(GUICtrlCreatePic('', 0, 0, $iW, $iH))

	Local $__hOK = GUICtrlCreateButton("OK", $iW - 130, $iH + 10, 120, 25)
	GUICtrlCreateButton("s", $iW + 100, $iH, 50, 20)
	GUICtrlSetState(-1, $GUI_FOCUS)

	Local $__hCheckForUpdate = GUICtrlCreateButton("Check for Update", $iW - 130 - 130, $iH + 10, 120, 25)
	GUICtrlCreateButton("s", $iW + 100, $iH, 50, 20)

	Local $__hLoading = GUICtrlCreatePic('', $iW - 130 - 130, $iH + 10 + 2, 20, 20)
	_GUICtrlPic_SetImage(-1, $sLoading)
	GUICtrlSetState(-1, $GUI_HIDE)
	Local $__hState = GUICtrlCreateLabel("Checking...", $iW - 130 - 130 + 20 + 5, $iH + 13, 120 - 25, 25)
	GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
	GUICtrlSetState(-1, $GUI_HIDE)
	Local $__hPercent = GUICtrlCreateLabel('0%', $iW - 130 - 130 - 15, $iH + 10 + 3, 35, 25, $SS_RIGHT)
	GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
	GUICtrlSetState(-1, $GUI_HIDE)

	Local $__hLink = GUICtrlCreateLabel("License User Agreement", 10, $iH + 20, 130, 20, BitOR($SS_NOTIFY, $SS_LEFT))
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetColor(-1, 0x3399FF)
	GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
	GUICtrlSetCursor(-1, 0)


	_GDIPlus_Startup()

	Local $hScanBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
	Local $hContext = _GDIPlus_ImageGetGraphicsContext($hScanBitmap)
	_GDIPlus_GraphicsSetSmoothingMode($hContext, $GDIP_SMOOTHINGMODE_HIGHQUALITY)
	_GDIPlus_GraphicsClear($hContext, 0xFFFFFFFF) ;clear bitmap with color white


	$hBrush = _GDIPlus_BrushCreateSolid(0xFF444444)
	$hFormat = _GDIPlus_StringFormatCreate()
	$hFamily = _GDIPlus_FontFamilyCreate("Segoe UI")
	$hFont = _GDIPlus_FontCreate($hFamily, 14, 0, 0)
	$tLayout = _GDIPlus_RectFCreate(10, 10, 0, 0)
	$aInfo = _GDIPlus_GraphicsMeasureString($hContext, $__sText, $hFont, $tLayout, $hFormat)
	_GDIPlus_GraphicsDrawStringEx($hContext, $__sText, $hFont, $aInfo[0], $hFormat, $hBrush)
	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_StringFormatDispose($hFormat)
	_GDIPlus_BrushDispose($hBrush)

	$hBrush = _GDIPlus_BrushCreateSolid(0xFF444444)
	$hFormat = _GDIPlus_StringFormatCreate()
	$hFamily = _GDIPlus_FontFamilyCreate("Segoe UI")
	$hFont = _GDIPlus_FontCreate($hFamily, 14, 0, 0)
	$tLayout = _GDIPlus_RectFCreate(365, 10, 0, 0)
	$aInfo = _GDIPlus_GraphicsMeasureString($hContext, $__sText, $hFont, $tLayout, $hFormat)
	_GDIPlus_GraphicsDrawStringEx($hContext, $__sTextInfo, $hFont, $aInfo[0], $hFormat, $hBrush)

	$hLOGO = _GDIPlus_BitmapCreateFromFile(@ScriptDir & "\Resources\logo.png")
	_GDIPlus_GraphicsDrawImage($hContext, $hLOGO, 70, 0)

	Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hScanBitmap)
	_SendMessage($__hPic, 0x0172, $IMAGE_BITMAP, $hBitmap) ;Set bitmap to control
	GUISetState()


	Local $bIsCheckingUpdate = False
	Local $sNowStatus[6] = ["Checking", "Checking.", "Checking..", "Checking...", "Checking....", "Checking....."], $iStatus = 0, $hTimeInt = 0, $bProcExit = False
	Local $pIDSubProc = 0
	While 1
		$nMsg = GUIGetMsg()
		If $bIsCheckingUpdate Then
			If TimerDiff($hTimeInt) >= 200 Then
				If Not ProcessExists($pIDSubProc) Then
					$bIsCheckingUpdate = False
					$nMsg = $GUI_EVENT_CLOSE
					$bProcExit = True
				EndIf
				GUICtrlSetData($__hState, $sNowStatus[$iStatus])
				$iStatus += 1
				If $iStatus = UBound($sNowStatus) Then $iStatus = 0
				$hTimeInt = TimerInit()
			EndIf
		EndIf
		Switch $nMsg
			Case $__hCheckForUpdate
				GUICtrlSetState($__hCheckForUpdate, $GUI_HIDE)
				GUICtrlSetState($__hOK, $GUI_DISABLE)
				GUICtrlSetState($__hLoading, $GUI_SHOW)
				GUICtrlSetState($__hState, $GUI_SHOW)
				$pIDSubProc = ShellExecute($TEMP_UPDATE_EXE, @ScriptDir)
				$bIsCheckingUpdate = True
			Case $__hLink
				ShellExecute(@ScriptDir & "\License User Agreement.rtf", "", "", "", @SW_MAXIMIZE)
			Case $GUI_EVENT_CLOSE, $__hOK
				If $bIsCheckingUpdate Then
					If MsgBox(64 + 4, "Wallpaper Explorer Update", "Do you want to cancel update?", 0, $__AboutGUI) = 7 Then ContinueCase
				EndIf
				ShellExecute($TEMP_UPDATE_EXE, "/StopRightNow")
				GUICtrlSetState($__hCheckForUpdate, $GUI_SHOW)
				GUICtrlSetState($__hOK, $GUI_ENABLE)
				GUICtrlSetState($__hLoading, $GUI_HIDE)
				GUICtrlSetState($__hState, $GUI_HIDE)
				If Not $bProcExit Then
					ExitLoop
				Else
					$bProcExit = False
				EndIf
		EndSwitch
	WEnd

	; Clean up resources
	_WinAPI_DeleteObject($hBitmap)
	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_StringFormatDispose($hFormat)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_GraphicsDispose($hContext)
	_GDIPlus_BitmapDispose($hScanBitmap)
	_GDIPlus_Shutdown()
EndFunc   ;==>Example
