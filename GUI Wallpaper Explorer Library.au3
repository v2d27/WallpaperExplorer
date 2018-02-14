
#include <SendMessage.au3>
#include <StaticConstants.au3>

Func _CreateWindowBoder($hWnd)
	Local $Form1 = GUICreate(WinGetTitle($hWnd), 100, 100, -1, -1, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hWnd)
	GUISetBkColor(0x010101)
	GUISetFont(12, 400, 0, "Segoe UI")
	GUISetState(@SW_SHOWNOACTIVATE)
	_WinAPI_SetLayeredWindowAttributes($Form1, 0x010101, 255)

	Global Const $iBoderColor = 0x30B3FF
	Global Const $iCaptionColor = 0xFAFAFA
	Global Const $iBoderSize = 2
	Global Const $iCaptionSize = 26
	Global Const $iCW = _WinAPI_GetClientWidth($Form1), $iCH = _WinAPI_GetClientHeight($Form1)

	Global $hCaption = GUICtrlCreateLabel("My program name", $iBoderSize, $iBoderSize, $iCW - $iBoderSize * 2, $iCaptionSize, BitOR(0x01, 0x0200), $GUI_WS_EX_PARENTDRAG)
	GUICtrlSetBkColor(-1, $iCaptionColor)
	GUICtrlSetColor(-1, 0x44444F)
	GUICtrlSetResizing(-1, BitOR($GUI_DOCKLEFT, $GUI_DOCKHEIGHT, $GUI_DOCKTOP, $GUI_DOCKRIGHT))


	;Draw left
	GUICtrlCreateLabel('', 0, 0, $iBoderSize, $iCH, -1, $GUI_WS_EX_PARENTDRAG)
	GUICtrlSetBkColor(-1, $iBoderColor)
 	GUICtrlSetResizing(-1, BitOR($GUI_DOCKLEFT, $GUI_DOCKWIDTH))
	;GUICtrlSetCursor(-1, 13)
	;Draw top
	GUICtrlCreateLabel('', 0, 0, $iCW, $iBoderSize, -1, $GUI_WS_EX_PARENTDRAG)
	GUICtrlSetBkColor(-1, $iBoderColor)
	GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKHEIGHT))
	;GUICtrlSetCursor(-1, 11)
	;Draw right
	GUICtrlCreateLabel('', $iCW - $iBoderSize, 0, $iBoderSize, $iCH, -1, $GUI_WS_EX_PARENTDRAG)
	GUICtrlSetBkColor(-1, $iBoderColor)
 	GUICtrlSetResizing(-1, BitOR($GUI_DOCKRIGHT, $GUI_DOCKWIDTH, $GUI_DOCKTOP))
	;GUICtrlSetCursor(-1, 13)
	;Draw bottom
	GUICtrlCreateLabel('', 0, $iCH - $iBoderSize, $iCW, $iBoderSize, -1, $GUI_WS_EX_PARENTDRAG)
	GUICtrlSetBkColor(-1, $iBoderColor)
	GUICtrlSetResizing(-1, BitOR($GUI_DOCKBOTTOM, $GUI_DOCKLEFT, $GUI_DOCKHEIGHT))
	;GUICtrlSetCursor(-1, 11)

	Return $Form1
EndFunc   ;==>_CreateWindowBoder

Func _SetBitmapAdjust($hWnd, $hBitmap, $aAdjust = 0)
	If Not IsHWnd($hWnd) Then
		$hWnd = GUICtrlGetHandle($hWnd)
		If Not $hWnd Then
			Return 0
		EndIf
	EndIf

	Local $tAdjust = 0
	If IsArray($aAdjust) Then
		$tAdjust = _WinAPI_CreateColorAdjustment($aAdjust[9], $aAdjust[10], $aAdjust[0], $aAdjust[1], $aAdjust[2], $aAdjust[3], $aAdjust[4], $aAdjust[5], $aAdjust[6], $aAdjust[7], $aAdjust[8])
	EndIf
	$hBitmap = _WinAPI_AdjustBitmap($hBitmap, -1, -1, $HALFTONE, $tAdjust)
	If @error Then
		Return 0
	EndIf
	Local $hPrev = _SendMessage($hWnd, $STM_SETIMAGE, $IMAGE_BITMAP, $hBitmap)
	If $hPrev Then
		_WinAPI_DeleteObject($hPrev)
	EndIf
	$hPrev = _SendMessage($hWnd, $STM_GETIMAGE)
	If $hPrev <> $hBitmap Then
		_WinAPI_DeleteObject($hBitmap)
	EndIf
	Return 1
EndFunc   ;==>_SetBitmapAdjust

Func _WinAPI_ConvertPointClientToScreen($hWnd, $iX, $iY)
	Local $tPoint = DllStructCreate("int X;int Y")
	Local $arrPoint[2]
	DllStructSetData($tPoint, "X", $iX)
	DllStructSetData($tPoint, "Y", $iY)
	_WinAPI_ClientToScreen($hWnd,$tPoint)
	$arrPoint[0] = DllStructGetData($tPoint, "X")
	$arrPoint[1] = DllStructGetData($tPoint, "Y")
	Return $arrPoint
EndFunc

Func _WinAPI_ConvertPointScreenToClient($hWnd, $iX, $iY)
	Local $tPoint = DllStructCreate("int X;int Y")
	Local $arrPoint[2]
	DllStructSetData($tPoint, "X", $iX)
	DllStructSetData($tPoint, "Y", $iY)
	_WinAPI_ScreenToClient($hWnd,$tPoint)
	$arrPoint[0] = DllStructGetData($tPoint, "X")
	$arrPoint[1] = DllStructGetData($tPoint, "Y")
	Return $arrPoint
EndFunc

Func SetBitmap($hWnd, $hImage, $iOpacity)
	Local $hScrDC, $hMemDC, $hBitmap, $hOld, $pSize, $tSize, $pSource, $tSource, $pBlend, $tBlend

	$hScrDC = _WinAPI_GetDC(0)
	$hMemDC = _WinAPI_CreateCompatibleDC($hScrDC)
	$hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImage)
	$hOld = _WinAPI_SelectObject($hMemDC, $hBitmap)
	$tSize = DllStructCreate($tagSIZE)
	$pSize = DllStructGetPtr($tSize)
	DllStructSetData($tSize, "X", _GDIPlus_ImageGetWidth($hImage))
	DllStructSetData($tSize, "Y", _GDIPlus_ImageGetHeight($hImage))
	$tSource = DllStructCreate($tagPOINT)
	$pSource = DllStructGetPtr($tSource)
	$tBlend = DllStructCreate($tagBLENDFUNCTION)
	$pBlend = DllStructGetPtr($tBlend)
	DllStructSetData($tBlend, "Alpha", $iOpacity)
	DllStructSetData($tBlend, "Format", 1)
	_WinAPI_UpdateLayeredWindow($hWnd, $hScrDC, 0, $pSize, $hMemDC, $pSource, 0, $pBlend, $ULW_ALPHA)
	_WinAPI_ReleaseDC(0, $hScrDC)
	_WinAPI_SelectObject($hMemDC, $hOld)
	_WinAPI_DeleteObject($hBitmap)
	_WinAPI_DeleteDC($hMemDC)
EndFunc ;==>SetBitmap


Func _GetPathWallpaper()
	Local $data
	If @OSVersion = "WIN_7" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_XP" Or @OSVersion = "WIN_XPe" Then
		$data = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\General", "WallpaperSource")
		If Not FileExists($data) Then
			$data = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper")
		EndIf
	Else
		$data = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache")
		$data = BinaryToString(StringReplace("0x" & StringTrimLeft($data, 50), "00", ""))
	EndIf

	If FileExists($data) Then
		;ConsoleWrite(StringFormat("$sWallpaper=%s\n", $data))
		Return $data
	EndIf
	MsgBox(14, "Missing error!", 'Can not find location of wallpaper' & @CRLF & 'Program is unable to start correctly.')
	Exit
EndFunc