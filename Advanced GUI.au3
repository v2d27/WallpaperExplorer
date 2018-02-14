#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GDIPlus.au3>
#include <WinAPISys.au3>
#include <WinAPIGdi.au3>
_GDIPlus_Startup()

$Form1 = GUICreate("Form1", 615, 437, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_SYSMENU, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_DLGMODALFRAME, $WS_EX_WINDOWEDGE, $WS_EX_LAYERED))
$Pic = GUICtrlCreateButton("This is some text!", 0, 0, 100, 20)
GUIRegisterMsg($WM_SIZING, "WM_SIZING")
GUIRegisterMsg($WM_EXITSIZEMOVE, "WM_SIZING")
_UpdateGUI($Form1)

Func _UpdateGUI($hWnd)
	Local Const $iPenSize = 4
	Local Const $iBkColor = 0xFF444444
	Local $iW = _WinAPI_GetWindowWidth($hWnd)
	Local $iH = _WinAPI_GetWindowHeight($hWnd)
	Local $Bitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
    Local $hBmpCtxt = _GDIPlus_ImageGetGraphicsContext($Bitmap)
    _GDIPlus_GraphicsSetSmoothingMode($hBmpCtxt, $GDIP_SMOOTHINGMODE_HIGHQUALITY)
	Local $iDwmColor = "0xF0" & StringTrimRight(StringRight(Hex(_WinAPI_DwmGetColorizationColor()), 6), 1) & "E"
	Local $hPen = _GDIPlus_PenCreate($iDwmColor, $iPenSize, 2)
	;ConsoleWrite(StringFormat("%s\n%s\n", $iDwmColor, Hex($iDwmColor)))
	Local $hBrush = _GDIPlus_BrushCreateSolid($iBkColor)
	_GDIPlus_GraphicsDrawRect($hBmpCtxt, $iPenSize, $iPenSize, $iW-$iPenSize*2, $iH-$iPenSize*2, $hPen)
	_GDIPlus_GraphicsFillRect($hBmpCtxt, $iPenSize, $iPenSize, $iW-$iPenSize*2, $iH-$iPenSize*2, $hBrush)
	SetBitmap($hWnd, $Bitmap, 255)
EndFunc

GUISetState(@SW_SHOW)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd

Func WM_SIZING($hWnd, $iMsg, $wParam, $lParam)
	;ConsoleWrite(StringFormat("Client W:H = %s:%s, Window W:H = %s:%s\n", _WinAPI_GetClientWidth($hWnd), _WinAPI_GetClientHeight($hWnd), _
	;_WinAPI_GetWindowWidth($hWnd), _WinAPI_GetWindowHeight($hWnd)))
	Return _UpdateGUI($hWnd)

EndFunc

Func SetBitmap($hGUI, $hImage, $iOpacity)
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
	_WinAPI_UpdateLayeredWindow($hGUI, $hScrDC, 0, $pSize, $hMemDC, $pSource, 0, $pBlend, $ULW_ALPHA)
	_WinAPI_ReleaseDC(0, $hScrDC)
	_WinAPI_SelectObject($hMemDC, $hOld)
	_WinAPI_DeleteObject($hBitmap)
	_WinAPI_DeleteDC($hMemDC)
EndFunc ;==>SetBitmap