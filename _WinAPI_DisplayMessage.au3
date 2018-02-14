
#include<WinAPI.au3>
#include<GDIPlus.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
_GDIPlus_Startup()
_WinAPI_DisplayMessage("Hello, I will always find you...")






Func _WinAPI_DisplayMessage($sText, $iTime = 3000)
	Local Const $iW = StringLen($sText)*10+40, $iH = 30
	Local Const $iX = @DesktopWidth/2-$iW/2, $IY = @DesktopHeight/2 -$iH/2
	Local Const $iBackgroundColor = 0x00F239
	Local $hDesktop = _WinAPI_GetDesktopWindow()
	Local $__hControl = GUICreate("", $iW, $iH, $iX, $IY, $WS_POPUP, BitOR($WS_EX_TOPMOST,$WS_EX_MDICHILD, $WS_EX_TOOLWINDOW), $hDesktop)
	GUISetBkColor(0xFF0000)
	Local $hDC = _WinAPI_GetDC($__hControl)
	Local $iStyle = BitOR($DT_CENTER, $DT_SINGLELINE, $DT_TOP, $DT_HIDEPREFIX, $DT_VCENTER)

	GUISetState(@SW_SHOW, $__hControl)
	Local $tRect = DllStructCreate($tagRECT)
	DllStructSetData($tRect, "left", 0)
	DllStructSetData($tRect, "top", 0)
	DllStructSetData($tRect, "right", $iW)
	DllStructSetData($tRect, "bottom", $iH)
	Local $pRect = DllStructGetPtr($tRect)
	Local $hBrushBackground = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($iBackgroundColor))

	$hBrushBorder = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($iBackgroundColor))
	$__hBUTTON_FONT_Caption = _WinAPI_CreateFont(18, 0, 0, 0, 400, False, False, False, $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $PROOF_QUALITY, 0, 'Segoe UI')
	$hOldFont = _WinAPI_SelectObject($hDC, $__hBUTTON_FONT_Caption)
	_WinAPI_SetTextColor($hDC, 0xFFFFFF)

	_WinAPI_SetBkMode($hDC, $TRANSPARENT)
	Local $hTimeInt = TimerInit()
	Do
		_WinAPI_SetBkColor($hDC, $iBackgroundColor)
		_WinAPI_DrawText($hDC, $sText, $tRect, $iStyle)
	Until TimerDiff($hTimeInt)>= $iTime

EndFunc