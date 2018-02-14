#include-once
opt("WinWaitDelay", 10)
#include "UDFGlobalID.au3"
#include "WinAPI.au3"
#include "Constants.au3"
#include "WindowsConstants.au3"
#include "WinAPIGdi.au3"
#include "FontConstants.au3"
#include "WinAPIRes.au3"

; #CONSTANTS# ===================================================================================================================
Global Const $__METROBUTTON_CONSTANT_ClassName = "MetroButton"
Global Const $__METROBUTTON_CONSTANT_SIZEOFPTR = DllStructGetSize(DllStructCreate("ptr"))

; Notifications for WM_COMMAND
Global Const $WM_MOUSEINSIDE = 27081997
Global Const $WM_MOUSEOUTSIDE = 27081998

Global Const $__METROBUTTON_CONSTANT_tagWNDCLASSEX = "uint Size; " & _
		"uint style; " & _
		"ptr WndProc; " & _
		"int ClsExtra; " & _
		"int WndExtra; " & _
		"HANDLE Instance; " & _
		"HANDLE Icon; " & _
		"HANDLE Cursor; " & _
		"HANDLE Background; " & _
		"ptr MenuName; " & _
		"ptr ClassName; " & _
		"HANDLE IconSm"


Global Const $__ARR_BUTTON_COLOR[3][2] = [ [0xFFFFFF, 0x171717], [0xDBDBDB, 0x171717], [0xB3B3B3, 0x171717] ] ; The color of The Background and Text in Normal, Hover, Clicked state

;Metro button animation
Global $xBackup = 0, $iDeltaMoveButton = 1, $hButton_Moved = ""
Global $tMousePoint
Global Const $WM_NEXTANIMATION = 27081999
Global Const $WM_ENDANIMATION = 27082000
Global $__STOP_ANIMATION = False
Global $__IsAnimation = False
Global $__MouseInsideTimeStart = TimerInit()




; Only function Mouse hover ================================================================================================
Global Const $MOUSE_BUTTON_CALLBACK = DllCallbackRegister("__MOUSE_BUTTON_PROC", "long", "int;wparam;lparam")
Global Const $H_BUTTON_MOD = _WinAPI_GetModuleHandle(0)
Global Const $H_BUTTON_HOOK = _WinAPI_SetWindowsHookEx($WH_MOUSE, DllCallbackGetPtr($MOUSE_BUTTON_CALLBACK), $H_BUTTON_MOD, _WinAPI_GetCurrentThreadId())
Func __MOUSE_BUTTON_PROC($nCode, $wParam, $lParam)
	If $__IsAnimation Or $nCode < 0 Then Return _WinAPI_CallNextHookEx($H_BUTTON_HOOK, $nCode, $wParam, $lParam)
	Local Static $LAST_WINDOWS = ""
	;Local $arrCursorInfo = _WinAPI_GetCursorInfo()
	$tMousePoint = DllStructCreate("struct; long X;long Y; endstruct")
	DllStructSetData($tMousePoint, "X", MouseGetPos(0))
	DllStructSetData($tMousePoint, "Y", MouseGetPos(1))
	Switch $wParam
		Case $WM_MOUSEMOVE
			If TimerDiff($__MouseInsideTimeStart) >= 20 Then
				Local $PRESENT_WINDOWS = _WinAPI_WindowFromPoint($tMousePoint)
				If $PRESENT_WINDOWS <> $LAST_WINDOWS Then
					_WinAPI_PostMessage($LAST_WINDOWS, $WM_COMMAND, $WM_MOUSEOUTSIDE, $__METROBUTTON_CONSTANT_ClassName)
					;ConsoleWrite(StringFormat("Outside: %s\n", StringLeft(_WinAPI_GetWindowText($PRESENT_WINDOWS), 1)))
					$LAST_WINDOWS = ""
				EndIf
;~ 				ConsoleWrite(StringFormat("Inside: %s\n", StringLeft(_WinAPI_GetWindowText($PRESENT_WINDOWS), 1)))
				If _WinAPI_IsClassName($PRESENT_WINDOWS, $__METROBUTTON_CONSTANT_ClassName) Then
					If $PRESENT_WINDOWS <> $LAST_WINDOWS Then
						_WinAPI_PostMessage($PRESENT_WINDOWS, $WM_COMMAND, $WM_MOUSEINSIDE, $__METROBUTTON_CONSTANT_ClassName)
						;ConsoleWrite(StringFormat("Inside: %s\n", StringLeft(_WinAPI_GetWindowText($PRESENT_WINDOWS), 1)))
						$LAST_WINDOWS = $PRESENT_WINDOWS
					EndIf
				EndIf
				$__MouseInsideTimeStart = TimerInit()
			EndIf
	EndSwitch
	Return _WinAPI_CallNextHookEx($H_BUTTON_HOOK, $nCode, $wParam, $lParam)
EndFunc
;===========================================================================================================================

Global $__sBUTTON_TEXT = ""
Global $__hBUTTON_IMAGE = 0
Global $__sBUTTON_CANNOTUPDATE = ""
Global $__bBUTTON_ALLCANUPDATE = True




Func _GUICtrlMetroButton_Create($hWnd, $sText, $iX, $iY, $iWidth = 60, $iHeight = 40, $sIcon = "")
	Local $nCtrlID
	Local Static $iAtom = 0
	If Not IsHWnd($hWnd) Or Not WinExists($hWnd) Then Return SetError(1, 0, 0) ; Parent window is not valid
	; Register the class
	If Not $iAtom Then
		Local $tClassName, $tWC, $aRes

		If Not OnAutoItExitRegister("__GUICtrlMetroButton_OnExit") Then Return SetError(2, 0, 0) ; Unable to register OnExit function

		Global $H_BUTTON_CALLBACK = DllCallbackRegister("__GUICtrlMetroButton_WndProc", "lresult", "hwnd;uint;wparam;lparam")
		If @error Then Return SetError(3, 0, 0) ; Unable to register window procedure. Possibly obfuscated?

		$tClassName = DllStructCreate("wchar[" & (StringLen($__METROBUTTON_CONSTANT_ClassName) + 1) & "]")
		DllStructSetData($tClassName, 1, $__METROBUTTON_CONSTANT_ClassName)

		$tWC = DllStructCreate($__METROBUTTON_CONSTANT_tagWNDCLASSEX)

		DllStructSetData($tWC, "Size", DllStructGetSize($tWC))
		DllStructSetData($tWC, "style", 3) ; BitOR($CS_HREDRAW, $CS_VREDRAW))
		DllStructSetData($tWC, "WndProc", DllCallbackGetPtr($H_BUTTON_CALLBACK))
		DllStructSetData($tWC, "ClsExtra", 0)
		DllStructSetData($tWC, "WndExtra", 16 * (@AutoItX64 + 1)) ; ptr IconFull, ptr IconEmpty, ptr Cursor, ptr CurrentWnd
		DllStructSetData($tWC, "Instance", _WinAPI_GetModuleHandle(0))
		DllStructSetData($tWC, "Icon", 0)
		DllStructSetData($tWC, "Cursor", _WinAPI_LoadImage(0, 32512, $IMAGE_CURSOR, 0, 0, BitOR($LR_DEFAULTCOLOR, $LR_SHARED)))
		DllStructSetData($tWC, "Background", 0)
		DllStructSetData($tWC, "MenuName", 0)
		DllStructSetData($tWC, "ClassName", DllStructGetPtr($tClassName))
		DllStructSetData($tWC, "IconSm", 0)

		$aRes = DllCall("user32.dll", "word", "RegisterClassExW", "ptr", DllStructGetPtr($tWC))
		If @error Or Not $aRes[0] Then Return SetError(4, @error, 0) ; Registering the class failed. @extended contains the DllCall @error code.

		$iAtom = $aRes[0]
	EndIf

	$nCtrlID = __UDF_GetNextGlobalID($hWnd)
	If @error Then Return SetError(5, @error, 0)

	$hButton = _WinAPI_CreateWindowEx(0, $__METROBUTTON_CONSTANT_ClassName, $sText, BitOR($__UDFGUICONSTANT_WS_CHILD, $__UDFGUICONSTANT_WS_VISIBLE), $iX, $iY, $iWidth, $iHeight, $hWnd, $nCtrlID)
	If Not $hButton Then Return SetError(6, @error, 0)
	$__sBUTTON_TEXT &= "|" & $hButton & '="' & $sText & '"|'

	If $sIcon <> "" And FileExists($sIcon) Then
		Local $hIcon = _WinAPI_LoadImage(0, $sIcon, $IMAGE_ICON, 32, 32, BitOR($LR_DEFAULTCOLOR, $LR_LOADFROMFILE, $LR_LOADTRANSPARENT))
		$__sBUTTON_TEXT &= "|" & $hButton & '=="' & $hIcon & '"|'
	EndIf

	Return $hButton
EndFunc   ;==>_GUICtrlMetroButton_Create

Func _GUICtrlMetroButton_Startup()
	Local $arrButton = StringRegExp($__sBUTTON_TEXT, '0x[0-9A-F]{8}', 3)
	Local $hLast = 0
	If IsArray($arrButton) Then
		For $i = 0 To UBound($arrButton)-1
			If _WinAPI_IsWindow($arrButton[$i]) And $arrButton[$i] <> $hLast Then
				_SetBkButton($arrButton[$i], 0) ;Normal state
				$hLast = $arrButton[$i]
			EndIf
		Next
	EndIf
EndFunc

Func __GUICtrlMetroButton_WndProc($hWnd, $iMsg, $wParam, $lParam)
;~ 	ConsoleWrite(StringFormat("$hWnd =%s, $iMsg = %s, $wParam = %s, $lParam = %s", $hWnd, $iMsg, $wParam, $lParam) & @CRLF)
	If $__bBUTTON_ALLCANUPDATE Then
		If $iMsg = $WM_COMMAND Then
			If $wParam = $WM_MOUSEINSIDE Then; Mouse inside
				_SetBkButton($hWnd, 1) ; Hover state
			ElseIf $wParam = $WM_MOUSEOUTSIDE Then; Mouse outside
				_SetBkButton($hWnd, 0) ; Normal state
			ElseIf $wParam = $WM_NEXTANIMATION Then
					If Not $__STOP_ANIMATION Then
						_GUICtrlMetroButton_Animation($hWnd)
					Else
						$__STOP_ANIMATION = False
					EndIf
			ElseIf $wParam = $WM_ENDANIMATION Then
					$__STOP_ANIMATION = True
			EndIf
			Return _WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam)
		EndIf
		Switch $iMsg
			Case $WM_LBUTTONDOWN
				_SetBkButton($hWnd, 2) ; Clicked state
			Case $WM_LBUTTONUP
				_SetBkButton($hWnd, 1) ; Hover state
				ConsoleWrite("WM_LBUTTONUP Handle:" & $hWnd & @CRLF)
				_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_LBUTTONUP, $hWnd)
			Case $WM_PAINT, $WM_MOVE
				If Not StringInStr($__sBUTTON_CANNOTUPDATE, $hWnd) Then
					_SetBkButton($hWnd, 0) ; Normal state
				EndIf
			Case $WM_DESTROY
				_WinAPI_DestroyWindow($hWnd)
		EndSwitch
	EndIf
	Return _WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__GUICtrlFinder_WndProc

Func _GUICtrlMetroButton_Animation($hWnd)
	$__IsAnimation = True
	__WinMove($hWnd, _GUICtrlMetroButton_GetX(_WinAPI_GetParent($hWnd), $hWnd), _GUICtrlMetroButton_GetY(_WinAPI_GetParent($hWnd), $hWnd), _WinAPI_GetWindowWidth($hWnd), _WinAPI_GetWindowHeight($hWnd))
	$__IsAnimation = False
EndFunc


Func __WinMove($hWnd, $x, $y, $w, $h)
	Local $iOldDelta = 1
	;ConsoleWrite(StringFormat("%s: %s - %s - %s - %s| $iDeltaMoveButton = %s ", $hWnd, $x, $y, $w, $h, $iDeltaMoveButton) & @CRLF)

	If Not StringInStr($hButton_Moved, $hWnd) Then ;Showing
		_WinAPI_MoveWindow($hWnd, $x-$iDeltaMoveButton, $y, $w, $h, False) ;Face out
		If Abs($x) >= $w Then ; End animation
			$hButton_Moved &= $hWnd
			$iDeltaMoveButton = 1
			_WinAPI_MoveWindow($hWnd, $xBackup-$w, $y, $w, $h)
			;ConsoleWrite(StringFormat("-> End animation face out. $hWnd = %s " & @CRLF & @CRLF, $hWnd))

			;Call next animation button
			_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_NEXTANIMATION, $hWnd)

			;end function
			Return
		EndIf

		; Animation
		$iOldDelta = $iDeltaMoveButton
		If Abs($x) <= Int($w/30) Then
			$iDeltaMoveButton = 0.1
		ElseIf Abs($x) > Int($w/30) And Abs($x) <= Int($w/15) Then
			$iDeltaMoveButton = 0.5
		ElseIf Abs($x) > Int($w*0.75) Then
			$iDeltaMoveButton = 45
		Else
			$iDeltaMoveButton = 20
		EndIf
	ElseIf StringInStr($hButton_Moved, $hWnd) Then ;Hiding
		_WinAPI_MoveWindow($hWnd, $x+$iDeltaMoveButton, $y, $w, $h, False) ;Face in
		If $x >= $xBackup Then ; End animation
			$hButton_Moved = StringReplace($hButton_Moved, $hWnd, "") ; Delete button moved variably
			$iDeltaMoveButton = 1

			;Set button to original position
			_WinAPI_MoveWindow($hWnd, $xBackup, $y, $w, $h)
			;ConsoleWrite(StringFormat("-> End animation face in. $hWnd = %s " & @CRLF & @CRLF, $hWnd))

			;Call next animation button
			_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_NEXTANIMATION, $hWnd)

			;end function
			Return
		EndIf

		; Animation
		$iOldDelta = $iDeltaMoveButton
		If Abs($x) <= Int($w/15) Then
			$iDeltaMoveButton = 0.1
		ElseIf Abs($x) > Int($w/15) And Abs($x) <= Int($w/8) Then
			$iDeltaMoveButton = 0.3
		ElseIf Abs($x) > Int($w*0.75) Then
			$iDeltaMoveButton = 35
		Else
			$iDeltaMoveButton = 12
		EndIf
	EndIf

	If Not StringInStr($hButton_Moved, $hWnd) Then
		__WinMove($hWnd, $x-$iOldDelta, $y, $w, $h)
	Else
		__WinMove($hWnd, $x+$iOldDelta, $y, $w, $h)
	EndIf
EndFunc



















Func _GUICtrlMetroButton_SetUpdateState($hWnd, $bState = True)
	If $bState Then ; Can update
		$__sBUTTON_CANNOTUPDATE = StringReplace($__sBUTTON_CANNOTUPDATE, $hWnd & " ", "")
	Else ; Can not update
		$__sBUTTON_CANNOTUPDATE &= $hWnd & " "
	EndIf
EndFunc

Func _GUICtrlMetroButton_SetAllUpdateState($bState = True)
	If $bState Then ; Can update
		$__bBUTTON_ALLCANUPDATE = True
	Else ; Can not update
		$__bBUTTON_ALLCANUPDATE = False
	EndIf
EndFunc


Func _GUICtrlMetroButton_GetX($hGUI, $hWnd)
	Local $tRect = _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
	Return DllStructGetData($tRect, "left")
EndFunc

Func _GUICtrlMetroButton_GetY($hGUI, $hWnd)
	Local $tRect = _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
	Return DllStructGetData($tRect, "top")
EndFunc

Func _GUICtrlMetroButton_GetWidth($hWnd)
	Return _WinAPI_GetWindowWidth($hWnd)
EndFunc

Func _GUICtrlMetroButton_GetHeight($hWnd)
	Return _WinAPI_GetWindowHeight($hWnd)
EndFunc

Func _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
	Local $tPoint = DllStructCreate("int X;int Y")
	Local $tRect = _WinAPI_GetWindowRect($hWnd)
	Local $iLeftBk = DllStructGetData($tRect, "left"), $iTopBk = DllStructGetData($tRect, "top")
	DllStructSetData($tPoint, "X", $iLeftBk)
	DllStructSetData($tPoint, "Y", $iTopBk)
	_WinAPI_ScreenToClient($hGUI, $tPoint)
	DllStructSetData($tRect, "left",DllStructGetData($tPoint, "X"))
	DllStructSetData($tRect, "top",DllStructGetData($tPoint, "Y"))
	DllStructSetData($tRect, "right", DllStructGetData($tRect, "right") - ($iLeftBk - DllStructGetData($tPoint, "X")))
	DllStructSetData($tRect, "bottom", DllStructGetData($tRect, "bottom") - ($iTopBk - DllStructGetData($tPoint, "Y")))
	Return $tRect
EndFunc

Func __GUICtrlMetroButton_OnExit()
	; Close all windows
	Local $a = _WinAPI_EnumWindows()
	For $i = 1 To $a[0][0]
		If $a[$i][1] = $__METROBUTTON_CONSTANT_ClassName And WinGetProcess($a[$i][0]) = @AutoItPID Then
			_WinAPI_PostMessage($a[$i][0], $WM_CLOSE, 0, 0)
		EndIf
	Next
	DllCallbackFree($H_BUTTON_CALLBACK)
	DllCallbackFree($MOUSE_BUTTON_CALLBACK)
	_WinAPI_UnhookWindowsHookEx($H_BUTTON_HOOK)
	DllCall("user32.dll", "int", "UnregisterClassW", _
			"wstr", $__METROBUTTON_CONSTANT_ClassName, _
			"HANDLE", _WinAPI_GetModuleHandle(0))
EndFunc   ;==>__GUICtrlMetroButton_OnExit


;Normal, hover, clicked
; 0, 1, 2
Func _SetBkButton($hWnd, $idArrColor)
	Local Static $idLastArrColor = $idArrColor

	Local $tRect = DllStructCreate("int;int;int;int")
	DllStructSetData($tRect, 1, 0)
	DllStructSetData($tRect, 2, 0)
	DllStructSetData($tRect, 3, _WinAPI_GetWindowWidth($hWnd))
	DllStructSetData($tRect, 4, _WinAPI_GetWindowHeight($hWnd))
	Local $hBrushBackground, $hBrushBorder, $__hBUTTON_FONT_Caption, $hOldFont
	Local $sText = StringRegExp($__sBUTTON_TEXT, "\|" & $hWnd & '="(.+?)"\|', 3)
	If IsArray($sText) Then
		If StringLen($sText[0]) = 0 Or StringIsSpace($sText[0]) Then
			$sText = ""
		Else
			$sText = $sText[0]
		EndIf
	Else
		$sText = ""
	EndIf

	Local $hDC = _WinAPI_GetDC($hWnd), $pRect = DllStructGetPtr($tRect)
	Local $hIcon = StringRegExp($__sBUTTON_TEXT, "\|" & $hWnd & '=="(.+?)"\|', 3) ; For drawing icon
	Local $iStart, $iEnd, $iStep, $iColor

	; For drawing button's text
	$hBrushBorder = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($__ARR_BUTTON_COLOR[$idArrColor][1]))
	$__hBUTTON_FONT_Caption = _WinAPI_CreateFont(18, 0, 0, 0, 400, False, False, False, $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $PROOF_QUALITY, 0, 'Segoe UI')
	$hOldFont = _WinAPI_SelectObject($hDC, $__hBUTTON_FONT_Caption)
	_WinAPI_SetTextColor($hDC, $__ARR_BUTTON_COLOR[$idArrColor][1])

	If $idArrColor = 1 And $idLastArrColor = 0 Then ; From hover to normal | From gray to white
;~; 		ConsoleWrite("-> Starting fill background..." & @CRLF)
		$iStart = 0
		$iEnd = 99
		$iStep = $iEnd*20
;~ 	ElseIf $idArrColor =
	Else ; original color
		$iStart = 0
		$iEnd = 0
		$iStep = 1
	EndIf

	Local $iStyle = ""
;~ 	Local $hTime= TimerInit()
	For $i = $iStart To $iEnd Step $iStep
		; Drawing background
		If $__ARR_BUTTON_COLOR[$idArrColor][0] <> "" Then
			$iColor = _WinAPI_ColorAdjustLuma($__ARR_BUTTON_COLOR[$idArrColor][0], $i, True)
;~ 			$iColor = _WinAPI_ColorAdjustLuma($__ARR_BUTTON_COLOR[$idArrColor][0], $iEnd, True)
			; Drawing frame
			$hBrushBackground = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($iColor))
			_WinAPI_FillRect($hDC, $pRect, $hBrushBackground)
			_WinAPI_SetBkColor($hDC, _WinAPI_SwitchColor($iColor))
			_WinAPI_DeleteObject($hBrushBackground)
		Else
			If $__ARR_BUTTON_COLOR[$idArrColor][0] = "" Then 	;Set transparent background
				_WinAPI_SetStretchBltMode($hDC,  $COLORONCOLOR)
				_WinAPI_SetBkMode($hDC, $TRANSPARENT)
			EndIf
		EndIf

		; Drawing icon and text at the same
		If IsArray($hIcon) And StringLen($hIcon[0]) > 0 Then
			_WinAPI_DrawIconEx($hDC, 15, DllStructGetData($tRect, 4)/2-32/2, $hIcon[0])
			$iStyle = BitOR($DT_LEFT, $DT_VCENTER, $DT_SINGLELINE, $DT_TOP, $DT_HIDEPREFIX)
			DllStructSetData($tRect, 1, 10+40)
		Else
			$iStyle = BitOR($DT_CENTER, $DT_SINGLELINE, $DT_TOP, $DT_HIDEPREFIX, $DT_VCENTER)
		EndIf
		_WinAPI_DrawText($hDC, $sText, $tRect, $iStyle)
	Next
;~ 	ConsoleWrite(StringFormat("Time loop: %s" & @CRLF, Round(TimerDiff($hTime), 4)))

	$idLastArrColor = $idArrColor
	; Clean up resources
	_WinAPI_DeleteObject($hOldFont)
	_WinAPI_DeleteObject($hBrushBorder)
	;_WinAPI_DeleteObject($hBrushBackground)
	_WinAPI_DeleteObject($__hBUTTON_FONT_Caption)
	_WinAPI_ReleaseDC($hWnd, $hDC)
EndFunc


;~ Test()
;~ Func Test()
;~ 	$hGUI = GUICreate("_GUICtrlMetroButton_Create Example", 400, 200)
;~ 	GUISetBkColor(0x444444)
;~ 	$hButton1 = _GUICtrlMetroButton_Create($hGUI, "Button Metro 1", 10, 50, 200, 60, @ScriptDir & "\ICON_COPY.ICO")
;~ 	ConsoleWrite("$Button1 @error = " & @error & @TAB & "Result: " & $hButton1 & @CRLF)
;~ 	$hButton2 = _GUICtrlMetroButton_Create($hGUI, "Button Metro 2", 10, 110, 200, 60, @ScriptDir & "\ICON_DELETE.ICO")
;~ 	ConsoleWrite("$Button2 @error = " & @error & @TAB & "Result: " & $hButton2 & @CRLF)
;~ 	$hBtn = GUICtrlCreateButton("Read & Save!", 40, 6, 150, 30)
;~ 	GUISetState()
;~ 	_GUICtrlMetroButton_Startup()
;~ 	While True
;~ 		Switch GUIGetMsg()
;~ 			Case -3
;~ 				ExitLoop
;~ 		EndSwitch
;~ 	WEnd
;~ EndFunc
