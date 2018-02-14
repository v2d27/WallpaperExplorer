#include <GUIConstantsEx.au3>
#include <GUISysLink.au3>
#include <GUIStatusBar.au3>
#include <WindowsConstants.au3>

$Text = 'The SysLink controls provides a convenient way to embed hypertext links in a window. For more information, click <A HREF="http://msdn.microsoft.com/en-us/library/bb760706(VS.85).aspx">here</A>.' & @CRLF & @CRLF & _
		'To learn how to use the SysLink controls in AutoIt, click <A HREF="http://www.autoitscript.com/forum/">here</A>.'

$hForm = GUICreate("MyGUI", 428, 160)
$hStatusBar = _GUICtrlStatusBar_Create($hForm)
$hSysLink = _GUICtrlSysLink_Create($hForm, $Text, 10, 10, 408, 54)
$Dummy = GUICtrlCreateDummy()
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUISetState()

While 1
	_SysLink_Over()
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Dummy
			ShellExecute(_GUICtrlSysLink_GetItemUrl($hSysLink, GUICtrlRead($Dummy)))
	EndSwitch
WEnd

Func _SysLink_Over()

	Local $Link = _GUICtrlSysLink_HitTest($hSysLink)

	For $i = 0 To 1
		If $i = $Link Then
			If Not _GUICtrlSysLink_GetItemHighlighted($hSysLink, $i) Then
				_GUICtrlSysLink_SetItemHighlighted($hSysLink, $i, 1)
				_GUICtrlStatusBar_SetText($hStatusBar, _GUICtrlSysLink_GetItemUrl($hSysLink, $i))
			EndIf
		Else
			If _GUICtrlSysLink_GetItemHighlighted($hSysLink, $i) Then
				_GUICtrlSysLink_SetItemHighlighted($hSysLink, $i, 0)
				_GUICtrlStatusBar_SetText($hStatusBar, "")
			EndIf
		EndIf
	Next
EndFunc   ;==>_SysLink_Over

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)

	Local $tNMLINK = DllStructCreate($tagNMLINK, $lParam)
	Local $hFrom = DllStructGetData($tNMLINK, "hWndFrom")
	Local $ID = DllStructGetData($tNMLINK, "Code")

	Switch $hFrom
		Case $hSysLink
			Switch $ID
				Case $NM_CLICK, $NM_RETURN
					GUICtrlSendToDummy($Dummy, DllStructGetData($tNMLINK, "Link"))
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY
