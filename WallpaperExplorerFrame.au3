
Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
Global Const $sVersion = StringTrimRight(FileGetVersion($sFile), 2)
Global Const $sTitle = "Wallpaper Explorer v" & $sVersion
Global Const $pIDWE = ShellExecute($sFile, "", @ScriptDir, "")
WinWaitActive($sTitle)
Global Const $hWpEx = WinGetHandle($sTitle)
ConsoleWrite(StringFormat("Title is %s \n", $hWpEx))

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 1114, 657, -1, -1, $WS_SIZEBOX)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

_WinAPI_SetParent($hWpEx, $Form1)
WinMove($hWpEx, "", 5, 5, _WinAPI_GetClientWidth($Form1), _WinAPI_GetClientHeight($Form1))
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			ExitLoop

	EndSwitch
WEnd
WinClose($hWpEx)
ProcessClose($pIDWE)