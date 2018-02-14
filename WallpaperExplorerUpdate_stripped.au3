#RequireAdmin
#NoTrayIcon
Global $ScriptDir = ""
If StringLen($cmdlineraw) <> 0 Then $ScriptDir = $cmdlineraw
If Not FileExists($ScriptDir & "\WallPaperExplorer.exe") Then
$ScriptDir = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Explorer", "InstallLocation")
If @error Then $ScriptDir = @ProgramFilesDir & "\Wallpaper Explorer"
EndIf
_CloseAllProc()
DirCreate($ScriptDir)
DirCreate($ScriptDir & "\Resources")
FileDelete($ScriptDir & "\WE Updater.exe")
FileDelete($ScriptDir & "\Wallpaper Explorer.ini")
If Not FileExists($ScriptDir) Then Exit MsgBox(16, "Error Message", "Can not create folder " & $ScriptDir & @CRLF & "Update Wallpaper Explorer will exit.")
If Not FileExists($ScriptDir & "\Resources") Then Exit MsgBox(16, "Error Message", "Can not create folder " & $ScriptDir & "\Resources" & @CRLF & "Update Wallpaper Explorer will exit.")
Global $arrResult = ""
$temp = FileInstall(".\About.exe", $ScriptDir & "\About.exe", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file About.exe' & @CRLF
$temp = FileInstall(".\config.exe", $ScriptDir & "\config.exe", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file config.exe' & @CRLF
$temp = FileInstall(".\WallPaperExplorer.exe", $ScriptDir & "\WallPaperExplorer.exe", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file WallPaperExplorer.exe' & @CRLF
$temp = FileInstall(".\License User Agreement.rtf", $ScriptDir & "\License User Agreement.rtf", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file License User Agreement.rtf' & @CRLF
If StringLen($arrResult) <> 0 Then
MsgBox(16, "Wallpaper Explorer Error Report", "Update Wallpaper Explorer program might not complete due to some errors below:" & @CRLF & @CRLF & $arrResult & @CRLF)
IniWrite(@TempDir & "\WE.ini", "UPDATE", "STATE", "FAILED")
Else
IniWrite(@TempDir & "\WE.ini", "UPDATE", "STATE", "OK")
MsgBox(64, "Update Message", "Update Wallpaper Explorer program to new version " & FileGetVersion($ScriptDir & "\WallPaperExplorer.exe") & " successfully!" & @CRLF & @CRLF & "Click OK to open program")
ShellExecute($ScriptDir & "\WallPaperExplorer.exe")
EndIf
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Explorer", "DisplayVersion", "REG_SZ", FileGetVersion($ScriptDir & "\WallPaperExplorer.exe"))
Func _CloseAllProc()
Local $ARR_CHECK[3] = ["About.exe", "config.exe", "WallPaperExplorer.exe"]
For $i = 0 To UBound($ARR_CHECK) - 1
Do
ProcessClose($ARR_CHECK[$i])
Until Not ProcessExists($ARR_CHECK[$i])
Next
EndFunc
