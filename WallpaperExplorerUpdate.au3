#RequireAdmin
#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\LOGO.ICO
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Wallpaper Explorer Update Package
#AutoIt3Wrapper_Res_Fileversion=3.7.3.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.3.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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

$temp = FileInstall(".\Resources\MENU.PNG", $ScriptDir & "\Resources\MENU.PNG", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\MENU.PNG' & @CRLF
$temp = FileInstall(".\Resources\logo.png", $ScriptDir & "\Resources\logo.png", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\logo.png' & @CRLF
$temp = FileInstall(".\Resources\logo.ico", $ScriptDir & "\Resources\logo.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\logo.ico' & @CRLF
$temp = FileInstall(".\Resources\Loading.gif", $ScriptDir & "\Resources\Loading.gif", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\Loading.gif' & @CRLF
$temp = FileInstall(".\Resources\icon_share.ico", $ScriptDir & "\Resources\icon_share.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_share.ico' & @CRLF
$temp = FileInstall(".\Resources\icon_picture.ico", $ScriptDir & "\Resources\icon_picture.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_picture.ico' & @CRLF
$temp = FileInstall(".\Resources\icon_info.ico", $ScriptDir & "\Resources\icon_info.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_info.ico' & @CRLF
$temp = FileInstall(".\Resources\icon_folder.ico", $ScriptDir & "\Resources\icon_folder.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_folder.ico' & @CRLF
$temp = FileInstall(".\Resources\ICON_PROP.ICO", $ScriptDir & "\Resources\ICON_PROP.ICO", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\ICON_PROP.ICO' & @CRLF
$temp = FileInstall(".\Resources\icon_details.ico", $ScriptDir & "\Resources\icon_details.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_details.ico' & @CRLF
$temp = FileInstall(".\Resources\icon_delete.ico", $ScriptDir & "\Resources\icon_delete.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_delete.ico' & @CRLF
$temp = FileInstall(".\Resources\icon_copy.ico", $ScriptDir & "\Resources\icon_copy.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_copy.ico' & @CRLF
$temp = FileInstall(".\Resources\HELP_USER.PNG", $ScriptDir & "\Resources\HELP_USER.PNG", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\HELP_USER.PNG' & @CRLF
$temp = FileInstall(".\Resources\icon_001.ico", $ScriptDir & "\Resources\icon_001.ico", 1)
If $temp = 0 Then $arrResult &= ' - Can not create file Resources\icon_001.ico' & @CRLF
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
	MsgBox(64, "Update Message", "Update Wallpaper Explorer program to new version " & FileGetVersion($ScriptDir & "\WallPaperExplorer.exe") & " successfully!" & @CRLF & @CRLF & _
			"Click OK to open program")

	ShellExecute($ScriptDir & "\WallPaperExplorer.exe")
EndIf

RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Explorer", "DisplayVersion", "REG_SZ", FileGetVersion($ScriptDir & "\WallPaperExplorer.exe"))

Func _CloseAllProc()
	Local $ARR_CHECK[3] = ["About.exe", "config.exe", "WallPaperExplorer.exe"]
	Local $sArrList
	For $i = 0 To UBound($ARR_CHECK) - 1
		Do
			ProcessClose($ARR_CHECK[$i])
		Until Not ProcessExists($ARR_CHECK[$i])
	Next
EndFunc   ;==>_CloseAllProc


