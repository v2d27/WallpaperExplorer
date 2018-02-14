#NoTrayIcon
#RequireAdmin
Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
If Not FileExists($sFile) Then Exit MsgBox(16, "Wallpaper Explorer Error Message", 'Sorry, we can not find "WallPaperExplorer.exe" file.')
If StringInStr($cmdlineraw, "/uninstall") Then
Local $a = RegDelete("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer")
Exit
EndIf
If RegRead("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer\command", "") = '"' & $sFile & '"' Then
Exit
EndIf
RegWrite("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer", "icon", "REG_SZ", '"' & $sFile & '", 0')
RegWrite("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer", "position", "REG_SZ", "middle")
RegWrite("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer\command", "", "REG_SZ", '"' & $sFile & '"')
If RegRead("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer\command", "") = '"' & $sFile & '"' Then
MsgBox(64, "Wallpaper Explorer Message", "Successful installation!")
Else
MsgBox(16, "Wallpaper Explorer Error Message", "Unable to install setting... Please re-open this file as administrator. Thank you! :'( ")
EndIf
