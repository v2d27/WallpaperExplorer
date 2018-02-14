#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Config Wallpaper Explorer v3.7.6
#AutoIt3Wrapper_Res_Fileversion=3.7.6.0
#AutoIt3Wrapper_Res_ProductVersion=3.7.6.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2015 by VuVanDuc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Productname|WallpaperExplorer
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
Global Const $INI_FILE = @ScriptDir & "\Wallpaper Explorer.ini"
Global Const $sURL_Folder = "http://googledrive.com/host/0B4S2pTQPGLo6NjZjZFBmOWpaMEk"
Global Const $sURL_Check = $sURL_Folder & "/WallpaperExplorerUpdate.txt"
Global Const $sLocalFileDir = @TempDir & "\CheckForUpdate.ini"
Global Const $sLocalPackageDir = @TempDir & "\WEUpdateInstaller.exe"
Global $sSize, $sUpdateFileName, $sNewVersion

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