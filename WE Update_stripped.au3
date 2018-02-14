#NoTrayIcon
Global Const $GDIP_PXF32ARGB = 0x0026200A
Global Const $GDIP_SMOOTHINGMODE_DEFAULT = 0
Global Const $GDIP_SMOOTHINGMODE_ANTIALIAS8X8 = 5
Global Const $GDIP_PIXELOFFSETMODE_HIGHQUALITY = 2
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagGDIPRECTF = "struct;float X;float Y;float Width;float Height;endstruct"
Global Const $tagGDIPSTARTUPINPUT = "uint Version;ptr Callback;bool NoThread;bool NoCodecs"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $IMAGE_BITMAP = 0
Func _WinAPI_DeleteObject($hObject)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Global $__g_hGDIPBrush = 0
Global $__g_hGDIPDll = 0
Global $__g_iGDIPRef = 0
Global $__g_iGDIPToken = 0
Global $__g_bGDIP_V1_0 = True
Func _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight, $iPixelFormat = $GDIP_PXF32ARGB, $iStride = 0, $pScan0 = 0)
Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $iWidth, "int", $iHeight, "int", $iStride, "int", $iPixelFormat, "struct*", $pScan0, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[6]
EndFunc
Func _GDIPlus_BitmapCreateHBITMAPFromBitmap($hBitmap, $iARGB = 0xFF000000)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateHBITMAPFromBitmap", "handle", $hBitmap, "handle*", 0, "dword", $iARGB)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_BitmapDispose($hBitmap)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDisposeImage", "handle", $hBitmap)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_BrushCreateSolid($iARGB = 0xFF000000)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateSolidFill", "int", $iARGB, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_BrushDispose($hBrush)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteBrush", "handle", $hBrush)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_FontCreate($hFamily, $fSize, $iStyle = 0, $iUnit = 3)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateFont", "handle", $hFamily, "float", $fSize, "int", $iStyle, "int", $iUnit, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[5]
EndFunc
Func _GDIPlus_FontDispose($hFont)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteFont", "handle", $hFont)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_FontFamilyCreate($sFamily, $pCollection = 0)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateFontFamilyFromName", "wstr", $sFamily, "ptr", $pCollection, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[3]
EndFunc
Func _GDIPlus_FontFamilyDispose($hFamily)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteFontFamily", "handle", $hFamily)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsClear($hGraphics, $iARGB = 0xFF000000)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGraphicsClear", "handle", $hGraphics, "dword", $iARGB)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsDispose($hGraphics)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteGraphics", "handle", $hGraphics)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsDrawImageRect($hGraphics, $hImage, $nX, $nY, $nW, $nH)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDrawImageRect", "handle", $hGraphics, "handle", $hImage, "float", $nX, "float", $nY, "float", $nW, "float", $nH)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsDrawStringEx($hGraphics, $sString, $hFont, $tLayout, $hFormat, $hBrush)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDrawString", "handle", $hGraphics, "wstr", $sString, "int", -1, "handle", $hFont, "struct*", $tLayout, "handle", $hFormat, "handle", $hBrush)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsFillEllipse($hGraphics, $nX, $nY, $nWidth, $nHeight, $hBrush = 0)
__GDIPlus_BrushDefCreate($hBrush)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipFillEllipse", "handle", $hGraphics, "handle", $hBrush, "float", $nX, "float", $nY, "float", $nWidth, "float", $nHeight)
__GDIPlus_BrushDefDispose()
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsFillRect($hGraphics, $nX, $nY, $nWidth, $nHeight, $hBrush = 0)
__GDIPlus_BrushDefCreate($hBrush)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipFillRectangle", "handle", $hGraphics, "handle", $hBrush, "float", $nX, "float", $nY, "float", $nWidth, "float", $nHeight)
__GDIPlus_BrushDefDispose()
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsSetPixelOffsetMode($hGraphics, $iPixelOffsetMode)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetPixelOffsetMode", "handle", $hGraphics, "int", $iPixelOffsetMode)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsSetSmoothingMode($hGraphics, $iSmooth)
If $iSmooth < $GDIP_SMOOTHINGMODE_DEFAULT Or $iSmooth > $GDIP_SMOOTHINGMODE_ANTIALIAS8X8 Then $iSmooth = $GDIP_SMOOTHINGMODE_DEFAULT
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetSmoothingMode", "handle", $hGraphics, "int", $iSmooth)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsSetTextRenderingHint($hGraphics, $iTextRenderingHint)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetTextRenderingHint", "handle", $hGraphics, "int", $iTextRenderingHint)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_GraphicsSetTransform($hGraphics, $hMatrix)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetWorldTransform", "handle", $hGraphics, "handle", $hMatrix)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_ImageGetGraphicsContext($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageGraphicsContext", "handle", $hImage, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_LineBrushCreate($nX1, $nY1, $nX2, $nY2, $iARGBClr1, $iARGBClr2, $iWrapMode = 0)
Local $tPointF1, $tPointF2, $aResult
$tPointF1 = DllStructCreate("float;float")
$tPointF2 = DllStructCreate("float;float")
DllStructSetData($tPointF1, 1, $nX1)
DllStructSetData($tPointF1, 2, $nY1)
DllStructSetData($tPointF2, 1, $nX2)
DllStructSetData($tPointF2, 2, $nY2)
$aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateLineBrush", "struct*", $tPointF1, "struct*", $tPointF2, "uint", $iARGBClr1, "uint", $iARGBClr2, "int", $iWrapMode, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[6]
EndFunc
Func _GDIPlus_LineBrushSetGammaCorrection($hLineGradientBrush, $bUseGammaCorrection = True)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetLineGammaCorrection", "handle", $hLineGradientBrush, "int", $bUseGammaCorrection)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_MatrixCreate()
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateMatrix", "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[1]
EndFunc
Func _GDIPlus_MatrixDispose($hMatrix)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteMatrix", "handle", $hMatrix)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_MatrixRotate($hMatrix, $fAngle, $bAppend = False)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipRotateMatrix", "handle", $hMatrix, "float", $fAngle, "int", $bAppend)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_MatrixTranslate($hMatrix, $fOffsetX, $fOffsetY, $bAppend = False)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipTranslateMatrix", "handle", $hMatrix, "float", $fOffsetX, "float", $fOffsetY, "int", $bAppend)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathAddEllipse($hPath, $nX, $nY, $nWidth, $nHeight)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipAddPathEllipse", "handle", $hPath, "float", $nX, "float", $nY, "float", $nWidth, "float", $nHeight)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathBrushCreateFromPath($hPath)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreatePathGradientFromPath", "handle", $hPath, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_PathBrushGetRect($hPathGradientBrush)
Local $tRECTF = DllStructCreate($tagGDIPRECTF)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetPathGradientRect", "handle", $hPathGradientBrush, "struct*", $tRECTF)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Local $aRectF[4]
For $iI = 1 To 4
$aRectF[$iI - 1] = DllStructGetData($tRECTF, $iI)
Next
Return $aRectF
EndFunc
Func _GDIPlus_PathBrushSetBlend($hPathGradientBrush, $aBlends)
Local $iCount = $aBlends[0][0]
Local $tFactors = DllStructCreate("float[" & $iCount & "]")
Local $tPositions = DllStructCreate("float[" & $iCount & "]")
For $iI = 1 To $iCount
DllStructSetData($tFactors, 1, $aBlends[$iI][0], $iI)
DllStructSetData($tPositions, 1, $aBlends[$iI][1], $iI)
Next
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetPathGradientBlend", "handle", $hPathGradientBrush, "struct*", $tFactors, "struct*", $tPositions, "int", $iCount)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathBrushSetCenterColor($hPathGradientBrush, $iARGB)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetPathGradientCenterColor", "handle", $hPathGradientBrush, "uint", $iARGB)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathBrushSetGammaCorrection($hPathGradientBrush, $bUseGammaCorrection)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetPathGradientGammaCorrection", "handle", $hPathGradientBrush, "int", $bUseGammaCorrection)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathBrushSetSurroundColor($hPathGradientBrush, $iARGB)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetPathGradientSurroundColorsWithCount", "handle", $hPathGradientBrush, "uint*", $iARGB, "int*", 1)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_PathCreate($iFillMode = 0)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreatePath", "int", $iFillMode, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_PathDispose($hPath)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeletePath", "handle", $hPath)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_RectFCreate($nX = 0, $nY = 0, $nWidth = 0, $nHeight = 0)
Local $tRECTF = DllStructCreate($tagGDIPRECTF)
DllStructSetData($tRECTF, "X", $nX)
DllStructSetData($tRECTF, "Y", $nY)
DllStructSetData($tRECTF, "Width", $nWidth)
DllStructSetData($tRECTF, "Height", $nHeight)
Return $tRECTF
EndFunc
Func _GDIPlus_Shutdown()
If $__g_hGDIPDll = 0 Then Return SetError(-1, -1, False)
$__g_iGDIPRef -= 1
If $__g_iGDIPRef = 0 Then
DllCall($__g_hGDIPDll, "none", "GdiplusShutdown", "ulong_ptr", $__g_iGDIPToken)
DllClose($__g_hGDIPDll)
$__g_hGDIPDll = 0
EndIf
Return True
EndFunc
Func _GDIPlus_Startup($sGDIPDLL = Default, $bRetDllHandle = False)
$__g_iGDIPRef += 1
If $__g_iGDIPRef > 1 Then Return True
If $sGDIPDLL = Default Then $sGDIPDLL = "gdiplus.dll"
$__g_hGDIPDll = DllOpen($sGDIPDLL)
If $__g_hGDIPDll = -1 Then
$__g_iGDIPRef = 0
Return SetError(1, 2, False)
EndIf
Local $sVer = FileGetVersion($sGDIPDLL)
$sVer = StringSplit($sVer, ".")
If $sVer[1] > 5 Then $__g_bGDIP_V1_0 = False
Local $tInput = DllStructCreate($tagGDIPSTARTUPINPUT)
Local $tToken = DllStructCreate("ulong_ptr Data")
DllStructSetData($tInput, "Version", 1)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdiplusStartup", "struct*", $tToken, "struct*", $tInput, "ptr", 0)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
$__g_iGDIPToken = DllStructGetData($tToken, "Data")
If $bRetDllHandle Then Return $__g_hGDIPDll
Return SetExtended($sVer[1], True)
EndFunc
Func _GDIPlus_StringFormatCreate($iFormat = 0, $iLangID = 0)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateStringFormat", "int", $iFormat, "word", $iLangID, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[3]
EndFunc
Func _GDIPlus_StringFormatDispose($hFormat)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteStringFormat", "handle", $hFormat)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_StringFormatSetAlign($hStringFormat, $iFlag)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetStringFormatAlign", "handle", $hStringFormat, "int", $iFlag)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_StringFormatSetLineAlign($hStringFormat, $iStringAlign)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetStringFormatLineAlign", "handle", $hStringFormat, "int", $iStringAlign)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func __GDIPlus_BrushDefCreate(ByRef $hBrush)
If $hBrush = 0 Then
$__g_hGDIPBrush = _GDIPlus_BrushCreateSolid()
$hBrush = $__g_hGDIPBrush
EndIf
EndFunc
Func __GDIPlus_BrushDefDispose($iCurError = @error, $iCurExtended = @extended)
If $__g_hGDIPBrush <> 0 Then
_GDIPlus_BrushDispose($__g_hGDIPBrush)
$__g_hGDIPBrush = 0
EndIf
Return SetError($iCurError, $iCurExtended)
EndFunc
Global Const $GUI_DISABLE = 128
Global Const $WS_POPUP = 0x80000000
Global Const $WS_EX_TOOLWINDOW = 0x00000080
Global Const $WM_TIMER = 0x0113
If StringLen($cmdlineraw) = 0 Then Exit
_GDIPlus_Startup()
Global Const $ScriptDir = $cmdlineraw
Global Const $sFile = $ScriptDir & "\WallPaperExplorer.exe"
If Not FileExists($sFile) Then Exit
Global Const $INI_FILE = $ScriptDir & "\Wallpaper Explorer.ini"
Global Const $sURL_Folder = "http://googledrive.com/host/0B4S2pTQPGLo6NjZjZFBmOWpaMEk"
Global Const $sURL_Check = $sURL_Folder & "/WallpaperExplorerUpdate.txt"
Global Const $sLocalFileDir = @TempDir & "\CheckForUpdate.ini"
Global Const $sLocalPackageDir = @TempDir & "\WallpaperExplorerUpdate.exe"
Global Const $STM_SETIMAGE = 0x0172
If StringInStr($cmdlineraw, "/StopRightNow") Then Exit IniWrite($INI_FILE, "UPDATE", "StopNow", "True")
Global $sSize, $sUpdateFileName, $sNewVersion
If UBound(ProcessList("WE Updater.exe")) > 2 Then Exit MsgBox(64, "", "Oops, another process of Wallpaper Explorer Updater is running..!")
Global $iW = 400, $iH = 250
Global Const $hGUI = GUICreate("Wallpaper Explorer Installation", $iW, $iH, -1, -1, $WS_POPUP, $WS_EX_TOOLWINDOW)
Global Const $iPic = GUICtrlCreatePic("", 0, 0, $iW, $iH)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $hHBmp_BG, $hB, $iPerc = 0, $iSleep = 80, $s = 0, $t, $m = 0, $bAnimate = True
Global $arrText[3], $iText = 0, $hTimeInt = TimerInit()
_SetText("Checking", True)
GUIRegisterMsg($WM_TIMER, "PlayAnim")
_UpdateFunction()
GUIRegisterMsg($WM_TIMER, "")
_WinAPI_DeleteObject($hHBmp_BG)
_GDIPlus_Shutdown()
GUIDelete()
Exit
Func _CloseAllProc()
Local $ARR_CHECK[3] = ["About.exe", "config.exe", "WallPaperExplorer.exe"]
For $i = 0 To UBound($ARR_CHECK) - 1
Do
ProcessClose($ARR_CHECK[$i])
Until Not ProcessExists($ARR_CHECK[$i])
Next
EndFunc
Func _UpdateFunction()
If FileExists($sLocalPackageDir) Then FileDelete($sLocalPackageDir)
If InetGet($sURL_Check, $sLocalFileDir, 1, 0) = 0 Then
_SetText("Error", False)
MsgBox(16, "Wallpaper Explorer Error Message", "Oops,check your internet connections and try again.", $hGUI)
Return
EndIf
$sNewVersion = IniRead($sLocalFileDir, "Info", "Version", "")
$sUpdateFileName = IniRead($sLocalFileDir, "Info", "File", "")
FileDelete($sLocalFileDir)
If StringLeft($sUpdateFileName, 5) = "https" Then
$sUpdateFileName = $sUpdateFileName
Else
$sUpdateFileName = $sURL_Folder & "/" & $sUpdateFileName
EndIf
$sSize = InetGetSize($sUpdateFileName)
If Not $sSize = 0 Then
$iPresentVersion = FileGetVersion($sFile)
If StringReplace($sNewVersion, ".", "") > StringReplace($iPresentVersion, ".", "") Then
If MsgBox(64 + 1, "Wallpaper Explorer Message", "Do you want to update Wallpaper Explorer program from version " & $iPresentVersion & " to new version " & $sNewVersion & "?" & @CRLF & "Update package size: " & Round($sSize /(1024 * 1024), 4) & " MB", $hGUI) = 1 Then
_DownloadUpdateGUI()
Sleep(1000)
EndIf
Else
MsgBox(64, "Wallpaper Explorer Message", "Thank you! This is the lastest version.", $hGUI)
_SetText("Thanks!", True)
EndIf
Else
MsgBox(16, "Wallpaper Explorer Error Message", "Sorry, some problems has occurred while downloading file.", $hGUI)
_SetText("Error!", True)
EndIf
EndFunc
Func _DownloadUpdateGUI()
_CloseAllProc()
_SetText("Downloading", True)
DllCall("user32.dll", "int", "SetTimer", "hwnd", $hGUI, "int", 0, "int", $iSleep, "int", 0)
GUISetState(@SW_SHOW, $hGUI)
Local Const $hDownload = InetGet($sUpdateFileName, $sLocalPackageDir, 1, 1)
Local $iPercent = 0
While True
$iPercent = Round(InetGetInfo($hDownload, 0) / $sSize, 2)
_SetText("Downloading" & @CRLF & $iPercent * 100 & "%")
Sleep(100)
If InetGetInfo($hDownload, 2) Then
_SetText("Closing App", False)
_CloseAllProc()
_SetText("Installing", True)
ShellExecuteWait($sLocalPackageDir, $ScriptDir)
If IniRead(@TempDir & "\WE.ini", "UPDATE", "STATE", "OK") = "OK" Then
_SetText("Successful!" & @CRLF & ":')", False)
Else
_SetText("Failed!" & @CRLF & ":'(", False)
EndIf
WinSetOnTop($hGUI, "", 1)
Sleep(1000)
FileDelete(@TempDir & "\WE.ini")
FileDelete($sLocalPackageDir)
ExitLoop
ElseIf InetGetInfo($hDownload, 4) Then
_SetText("Error!", True)
ExitLoop
EndIf
If IniRead($INI_FILE, "UPDATE", "StopNow", "False") = "True" Then
InetClose($hDownload)
ExitLoop
EndIf
WEnd
EndFunc
Func PlayAnim()
$hHBmp_BG = _GDIPlus_MultiColorLoader($iW, $iH, $arrText[$iText])
If TimerDiff($hTimeInt) >= 1000 Then
$iText += 1
$hTimeInt = TimerInit()
EndIf
If $iText >= UBound($arrText) Then $iText = 0
$hB = GUICtrlSendMsg($iPic, $STM_SETIMAGE, $IMAGE_BITMAP, $hHBmp_BG)
If $hB Then _WinAPI_DeleteObject($hB)
_WinAPI_DeleteObject($hHBmp_BG)
EndFunc
Func _SetText($sText = "Loading", $bPoint = True)
If $bPoint = True Then
$arrText[0] = $sText & "."
$arrText[1] = $sText & ".."
$arrText[2] = $sText & "..."
$bAnimate = True
Else
$arrText[0] = $sText
$arrText[1] = $sText
$arrText[2] = $sText
$bAnimate = False
EndIf
EndFunc
Func _GDIPlus_MultiColorLoader($iW, $iH, $sText = "", $sFont = "Segoe UI", $bHBitmap = True)
Local Const $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
Local Const $hGfx = _GDIPlus_ImageGetGraphicsContext($hBitmap)
_GDIPlus_GraphicsSetSmoothingMode($hGfx, 4 +(@OSBuild > 5999))
_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 3)
_GDIPlus_GraphicsSetPixelOffsetMode($hGfx, $GDIP_PIXELOFFSETMODE_HIGHQUALITY)
_GDIPlus_GraphicsClear($hGfx, 0xFF232323)
Local $iRadius =($iW > $iH) ? $iH * 0.6 : $iW * 0.6
Local Const $hPath = _GDIPlus_PathCreate()
_GDIPlus_PathAddEllipse($hPath,($iW -($iRadius + 24)) / 2,($iH -($iRadius + 24)) / 2, $iRadius + 24, $iRadius + 24)
Local $hBrush = _GDIPlus_PathBrushCreateFromPath($hPath)
_GDIPlus_PathBrushSetCenterColor($hBrush, 0xFFFFFFFF)
_GDIPlus_PathBrushSetSurroundColor($hBrush, 0x08101010)
_GDIPlus_PathBrushSetGammaCorrection($hBrush, True)
Local $aBlend[4][2] = [[3]]
$aBlend[1][0] = 0
$aBlend[1][1] = 0
$aBlend[2][0] = 0.33
$aBlend[2][1] = 0.1
$aBlend[3][0] = 1
$aBlend[3][1] = 1
_GDIPlus_PathBrushSetBlend($hBrush, $aBlend)
Local $aRect = _GDIPlus_PathBrushGetRect($hBrush)
_GDIPlus_GraphicsFillRect($hGfx, $aRect[0], $aRect[1], $aRect[2], $aRect[3], $hBrush)
_GDIPlus_PathDispose($hPath)
_GDIPlus_BrushDispose($hBrush)
Local Const $hBrush_Black = _GDIPlus_BrushCreateSolid(0xFF161616)
_GDIPlus_GraphicsFillEllipse($hGfx,($iW -($iRadius + 10)) / 2,($iH -($iRadius + 10)) / 2, $iRadius + 10, $iRadius + 10, $hBrush_Black)
Local Const $hBitmap_Gradient = _GDIPlus_BitmapCreateFromScan0($iRadius, $iRadius)
Local Const $hGfx_Gradient = _GDIPlus_ImageGetGraphicsContext($hBitmap_Gradient)
_GDIPlus_GraphicsSetSmoothingMode($hGfx_Gradient, 4 +(@OSBuild > 5999))
Local Const $hMatrix = _GDIPlus_MatrixCreate()
Local Static $r = 0
_GDIPlus_MatrixTranslate($hMatrix, $iRadius / 2, $iRadius / 2)
_GDIPlus_MatrixRotate($hMatrix, $r)
_GDIPlus_MatrixTranslate($hMatrix, -$iRadius / 2, -$iRadius / 2)
_GDIPlus_GraphicsSetTransform($hGfx_Gradient, $hMatrix)
$r += 10
Local Const $hBrush_Gradient = _GDIPlus_LineBrushCreate($iRadius, $iRadius / 2, $iRadius, $iRadius, 0xFF000000, 0xFF33CAFD, 1)
_GDIPlus_LineBrushSetGammaCorrection($hBrush_Gradient)
_GDIPlus_GraphicsFillEllipse($hGfx_Gradient, 0, 0, $iRadius, $iRadius, $hBrush_Gradient)
_GDIPlus_GraphicsFillEllipse($hGfx_Gradient, 4, 4, $iRadius - 8, $iRadius - 8, $hBrush_Black)
_GDIPlus_GraphicsDrawImageRect($hGfx, $hBitmap_Gradient,($iW - $iRadius) / 2,($iH - $iRadius) / 2, $iRadius, $iRadius)
_GDIPlus_BrushDispose($hBrush_Gradient)
_GDIPlus_BrushDispose($hBrush_Black)
_GDIPlus_GraphicsDispose($hGfx_Gradient)
_GDIPlus_BitmapDispose($hBitmap_Gradient)
_GDIPlus_MatrixDispose($hMatrix)
Local Const $hFormat = _GDIPlus_StringFormatCreate()
Local Const $hFamily = _GDIPlus_FontFamilyCreate($sFont)
Local Const $hFont = _GDIPlus_FontCreate($hFamily, $iRadius / 10)
_GDIPlus_StringFormatSetAlign($hFormat, 1)
_GDIPlus_StringFormatSetLineAlign($hFormat, 1)
Local $tLayout = _GDIPlus_RectFCreate(0, 0, $iW, $iH)
Local Static $iColor = 0x00, $iDir = 13
Local $hBrush_txt = _GDIPlus_BrushCreateSolid(0xFF000000 + 0x010000 * $iColor + 0x0100 * $iColor + $iColor)
_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hFont, $tLayout, $hFormat, $hBrush_txt)
If $bAnimate Then
$iColor += $iDir
If $iColor > 0xFF Then
$iColor = 0xFF
$iDir *= -1
ElseIf $iColor < 0x16 Then
$iDir *= -1
$iColor = 0x16
EndIf
Else
$iColor = 0xFF
EndIf
_GDIPlus_BrushDispose($hBrush_txt)
_GDIPlus_FontDispose($hFont)
_GDIPlus_FontFamilyDispose($hFamily)
_GDIPlus_StringFormatDispose($hFormat)
_GDIPlus_GraphicsDispose($hGfx)
If $bHBitmap Then
Local $hHBITMAP = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hBitmap)
_GDIPlus_BitmapDispose($hBitmap)
Return $hHBITMAP
EndIf
Return $hBitmap
EndFunc
