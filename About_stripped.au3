#NoTrayIcon
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_SHOW = 16
Global Const $GUI_HIDE = 32
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Global Const $GUI_FOCUS = 256
Global Const $GUI_BKCOLOR_TRANSPARENT = -2
Global Const $GMEM_MOVEABLE = 0x0002
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagGDIPRECTF = "struct;float X;float Y;float Width;float Height;endstruct"
Global Const $tagGDIPSTARTUPINPUT = "uint Version;ptr Callback;bool NoThread;bool NoCodecs"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagGUID = "struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];endstruct"
Func _MemGlobalAlloc($iBytes, $iFlags = 0)
Local $aResult = DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", $iFlags, "ulong_ptr", $iBytes)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemGlobalFree($hMem)
Local $aResult = DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", $hMem)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemGlobalLock($hMem)
Local $aResult = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", $hMem)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemGlobalUnlock($hMem)
Local $aResult = DllCall("kernel32.dll", "bool", "GlobalUnlock", "handle", $hMem)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
Return $aResult
EndFunc
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $IMAGE_BITMAP = 0
Func _WinAPI_CloseHandle($hObject)
Local $aResult = DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteObject($hObject)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_GetClientHeight($hWnd)
Local $tRect = _WinAPI_GetClientRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRect, "Bottom") - DllStructGetData($tRect, "Top")
EndFunc
Func _WinAPI_GetClientWidth($hWnd)
Local $tRect = _WinAPI_GetClientRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRect, "Right") - DllStructGetData($tRect, "Left")
EndFunc
Func _WinAPI_GetClientRect($hWnd)
Local $tRect = DllStructCreate($tagRECT)
Local $aRet = DllCall("user32.dll", "bool", "GetClientRect", "hwnd", $hWnd, "struct*", $tRect)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Return $tRect
EndFunc
Func _WinAPI_GetCurrentProcess()
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentProcess")
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDC($hWnd)
Local $aResult = DllCall("user32.dll", "handle", "GetDC", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetParent($hWnd)
Local $aResult = DllCall("user32.dll", "hwnd", "GetParent", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_ReleaseDC($hWnd, $hDC)
Local $aResult = DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_StringFromGUID($pGUID)
Local $aResult = DllCall("ole32.dll", "int", "StringFromGUID2", "struct*", $pGUID, "wstr", "", "int", 40)
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, "")
Return SetExtended($aResult[0], $aResult[2])
EndFunc
Global $__g_pGRC_StreamFromFileCallback = DllCallbackRegister("__GCR_StreamFromFileCallback", "dword", "long_ptr;ptr;long;ptr")
Global $__g_pGRC_StreamFromVarCallback = DllCallbackRegister("__GCR_StreamFromVarCallback", "dword", "long_ptr;ptr;long;ptr")
Global $__g_pGRC_StreamToFileCallback = DllCallbackRegister("__GCR_StreamToFileCallback", "dword", "long_ptr;ptr;long;ptr")
Global $__g_pGRC_StreamToVarCallback = DllCallbackRegister("__GCR_StreamToVarCallback", "dword", "long_ptr;ptr;long;ptr")
Global $__g_pGRC_sStreamVar
Global $__g_hLib_RichCom_OLE32 = DllOpen("OLE32.DLL")
Global $__g_pRichCom_Object_QueryInterface = DllCallbackRegister("__RichCom_Object_QueryInterface", "long", "ptr;dword;dword")
Global $__g_pRichCom_Object_AddRef = DllCallbackRegister("__RichCom_Object_AddRef", "long", "ptr")
Global $__g_pRichCom_Object_Release = DllCallbackRegister("__RichCom_Object_Release", "long", "ptr")
Global $__g_pRichCom_Object_GetNewStorage = DllCallbackRegister("__RichCom_Object_GetNewStorage", "long", "ptr;ptr")
Global $__g_pRichCom_Object_GetInPlaceContext = DllCallbackRegister("__RichCom_Object_GetInPlaceContext", "long", "ptr;dword;dword;dword")
Global $__g_pRichCom_Object_ShowContainerUI = DllCallbackRegister("__RichCom_Object_ShowContainerUI", "long", "ptr;long")
Global $__g_pRichCom_Object_QueryInsertObject = DllCallbackRegister("__RichCom_Object_QueryInsertObject", "long", "ptr;dword;ptr;long")
Global $__g_pRichCom_Object_DeleteObject = DllCallbackRegister("__RichCom_Object_DeleteObject", "long", "ptr;ptr")
Global $__g_pRichCom_Object_QueryAcceptData = DllCallbackRegister("__RichCom_Object_QueryAcceptData", "long", "ptr;ptr;dword;dword;dword;ptr")
Global $__g_pRichCom_Object_ContextSensitiveHelp = DllCallbackRegister("__RichCom_Object_ContextSensitiveHelp", "long", "ptr;long")
Global $__g_pRichCom_Object_GetClipboardData = DllCallbackRegister("__RichCom_Object_GetClipboardData", "long", "ptr;ptr;dword;ptr")
Global $__g_pRichCom_Object_GetDragDropEffect = DllCallbackRegister("__RichCom_Object_GetDragDropEffect", "long", "ptr;dword;dword;dword")
Global $__g_pRichCom_Object_GetContextMenu = DllCallbackRegister("__RichCom_Object_GetContextMenu", "long", "ptr;short;ptr;ptr;ptr")
Global Const $_GCR_S_OK = 0
Global Const $_GCR_E_NOTIMPL = 0x80004001
Func __GCR_StreamFromFileCallback($hFile, $pBuf, $iBuflen, $pQbytes)
Local $tQbytes = DllStructCreate("long", $pQbytes)
DllStructSetData($tQbytes, 1, 0)
Local $tBuf = DllStructCreate("char[" & $iBuflen & "]", $pBuf)
Local $sBuf = FileRead($hFile, $iBuflen - 1)
If @error <> 0 Then Return 1
DllStructSetData($tBuf, 1, $sBuf)
DllStructSetData($tQbytes, 1, StringLen($sBuf))
Return 0
EndFunc
Func __GCR_StreamFromVarCallback($iCookie, $pBuf, $iBuflen, $pQbytes)
#forceref $iCookie
Local $tQbytes = DllStructCreate("long", $pQbytes)
DllStructSetData($tQbytes, 1, 0)
Local $tCtl = DllStructCreate("char[" & $iBuflen & "]", $pBuf)
Local $sCtl = StringLeft($__g_pGRC_sStreamVar, $iBuflen - 1)
If $sCtl = "" Then Return 1
DllStructSetData($tCtl, 1, $sCtl)
Local $iLen = StringLen($sCtl)
DllStructSetData($tQbytes, 1, $iLen)
$__g_pGRC_sStreamVar = StringMid($__g_pGRC_sStreamVar, $iLen + 1)
Return 0
EndFunc
Func __GCR_StreamToFileCallback($hFile, $pBuf, $iBuflen, $pQbytes)
Local $tQbytes = DllStructCreate("long", $pQbytes)
DllStructSetData($tQbytes, 1, 0)
Local $tBuf = DllStructCreate("char[" & $iBuflen & "]", $pBuf)
Local $s = DllStructGetData($tBuf, 1)
FileWrite($hFile, $s)
DllStructSetData($tQbytes, 1, StringLen($s))
Return 0
EndFunc
Func __GCR_StreamToVarCallback($iCookie, $pBuf, $iBuflen, $pQbytes)
#forceref $iCookie
Local $tQbytes = DllStructCreate("long", $pQbytes)
DllStructSetData($tQbytes, 1, 0)
Local $tBuf = DllStructCreate("char[" & $iBuflen & "]", $pBuf)
Local $s = DllStructGetData($tBuf, 1)
$__g_pGRC_sStreamVar &= $s
Return 0
EndFunc
Func __RichCom_Object_QueryInterface($pObject, $iREFIID, $pPvObj)
#forceref $pObject, $iREFIID, $pPvObj
Return $_GCR_S_OK
EndFunc
Func __RichCom_Object_AddRef($pObject)
Local $tData = DllStructCreate("ptr;dword", $pObject)
DllStructSetData($tData, 2, DllStructGetData($tData, 2) + 1)
Return DllStructGetData($tData, 2)
EndFunc
Func __RichCom_Object_Release($pObject)
Local $tData = DllStructCreate("ptr;dword", $pObject)
If DllStructGetData($tData, 2) > 0 Then
DllStructSetData($tData, 2, DllStructGetData($tData, 2) - 1)
Return DllStructGetData($tData, 2)
EndIf
EndFunc
Func __RichCom_Object_GetInPlaceContext($pObject, $pPFrame, $pPDoc, $pFrameInfo)
#forceref $pObject, $pPFrame, $pPDoc, $pFrameInfo
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_ShowContainerUI($pObject, $bShow)
#forceref $pObject, $bShow
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_QueryInsertObject($pObject, $pClsid, $tStg, $vCp)
#forceref $pObject, $pClsid, $tStg, $vCp
Return $_GCR_S_OK
EndFunc
Func __RichCom_Object_DeleteObject($pObject, $pOleobj)
#forceref $pObject, $pOleobj
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_QueryAcceptData($pObject, $pDataobj, $pCfFormat, $vReco, $bReally, $hMetaPict)
#forceref $pObject, $pDataobj, $pCfFormat, $vReco, $bReally, $hMetaPict
Return $_GCR_S_OK
EndFunc
Func __RichCom_Object_ContextSensitiveHelp($pObject, $bEnterMode)
#forceref $pObject, $bEnterMode
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_GetClipboardData($pObject, $pChrg, $vReco, $pPdataobj)
#forceref $pObject, $pChrg, $vReco, $pPdataobj
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_GetDragDropEffect($pObject, $bDrag, $iGrfKeyState, $piEffect)
#forceref $pObject, $bDrag, $iGrfKeyState, $piEffect
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_GetContextMenu($pObject, $iSeltype, $pOleobj, $pChrg, $pHmenu)
#forceref $pObject, $iSeltype, $pOleobj, $pChrg, $pHmenu
Return $_GCR_E_NOTIMPL
EndFunc
Func __RichCom_Object_GetNewStorage($pObject, $pPstg)
#forceref $pObject
Local $aSc = DllCall($__g_hLib_RichCom_OLE32, "dword", "CreateILockBytesOnHGlobal", "hwnd", 0, "int", 1, "ptr*", 0)
Local $pLockBytes = $aSc[3]
$aSc = $aSc[0]
If $aSc Then Return $aSc
$aSc = DllCall($__g_hLib_RichCom_OLE32, "dword", "StgCreateDocfileOnILockBytes", "ptr", $pLockBytes, "dword", BitOR(0x10, 2, 0x1000), "dword", 0, "ptr*", 0)
Local $tStg = DllStructCreate("ptr", $pPstg)
DllStructSetData($tStg, 1, $aSc[4])
$aSc = $aSc[0]
If $aSc Then
Local $tObj = DllStructCreate("ptr", $pLockBytes)
Local $tUnknownFuncTable = DllStructCreate("ptr[3]", DllStructGetData($tObj, 1))
Local $pReleaseFunc = DllStructGetData($tUnknownFuncTable, 3)
DllCallAddress("long", $pReleaseFunc, "ptr", $pLockBytes)
EndIf
Return $aSc
EndFunc
Global Const $WS_EX_TRANSPARENT = 0x00000020
Global Const $WM_DESTROY = 0x0002
Global Const $GDIP_PXF32ARGB = 0x0026200A
Global Const $GDIP_IMAGEFORMAT_UNDEFINED = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_MEMORYBMP = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_EMF = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_WMF = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_EXIF = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_IMAGEFORMAT_ICON = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"
Global Const $GDIP_SMOOTHINGMODE_DEFAULT = 0
Global Const $GDIP_SMOOTHINGMODE_HIGHQUALITY = 2
Global Const $GDIP_SMOOTHINGMODE_ANTIALIAS8X8 = 5
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Global $__g_hGDIPDll = 0
Global $__g_iGDIPRef = 0
Global $__g_iGDIPToken = 0
Global $__g_bGDIP_V1_0 = True
Func _GDIPlus_BitmapCreateFromFile($sFileName)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateBitmapFromFile", "wstr", $sFileName, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight, $iPixelFormat = $GDIP_PXF32ARGB, $iStride = 0, $pScan0 = 0)
Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $iWidth, "int", $iHeight, "int", $iStride, "int", $iPixelFormat, "ptr", $pScan0, "handle*", 0)
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
Func _GDIPlus_GraphicsDrawImage($hGraphics, $hImage, $nX, $nY)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDrawImage", "handle", $hGraphics, "handle", $hImage, "float", $nX, "float", $nY)
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
Func _GDIPlus_GraphicsMeasureString($hGraphics, $sString, $hFont, $tLayout, $hFormat)
Local $tRectF = DllStructCreate($tagGDIPRECTF)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipMeasureString", "handle", $hGraphics, "wstr", $sString, "int", -1, "handle", $hFont, "struct*", $tLayout, "handle", $hFormat, "struct*", $tRectF, "int*", 0, "int*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Local $aInfo[3]
$aInfo[0] = $tRectF
$aInfo[1] = $aResult[8]
$aInfo[2] = $aResult[9]
Return $aInfo
EndFunc
Func _GDIPlus_GraphicsSetSmoothingMode($hGraphics, $iSmooth)
If $iSmooth < $GDIP_SMOOTHINGMODE_DEFAULT Or $iSmooth > $GDIP_SMOOTHINGMODE_ANTIALIAS8X8 Then $iSmooth = $GDIP_SMOOTHINGMODE_DEFAULT
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetSmoothingMode", "handle", $hGraphics, "int", $iSmooth)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_ImageDispose($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDisposeImage", "handle", $hImage)
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
Func _GDIPlus_ImageGetHeight($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageHeight", "handle", $hImage, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
EndFunc
Func _GDIPlus_ImageGetRawFormat($hImage)
Local $aGuid[2]
If($hImage = -1) Or(Not $hImage) Then Return SetError(11, 0, $aGuid)
Local $aImageType[11][2] = [["UNDEFINED", $GDIP_IMAGEFORMAT_UNDEFINED], ["MEMORYBMP", $GDIP_IMAGEFORMAT_MEMORYBMP], ["BMP", $GDIP_IMAGEFORMAT_BMP], ["EMF", $GDIP_IMAGEFORMAT_EMF], ["WMF", $GDIP_IMAGEFORMAT_WMF], ["JPEG", $GDIP_IMAGEFORMAT_JPEG], ["PNG", $GDIP_IMAGEFORMAT_PNG], ["GIF", $GDIP_IMAGEFORMAT_GIF], ["TIFF", $GDIP_IMAGEFORMAT_TIFF], ["EXIF", $GDIP_IMAGEFORMAT_EXIF], ["ICON", $GDIP_IMAGEFORMAT_ICON]]
Local $tStruct = DllStructCreate("byte[16]")
Local $aResult1 = DllCall($__g_hGDIPDll, "int", "GdipGetImageRawFormat", "handle", $hImage, "struct*", $tStruct)
If @error Then Return SetError(@error, @extended, $aGuid)
If $aResult1[0] Then Return SetError(10, $aResult1[0], $aGuid)
Local $sResult2 = _WinAPI_StringFromGUID($aResult1[2])
If @error Then Return SetError(@error + 20, @extended, $aGuid)
If $sResult2 = "" Then Return SetError(12, 0, $aGuid)
For $i = 0 To 10
If $aImageType[$i][1] == $sResult2 Then
$aGuid[0] = $aImageType[$i][1]
$aGuid[1] = $aImageType[$i][0]
Return $aGuid
EndIf
Next
Return SetError(13, 0, $aGuid)
EndFunc
Func _GDIPlus_ImageGetWidth($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageWidth", "handle", $hImage, "uint*", -1)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
EndFunc
Func _GDIPlus_ImageLoadFromFile($sFileName)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipLoadImageFromFile", "wstr", $sFileName, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[2]
EndFunc
Func _GDIPlus_RectFCreate($nX = 0, $nY = 0, $nWidth = 0, $nHeight = 0)
Local $tRectF = DllStructCreate($tagGDIPRECTF)
DllStructSetData($tRectF, "X", $nX)
DllStructSetData($tRectF, "Y", $nY)
DllStructSetData($tRectF, "Width", $nWidth)
DllStructSetData($tRectF, "Height", $nHeight)
Return $tRectF
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
If $sGDIPDLL = Default Then
If @OSBuild > 4999 And @OSBuild < 7600 Then
$sGDIPDLL = @WindowsDir & "\winsxs\x86_microsoft.windows.gdiplus_6595b64144ccf1df_1.1.6000.16386_none_8df21b8362744ace\gdiplus.dll"
Else
$sGDIPDLL = "gdiplus.dll"
EndIf
EndIf
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
Return True
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
Global Const $SS_LEFT = 0x0
Global Const $SS_RIGHT = 0x2
Global Const $SS_BITMAP = 0xE
Global Const $SS_NOTIFY = 0x0100
Global Const $ILC_MASK = 0x00000001
Global Const $ILC_COLOR = 0x00000000
Global Const $ILC_COLORDDB = 0x000000FE
Global Const $ILC_COLOR4 = 0x00000004
Global Const $ILC_COLOR8 = 0x00000008
Global Const $ILC_COLOR16 = 0x00000010
Global Const $ILC_COLOR24 = 0x00000018
Global Const $ILC_COLOR32 = 0x00000020
Global Const $ILC_MIRROR = 0x00002000
Global Const $ILC_PERITEMMIRROR = 0x00008000
Func _GUIImageList_Create($iCX = 16, $iCY = 16, $iColor = 4, $iOptions = 0, $iInitial = 4, $iGrow = 4)
Local Const $aColor[7] = [$ILC_COLOR, $ILC_COLOR4, $ILC_COLOR8, $ILC_COLOR16, $ILC_COLOR24, $ILC_COLOR32, $ILC_COLORDDB]
Local $iFlags = 0
If BitAND($iOptions, 1) <> 0 Then $iFlags = BitOR($iFlags, $ILC_MASK)
If BitAND($iOptions, 2) <> 0 Then $iFlags = BitOR($iFlags, $ILC_MIRROR)
If BitAND($iOptions, 4) <> 0 Then $iFlags = BitOR($iFlags, $ILC_PERITEMMIRROR)
$iFlags = BitOR($iFlags, $aColor[$iColor])
Local $aResult = DllCall("comctl32.dll", "handle", "ImageList_Create", "int", $iCX, "int", $iCY, "uint", $iFlags, "int", $iInitial, "int", $iGrow)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _GUIImageList_Destroy($hWnd)
Local $aResult = DllCall("comctl32.dll", "bool", "ImageList_Destroy", "handle", $hWnd)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0] <> 0
EndFunc
Func _ArrayMax(Const ByRef $avArray, $iCompNumeric = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
Local $iResult = _ArrayMaxIndex($avArray, $iCompNumeric, $iStart, $iEnd, $iSubItem)
If @error Then Return SetError(@error, 0, "")
If UBound($avArray, $UBOUND_DIMENSIONS) = 1 Then
Return $avArray[$iResult]
Else
Return $avArray[$iResult][$iSubItem]
EndIf
EndFunc
Func _ArrayMaxIndex(Const ByRef $avArray, $iCompNumeric = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iCompNumeric <> 1 Then $iCompNumeric = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If Not IsArray($avArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($avArray, $UBOUND_ROWS) - 1
If $iEnd = 0 Then $iEnd = $iDim_1
If $iStart < 0 Or $iEnd < 0 Then Return SetError(3, 0, -1)
If $iStart > $iDim_1 Or $iEnd > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart > $iEnd Then Return SetError(4, 0, -1)
If $iDim_1 < 1 Then Return SetError(5, 0, -1)
Local $iMaxIndex = $iStart
Switch UBound($avArray, $UBOUND_DIMENSIONS)
Case 1
If $iCompNumeric Then
For $i = $iStart To $iEnd
If Number($avArray[$iMaxIndex]) < Number($avArray[$i]) Then $iMaxIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $avArray[$iMaxIndex] < $avArray[$i] Then $iMaxIndex = $i
Next
EndIf
Case 2
If $iSubItem < 0 Or $iSubItem > UBound($avArray, $UBOUND_COLUMNS) - 1 Then Return SetError(6, 0, -1)
If $iCompNumeric Then
For $i = $iStart To $iEnd
If Number($avArray[$iMaxIndex][$iSubItem]) < Number($avArray[$i][$iSubItem]) Then $iMaxIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $avArray[$iMaxIndex][$iSubItem] < $avArray[$i][$iSubItem] Then $iMaxIndex = $i
Next
EndIf
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return $iMaxIndex
EndFunc
Func _ArrayMin(Const ByRef $avArray, $iCompNumeric = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
Local $iResult = _ArrayMinIndex($avArray, $iCompNumeric, $iStart, $iEnd, $iSubItem)
If @error Then Return SetError(@error, 0, "")
If UBound($avArray, $UBOUND_DIMENSIONS) = 1 Then
Return $avArray[$iResult]
Else
Return $avArray[$iResult][$iSubItem]
EndIf
EndFunc
Func _ArrayMinIndex(Const ByRef $avArray, $iCompNumeric = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If Not IsArray($avArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($avArray, $UBOUND_ROWS) - 1
If $iEnd = 0 Then $iEnd = $iDim_1
If $iStart < 0 Or $iEnd < 0 Then Return SetError(3, 0, -1)
If $iStart > $iDim_1 Or $iEnd > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart > $iEnd Then Return SetError(4, 0, -1)
If $iDim_1 < 1 Then Return SetError(5, 0, -1)
Local $iMinIndex = $iStart
Switch UBound($avArray, $UBOUND_DIMENSIONS)
Case 1
If $iCompNumeric Then
For $i = $iStart To $iEnd
If Number($avArray[$iMinIndex]) > Number($avArray[$i]) Then $iMinIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $avArray[$iMinIndex] > $avArray[$i] Then $iMinIndex = $i
Next
EndIf
Case 2
If $iSubItem < 0 Or $iSubItem > UBound($avArray, $UBOUND_COLUMNS) - 1 Then Return SetError(6, 0, -1)
If $iCompNumeric Then
For $i = $iStart To $iEnd
If Number($avArray[$iMinIndex][$iSubItem]) > Number($avArray[$i][$iSubItem]) Then $iMinIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $avArray[$iMinIndex][$iSubItem] > $avArray[$i][$iSubItem] Then $iMinIndex = $i
Next
EndIf
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return $iMinIndex
EndFunc
Global Const $GIS_ASPECTRATIOFIX = 0x2000
Global Const $GIS_HALFTRANSPARENCY = 0x4014
Global Const $GIS_FULLTRANSPARENCY = 0x40015
Global Const $GIS_SS_DEFAULT_PIC = BitOR($GIS_HALFTRANSPARENCY, $SS_NOTIFY)
Global $GIS_EX_DEFAULTRENDER = 0x21
Global $GIS_EX_CTRLSNDRENDER = 0x22
Global $avGRP_CTRLIDS[1][20]
Global $vGRP_TEMP, $iGRP_MSG = 0
Global $pGRP_tGUID = DllStructCreate($tagGUID)
Global Const $iGRP_ptTypeByte = 1
Global Const $iGRP_ptTypeShort = 3
Global Const $iGRP_ptTypeLong = 4
Global Const $iGRP_ptLoopCount = 0x5101
Global Const $iGRP_ptFrameDelay = 0x5100
Global Const $tGRP_PropTagItem = "long id; long length; int Type; ptr value"
Global Const $hGRP_GDI32 = DllOpen("gdi32.dll")
Global Const $hGRP_USER32 = DllOpen("user32.dll")
Global Const $hGRP_COMCTL32 = DllOpen("comctl32.dll")
OnAutoItExitRegister("__GRP_ShutDown")
Func _GUICtrlPic_Create($sFileName, $iLeft = 0, $iTop = 0, $iWidth = -1, $iHeight = -1, $iStyle = -1, $iExStyle = -1, $INTERNALID = 0)
Local $iCtrlID, $hCtrlID, $aCtrlStyle, $aCtrlExStyle, $iCtrlSize = 0x40
Local $hOImage, $pDimensionIDs, $iTransparency = 1, $iFrameCount, $aiFrameDelays, $iLoopCount, $hWndForm = 0
Local $iIndex, $iOWidth, $iOHeight, $iReSize = 1, $asImgType, $tCtrlInfo, $iResizeMode
Local $t_Style, $t_ExStyle, $iDefaultRender = 1, $iTransMode = 0, $hHGMem
If $INTERNALID Then
If Not GUICtrlGetHandle($INTERNALID) Then Return SetError(0, 0, 0)
$iCtrlID = $INTERNALID
EndIf
If $__g_hGDIPDll = 0 Then _GDIPlus_Startup()
If Not __GRP_GetFileNameType($sFileName, $hOImage, $hHGMem) Then Return SetError(0, 0, 0)
$asImgType = _GDIPlus_ImageGetRawFormat($hOImage)
$asImgType = $asImgType[1]
If $asImgType = "GIF" Then
If Not __GRP_GetGifInfo($pDimensionIDs, $hOImage, $iFrameCount, $iLoopCount, $aiFrameDelays, $iTransparency) Then
__GRP_FreeMem($hOImage, $hHGMem)
Return SetError(0, 0, 0)
EndIf
EndIf
If $iStyle = -1 Then $iStyle = $GIS_SS_DEFAULT_PIC
If $iExStyle = -1 Then $iExStyle = $WS_EX_TRANSPARENT
If BitAND($iStyle, $GIS_FULLTRANSPARENCY) = $GIS_FULLTRANSPARENCY Then
If $iTransparency Then $iTransMode = 1
$t_Style = BitOR($iStyle, $GIS_FULLTRANSPARENCY)
$iStyle = BitXOR($iStyle, $GIS_FULLTRANSPARENCY)
EndIf
If BitAND($iExStyle, $GIS_EX_CTRLSNDRENDER) = $GIS_EX_CTRLSNDRENDER Then
$iDefaultRender = 0
If $iTransparency Then $iTransMode = 1
$t_ExStyle = BitOR($iExStyle, $GIS_EX_CTRLSNDRENDER)
$iExStyle = BitXOR($iExStyle, $GIS_EX_CTRLSNDRENDER)
EndIf
$iOWidth = _GDIPlus_ImageGetWidth($hOImage)
$iOHeight = _GDIPlus_ImageGetHeight($hOImage)
Select
Case($iWidth = 0 And $iHeight = 0) Or($iWidth = -1 And $iHeight = -1) Or($iWidth = $iOWidth And $iHeight = $iOHeight)
$iReSize = 0
$iCtrlSize = 0x800
$iWidth = $iOWidth
$iHeight = $iOHeight
If BitAND($iStyle, $GIS_ASPECTRATIOFIX) = $GIS_ASPECTRATIOFIX Then
$t_Style = BitOR($iStyle, $GIS_ASPECTRATIOFIX)
$iStyle = BitXOR($iStyle, $GIS_ASPECTRATIOFIX)
EndIf
Case BitAND($iStyle, $GIS_ASPECTRATIOFIX) = $GIS_ASPECTRATIOFIX
Local $iAspectW = Int($iOWidth * $iHeight / $iOHeight)
Local $iAspectH = Int($iOHeight * $iWidth / $iOWidth)
Switch($iAspectW > $iWidth)
Case True
$iHeight = $iAspectH
Case False
$iWidth = $iAspectW
EndSwitch
$t_Style = BitOR($iStyle, $GIS_ASPECTRATIOFIX)
$iStyle = BitXOR($iStyle, $GIS_ASPECTRATIOFIX)
EndSelect
If BitAND($iStyle, $GIS_HALFTRANSPARENCY) = $GIS_HALFTRANSPARENCY Then $iStyle = BitXOR($iStyle, $GIS_HALFTRANSPARENCY)
If BitAND($iExStyle, $GIS_EX_DEFAULTRENDER) = $GIS_EX_DEFAULTRENDER Then $iExStyle = BitXOR($iExStyle, $GIS_EX_DEFAULTRENDER)
Switch $INTERNALID
Case 0
$iCtrlID = GUICtrlCreateLabel("", $iLeft, $iTop, $iWidth, $iHeight, BitOR($iStyle, $iCtrlSize, $SS_BITMAP), $iExStyle)
$iResizeMode = Opt("GUIResizeMode")
Select
Case Opt("GUIEventOptions") = 0 And $iResizeMode = 0
GUICtrlSetResizing(-1, 768)
Case $iResizeMode > 0
GUICtrlSetResizing(-1, $iResizeMode)
EndSelect
Case Else
GUICtrlSetPos($iCtrlID, Default, Default, $iWidth, $iHeight)
GUICtrlSetStyle($iCtrlID, BitOR($iStyle, $iCtrlSize, $SS_BITMAP), $iExStyle)
EndSwitch
GUICtrlSetBkColor($iCtrlID, $GUI_BKCOLOR_TRANSPARENT)
$hCtrlID = GUICtrlGetHandle($iCtrlID)
$aCtrlStyle = DllCall($hGRP_USER32, "long", "GetWindowLong", "hwnd", $hCtrlID, "int", -16)
$aCtrlExStyle = DllCall($hGRP_USER32, "long", "GetWindowLong", "hwnd", $hCtrlID, "int", -20)
$hWndForm = _WinAPI_GetParent($hCtrlID)
If IsBinary($sFileName) Then $sFileName = "Binary"
$tCtrlInfo = DllStructCreate("hwnd CtrlHandle;int Style;int ExStyle;wchar FileName[" & StringLen($sFileName) & "];" & "long Left;long Top;long Width;long Height;long OWidht;long OHeight;int tStyle;int tExStyle;wchar Function[15];long HGMEM")
DllStructSetData($tCtrlInfo, "CtrlHandle", $hCtrlID)
DllStructSetData($tCtrlInfo, "Style", $aCtrlStyle[0])
DllStructSetData($tCtrlInfo, "ExStyle", $aCtrlExStyle[0])
DllStructSetData($tCtrlInfo, "FileName", $sFileName)
DllStructSetData($tCtrlInfo, "Left", $iLeft)
DllStructSetData($tCtrlInfo, "Top", $iTop)
DllStructSetData($tCtrlInfo, "Width", $iWidth)
DllStructSetData($tCtrlInfo, "Height", $iHeight)
DllStructSetData($tCtrlInfo, "OWidht", $iOWidth)
DllStructSetData($tCtrlInfo, "OHeight", $iOHeight)
DllStructSetData($tCtrlInfo, "tStyle", $t_Style)
DllStructSetData($tCtrlInfo, "tExStyle", $t_ExStyle)
DllStructSetData($tCtrlInfo, "Function", "__GRP_BuildList")
DllStructSetData($tCtrlInfo, "HGMEM", $hHGMem)
$iIndex = $avGRP_CTRLIDS[0][0] + 1
ReDim $avGRP_CTRLIDS[$iIndex + 1][20]
$avGRP_CTRLIDS[0][0] = $iIndex
$avGRP_CTRLIDS[$iIndex][0] = $iCtrlID
$avGRP_CTRLIDS[$iIndex][1] = $tCtrlInfo
$avGRP_CTRLIDS[$iIndex][2] = $hOImage
$avGRP_CTRLIDS[$iIndex][3] = $pDimensionIDs
$avGRP_CTRLIDS[$iIndex][4] = $iFrameCount
$avGRP_CTRLIDS[$iIndex][5] = $aiFrameDelays
$avGRP_CTRLIDS[$iIndex][6] = $iLoopCount * $iFrameCount
$avGRP_CTRLIDS[$iIndex][7] = 0
$avGRP_CTRLIDS[$iIndex][8] = $iDefaultRender
$avGRP_CTRLIDS[$iIndex][9] = $iReSize
$avGRP_CTRLIDS[$iIndex][10] = $iTransMode
$avGRP_CTRLIDS[$iIndex][11] = 0
$avGRP_CTRLIDS[$iIndex][12] = $hWndForm
$avGRP_CTRLIDS[$iIndex][13] = 0
$avGRP_CTRLIDS[$iIndex][14] = 0
$avGRP_CTRLIDS[$iIndex][15] = 0
$avGRP_CTRLIDS[$iIndex][16] = 0
$avGRP_CTRLIDS[$iIndex][17] = 0
$avGRP_CTRLIDS[$iIndex][18] = $iTransparency
$avGRP_CTRLIDS[$iIndex][19] = 0
Switch $iFrameCount
Case 0
If $iReSize Then
Local $hOBitmap = __GDIPCreateBitmapFromScan0($iWidth, $iHeight, 0, $GDIP_PXF32ARGB, 0)
Local $hGraphic = _GDIPlus_ImageGetGraphicsContext($hOBitmap)
DllCall($__g_hGDIPDll, "uint", "GdipSetInterpolationMode", "hwnd", $hGraphic, "int", 7)
_GDIPlus_GraphicsDrawImageRect($hGraphic, $hOImage, 0, 0, $iWidth, $iHeight)
_GDIPlus_ImageDispose($hOImage)
_GDIPlus_GraphicsDispose($hGraphic)
$hOImage = $hOBitmap
$avGRP_CTRLIDS[$iIndex][2] = $hOImage
EndIf
__GRP_hImgToCtrl($iCtrlID, $hOImage, $pDimensionIDs, 0)
Case Else
Local $iLowTime = _ArrayMin($aiFrameDelays, 1, 0), $iHiTime = _ArrayMax($aiFrameDelays, 1, 0)
If $iLowTime <> $iHiTime Then $avGRP_CTRLIDS[$iIndex][17] = 1
Switch $iDefaultRender
Case 1
$avGRP_CTRLIDS[$iIndex][13] = _WinAPI_GetDC($hCtrlID)
$avGRP_CTRLIDS[$iIndex][14] = _GUIImageList_Create($iWidth, $iHeight, 5, 0, $iFrameCount, 0)
Case 0
Local $HBITMAP[$iFrameCount]
$avGRP_CTRLIDS[$iIndex][14] = $HBITMAP
EndSwitch
__GRP_CreateTimer($hWndForm, $iLowTime, "__GRP_BuildList", $iIndex)
EndSwitch
If Not $iGRP_MSG Then
GUIRegisterMsg($WM_DESTROY, "__GRP_WM_DESTROY")
$iGRP_MSG = 1
Local $hHandle = _WinAPI_GetCurrentProcess()
DllCall("kernel32.dll", "int", "SetProcessWorkingSetSize", "hwnd", $hHandle, "int", -1, "int", -1)
_WinAPI_CloseHandle($hHandle)
EndIf
$tCtrlInfo = 0
Return SetError(0, 0, $iCtrlID)
EndFunc
Func _GUICtrlPic_SetImage($iCtrlID, $sFileName, $lFixSize = False)
Local $hOImage, $pDimensionIDs, $iTransparency = 1, $iTransMode = 0
Local $iFrameCount, $aiFrameDelays, $iLoopCount, $asImgType
Local $iWidth = -1, $iHeight = -1, $iOWidth, $iOHeight, $iReSize = 1, $tCtrlInfo
Local $iIndex, $iStyle, $iExStyle, $hHGMem
If $iCtrlID = -1 Then $iCtrlID = _WinAPI_GetDlgCtrlID(GUICtrlGetHandle(-1))
$iIndex = __GRP_GetCtrlIndex($iCtrlID)
If Not $iIndex Then
Local $aCtrlPos, $hCtrlID
$hCtrlID = GUICtrlGetHandle($iCtrlID)
If Not $hCtrlID Then Return SetError(0, 0, 0)
$aCtrlPos = WinGetPos($hCtrlID)
If @error Then Return SetError(0, 0, 0)
If $lFixSize Then
$aCtrlPos[2] = 0
$aCtrlPos[3] = 0
EndIf
$iStyle = DllCall($hGRP_USER32, "long", "GetWindowLong", "hwnd", $hCtrlID, "int", -16)
$iExStyle = DllCall($hGRP_USER32, "long", "GetWindowLong", "hwnd", $hCtrlID, "int", -20)
Return SetError(0, _GUICtrlPic_Create($sFileName, $aCtrlPos[0], $aCtrlPos[1], $aCtrlPos[2], $aCtrlPos[3], $iStyle[0], $iExStyle[0], $iCtrlID), 1)
EndIf
If Not __GRP_GetFileNameType($sFileName, $hOImage, $hHGMem) Then Return SetError(0, 0, 0)
__GRP_ReleaseGIF($iIndex)
$asImgType = _GDIPlus_ImageGetRawFormat($hOImage)
$asImgType = $asImgType[1]
If $asImgType = "GIF" Then
If Not __GRP_GetGifInfo($pDimensionIDs, $hOImage, $iFrameCount, $iLoopCount, $aiFrameDelays, $iTransparency) Then
_GDIPlus_ImageDispose($hOImage)
Return SetError(0, 0, 0)
EndIf
EndIf
$tCtrlInfo = $avGRP_CTRLIDS[$iIndex][1]
$iStyle = DllStructGetData($tCtrlInfo, "tStyle")
$iExStyle = DllStructGetData($tCtrlInfo, "tExStyle")
If BitAND($iStyle, $GIS_FULLTRANSPARENCY) = $GIS_FULLTRANSPARENCY Or BitAND($iExStyle, $GIS_EX_CTRLSNDRENDER) = $GIS_EX_CTRLSNDRENDER Then
If $iTransparency Then $iTransMode = 1
EndIf
$iOWidth = _GDIPlus_ImageGetWidth($hOImage)
$iOHeight = _GDIPlus_ImageGetHeight($hOImage)
$iWidth = DllStructGetData($tCtrlInfo, "Width")
$iHeight = DllStructGetData($tCtrlInfo, "Height")
Select
Case $lFixSize
$iReSize = 0
$iWidth = $iOWidth
$iHeight = $iOHeight
GUICtrlSetPos($iCtrlID, Default, Default, $iWidth, $iHeight)
Case BitAND($iStyle, $GIS_ASPECTRATIOFIX) = $GIS_ASPECTRATIOFIX
Local $iAspectW = Int($iOWidth * $iHeight / $iOHeight)
Local $iAspectH = Int($iOHeight * $iWidth / $iOWidth)
Switch($iAspectW > $iWidth)
Case True
$iHeight = $iAspectH
Case False
$iWidth = $iAspectW
EndSwitch
EndSelect
If Not IsBinary($sFileName) Then DllStructSetData($tCtrlInfo, "FileName", $sFileName)
DllStructSetData($tCtrlInfo, "Width", $iWidth)
DllStructSetData($tCtrlInfo, "Height", $iHeight)
DllStructSetData($tCtrlInfo, "OWidht", $iOWidth)
DllStructSetData($tCtrlInfo, "OHeight", $iOHeight)
DllStructSetData($tCtrlInfo, "Function", "__GRP_BuildList")
DllStructSetData($tCtrlInfo, "HGMEM", $hHGMem)
$avGRP_CTRLIDS[$iIndex][2] = $hOImage
$avGRP_CTRLIDS[$iIndex][3] = $pDimensionIDs
$avGRP_CTRLIDS[$iIndex][4] = $iFrameCount
$avGRP_CTRLIDS[$iIndex][5] = $aiFrameDelays
$avGRP_CTRLIDS[$iIndex][6] = $iLoopCount * $iFrameCount
$avGRP_CTRLIDS[$iIndex][7] = 0
$avGRP_CTRLIDS[$iIndex][9] = $iReSize
$avGRP_CTRLIDS[$iIndex][10] = $iTransMode
$avGRP_CTRLIDS[$iIndex][11] = 0
$avGRP_CTRLIDS[$iIndex][13] = 0
$avGRP_CTRLIDS[$iIndex][14] = 0
$avGRP_CTRLIDS[$iIndex][17] = 0
$avGRP_CTRLIDS[$iIndex][18] = $iTransparency
$avGRP_CTRLIDS[$iIndex][19] = 0
Local $iState = GUICtrlGetState($iCtrlID)
GUICtrlSetState($iCtrlID, $GUI_HIDE)
Switch $iFrameCount
Case 0
If $iReSize Then
Local $hOBitmap = __GDIPCreateBitmapFromScan0($iWidth, $iHeight, 0, $GDIP_PXF32ARGB, 0)
Local $hGraphic = _GDIPlus_ImageGetGraphicsContext($hOBitmap)
DllCall($__g_hGDIPDll, "uint", "GdipSetInterpolationMode", "hwnd", $hGraphic, "int", 7)
_GDIPlus_GraphicsDrawImageRect($hGraphic, $hOImage, 0, 0, $iWidth, $iHeight)
_GDIPlus_ImageDispose($hOImage)
_GDIPlus_GraphicsDispose($hGraphic)
$hOImage = $hOBitmap
$avGRP_CTRLIDS[$iIndex][2] = $hOImage
EndIf
__GRP_hImgToCtrl($iCtrlID, $hOImage, $pDimensionIDs, 0)
Case Else
Local $iLowTime = _ArrayMin($aiFrameDelays, 1, 0), $iHiTime = _ArrayMax($aiFrameDelays, 1, 0)
If $iLowTime <> $iHiTime Then $avGRP_CTRLIDS[$iIndex][17] = 1
Switch $avGRP_CTRLIDS[$iIndex][8]
Case 1
$avGRP_CTRLIDS[$iIndex][13] = _WinAPI_GetDC(DllStructGetData($tCtrlInfo, "CtrlHandle"))
$avGRP_CTRLIDS[$iIndex][14] = _GUIImageList_Create(DllStructGetData($tCtrlInfo, "Width"), DllStructGetData($tCtrlInfo, "Height"), 5, 0, $iFrameCount, 0)
Case 0
Local $HBITMAP[$iFrameCount]
$avGRP_CTRLIDS[$iIndex][14] = $HBITMAP
EndSwitch
__GRP_CreateTimer($avGRP_CTRLIDS[$iIndex][12], $iLowTime, "__GRP_BuildList", $iIndex)
EndSwitch
GUICtrlSetState($iCtrlID, $iState)
$tCtrlInfo = 0
Return SetError(0, 0, 1)
EndFunc
Func __GRP_GetGifInfo(ByRef $pDimensionIDs, ByRef $hOImage, ByRef $iFrameCount, ByRef $iLoopCount, ByRef $aiFrameDelays, ByRef $iTransparency)
Local $aiDimensionsCount, $aiFrameCount, $aGetPixel
$pDimensionIDs = DllStructGetPtr($pGRP_tGUID)
$aiDimensionsCount = DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameDimensionsCount", "ptr", $hOImage, "int*", 0)
If @error Or Not $aiDimensionsCount[2] Then Return 0
DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameDimensionsList", "ptr", $hOImage, "ptr", $pDimensionIDs, "int", $aiDimensionsCount[2])
If @error Then Return 0
$aiFrameCount = DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameCount", "int", $hOImage, "ptr", $pDimensionIDs, "int*", 0)
If @error Or Not $aiFrameCount[3] Then Return 0
$iFrameCount = $aiFrameCount[3]
If $iFrameCount Then
$aiFrameDelays = __GRP_GetGifFrameDelays($hOImage, $iFrameCount)
$iLoopCount = __GRP_GetGifLoopCount($hOImage)
Else
$iFrameCount = 0
EndIf
$aGetPixel = DllCall($__g_hGDIPDll, "dword", "GdipBitmapGetPixel", "ptr", $hOImage, "int", 0, "int", 0, "dword*", 0)
If Not @error Or $aGetPixel[0] = 0 Then
If BitShift($aGetPixel[4], 24) Then $iTransparency = 0
EndIf
Return 1
EndFunc
Func __GRP_hImgToCtrl($iCtrlID, $hOImage, $pDimensionIDs, $iFrame)
Local $hHBitmap
If $pDimensionIDs And $iFrame Then
DllCall($__g_hGDIPDll, "int", "GdipImageSelectActiveFrame", "ptr", $hOImage, "ptr", $pDimensionIDs, "int", $iFrame)
EndIf
$hHBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hOImage)
__GRP_DeleteObj(GUICtrlSendMsg($iCtrlID, 0x0172, 0, $hHBitmap))
__GRP_DeleteObj($hHBitmap)
EndFunc
Func __GRP_DeleteObj($hObj)
DllCall($hGRP_GDI32, "bool", "DeleteObject", "handle", $hObj)
EndFunc
Func __GRP_CreateTimer($hWnd, $iElapse, $sTimerFunc, $iIndex)
Local $hCallBack = 0, $pTimerFunc = 0, $aResult[1] = [0]
$hCallBack = DllCallbackRegister($sTimerFunc, "none", "hwnd;int;uint_ptr;dword")
$pTimerFunc = DllCallbackGetPtr($hCallBack)
$aResult = DllCall($hGRP_USER32, "uint_ptr", "SetTimer", "hwnd", $hWnd, "uint_ptr", $iIndex, "uint", $iElapse, "ptr", $pTimerFunc)
If @error Or $aResult[0] = 0 Then
DllCallbackFree($hCallBack)
Return SetError(@error, @extended, 0)
EndIf
$avGRP_CTRLIDS[$iIndex][15] = $hCallBack
$avGRP_CTRLIDS[$iIndex][16] = $pTimerFunc
Return $aResult[0]
EndFunc
Func __GRP_KillTimer($hWnd, $iIndex)
Local $aResult[1] = [0], $hCallBack = 0
$aResult = DllCall($hGRP_USER32, "bool", "KillTimer", "hwnd", $hWnd, "uint_ptr", $iIndex)
$hCallBack = $avGRP_CTRLIDS[$iIndex][15]
If $hCallBack <> 0 Then DllCallbackFree($hCallBack)
$avGRP_CTRLIDS[$iIndex][15] = 0
$avGRP_CTRLIDS[$iIndex][16] = 0
Return $aResult[0] <> 0
EndFunc
Func __GRP_ReleaseGIF($iIndex)
Local $hHGMem
If $avGRP_CTRLIDS[$iIndex][4] Then
__GRP_KillTimer($avGRP_CTRLIDS[$iIndex][12], $iIndex)
Switch $avGRP_CTRLIDS[$iIndex][8]
Case 1
_WinAPI_ReleaseDC($avGRP_CTRLIDS[$iIndex][12], $avGRP_CTRLIDS[$iIndex][13])
_GUIImageList_Destroy($avGRP_CTRLIDS[$iIndex][14])
Case 0
Local $HBITMAP = $avGRP_CTRLIDS[$iIndex][14]
For $i = 0 To $avGRP_CTRLIDS[$iIndex][4] - 1
__GRP_DeleteObj($HBITMAP[$i])
Next
$HBITMAP = 0
EndSwitch
EndIf
$hHGMem = DllStructGetData($avGRP_CTRLIDS[$iIndex][1], "HGMEM")
__GRP_FreeMem($avGRP_CTRLIDS[$iIndex][2], $hHGMem)
EndFunc
Func __GRP_WM_DESTROY($hWnd)
Local $iIndex = 1
While $iIndex <= $avGRP_CTRLIDS[0][0]
If $avGRP_CTRLIDS[$iIndex][12] = $hWnd Then
__GRP_ReleaseGIF($iIndex)
For $i = $iIndex To UBound($avGRP_CTRLIDS) - 2
For $j = 0 To 19
$avGRP_CTRLIDS[$i][$j] = $avGRP_CTRLIDS[$i + 1][$j]
Next
Next
ReDim $avGRP_CTRLIDS[$avGRP_CTRLIDS[0][0]][20]
$avGRP_CTRLIDS[0][0] -= 1
$iIndex -= 1
ContinueLoop
EndIf
$iIndex += 1
WEnd
Return $GUI_RUNDEFMSG
EndFunc
Func __GRP_ShutDown()
If $avGRP_CTRLIDS[0][0] Then
For $iIndex = 1 To $avGRP_CTRLIDS[0][0]
__GRP_ReleaseGIF($iIndex)
Next
EndIf
_GDIPlus_Shutdown()
DllClose($hGRP_GDI32)
DllClose($hGRP_USER32)
DllClose($hGRP_COMCTL32)
Exit
EndFunc
Func __GRP_GetCtrlIndex($iCtrlID)
If $avGRP_CTRLIDS[0][0] Then
For $iIndex = 1 To $avGRP_CTRLIDS[0][0]
If $avGRP_CTRLIDS[$iIndex][0] = $iCtrlID Then
Return $iIndex
EndIf
Next
EndIf
Return 0
EndFunc
Func __GRP_GetGifFrameDelays($hOImage, $iFrameCount)
Local $aPropSize, $tPropItem, $aPropItem, $tSize, $tProp, $iDelay
Local $tFrameDelay[$iFrameCount]
$aPropSize = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItemSize", "ptr", $hOImage, "dword", $iGRP_ptFrameDelay, "uint*", 0)
If $aPropSize[0] = 0 Then
$tPropItem = DllStructCreate($tGRP_PropTagItem & ";byte[" & $aPropSize[3] & "]")
$aPropItem = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItem", "ptr", $hOImage, "dword", $iGRP_ptFrameDelay, "dword", $aPropSize[3], "ptr", DllStructGetPtr($tPropItem))
If $aPropItem[0] <> 0 Then Return SetError(1, 0, 0)
$tSize = DllStructGetData($tPropItem, "length")
Switch DllStructGetData($tPropItem, "Type")
Case $iGRP_ptTypeByte
$tProp = DllStructCreate("byte[" & $tSize & "]", DllStructGetData($tPropItem, "value"))
Case $iGRP_ptTypeShort
$tProp = DllStructCreate("short[" & Ceiling($tSize / 2) & "]", DllStructGetData($tPropItem, "value"))
Case $iGRP_ptTypeLong
$tProp = DllStructCreate("long[" & Ceiling($tSize / 4) & "]", DllStructGetData($tPropItem, "value"))
EndSwitch
For $iIndex = 0 To $iFrameCount - 1
$iDelay = DllStructGetData($tProp, 1, $iIndex + 1) * 10
If Not $iDelay Or $iDelay < 50 Then
$iDelay = 50
EndIf
$tFrameDelay[$iIndex] = $iDelay
Next
EndIf
Return $tFrameDelay
EndFunc
Func __GRP_GetGifLoopCount($hOImage)
Local $tPropItem, $tProp, $iSize, $aPropItem, $aPropSize, $iCount
$aPropSize = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItemSize", "ptr", $hOImage, "dword", $iGRP_ptLoopCount, "uint*", 0)
If $aPropSize[0] <> 0 Then Return 0
$tPropItem = DllStructCreate($tGRP_PropTagItem & ";byte[" & $aPropSize[3] & "]")
$aPropItem = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItem", "ptr", $hOImage, "dword", $iGRP_ptLoopCount, "dword", $aPropSize[3], "ptr", DllStructGetPtr($tPropItem))
If $aPropItem[0] <> 0 Then Return SetError(1, 0, 0)
$iSize = DllStructGetData($tPropItem, "length")
Switch DllStructGetData($tPropItem, "Type")
Case $iGRP_ptTypeByte
$tProp = DllStructCreate("byte[" & $iSize & "]", DllStructGetData($tPropItem, "value"))
Case $iGRP_ptTypeShort
$tProp = DllStructCreate("short[" & Ceiling($iSize / 2) & "]", DllStructGetData($tPropItem, "value"))
Case $iGRP_ptTypeLong
$tProp = DllStructCreate("long[" & Ceiling($iSize / 4) & "]", DllStructGetData($tPropItem, "value"))
EndSwitch
$iCount = DllStructGetData($tProp, 1, 1)
If $iCount < 0 Then $iCount = 0
Return $iCount
EndFunc
Func __GRP_GetFileNameType($sFileName, ByRef $hOImage, ByRef $hHGMem)
Select
Case IsBinary($sFileName)
$hHGMem = _MemGlobalAllocFromBinary($sFileName)
$hOImage = _GDIPlus_ImageLoadFromHGlobal($hHGMem)
If @error Then
__GRP_FreeMem($hOImage, $hHGMem)
Return SetError(1, 0, 0)
EndIf
Return SetError(0, 0, 1)
Case StringInStr($sFileName, "/", 0, 1)
$sFileName = InetRead($sFileName)
If @error Then Return SetError(2, 0, 0)
$hHGMem = _MemGlobalAllocFromBinary($sFileName)
$hOImage = _GDIPlus_ImageLoadFromHGlobal($hHGMem)
If @error Then
__GRP_FreeMem($hOImage, $hHGMem)
Return SetError(2, 0, 0)
EndIf
Return SetError(0, 0, 1)
Case StringInStr($sFileName, "|", 0, 1)
Local $asFile = StringSplit($sFileName, "|")
If $asFile[0] < 3 Or Not FileExists($asFile[1]) Then Return SetError(3, 0, 0)
$sFileName = __GRP_LoadResource($asFile[1], $asFile[2], $asFile[3])
If @error Then Return SetError(3, 0, 0)
$hHGMem = _MemGlobalAllocFromBinary(Binary($sFileName))
$hOImage = _GDIPlus_ImageLoadFromHGlobal($hHGMem)
If @error Then
__GRP_FreeMem($hOImage, $hHGMem)
Return SetError(3, 0, 0)
EndIf
Return SetError(0, 0, 1)
Case Else
If Not StringInStr($sFileName, "\") Then $sFileName = @ScriptDir & "\" & $sFileName
If StringInStr($sFileName, ".\") Then $sFileName = StringReplace($sFileName, ".\", @ScriptDir & "\")
If Not FileExists($sFileName) Then Return SetError(4, 0, 0)
$hHGMem = 0
$hOImage = _GDIPlus_ImageLoadFromFile($sFileName)
If @error Then
__GRP_FreeMem($hOImage, $hHGMem)
Return SetError(4, 0, 0)
EndIf
Return SetError(0, 0, 1)
EndSelect
Return SetError(5, 0, 0)
EndFunc
Func __GRP_FreeMem(ByRef $hOImage, ByRef $hMem)
_GDIPlus_ImageDispose($hOImage)
If $hMem Then _MemGlobalFree($hMem)
$hOImage = 0
$hMem = 0
EndFunc
Func __GRP_LoadResource($sFileName, $vResName, $iResType = 10)
Local $hInstance, $InfoBlock, $MemBlock, $pMemPtr
Local $ResStruct, $bResSize, $bResource
$hInstance = DllCall("Kernel32.dll", "ptr", "LoadLibraryW", "wstr", $sFileName)
If @error Or $hInstance[0] = 0 Then Return SetError(1, 0, 0)
$hInstance = $hInstance[0]
$InfoBlock = DllCall("kernel32.dll", "ptr", "FindResourceW", "ptr", $hInstance, "wstr", $vResName, "long", $iResType)
If @error Or $InfoBlock[0] = 0 Then Return SetError(2, 0, 0)
$InfoBlock = $InfoBlock[0]
$bResSize = DllCall("kernel32.dll", "dword", "SizeofResource", "ptr", $hInstance, "ptr", $InfoBlock)
If @error Or $bResSize[0] = 0 Then Return SetError(3, 0, 0)
$bResSize = $bResSize[0]
$MemBlock = DllCall("kernel32.dll", "ptr", "LoadResource", "ptr", $hInstance, "ptr", $InfoBlock)
If @error Or $MemBlock[0] = 0 Then Return SetError(4, 0, 0)
$MemBlock = $MemBlock[0]
$pMemPtr = DllCall("kernel32.dll", "ptr", "LockResource", "ptr", $MemBlock)
If @error Or $pMemPtr[0] = 0 Then
DllCall("Kernel32.dll", "BOOL", "FreeResource", "ptr", $MemBlock)
Return SetError(5, 0, 0)
EndIf
$pMemPtr = $pMemPtr[0]
$ResStruct = DllStructCreate("byte[" & $bResSize & "]", $pMemPtr)
If @error Then Return SetError(6, 0, 0)
$bResSize = DllStructGetSize($ResStruct)
$bResource = DllStructGetData($ResStruct, 1)
$ResStruct = 0
DllCall("Kernel32.dll", "BOOL", "FreeResource", "ptr", $MemBlock)
DllCall("Kernel32.dll", "BOOL", "FreeResource", "ptr", $hInstance)
Return SetError(0, $bResSize, $bResource)
EndFunc
Func __GDIPCreateBitmapFromScan0($iWidth, $iHeight, $iStride = 0, $iPixelFormat = 0x0026200A, $pScan0 = 0)
Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $iWidth, "int", $iHeight, "int", $iStride, "int", $iPixelFormat, "ptr", $pScan0, "int*", 0)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[6]
EndFunc
Func _GDIPlus_ImageLoadFromHGlobal($hGlobal)
Local $aResult = DllCall("ole32.dll", "int", "CreateStreamOnHGlobal", "handle", $hGlobal, "bool", True, "ptr*", 0)
If @error Or $aResult[0] <> 0 Or $aResult[3] = 0 Then Return SetError(1, @error, 0)
Local $hImage = DllCall($__g_hGDIPDll, "uint", "GdipLoadImageFromStream", "ptr", $aResult[3], "int*", 0)
Local $error = @error
Local $tVARIANT = DllStructCreate("word vt;word r1;word r2;word r3;ptr data; ptr")
DllCall("oleaut32.dll", "long", "DispCallFunc", "ptr", $aResult[3], "dword", 8 + 8 * @AutoItX64, "dword", 4, "dword", 23, "dword", 0, "ptr", 0, "ptr", 0, "ptr", DllStructGetPtr($tVARIANT))
If $error Then Return SetError(2, $error, 0)
If $hImage[2] = 0 Then Return SetError(3, 0, $hImage[2])
Return $hImage[2]
EndFunc
Func _MemGlobalAllocFromBinary(Const $bBinary)
Local $iLen = BinaryLen($bBinary)
If $iLen = 0 Then Return SetError(1, 0, 0)
Local $hMem = _MemGlobalAlloc($iLen, $GMEM_MOVEABLE)
If @error Or Not $hMem Then Return SetError(2, 0, 0)
DllStructSetData(DllStructCreate("byte[" & $iLen & "]", _MemGlobalLock($hMem)), 1, $bBinary)
If @error Then
_MemGlobalUnlock($hMem)
_MemGlobalFree($hMem)
Return SetError(3, 0, 0)
EndIf
_MemGlobalUnlock($hMem)
Return $hMem
EndFunc
Global Const $sVersion = "3.2.5"
Global Const $sLoading = @ScriptDir & "\Resources\Loading.gif"
Global Const $INI_FILE = @ScriptDir & "\Wallpaper Explorer.ini"
Example()
Func Example()
Local Const $__sText = "" & "                                                            Product name     : Wallpaper Explorer" & @CRLF & "                                                            Product version   : " & $sVersion & @CRLF & "                                                            Date publish        : 20/07/2015" & @CRLF & "                                                            Author                : Vu Van Duc" & @CRLF & "                                                            Email                   : ducduc08@gmail.com" & @CRLF & @CRLF & @CRLF & "When we use the computer, we often choose desktop bakground in slide show mode, so we can" & @CRLF & "view many pictures stored in our hard disk drivers and work more comfortably. Some of these" & @CRLF & "picture make us do like whereas others do not. Sometimes, we want to delete them immediately" & @CRLF & "from desktop without finding them in their folders to save our time. Unfortunately, none of " & @CRLF & "windows versions supports this feature. And now, with Wallpaper Explorer, you can do it only by" & @CRLF & "one mouse click. I think this program is so useful for you. Let try it..." & @CRLF & "I express my heartfelt thanks to these web-pages below, who help me to write this program: " & @CRLF & "        www.stackoverflow.com*" & @CRLF & "        www.autoitscript.com*" & @CRLF & "        www.msdn.microsoft.com" & @CRLF & "        www.codeproject.com" & @CRLF
Local $__AboutGUI = GUICreate("About Wallpaper Explorer v" & $sVersion, 640, 400)
GUISetBkColor(0xFFFFFF, $__AboutGUI)
Local $iW = _WinAPI_GetClientWidth($__AboutGUI)
Local $iH = _WinAPI_GetClientHeight($__AboutGUI) - 45
Local $__hPic = GUICtrlGetHandle(GUICtrlCreatePic('', 0, 0, $iW, $iH))
Local $__hOK = GUICtrlCreateButton("OK", $iW - 130, $iH + 10, 120, 25)
GUICtrlCreateButton("s", $iW + 100, $iH, 50, 20)
GUICtrlSetState(-1, $GUI_FOCUS)
Local $__hCheckForUpdate = GUICtrlCreateButton("Check for Update", $iW - 130 - 130, $iH + 10, 120, 25)
GUICtrlCreateButton("s", $iW + 100, $iH, 50, 20)
Local $__hLoading = GUICtrlCreatePic('', $iW - 130 - 130, $iH + 10 + 2, 20, 20)
_GUICtrlPic_SetImage(-1, $sLoading)
GUICtrlSetState(-1, $GUI_HIDE)
Local $__hState = GUICtrlCreateLabel("Checking...", $iW - 130 - 130 + 20 + 5, $iH + 13, 120 - 25, 25)
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
GUICtrlSetState(-1, $GUI_HIDE)
Local $__hPercent = GUICtrlCreateLabel('0%', $iW - 130 - 130 - 15, $iH + 10 + 3, 35, 25, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
GUICtrlSetState(-1, $GUI_HIDE)
Local $__hLink = GUICtrlCreateLabel("License User Agreement", 10, $iH + 20, 130, 20, BitOR($SS_NOTIFY, $SS_LEFT))
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlSetColor(-1, 0x3399FF)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlSetCursor(-1, 0)
_GDIPlus_Startup()
Local $hScanBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
Local $hContext = _GDIPlus_ImageGetGraphicsContext($hScanBitmap)
_GDIPlus_GraphicsSetSmoothingMode($hContext, $GDIP_SMOOTHINGMODE_HIGHQUALITY)
_GDIPlus_GraphicsClear($hContext, 0xFFFFFFFF)
$hBrush = _GDIPlus_BrushCreateSolid(0xFF444444)
$hFormat = _GDIPlus_StringFormatCreate()
$hFamily = _GDIPlus_FontFamilyCreate("Segoe UI")
$hFont = _GDIPlus_FontCreate($hFamily, 14, 0, 0)
$tLayout = _GDIPlus_RectFCreate(10, 10, 0, 0)
$aInfo = _GDIPlus_GraphicsMeasureString($hContext, $__sText, $hFont, $tLayout, $hFormat)
_GDIPlus_GraphicsDrawStringEx($hContext, $__sText, $hFont, $aInfo[0], $hFormat, $hBrush)
$hLOGO = _GDIPlus_BitmapCreateFromFile(@ScriptDir & "\Resources\logo.png")
_GDIPlus_GraphicsDrawImage($hContext, $hLOGO, 70, 0)
Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hScanBitmap)
_SendMessage($__hPic, 0x0172, $IMAGE_BITMAP, $hBitmap)
GUISetState()
Local $bIsCheckingUpdate = False
Local $sNowStatus = "Checking...", $temp = "", $bIsHideLoading = False, $hTimeInt = 0
While 1
$nMsg = GUIGetMsg()
Switch $nMsg
Case $__hCheckForUpdate
GUICtrlSetState($__hCheckForUpdate, $GUI_HIDE)
GUICtrlSetState($__hLoading, $GUI_SHOW)
GUICtrlSetState($__hState, $GUI_SHOW)
GUICtrlSetData($__hOK, "Cancel")
ShellExecute(@ScriptDir & '\WE Updater.exe', "/update")
$bIsCheckingUpdate = True
Sleep(100)
GUICtrlSetState($__hOK, $GUI_DISABLE)
Case $__hLink
ShellExecute(@ScriptDir & "\License User Agreement.rtf")
Case $GUI_EVENT_CLOSE, $__hOK
If $bIsCheckingUpdate Then
If MsgBox(64 + 4, "Wallpaper Explorer Update", "Do you want to cancel update program?", 0, $__AboutGUI) = 7 Then
ContinueCase
Else
GUICtrlSetState($__hCheckForUpdate, $GUI_SHOW)
GUICtrlSetState($__hLoading, $GUI_HIDE)
GUICtrlSetState($__hState, $GUI_HIDE)
GUICtrlSetState($__hPercent, $GUI_HIDE)
GUICtrlSetData($__hOK, "OK")
IniWrite($INI_FILE, "UPDATE", "StopNow", "True")
$bIsCheckingUpdate = False
$bIsHideLoading = False
GUICtrlSetState($__hOK, $GUI_ENABLE)
EndIf
Else
ExitLoop
EndIf
EndSwitch
If $bIsCheckingUpdate Then
If TimerDiff($hTimeInt) >= 100 Then
$temp = IniRead($INI_FILE, "UPDATE", "State", "")
If $temp <> $sNowStatus And $temp <> "" Then
GUICtrlSetData($__hState, $temp)
$sNowStatus = $temp
EndIf
If $sNowStatus = "Downloading..." Then
If Not $bIsHideLoading Then
GUICtrlSetState($__hLoading, $GUI_HIDE)
GUICtrlSetState($__hPercent, $GUI_SHOW)
GUICtrlSetState($__hOK, $GUI_ENABLE)
$bIsHideLoading = True
EndIf
GUICtrlSetData($__hPercent, IniRead($INI_FILE, "UPDATE", "Percent", "%s\%"))
EndIf
If $sNowStatus = "Complete" Or $sNowStatus = "Error" Or $sNowStatus = "Finish" Then
GUICtrlSetState($__hCheckForUpdate, $GUI_SHOW)
GUICtrlSetState($__hLoading, $GUI_HIDE)
GUICtrlSetState($__hState, $GUI_HIDE)
GUICtrlSetState($__hPercent, $GUI_HIDE)
GUICtrlSetData($__hOK, "OK")
IniWrite($INI_FILE, "UPDATE", "StopNow", "True")
$bIsCheckingUpdate = False
$bIsHideLoading = False
GUICtrlSetState($__hOK, $GUI_ENABLE)
EndIf
$hTimeInt = TimerInit()
EndIf
EndIf
WEnd
_WinAPI_DeleteObject($hBitmap)
_GDIPlus_FontDispose($hFont)
_GDIPlus_FontFamilyDispose($hFamily)
_GDIPlus_StringFormatDispose($hFormat)
_GDIPlus_BrushDispose($hBrush)
_GDIPlus_GraphicsDispose($hContext)
_GDIPlus_BitmapDispose($hScanBitmap)
_GDIPlus_Shutdown()
EndFunc
