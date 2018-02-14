#NoTrayIcon
Global Const $__FILE_BEGIN = 0
Global Const $__FILE_CURRENT = 1
Func _ImageGetInfo($sFile)
Local $sInfo = "", $hFile, $nClr, $error
Local $nFileSize = FileGetSize($sFile)
Local $ret = DllCall("kernel32.dll","int","CreateFile", "str",$sFile, "int",0x80000000, "int",0, "ptr",0, "int",3, "int",0x80, "ptr",0)
If @error OR($ret[0] = 0) Then Return SetError(1, 0, "")
$hFile = $ret[0]
Local $asIdent[7] = ["0xFFD8", "0x424D", "0x89504E470D0A1A0A", "0x4749463839", "0x4749463837", "0x4949", "0x4D4D"]
Local $p = _FileReadToStructAtOffset("ubyte[54]", $hFile, 0)
For $i = 0 To UBound($asIdent) - 1
If BinaryMid(DllStructGetData($p, 1), 1, BinaryLen($asIdent[$i])) = $asIdent[$i] Then
Switch $i
Case 0
$sInfo = _ParseJPEG($hFile, $nFileSize)
$error = @error
Case 1
Local $t = DllStructCreate("int;int;short;short;dword;dword;dword;dword", DllStructGetPtr($p) + 18)
_Add($sInfo, "Width", DllStructGetData($t, 1))
_Add($sInfo, "Height", DllStructGetData($t, 2))
_Add($sInfo, "ColorDepth", DllStructGetData($t, 4))
_Add($sInfo, "XResolution", Round(DllStructGetData($t, 7)/39.37))
_Add($sInfo, "YResolution", Round(DllStructGetData($t, 8)/39.37))
_Add($sInfo, "ResolutionUnit", "Inch")
Case 2
$sInfo = _ParsePNG($hFile, $nFileSize)
$error = @error
Case 3, 4
$t = DllStructCreate("short;short;ubyte", DllStructGetPtr($p) + 6)
_Add($sInfo, "Width", DllStructGetData($t, 1))
_Add($sInfo, "Height", DllStructGetData($t, 2))
$nClr = DllStructGetData($t, 3)
_Add($sInfo, "ColorDepth", _IsBitSet($nClr, 0) + _IsBitSet($nClr, 1)*2 + _IsBitSet($nClr, 2)*4 + 1)
Case 5, 6
$sInfo = _ParseTIFF($hFile, 0)
EndSwitch
Exitloop
Endif
Next
DllCall("kernel32.dll","int","CloseHandle","int", $hFile)
Return SetError($error, 0, $sInfo)
EndFunc
Func _ParsePNG($hFile, $nFileSize)
Local $sInfo = "", $nBlockSize, $nID = 0, $t
Local $nBPP, $nCol, $sAlpha = ""
Local $nXRes, $nYRes
Local $sKeyword, $nKWLen
Local $pBlockID = DllStructCreate("ulong;ulong")
_FileSetPointer($hFile, 8, $__FILE_BEGIN)
While $nID <> 0x49454E44
$pBlockID = _FileReadToStruct($pBlockID, $hFile)
$nBlockSize = _IntR32(DllStructGetData($pBlockID, 1))
If $nBlockSize > $nFileSize Then Return SetError(2, 0, $sInfo)
$nID = _IntR32(DllStructGetData($pBlockID, 2))
Switch $nID
Case 0x49484452
$t = _FileReadToStruct("align 1;ulong;ulong;byte;byte;byte;byte;byte", $hFile)
_Add($sInfo, "Width",  _IntR32(DllStructGetData($t, 1)))
_Add($sInfo, "Height", _IntR32(DllStructGetData($t, 2)))
$nBPP = DllStructGetData($t, 3)
$nCol = DllStructGetData($t, 4)
If $nCol > 3 Then
$nCol = $nCol - 4
$sAlpha = " + alpha"
Endif
If $nCol < 3 Then $nBPP =($nCol + 1) * $nBPP
_Add($sInfo, "ColorDepth", $nBPP & $sAlpha)
_Add($sInfo, "Interlace", DllStructGetData($t, 7))
Case 0x70485973
$t = _FileReadToStruct("align 1;ulong;ulong;ubyte", $hFile)
$nXRes = _IntR32(DllStructGetData($t, 1))
$nYRes = _IntR32(DllStructGetData($t, 2))
If DllStructGetData($t, 3) = 1 Then
$nXRes = Round($nXRes/39.37)
$nYRes = Round($nYRes/39.37)
Endif
_Add($sInfo, "XResolution", $nXRes)
_Add($sInfo, "YResolution", $nYRes)
_Add($sInfo, "ResolutionUnit", "Inch")
Case 0x74455874
$t = _FileReadToStruct("char[80]", $hFile)
$sKeyword = DllStructGetData($t, 1)
$nKWLen = StringLen($sKeyword) + 1
_FileSetPointer($hFile, -80 + $nKWLen, $__FILE_CURRENT)
$t = _FileReadToStruct("char[" & $nBlockSize - $nKWLen & "]", $hFile)
_Add($sInfo, $sKeyword, DllStructGetData($t, 1))
Case 0x74494D45
$t = _FileReadToStruct("align 1;ushort;ubyte;ubyte;ubyte;ubyte;ubyte", $hFile)
_Add($sInfo, "DateTime", StringFormat("%4d-%02d-%02d %02d:%02d:%02d", DllStructGetData($t, 1), DllStructGetData($t, 2), DllStructGetData($t, 3), DllStructGetData($t, 4), DllStructGetData($t, 5), DllStructGetData($t, 6)))
Case Else
_FileSetPointer($hFile, $nBlockSize + 4, $__FILE_CURRENT)
ContinueLoop
EndSwitch
_FileSetPointer($hFile, 4, $__FILE_CURRENT)
Wend
Return $sInfo
EndFunc
Func _ParseJPEG($hFile, $nFileSize)
Local $nPos = 2, $nMarker = 0, $sInfo = ""
Local $nSegSize, $pSegment, $t
_FileSetPointer($hFile, 2, $__FILE_BEGIN)
Local $p = DllStructCreate("align 1;ubyte;ubyte;ushort")
While($nMarker <> 0xDA) and($nPos < $nFileSize)
$p = _FileReadToStruct($p, $hFile)
If DllStructGetData($p, 1) = 0xFF Then
$nMarker = DllStructGetData($p, 2)
$nSegSize = _IntR16(DllStructGetData($p, 3))
$pSegment = _FileReadToStruct("byte[" & $nSegSize - 2 & "]", $hFile)
Switch $nMarker
Case 0xC0 To 0xC3, 0xC5 To 0xC7, 0xCB, 0xCD To 0xCF
Switch $nMarker
Case 0xC2, 0xC6, 0xCA, 0xCE
_Add($sInfo, "Progressive", "True")
Case 0xC3, 0xC7, 0xCB, 0xCF
_Add($sInfo, "Lossless", "True")
EndSwitch
Switch $nMarker
Case 0xC0 To 0xC3, 0xC5 To 0xC7
_Add($sInfo, "Compression", "Huffman")
Case 0xC9 To 0xCB, 0xCD To 0xCF
_Add($sInfo, "Compression", "Arithmetic")
EndSwitch
$t = DllStructCreate("align 1;byte;ushort;ushort", DllStructGetPtr($pSegment))
_Add($sInfo, "Width",  _IntR16(DllStructGetData($t, 3)))
_Add($sInfo, "Height", _IntR16(DllStructGetData($t, 2)))
Case 0xE0
$t = DllStructCreate("byte[5];byte;byte;ubyte;ushort;ushort", DllStructGetPtr($pSegment))
_Add($sInfo, "XResolution", _IntR16(DllStructGetData($t, 5)))
_Add($sInfo, "YResolution", _IntR16(DllStructGetData($t, 6)))
_AddArray($sInfo, "ResolutionUnit", _IntR16(DllStructGetData($t, 4)), "Pixel;Inch;Cm")
Case 0xE1
$sInfo = $sInfo & _ParseTIFF($hFile, $nPos + 10)
_FileSetPointer($hFile, $nPos + $nSegSize + 2, $__FILE_BEGIN)
Case 0xFE
$t = DllStructCreate("char[" & $nSegSize - 2 & "]", DllStructGetPtr($pSegment))
_Add($sInfo, "Comment", StringStripWS(DllStructGetData($t, 1), 3))
EndSwitch
$nPos= $nPos + $nSegSize + 2
Else
Return SetError(2, 0, $sInfo)
Endif
Wend
Return($sInfo)
EndFunc
Func _ParseTIFF($hFile, $nTiffHdrOffset = 0)
Local $pHdr, $nIFDOffset, $pCnt, $nIFDCount, $pTag, $nCnt, $nID, $nEIFDCount
Local $ByteOrder = 0, $sInfo = ""
Local $nEIFDOffset, $nSubID
$pHdr = _FileReadToStructAtOffset("short;short;dword", $hFile, $nTiffHdrOffset)
If DllStructGetData($pHdr, 1) = 0x4D4D then $ByteOrder = 1
$nIFDOffset = _IntR32(DllStructGetData($pHdr, 3), $ByteOrder)
$pCnt = _FileReadToStructAtOffset("ushort", $hFile, $nTiffHdrOffset + $nIFDOffset)
$nIFDCount = _IntR16(DllStructGetData($pCnt, 1), $ByteOrder)
$pTag = DllStructCreate("ushort;ushort;ulong;ulong")
For $nCnt = 0 To $nIFDCount - 1
$pTag = _FileReadToStructAtOffset($pTag, $hFile, $nTiffHdrOffset + $nIFDOffset + 2 + $nCnt * 12)
$nID = _IntR16(DllStructGetData($pTag, 1), $ByteOrder)
Switch $nID
Case 0x8769, 0x8825
$nEIFDOffset = _ReadTag($hFile, $pTag, $nTiffHdrOffset, $ByteOrder)
$pCnt = _FileReadToStructAtOffset($pCnt, $hFile, $nTiffHdrOffset + $nEIFDOffset)
$nEIFDCount = _IntR16(DllStructGetData($pCnt, 1), $ByteOrder)
For $i = 0 To $nEIFDCount - 1
$pTag = _FileReadToStructAtOffset($pTag, $hFile, $nTiffHdrOffset + $nEIFDOffset + 2 + $i * 12)
$nSubID = _IntR16(DllStructGetData($pTag, 1), $ByteOrder)
Switch $nID
Case 0x8769
_AddExifTag($sInfo, $nSubID, _ReadTag($hFile, $pTag, $nTiffHdrOffset, $ByteOrder))
Case 0x8825
_AddGpsTag($sInfo, $nSubID, _ReadTag($hFile, $pTag, $nTiffHdrOffset, $ByteOrder))
EndSwitch
Next
Case 0xA005
Case 0x014A
Case 0x0190
Case Else
_AddTiffTag($sInfo, $nID, _ReadTag($hFile, $pTag, $nTiffHdrOffset, $ByteOrder))
EndSwitch
Next
Return($sInfo)
EndFunc
Func _ReadTag($hFile, $pTag, $nHdrOffset, $ByteOrder)
Local $nType = _IntR16(DllStructGetData($pTag, 2), $ByteOrder)
Local $nCount = _IntR32(DllStructGetData($pTag, 3), $ByteOrder)
Local $nOffset = _IntR32(DllStructGetData($pTag, 4), $ByteOrder) + $nHdrOffset
Local $p
Switch $nType
Case 2
$p = _FileReadToStructAtOffset("char[" & $nCount & "]", $hFile, $nOffset)
Return DllStructGetData($p, 1)
Case 1
Return BitAND(DllStructGetData($pTag, 4), 0xFF)
Case 3
If $nCount > 1 Then
If $nCount > 2 Then
$p = _FileReadToStructAtOffset("short[" & $nCount & "]", $hFile, $nOffset)
Else
$p = DllStructCreate("short[2]", DllStructGetPtr($pTag, 4))
EndIf
Local $aRet[$nCount]
For $i = 1 To $nCount
$aRet[$i-1] = _IntR16(DllStructGetData($p, 1, $i), $ByteOrder)
Next
Return $aRet
Else
Return _IntR16(BitAND(DllStructGetData($pTag, 4), 0xFFFF), $ByteOrder)
EndIf
Case 4
If $nCount > 1 Then
$p = _FileReadToStructAtOffset("short[" & $nCount & "]", $hFile, $nOffset)
Local $aRet[$nCount]
For $i = 1 To $nCount
$aRet[$i-1] = _IntR32(DllStructGetData($p, 1, $i), $ByteOrder)
Next
Return $aRet
Else
Return _IntR32(DllStructGetData($pTag, 4), $ByteOrder)
EndIf
Case 5
$p = _FileReadToStructAtOffset("ulong;ulong", $hFile, $nOffset)
Return _IntR32(DllStructGetData($p, 1), $ByteOrder) / _IntR32(DllStructGetData($p, 2), $ByteOrder)
Case 7
$p = _FileReadToStructAtOffset("byte[" & $nCount & "]", $hFile, $nOffset)
Return DllStructGetData($p, 1)
Case 9
$p = DllStructCreate("long", DllStructGetPtr($pTag, 4))
Return _IntR32(DllStructGetData($p, 1), $ByteOrder)
Case 10
$p = _FileReadToStructAtOffset("dword;dword", $hFile, $nOffset)
Return _IntR32(DllStructGetData($p, 1), $ByteOrder) / _IntR32(DllStructGetData($p, 2), $ByteOrder)
Case Else
Return SetError(1, 0, "")
EndSwitch
EndFunc
Func _AddTiffTag(ByRef $sInfo, $nID, $vTag)
Local $p, $vTemp = ""
Switch $nID
Case 0x00FE
Select
Case _IsBitSet($vTag, 0)
$vTemp = "Reduced image"
Case _IsBitSet($vTag, 1)
$vTemp = "Page"
Case _IsBitSet($vTag, 2)
$vTemp = "Mask"
EndSelect
_Add($sInfo, "NewSubfileType", $vTemp)
Case 0x0100
_Add($sInfo, "Width", $vTag)
Case 0x0101
_Add($sInfo, "Height", $vTag)
Case 0x0102
If IsArray($vTag) Then
$vTemp = 0
For $i = 0 To UBound($vTag) - 1
$vTemp += $vTag[$i]
Next
_Add($sInfo, "Colordepth", $vTemp)
Else
_Add($sInfo, "Colordepth", $vTag)
EndIf
Case 0x0103
Switch $vTag
Case 32773
_Add($sInfo, "Compression", "Packbits")
Case Else
_AddArray($sInfo, "Compression", $vTag, ";No compression;CCITT modified Huffman RLE;CCITT Group 3 fax encoding;" & "CCITT Group 4 fax encoding;LZW;JPEG (old);JPEG;Deflate;JBIG on B/W;JBIG on color","Unknown")
EndSwitch
Case 0x0106
Switch $vTag
Case 32803
_Add($sInfo, "PhotometricInterpretation", "LOGL")
$vTemp = "LOGL"
Case 34892
_Add($sInfo, "PhotometricInterpretation", "LOGLUV")
Case Else
_AddArray($sInfo, "PhotometricInterpretation", $vTag, "MINISWHITE;MINISBLACK;RGB;PALETTE;MASK;SEPARATED;YCBCR;CIELAB;ICCLAB;ITULAB")
EndSwitch
Case 0x0107
_AddArray($sInfo, "Threshholding", $vTag, ";BiLevel;Halftone;Error Diffusion")
Case 0x0108
_Add($sInfo, "CellWidth", $vTemp)
Case 0x0109
_Add($sInfo, "CellLength", $vTemp)
Case 0x010E
_Add($sInfo, "ImageDescription", $vTag)
Case 0x010F
_Add($sInfo, "Make", $vTag)
Case 0x0110
_Add($sInfo, "Model", $vTag)
Case 0x0112
_AddArray($sInfo, "Orientation", $vTag, ";Normal;Mirrored;180°;180° and mirrored;90° left and mirrored;90° right;90° right and mirrored;90° left")
Case 0x0115
_Add($sInfo, "SamplesPerPixel", $vTag)
Case 0x011A
_Add($sInfo, "XResolution", $vTag)
Case 0x011B
_Add($sInfo, "YResolution", $vTag)
Case 0x0128
_AddArray($sInfo, "ResolutionUnit", $vTag, ";;Inch;Centimeter")
Case 0x0131
_Add($sInfo, "Software", $vTag)
Case 0x0132
_Add($sInfo, "DateTime", $vTag)
Case 0x013B
_Add($sInfo, "Artist", $vTag)
Case 0x013C
_Add($sInfo, "HostComputer", $vTag)
Case 0x0152
_AddArray($sInfo, "ExtraSamples", $vTag, ";Associated alpha;Unassociated alpha","Unspecified")
Case 0x8298
_Add($sInfo, "Copyright", $vTag)
Case 0x010D
_Add($sInfo, "DocumentName", $vTag)
Case 0x011D
_Add($sInfo, "PageName", $vTag)
Case 0x011E
_Add($sInfo, "XPosition", $vTag)
Case 0x011F
_Add($sInfo, "YPosition", $vTag)
Case 0x0129
If IsArray($vTag) Then
_Add($sInfo, "PageNumber", $vTag[0] & "/" & $vTag[1])
EndIf
Case 0x013E
_Add($sInfo, "WhitePoint", $vTag)
Case 0x0146
_Add($sInfo, "BadFaxLines", $vTag)
Case 0x0147
_AddArray($sInfo, "CleanFaxData", $vTag, "Clean;Regenerated;Unclean")
Case 0x0148
_Add($sInfo, "ConsecutiveBadFaxLines", $vTag)
Case 0x014C
_AddArray($sInfo, "InkSet", $vTag, ";CMYK", "No CMYK")
Case 0x014E
_Add($sInfo, "NumberOfInks", $vTag)
Case 0x0151
_Add($sInfo, "TargetPrinter", $vTag)
Case 0x015A
_AddArray($sInfo, "Indexed", $vTag, "Not indexed;Indexed")
Case 0x0191
_AddArray($sInfo, "ProfileType", $vTag, ";Group 3 fax","Unspecified")
Case 0x02BC
_Add($sInfo, "XMP", $vTag)
EndSwitch
EndFunc
Func _AddExifTag(ByRef $sInfo, $nID, $vTag)
Local $vTemp = ""
Switch $nID
Case 0x829A
_Add($sInfo, "ExposureTime", $vTag)
Case 0x829D
_Add($sInfo, "FNumber", $vTag)
Case 0x8822
_AddArray($sInfo, "ExposureProgram", $vTag, "Undefined;Manual;Normal;Aperture priority;Shutter priority;Creative;Action;Portrait mode;Landscape mode")
Case 0x8824
_Add($sInfo, "SpectralSensitivity", $vTag)
Case 0x8827
_Add($sInfo, "ISO", $vTag)
Case 0x9000
_Add($sInfo, "ExifVersion", $vTag)
Case 0x9003
_Add($sInfo, "DateTimeOriginal", $vTag)
Case 0x9004
_Add($sInfo, "DateTimeDigitized", $vTag)
Case 0x9101
Switch BinaryMid($vTag, 1, 1)
Case 0x34
$vTemp = "RGB"
Case 0x31
$vTemp = "YCbCr"
Case Else
$vTemp = "Undefined"
EndSwitch
_Add($sInfo, "ComponentsConfiguration", $vTemp)
Case 0x9102
_Add($sInfo, "CompressedBitsPerPixel", $vTag)
Case 0x9201
_Add($sInfo, "ShutterSpeedValue", $vTag)
Case 0x9202
_Add($sInfo, "ApertureValue", $vTag)
Case 0x9203
_Add($sInfo, "BrightnessValue", $vTag)
Case 0x9204
_Add($sInfo, "ExposureBiasValue", $vTag)
Case 0x9205
_Add($sInfo, "MaxApertureValue", $vTag)
Case 0x9206
_Add($sInfo, "SubjectDistance", $vTag)
Case 0x9207
_AddArray($sInfo, "MeteringMode", $vTag, "Unknown;Average;Center Weighted Average;Spot;MultiSpot;MultiSegment;Partial;Other")
Case 0x9208
If $vTag = 255 Then
_Add($sInfo, "LightSource", "Other")
Else
_AddArray($sInfo, "LightSource", $vTag, "Unknown;Daylight;Fluorescent;Tungsten;Flash;;;;;Fine weather;Cloudy weather;Shade;" & "Daylight fluorescent;Day white fluorescent;Cool white fluorescent;White fluorescent;;" & "Standard light A;Standard light B;Standard light C;D55;D65;D75;D50;ISO studio tungsten")
EndIf
Case 0x9209
If _IsBitSet($vTag, 0) Then
$vTemp = "Fired, "
Else
$vTemp = "Not fired, "
EndIf
Switch _IsBitSet($vTag, 4) * 2 + _IsBitSet($vTag, 3)
Case 1
$vTemp &= "Forced ON, "
Case 2
$vTemp &= "Forced OFF, "
Case 3
$vTemp &= "Auto, "
EndSwitch
If _IsBitSet($vTag, 6) Then $vTemp &= "Red-eye reduction, "
_Add($sInfo, "Flash", StringTrimRight($vTemp, 2))
Case 0x920A
_Add($sInfo, "FocalLength", $vTag)
Case 0x9214
If IsArray($vTag) Then
Switch UBound($vTag)
Case 2
$vTemp = StringFormat("Point, %d:%d", $vTag[0], $vTag[1])
Case 3
$vTemp = StringFormat("Circle, Center %d:%d, Diameter %d", $vTag[0], $vTag[1], $vTag[2])
Case 4
$vTemp = StringFormat("Square, Center %d:%d, Width %d, Height %d", $vTag[0], $vTag[1], $vTag[2], $vTag[3])
EndSwitch
EndIf
_Add($sInfo, "SubjectArea", $vTemp)
Case 0x9286
Switch BinaryMid($vTag, 1, 8)
Case "0x4153434949000000"
$vTemp = BinaryToString(BinaryMid($vTag, 9))
Case "0x554E49434F444500"
$vTemp = BinaryMid($vTag, 9)
Case Else
$vTemp = BinaryMid($vTag, 9)
EndSwitch
_Add($sInfo, "UserComment", $vTemp)
Case 0x9290
_Add($sInfo, "SubsecTime", $vTag)
Case 0x9291
_Add($sInfo, "SubsecTimeOriginal", $vTag)
Case 0x9292
_Add($sInfo, "SubsecTimeDigitized", $vTag)
Case 0xA001
_AddArray($sInfo, "ColorSpace", $vTag, ";RGB", "Uncalibrated")
Case 0xA002
_Add($sInfo, "PixelXDimension", $vTag)
Case 0xA003
_Add($sInfo, "PixelYDimension", $vTag)
Case 0xA004
_Add($sInfo, "RelatedSoundFile", $vTag)
Case 0xA20B
_Add($sInfo, "FlashEnergy", $vTag)
Case 0xA20E
_Add($sInfo, "FocalPlaneXResolution", $vTag)
Case 0xA20F
_Add($sInfo, "FocalPlaneYResolution", $vTag)
Case 0xA210
_AddArray($sInfo, "FocalPlaneResolutionUnit", $vTag, ";;Inch;Centimeter")
Case 0xA214
If IsArray($vTag) Then
_Add($sInfo, "SubjectLocation", $vTag[0] & ":" & $vTag[1])
EndIf
Case 0xA215
_Add($sInfo, "ExposureIndex", $vTag)
Case 0xA217
_AddArray($sInfo, "SensingMethod", $vTag, "Undefined;Undefined;OneChipColorArea;TwoChipColorArea;ThreeChipColorArea;ColorSequentialArea;Undefined;Trilinear;ColorSequentialLinear")
Case 0xA300
If $vTag = 0x3 Then _Add($sInfo, "FileSource", "DSC")
Case 0xA301
If $vTag = 0x1 Then _Add($sInfo, "SceneType", "Directly recorded")
Case 0xA401
_AddArray($sInfo, "CustomRendered", "No;Yes","No")
Case 0xA402
_AddArray($sInfo, "ExposureMode", $vTag, "Auto;Manual;Auto bracket")
Case 0xA403
_AddArray($sInfo, "WhiteBalance", $vTag, "Auto;Manual")
Case 0xA404
_Add($sInfo, "DigitalZoomRatio", $vTag)
Case 0xA405
_AddArray($sInfo, "FocalLengthIn35mmFilm", $vTag, "Unknown", $vTag)
Case 0xA406
_AddArray($sInfo, "SceneCaptureType", $vTag, "Standard;Landscape;Portrait;Night scene")
Case 0xA407
_AddArray($sInfo, "GainControl", $vTag, "None;Low gain up;High gain up;Low gain down;High gain down")
Case 0xA408
_AddArray($sInfo, "Contrast", $vTag, "Normal;Soft;Hard")
Case 0xA409
_AddArray($sInfo, "Saturation", $vTag, "Normal;Low;High")
Case 0xA40A
_AddArray($sInfo, "Sharpness", $vTag, "Normal;Soft;Hard")
Case 0xA40B
_Add($sInfo, "DeviceSettingDescription", $vTag)
Case 0xA40C
_AddArray($sInfo, "SubjectDistanceRange", $vTag, "Unknown;Macro;Close view;Distant view")
Case 0xA420
_Add($sInfo, "ImageUniqueID", $vTag)
EndSwitch
EndFunc
Func _AddGpsTag(ByRef $sInfo, $nID, $vTag)
Switch $nID
Case 0x0000
_Add($sInfo, "GPSVersionID", $vTag)
Case 0x0001
_Add($sInfo, "GPSLatitudeRef", $vTag)
Case 0x0002
If IsArray($vTag) And(UBound($vTag) = 3) Then
_Add($sInfo, "GPSLatitude", StringFormat("%f/%f/%f", $vTag[0], $vTag[1], $vTag[2]))
EndIf
Case 0x0003
_Add($sInfo, "GPSLongitudeRef", $vTag)
Case 0x0004
If IsArray($vTag) And(UBound($vTag) = 3) Then
_Add($sInfo, "GPSLongitude", StringFormat("%f/%f/%f", $vTag[0], $vTag[1], $vTag[2]))
EndIf
Case 0x0005
_AddArray($sInfo, "GPSAltitudeRef", $vTag, "Above sea level,Below sea level")
Case 0x0006
_Add($sInfo, "GPSAltitude", $vTag)
Case 0x0007
If IsArray($vTag) And(UBound($vTag) = 3) Then
_Add($sInfo, "GPSTimeStamp", StringFormat("UTC %f:%f:%f", $vTag[0], $vTag[1], $vTag[2]))
EndIf
Case 0x0008
_Add($sInfo, "GPSSatellites", $vTag)
Case 0x0009
_Add($sInfo, "GPSStatus", $vTag)
Case 0x000A
_Add($sInfo, "GPSMeasureMode", $vTag)
Case 0x000B
_Add($sInfo, "GPSDOP", $vTag)
Case 0x000C
_Add($sInfo, "GPSSpeedRef", $vTag)
Case 0x000D
_Add($sInfo, "GPSSpeed", $vTag)
Case 0x000E
_Add($sInfo, "GPSTrackRef", $vTag)
Case 0x000F
_Add($sInfo, "GPSTrack", $vTag)
Case 0x0010
_Add($sInfo, "GPSImgDirectionRef", $vTag)
Case 0x0011
_Add($sInfo, "GPSImgDirection", $vTag)
Case 0x0012
_Add($sInfo, "GPSMapDatum", $vTag)
Case 0x0013
_Add($sInfo, "GPSDestLatitudeRef", $vTag)
Case 0x0014
If IsArray($vTag) And(UBound($vTag) = 3) Then
_Add($sInfo, "GPSDestLatitude", StringFormat("%f/%f/%f", $vTag[0], $vTag[1], $vTag[2]))
EndIf
Case 0x0015
_Add($sInfo, "GPSDestLongitudeRef", $vTag)
Case 0x0016
If IsArray($vTag) And(UBound($vTag) = 3) Then
_Add($sInfo, "GPSDestLongitude", StringFormat("%f/%f/%f", $vTag[0], $vTag[1], $vTag[2]))
EndIf
Case 0x0017
_Add($sInfo, "GPSDestBearingRef", $vTag)
Case 0x0018
_Add($sInfo, "GPSDestBearing", $vTag)
Case 0x0019
_Add($sInfo, "GPSDestDistanceRef", $vTag)
Case 0x001A
_Add($sInfo, "GPSDestDistance", $vTag)
Case 0x001D
_Add($sInfo, "GPSDateStamp", $vTag)
Case 0x001E
_AddArray($sInfo, "GPSDifferential", $vTag, "No correction;Correction applied")
EndSwitch
EndFunc
Func _AddArray(ByRef $sInfo, $sField, $vTag, $sValues, $sDefault = "Undefined")
Local $aArray = StringSplit($sValues, ";", 2)
Switch $vTag
Case 0 To UBound($aArray) - 1
If $aArray[$vTag] = "" Then $aArray[$vTag] = $sDefault
_Add($sInfo, $sField, $aArray[$vTag])
Case Else
_Add($sInfo, $sField, $sDefault)
EndSwitch
EndFunc
Func _IsBitSet($nNum, $nBit)
Return BitAND(BitShift($nNum, $nBit), 1)
EndFunc
Func _Add(ByRef $sInfo, $sField, $nValue)
$sInfo = $sInfo & $sField & "=" & $nValue & @LF
EndFunc
Func _IntR32($nInt, $nOrder = 1)
If not $nOrder Then Return $nInt
Return BitOR(BitAND(BitShift($nInt,24), 0x000000FF), BitAND(BitShift($nInt, 8), 0x0000FF00), BitAND(BitShift($nInt,-8), 0x00FF0000), BitAND(BitShift($nInt,-24),0xFF000000))
EndFunc
Func _IntR16($nInt, $nOrder = 1)
If not $nOrder Then Return $nInt
Return BitOR(BitAND(BitShift($nInt, 8), 0xFF), BitAND(BitShift($nInt, -8), 0xFF00))
EndFunc
Func _FileSetPointer($hFile, $nOffset, $nStartPosition)
Local $ret = DllCall("kernel32.dll","int","SetFilePointer", "int",$hFile, "int",$nOffset, "ptr",0, "int",$nStartPosition)
EndFunc
Func _FileReadToStruct($vStruct, $hFile)
If not DllStructGetSize($vStruct) Then $vStruct = DllStructCreate($vStruct)
Local $nLen = DllStructGetSize($vStruct)
Local $pRead = DllStructCreate("dword")
Local $ret = DllCall("kernel32.dll","int","ReadFile", "int",$hFile, "ptr",DllStructGetPtr($vStruct), "int", $nLen, "ptr",DllStructGetPtr($pRead), "ptr",0)
Local $nRead = DllStructGetData($pRead, 1)
SetExtended($nRead)
If not($nRead = $nLen) Then SetError(2)
Return $vStruct
EndFunc
Func _FileReadToStructAtOffset($vStruct, $hFile, $nOffset)
_FileSetPointer($hFile, $nOffset, $__FILE_BEGIN)
Return SetError(@error, @extended, _FileReadToStruct($vStruct, $hFile))
EndFunc
Global Const $UBOUND_COLUMNS = 2
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Func _WinAPI_GetLastError($iError = @error, $iExtended = @extended)
Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($iError, $iExtended, $aResult[0])
EndFunc
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $pNewState, $iBufferLen, $pPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $pNewState, "dword", $iBufferLen, "struct*", $pPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
If $hThread = 0 Then
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 10, @extended, 0)
$hThread = $aResult[0]
EndIf
Local $aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then Return SetError(3, @error, False)
Return True
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
Return $aResult
EndFunc
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagSIZE = "struct;long X;long Y;endstruct"
Global Const $tagGDIPSTARTUPINPUT = "uint Version;ptr Callback;bool NoThread;bool NoCodecs"
Global Const $tagGDIPIMAGECODECINFO = "byte CLSID[16];byte FormatID[16];ptr CodecName;ptr DllName;ptr FormatDesc;ptr FileExt;" & "ptr MimeType;dword Flags;dword Version;dword SigCount;dword SigSize;ptr SigPattern;ptr SigMask"
Global Const $tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagBLENDFUNCTION = "byte Op;byte Flags;byte Alpha;byte Format"
Global Const $tagGUID = "struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];endstruct"
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $ULW_ALPHA = 0x02
Global Const $WH_MOUSE_LL = 14
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $DI_MASK = 0x0001
Global Const $DI_IMAGE = 0x0002
Global Const $DI_NORMAL = 0x0003
Global Const $DI_COMPAT = 0x0004
Global Const $DI_DEFAULTSIZE = 0x0008
Global Const $DI_NOMIRROR = 0x0010
Global Const $GW_HWNDNEXT = 2
Global Const $GW_CHILD = 5
Global Const $IMAGE_BITMAP = 0
Global Const $IMAGE_ICON = 1
Global Const $IMAGE_CURSOR = 2
Global Const $LR_DEFAULTCOLOR = 0x0000
Global Const $LR_LOADFROMFILE = 0x0010
Global Const $LR_LOADTRANSPARENT = 0x0020
Global Const $LR_SHARED = 0x8000
Global $__g_aInProcess_WinAPI[64][2] = [[0, 0]]
Global $__g_aWinList_WinAPI[64][2] = [[0, 0]]
Global Const $__WINAPICONSTANT_FW_NORMAL = 400
Global Const $__WINAPICONSTANT_DEFAULT_CHARSET = 1
Global Const $__WINAPICONSTANT_OUT_DEFAULT_PRECIS = 0
Global Const $__WINAPICONSTANT_CLIP_DEFAULT_PRECIS = 0
Global Const $__WINAPICONSTANT_DEFAULT_QUALITY = 0
Func _WinAPI_CallNextHookEx($hHk, $iCode, $wParam, $lParam)
Local $aResult = DllCall("user32.dll", "lresult", "CallNextHookEx", "handle", $hHk, "int", $iCode, "wparam", $wParam, "lparam", $lParam)
If @error Then Return SetError(@error, @extended, -1)
Return $aResult[0]
EndFunc
Func _WinAPI_ClientToScreen($hWnd, ByRef $tPoint)
Local $aRet = DllCall("user32.dll", "bool", "ClientToScreen", "hwnd", $hWnd, "struct*", $tPoint)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Return $tPoint
EndFunc
Func _WinAPI_CreateCompatibleDC($hDC)
Local $aResult = DllCall("gdi32.dll", "handle", "CreateCompatibleDC", "handle", $hDC)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreateFont($iHeight, $iWidth, $iEscape = 0, $iOrientn = 0, $iWeight = $__WINAPICONSTANT_FW_NORMAL, $bItalic = False, $bUnderline = False, $bStrikeout = False, $iCharset = $__WINAPICONSTANT_DEFAULT_CHARSET, $iOutputPrec = $__WINAPICONSTANT_OUT_DEFAULT_PRECIS, $iClipPrec = $__WINAPICONSTANT_CLIP_DEFAULT_PRECIS, $iQuality = $__WINAPICONSTANT_DEFAULT_QUALITY, $iPitch = 0, $sFace = 'Arial')
Local $aResult = DllCall("gdi32.dll", "handle", "CreateFontW", "int", $iHeight, "int", $iWidth, "int", $iEscape, "int", $iOrientn, "int", $iWeight, "dword", $bItalic, "dword", $bUnderline, "dword", $bStrikeout, "dword", $iCharset, "dword", $iOutputPrec, "dword", $iClipPrec, "dword", $iQuality, "dword", $iPitch, "wstr", $sFace)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreateSolidBrush($nColor)
Local $aResult = DllCall("gdi32.dll", "handle", "CreateSolidBrush", "INT", $nColor)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreateWindowEx($iExStyle, $sClass, $sName, $iStyle, $iX, $iY, $iWidth, $iHeight, $hParent, $hMenu = 0, $hInstance = 0, $pParam = 0)
If $hInstance = 0 Then $hInstance = _WinAPI_GetModuleHandle("")
Local $aResult = DllCall("user32.dll", "hwnd", "CreateWindowExW", "dword", $iExStyle, "wstr", $sClass, "wstr", $sName, "dword", $iStyle, "int", $iX, "int", $iY, "int", $iWidth, "int", $iHeight, "hwnd", $hParent, "handle", $hMenu, "handle", $hInstance, "ptr", $pParam)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_DefWindowProc($hWnd, $iMsg, $iwParam, $ilParam)
Local $aResult = DllCall("user32.dll", "lresult", "DefWindowProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iwParam, "lparam", $ilParam)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteDC($hDC)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteDC", "handle", $hDC)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteObject($hObject)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DestroyWindow($hWnd)
Local $aResult = DllCall("user32.dll", "bool", "DestroyWindow", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DrawIconEx($hDC, $iX, $iY, $hIcon, $iWidth = 0, $iHeight = 0, $iStep = 0, $hBrush = 0, $iFlags = 3)
Local $iOptions
Switch $iFlags
Case 1
$iOptions = $DI_MASK
Case 2
$iOptions = $DI_IMAGE
Case 3
$iOptions = $DI_NORMAL
Case 4
$iOptions = $DI_COMPAT
Case 5
$iOptions = $DI_DEFAULTSIZE
Case Else
$iOptions = $DI_NOMIRROR
EndSwitch
Local $aResult = DllCall("user32.dll", "bool", "DrawIconEx", "handle", $hDC, "int", $iX, "int", $iY, "handle", $hIcon, "int", $iWidth, "int", $iHeight, "uint", $iStep, "handle", $hBrush, "uint", $iOptions)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DrawText($hDC, $sText, ByRef $tRect, $iFlags)
Local $aResult = DllCall("user32.dll", "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "struct*", $tRect, "uint", $iFlags)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_EnumWindows($bVisible = True, $hWnd = Default)
__WinAPI_EnumWindowsInit()
If $hWnd = Default Then $hWnd = _WinAPI_GetDesktopWindow()
__WinAPI_EnumWindowsChild($hWnd, $bVisible)
Return $__g_aWinList_WinAPI
EndFunc
Func __WinAPI_EnumWindowsAdd($hWnd, $sClass = "")
If $sClass = "" Then $sClass = _WinAPI_GetClassName($hWnd)
$__g_aWinList_WinAPI[0][0] += 1
Local $iCount = $__g_aWinList_WinAPI[0][0]
If $iCount >= $__g_aWinList_WinAPI[0][1] Then
ReDim $__g_aWinList_WinAPI[$iCount + 64][2]
$__g_aWinList_WinAPI[0][1] += 64
EndIf
$__g_aWinList_WinAPI[$iCount][0] = $hWnd
$__g_aWinList_WinAPI[$iCount][1] = $sClass
EndFunc
Func __WinAPI_EnumWindowsChild($hWnd, $bVisible = True)
$hWnd = _WinAPI_GetWindow($hWnd, $GW_CHILD)
While $hWnd <> 0
If(Not $bVisible) Or _WinAPI_IsWindowVisible($hWnd) Then
__WinAPI_EnumWindowsAdd($hWnd)
__WinAPI_EnumWindowsChild($hWnd, $bVisible)
EndIf
$hWnd = _WinAPI_GetWindow($hWnd, $GW_HWNDNEXT)
WEnd
EndFunc
Func __WinAPI_EnumWindowsInit()
ReDim $__g_aWinList_WinAPI[64][2]
$__g_aWinList_WinAPI[0][0] = 0
$__g_aWinList_WinAPI[0][1] = 64
EndFunc
Func _WinAPI_FillRect($hDC, $pRect, $hBrush)
Local $aResult
If IsPtr($hBrush) Then
$aResult = DllCall("user32.dll", "int", "FillRect", "handle", $hDC, "struct*", $pRect, "handle", $hBrush)
Else
$aResult = DllCall("user32.dll", "int", "FillRect", "handle", $hDC, "struct*", $pRect, "dword_ptr", $hBrush)
EndIf
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_GetClassName($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, '')
Return SetExtended($aResult[0], $aResult[2])
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
Func _WinAPI_GetDC($hWnd)
Local $aResult = DllCall("user32.dll", "handle", "GetDC", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDesktopWindow()
Local $aResult = DllCall("user32.dll", "hwnd", "GetDesktopWindow")
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetModuleHandle($sModuleName)
Local $sModuleNameType = "wstr"
If $sModuleName = "" Then
$sModuleName = 0
$sModuleNameType = "ptr"
EndIf
Local $aResult = DllCall("kernel32.dll", "handle", "GetModuleHandleW", $sModuleNameType, $sModuleName)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetMousePos($bToClient = False, $hWnd = 0)
Local $iMode = Opt("MouseCoordMode", 1)
Local $aPos = MouseGetPos()
Opt("MouseCoordMode", $iMode)
Local $tPoint = DllStructCreate($tagPOINT)
DllStructSetData($tPoint, "X", $aPos[0])
DllStructSetData($tPoint, "Y", $aPos[1])
If $bToClient And Not _WinAPI_ScreenToClient($hWnd, $tPoint) Then Return SetError(@error + 20, @extended, 0)
Return $tPoint
EndFunc
Func _WinAPI_GetParent($hWnd)
Local $aResult = DllCall("user32.dll", "hwnd", "GetParent", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetStockObject($iObject)
Local $aResult = DllCall("gdi32.dll", "handle", "GetStockObject", "int", $iObject)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindow($hWnd, $iCmd)
Local $aResult = DllCall("user32.dll", "hwnd", "GetWindow", "hwnd", $hWnd, "uint", $iCmd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowHeight($hWnd)
Local $tRect = _WinAPI_GetWindowRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRect, "Bottom") - DllStructGetData($tRect, "Top")
EndFunc
Func _WinAPI_GetWindowRect($hWnd)
Local $tRect = DllStructCreate($tagRECT)
Local $aRet = DllCall("user32.dll", "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRect)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Return $tRect
EndFunc
Func _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error, @extended, 0)
$iPID = $aResult[2]
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowWidth($hWnd)
Local $tRect = _WinAPI_GetWindowRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRect, "Right") - DllStructGetData($tRect, "Left")
EndFunc
Func _WinAPI_GUIDFromString($sGUID)
Local $tGUID = DllStructCreate($tagGUID)
_WinAPI_GUIDFromStringEx($sGUID, $tGUID)
If @error Then Return SetError(@error + 10, @extended, 0)
Return $tGUID
EndFunc
Func _WinAPI_GUIDFromStringEx($sGUID, $pGUID)
Local $aResult = DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $pGUID)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
If $hWnd = $hLastWnd Then Return True
For $iI = $__g_aInProcess_WinAPI[0][0] To 1 Step -1
If $hWnd = $__g_aInProcess_WinAPI[$iI][0] Then
If $__g_aInProcess_WinAPI[$iI][1] Then
$hLastWnd = $hWnd
Return True
Else
Return False
EndIf
EndIf
Next
Local $iPID
_WinAPI_GetWindowThreadProcessId($hWnd, $iPID)
Local $iCount = $__g_aInProcess_WinAPI[0][0] + 1
If $iCount >= 64 Then $iCount = 1
$__g_aInProcess_WinAPI[0][0] = $iCount
$__g_aInProcess_WinAPI[$iCount][0] = $hWnd
$__g_aInProcess_WinAPI[$iCount][1] =($iPID = @AutoItPID)
Return $__g_aInProcess_WinAPI[$iCount][1]
EndFunc
Func _WinAPI_IsClassName($hWnd, $sClassName)
Local $sSeparator = Opt("GUIDataSeparatorChar")
Local $aClassName = StringSplit($sClassName, $sSeparator)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $sClassCheck = _WinAPI_GetClassName($hWnd)
For $x = 1 To UBound($aClassName) - 1
If StringUpper(StringMid($sClassCheck, 1, StringLen($aClassName[$x]))) = StringUpper($aClassName[$x]) Then Return True
Next
Return False
EndFunc
Func _WinAPI_IsWindowVisible($hWnd)
Local $aResult = DllCall("user32.dll", "bool", "IsWindowVisible", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_InvalidateRect($hWnd, $tRect = 0, $bErase = True)
Local $aResult = DllCall("user32.dll", "bool", "InvalidateRect", "hwnd", $hWnd, "struct*", $tRect, "bool", $bErase)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_LoadImage($hInstance, $sImage, $iType, $iXDesired, $iYDesired, $iLoad)
Local $aResult, $sImageType = "int"
If IsString($sImage) Then $sImageType = "wstr"
$aResult = DllCall("user32.dll", "handle", "LoadImageW", "handle", $hInstance, $sImageType, $sImage, "uint", $iType, "int", $iXDesired, "int", $iYDesired, "uint", $iLoad)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_MoveWindow($hWnd, $iX, $iY, $iWidth, $iHeight, $bRepaint = True)
Local $aResult = DllCall("user32.dll", "bool", "MoveWindow", "hwnd", $hWnd, "int", $iX, "int", $iY, "int", $iWidth, "int", $iHeight, "bool", $bRepaint)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_PostMessage($hWnd, $iMsg, $iwParam, $ilParam)
Local $aResult = DllCall("user32.dll", "bool", "PostMessage", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iwParam, "lparam", $ilParam)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_ReleaseDC($hWnd, $hDC)
Local $aResult = DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_ScreenToClient($hWnd, ByRef $tPoint)
Local $aResult = DllCall("user32.dll", "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tPoint)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_SelectObject($hDC, $hGDIObj)
Local $aResult = DllCall("gdi32.dll", "handle", "SelectObject", "handle", $hDC, "handle", $hGDIObj)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_SetBkColor($hDC, $iColor)
Local $aResult = DllCall("gdi32.dll", "INT", "SetBkColor", "handle", $hDC, "INT", $iColor)
If @error Then Return SetError(@error, @extended, -1)
Return $aResult[0]
EndFunc
Func _WinAPI_SetBkMode($hDC, $iBkMode)
Local $aResult = DllCall("gdi32.dll", "int", "SetBkMode", "handle", $hDC, "int", $iBkMode)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_SetTextColor($hDC, $iColor)
Local $aResult = DllCall("gdi32.dll", "INT", "SetTextColor", "handle", $hDC, "INT", $iColor)
If @error Then Return SetError(@error, @extended, -1)
Return $aResult[0]
EndFunc
Func _WinAPI_SetWindowsHookEx($idHook, $pFn, $hMod, $iThreadId = 0)
Local $aResult = DllCall("user32.dll", "handle", "SetWindowsHookEx", "int", $idHook, "ptr", $pFn, "handle", $hMod, "dword", $iThreadId)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_StringFromGUID($pGUID)
Local $aResult = DllCall("ole32.dll", "int", "StringFromGUID2", "struct*", $pGUID, "wstr", "", "int", 40)
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, "")
Return SetExtended($aResult[0], $aResult[2])
EndFunc
Func _WinAPI_UnhookWindowsHookEx($hHk)
Local $aResult = DllCall("user32.dll", "bool", "UnhookWindowsHookEx", "handle", $hHk)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_UpdateLayeredWindow($hWnd, $hDCDest, $pPTDest, $pSize, $hDCSrce, $pPTSrce, $iRGB, $pBlend, $iFlags)
Local $aResult = DllCall("user32.dll", "bool", "UpdateLayeredWindow", "hwnd", $hWnd, "handle", $hDCDest, "ptr", $pPTDest, "ptr", $pSize, "handle", $hDCSrce, "ptr", $pPTSrce, "dword", $iRGB, "ptr", $pBlend, "dword", $iFlags)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_WideCharToMultiByte($pUnicode, $iCodePage = 0, $bRetString = True)
Local $sUnicodeType = "wstr"
If Not IsString($pUnicode) Then $sUnicodeType = "struct*"
Local $aResult = DllCall("kernel32.dll", "int", "WideCharToMultiByte", "uint", $iCodePage, "dword", 0, $sUnicodeType, $pUnicode, "int", -1, "ptr", 0, "int", 0, "ptr", 0, "ptr", 0)
If @error Or Not $aResult[0] Then Return SetError(@error + 20, @extended, "")
Local $tMultiByte = DllStructCreate("char[" & $aResult[0] & "]")
$aResult = DllCall("kernel32.dll", "int", "WideCharToMultiByte", "uint", $iCodePage, "dword", 0, $sUnicodeType, $pUnicode, "int", -1, "struct*", $tMultiByte, "int", $aResult[0], "ptr", 0, "ptr", 0)
If @error Or Not $aResult[0] Then Return SetError(@error + 10, @extended, "")
If $bRetString Then Return DllStructGetData($tMultiByte, 1)
Return $tMultiByte
EndFunc
Func _WinAPI_WindowFromPoint(ByRef $tPoint)
Local $aResult = DllCall("user32.dll", "hwnd", "WindowFromPoint", "struct", $tPoint)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Global Const $GDIP_PXF32ARGB = 0x0026200A
Global Const $GDIP_SMOOTHINGMODE_DEFAULT = 0
Global Const $GDIP_SMOOTHINGMODE_HIGHQUALITY = 2
Global Const $GDIP_SMOOTHINGMODE_ANTIALIAS8X8 = 5
Global Const $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC = 7
Global Const $CA_NEGATIVE = 0x01
Global Const $CA_LOG_FILTER = 0x02
Global Const $COLORONCOLOR = 3
Global Const $HALFTONE = 4
Global $__g_hHeap = 0, $__g_iRGBMode = 1
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func _WinAPI_SwitchColor($iColor)
If $iColor = -1 Then Return $iColor
Return BitOR(BitAND($iColor, 0x00FF00), BitShift(BitAND($iColor, 0x0000FF), -16), BitShift(BitAND($iColor, 0xFF0000), 16))
EndFunc
Func __Iif($bTest, $vTrue, $vFalse)
Return $bTest ? $vTrue : $vFalse
EndFunc
Func __RGB($iColor)
If $__g_iRGBMode Then
$iColor = _WinAPI_SwitchColor($iColor)
EndIf
Return $iColor
EndFunc
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Global Const $tagBITMAP = 'struct;long bmType;long bmWidth;long bmHeight;long bmWidthBytes;ushort bmPlanes;ushort bmBitsPixel;ptr bmBits;endstruct'
Global Const $tagCOLORADJUSTMENT = 'ushort Size;ushort Flags;ushort IlluminantIndex;ushort RedGamma;ushort GreenGamma;ushort BlueGamma;ushort ReferenceBlack;ushort ReferenceWhite;short Contrast;short Brightness;short Colorfulness;short RedGreenTint'
Func _WinAPI_AdjustBitmap($hBitmap, $iWidth, $iHeight, $iMode = 3, $tAdjustment = 0)
Local $tObj = DllStructCreate($tagBITMAP)
Local $aRet = DllCall('gdi32.dll', 'int', 'GetObject', 'handle', $hBitmap, 'int', DllStructGetSize($tObj), 'struct*', $tObj)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
If $iWidth = -1 Then
$iWidth = DllStructGetData($tObj, 'bmWidth')
EndIf
If $iHeight = -1 Then
$iHeight = DllStructGetData($tObj, 'bmHeight')
EndIf
$aRet = DllCall('user32.dll', 'handle', 'GetDC', 'hwnd', 0)
Local $hDC = $aRet[0]
$aRet = DllCall('gdi32.dll', 'handle', 'CreateCompatibleDC', 'handle', $hDC)
Local $hDestDC = $aRet[0]
$aRet = DllCall('gdi32.dll', 'handle', 'CreateCompatibleBitmap', 'handle', $hDC, 'int', $iWidth, 'int', $iHeight)
Local $hBmp = $aRet[0]
$aRet = DllCall('gdi32.dll', 'handle', 'SelectObject', 'handle', $hDestDC, 'handle', $hBmp)
Local $hDestSv = $aRet[0]
$aRet = DllCall('gdi32.dll', 'handle', 'CreateCompatibleDC', 'handle', $hDC)
Local $hSrcDC = $aRet[0]
$aRet = DllCall('gdi32.dll', 'handle', 'SelectObject', 'handle', $hSrcDC, 'handle', $hBitmap)
Local $hSrcSv = $aRet[0]
If _WinAPI_SetStretchBltMode($hDestDC, $iMode) Then
Switch $iMode
Case 4
If IsDllStruct($tAdjustment) Then
If Not _WinAPI_SetColorAdjustment($hDestDC, $tAdjustment) Then
EndIf
EndIf
Case Else
EndSwitch
EndIf
$aRet = _WinAPI_StretchBlt($hDestDC, 0, 0, $iWidth, $iHeight, $hSrcDC, 0, 0, DllStructGetData($tObj, 'bmWidth'), DllStructGetData($tObj, 'bmHeight'), 0x00CC0020)
DllCall('user32.dll', 'int', 'ReleaseDC', 'hwnd', 0, 'handle', $hDC)
DllCall('gdi32.dll', 'handle', 'SelectObject', 'handle', $hDestDC, 'handle', $hDestSv)
DllCall('gdi32.dll', 'handle', 'SelectObject', 'handle', $hSrcDC, 'handle', $hSrcSv)
DllCall('gdi32.dll', 'bool', 'DeleteDC', 'handle', $hDestDC)
DllCall('gdi32.dll', 'bool', 'DeleteDC', 'handle', $hSrcDC)
If Not $aRet Then Return SetError(10, 0, 0)
Return $hBmp
EndFunc
Func _WinAPI_ColorAdjustLuma($iRGB, $iPercent, $bScale = True)
If $iRGB = -1 Then Return SetError(10, 0, -1)
If $bScale Then
$iPercent = Floor($iPercent * 10)
EndIf
Local $aRet = DllCall('shlwapi.dll', 'dword', 'ColorAdjustLuma', 'dword', __RGB($iRGB), 'int', $iPercent, 'bool', $bScale)
If @error Then Return SetError(@error, @extended, -1)
Return __RGB($aRet[0])
EndFunc
Func _WinAPI_CreateColorAdjustment($iFlags = 0, $iIlluminant = 0, $iGammaR = 10000, $iGammaG = 10000, $iGammaB = 10000, $iBlack = 0, $iWhite = 10000, $iContrast = 0, $iBrightness = 0, $iColorfulness = 0, $iTint = 0)
Local $tCA = DllStructCreate($tagCOLORADJUSTMENT)
DllStructSetData($tCA, 1, DllStructGetSize($tCA))
DllStructSetData($tCA, 2, $iFlags)
DllStructSetData($tCA, 3, $iIlluminant)
DllStructSetData($tCA, 4, $iGammaR)
DllStructSetData($tCA, 5, $iGammaG)
DllStructSetData($tCA, 6, $iGammaB)
DllStructSetData($tCA, 7, $iBlack)
DllStructSetData($tCA, 8, $iWhite)
DllStructSetData($tCA, 9, $iContrast)
DllStructSetData($tCA, 10, $iBrightness)
DllStructSetData($tCA, 11, $iColorfulness)
DllStructSetData($tCA, 12, $iTint)
Return $tCA
EndFunc
Func _WinAPI_SetColorAdjustment($hDC, $tAdjustment)
Local $aRet = DllCall('gdi32.dll', 'bool', 'SetColorAdjustment', 'handle', $hDC, 'struct*', $tAdjustment)
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_SetStretchBltMode($hDC, $iMode)
Local $aRet = DllCall('gdi32.dll', 'int', 'SetStretchBltMode', 'handle', $hDC, 'int', $iMode)
If @error Or Not $aRet[0] Or($aRet[0] = 87) Then Return SetError(@error + 10, $aRet[0], 0)
Return $aRet[0]
EndFunc
Func _WinAPI_StretchBlt($hDestDC, $iXDest, $iYDest, $iWidthDest, $iHeightDest, $hSrcDC, $iXSrc, $iYSrc, $iWidthSrc, $iHeightSrc, $iRop)
Local $aRet = DllCall('gdi32.dll', 'bool', 'StretchBlt', 'handle', $hDestDC, 'int', $iXDest, 'int', $iYDest, 'int', $iWidthDest, 'int', $iHeightDest, 'hwnd', $hSrcDC, 'int', $iXSrc, 'int', $iYSrc, 'int', $iWidthSrc, 'int', $iHeightSrc, 'dword', $iRop)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
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
Func _GDIPlus_BitmapCreateFromHBITMAP($hBmp, $hPal = 0)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateBitmapFromHBITMAP", "handle", $hBmp, "handle", $hPal, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[3]
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
Func _GDIPlus_Encoders()
Local $iCount = _GDIPlus_EncodersGetCount()
Local $iSize = _GDIPlus_EncodersGetSize()
Local $tBuffer = DllStructCreate("byte[" & $iSize & "]")
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncoders", "uint", $iCount, "uint", $iSize, "struct*", $tBuffer)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tCodec, $aInfo[$iCount + 1][14]
$aInfo[0][0] = $iCount
For $iI = 1 To $iCount
$tCodec = DllStructCreate($tagGDIPIMAGECODECINFO, $pBuffer)
$aInfo[$iI][1] = _WinAPI_StringFromGUID(DllStructGetPtr($tCodec, "CLSID"))
$aInfo[$iI][2] = _WinAPI_StringFromGUID(DllStructGetPtr($tCodec, "FormatID"))
$aInfo[$iI][3] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "CodecName"))
$aInfo[$iI][4] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "DllName"))
$aInfo[$iI][5] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "FormatDesc"))
$aInfo[$iI][6] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "FileExt"))
$aInfo[$iI][7] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "MimeType"))
$aInfo[$iI][8] = DllStructGetData($tCodec, "Flags")
$aInfo[$iI][9] = DllStructGetData($tCodec, "Version")
$aInfo[$iI][10] = DllStructGetData($tCodec, "SigCount")
$aInfo[$iI][11] = DllStructGetData($tCodec, "SigSize")
$aInfo[$iI][12] = DllStructGetData($tCodec, "SigPattern")
$aInfo[$iI][13] = DllStructGetData($tCodec, "SigMask")
$pBuffer += DllStructGetSize($tCodec)
Next
Return $aInfo
EndFunc
Func _GDIPlus_EncodersGetCLSID($sFileExt)
Local $aEncoders = _GDIPlus_Encoders()
If @error Then Return SetError(@error, 0, "")
For $iI = 1 To $aEncoders[0][0]
If StringInStr($aEncoders[$iI][6], "*." & $sFileExt) > 0 Then Return $aEncoders[$iI][1]
Next
Return SetError(-1, -1, "")
EndFunc
Func _GDIPlus_EncodersGetCount()
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncodersSize", "uint*", 0, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[1]
EndFunc
Func _GDIPlus_EncodersGetSize()
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncodersSize", "uint*", 0, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
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
Func _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $iInterpolationMode)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSetInterpolationMode", "handle", $hGraphics, "int", $iInterpolationMode)
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
Func _GDIPlus_ImageSaveToFile($hImage, $sFileName)
Local $sExt = __GDIPlus_ExtractFileExt($sFileName)
Local $sCLSID = _GDIPlus_EncodersGetCLSID($sExt)
If $sCLSID = "" Then Return SetError(-1, 0, False)
Local $bRet = _GDIPlus_ImageSaveToFileEx($hImage, $sFileName, $sCLSID, 0)
Return SetError(@error, @extended, $bRet)
EndFunc
Func _GDIPlus_ImageSaveToFileEx($hImage, $sFileName, $sEncoder, $pParams = 0)
Local $tGUID = _WinAPI_GUIDFromString($sEncoder)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSaveImageToFile", "handle", $hImage, "wstr", $sFileName, "struct*", $tGUID, "struct*", $pParams)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_ImageResize($hImage, $iNewWidth, $iNewHeight, $iInterpolationMode = $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC)
Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iNewWidth, $iNewHeight)
If @error Then Return SetError(1, 0, 0)
Local $hBmpCtxt = _GDIPlus_ImageGetGraphicsContext($hBitmap)
If @error Then
_GDIPlus_BitmapDispose($hBitmap)
Return SetError(2, @extended, 0)
EndIf
_GDIPlus_GraphicsSetInterpolationMode($hBmpCtxt, $iInterpolationMode)
If @error Then
_GDIPlus_GraphicsDispose($hBmpCtxt)
_GDIPlus_BitmapDispose($hBitmap)
Return SetError(3, @extended, 0)
EndIf
_GDIPlus_GraphicsDrawImageRect($hBmpCtxt, $hImage, 0, 0, $iNewWidth, $iNewHeight)
If @error Then
_GDIPlus_GraphicsDispose($hBmpCtxt)
_GDIPlus_BitmapDispose($hBitmap)
Return SetError(4, @extended, 0)
EndIf
_GDIPlus_GraphicsDispose($hBmpCtxt)
Return $hBitmap
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
Func __GDIPlus_ExtractFileExt($sFileName, $bNoDot = True)
Local $iIndex = __GDIPlus_LastDelimiter(".\:", $sFileName)
If($iIndex > 0) And(StringMid($sFileName, $iIndex, 1) = '.') Then
If $bNoDot Then
Return StringMid($sFileName, $iIndex + 1)
Else
Return StringMid($sFileName, $iIndex)
EndIf
Else
Return ""
EndIf
EndFunc
Func __GDIPlus_LastDelimiter($sDelimiters, $sString)
Local $sDelimiter, $iN
For $iI = 1 To StringLen($sDelimiters)
$sDelimiter = StringMid($sDelimiters, $iI, 1)
$iN = StringInStr($sString, $sDelimiter, 0, -1)
If $iN > 0 Then Return $iN
Next
EndFunc
Global Const $WS_MAXIMIZEBOX = 0x00010000
Global Const $WS_MINIMIZEBOX = 0x00020000
Global Const $WS_SIZEBOX = 0x00040000
Global Const $WS_SYSMENU = 0x00080000
Global Const $WS_CAPTION = 0x00C00000
Global Const $WS_POPUP = 0x80000000
Global Const $WS_EX_APPWINDOW = 0x00040000
Global Const $WS_EX_LAYERED = 0x00080000
Global Const $WS_EX_MDICHILD = 0x00000040
Global Const $WM_DESTROY = 0x0002
Global Const $WM_MOVE = 0x0003
Global Const $WM_PAINT = 0x000F
Global Const $WM_CLOSE = 0x0010
Global Const $WM_COMMAND = 0x0111
Global Const $WM_HSCROLL = 0x0114
Global Const $WM_ENTERSIZEMOVE = 0x0231
Global Const $WM_EXITSIZEMOVE = 0x0232
Global Const $WM_MOUSEMOVE = 0x0200
Global Const $WM_LBUTTONDOWN = 0x0201
Global Const $WM_LBUTTONUP = 0x0202
Global Const $TRANSPARENT = 1
Global Const $DT_CENTER = 0x1
Global Const $DT_HIDEPREFIX = 0x100000
Global Const $DT_LEFT = 0x0
Global Const $DT_SINGLELINE = 0x20
Global Const $DT_TOP = 0x0
Global Const $DT_VCENTER = 0x4
Global Const $GUI_SS_DEFAULT_GUI = BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU)
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_EVENT_RESTORE = -5
Global Const $GUI_EVENT_MAXIMIZE = -6
Global Const $GUI_EVENT_PRIMARYUP = -8
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_CHECKED = 1
Global Const $GUI_UNCHECKED = 4
Global Const $GUI_FOCUS = 256
Global Const $_UDF_GlobalIDs_OFFSET = 2
Global Const $_UDF_GlobalID_MAX_WIN = 16
Global Const $_UDF_STARTID = 10000
Global Const $_UDF_GlobalID_MAX_IDS = 55535
Global Const $__UDFGUICONSTANT_WS_VISIBLE = 0x10000000
Global Const $__UDFGUICONSTANT_WS_CHILD = 0x40000000
Global $__g_aUDF_GlobalIDs_Used[$_UDF_GlobalID_MAX_WIN][$_UDF_GlobalID_MAX_IDS + $_UDF_GlobalIDs_OFFSET + 1]
Func __UDF_GetNextGlobalID($hWnd)
Local $nCtrlID, $iUsedIndex = -1, $bAllUsed = True
If Not WinExists($hWnd) Then Return SetError(-1, -1, 0)
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] <> 0 Then
If Not WinExists($__g_aUDF_GlobalIDs_Used[$iIndex][0]) Then
For $x = 0 To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
$__g_aUDF_GlobalIDs_Used[$iIndex][$x] = 0
Next
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
EndIf
EndIf
Next
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd Then
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
If $iUsedIndex = -1 Then
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = 0 Then
$__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
EndIf
If $iUsedIndex = -1 And $bAllUsed Then Return SetError(16, 0, 0)
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] = $_UDF_STARTID + $_UDF_GlobalID_MAX_IDS Then
For $iIDIndex = $_UDF_GlobalIDs_OFFSET To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = 0 Then
$nCtrlID =($iIDIndex - $_UDF_GlobalIDs_OFFSET) + 10000
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = $nCtrlID
Return $nCtrlID
EndIf
Next
Return SetError(-1, $_UDF_GlobalID_MAX_IDS, 0)
EndIf
$nCtrlID = $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1]
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] += 1
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][($nCtrlID - 10000) + $_UDF_GlobalIDs_OFFSET] = $nCtrlID
Return $nCtrlID
EndFunc
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $DEFAULT_CHARSET = 1
Global Const $OUT_DEFAULT_PRECIS = 0
Global Const $CLIP_DEFAULT_PRECIS = 0
Global Const $PROOF_QUALITY = 2
Global Const $__METROBUTTON_CONSTANT_ClassName = "MetroButton"
Global Const $__METROBUTTON_CONSTANT_SIZEOFPTR = DllStructGetSize(DllStructCreate("ptr"))
Global Const $WM_MOUSEINSIDE = 27081997
Global Const $WM_MOUSEOUTSIDE = 27081998
Global Const $__METROBUTTON_CONSTANT_tagWNDCLASSEX = "uint Size; " & "uint style; " & "ptr WndProc; " & "int ClsExtra; " & "int WndExtra; " & "HANDLE Instance; " & "HANDLE Icon; " & "HANDLE Cursor; " & "HANDLE Background; " & "ptr MenuName; " & "ptr ClassName; " & "HANDLE IconSm"
Global Const $__ARR_BUTTON_COLOR[3][2] = [ [0xFFFFFF, 0x171717], [0xDBDBDB, 0x171717], [0xB3B3B3, 0x171717] ]
Global $xBackup = 0, $iDeltaMoveButton = 1, $hButton_Moved = ""
Global Const $WM_NEXTANIMATION = 27081999
Global Const $WM_ENDANIMATION = 27082000
Global $__STOP_ANIMATION = False
Global $__IsAnimation = False
Global Const $MOUSE_BUTTON_CALLBACK = DllCallbackRegister("__MOUSE_BUTTON_PROC", "long", "int;wparam;lparam")
Global Const $H_BUTTON_MOD = _WinAPI_GetModuleHandle(0)
Global Const $H_BUTTON_HOOK = _WinAPI_SetWindowsHookEx($WH_MOUSE_LL, DllCallbackGetPtr($MOUSE_BUTTON_CALLBACK), $H_BUTTON_MOD)
Func __MOUSE_BUTTON_PROC($nCode, $wParam, $lParam)
If $__IsAnimation Then Return
Local Static $LAST_WINDOWS = ""
Local $tPoint = _WinAPI_GetMousePos()
Local $PRESENT_WINDOWS = _WinAPI_WindowFromPoint($tPoint)
Switch $wParam
Case $WM_MOUSEMOVE
If $PRESENT_WINDOWS <> $LAST_WINDOWS Then
If _WinAPI_IsClassName($PRESENT_WINDOWS, $__METROBUTTON_CONSTANT_ClassName) Then
_WinAPI_PostMessage($PRESENT_WINDOWS, $WM_COMMAND, $WM_MOUSEINSIDE, $__METROBUTTON_CONSTANT_ClassName)
EndIf
If _WinAPI_IsClassName($LAST_WINDOWS, $__METROBUTTON_CONSTANT_ClassName) Then
_WinAPI_PostMessage($LAST_WINDOWS, $WM_COMMAND, $WM_MOUSEOUTSIDE, $__METROBUTTON_CONSTANT_ClassName)
EndIf
$LAST_WINDOWS = $PRESENT_WINDOWS
EndIf
EndSwitch
Return _WinAPI_CallNextHookEx($H_BUTTON_HOOK, $nCode, $wParam, $lParam)
EndFunc
Global $__sBUTTON_TEXT = ""
Global $__sBUTTON_CANNOTUPDATE = ""
Global $__bBUTTON_ALLCANUPDATE = True
Func _GUICtrlMetroButton_Create($hWnd, $sText, $iX, $iY, $iWidth = 60, $iHeight = 40, $sIcon = "")
Local $nCtrlID
Global $H_BUTTON_CALLBACK
Local Static $iAtom = 0
If Not IsHWnd($hWnd) Or Not WinExists($hWnd) Then Return SetError(1, 0, 0)
If Not $iAtom Then
Local $tClassName, $tWC, $aRes
If Not OnAutoItExitRegister("__GUICtrlMetroButton_OnExit") Then Return SetError(2, 0, 0)
$H_BUTTON_CALLBACK = DllCallbackRegister("__GUICtrlMetroButton_WndProc", "lresult", "hwnd;uint;wparam;lparam")
If @error Then Return SetError(3, 0, 0)
$tClassName = DllStructCreate("wchar[" &(StringLen($__METROBUTTON_CONSTANT_ClassName) + 1) & "]")
DllStructSetData($tClassName, 1, $__METROBUTTON_CONSTANT_ClassName)
$tWC = DllStructCreate($__METROBUTTON_CONSTANT_tagWNDCLASSEX)
DllStructSetData($tWC, "Size", DllStructGetSize($tWC))
DllStructSetData($tWC, "style", 3)
DllStructSetData($tWC, "WndProc", DllCallbackGetPtr($H_BUTTON_CALLBACK))
DllStructSetData($tWC, "ClsExtra", 0)
DllStructSetData($tWC, "WndExtra", 16 *(@AutoItX64 + 1))
DllStructSetData($tWC, "Instance", _WinAPI_GetModuleHandle(0))
DllStructSetData($tWC, "Icon", 0)
DllStructSetData($tWC, "Cursor", _WinAPI_LoadImage(0, 32512, $IMAGE_CURSOR, 0, 0, BitOR($LR_DEFAULTCOLOR, $LR_SHARED)))
DllStructSetData($tWC, "Background", 0)
DllStructSetData($tWC, "MenuName", 0)
DllStructSetData($tWC, "ClassName", DllStructGetPtr($tClassName))
DllStructSetData($tWC, "IconSm", 0)
$aRes = DllCall("user32.dll", "word", "RegisterClassExW", "ptr", DllStructGetPtr($tWC))
If @error Or Not $aRes[0] Then Return SetError(4, @error, 0)
$iAtom = $aRes[0]
EndIf
$nCtrlID = __UDF_GetNextGlobalID($hWnd)
If @error Then Return SetError(5, @error, 0)
$hButton = _WinAPI_CreateWindowEx(0, $__METROBUTTON_CONSTANT_ClassName, "", BitOR($__UDFGUICONSTANT_WS_CHILD, $__UDFGUICONSTANT_WS_VISIBLE), $iX, $iY, $iWidth, $iHeight, $hWnd, $nCtrlID)
If Not $hButton Then Return SetError(6, @error, 0)
$__sBUTTON_TEXT &= "|" & $hButton & '="' & $sText & '"|'
If $sIcon <> "" And FileExists($sIcon) Then
Local $hIcon = _WinAPI_LoadImage(0, $sIcon, $IMAGE_ICON, 32, 32, BitOR($LR_DEFAULTCOLOR, $LR_LOADFROMFILE, $LR_LOADTRANSPARENT))
$__sBUTTON_TEXT &= "|" & $hButton & '=="' & $hIcon & '"|'
EndIf
Return $hButton
EndFunc
Func __GUICtrlMetroButton_WndProc($hWnd, $iMsg, $wParam, $lParam)
If $__bBUTTON_ALLCANUPDATE Then
Switch $iMsg
Case $WM_COMMAND
Switch $wParam
Case $WM_MOUSEINSIDE
_SetBkButton($hWnd, 1)
Case $WM_MOUSEOUTSIDE
_SetBkButton($hWnd, 0)
Case $WM_NEXTANIMATION
If Not $__STOP_ANIMATION Then
_GUICtrlMetroButton_Animation($hWnd)
Else
$__STOP_ANIMATION = False
EndIf
Case $WM_ENDANIMATION
$__STOP_ANIMATION = True
EndSwitch
Case $WM_LBUTTONDOWN
_SetBkButton($hWnd, 2)
Case $WM_LBUTTONUP
_SetBkButton($hWnd, 1)
ConsoleWrite("WM_LBUTTONUP Handle:" & $hWnd & @CRLF)
_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_LBUTTONUP, $hWnd)
Case $WM_PAINT, $WM_MOVE
If Not StringInStr($__sBUTTON_CANNOTUPDATE, $hWnd) Then
_SetBkButton($hWnd, 0)
EndIf
Case $WM_DESTROY
_WinAPI_DestroyWindow($hWnd)
EndSwitch
EndIf
Return _WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc
Func _GUICtrlMetroButton_Animation($hWnd)
$__IsAnimation = True
__WinMove($hWnd, _GUICtrlMetroButton_GetX(_WinAPI_GetParent($hWnd), $hWnd), _GUICtrlMetroButton_GetY(_WinAPI_GetParent($hWnd), $hWnd), _WinAPI_GetWindowWidth($hWnd), _WinAPI_GetWindowHeight($hWnd))
$__IsAnimation = False
EndFunc
Func __WinMove($hWnd, $x, $y, $w, $h)
Local $iOldDelta = 1
If Not StringInStr($hButton_Moved, $hWnd) Then
_WinAPI_MoveWindow($hWnd, $x-$iDeltaMoveButton, $y, $w, $h, False)
If Abs($x) >= $w Then
$hButton_Moved &= $hWnd
$iDeltaMoveButton = 1
_WinAPI_MoveWindow($hWnd, $xBackup-$w, $y, $w, $h)
_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_NEXTANIMATION, $hWnd)
Return
EndIf
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
ElseIf StringInStr($hButton_Moved, $hWnd) Then
_WinAPI_MoveWindow($hWnd, $x+$iDeltaMoveButton, $y, $w, $h, False)
If $x >= $xBackup Then
$hButton_Moved = StringReplace($hButton_Moved, $hWnd, "")
$iDeltaMoveButton = 1
_WinAPI_MoveWindow($hWnd, $xBackup, $y, $w, $h)
_WinAPI_PostMessage(_WinAPI_GetParent($hWnd), $WM_COMMAND, $WM_NEXTANIMATION, $hWnd)
Return
EndIf
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
Func _GUICtrlMetroButton_GetX($hGUI, $hWnd)
Local $tRect = _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
Return DllStructGetData($tRect, "left")
EndFunc
Func _GUICtrlMetroButton_GetY($hGUI, $hWnd)
Local $tRect = _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
Return DllStructGetData($tRect, "top")
EndFunc
Func _GUICtrlMetroButton_GetRect($hGUI, $hWnd)
Local $tPoint = DllStructCreate("int X;int Y")
Local $tRect = _WinAPI_GetWindowRect($hWnd)
DllStructSetData($tPoint, "X", DllStructGetData($tRect, "left"))
DllStructSetData($tPoint, "Y", DllStructGetData($tRect, "top"))
_WinAPI_ScreenToClient($hGUI,$tPoint)
DllStructSetData($tRect, "left",DllStructGetData($tPoint, "X"))
DllStructSetData($tRect, "top",DllStructGetData($tPoint, "Y"))
Return $tRect
EndFunc
Func __GUICtrlMetroButton_OnExit()
Local $a = _WinAPI_EnumWindows()
For $i = 1 To $a[0][0]
If $a[$i][1] = $__METROBUTTON_CONSTANT_ClassName And WinGetProcess($a[$i][0]) = @AutoItPID Then
_WinAPI_PostMessage($a[$i][0], $WM_CLOSE, 0, 0)
EndIf
Next
DllCallbackFree($H_BUTTON_CALLBACK)
DllCallbackFree($MOUSE_BUTTON_CALLBACK)
_WinAPI_UnhookWindowsHookEx($H_BUTTON_HOOK)
DllCall("user32.dll", "int", "UnregisterClassW", "wstr", $__METROBUTTON_CONSTANT_ClassName, "HANDLE", _WinAPI_GetModuleHandle(0))
EndFunc
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
Local $hIcon = StringRegExp($__sBUTTON_TEXT, "\|" & $hWnd & '=="(.+?)"\|', 3)
Local $iStart, $iEnd, $iStep, $iColor
$hBrushBorder = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($__ARR_BUTTON_COLOR[$idArrColor][1]))
$__hBUTTON_FONT_Caption = _WinAPI_CreateFont(18, 0, 0, 0, 400, False, False, False, $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $PROOF_QUALITY, 0, 'Segoe UI')
$hOldFont = _WinAPI_SelectObject($hDC, $__hBUTTON_FONT_Caption)
_WinAPI_SetTextColor($hDC, $__ARR_BUTTON_COLOR[$idArrColor][1])
If $idArrColor = 1 And $idLastArrColor = 0 Then
$iStart = 0
$iEnd = 99
$iStep = $iEnd*20
Else
$iStart = 0
$iEnd = 0
$iStep = 1
EndIf
Local $iStyle = ""
For $i = $iStart To $iEnd Step $iStep
If $__ARR_BUTTON_COLOR[$idArrColor][0] <> "" Then
$iColor = _WinAPI_ColorAdjustLuma($__ARR_BUTTON_COLOR[$idArrColor][0], $i, True)
$hBrushBackground = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($iColor))
_WinAPI_FillRect($hDC, $pRect, $hBrushBackground)
_WinAPI_SetBkColor($hDC, _WinAPI_SwitchColor($iColor))
_WinAPI_DeleteObject($hBrushBackground)
ElseIf $__ARR_BUTTON_COLOR[$idArrColor][0] = "" Then
_WinAPI_SetStretchBltMode($hDC, $COLORONCOLOR)
_WinAPI_SetBkMode($hDC, $TRANSPARENT)
EndIf
If IsArray($hIcon) And StringLen($hIcon[0]) > 0 Then
_WinAPI_DrawIconEx($hDC, 15, DllStructGetData($tRect, 4)/2-32/2, $hIcon[0])
$iStyle = BitOR($DT_LEFT, $DT_VCENTER, $DT_SINGLELINE, $DT_TOP, $DT_HIDEPREFIX)
DllStructSetData($tRect, 1, 10+40)
Else
$iStyle = BitOR($DT_CENTER, $DT_SINGLELINE, $DT_TOP, $DT_HIDEPREFIX, $DT_VCENTER)
EndIf
_WinAPI_DrawText($hDC, $sText, $tRect, $iStyle)
Next
$idLastArrColor = $idArrColor
_WinAPI_DeleteObject($hOldFont)
_WinAPI_DeleteObject($hBrushBorder)
_WinAPI_DeleteObject($__hBUTTON_FONT_Caption)
_WinAPI_ReleaseDC($hWnd, $hDC)
EndFunc
Global Const $STM_SETIMAGE = 0x0172
Global Const $STM_GETIMAGE = 0x0173
Func _SetBitmapAdjust($hWnd, $hBitmap, $aAdjust = 0)
If Not IsHWnd($hWnd) Then
$hWnd = GUICtrlGetHandle($hWnd)
If Not $hWnd Then
Return 0
EndIf
EndIf
Local $tAdjust = 0
If IsArray($aAdjust) Then
$tAdjust = _WinAPI_CreateColorAdjustment($aAdjust[9], $aAdjust[10], $aAdjust[0], $aAdjust[1], $aAdjust[2], $aAdjust[3], $aAdjust[4], $aAdjust[5], $aAdjust[6], $aAdjust[7], $aAdjust[8])
EndIf
$hBitmap = _WinAPI_AdjustBitmap($hBitmap, -1, -1, $HALFTONE, $tAdjust)
If @error Then
Return 0
EndIf
Local $hPrev = _SendMessage($hWnd, $STM_SETIMAGE, $IMAGE_BITMAP, $hBitmap)
If $hPrev Then
_WinAPI_DeleteObject($hPrev)
EndIf
$hPrev = _SendMessage($hWnd, $STM_GETIMAGE)
If $hPrev <> $hBitmap Then
_WinAPI_DeleteObject($hBitmap)
EndIf
Return 1
EndFunc
Func _WinAPI_ConvertPointClientToScreen($hWnd, $iX, $iY)
Local $tPoint = DllStructCreate("int X;int Y")
Local $arrPoint[2]
DllStructSetData($tPoint, "X", $iX)
DllStructSetData($tPoint, "Y", $iY)
_WinAPI_ClientToScreen($hWnd,$tPoint)
$arrPoint[0] = DllStructGetData($tPoint, "X")
$arrPoint[1] = DllStructGetData($tPoint, "Y")
Return $arrPoint
EndFunc
Func SetBitmap($hWnd, $hImage, $iOpacity)
Local $hScrDC, $hMemDC, $hBitmap, $hOld, $pSize, $tSize, $pSource, $tSource, $pBlend, $tBlend
$hScrDC = _WinAPI_GetDC(0)
$hMemDC = _WinAPI_CreateCompatibleDC($hScrDC)
$hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImage)
$hOld = _WinAPI_SelectObject($hMemDC, $hBitmap)
$tSize = DllStructCreate($tagSIZE)
$pSize = DllStructGetPtr($tSize)
DllStructSetData($tSize, "X", _GDIPlus_ImageGetWidth($hImage))
DllStructSetData($tSize, "Y", _GDIPlus_ImageGetHeight($hImage))
$tSource = DllStructCreate($tagPOINT)
$pSource = DllStructGetPtr($tSource)
$tBlend = DllStructCreate($tagBLENDFUNCTION)
$pBlend = DllStructGetPtr($tBlend)
DllStructSetData($tBlend, "Alpha", $iOpacity)
DllStructSetData($tBlend, "Format", 1)
_WinAPI_UpdateLayeredWindow($hWnd, $hScrDC, 0, $pSize, $hMemDC, $pSource, 0, $pBlend, $ULW_ALPHA)
_WinAPI_ReleaseDC(0, $hScrDC)
_WinAPI_SelectObject($hMemDC, $hOld)
_WinAPI_DeleteObject($hBitmap)
_WinAPI_DeleteDC($hMemDC)
EndFunc
Func _GetPathWallpaper()
Local $data
If @OSVersion = "WIN_7" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_XP" Or @OSVersion = "WIN_XPe" Then
$data = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\General", "WallpaperSource")
Else
$data = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache")
$data = BinaryToString(StringReplace("0x" & StringTrimLeft($data, 50), "00", ""))
EndIf
If FileExists($data) Then Return $data
Return ""
EndFunc
Global Const $SS_RIGHT = 0x2
Global Const $SS_NOPREFIX = 0x0080
Global Const $SS_NOTIFY = 0x0100
Global Const $SS_CENTERIMAGE = 0x0200
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $PAGE_READWRITE = 0x00000004
Global Const $MEM_RELEASE = 0x00008000
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aResult = DllCall("User32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aResult[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "struct*")
If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
Local $aResult = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iProcessID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iProcessID)
If @error Then Return SetError(@error + 10, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return 0
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 20, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iLastError = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iProcessID)
$iError = @error
$iLastError = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 30
$iLastError = @extended
EndIf
Else
$iError = @error + 40
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iLastError, $iRet)
EndFunc
Global Const $LVCF_FMT = 0x0001
Global Const $LVCF_IMAGE = 0x0010
Global Const $LVCF_TEXT = 0x0004
Global Const $LVCF_WIDTH = 0x0002
Global Const $LVCFMT_BITMAP_ON_RIGHT = 0x1000
Global Const $LVCFMT_CENTER = 0x0002
Global Const $LVCFMT_COL_HAS_IMAGES = 0x8000
Global Const $LVCFMT_IMAGE = 0x0800
Global Const $LVCFMT_LEFT = 0x0000
Global Const $LVCFMT_RIGHT = 0x0001
Global Const $LVIF_IMAGE = 0x00000002
Global Const $LVIF_PARAM = 0x00000004
Global Const $LVIF_TEXT = 0x00000001
Global Const $LVS_DEFAULT = 0x0000000D
Global Const $LVS_EX_FULLROWSELECT = 0x00000020
Global Const $LVS_EX_GRIDLINES = 0x00000001
Global Const $LVM_FIRST = 0x1000
Global Const $LVM_DELETEALLITEMS =($LVM_FIRST + 9)
Global Const $LVM_GETHEADER =($LVM_FIRST + 31)
Global Const $LVM_GETITEMA =($LVM_FIRST + 5)
Global Const $LVM_GETITEMW =($LVM_FIRST + 75)
Global Const $LVM_GETITEMCOUNT =($LVM_FIRST + 4)
Global Const $LVM_GETUNICODEFORMAT = 0x2000 + 6
Global Const $LVM_INSERTCOLUMNA =($LVM_FIRST + 27)
Global Const $LVM_INSERTCOLUMNW =($LVM_FIRST + 97)
Global Const $LVM_INSERTITEMA =($LVM_FIRST + 7)
Global Const $LVM_INSERTITEMW =($LVM_FIRST + 77)
Global Const $LVM_SETCOLUMNA =($LVM_FIRST + 26)
Global Const $LVM_SETCOLUMNW =($LVM_FIRST + 96)
Global Const $LVM_SETCOLUMNWIDTH =($LVM_FIRST + 30)
Global Const $LVM_SETEXTENDEDLISTVIEWSTYLE =($LVM_FIRST + 54)
Global Const $LVM_SETITEMA =($LVM_FIRST + 6)
Global Const $LVM_SETITEMW =($LVM_FIRST + 76)
Global $__g_hLVLastWnd
Global Const $__LISTVIEWCONSTANT_ClassName = "SysListView32"
Global Const $__LISTVIEWCONSTANT_WM_SETREDRAW = 0x000B
Global Const $__LISTVIEWCONSTANT_WM_SETFONT = 0x0030
Global Const $__LISTVIEWCONSTANT_DEFAULT_GUI_FONT = 17
Global Const $tagLVCOLUMN = "uint Mask;int Fmt;int CX;ptr Text;int TextMax;int SubItem;int Image;int Order;int cxMin;int cxDefault;int cxIdeal"
Func _GUICtrlListView_BeginUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, False) = 0
EndFunc
Func _GUICtrlListView_Create($hWnd, $sHeaderText, $iX, $iY, $iWidth = 150, $iHeight = 150, $iStyle = 0x0000000D, $iExStyle = 0x00000000, $bCoInit = False)
If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)
If Not IsString($sHeaderText) Then Return SetError(2, 0, 0)
If $iWidth = -1 Then $iWidth = 150
If $iHeight = -1 Then $iHeight = 150
If $iStyle = -1 Then $iStyle = $LVS_DEFAULT
If $iExStyle = -1 Then $iExStyle = 0x00000000
Local Const $S_OK = 0x0
Local Const $S_FALSE = 0x1
Local Const $RPC_E_CHANGED_MODE = 0x80010106
Local Const $E_INVALIDARG = 0x80070057
Local Const $E_OUTOFMEMORY = 0x8007000E
Local Const $E_UNEXPECTED = 0x8000FFFF
Local $sSeparatorChar = Opt('GUIDataSeparatorChar')
Local Const $COINIT_APARTMENTTHREADED = 0x02
Local $iStr_len = StringLen($sHeaderText)
If $iStr_len Then $sHeaderText = StringSplit($sHeaderText, $sSeparatorChar)
$iStyle = BitOR($__UDFGUICONSTANT_WS_CHILD, $__UDFGUICONSTANT_WS_VISIBLE, $iStyle)
If $bCoInit Then
Local $aResult = DllCall('ole32.dll', 'long', 'CoInitializeEx', 'ptr', 0, 'dword', $COINIT_APARTMENTTHREADED)
If @error Then Return SetError(@error, @extended, 0)
Switch $aResult[0]
Case $S_OK
Case $S_FALSE
Case $RPC_E_CHANGED_MODE
Case $E_INVALIDARG
Case $E_OUTOFMEMORY
Case $E_UNEXPECTED
EndSwitch
EndIf
Local $nCtrlID = __UDF_GetNextGlobalID($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Local $hList = _WinAPI_CreateWindowEx($iExStyle, $__LISTVIEWCONSTANT_ClassName, "", $iStyle, $iX, $iY, $iWidth, $iHeight, $hWnd, $nCtrlID)
_SendMessage($hList, $__LISTVIEWCONSTANT_WM_SETFONT, _WinAPI_GetStockObject($__LISTVIEWCONSTANT_DEFAULT_GUI_FONT), True)
If $iStr_len Then
For $x = 1 To $sHeaderText[0]
_GUICtrlListView_InsertColumn($hList, $x - 1, $sHeaderText[$x], 75)
Next
EndIf
Return $hList
EndFunc
Func _GUICtrlListView_DeleteAllItems($hWnd)
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
Local Const $LV_WM_SETREDRAW = 0x000B
Local $vCID = 0
If IsHWnd($hWnd) Then
$vCID = _WinAPI_GetDlgCtrlID($hWnd)
Else
$vCID = $hWnd
$hWnd = GUICtrlGetHandle($hWnd)
EndIf
If $vCID Then
GUICtrlSendMsg($vCID, $LV_WM_SETREDRAW, False, 0)
Local $iParam = 0
For $iIndex = _GUICtrlListView_GetItemCount($hWnd) - 1 To 0 Step -1
$iParam = _GUICtrlListView_GetItemParam($hWnd, $iIndex)
If GUICtrlGetState($iParam) > 0 And GUICtrlGetHandle($iParam) = 0 Then
GUICtrlDelete($iParam)
EndIf
Next
GUICtrlSendMsg($vCID, $LV_WM_SETREDRAW, True, 0)
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
EndIf
Return _SendMessage($hWnd, $LVM_DELETEALLITEMS) <> 0
EndFunc
Func _GUICtrlListView_EndUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, True) = 0
EndFunc
Func _GUICtrlListView_GetColumnCount($hWnd)
Return _SendMessage(_GUICtrlListView_GetHeader($hWnd), 0x1200)
EndFunc
Func _GUICtrlListView_GetHeader($hWnd)
If IsHWnd($hWnd) Then
Return HWnd(_SendMessage($hWnd, $LVM_GETHEADER))
Else
Return HWnd(GUICtrlSendMsg($hWnd, $LVM_GETHEADER, 0, 0))
EndIf
EndFunc
Func _GUICtrlListView_GetItemCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETITEMCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETITEMCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetItemEx($hWnd, ByRef $tItem)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
$iRet = _SendMessage($hWnd, $LVM_GETITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem, $tMemMap)
_MemWrite($tMemMap, $tItem)
If $bUnicode Then
_SendMessage($hWnd, $LVM_GETITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
_SendMessage($hWnd, $LVM_GETITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemRead($tMemMap, $pMemory, $tItem, $iItem)
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_GetItemParam($hWnd, $iIndex)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Mask", $LVIF_PARAM)
DllStructSetData($tItem, "Item", $iIndex)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Return DllStructGetData($tItem, "Param")
EndFunc
Func _GUICtrlListView_GetUnicodeFormat($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETUNICODEFORMAT) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETUNICODEFORMAT, 0, 0) <> 0
EndIf
EndFunc
Func _GUICtrlListView_InsertColumn($hWnd, $iIndex, $sText, $iWidth = 50, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = BitOR($LVCF_FMT, $LVCF_WIDTH, $LVCF_TEXT)
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
Local $iFmt = $aAlign[$iAlign]
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "TextMax", $iBuffer)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tColumn, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $tColumn, 0, "wparam", "struct*")
Else
Local $iColumn = DllStructGetSize($tColumn)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iColumn + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iColumn
DllStructSetData($tColumn, "Text", $pText)
_MemWrite($tMemMap, $tColumn, $pMemory, $iColumn)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pColumn = DllStructGetPtr($tColumn)
DllStructSetData($tColumn, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $pColumn)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTCOLUMNA, $iIndex, $pColumn)
EndIf
EndIf
If $iAlign > 0 Then _GUICtrlListView_SetColumn($hWnd, $iRet, $sText, $iWidth, $iAlign, $iImage, $bOnRight)
Return $iRet
EndFunc
Func _GUICtrlListView_InsertItem($hWnd, $sText, $iIndex = -1, $iImage = -1, $iParam = 0)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iBuffer, $tBuffer, $iRet
If $iIndex = -1 Then $iIndex = 999999999
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Param", $iParam)
$iBuffer = StringLen($sText) + 1
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Text", DllStructGetPtr($tBuffer))
DllStructSetData($tItem, "TextMax", $iBuffer)
Local $iMask = BitOR($LVIF_TEXT, $LVIF_PARAM)
If $iImage >= 0 Then $iMask = BitOR($iMask, $LVIF_IMAGE)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "Image", $iImage)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Or($sText = -1) Then
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet
EndFunc
Func _GUICtrlListView_SetColumn($hWnd, $iIndex, $sText, $iWidth = -1, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = $LVCF_TEXT
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
$iMask = BitOR($iMask, $LVCF_FMT)
Local $iFmt = $aAlign[$iAlign]
If $iWidth <> -1 Then $iMask = BitOR($iMask, $LVCF_WIDTH)
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
Else
$iImage = 0
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "TextMax", $iBuffer)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tColumn, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNW, $iIndex, $tColumn, 0, "wparam", "struct*")
Else
Local $iColumn = DllStructGetSize($tColumn)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iColumn + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iColumn
DllStructSetData($tColumn, "Text", $pText)
_MemWrite($tMemMap, $tColumn, $pMemory, $iColumn)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pColumn = DllStructGetPtr($tColumn)
DllStructSetData($tColumn, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNW, $iIndex, $pColumn)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNA, $iIndex, $pColumn)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetColumnWidth($hWnd, $iCol, $iWidth)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_SETCOLUMNWIDTH, $iCol, $iWidth)
Else
Return GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNWIDTH, $iCol, $iWidth)
EndIf
EndFunc
Func _GUICtrlListView_SetExtendedListViewStyle($hWnd, $iExStyle, $iExMask = 0)
Local $iRet
If IsHWnd($hWnd) Then
$iRet = _SendMessage($hWnd, $LVM_SETEXTENDEDLISTVIEWSTYLE, $iExMask, $iExStyle)
_WinAPI_InvalidateRect($hWnd)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETEXTENDEDLISTVIEWSTYLE, $iExMask, $iExStyle)
_WinAPI_InvalidateRect(GUICtrlGetHandle($hWnd))
EndIf
Return $iRet
EndFunc
Func _GUICtrlListView_SetItemText($hWnd, $iIndex, $sText, $iSubItem = 0)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iRet
If $iSubItem = -1 Then
Local $sSeparatorChar = Opt('GUIDataSeparatorChar')
Local $i_Cols = _GUICtrlListView_GetColumnCount($hWnd)
Local $a_Text = StringSplit($sText, $sSeparatorChar)
If $i_Cols > $a_Text[0] Then $i_Cols = $a_Text[0]
For $i = 1 To $i_Cols
$iRet = _GUICtrlListView_SetItemText($hWnd, $iIndex, $a_Text[$i], $i - 1)
If Not $iRet Then ExitLoop
Next
Return $iRet
EndIf
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Mask", $LVIF_TEXT)
DllStructSetData($tItem, "item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tItem, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_SETITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_SETITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_SETITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
DllStructSetData($tItem, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _WinAPI_SetWindowTheme($hWnd, $sName = 0, $sList = 0)
Local $sTypeOfName = 'wstr', $sTypeOfList = 'wstr'
If Not IsString($sName) Then
$sTypeOfName = 'ptr'
$sName = 0
EndIf
If Not IsString($sList) Then
$sTypeOfList = 'ptr'
$sList = 0
EndIf
Local $aRet = DllCall('uxtheme.dll', 'long', 'SetWindowTheme', 'hwnd', $hWnd, $sTypeOfName, $sName, $sTypeOfList, $sList)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then Return SetError(10, $aRet[0], 0)
Return 1
EndFunc
Func _WinAPI_PathStripPath($sPath)
Local $aRet = DllCall('shlwapi.dll', 'none', 'PathStripPathW', 'wstr', $sPath)
If @error Then Return SetError(@error, @extended, '')
Return $aRet[1]
EndFunc
Global Const $OAIF_EXEC = 0x00000004
Global Const $tagPRINTDLG = __Iif(@AutoItX64, '', 'align 2;') & 'dword Size;hwnd hOwner;handle hDevMode;handle hDevNames;handle hDC;dword Flags;word FromPage;word ToPage;word MinPage;word MaxPage;word Copies;handle hInstance;lparam lParam;ptr PrintHook;ptr SetupHook;ptr PrintTemplateName;ptr SetupTemplateName;handle hPrintTemplate;handle hSetupTemplate'
Func _WinAPI_ShellOpenWithDlg($sFile, $iFlags = 0, $hParent = 0)
Local $tOPENASINFO = DllStructCreate('ptr;ptr;dword;wchar[' &(StringLen($sFile) + 1) & ']')
DllStructSetData($tOPENASINFO, 1, DllStructGetPtr($tOPENASINFO, 4))
DllStructSetData($tOPENASINFO, 2, 0)
DllStructSetData($tOPENASINFO, 3, $iFlags)
DllStructSetData($tOPENASINFO, 4, $sFile)
Local $aRet = DllCall('shell32.dll', 'long', 'SHOpenWithDialog', 'hwnd', $hParent, 'struct*', $tOPENASINFO)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then Return SetError(10, $aRet[0], 0)
Return 1
EndFunc
_GDIPlus_Startup()
Global Const $sVersion = "3.7.2"
Global $ARR_CHECK[15] = ["Resources\MENU.PNG", "Resources\logo.png", "Resources\logo.ico", "Resources\icon_share.ico", "Resources\icon_picture.ico", "Resources\icon_info.ico", "Resources\icon_folder.ico", "Resources\ICON_PROP.ICO", "Resources\icon_details.ico", "Resources\icon_delete.ico", "Resources\icon_copy.ico", "Resources\HELP_USER.PNG", "Resources\icon_001.ico", "About.exe", "config.exe"]
For $i = 0 To UBound($ARR_CHECK) - 1
If Not FileExists(@ScriptDir & "\" & $ARR_CHECK[$i]) Then
Exit MsgBox(16, "Wallpaper Explorer Error Message", 'Oops! We can not find "' & StringUpper($ARR_CHECK[$i]) & '" file' & @CRLF & "Program Version: " & FileGetVersion(@ScriptFullPath) & @CRLF & @CRLF & @CRLF & "@errorcode = 0x" & Hex($i + 50, 2) & Hex(@MDAY, 2) & Hex(@MON, 2) & Hex(StringTrimLeft(@YEAR, 2), 2) & @CRLF)
EndIf
Next
Global Const $INI_FILE = @ScriptDir & "\Wallpaper Explorer.ini"
Global Const $sFile = @ScriptDir & "\WallPaperExplorer.exe"
Global Const $sURL_Folder = "http://googledrive.com/host/0B4S2pTQPGLo6NjZjZFBmOWpaMEk"
Global Const $sURL_Check = $sURL_Folder & "/WallpaperExplorerUpdate.txt"
Global Const $sLocalFileDir = @TempDir & "\CheckForUpdate.ini"
Global $hCheckUpdate = InetGet($sURL_Check, $sLocalFileDir, 1, 1)
If Not FileExists($INI_FILE) Then FileOpen($INI_FILE, 8)
Local $sRegRead = RegRead("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Wallpaper Explorer\command", "")
If($sRegRead <> '"' & @ScriptFullPath & '"' Or @error) Then
$Form1 = GUICreate("Install Context Menu", 580, 450, -1, -1, $WS_SYSMENU)
GUISetBkColor(0xFFFFFF)
GUICtrlCreateLabel("Do you want to create context menu of Wallpaper Explorer program in desktop?", 10, 5, 600, 25)
GUICtrlSetColor(-1, 0x444444)
GUICtrlSetFont(-1, 12, 400, 0, "Segoe UI")
$hPic1 = GUICtrlGetHandle(GUICtrlCreatePic("", 0, 40, _WinAPI_GetClientWidth($Form1), _WinAPI_GetClientHeight($Form1) - 80))
Local $hScanBitmap = _GDIPlus_BitmapCreateFromScan0(_WinAPI_GetClientWidth($Form1), _WinAPI_GetClientHeight($Form1) - 80)
Local $hContext = _GDIPlus_ImageGetGraphicsContext($hScanBitmap)
_GDIPlus_GraphicsSetSmoothingMode($hContext, $GDIP_SMOOTHINGMODE_HIGHQUALITY)
_GDIPlus_GraphicsClear($hContext, 0xFFFFFF)
Local $hUser_Help = _GDIPlus_BitmapCreateFromFile(@ScriptDir & "\Resources\HELP_USER.PNG")
_GDIPlus_GraphicsDrawImage($hContext, $hUser_Help, 160, 0)
Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hScanBitmap)
_SendMessage($hPic1, 0x0172, $IMAGE_BITMAP, $hBitmap)
$hButton1 = GUICtrlCreateButton("OK", 355, 390, 100, 25)
$hButton2 = GUICtrlCreateButton("Cancel", 465, 390, 100, 25)
GUISetState(@SW_SHOW)
While 1
$nMsg = GUIGetMsg()
Switch $nMsg
Case $hButton1
ShellExecuteWait("config.exe", "/start")
$nMsg = $GUI_EVENT_CLOSE
ContinueCase
Case $hButton2
IniWrite($INI_FILE, "INFO", "DISPLAY", False)
$nMsg = $GUI_EVENT_CLOSE
ContinueCase
Case $GUI_EVENT_CLOSE
_GDIPlus_GraphicsDispose($hContext)
_GDIPlus_BitmapDispose($hScanBitmap)
_GDIPlus_BitmapDispose($hUser_Help)
_WinAPI_DeleteObject($hBitmap)
GUIDelete($Form1)
ExitLoop
EndSwitch
WEnd
EndIf
Global $bSpecialOsVersion = False
If StringInStr("WIN_7_ WIN_VISTA_ WIN_XP_ WIN_XPe_", @OSVersion & "_") Then $bSpecialOsVersion = True
Global $sPathWallpaper, $hImageSource = 0, $temp_bitmap, $sImageInfo = ""
Global $__EditingWallpaper = False
Global $__WallpaperEdited = False
Global $__hGUI2_IsShow = True
Global Const $hMenuIcon = _GDIPlus_ImageResize(_GDIPlus_ImageLoadFromFile(@ScriptDir & "\Resources\MENU.PNG"), 45, 45)
Global Const $hGUI = GUICreate("Wallpaper Explorer v" & $sVersion, 0, 0, -200, 0, BitOR($WS_MAXIMIZEBOX, $WS_SIZEBOX, $GUI_SS_DEFAULT_GUI, $WS_EX_APPWINDOW))
Global Const $hGUI2 = GUICreate("", 200, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $hGUI3 = GUICreate("", _GDIPlus_ImageGetWidth($hMenuIcon), _GDIPlus_ImageGetHeight($hMenuIcon), 10, 10, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hGUI2)
Global Const $hGUI4 = GUICreate("", 250, 0, 0, 0, $WS_POPUP, $WS_EX_MDICHILD, $hGUI)
Global Const $__ImageInfoGUI = GUICreate("Details", 600, 500, 1, 1, $WS_SYSMENU, $WS_EX_MDICHILD, $hGUI)
GUISwitch($hGUI)
Global Const $hPic = GUICtrlGetHandle(GUICtrlCreatePic('', 0, 0, 500, 200, BitOR($SS_CENTERIMAGE, $SS_NOPREFIX, $SS_NOTIFY)))
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUISwitch($hGUI2)
Global Const $ARR_BUTTON_TEXT[8] = ["Open with...", "Delete", "Recycle bin", "Image Details", "Quick edit", "Open location", "Properties", "About"]
Global Const $ARR_BUTTON_ICON[8] = ["ICON_SHARE.ico", "ICON_DELETE.ICO", "ICON_DELETE.ICO", "ICON_DETAILS.ico", "icon_picture.ico", "ICON_FOLDER.ico", "ICON_PROP.ICO", "icon_info.ico"]
Global $ARR_HANDLE_BUTTON[8]
For $i = 0 To UBound($ARR_BUTTON_TEXT) - 1
$ARR_HANDLE_BUTTON[$i] = _GUICtrlMetroButton_Create($hGUI2, $ARR_BUTTON_TEXT[$i], 0, 60 + 50 * $i, 200, 50, @ScriptDir & "\Resources\" & $ARR_BUTTON_ICON[$i])
Next
GUISwitch($hGUI3)
SetBitmap($hGUI3, $hMenuIcon, 255)
GUISwitch($hGUI4)
Global Const $TBS_BOTH = 0x0008
Global Const $TBS_NOTICKS = 0x0010
Global Const $TBS_TOOLTIPS = 0x100
Global $g_aDefault[11] = [10000, 10000, 10000, 0, 10000, 0, 0, 0, 0, 0, 0]
Global $g_aAdjust = $g_aDefault
Global Const $hGUI4_BUTTONBACK = _GUICtrlMetroButton_Create($hGUI4, "Back", 0, 0, _WinAPI_GetClientWidth($hGUI4), 50, @ScriptDir & "\Resources\icon_001.ico")
Global $g_aidSlider[9][6] = [[0, 0, 'Red', 25, 650, 100], [0, 0, 'Green', 25, 650, 100], [0, 0, 'Blue', 25, 650, 100], [0, 0, 'Black', 0, 400, 10], [0, 0, 'White', 600, 1000, 10], [0, 0, 'Contrast', -100, 100, 1], [0, 0, 'Brightness', -100, 100, 1], [0, 0, 'Colorfulness', -100, 100, 1], [0, 0, 'Tint', -100, 100, 1]]
Global Const $iDeltaHeight = 50
For $i = 0 To 8
GUICtrlCreateLabel($g_aidSlider[$i][2] & ':', 10, 21 + $i * 31 + $iDeltaHeight, 60, 14, $SS_RIGHT)
GUICtrlSetBkColor(-1, 0xFFFFFF)
$g_aidSlider[$i][0] = GUICtrlCreateSlider(71, 19 + $i * 31 + $iDeltaHeight, 170, 20, BitOR($TBS_BOTH, $TBS_TOOLTIPS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, $g_aidSlider[$i][4], $g_aidSlider[$i][3])
$g_aidSlider[$i][1] = GUICtrlGetHandle(-1)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Next
GUICtrlCreateLabel('Illuminant:', 10, 310 + $iDeltaHeight, 60, 14, $SS_RIGHT)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global Const $CBS_DROPDOWNLIST = 0x3
Global Const $CB_ADDSTRING = 0x143
Global Const $CB_GETCURSEL = 0x147
Global Const $CB_SETCURSEL = 0x14E
Func _GUICtrlComboBox_AddString($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_ADDSTRING, 0, $sText, 0, "wparam", "wstr")
EndFunc
Func _GUICtrlComboBox_GetCurSel($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETCURSEL)
EndFunc
Func _GUICtrlComboBox_SetCurSel($hWnd, $iIndex = -1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_SETCURSEL, $iIndex)
EndFunc
Global $g_idCombo = GUICtrlCreateCombo('', 77, 306 + $iDeltaHeight, 158, 160, $CBS_DROPDOWNLIST)
Local $aLight = StringSplit('Default|Tungsten lamp|Noon sunlight|NTSC daylight|Normal print|Bond paper print|Standard daylight|Northern daylight|Cool white lamp', '|')
For $i = 1 To $aLight[0]
_GUICtrlComboBox_AddString(-1, $aLight[$i])
Next
GUICtrlCreateLabel('Filters:', 10, 347 + $iDeltaHeight, 60, 14, $SS_RIGHT)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $g_aidCheck[3]
$g_aidCheck[0] = GUICtrlCreateCheckbox('Negative (invert colors)', 77, 343 + $iDeltaHeight, 131, 21)
GUICtrlSetBkColor(-1, 0xFFFFFF)
$g_aidCheck[1] = GUICtrlCreateCheckbox('Logarithmic', 77, 368 + $iDeltaHeight, 76, 21)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Local $aidButton[2]
$aidButton[0] = GUICtrlCreateButton('Reset', 136 - 100, 448 + $iDeltaHeight, 75, 25)
$aidButton[1] = GUICtrlCreateButton('Save...', 218 - 100, 448 + $iDeltaHeight, 75, 25)
GUICtrlSetState($g_aidCheck[0], $GUI_FOCUS)
_Reset()
GUISwitch($__ImageInfoGUI)
Global Const $hListView = _GUICtrlListView_Create($__ImageInfoGUI, " Property | Value ", 0, 0, _WinAPI_GetClientWidth($__ImageInfoGUI), _WinAPI_GetClientHeight($__ImageInfoGUI))
_GUICtrlListView_SetExtendedListViewStyle($hListView, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
_GUICtrlListView_SetColumnWidth($hListView, 0, 200)
_GUICtrlListView_SetColumnWidth($hListView, 1, _WinAPI_GetClientWidth($__ImageInfoGUI) - 200)
_WinAPI_SetWindowTheme($hListView, "Explorer")
GUIRegisterMsg($WM_ENTERSIZEMOVE, "WM_ENTERSIZEMOVE")
GUIRegisterMsg($WM_EXITSIZEMOVE, "WM_EXITSIZEMOVE")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_HSCROLL, 'WM_HSCROLL')
GUISetBkColor(0xFFFFFF, $hGUI2)
WinSetTrans($hGUI2, "", 200)
GUISetBkColor(0xFFFFFF, $hGUI)
GUISetBkColor(0xFFFFFF, $hGUI4)
WinSetTrans($hGUI4, "", 230)
GUISetState(@SW_SHOW, $hGUI)
Sleep(300)
GUISetState(@SW_SHOW, $hGUI2)
GUISetState(@SW_SHOW, $hGUI3)
While True
$nMsg = GUIGetMsg(1)
Switch $nMsg[1]
Case $hGUI
Switch $nMsg[0]
Case $GUI_EVENT_CLOSE
ExitLoop
Case $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESTORE
WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
EndSwitch
Case $hGUI3
Switch $nMsg[0]
Case $GUI_EVENT_PRIMARYUP
$arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
_WinAPI_MoveWindow($hGUI2, $arrPos[0], $arrPos[1], _WinAPI_GetWindowWidth($hGUI2), _WinAPI_GetClientHeight($hGUI))
If $__hGUI2_IsShow Then
_Animation_FaceOut($hGUI2)
Else
_Animation_FaceIn($hGUI2)
EndIf
__hGUI2State_Toggle()
EndSwitch
Case $hGUI4
Switch $nMsg[0]
Case $aidButton[0]
_Reset()
Case $aidButton[1]
_SaveImage()
Case $g_aidCheck[0]
If GUICtrlRead($g_aidCheck[0]) = $GUI_CHECKED Then
$g_aAdjust[9] = BitOR($g_aAdjust[9], $CA_NEGATIVE)
Else
$g_aAdjust[9] = BitAND($g_aAdjust[9], BitNOT($CA_NEGATIVE))
EndIf
_Update()
Case $g_aidCheck[1]
If GUICtrlRead($g_aidCheck[1]) = $GUI_CHECKED Then
$g_aAdjust[9] = BitOR($g_aAdjust[9], $CA_LOG_FILTER)
Else
$g_aAdjust[9] = BitAND($g_aAdjust[9], BitNOT($CA_LOG_FILTER))
EndIf
_Update()
Case $g_idCombo
$sData = _GUICtrlComboBox_GetCurSel($g_idCombo)
If $g_aAdjust[10] <> $sData Then
$g_aAdjust[10] = $sData
_Update()
EndIf
EndSwitch
Case $__ImageInfoGUI
Switch $nMsg[0]
Case $GUI_EVENT_CLOSE
GUISetState(@SW_HIDE, $__ImageInfoGUI)
GUISetState(@SW_ENABLE, $hGUI)
GUISetState(@SW_ENABLE, $hGUI2)
GUISetState(@SW_ENABLE, $hGUI3)
GUISetState(@SW_ENABLE, $hGUI4)
WinActivate($hGUI)
EndSwitch
EndSwitch
If __WallpaperIsChanged() And Not $__EditingWallpaper Then
__ResetImage()
WinSetTitle($__ImageInfoGUI, '', 'Details of "' & _WinAPI_PathStripPath($sPathWallpaper) & '"')
__ImageInfoUpdate()
EndIf
If InetGetInfo($hCheckUpdate, 2) Then
Local $iPresentVersion = FileGetVersion($sFile)
Local $sNewVersion = IniRead($sLocalFileDir, "Info", "Version", "")
If StringReplace($sNewVersion, ".", "") > StringReplace($iPresentVersion, ".", "") Then MsgBox(64, "Message", "Hi " & @ComputerName & ", new update is now available. Go to About -> Check for update to update program. Thank you!", $hGUI)
InetClose($hCheckUpdate)
$hCheckUpdate = 0
EndIf
WEnd
Func _SaveImage()
$sData = FileSaveDialog('Save Image', @WorkingDir, 'Image Files (*.bmp;*.dib;*.gif;*.jpg;*.tif)|All Files (*.*)', 2 + 16, @ScriptDir & '\Wallpaper_' & Random(1, 999999999, 1) & '.jpg', $hGUI)
GUICtrlSetState($aidButton[1], $GUI_FOCUS)
If Not $sData Then
Return
EndIf
Local $aAdjust = $g_aAdjust
If IsArray($aAdjust) Then
$tAdjust = _WinAPI_CreateColorAdjustment($aAdjust[9], $aAdjust[10], $aAdjust[0], $aAdjust[1], $aAdjust[2], $aAdjust[3], $aAdjust[4], $aAdjust[5], $aAdjust[6], $aAdjust[7], $aAdjust[8])
EndIf
Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImageSource)
Local $hAdjustBitmap = _WinAPI_AdjustBitmap($hBitmap, -1, -1, $HALFTONE, $tAdjust)
If $hAdjustBitmap Then
Local $hBitmap_d = _GDIPlus_BitmapCreateFromHBITMAP($hAdjustBitmap)
If _GDIPlus_ImageSaveToFile($hBitmap_d, $sData) Then
$sData = 0
EndIf
_WinAPI_DeleteObject($hAdjustBitmap)
_GDIPlus_BitmapDispose($hBitmap_d)
EndIf
_WinAPI_DeleteObject($hBitmap)
If $sData Then
MsgBox(16, "Message", "Unable to save image")
Else
MsgBox(64, "Message", "Successful to save image")
EndIf
EndFunc
_GDIPlus_BitmapDispose($hMenuIcon)
_WinAPI_DeleteObject($temp_bitmap)
_GDIPlus_ImageDispose($hImageSource)
_GDIPlus_Shutdown()
GUIRegisterMsg($WM_ENTERSIZEMOVE, "")
GUIRegisterMsg($WM_EXITSIZEMOVE, "")
GUIRegisterMsg($WM_COMMAND, "")
GUIRegisterMsg($WM_HSCROLL, "")
GUIDelete($__ImageInfoGUI)
GUIDelete($hGUI4)
GUIDelete($hGUI3)
GUIDelete($hGUI2)
GUIDelete($hGUI)
FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
InetClose($hCheckUpdate)
ProcessClose(ProcessExists("about.exe"))
Func _Animation_FaceOut($hWnd, $ihGUI4 = False)
If Not $ihGUI4 Then
For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
_WinAPI_PostMessage($ARR_HANDLE_BUTTON[$i], $WM_COMMAND, $WM_NEXTANIMATION, 0)
Next
EndIf
Local $iW = _WinAPI_GetWindowWidth($hWnd), $iH = _WinAPI_GetWindowHeight($hWnd)
Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
Local $iOW = $iW
While 1
_WinAPI_MoveWindow($hWnd, $arrPos[0], $arrPos[1], $iW, $iH, False)
If $iW = 0 Then ExitLoop
If $bSpecialOsVersion Then
If $iW > $iOW * 1 / 6 Then
$iW -= 1
ElseIf $iW > $iOW / 3 And $iW <= $iOW * 1 / 6 Then
$iW -= 3
ElseIf $iW > $iOW / 2 And $iW <= $iOW / 3 Then
$iW -= 8
Else
$iW -= 20
EndIf
Else
If $iW > $iOW * 1 / 6 Then
$iW -= 2
ElseIf $iW > $iOW / 3 And $iW <= $iOW * 1 / 6 Then
$iW -= 5
ElseIf $iW > $iOW / 2 And $iW <= $iOW / 3 Then
$iW -= 10
Else
$iW -= 20
EndIf
EndIf
If $iW <= 0 Then $iW = 0
WEnd
EndFunc
Func _Animation_FaceIn($hWnd)
Local $iOW = 200, $iH = _WinAPI_GetWindowHeight($hWnd)
Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
Local $iW = 0
While 1
_WinAPI_MoveWindow($hWnd, $arrPos[0], $arrPos[1], $iW, $iH, True)
If $iW = $iOW Then ExitLoop
If $bSpecialOsVersion Then
If $iW > $iOW * 5 / 6 Then
$iW += 0.1
ElseIf $iW > $iOW / 4 And $iW <= $iOW * 5 / 6 Then
$iW += 0.5
ElseIf $iW > $iOW / 3 And $iW <= $iOW / 4 Then
$iW += 5
Else
$iW += 15
EndIf
Else
If $iW > $iOW * 5 / 6 Then
$iW += 0.5
ElseIf $iW > $iOW / 4 And $iW <= $iOW * 5 / 6 Then
$iW += 5
ElseIf $iW > $iOW / 3 And $iW <= $iOW / 4 Then
$iW += 8
Else
$iW += 15
EndIf
EndIf
If $iW >= $iOW Then $iW = $iOW
WEnd
For $i = 0 To UBound($ARR_HANDLE_BUTTON) - 1
_WinAPI_PostMessage($ARR_HANDLE_BUTTON[$i], $WM_COMMAND, $WM_NEXTANIMATION, 0)
Next
EndFunc
Func __ImageInfoUpdate()
Local $__arrInfo = StringRegExp($sImageInfo, "(?m)^.*$", 3)
Local $__tempString = 0
If Not IsArray($__arrInfo) Then Return ConsoleWrite("Not array..." & @CRLF)
_GUICtrlListView_BeginUpdate($hListView)
_GUICtrlListView_DeleteAllItems($hListView)
For $i = 0 To UBound($__arrInfo) - 1
$__tempString = StringSplit($__arrInfo[$i], "=")
_GUICtrlListView_InsertItem($hListView, $__tempString[1], $i)
_GUICtrlListView_SetItemText($hListView, $i, $__tempString[2], 1)
Next
_GUICtrlListView_EndUpdate($hListView)
EndFunc
Func WM_COMMAND($hWnd, $iMsg, $lParam, $wParam)
Switch $hWnd
Case $hGUI2
Switch $lParam
Case $WM_LBUTTONUP
Switch $wParam
Case $ARR_HANDLE_BUTTON[0]
_WinAPI_ShellOpenWithDlg($sPathWallpaper, $OAIF_EXEC, $hGUI)
Case $ARR_HANDLE_BUTTON[1]
If MsgBox(64 + 1, "Wallpaper Explorer Message", "Do your want to delete your wallpaper file? You can't restore it. Still continue?" & @CRLF & @CRLF) = 1 Then
FileDelete($sPathWallpaper)
If FileExists($sPathWallpaper) Then
Local $sText = ""
$sText = "- We have not permission to delete file in system folder or system disk" & @CRLF & "- This file is used by other processes" & @CRLF & "- Attribute of your disk driver or your file is read-only" & @CRLF & "- Your file has been crashed. "
MsgBox(16, "Wallpaper Explorer Error Message", "Oops, we are meeting some errors when deleting your wallpaper file. Please checking again some following reasons:" & @CRLF & @CRLF & $sText & @CRLF)
Else
MsgBox(64, "Wallpaper Explorer Message", "Successful deleting file!")
EndIf
EndIf
Case $ARR_HANDLE_BUTTON[2]
FileRecycle($sPathWallpaper)
If FileExists($sPathWallpaper) Then
If FileExists($sPathWallpaper) Then
Local $sText = ""
$sText = "- We have not permission to delete file in system folder or system disk" & @CRLF & "- This file is used by other processes" & @CRLF & "- Attribute of your disk driver or your file is read-only" & @CRLF & "- Your file has been crashed. "
MsgBox(16, "Wallpaper Explorer Error Message", "Sorry, your file can not be moved to recycle bin. Please checking again some following reasons:" & @CRLF & @CRLF & $sText & @CRLF)
Else
MsgBox(64, "Wallpaper Explorer Message", "Delete file successfully!")
EndIf
Else
MsgBox(64, "Message", "Move this file to recycle bin successfully!")
EndIf
Case $ARR_HANDLE_BUTTON[3]
GUISetState(@SW_DISABLE, $hGUI)
GUISetState(@SW_DISABLE, $hGUI2)
GUISetState(@SW_DISABLE, $hGUI3)
GUISetState(@SW_DISABLE, $hGUI4)
__ImageInfoUpdate()
Local $iW = _WinAPI_GetWindowWidth($__ImageInfoGUI)
Local $iH = _WinAPI_GetWindowHeight($__ImageInfoGUI)
_WinAPI_MoveWindow($__ImageInfoGUI, @DesktopWidth / 2 - $iW / 2, @DesktopHeight / 2 - $iH / 2, $iW, $iH)
WinSetTitle($__ImageInfoGUI, '', 'Details of "' & _WinAPI_PathStripPath($sPathWallpaper) & '"')
GUISetState(@SW_SHOWNORMAL, $__ImageInfoGUI)
Case $ARR_HANDLE_BUTTON[4]
Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
_WinAPI_MoveWindow($hGUI4, $arrPos[0], $arrPos[1], 250, _WinAPI_GetClientHeight($hGUI))
GUISetState(@SW_SHOW, $hGUI4)
GUISetState(@SW_HIDE, $hGUI3)
GUISetState(@SW_HIDE, $hGUI2)
$__EditingWallpaper = True
_Reset()
Case $ARR_HANDLE_BUTTON[5]
ShellExecute("explorer.exe", '/select,"' & $sPathWallpaper & '"')
Case $ARR_HANDLE_BUTTON[6]
$oshellApp = ObjCreate("shell.application")
$oshellApp.namespace(0).parsename($sPathWallpaper).invokeverb("Properties")
Case $ARR_HANDLE_BUTTON[7]
ShellExecute(@ScriptDir & "\about.exe", "/start")
EndSwitch
EndSwitch
Case $hGUI4
Switch $lParam
Case $WM_LBUTTONUP
Switch $wParam
Case $hGUI4_BUTTONBACK
If $__WallpaperEdited Then
If MsgBox(64 + 4, "Wallpaper Explorer Message", "This image has been changed. Do you want to save it?") = 6 Then
_SaveImage()
EndIf
EndIf
GUISetState(@SW_HIDE, $hGUI4)
GUISetState(@SW_SHOW, $hGUI2)
GUISetState(@SW_SHOW, $hGUI3)
$__EditingWallpaper = False
WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
EndSwitch
EndSwitch
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func _Reset()
$g_aAdjust = $g_aDefault
For $i = 0 To 8
GUICtrlSetData($g_aidSlider[$i][0], $g_aAdjust[$i] / $g_aidSlider[$i][5])
GUICtrlSetData($g_aidSlider[$i][2], $g_aAdjust[$i])
Next
_GUICtrlComboBox_SetCurSel($g_idCombo, $g_aAdjust[10])
If BitAND($g_aAdjust[9], $CA_LOG_FILTER) Then
GUICtrlSetState($g_aidCheck[1], $GUI_CHECKED)
Else
GUICtrlSetState($g_aidCheck[1], $GUI_UNCHECKED)
EndIf
If BitAND($g_aAdjust[9], $CA_NEGATIVE) Then
GUICtrlSetState($g_aidCheck[0], $GUI_CHECKED)
Else
GUICtrlSetState($g_aidCheck[0], $GUI_UNCHECKED)
EndIf
_Update()
$__WallpaperEdited = False
EndFunc
Func _Update()
_SetBitmapAdjust($hPic, $temp_bitmap, $g_aAdjust)
$__WallpaperEdited = True
EndFunc
Func WM_HSCROLL($hWnd, $iMsg, $wParam, $lParam)
#forceref $iMsg, $wParam
Switch $hWnd
Case $hGUI4
For $i = 0 To 8
If $g_aidSlider[$i][1] = $lParam Then
$g_aAdjust[$i] = $g_aidSlider[$i][5] * GUICtrlRead($g_aidSlider[$i][0])
GUICtrlSetData($g_aidSlider[$i][2], $g_aAdjust[$i])
_Update()
ExitLoop
EndIf
Next
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func __hGUI2State_Toggle()
$__hGUI2_IsShow = Not $__hGUI2_IsShow
Return(Not $__hGUI2_IsShow)
EndFunc
Func WM_ENTERSIZEMOVE($hWnd, $iMsg, $lParam, $wParam)
Switch $hWnd
Case $hGUI
Local $arrPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
_WinAPI_MoveWindow($hGUI2, $arrPos[0], $arrPos[1], 0, 400)
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func WM_EXITSIZEMOVE($hWnd, $iMsg, $lParam, $wParam)
Switch $hWnd
Case $hGUI
Local $iW = _WinAPI_GetClientWidth($hGUI)
Local $iH = _WinAPI_GetClientHeight($hGUI)
Local $temp_bitmap_old = $temp_bitmap
Local $__tempimage = $hImageSource
If $__EditingWallpaper Then $__tempimage = _GDIPlus_BitmapCreateFromHBITMAP(_SendMessage($hPic, $STM_GETIMAGE))
$temp_bitmap = _GDIPlus_ImageResizeToBitmap($__tempimage, $iW, $iH)
_WinAPI_MoveWindow($hPic, 0, 0, $iW, $iH)
_SetBitmapAdjust($hPic, $temp_bitmap, 0)
_WinAPI_DeleteObject($temp_bitmap_old)
Local $tRect = _WinAPI_GetClientRect($hGUI)
Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
If $__hGUI2_IsShow Then _WinAPI_MoveWindow($hGUI2, $aGUIPos[0], $aGUIPos[1], 200, _WinAPI_GetClientHeight($hGUI))
_WinAPI_MoveWindow($hGUI4, $aGUIPos[0], $aGUIPos[1], _WinAPI_GetClientWidth($hGUI4), _WinAPI_GetClientHeight($hGUI))
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func _GDIPlus_ImageResizeToBitmap($hImage, $iW, $iH)
Local $temp_image = _GDIPlus_ImageResize($hImage, $iW, $iH)
Local $temp_bitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($temp_image)
_GDIPlus_ImageDispose($temp_image)
Return $temp_bitmap
EndFunc
Func __WallpaperIsChanged()
Local $Now_Wallpaper = _GetPathWallpaper()
If $Now_Wallpaper <> $sPathWallpaper And StringLen($Now_Wallpaper) > 0 Then
Return True
Else
Return False
EndIf
EndFunc
Func __ResetImage()
_GDIPlus_ImageDispose($hImageSource)
FileDelete(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
$sPathWallpaper = _GetPathWallpaper()
$sImageInfo = _ImageGetInfo($sPathWallpaper)
FileCopy($sPathWallpaper, @TempDir & "\", 1)
$hImageSource = _GDIPlus_ImageLoadFromFile(@TempDir & "\" & _WinAPI_PathStripPath($sPathWallpaper))
If $hImageSource = -1 Then Exit ConsoleWrite("! Can't load image file..." & @CRLF)
Local $iW = _GDIPlus_ImageGetWidth($hImageSource)
Local $iH = _GDIPlus_ImageGetHeight($hImageSource)
Local $iTrayHeight = _WinAPI_GetWindowHeight(WinGetHandle("[CLASS:Shell_TrayWnd]"))
If $iW > @DesktopWidth * 0.75 Or $iH > @DesktopHeight * 0.75 Then
$iW = @DesktopWidth * 0.75
$iH = @DesktopHeight * 0.75
EndIf
_WinAPI_MoveWindow($hGUI, @DesktopWidth / 2 - $iW / 2,(@DesktopHeight - $iTrayHeight) / 2 - $iH / 2, $iW, $iH)
Local $tRect = _WinAPI_GetClientRect($hGUI)
Local $aGUIPos = _WinAPI_ConvertPointClientToScreen($hGUI, 0, 0)
_WinAPI_MoveWindow($hGUI2, $aGUIPos[0], $aGUIPos[1], _WinAPI_GetWindowWidth($hGUI2), DllStructGetData($tRect, "bottom") - DllStructGetData($tRect, "top"))
WM_EXITSIZEMOVE($hGUI, 0, 0, 0)
EndFunc
