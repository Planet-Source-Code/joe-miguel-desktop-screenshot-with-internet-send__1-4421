Attribute VB_Name = "SetDIBits"
Type BITMAPINFOHEADER   '40 bytes
   biSize As Long '
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors(0 To 255) As RGBQUAD
End Type


Type BITMAP  '24 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type COLORQUAD
  rgbB As Byte
  rgbG As Byte
  rgbR As Byte
  rgbP As Byte
End Type

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&
Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0&

Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal _
hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, _
lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal _
hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As _
Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As _
Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) _
As Long

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal _
hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal _
X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal _
nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As _
Long

Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap _
As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As _
Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long


  Dim hand As Long, oldhand As Long
  Dim bmap As BITMAP
  Dim srcewid As Long, srcehgt As Long
  Dim srcedibbmap As BITMAPINFO
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim icol As Integer, irow As Integer
  Dim lsuccess As Long
  Dim hdcNew As Long
  Dim srceqarr() As COLORQUAD
  Dim thiscolor As COLORQUAD

    Dim X As Long

Public Sub SetDIB(Array3d() As Byte, Picture As Control)

    hand = Picture.Image.Handle

  'Create a device context compatible with the Desktop.
    hdcNew = Picture.hdc
  
  'Get the source bitmap width and height, in pixels, from BITMAP
  'structure.
    srcewid = Picture.ScaleWidth
    srcehgt = Picture.ScaleHeight
  
  With srcedibbmap.bmiHeader
    .biSize = 40
    .biWidth = srcewid
    .biHeight = -srcehgt
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With

lsuccess = SetDIBits(hdcNew, hand, 0, srcehgt, _
               Array3d(1, 1, 1), srcedibbmap, DIB_RGB_COLORS)

'StretchBlt to PictureBox so we can see the result
lsuccess = StretchBlt(Picture.hdc, _
                0, 0, Picture.ScaleWidth, _
                Picture.ScaleHeight, _
                hdcNew, _
                0, 0, srcewid, srcehgt, _
                vbSrcCopy)

End Sub



