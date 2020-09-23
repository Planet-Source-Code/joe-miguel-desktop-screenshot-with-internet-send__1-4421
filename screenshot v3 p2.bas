Attribute VB_Name = "API"
' Enumerated raster operation constants
    Public Enum RasterOps
        ' Copies the source bitmap to destination bitmap
         SRCCOPY = &HCC0020
        '
        ' Combines pixels of the destination with source bitmap using the Boolean AND operator.
         SRCAND = &H8800C6
        '
        ' Combines pixels of the destination with source bitmap using the Boolean XOR operator.
         SRCINVERT = &H660046
        '
        ' Combines pixels of the destination with source bitmap using the Boolean OR operator.
         SRCPAINT = &HEE0086
        '
        ' Inverts the destination bitmap and then combines the results with the source bitmap
        ' using the Boolean AND operator.
         SRCERASE = &H4400328
        '
        ' Turns all output white.
         WHITENESS = &HFF0062
        '
        ' Turn output black.
         BLACKNESS = &H42
    End Enum
    

Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As RasterOps _
        ) As Long
        
Public Declare Function StretchBlt Lib "gdi32" ( _
        ByVal hDestdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, _
        ByVal dwRop As Long _
        ) As Long


Declare Function GetDesktopWindow Lib "user32" () As Long


Declare Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long


Public Declare Function GetLastError Lib "kernel32" () As Long























