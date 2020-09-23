Attribute VB_Name = "mImageList"
Option Explicit

Public Enum eSpecialColors
    CLR_NONE = &HFFFFFFFF
    CLR_DEFAULT = &HFF000000
End Enum

Public Enum eDrawTypes
    ILD_HIGHLIGHT50 = &H4
    ILD_HIGHLIGHT25 = &H2
    ILD_TRANSPARENT = &H1
End Enum

Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long

Public Function DrawImage(ByVal HDC As Long, ByVal hImagelist As Long, ByVal IconIndex As Long, Optional Left As Long, Optional Top As Long, Optional Width As Long, Optional Height As Long, Optional ByVal DrawType As eDrawTypes = ILD_TRANSPARENT, Optional ByVal BackGround As eSpecialColors = CLR_DEFAULT, Optional ByVal ForeGround As eSpecialColors = CLR_DEFAULT)
    DrawImage = ImageList_DrawEx(hImagelist, IconIndex, HDC, Left, Top, Width, Height, BackGround, ForeGround, DrawType) <> 0
End Function

