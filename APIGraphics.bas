Attribute VB_Name = "modAPIGraphics"
'=================================
' Date : 12/11/2001
'----------------------------------------------
' Author : Eric O'Sullivan
'----------------------------------------------
' Contact : DiskJunky@hotmail.com
'----------------------------------------------
' Comments :
'This module was made for using api graphics functions in your
'programs. With the following api calls and function and procedures
'written by me, you have to tools to do almost anything. The only api
'function listed here that is not directly used by any piece of code
'in this module is BitBlt. You have the tools to create and manipulate
'graphics, but it is still necessary to display them manually. The
'functions themselves mostly need hDc or a handle to work. You can
'find this hDc in both a forms and pictureboxs' properties. I have
'also set up a data type called BitmapStruc. For my programs, I have
'used this structure almost exclusivly for the graphics. The structure
'holds all the information needed to reference a bitmap created using
'this module (CreateNewBitmap, DeleteBitmap).
'Please keep in mind that any object (bitmap, brush, pen etc) needs to
'be deleted after use or else it will stay in memory until the program is
'finished. Not doing so will eventually cause your program to take up
'ALL your computers recources.
'Also for anyone using optional paramters, it is probably better to use
'a default parameter values to determine whether or not a parameter
'has been passed than the function IsMissing().
'
'Thank you,
'Eric
'----------------------------------------------
'=================================

Option Explicit
Option Private Module

'--------------------------------------------------------------------------
'API calls
'--------------------------------------------------------------------------
'These functions are sorted alphabetically.

'This is used to copy bitmaps
Public Declare Function BitBlt _
       Lib "gdi32" _
            (ByVal hDestDC As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal dwRop As Long) _
             As Long

'used to change various windows settings
Public Declare Function ChangeDisplaySettings _
       Lib "user32" _
       Alias "ChangeDisplaySettingsA" _
            (ByRef wef As Any, _
             ByVal i As Long) _
             As Long

'creates a brush object which can be applied to a bitmap
Public Declare Function CreateBrushIndirect _
       Lib "gdi32" _
            (lpLogBrush As LogBrush) _
             As Long

'creates a colourspace object which can be applied to a bitmap
Public Declare Function CreateColorSpace _
       Lib "gdi32" _
       Alias "CreateColorSpaceA" _
            (lplogcolorspace As LogColorSpace) _
             As Long

'the will create a bitmap compatable with the passed hDc
Public Declare Function CreateCompatibleBitmap _
       Lib "gdi32" _
            (ByVal hDc As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long) _
            As Long

'this create a compatable device context with the specified
'windows handle
Public Declare Function CreateCompatibleDC _
       Lib "gdi32" _
            (ByVal hDc As Long) _
             As Long

'creates an elliptical region in a hDc
Public Declare Function CreateEllipticRgn _
       Lib "gdi32" _
            (ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'creates an elliptical region in a hDc
Public Declare Function CreateEllipticRgnIndirect _
       Lib "gdi32" _
            (EllRect As Rect) _
             As Long

'creates a font compatable with the specified device context
Public Declare Function CreateFontIndirect _
       Lib "gdi32" _
       Alias "CreateFontIndirectA" _
            (lpLogFont As LogFont) _
             As Long

'creates a pen that can be applied to a hDc
Public Declare Function CreatePen _
       Lib "gdi32" _
            (ByVal nPenStyle As Long, _
             ByVal nWidth As Long, _
             ByVal crColor As Long) _
             As Long

'creates a pen that can be applied to a hDc
Public Declare Function CreatePenIndirect _
       Lib "gdi32" _
            (lpLogPen As LogPen) _
             As Long

'creates a rectangular region on a hDc
Public Declare Function CreateRectRgn _
       Lib "gdi32" _
            (Left As Integer, _
             Top As Integer, _
             Right As Integer, _
             Bottom As Integer) _
             As Long

'creates a solid colour brush to be applied to
'a bitmap
Public Declare Function CreateSolidBrush _
       Lib "gdi32" _
            (ColorRef As Long) _
             As Long

'removes a device context from memory
Public Declare Function DeleteDC _
       Lib "gdi32" _
            (ByVal hDc As Long) _
             As Long

'removes an object such as a brush or bitmap from memory
Public Declare Function DeleteObject _
       Lib "gdi32" _
            (ByVal hObject As Long) _
             As Long

'this draws the animated minimize/maximize rectangeles
Public Declare Function DrawAnimatedRects _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal idAni As Long, _
             lprcFrom As Rect, _
             lprcTo As Rect) _
             As Long

'this draws an icon onto a surphase (eg, a bitmap)
Public Declare Function DrawIconEx _
       Lib "user32" _
            (ByVal hDc As Long, _
             ByVal xLeft As Long, _
             ByVal yTop As Long, _
             ByVal hIcon As Long, _
             ByVal cxWidth As Long, _
             ByVal cyWidth As Long, _
             ByVal istepIfAniCur As Long, _
             ByVal hbrFlickerFreeDraw As Long, _
             ByVal diFlags As Long) _
             As Long

'this draws text onto the bitmap
Public Declare Function DrawText _
       Lib "user32" _
       Alias "DrawTextA" _
            (ByVal hDc As Long, _
             ByVal lpStr As String, _
             ByVal nCount As Long, _
             lpRect As Rect, _
             ByVal wFormat As Long) _
             As Long

'this draws an ellipse onto the bitmap
Public Declare Function Ellipse _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Boolean

'this will set the display settings
Public Declare Function EnumDisplaySettings _
       Lib "user32" _
       Alias "EnumDisplaySettingsA" _
            (ByVal A As Long, _
             ByVal B As Long, _
             wef As DEVMODE) _
             As Boolean

'this provides more control than the CreatePen function
Public Declare Function ExtCreatePen _
       Lib "gdi32" _
            (ByVal dwPenStyle As Long, _
             ByVal dwWidth As Long, _
             lplb As LogBrush, _
             ByVal dwStyleCount As Long, _
             lpStyle As Long) _
             As Long

'this will fill the rectangle with the brush applied
'to the hDc
Public Declare Function FillRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             Fill As Rect, _
             hBrush As Long) _
             As Integer

'this will fill a region with the brush specified
Public Declare Function FillRgn _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal HRgn As Long, _
             hBrush As Long) _
             As Boolean

'this will find a window based on its class and
'window name
Public Declare Function FindWindow _
       Lib "user32" _
       Alias "FindWindowA" _
            (ByVal lpClassName As String, _
             ByVal lpWindowName As String) _
             As Long

'this will return the handle of the top-most window
Public Declare Function GetActiveWindow _
       Lib "user32" _
            () _
             As Long

'this will get the state of any specified key - even
'the mouse buttons
Public Declare Function GetAsyncKeyState _
       Lib "user32" _
            (ByVal vKey As Long) _
             As Integer

'this will get the dimensions of the specified bitmap
Public Declare Function GetBitmapDimensionEx _
       Lib "gdi32" _
            (ByVal hBitmap As Long, _
             lpDimension As SizeType) _
             As Long

'this will get the class name from the handle of the
'window specified
Public Declare Function GetClassName _
       Lib "user32" _
       Alias "GetClassNameA" _
            (ByVal hWnd As Long, _
             ByVal lpClassName As String, _
             ByVal nMaxCount As Long) _
             As Long

'this will capture a screen shot of the specified
'area of the client
Public Declare Function GetClientRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             lpRect As Rect) _
             As Long

'this gets the cursors icon picture
Public Declare Function GetCursor _
       Lib "user32" _
            () _
             As Long

'this gets the position of the cursor on the screen
'(given in pixels)
Public Declare Function GetCursorPos _
       Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'gets a hDc of the specified window
Public Declare Function GetDC _
       Lib "user32" _
            (ByVal hWnd As Long) _
            As Long

'gets the entire screen area
Public Declare Function GetDesktopWindow _
       Lib "user32" _
            () _
             As Long

'this will get the current devices' capabilities
Public Declare Function GetDeviceCaps _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal nIndex As Long) _
             As Long

'get the last error to occur from within the api
Public Declare Function GetLastError _
       Lib "kernel32" _
            () _
             As Long

'get the handle of the menu bar on a window
Private Declare Function GetMenu _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'Get the sub menu ID number. This is used to
'reference sub menus along with their handle
Private Declare Function GetMenuItemID _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get information about the specified object such as
'its dimensions etc.
Public Declare Function GetObjectAPI _
       Lib "gdi32" _
       Alias "GetObjectA" _
            (ByVal hObject As Long, _
             ByVal nCount As Long, _
             lpObject As Any) _
             As Long

'get the colour of the specified pixel
Public Declare Function GetPixel _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'this will get the handle of a specified
'sub menu using the handle of the menu
'and the item ID
Private Declare Function GetSubMenu _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get the dimensions of the applied text metrics for
'the device context
Public Declare Function GetTextMetrics _
       Lib "gdi32" _
       Alias "GetTextMetricsA" _
            (ByVal hDc As Long, _
             lpMetrics As TEXTMETRIC) _
             As Long

'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Public Declare Function GetTickCount _
       Lib "kernel32" _
            () _
             As Long  'very usefull timing function ;)

'retrieves the handle of a window that has the
'specified relationship to the specified window.
Public Declare Function GetWindow _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal wCmd As Long) _
             As Long

'gets the area the specified window takes up
Public Declare Function GetWindowRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             lpRect As Rect) _
             As Long

'increases the size of a rect structure
Public Declare Function InflateRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'gets any intersection of two rectangles
Public Declare Function IntersectRect _
       Lib "user32" _
            (lpDestRect As Rect, _
             lpSrc1Rect As Rect, _
             lpSrc2Rect As Rect) _
             As Long

'draws a line from the current position to the
'specified co-ordinates
Public Declare Function LineTo _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             XEnd As Integer, _
             YEnd As Integer) _
             As Boolean

'this will load a cursor into a device context
Public Declare Function LoadCursor _
       Lib "user32" _
       Alias "LoadCursorA" _
            (ByVal hInstance As Long, _
             ByVal lpCursorName As Any) _
             As Long

'this will load an image into a device context
Public Declare Function LoadImage _
       Lib "user32" _
       Alias "LoadImageA" _
            (ByVal hInst As Long, _
             ByVal lpsz As String, _
             ByVal un1 As Long, _
             ByVal n1 As Long, _
             ByVal n2 As Long, _
             ByVal un2 As Long) _
             As Long

'This stops the specified window from updating
'its display. This is mainly used to help cut out
'flicker but does not work on controls.
Public Declare Function LockWindowUpdate _
       Lib "user32" _
            (ByVal hwndLock As Long) _
             As Long
'changes some of a menu's properties
Private Declare Function ModifyMenu _
        Lib "user32" _
        Alias "ModifyMenuA" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal wIDNewItem As Long, _
             ByVal lpString As Any) _
             As Long

'moves the current position to the specified
'point
Public Declare Function MoveToEx _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             X As Integer, _
             Y As Integer, _
             lpPoint As PointAPI) _
             As Boolean

'multiplies two numbers and divides them by a third
'numbers
Public Declare Function MulDiv _
       Lib "kernel32" _
            (ByVal nNumber As Long, _
             ByVal nNumerator As Long, _
             ByVal nDenominator As Long) _
             As Long

'This will increase or decrease a rectangles
'co-ordinates by the specified amount. Usefull
'for moving graphic blocks as rect structures.
Public Declare Function OffsetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'Pattern Blitter. Used to draw a pattern onto
'a device context
Public Declare Function PatBlt _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal dwRop As Long) _
             As Long

'This draws a polygon onto a device context
'useing an array.
Public Declare Function Polygon _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines to the specifed
'points
Public Declare Function Polyline _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines starting from
'the current position on the device context.
Public Declare Function PolylineTo _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lppt As PointAPI, _
             ByVal cCount As Long) _
             As Long

'This draws a rectangle onto the device
'context
Public Declare Function Rectangle _
       Lib "gdi32" _
            (ByVal hWnd As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Long

'This will release a device context from
'memory. Not to be confused with DeleteDC
Public Declare Function ReleaseDC _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal hDc As Long) _
             As Long

'this will draw a round-cornered rectangle
'onto the specified device context
Public Declare Function RoundRect _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal Left As Long, _
             ByVal Top As Long, _
             ByVal Right As Long, _
             ByVal Bottom As Long, _
             ByVal X3 As Long, _
             ByVal Y3 As Long) _
             As Long

'this will select the specified object to
'a window or device context
Public Declare Function SelectObject _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal hObject As Long) _
             As Long

'This sets the background colour on a bitmap
Public Declare Function SetBkColor _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal crColor As Long) _
             As Long

'This sets the background mode on a bitmap
'(eg, transparent, solid etc)
Public Declare Function SetBkMode _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal nBkMode As Long) _
             As Long

'The color adjustment values are used to adjust
'the input color of the source bitmap for calls
'to the StretchBlt and StretchDIBits functions
'when HALFTONE mode is set.
Public Declare Function SetColorAdjustment _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpca As COLORADJUSTMENT) _
             As Long

'sets the colourspace to a device context
Public Declare Function SetColorSpace _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal hcolorspace As Long) _
             As Long

'sets the bitmap of a menu
Private Declare Function SetMenuItemBitmaps _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal hBitmapUnchecked As Long, _
             ByVal hBitmapChecked As Long) _
             As Long

'sets the current information about the selected menu
Private Declare Function SetMenuItemInfo _
        Lib "user32" _
        Alias "SetMenuItemInfoA" _
            (ByVal hMenu As Long, _
             ByVal uItem As Long, _
             ByVal fByPosition As Long, _
             lpmii As MENUITEMINFO) _
             As Long

'sets the colour of the specified pixel
Public Declare Function SetPixel _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal crColor As Long) _
             As Long

'sets the rectangles size and position
Public Declare Function SetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'sets the current text colour
Public Declare Function SetTextColor _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal crColor As Long) _
             As Long

'pauses the execution of the programs thread
'for a specified amount of milliseconds
Public Declare Sub Sleep _
       Lib "kernel32" _
            (ByVal dwMilliseconds As Long)

'used to stretch or shrink a bitmap
Public Declare Function StretchBlt _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal nSrcWidth As Long, _
             ByVal nSrcHeight As Long, _
             ByVal dwRop As Long) _
             As Long


'--------------------------------------------------------------------------
'enumerator section
'--------------------------------------------------------------------------

'the direction of the gradient
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum

'in twips or pixels
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum

'The key values of the mouse buttons
Public Enum MouseKeys
    MouseLeft = 1
    MouseRight = 2
    MouseMiddle = 4
End Enum

'text alignment constants
Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

'bitmap flip constants
Public Enum FlipType
    FlipHorizontally = 0
    FlipVertically = 1
End Enum

'image load constants
Public Enum LoadType
    IMAGE_BITMAP& = 0
End Enum
    
'rotate bitmap constants
Public Enum RotateType
    RotateRight = 0
    RotateLeft = 1
    Rotate180 = 2
End Enum

'--------------------------------------------------------------------------
'Programmer defined data types
'--------------------------------------------------------------------------

'Bitmap structue for menu information
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

'size structure
Public Type SizeType
        cx As Long
        cy As Long
End Type

'Text metrics
Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type


Public Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Public Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Public Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type

Public Type LogColorSpace
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSType As Long
    lcsIntent As Long
    lcsEndPoints As CIEXYZTRIPLE
    lcsGammaRed As Long
    lcsGammaGreen As Long
    lcsGammaBlue As Long
    lcsFileName As String * 26 'MAX_PATH
End Type

'display settings (800x600 etc)
Public Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BitmapStruc
    hDcMemory As Long
    hDcBitmap As Long
    hDcPointer As Long
    Area As Rect
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Type LogPen
        lopnStyle As Long
        lopnWidth As PointAPI
        lopnColor As Long
End Type

Public Type LogBrush
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Type FontStruc
    Name As String
    Alignment As AlignText
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    StrikeThru As Boolean
    PointSize As Byte
    Colour As Long
End Type

Public Type LogFont
    'for the DrawText api call
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type

Public Type Point
    'you'll need this to reference a point on the
    'screen'
    X As Integer
    Y As Integer
End Type

'To hold the RGB value
Public Type RGBVal
    Red As Single
    Green As Single
    Blue As Single
End Type

'bitmap structure for the GetObject api call
Public Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

'--------------------------------------------------------------------------
'Constants section
'--------------------------------------------------------------------------

'general constants
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const WM_USER = &H400

'Image load constants
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const LR_DEFAULTSIZE As Long = &H40

'PatBlt constants
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

'Display constants
Public Const CDS_FULLSCREEN = 4
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000

'DrawText constants
Public Const DT_CENTER = &H1
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'CreateBrushIndirect constants
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_DIBPATTERNPT = 6
Public Const BS_HATCHED = 2
Public Const BS_HOLLOW = 1
Public Const BS_NULL = 1
Public Const BS_PATTERN = 3
Public Const BS_PATTERN8X8 = 7
Public Const BS_SOLID = 0
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_NOSHADE = 17
Public Const HS_SOLID = 8
Public Const HS_SOLIDBKCLR = 23
Public Const HS_SOLIDCLR = 19
Public Const HS_VERTICAL = 1                '  |||||

'BitBlt constants
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)

'LogFont constants
Public Const LF_FACESIZE = 32
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100
Public Const DEFAULT_CHARSET = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_MASK = &HF
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_TT_ALWAYS = 32
Public Const WM_SETFONT = &H30
Public Const LF_FULLFACESIZE = 64
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_QUALITY = 0
Public Const PROOF_QUALITY = 2

'GetDeviceCaps constants
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

'colourspace constants
Public Const MAX_PATH = 260

'pen constants
Public Const PS_COSMETIC = &H0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_DOT = 2                     '  .......
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_GEOMETRIC = &H10000
Public Const PS_INSIDEFRAME = 6
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0
Public Const PS_SOLID = 0

'mouse cursor constants
Public Const IDC_APPSTARTING = 32650&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZE = 32640&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

'menu constants
Private Const MFT_RADIOCHECK = &H200&
Private Const MF_BITMAP = &H4&
Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4

'some key values for GetASyncKeyState
Public Const KLeft = 37
Public Const KUp = 38
Public Const KRight = 39
Public Const KDown = 40

'some mathimatical constants
Public Const Pi = 3.14159265358979

'--------------------------------------------------------------------------
'Variable declarations section
'--------------------------------------------------------------------------

'This stores the various points of a polygon. Use DrawPoly to draw
'the polygon with the points specified in the array.
'Also see AddToPoly, ClearPoly and DelFromPoly
Private PolygonPoints() As PointAPI

'some general variables used by some api calls
Private rctFrom As Rect
Private rctTo As Rect
Private lngTrayHand As Long
Private lngStartMenuHand As Long
Private lngChildHand As Long
Private strClass As String * 255
Private lngClassNameLen As Long
Private lngRetVal As Long
Private blnResChanged As Boolean

'--------------------------------------------------------------------------
'Procedures/functions section
'--------------------------------------------------------------------------

Public Sub BFAlphaBlend(ByVal lngDesthDc As Long, _
                        ByVal lngPic1hDc As Long, _
                        ByVal lngPic2hDc As Long, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        ByVal lngPic1X As Integer, _
                        ByVal lngPic1Y As Integer, _
                        ByVal lngPic2X As Integer, _
                        ByVal lngPic2Y As Integer, _
                        Optional ByVal sngBlendAmount As Single = 0.5, _
                        Optional ByVal lngMaskhDc As Long = 0, _
                        Optional ByVal intMeasurement As Scaling = InPixels)

    'This is a "brute force" alpha blend function. Because it's written in
    'vb, this function is not as fast at it might otherwise be in another
    'language like C++ or Fox.
    'The purpose of the function is to mix the colours of two bitmaps to
    'produce a result that looks like both pictures. Think of it as fading
    'one picture into another. I get the pixel colour of a point in picture1,
    'and the colour of the corresponding pixel in pixture2, find the 'middle'
    'colour and put it into the destination bitmap. There are no calls to
    'other procedures or functions other than api calls. This is to improve
    'speed as all calculations are made internally.
    'The parameter sngBlendAmount MUST be between 1 and 0. If not then
    'the value is rounded to 1 or zero.
    'However, sngBlendAmount is ignored if a Mask bitmap has been specified.
    'Please note that if the mask used only contains black or white pixels,
    'then it is recommended that you use the MergeBitmaps procedure as
    'it will process the bitmaps much faster (by about 15 to 30 times).
    'Keep in mind that using a mask bitmap is about 25% slower than
    'specifying the blend amount.

    Dim TempBmp As BitmapStruc      'a temperory bitmap
    Dim lngResult As Long           'any result returned from an api call
    Dim Col1 As RGBVal              'used to store a pixel colour in RGB format
    Dim Col2 As RGBVal              'used to store a pixel colour in RGB format
    Dim BlendCol As RGBVal          'used to store a pixel colour in RGB format
    Dim MaskCol As RGBVal           'used to store a pixel colour in RGB format
    Dim lngCounterX As Long         'scan the rows of the bitmap
    Dim lngCounterY As Long         'scan the columns of the bitmap
    Dim intPxlHeight As Integer     'the pixel height in twips
    Dim intPxlWidth As Integer      'the pixel width in twips
    Dim lngBlendCol As Long         'the blended colour calculated from the two pixel colours of the bitmaps
    Dim lngCol1 As Long             'used to store a pixel colour in Long format
    Dim lngCol2 As Long             'used to store a pixel colour in Long format
    Dim lngMaskCol As Long          'used to store a pixel colour in Long format
    
    'first convert the passed values if they
    'need to be converted.
    If intMeasurement = InTwips Then
        'get the pixel intHeight and intWidth
        'values per twips in the current
        'screen resolution.
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'start converting the twip values to pixels
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        lngPic1X = lngPic1X / intPxlWidth
        lngPic2X = lngPic2X / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        lngPic1Y = lngPic1Y / intPxlHeight
        lngPic2Y = lngPic2Y / intPxlHeight
    End If
    
    'validate the sngBlendAmount parameter.
    'It must be a values between 0 and
    '1. If the parameter is outside these
    'bounds, then round to nearist
    'bounding value (0 or 1)
    Select Case sngBlendAmount
    Case Is >= 1
        sngBlendAmount = 1
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngPic2hDc, _
                           lngPic2X, _
                           lngPic2Y, _
                           SRCCOPY)
        Exit Sub
    Case Is <= 0
        sngBlendAmount = 0
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngPic1hDc, _
                           lngPic1X, _
                           lngPic1Y, _
                           SRCCOPY)
        Exit Sub
    End Select
    
    'create a temperory destination
    'bitmap
    TempBmp.Area.Right = intWidth
    TempBmp.Area.Bottom = intHeight
    Call CreateNewBitmap(TempBmp.hDcMemory, _
                         TempBmp.hDcBitmap, _
                         TempBmp.hDcPointer, _
                         TempBmp.Area, lngDesthDc)
    
    'start going through the 2 source
    'bitmaps and blending the colours.
    For lngCounterX = 0 To intWidth
        For lngCounterY = 0 To intHeight
            'get the pixel colours
            lngCol1 = GetPixel(lngPic1hDc, _
                               lngPic1X + lngCounterX, _
                               lngPic1Y + lngCounterY)
            lngCol2 = GetPixel(lngPic2hDc, _
                               lngPic2X + lngCounterX, _
                               lngPic2Y + lngCounterY)
            
            'if a blend mask has been specified,
            'then get the blend amount
            'from the bitmap.
            If lngMaskhDc <> 0 Then
                lngMaskCol = GetPixel(lngMaskhDc, _
                                      lngCounterX, _
                                      lngCounterY)
                
                'convert the long value into
                'the blend amount
                MaskCol.Blue = lngMaskCol \ 65536
                MaskCol.Green = ((lngMaskCol - (MaskCol.Blue * 65536)) \ 256)
                MaskCol.Red = (lngMaskCol - (MaskCol.Blue * 65536) - _
                               (MaskCol.Green * 256))
                
                'now convert rgb value to
                'value between 0 and 1
                '(divide by 3 for the average
                'rgb and 255 to a value between
                '1 and 0 (3 * 255 = 765) )
                sngBlendAmount = (MaskCol.Red + MaskCol.Green + MaskCol.Blue) \ 765
            End If
            
            'convert long values to rgb values
            Col1.Blue = lngCol1 \ 65536
            Col1.Green = ((lngCol1 - (Col1.Blue * 65536)) \ 256)
            Col1.Red = (lngCol1 - (Col1.Blue * 65536) - (Col1.Green * 256))
            Col2.Blue = lngCol2 \ 65536
            Col2.Green = ((lngCol2 - (Col2.Blue * 65536)) \ 256)
            Col2.Red = (lngCol2 - (Col2.Blue * 65536) - (Col2.Green * 256))
    
            'average the colours by blend amount
            If (Col1.Red <> Col2.Red) _
                Or (Col1.Green <> Col2.Green) _
                Or (Col1.Blue <> Col2.Blue) _
                Then
                BlendCol.Red = Col1.Red - ((Col1.Red - Col2.Red) * sngBlendAmount)
                BlendCol.Green = Col1.Green - ((Col1.Green - Col2.Green) * sngBlendAmount)
                BlendCol.Blue = Col1.Blue - ((Col1.Blue - Col2.Blue) * sngBlendAmount)
            Else
                'there is no point in blending
                'colours that are the same
                BlendCol = Col1
            End If
            
            'convert the BlendCol RGB values
            'to a long
            lngBlendCol = (CLng(BlendCol.Blue) * 65536) + _
                          (CLng(BlendCol.Green) * 256) + _
                          BlendCol.Red
            
            'set the corresponding pixel colour
            'on the temperory bitmap
            lngResult = SetPixel(TempBmp.hDcMemory, _
                                 lngCounterX, _
                                 lngCounterY, _
                                 lngBlendCol)
        Next lngCounterY
    Next lngCounterX
    
    'copy the blended picture to the
    'destination bitmap
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       intWidth, _
                       intHeight, _
                       TempBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    'remove the temperory bitmap
    'from memory
    Call DeleteBitmap(TempBmp.hDcMemory, _
                      TempBmp.hDcBitmap, _
                      TempBmp.hDcPointer)
End Sub

Public Sub FlipBitmap(ByVal lngDesthDc As Long, _
                      ByVal intDestX As Integer, _
                      ByVal intDestY As Integer, _
                      ByVal intWidth As Integer, _
                      ByVal intHeight As Integer, _
                      ByVal lngSourcehDc As Long, _
                      ByVal intSourceX As Integer, _
                      ByVal intSourceY As Integer, _
                      Optional ByVal Orientation As FlipType = FlipHorizontally, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)

    'This procedure will flip a picture either
    'horizontally or vertically. It copies
    'the bitmap eithre row by row or column
    'by column to improve speed.
    
    Dim bytPxlHeight As Byte
    Dim bytPxlWidth As Byte
    Dim intCounter As Integer
    Dim TempBmp As BitmapStruc
    Dim intFinish As Integer
    Dim lngResult As Long
    
    'convert the twips to pixel values if necessary
    If udtMeasurement = InTwips Then
        bytPxlWidth = Screen.TwipsPerPixelX
        bytPxlHeight = Screen.TwipsPerPixelY
        
        intDestX = intDestX / bytPxlWidth
        intWidth = intWidth / bytPxlWidth
        intSourceX = intSourceX / bytPxlWidth
        intDestY = intDestY / bytPxlHeight
        intHeight = intHeight / bytPxlHeight
        intSourceY = intSourceY / bytPxlHeight
    End If
    
    'create the temperory bitmap
    TempBmp.Area.Right = intWidth
    TempBmp.Area.Bottom = intHeight
    Call CreateNewBitmap(TempBmp.hDcMemory, _
                         TempBmp.hDcBitmap, _
                         TempBmp.hDcPointer, _
                         TempBmp.Area, _
                         lngSourcehDc)
    
    'define the bounds of the loop depending on the orientation (do I scan
    'the bitmap row by row or column by column)
    Select Case Orientation
    Case FlipHorizontally
        'scan column by column
        intFinish = intWidth - 1
    Case FlipVertically
        'scan row by row
        intFinish = intHeight - 1
    End Select
    
    For intCounter = 0 To intFinish
        'copy the row or column into the appropiate section of the bitmap
        If Orientation = FlipHorizontally Then
            'horizontal
            lngResult = BitBlt(TempBmp.hDcMemory, _
                               intFinish - intCounter, _
                               0, _
                               1, _
                               intHeight, _
                               lngSourcehDc, _
                               intSourceX + intCounter, _
                               intSourceY, _
                               SRCCOPY)
        Else
            'flip vertically
            lngResult = BitBlt(TempBmp.hDcMemory, _
                               0, _
                               intFinish - intCounter, _
                               intWidth, _
                               1, _
                               lngSourcehDc, _
                               intSourceX, _
                               intSourceY + intCounter, _
                               SRCCOPY)
        End If
    Next
    
    'copy the flipped bitmap onto the destination bitmap
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
                       
    'delete the temperory bitmap
    Call DeleteBitmap(TempBmp.hDcMemory, _
                      TempBmp.hDcBitmap, _
                      TempBmp.hDcPointer)
End Sub

Public Sub RotateBitmap(ByVal lngDesthDc As Long, _
                        ByVal lngSourcehDc As Long, _
                        ByVal Rotate As RotateType, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intSourceX As Integer, _
                        ByVal intSourceY As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)

    'This procedure will rotate a bitmap 90, 180 or 270 degrees.
    
    Dim lngResult As Long
    Dim intPxlWidth As Integer
    Dim intPxlHeight As Integer
    Dim intCounterX As Integer
    Dim intCounterY As Integer
    Dim TempBmp As BitmapStruc
    Dim lngBitCol As Long
    
    'convert twips values to pixels if necessary
    If udtMeasurement = InTwips Then
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'convert values
        intDestX = intDestX / intPxlWidth
        intSourceX = intSourceX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intSourceY = intSourceY / intPxlHeight
        intHeight = intHeight / intPxlHeight
    End If
    
    'create a temperory bitmap to draw on
    If Rotate = Rotate180 Then
        'the intWidth and intHeight dimensions are the same
        TempBmp.Area.Bottom = intHeight
        TempBmp.Area.Right = intWidth
    Else
        'rotate the dimensions 90 degrees
        TempBmp.Area.Bottom = intWidth
        TempBmp.Area.Right = intHeight
    End If
    Call CreateNewBitmap(TempBmp.hDcMemory, _
                         TempBmp.hDcBitmap, _
                         TempBmp.hDcPointer, _
                         TempBmp.Area, _
                         lngDesthDc)
    
    Select Case Rotate
    Case RotateRight To RotateLeft
        'rotate bitmap right or left
        For intCounterX = 0 To intWidth
            For intCounterY = 0 To intHeight
                'get the pixel colour
                lngBitCol = GetPixel(lngSourcehDc, _
                                     intSourceX + intCounterX, _
                                     intSourceY + intCounterY)
                
                'copy to appropiate part of the temperory bitmap
                If Rotate = RotateRight Then
                    'rotate right
                    lngResult = SetPixel(TempBmp.hDcMemory, _
                                         intHeight - intCounterY, _
                                         intCounterX, _
                                         lngBitCol)
                Else
                    'rotate left
                    lngResult = SetPixel(TempBmp.hDcMemory, _
                                         intCounterY, _
                                         intHeight - intCounterX, _
                                         lngBitCol)
                End If
            Next intCounterY
        Next intCounterX
    
    Case Rotate180
        'rotate bitmap 180 degrees
        
        'we rotate the bitmap 180 degrees by flipping it vertically and
        'horizontally. This is done fastest by calling the FlipBitmap procedure
        Call FlipBitmap(TempBmp.hDcMemory, _
                        0, _
                        0, _
                        intWidth, _
                        intHeight, _
                        lngSourcehDc, _
                        intSourceX, _
                        intSourceY)
        Call FlipBitmap(TempBmp.hDcMemory, _
                        0, _
                        0, _
                        intWidth, _
                        intHeight, _
                        TempBmp.hDcMemory, _
                        0, _
                        0, _
                        FlipVertically)
    End Select
    
    'copy the temperory bitmap to the destination Dc at the specified
    'co-ordinates
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    'remove the temperory bitmap from memory and exit
    Call DeleteBitmap(TempBmp.hDcMemory, _
                      TempBmp.hDcBitmap, _
                      TempBmp.hDcPointer)
End Sub

Public Sub DrawRect(ByVal lngHDC As Long, _
                    ByVal lngColour As Long, _
                    ByVal intLeft As Integer, _
                    ByVal intTop As Integer, _
                    ByVal intRight As Integer, _
                    ByVal intBottom As Integer, _
                    Optional ByVal udtMeasurement As Scaling = InPixels, _
                    Optional ByVal lngStyle As Long = BS_SOLID, _
                    Optional ByVal lngPattern As Long = HS_SOLID)
    
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Dim StartRect As Rect
    Dim lngResult As Long
    Dim lngJunk  As Long
    Dim lnghBrush As Long
    Dim BrushStuff As LogBrush
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intRight = intRight / Screen.TwipsPerPixelX
        intBottom = intBottom / Screen.TwipsPerPixelY
    End If
    
    'initalise values
    StartRect.Top = intTop
    StartRect.Left = intLeft
    StartRect.Bottom = intBottom
    StartRect.Right = intRight
    
    'create a brush
    BrushStuff.lbColor = lngColour
    BrushStuff.lbHatch = lngPattern
    BrushStuff.lbStyle = lngStyle
    
    'apply the brush to the device context
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lngHDC, lnghBrush)
    
    'draw a rectangle
    lngResult = PatBlt(lngHDC, _
                       intLeft, _
                       intTop, _
                       (intRight - intLeft), _
                       (intBottom - intTop), _
                       PATCOPY)
    
    'A "Brush" object was created. It must be removed from memory.
    lngJunk = SelectObject(lngHDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub DrawRoundRect(ByVal lngHDC As Long, _
                         ByVal lngColour As Long, _
                         ByVal intLeft As Integer, _
                         ByVal intTop As Integer, _
                         ByVal intRight As Integer, _
                         ByVal intBottom As Integer, _
                         ByVal intEdgeRadius As Integer, _
                         Optional ByVal udtMeasurement As Scaling = InPixels, _
                         Optional ByVal lngStyle As Long = BS_SOLID, _
                         Optional ByVal lngPattern As Long = HS_SOLID)
                         
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Const Width = 1 'the pixel width of the edge
    
    Dim lnghPen As Long
    Dim PenStuff As LogPen
    Dim lnghBrush As Long
    Dim BrushStuff As LogBrush
    Dim lngJunk  As Long
    Dim lngResult As Long
    Dim OffsetRect As Rect
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelX
        intRight = intRight / Screen.TwipsPerPixelY
        intBottom = intBottom / Screen.TwipsPerPixelY
        intEdgeRadius = intEdgeRadius / Screen.TwipsPerPixelX
    End If
    
    'Find out if a specific lngColour is
    'to be set. If so set it.
    
    'set lnghBrush settings (similar to
    'the FillColor property)
    BrushStuff.lbColor = lngColour
    BrushStuff.lbStyle = lngStyle
    BrushStuff.lbHatch = lngPattern  'ignored if lngStyle is solid
    
    'set lnghPen settings (similar to the
    'border properties on controls)
    PenStuff.lopnColor = lngColour
    PenStuff.lopnWidth.X = Width
    PenStuff.lopnStyle = PS_SOLID
    
    'apply the settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(lngHDC, lnghPen)
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lngHDC, lnghBrush)
    
    '---------in debug - 18/02/02
    'deflate the rectangle dimensions by
    'the radius amount
    'OffsetRect.intLeft = intLeft
    'OffsetRect.intTop = intTop
    'OffsetRect.intRight = intRight
    'OffsetRect.intBottom = intBottom
    'lngResult = InflateRect(OffsetRect, -intEdgeRadius, -intEdgeRadius)
    
    'draw the rounded rectangle
    lngResult = RoundRect(lngHDC, _
                          OffsetRect.Left, _
                          OffsetRect.Top, _
                          OffsetRect.Right, _
                          OffsetRect.Bottom, _
                          intEdgeRadius, _
                          intEdgeRadius)
    
    'delete the objects created (lnghPen and lnghBrush objects)
    lngJunk = SelectObject(lngHDC, lnghPen)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(lngHDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Function TitleToTray(frm As Form)

    'This function will draw the minimize animation from the form to
    'the title tray.
    
    'find the position of the title tray
    lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
    lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
    Do
        lngClassNameLen = GetClassName(lngChildHand, _
                                       strClass, _
                                       Len(strClass))
        If InStr(1, strClass, "TrayNotifyWnd") Then
            lngTrayHand = lngChildHand
            Exit Do
        End If
        lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
    Loop
    
    'animate the title bar to the title tray
    lngRetVal = GetWindowRect(frm.hWnd, rctFrom)
    lngRetVal = GetWindowRect(lngTrayHand, rctTo)
    lngRetVal = DrawAnimatedRects(frm.hWnd, _
                                  IDANI_OPEN Or IDANI_CAPTION, _
                                  rctFrom, _
                                  rctTo)
    
    'hide form
    frm.Visible = False
    frm.Hide
End Function

Public Function TrayToTitle(frm As Form)
    'This function draws the restore animation of the forms title bar, from
    'the system tray to the forms' position.
    
    'find the system trays position
    lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
    lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
    Do
        lngClassNameLen = GetClassName(lngChildHand, _
                                       strClass, _
                                       Len(strClass))
        
        If InStr(1, strClass, "TrayNotifyWnd") Then
            lngTrayHand = lngChildHand
            Exit Do
        End If
        
        lngChildHand = GetWindow(lngChildHand, _
        GW_HWNDNEXT)
    Loop
    
    'draw the animation
    lngRetVal = GetWindowRect(frm.hWnd, rctFrom)
    lngRetVal = GetWindowRect(lngTrayHand, rctTo)
    lngRetVal = DrawAnimatedRects(frm.hWnd, _
                                  IDANI_CLOSE Or IDANI_CAPTION, _
                                  rctTo, _
                                  rctFrom)
    
    'show the window
    frm.Visible = True
    frm.Show
End Function

Public Sub DrawLine(lngHDC As Long, _
                    ByVal intX1 As Integer, _
                    ByVal intY1 As Integer, _
                    ByVal intX2 As Integer, _
                    ByVal intY2 As Integer, _
                    Optional ByVal lngColour As Long = 0, _
                    Optional ByVal intWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InTwips)
                    
    'This will draw a line from point1 to point2
    
    Const NumOfPoints = 2
    
    Dim lngResult As Long
    Dim lnghPen As Long
    Dim PenStuff As LogPen
    Dim Junk  As Long
    Dim Points(NumOfPoints) As PointAPI
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intX1 = intX1 / Screen.TwipsPerPixelX
        intX2 = intX2 / Screen.TwipsPerPixelX
        intY1 = intY1 / Screen.TwipsPerPixelY
        intY2 = intY2 / Screen.TwipsPerPixelY
    End If
    
    'Find out if a specific lngColour is to be set. If so set it.
    PenStuff.lopnColor = lngColour
    PenStuff.lopnStyle = PS_GEOMETRIC
    PenStuff.lopnWidth.X = intWidth
    
    'apply the pen settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(lngHDC, lnghPen)
    
    'set the points
    Points(1).X = intX1
    Points(1).Y = intY1
    Points(2).X = intX2
    Points(2).Y = intY2
    
    'draw the line
    lngResult = Polyline(lngHDC, Points(1), NumOfPoints)
    lngResult = GetLastError
    
    'A "Pen" object was created. It must be removed from memory.
    Junk = SelectObject(lngHDC, lnghPen)
    Junk = DeleteObject(Junk)
End Sub

Public Sub DrawPoly(ByVal lngHDC As Long, _
                    Optional ByVal lngFillColour As Long = 0, _
                    Optional ByVal lngBorderColour As Long = 0, _
                    Optional ByVal intBorderWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This function will draw a polygon of size and colours specified.
    'Modified to draw and fill all the points specified in the array,
    'PolygonPoints(). Use;
    '
    '"ReDim Preserve PolygonPoints(UBound(PolygonPoints) + 1)"
    '
    'To add a point to the array, then specifiy the X and Y values
    'of your point.
    '
    'To delete a point from the array, use the same as above
    'except replace the + sign with a - sign.
    '
    'To clear the array of all data, use;
    '
    '"ReDim PolygonPoints(0)"
    '
    'Also see AddToPoly, DelFromPoly and ClearPoly
    '----------------
    '- Eric  15/12/2001
    '----------------
    
    Dim lngSuccessful As Long
    Dim intCounter As Integer
    Dim Temp() As PointAPI
    Dim intPointNum As Integer
    Dim lnghBrush As Long
    Dim BrushStuff As LogBrush
    Dim lnghPen As Long
    Dim PenStuff As LogPen
    
    'find the number of points stored
    intPointNum = UBound(PolygonPoints)
    
    'convert all points to pixels if necessary, otherwise, don't waste time.
    If udtMeasurement = InTwips Then
        'create a temperory array to hold the points
        ReDim Temp(intPointNum)
        
        'convert to pixels
        For intCounter = 0 To intPointNum
            Temp(intCounter).X = PolygonPoints(intCounter).X / Screen.TwipsPerPixelX
            Temp(intCounter).Y = PolygonPoints(intCounter).Y / Screen.TwipsPerPixelY
        Next intCounter
    End If
    
    'apply the border and background colours
    BrushStuff.lbColor = lngFillColour
    BrushStuff.lbHatch = 0
    BrushStuff.lbStyle = BS_SOLID
    
    PenStuff.lopnColor = lngBorderColour
    PenStuff.lopnWidth.X = intBorderWidth
    PenStuff.lopnStyle = PS_SOLID
    
    'create the objects necessary to apply the colour and draw settings
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghPen = CreatePenIndirect(PenStuff)
    lngSuccessful = SelectObject(lngHDC, lnghBrush)
    lngSuccessful = SelectObject(lngHDC, lnghPen)
    
    'draw the polygon
    If udtMeasurement = InTwips Then
        lngSuccessful = Polygon(lngHDC, Temp(0), intPointNum)
    Else
        lngSuccessful = Polygon(lngHDC, PolygonPoints(0), intPointNum)
    End If
    
    'remove the pen and brush objects
    lngSuccessful = SelectObject(lngHDC, lnghBrush)
    lngSuccessful = DeleteObject(lngSuccessful)
    lngSuccessful = SelectObject(lngHDC, lnghPen)
    lngSuccessful = DeleteObject(lngSuccessful)
End Sub

Public Sub ClearPoly()
    'This will reset the polgon array to contain no values, with no elements
    '(except element zero by default).
    'Also see DrawPoly, AddToPoly and DelFromPoly
    
    ReDim PolygonPoints(0, 0)
    
    'find the amount of points in the second array dimension
    'clear the points
    PolygonPoints(0).X = 0
    PolygonPoints(0).Y = 0
End Sub

Public Sub AddToPoly(ByVal intX As Integer, _
                     ByVal intY As Integer, _
                     Optional ByVal Point As Integer = -1)
                     
    'This procedure will add a point to the polygon at the specified
    'point. if no point os specified (point -1), then the point is added to
    'the end of the polygon array. It is best to call ClearPoly before using
    'AddToPoly for the first time, but it is not necessary as the appropiate
    'check is made here.
    'Also see DrawPoly, DelFromPoly and ClearPoly
    
    Dim Counter As Integer
    Dim MaintXPoints As Integer
    Dim ErrNum  As Integer
    
    'check for error 9 "subscript out of range". This means that the array
    'has not been initalized yet (there are no elements in the array).
    On Error Resume Next
    MaintXPoints = UBound(PolygonPoints)
    
    'trap error
    ErrNum = Err
    
    'resumt normal error handling
    On Error GoTo 0
    
    'set the array if it has not been initialized yet.
    If ErrNum <> 0 Then
        Call ClearPoly
    End If
    
    'eintXpand the array by one
    ReDim Preserve PolygonPoints(MaintXPoints + 1)
    MaintXPoints = MaintXPoints + 1
    
    
    'if no insert point was specified, then add the point to the end of
    'the current maintX array element number
    If Point < 0 Then
        'enter the new values
        PolygonPoints(MaintXPoints).X = intX
        PolygonPoints(MaintXPoints).Y = intY
    Else
        'insert the point into the array at the specified point, moving any
        'consecutive points "up".
        
        'move consecutive values up
        For Counter = Point To (MaintXPoints - 1)
            PolygonPoints(Counter + 1) = PolygonPoints(Counter)
        Next Counter
        
        'enter the new values
        PolygonPoints(Point).X = intX
        PolygonPoints(Point).Y = intY
    End If
End Sub

Public Sub DelFromPoly(Optional ByVal intPoint As Integer = -1)
    'This will remove the specified Point from the polygon array. if
    'no Point was specified (intPoint -1), then the procedure will remove
    'the last array position
    'Also see DrawPoly, AddToPoly and ClearPoly
    
    Dim intCounter As Integer
    Dim intMaxPoints As Integer
    
    'find the number of points
    intMaxPoints = UBound(PolygonPoints)
    
    'if there is only one intPoint, then clear the array and exit
    If intMaxPoints < 2 Then
        Call ClearPoly
        Exit Sub
    End If
    
    'move consecutive points in the array "down" on top of the specified
    'intPoint
    If intPoint >= 0 Then
        For intCounter = intPoint To (intMaxPoints - 1)
            PolygonPoints(intCounter) = PolygonPoints(intCounter + 1)
        Next intCounter
    End If
    
    'shrink the array size by one
    ReDim Preserve PolygonPoints(intMaxPoints - 1)
End Sub

Public Sub LockWindow(FormName As Form)
    'Prevent the form from updating its display
    
    Dim lngResult As Boolean
    
    lngResult = LockWindowUpdate(FormName.hWnd)
End Sub

Public Sub UnLockWindow()
    'see the procedure LockWindow
    
    Dim lngResult As Long
    
    'Let the form update its display
    lngResult = LockWindowUpdate(0)
End Sub

Public Sub CreateNewBitmap(ByRef hDcMemory As Long, _
                           ByRef hDcBitmap As Long, _
                           ByRef hDcPointer As Long, _
                           ByRef BmpArea As Rect, _
                           ByVal CompatableWithhDc As Long, _
                           Optional ByVal lngBackColour As Long = 0, _
                           Optional ByVal udtMeasurement As Scaling = InPixels)
                           
    'This procedure will create a new bitmap compatable with a given
    'form (you will also be able to then use this bitmap in a picturebox).
    'The space specified in "Area" should be in "Twips" and will be
    'converted into pixels in the following code.
    
    Dim lngResult As Long
    Dim Area As Rect
    
    'scale the bitmap points if necessary
    Area = BmpArea
    If udtMeasurement = InTwips Then
        Call RectToPixels(Area)
    End If
    
    'create the bitmap and its references
    hDcMemory = CreateCompatibleDC(CompatableWithhDc)
    hDcBitmap = CreateCompatibleBitmap(CompatableWithhDc, _
                                       (Area.Right - Area.Left), _
                                       (Area.Bottom - Area.Top))
    hDcPointer = SelectObject(hDcMemory, hDcBitmap)
    
    'set default colours and clear bitmap to selected colour
    lngResult = SetBkMode(hDcMemory, OPAQUE)
    lngResult = SetBkColor(hDcMemory, lngBackColour)
    Call DrawRect(hDcMemory, _
                  lngBackColour, _
                  0, _
                  0, _
                  (Area.Right - Area.Left), _
                  (Area.Bottom - Area.Top))
End Sub

Public Sub DeleteBitmap(ByRef hDcMemory As Long, _
                        ByRef hDcBitmap As Long, _
                        ByRef hDcPointer As Long)
                        
    'This will remove the bitmap that stored what was displayed before
    'the text was written to the screen, from memory.
    
    Dim lngJunk As Long
    
    If hDcMemory = 0 Then
        'there is nothing to delete. Exit the sub-routine
        Exit Sub
    End If
    
    'delete the device context
    lngJunk = SelectObject(hDcMemory, hDcPointer)
    lngJunk = DeleteObject(hDcBitmap)
    lngJunk = DeleteDC(hDcMemory)
    
    'show that the device context has been deleted by setting
    'all parameters passed to the procedure to zero
    hDcMemory = 0
    hDcBitmap = 0
    hDcPointer = 0
End Sub

Public Sub MergeBitmaps(ByVal hDcDest As Long, _
                        ByVal hDcTextureBack As Long, _
                        ByVal hDcTextureFore As Long, _
                        ByVal hDcMask As Long, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intTextureBackX As Integer, _
                        ByVal intTextureBackY As Integer, _
                        ByVal intTextureForeX As Integer, _
                        ByVal intTextureForeY As Integer, _
                        ByVal intMaskWidth As Integer, _
                        ByVal intMaskHeight As Integer, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)
                        
    'This procedure takes a monochrome (black & white) bitmap as a mask,
    '2 source texture bitmaps and a destination bitmap. It copies to the
    'destination bitmap a merge of the two source bitmaps, filling in the
    'pixels according to the mask. For each white pixel in the mask, it copies
    'the corresponding pixel from TextureBack. For each black pixel in the mask,
    'it copies the corresponding pixel from TextureFore to the destination
    'bitmap.
    'eg. Say the mask picture was a black square with a white letter "A" on
    'it. The first texture was of clouds and the second picture was of a
    'tartan design. The lngResult would be a cloud picture in the shape of an
    '"A" on a tartan background.
    
    Dim lngResult As Long
    Dim TempMaskBmp As BitmapStruc
    Dim TempBackBmp As BitmapStruc
    Dim intConvertX As Integer
    Dim intConvertY As Integer
    
    'convert passed values if necessary
    If udtMeasurement = InTwips Then
        intConvertX = Screen.TwipsPerPixelX
        intConvertY = Screen.TwipsPerPixelY
        
        intDestX = intDestX / intConvertX
        intDestY = intDestY / intConvertY
        intTextureBackX = intTextureBackX / intConvertX
        intTextureBackY = intTextureBackY / intConvertY
        intTextureForeX = intTextureForeX / intConvertX
        intTextureForeY = intTextureForeY / intConvertY
        intMaskWidth = intMaskWidth / intConvertX
        intMaskHeight = intMaskHeight / intConvertY
    End If
    
    'create temperory bitmaps
    TempMaskBmp.Area.Right = intMaskWidth
    TempMaskBmp.Area.Bottom = intMaskHeight
    TempBackBmp.Area = TempMaskBmp.Area
    Call CreateNewBitmap(TempMaskBmp.hDcMemory, _
                         TempMaskBmp.hDcBitmap, _
                         TempMaskBmp.hDcPointer, _
                         TempMaskBmp.Area, _
                         hDcMask, _
                         vbWhite, _
                         InPixels)
    Call CreateNewBitmap(TempBackBmp.hDcMemory, _
                         TempBackBmp.hDcBitmap, _
                         TempBackBmp.hDcPointer, _
                         TempBackBmp.Area, _
                         hDcMask, _
                         vbWhite, _
                         InPixels)
    
    'create a white bitmap with a mask shaped hole onto the
    'destination background
    lngResult = BitBlt(TempMaskBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureFore, _
                       intTextureForeX, _
                       intTextureForeY, _
                       SRCCOPY)
    lngResult = BitBlt(TempBackBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcMask, _
                       0, _
                       0, _
                       SRCINVERT)
    lngResult = BitBlt(TempMaskBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.hDcMemory, _
                       0, _
                       0, _
                       MERGEPAINT)
    
    'draw a white mask shape onto the second texture
    'where the black mask is
    lngResult = BitBlt(TempBackBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureBack, _
                       intTextureBackX, _
                       intTextureBackY, _
                       SRCCOPY)
    lngResult = BitBlt(TempBackBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcMask, _
                       0, _
                       0, _
                       MERGEPAINT)
    
    'merge the two masks
    lngResult = BitBlt(TempMaskBmp.hDcMemory, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCAND)
    
    'copy the merged lngResult to the destination
    lngResult = BitBlt(hDcDest, _
                       intDestX, _
                       intDestY, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempMaskBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    'remove the two temperory bitmaps from memory
    Call DeleteBitmap(TempMaskBmp.hDcMemory, _
                      TempMaskBmp.hDcBitmap, _
                      TempMaskBmp.hDcPointer)
    Call DeleteBitmap(TempBackBmp.hDcMemory, _
                      TempBackBmp.hDcBitmap, _
                      TempBackBmp.hDcPointer)
End Sub

Public Sub RectToTwips(ByRef TheRect As Rect)
    'converts pixels to twips in a rect structure
    
    TheRect.Left = TheRect.Left * Screen.TwipsPerPixelX
    TheRect.Right = TheRect.Right * Screen.TwipsPerPixelX
    TheRect.Top = TheRect.Top * Screen.TwipsPerPixelY
    TheRect.Bottom = TheRect.Bottom * Screen.TwipsPerPixelY
End Sub

Public Sub RectToPixels(ByRef TheRect As Rect)
    'converts twips to pixels in a rect structure
    
    TheRect.Left = TheRect.Left \ Screen.TwipsPerPixelX
    TheRect.Right = TheRect.Right \ Screen.TwipsPerPixelX
    TheRect.Top = TheRect.Top \ Screen.TwipsPerPixelY
    TheRect.Bottom = TheRect.Bottom \ Screen.TwipsPerPixelY
End Sub

Public Function AmIActive(TheForm As Form) As Boolean
    'This function returns wether or not the window is active
    
    If TheForm.hWnd = GetActiveWindow Then
        AmIActive = True
    Else
        AmIActive = False
    End If
End Function

Public Sub Gradient(ByVal lngDesthDc As Long, _
                    ByVal lngStartCol As Long, _
                    ByVal FinishCol As Long, _
                    ByVal intLeft As Integer, _
                    ByVal intTop As Integer, _
                    ByVal intWidth As Integer, _
                    ByVal intHeight As Integer, _
                    ByVal Direction As GradientTo, _
                    Optional ByVal udtMeasurement As Scaling = 1, _
                    Optional ByVal bytLineWidth As Byte = 1)
                    
    'draws a gradient from colour mblnStart to colour Finish, and assums
    'that all measurments passed to it are in pixels unless otherwise
    'specified.
    
    Dim intCounter As Integer
    Dim intBiggestDiff As Integer
    Dim Colour As RGBVal
    Dim mblnStart As RGBVal
    Dim Finish As RGBVal
    Dim sngAddRed As Single
    Dim sngAddGreen As Single
    Dim sngAddBlue As Single
    
    'perform all necessary calculations before drawing gradient
    'such as converting long to rgb values, and getting the correct
    'scaling for the bitmap.
    mblnStart = GetRGB(lngStartCol)
    Finish = GetRGB(FinishCol)
    
    If udtMeasurement = InTwips Then
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intWidth = intWidth / Screen.TwipsPerPixelX
        intHeight = intHeight / Screen.TwipsPerPixelY
    End If
    
    'draw the colour gradient
    Select Case Direction
    Case GradVertical
        intBiggestDiff = intWidth
    Case GradHorizontal
        intBiggestDiff = intHeight
    End Select
    
    'calculate how much to increment/decrement each colour per step
    sngAddRed = (bytLineWidth * ((Finish.Red) - mblnStart.Red) / intBiggestDiff)
    sngAddGreen = (bytLineWidth * ((Finish.Green) - mblnStart.Green) / intBiggestDiff)
    sngAddBlue = (bytLineWidth * ((Finish.Blue) - mblnStart.Blue) / intBiggestDiff)
    Colour = mblnStart
    
    'calculate the colour of each line before drawing it on the bitmap
    For intCounter = 0 To intBiggestDiff Step bytLineWidth
        'find the point between colour mblnStart and Colour Finish that
        'corresponds to the point between 0 and intBiggestDiff
        
        'check for overflow
        If Colour.Red > 255 Then
            Colour.Red = 255
        Else
            If Colour.Red < 0 Then
                Colour.Red = 0
            End If
        End If
        If Colour.Green > 255 Then
            Colour.Green = 255
        Else
            If Colour.Green < 0 Then
                Colour.Green = 0
            End If
        End If
        If Colour.Blue > 255 Then
            Colour.Blue = 255
        Else
            If Colour.Blue < 0 Then
                Colour.Blue = 0
            End If
        End If
        
        'draw the gradient in the proper orientation in the calculated colour
        Select Case Direction
        Case GradVertical
            Call DrawLine(lngDesthDc, _
                          intCounter + intLeft, _
                          intTop, _
                          intCounter + intLeft, _
                          intHeight + intTop, _
                          RGB(Colour.Red, Colour.Green, Colour.Blue), _
                          bytLineWidth, _
                          InPixels)
        Case GradHorizontal
            Call DrawLine(lngDesthDc, _
                          intLeft, _
                          intCounter + intTop, _
                          intLeft + intWidth, _
                          intTop + intCounter, _
                          RGB(Colour.Red, Colour.Green, Colour.Blue), _
                          bytLineWidth, _
                          InPixels)
        End Select
        
        'set next colour
        Colour.Red = Colour.Red + sngAddRed
        Colour.Green = Colour.Green + sngAddGreen
        Colour.Blue = Colour.Blue + sngAddBlue
    Next intCounter
End Sub

Public Sub FadeGradient(ByVal lngDesthDc As Long, _
                        ByVal intDestintLeft As Integer, _
                        ByVal intDestinttop As Integer, _
                        ByVal intDestWidth As Integer, _
                        ByVal intDestHeight As Integer, _
                        ByVal lngGradhDc As Long, _
                        ByVal lngStartFromA As Long, _
                        ByVal lngFinishToA As Long, _
                        ByVal lngStartFromB As Long, _
                        ByVal lngFinishToB As Long, _
                        ByVal intLeft As Integer, _
                        ByVal intTop As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        ByVal udtDirection As GradientTo, _
                        Optional ByVal udtMesurement As Scaling = 1, _
                        Optional ByVal bytLineWidth As Byte = 1)
                        
    'This procedure will call the Gradient function to fade it into
    'the udtColours specified.
    'Note : all mesurements must me of the same scale, ie they must all
    'be in pixels or all in twips.
    
    Dim udtColour(2) As RGBVal
    Dim udtStart(2) As RGBVal
    Dim udtFinish(2) As RGBVal
    Dim lngGradCol(2) As Long
    Dim intCounter As Integer
    Dim intBiggestDiff As Integer
    Dim intValue As Integer
    Dim intIndex As Integer
    Dim lngResult As Long
    
    Const A = 0
    Const B = 1
    
    'convert to RGB values
    udtStart(A) = GetRGB(lngStartFromA)
    udtStart(B) = GetRGB(lngStartFromB)
    udtFinish(A) = GetRGB(lngFinishToA)
    udtFinish(B) = GetRGB(lngFinishToB)
    
    'convert to pixels if necessary
    If udtMesurement = InTwips Then
        intDestintLeft = intDestintLeft / Screen.TwipsPerPixelX
        intDestinttop = intDestinttop / Screen.TwipsPerPixelY
        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intWidth = intWidth / Screen.TwipsPerPixelX
        intHeight = intHeight / Screen.TwipsPerPixelY
    End If
    
    
    'Find the largest difference between any two corresponding
    'udtColours, and use that as the number of steps to take in the loop,
    '(therefore guarenteing that it will cycle through all necessary
    'udtColours without jumping)
    For intIndex = A To B
        'test red
        intValue = PositVal(udtStart(intIndex).Red - udtFinish(intIndex).Red)
        If intValue > intBiggestDiff Then
            intBiggestDiff = intValue
        End If
        
        'test green
        intValue = PositVal(udtStart(intIndex).Green - udtFinish(intIndex).Green)
        If intValue > intBiggestDiff Then
            intBiggestDiff = intValue
        End If
        
        'test blue
        intValue = PositVal(udtStart(intIndex).Blue - udtFinish(intIndex).Blue)
        If intValue > intBiggestDiff Then
            intBiggestDiff = intValue
        End If
    Next intIndex
    
    'if there is no difference, then just draw one gradient and exit
    If intBiggestDiff = 0 Then
        Call Gradient(lngGradhDc, _
                      lngStartFromA, _
                      lngStartFromB, _
                      intLeft, _
                      intTop, _
                      intWidth, _
                      intHeight, _
                      udtDirection, _
                      InPixels, _
                      bytLineWidth)
        Exit Sub
    End If
    
    'fade the gradient
    For intCounter = 0 To intBiggestDiff
        'find the point between udtColour udtStart and udtColour udtFinish that
        'corresponds to the point between 0 and intBiggestDiff
        
        For intIndex = A To B
            udtColour(intIndex).Red = udtStart(intIndex).Red + (((udtFinish(intIndex).Red - udtStart(intIndex).Red) / intBiggestDiff) * intCounter)
            udtColour(intIndex).Green = udtStart(intIndex).Green + (((udtFinish(intIndex).Green - udtStart(intIndex).Green) / intBiggestDiff) * intCounter)
            udtColour(intIndex).Blue = udtStart(intIndex).Blue + (((udtFinish(intIndex).Blue - udtStart(intIndex).Blue) / intBiggestDiff) * intCounter)
        
            'convert to long intValue and store
            lngGradCol(intIndex) = RGB(udtColour(intIndex).Red, _
                                       udtColour(intIndex).Green, _
                                       udtColour(intIndex).Blue)
        Next intIndex
        
        'draw the gradient onto the bitmap
        Call Gradient(lngGradhDc, _
                      lngGradCol(A), _
                      lngGradCol(B), _
                      intLeft, _
                      intTop, _
                      intWidth, _
                      intHeight, _
                      udtDirection, _
                      InPixels, _
                      bytLineWidth)
    
        'blitt the bitmap to the screen
        lngResult = BitBlt(lngDesthDc, _
                           intDestintLeft, _
                           intDestinttop, _
                           intDestWidth, _
                           intDestHeight, _
                           lngGradhDc, _
                           intLeft, _
                           intTop, _
                           SRCCOPY)
        DoEvents
    Next intCounter
End Sub

Public Function PositVal(intValue As Integer) _
                         As Integer
    '-obsolete. Use the Abs function - Eric, 6/12/2001
    'Returns the positive value of a number
    
    'PosVal = Sqr(Value ^ 2)
    PositVal = Abs(intValue)
End Function

Public Function FromRGB(ByVal bytRed As Byte, _
                        ByVal bytGreen As Byte, _
                        ByVal bytBlue As Byte) _
                        As Long
    'Convert RGB to Long
     
    Dim lngMyVal As Long
    
    lngMyVal = (CLng(bytBlue) * 65536) + (CLng(bytGreen) * 256) + bytRed
    FromRGB = lngMyVal
End Function

Public Function GetRGB(ByVal lngColour As Long) _
                       As RGBVal
    'Convert Long to RGB:
    
    'if the lngcolour value is greater than acceptable then half the value
    If (lngColour > RGB(255, 255, 255)) Or (lngColour < (RGB(255, 255, 255) * -1)) Then
        Exit Function
    End If
    
    GetRGB.Blue = (lngColour \ 65536)
    GetRGB.Green = ((lngColour - (GetRGB.Blue * 65536)) \ 256)
    GetRGB.Red = (lngColour - (GetRGB.Blue * (65536)) - ((GetRGB.Green) * 256))
End Function

Public Sub Pause(lngTicks As Long)
    'pause execution of the program for a specified number of lngTicks
    
    If lngTicks < 0 Then
        lngTicks = 0
    End If
    Call Sleep(lngTicks)
End Sub

Public Sub DrawEllipse(ByVal hDc As Long, _
                       ByVal intCenterX As Integer, _
                       ByVal intCenterY As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal intWidth As Integer, _
                       Optional ByVal intTileAngle As Integer = 90, _
                       Optional ByVal lngColour As Long = 0, _
                       Optional intThickness As Integer = 1, _
                       Optional blnIsHollow As Boolean = True, _
                       Optional udtMesurement As Scaling = InPixels)
                       
    'This procedure will draw an ellipse of the specified dimensions and
    'lngColour, by drawing a line between each of the 360 points that make
    'up the ellipse.
    
    Const A = 1
    Const B = 2
    
    Dim sngMovecCenterX As Single
    Dim sngMovecCenterY As Single
    Dim sngCircleX As Single
    Dim sngCircleY As Single
    Dim sngCounter As Single
    Dim sngTiltX As Single
    Dim sngTiltY As Single
    Dim sngNumOfPoints As Single
    Dim udtEllipse() As PointAPI
    Dim udtBrushStuff As LogBrush
    Dim lnghBrush As Long
    Dim udtPenStuff As LogPen
    Dim lnghPen As Long
    Dim lngResult As Long
    Dim lngJunk As Long
    Dim sngDegPerPoint As Single
    
    'set scaling values
    If udtMesurement = InTwips Then
        'convert parameters to pixels
        intCenterX = (intCenterX / Screen.TwipsPerPixelX) '- intThickness
        intCenterY = (intCenterY / Screen.TwipsPerPixelY) '- intThickness
        intHeight = (intHeight / Screen.TwipsPerPixelY) - (intThickness * 2)
        intWidth = (intWidth / Screen.TwipsPerPixelX) - (intThickness * 2)
        
        'values are now in pixels
        udtMesurement = InPixels
    End If
    
    'calculate the radius for intWidth and intHeight
    intHeight = intHeight / 2
    intWidth = (intWidth / 2) - intHeight
    
    'calculate the starting point of the ellipse
    sngTiltX = Sin(intTileAngle * Pi / 180) * intWidth
    sngTiltY = Cos(intTileAngle * Pi / 180) * intWidth
    
    'draw the ellipse using one line for every three pixels
    'This will increase drawing speed on small ellipses and produce
    'detailed ones for large ellipses.
    sngNumOfPoints = (Pi * (intWidth + intHeight)) / 3  '2.Pi.r = circumfrence, /3=per 3 pixels
    
    
    'size the array to match the number of points to be calculated
    ReDim udtEllipse(sngNumOfPoints)
    
    'calculate the number of degrees between points
    sngDegPerPoint = (360 / sngNumOfPoints)
    
    For sngCounter = 0 To 360 Step sngDegPerPoint
        sngMovecCenterX = intCenterX + (Cos(sngCounter * Pi / 180) * sngTiltX)
        sngMovecCenterY = intCenterY + (Cos(sngCounter * Pi / 180) * sngTiltY)
        
        'calculate the new position
        sngCircleX = Sin((sngCounter + intTileAngle) * Pi / 180) * intHeight
        sngCircleY = Cos((sngCounter + intTileAngle) * Pi / 180) * intHeight
        
        'add the points
        udtEllipse(sngCounter / sngDegPerPoint).X = sngMovecCenterX + sngCircleX
        udtEllipse(sngCounter / sngDegPerPoint).Y = sngMovecCenterY + sngCircleY
    Next sngCounter
    
    'draw the ellipse as a polygon
    
    'first create a brush and pen to display the colours
    udtBrushStuff.lbColor = lngColour
    If blnIsHollow Then
        udtBrushStuff.lbHatch = BS_HOLLOW
    Else
        udtBrushStuff.lbHatch = BS_SOLID
    End If
    udtBrushStuff.lbStyle = 0
    lnghBrush = CreateBrushIndirect(udtBrushStuff)
    
    'apply brush
    lngResult = SelectObject(hDc, lnghBrush)
    
    'create the pen
    udtPenStuff.lopnColor = lngColour
    udtPenStuff.lopnWidth.X = intThickness
    udtPenStuff.lopnStyle = PS_SOLID And PS_INSIDEFRAME
    lnghPen = CreatePenIndirect(udtPenStuff)
    
    'apply pen
    lngResult = SelectObject(hDc, lnghPen)
    
    'now draw the ellipse onto the hDc
    lngResult = Polygon(hDc, udtEllipse(0), sngNumOfPoints)
    
    'delete the brush and pen objects from memory
    lngJunk = SelectObject(hDc, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(hDc, lnghPen)
    lngJunk = DeleteObject(lnghPen)
End Sub

Public Function GetAngle(ByVal sngX1 As Single, _
                         ByVal sngY1 As Single, _
                         ByVal sngX2 As Single, _
                         ByVal sngY2 As Single) _
                         As Integer
                         
    'returns the angle of point1 in relation to point2
    
    Dim sngTempAngle As Single
    
    'if the values are not over the center point, then calculate the angle
    If ((sngY1 - sngY2) <> 0) Or ((sngX1 - sngX2) <> 0) Then
        
        sngTempAngle = (Atn(Slope(sngX1, sngY1, sngX2, sngY2)) * 180 / Pi) Mod 360
        'Debug.Print sngTempAngle, sngX1, sngX2, sngY1, sngY2
        If sngTempAngle > 0 Then
            sngTempAngle = 90 - sngTempAngle
        Else
            sngTempAngle = Abs(sngTempAngle) + 90
        End If
        If sngX1 < sngX2 Then
            sngTempAngle = sngTempAngle + 180
        End If
        
        GetAngle = sngTempAngle
    End If
End Function

Public Function Slope(ByVal intX1 As Integer, _
                      ByVal intY1 As Integer, _
                      ByVal intX2 As Integer, _
                      ByVal intY2 As Integer) _
                      As Single
                      
    'This function finds the slope of a line, where the slope, m =
    '       intX1 - inty1
    'm = ------------
    '       intX2 - intY2
    
    Dim intXVal As Integer
    Dim intYVal As Integer
    
    intXVal = intX2 - intX1
    intYVal = intY2 - intY1
    If (intXVal = 0) And (intYVal = 0) Then
        'if both values were zero, then
        Slope = 0
        Exit Function
    Else
        'if only one value was zero then
        If (intXVal = 0) Or (intYVal = 0) Then
            'the slope = the other value
            Select Case 0
            Case intXVal
                Slope = intXVal
            Case intYVal
                Slope = intYVal
            End Select
            Exit Function
        End If
    End If
    
    If (intXVal <> 0) And (intYVal <> 0) Then
        'otherwise the slope = the formula
        Slope = (intY2 - intY1) / (intX2 - intX1)
    End If
End Function

Public Sub MakeText(ByVal hDcSurphase As Long, _
                    ByVal strText As String, _
                    ByVal intTop As Integer, _
                    ByVal intLeft As Integer, _
                    ByVal intHeight As Integer, _
                    ByVal intWidth As Integer, _
                    ByRef udtFont As FontStruc, _
                    Optional ByVal udtMeasurement As Scaling = 0)
                    
    'This procedure will draw strText onto the bitmap in the specified udtFont,
    'colour and position.
    
    Dim udtAPIFont As LogFont
    Dim lngAlignment As Long
    Dim udtTextRect As Rect
    Dim lngResult As Long
    Dim lngJunk As Long
    Dim hDcFont As Long
    Dim hDcOldFont As Long
    Dim intCounter As Integer
    
    'set Measurement values
    udtTextRect.Top = intTop
    udtTextRect.Left = intLeft
    udtTextRect.Right = intLeft + intWidth
    udtTextRect.Bottom = intTop + intHeight
    
    If udtMeasurement = InTwips Then
        'convert to pixels
        Call RectToPixels(udtTextRect)
    End If
    
    'Create details about the udtFont using the udtFont structure
    '====================
    
    'convert point size to pixels
    udtAPIFont.lfHeight = -((udtFont.PointSize * GetDeviceCaps(hDcSurphase, LOGPIXELSY)) / 72)
    udtAPIFont.lfCharSet = DEFAULT_CHARSET
    udtAPIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    udtAPIFont.lfEscapement = 0
    
    'move the name of the udtFont into the array
    For intCounter = 1 To Len(udtFont.Name)
        udtAPIFont.lfFaceName(intCounter) = Asc(Mid(udtFont.Name, intCounter, 1))
    Next intCounter
    'this has to be a Null terminated string
    udtAPIFont.lfFaceName(intCounter) = 0
    
    udtAPIFont.lfItalic = udtFont.Italic
    udtAPIFont.lfUnderline = udtFont.Underline
    udtAPIFont.lfStrikeOut = udtFont.StrikeThru
    udtAPIFont.lfOrientation = 0
    udtAPIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    udtAPIFont.lfPitchAndFamily = DEFAULT_PITCH
    udtAPIFont.lfQuality = PROOF_QUALITY
    
    If udtFont.Bold Then
        udtAPIFont.lfWeight = FW_BOLD
    Else
        udtAPIFont.lfWeight = FW_NORMAL
    End If
    
    udtAPIFont.lfWidth = 0
    hDcFont = CreateFontIndirect(udtAPIFont)
    hDcOldFont = SelectObject(hDcSurphase, hDcFont)
    '====================
    
    Select Case udtFont.Alignment
    Case vbLeftAlign
        lngAlignment = DT_LEFT
    Case vbCentreAlign
        lngAlignment = DT_CENTER
    Case vbRightAlign
        lngAlignment = DT_RIGHT
    End Select
    
    'Draw the strText into the off-screen bitmap before copying the
    'new bitmap (with the strText) onto the screen.
    lngResult = SetBkMode(hDcSurphase, TRANSPARENT)
    lngResult = SetTextColor(hDcSurphase, udtFont.Colour)
    lngResult = DrawText(hDcSurphase, _
                         strText, _
                         Len(strText), _
                         udtTextRect, _
                         lngAlignment)
    
    'clean up by deleting the off-screen bitmap and udtFont
    lngJunk = SelectObject(hDcSurphase, hDcOldFont)
    lngJunk = DeleteObject(hDcFont)
End Sub

Public Function GetTextHeight(ByVal hDc As Long) _
                              As Integer
    'This function will return the height of the text using the point size
    
    Dim udtMetrics As TEXTMETRIC
    Dim lngResult As Long
    
    lngResult = GetTextMetrics(hDc, udtMetrics)
    
    GetTextHeight = udtMetrics.tmHeight
End Function

Public Sub GetScreenRes(ByRef intWidth As Integer, _
                        ByRef intHeight As Integer)
    'This procedure sets the variable to the current screen dimensions.
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX
    intHeight = Screen.Height / Screen.TwipsPerPixelY
End Sub

Public Sub ReturnOldDisplay()
    'returns the display to what it was originally
    
    If blnResChanged Then
        Call ChangeDisplaySettings(Null, 0)
        blnResChanged = False
    End If
End Sub

Public Sub SetDisplay(ByVal intWidth As Integer, _
                      ByVal intHeight As Integer)
    'changes the resolution of the screen to new size
    
    Dim udtDevM    As DEVMODE
    Dim lngResult As Long
    
    lngResult = EnumDisplaySettings(0, 0, udtDevM)
    
    With udtDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = intWidth  'ScreenWidth
            .dmPelsHeight = intHeight 'ScreenHeight
            .dmBitsPerPel = 32 '(could be 8, 16, 32 or even 4)
    End With
    lngResult = ChangeDisplaySettings(udtDevM, 2)
    
    If lngResult = 0 Then
            Call ChangeDisplaySettings(udtDevM, 4)
    End If
    
    blnResChanged = True
End Sub

Public Sub LoadBmp(ByRef hDcMemory As Long, _
                   ByRef hDcBitmap As Long, _
                   ByRef hDcPointer As Long, _
                   ByRef udtArea As Rect, _
                   ByVal strFileName As String, _
                   Optional ByVal udtImageType As LoadType = IMAGE_BITMAP, _
                   Optional ByVal udtMeasurement As Scaling = InPixels)
    'This procedure will load the specified image into a bitmap.
    'Please keep in mind that this function creates a bitmap object. If the
    'handles passed to this procedure already contain a bitmap, then it is
    'automatically deleted. This is assuming that the load operation was
    'successful. If not, then the pointers are left untouched.
    '
    'NOTE: This function is untested code from another source
    '12/12/2001 Eric
    
    Dim hBitmap As Long
    Dim hMemory As Long
    Dim hPointer As Long
    Dim udtLoadArea As Rect
    Dim lngResult As Long
    Dim udtDimensions As BITMAP
    
    'we have to pass a Null terminated string
    strFileName = strFileName & Chr(0)
    hBitmap = LoadImage(0, _
                        strFileName, _
                        udtImageType, _
                        0, _
                        0, _
                        LR_LOADFROMFILE)
    
    'if the load operation was unsuccessful then exit the procedure.
    If hBitmap = Null Then
        'load unsuccessful
        Exit Sub
    End If
    
    'create a device context and assign the bitmap to it
    hMemory = CreateCompatibleDC(0)
    hPointer = SelectObject(hMemory, hBitmap)
    
    'get the size of the bitmap loaded
    lngResult = GetObjectAPI(hBitmap, 14, udtDimensions)
    
    'assign the udtDimensions to a rect structure and convert to twips if
    'specified.
    udtLoadArea.Left = 0
    udtLoadArea.Top = 0
    udtLoadArea.Right = udtDimensions.bmWidth
    udtLoadArea.Bottom = udtDimensions.bmHeight
    
    If udtMeasurement = InTwips Then
        'convert pixel values to twips
        Call RectToTwips(udtLoadArea)
    End If
    
    'set the bitmap structure
    udtArea = udtLoadArea
    
    'Now that a new bitmap has been created, then we have to check to
    'see if the passed parameters are already in use hlding a bitmap. If so
    'then we need to delete it from memory before replacing it.
    Call DeleteBitmap(hDcMemory, hDcBitmap, hDcPointer)
    
    'return the handles of the bitmap created
    hDcMemory = hMemory
    hDcBitmap = hBitmap
    hDcPointer = hPointer
End Sub

Public Function RectIntersects(udtRect1 As Rect, _
                               udtRect2 As Rect) _
                               As Boolean
    'This function will return True if the two passed Rect structers
    'intersect each other.
    
    Dim udtTempRect As Rect
    Dim lngResult As Long
    
    lngResult = IntersectRect(udtTempRect, udtRect1, udtRect2)
    RectIntersects = lngResult
End Function

Public Function MouseKeyPressed(ByVal udtKey As MouseKeys) _
                                As Boolean
    'This will return True if the specified udtKey was pressed
    
    Const KeyPressed = -32768
    
    Dim lngResult As Long
    Dim lngMyKey As Long
    
    lngMyKey = udtKey
    
    'find the udtKey state of the mouse udtKey specified
    lngResult = GetAsyncKeyState(lngMyKey)
    
    'if the udtKey was pressed, then
    If lngResult = KeyPressed Then
        'the udtKey is pressed
        MouseKeyPressed = True
    End If
End Function

Public Sub GetScreenShot(ByVal lngDesthDc As Long, _
                         ByVal intDestX As Integer, _
                         ByVal intDestY As Integer, _
                         ByVal intDestWidth As Integer, _
                         ByVal intDestHeight As Integer, _
                         Optional ByVal intScreenX As Integer = 0, _
                         Optional ByVal intScreenY As Integer = 0, _
                         Optional ByVal udtMeasurement As Scaling = InPixels)
                         
    'This will get a screen shot at the specified co-ordinates and copy
    'them into the specified destination co-ordinates.
    
    Dim lngResult As Long
    
    'set the scaling mode specified and convert parameters if necessary
    If udtMeasurement = InTwips Then
        intDestX = intDestX / Screen.TwipsPerPixelX
        intDestY = intDestY / Screen.TwipsPerPixelY
        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
        intScreenX = intScreenX / Screen.TwipsPerPixelX
        intScreenY = intScreenY / Screen.TwipsPerPixelY
    End If
    
    'copy the screen shot - GetDesktopWindow was previously
    'used to get the handle on the top window.
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       intDestWidth, _
                       intDestHeight, _
                       GetDC(0), _
                       intScreenX, _
                       intScreenY, _
                       SRCCOPY)
End Sub

Public Sub SetMenuGraphic(ByVal hDcGraphic As Long, _
                          ByVal hWnd As Long, _
                          ByVal lngTopPos As Long, _
                          ByVal lngSubPos1 As Long, _
                          Optional ByVal lngSubPos2 As Long = -1, _
                          Optional ByVal lngSubPos3 As Long = -1)
    'This will set the graphic of a menu item
    'to any image in the device context. The menu
    'item must NOT be a top level menu or have a
    'sub menu. Top-level menu positions are from
    'left to right, starting at 0. Sub-level menu
    'positions are from top down, starting at 0.
    'NOTE : hDcGraphic MUST be the Picture property
    'of a control
    
    Const MF_BYPOSITION = &H400&
    Const MF_BYCOMMAND = &H0&
    Const BMP_SIZE = 14
    Const BF_BITMAP = &H4
    
    Dim lngResult As Long           'any error message returned from an api call
    Dim hMenu As Long               'the handle of the current menu item
    Dim hSubMenu As Long            'the handle of the current sub menu
    Dim lngID As Long               'the sub menu's ID
    
    'Get the handle of the form's menu
    hMenu = GetMenu(hWnd)
    
    'Get the handle of the form's submenu
    hSubMenu = GetSubMenu(hMenu, lngTopPos)
    
    'get any sub menu's
    If lngSubPos2 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos2)
    End If
    If lngSubPos3 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos3)
    End If
    
    'if we were unable to get the sub menu handle
    'then exit
    If (hMenu = 0) _
       Or (hSubMenu = 0) _
       Or (hDcGraphic = 0) Then
        Exit Sub
    End If
    
    'set the graphic to the sub menu
    lngID = GetMenuItemID(hSubMenu, lngSubPos1)
    
    lngResult = SetMenuItemBitmaps(hSubMenu, _
                                   lngID, _
                                   MF_BYCOMMAND, _
                                   hDcGraphic, _
                                   hDcGraphic)
End Sub

Public Sub GenerateMask(ByVal hDcSource, _
                        ByVal hDcDestination, _
                        ByVal intLeft As Integer, _
                        ByVal intTop As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        Optional ByVal intDestLeft As Integer, _
                        Optional ByVal intDestTop As Integer, _
                        Optional ByVal enmMeasurement As Scaling = InPixels)
    'This will automatically produce a monochrome
    'bitmap that is a mask for the source bitmap.
    'The mask colour transparent must be white
    'for this to work properly. The Mask is then
    'returned to the destination bitmap.
    
    Dim lngResult As Long           'holds any error value returned from the api calls
    Dim udtTempBmp As BitmapStruc   'holds the inverse of the source bitmap. This is used to create the mask
    
    'check the source and destination bitmaps
    If (hDcSource = 0) Or (hDcDestination = 0) Then
        Exit Sub
    End If
    
    'check the scale mode
    If enmMeasurement = InTwips Then
        'convert values to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intWidth = intWidth / Screen.TwipsPerPixelX
        intDestLeft = intDestLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intHeight = intHeight / Screen.TwipsPerPixelY
        intDestTop = intDestTop / Screen.TwipsPerPixelY
    End If
    
    'copy the bitmap to the destination as normal
    lngResult = BitBlt(hDcDestination, _
                       intDestLeft, _
                       intDestTop, _
                       intWidth, _
                       intHeight, _
                       hDcSource, _
                       intLeft, _
                       intTop, _
                       SRCINVERT)
    
    'create the temperory bitmap
    With udtTempBmp
        .Area.Right = intWidth
        .Area.Bottom = intHeight
        
        Call CreateNewBitmap(.hDcMemory, _
                             .hDcBitmap, _
                             .hDcPointer, _
                             .Area, _
                             hDcSource)
                             
        'copy the inverse bitmap into the
        'temperory bitmap
        lngResult = BitBlt(.hDcMemory, _
                           0, _
                           0, _
                           intWidth, _
                           intHeight, _
                           hDcSource, _
                           intLeft, _
                           intTop, _
                           SRCCOPY)
    
        'create the mask by "adding" the two bitmaps
        'onto the source bitmap
        'Call MergeBitmaps(hDcDestination, _
                          .hDcMemory, _
                          hDcSource, _
                          .hDcMemory, _
                          intDestLeft, _
                          intDestTop, _
                          intDestLeft, _
                          intDestTop, _
                          0, _
                          0, _
                          intWidth, _
                          intHeight)
        lngResult = BitBlt(hDcDestination, _
                           intDestLeft, _
                           intDestTop, _
                           intWidth, _
                           intHeight, _
                           .hDcMemory, _
                           0, _
                           0, _
                           SRCPAINT)
    
        'remove the temperory bitmap
        Call DeleteBitmap(.hDcMemory, _
                          .hDcBitmap, _
                          .hDcPointer)
    End With
End Sub

