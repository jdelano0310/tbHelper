Attribute VB_Name = "APIs"

' GDI+ 
Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As Long
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As SIZE) As Long
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long

Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nWidthEllipse As Long, ByVal nHeightEllipse As Long) As LongPtr
Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As LongPtr
Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Boolean) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function CombineRgn Lib "gdi32" (ByVal hrgnDest As LongPtr, ByVal hrgnSrc1 As LongPtr, ByVal hrgnSrc2 As LongPtr, ByVal fnCombineMode As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'-- GDI+ Types
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type

' change the png to greyscale to denote disabled status
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (Imageattr As Long) As Long
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long

Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" ( _
    ByVal imageattr As Long, _
    ByVal ColorAdjustType As Long, _
    ByVal enableFlag As Long, _
    ByVal colorMatrix As Long, _
    ByVal grayMatrix As Long, _
    ByVal flags As Long) As Long

Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" ( _
    ByVal graphics As Long, _
    ByVal image As Long, _
    ByVal dstx As Long, _
    ByVal dsty As Long, _
    ByVal dstwidth As Long, _
    ByVal dstheight As Long, _
    ByVal srcx As Long, _
    ByVal srcy As Long, _
    ByVal srcwidth As Long, _
    ByVal srcheight As Long, _
    ByVal srcUnit As Long, _
    ByVal imageAttributes As Long) As Long

Private Const WM_GETFONT As Long = &H31

Public Const SM_CXVSCROLL = 2

' hWndInsertAfter constants
Private Const RGN_OR = 2
Private Const HWND_TOP As Long = 0

' uFlags constants
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Public GDIPlus_Ready As Boolean
Private GdiPlusToken As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
Public FO_DELETE As Long = &H3
Public FOF_ALLOWUNDO As Long = &H40

Public Const FO_COPY = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10

Public Declare Function SHFO_UnZip Lib "shell32" Alias "SHFileOperationW" (ByVal lpFileOp As Long) As Long

Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public BINDF_GETNEWESTVERSION As Long = &H10
    
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public SW_HIDE As Integer = 0

Public Sub DrawGDIPlusImageFromFile(hdc As Long, path As String, x As Long, y As Long, w As Long, h As Long)
    ' used by ucCustomButton to draw the image in the button whie using a png file with transparent background
    Dim img As Long, g As Long

    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        If GdipCreateFromHDC(hdc, g) = 0 Then
            Call GdipDrawImageRectI(g, img, x, y, w, h)
            Call GdipDeleteGraphics(g)
        End If
        Call GdipDisposeImage(img)
    End If
End Sub

Public Sub InitializeGDIPlus()
    If GDIPlus_Ready Then Exit Sub

    Dim GDIStart As GdiplusStartupInput
    GDIStart.GdiplusVersion = 1

    If GdiplusStartup(GdiPlusToken, GDIStart, 0) = 0 Then
        GDIPlus_Ready = True
    End If
End Sub

Public Sub ShutdownGDIPlus()
    If GDIPlus_Ready Then
        Call GdiplusShutdown(GdiPlusToken)
        GDIPlus_Ready = False
    End If
End Sub

Public Sub ApplyRoundedCorners(ctl As Object, ByVal RadiusX As Long, ByVal RadiusY As Long)
    Dim hRgn As LongPtr
    Dim clientRect As RECT
    
    ' CORRECTION: Use GetClientRect to reliably get dimensions in PIXELS.
    ' This works for all controls, including UserControls.
    GetClientRect ctl.hWnd, clientRect
    
    ' clientRect.Right IS the width in pixels.
    ' clientRect.Bottom IS the height in pixels.
    hRgn = CreateRoundRectRgn(0, 0, clientRect.Right, clientRect.Bottom, RadiusX, RadiusY)
    
    SetWindowRgn ctl.hWnd, hRgn, True
End Sub

Public Sub ApplyBottomRoundedCorners(ctl As Object, ByVal RadiusX As Long, ByVal RadiusY As Long, Optional ByVal hasVerticalScrollbar As Boolean = False)
    Dim hRgn_Combined As LongPtr
    Dim hRgn_TopRect As LongPtr
    Dim hRgn_BottomRoundRect As LongPtr
    Dim clientRect As RECT
    
    ' Get the dimensions of the control's CLIENT area (the part inside borders/scrollbars)
    GetClientRect ctl.hWnd, clientRect
    
    Dim PxWidth As Long
    Dim PxHeight As Long
    PxWidth = clientRect.Right   ' For a client rect, Right IS the width
    PxHeight = clientRect.Bottom ' and Bottom IS the height
    
    ' YOUR BRILLIANT IDEA: If the control has a scrollbar, add its width back
    ' to the region's total width to avoid clipping it.
    If hasVerticalScrollbar Then
        PxWidth = PxWidth + GetSystemMetrics(SM_CXVSCROLL)
    End If
    
    ' 1. Create a sharp rectangle for the top part
    hRgn_TopRect = CreateRectRgn(0, 0, PxWidth, PxHeight - RadiusY)

    ' 2. Create a rounded rectangle for the full area
    hRgn_BottomRoundRect = CreateRoundRectRgn(0, 0, PxWidth, PxHeight, RadiusX, RadiusY)
    
    ' 3. Create an empty region to hold the result
    hRgn_Combined = CreateRectRgn(0, 0, 0, 0)
    
    ' 4. Combine the two regions
    CombineRgn hRgn_Combined, hRgn_TopRect, hRgn_BottomRoundRect, RGN_OR
    
    ' 5. Apply the final combined shape to the control's window
    SetWindowRgn ctl.hWnd, hRgn_Combined, True
    
    ' 6. Clean up
    DeleteObject hRgn_TopRect
    DeleteObject hRgn_BottomRoundRect
End Sub

''' <summary>
''' Uses the SetWindowPos API to bring a window handle to the top of the Z-Order
''' without moving or resizing it. This is more forceful than the ZOrder method.
''' </summary>
Public Sub BringWindowToTop(ByVal hWnd As LongPtr)
    SetWindowPos hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub