Attribute VB_Name = "modAddRoundedCorners"

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
    ByVal XRadius As Long, ByVal YRadius As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub ApplyRoundedCorners(ByVal ctrl As Object, Optional ByVal radiusX As Long = 8, Optional ByVal radiusY As Long = 8)

    Dim rc As RECT
    If GetClientRect(ctrl.hWnd, rc) = 0 Then Exit Sub

    Dim pxWidth As Long, pxHeight As Long
    pxWidth = rc.Right - rc.Left
    pxHeight = rc.Bottom - rc.Top

    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, pxWidth, pxHeight, radiusX, radiusY)

    If hRgn <> 0 Then
        SetWindowRgn ctrl.hWnd, hRgn, True
        DeleteObject hRgn
    End If
End Sub


