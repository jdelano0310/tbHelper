Option Explicit

' Types
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

' API Declarations (gdiplus.dll)
Private Declare Function GdiplusStartup Lib "gdiplus" (token As LongPtr, input As GdiplusStartupInput, ByVal output As LongPtr) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long

' Module-level state
Private m_gdiToken As LongPtr    ' token returned by GdiplusStartup
Private m_refCount As Long       ' simple reference counter
Private m_started As Boolean

' Call this before any GDI+ calls (controls can call on Create/Init)
Public Function GDIPlus_EnsureStarted() As Boolean
    Dim ret As Long
    Dim s As GdiplusStartupInput

    If m_started Then
        m_refCount = m_refCount + 1
        GDIPlus_EnsureStarted = True
        Exit Function
    End If

    ' Fill startup input
    s.GdiplusVersion = 1
    s.DebugEventCallback = 0
    s.SuppressBackgroundThread = 0
    s.SuppressExternalCodecs = 0

    ' Start GDI+
    ret = GdiplusStartup(m_gdiToken, s, 0)
    If ret = 0 Then    ' Ok (Gdiplus::Ok == 0)
        m_started = True
        m_refCount = 1
        GDIPlus_EnsureStarted = True
    Else
        ' Failed to initialize GDI+. ret contains the GpStatus code.
        m_started = False
        m_refCount = 0
        GDIPlus_EnsureStarted = False
    End If
End Function

' Call this when a control/form no longer needs GDI+ (e.g., in Terminate/Unload)
Public Sub GDIPlus_Release()
    If Not m_started Then Exit Sub

    If m_refCount > 1 Then
        m_refCount = m_refCount - 1
        Exit Sub
    End If

    ' Last release -> shutdown
    On Error Resume Next
    GdiplusShutdown m_gdiToken
    m_gdiToken = 0
    m_refCount = 0
    m_started = False
    On Error GoTo 0
End Sub

' Optional: force shutdown (for debugging). Use with care.
Public Sub GDIPlus_ForceShutdown()
    If m_started Then
        On Error Resume Next
        GdiplusShutdown m_gdiToken
        m_gdiToken = 0
        m_refCount = 0
        m_started = False
        On Error GoTo 0
    End If
End Sub

' Query
Public Function GDIPlus_IsStarted() As Boolean
    GDIPlus_IsStarted = m_started
End Function
