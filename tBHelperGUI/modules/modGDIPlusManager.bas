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

Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const GWL_WNDPROC As Long = -4

Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private CallbackOwner As Object
Public OriginalCanvasProc As Long

Public ucDictionary As New Scripting.Dictionary   ' my own dictionary to hold window handle to object (user controls) 

Public Sub RegisterScrollableCanvas(ByVal hWnd As Long, ByVal ownerCtrl As Object)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Canvas_WindowProc)
    ucDictionary.Add hWnd, ownerCtrl
End Sub

Public Const ROW_ALT_COLOR = &HF8F8F8
Public Const CUST_BTN_BCOLOR = &HA2640C

Public Sub UpdateScrollOwnership(ByVal hWnd As LongPtr, ByVal newPos As Long)
        
    If ucDictionary.Exists(hWnd) Then

        Set CallbackOwner = ucDictionary(hWnd)
        
        On Error GoTo errorHandler   ' this is mainly used to stop the code from breaking while in design mode
        
        ' there is no need to track scrolls on frmMain controls if a panel is displayed
        If CallbackOwner.parent.name = "frmMain" Then
            
            If CallbackOwner.parent.isAPaneldisplayed And (CallbackOwner.tag <> "LogHistoryView" Or CallbackOwner.tag <> "RevertView") Then
                'WriteToDebugLogFile("   UpdateScrollOwnership: ignore call " & CallbackOwner.Name & " on " & CallbackOwner.parent.name)
                Exit Sub
            End If
        End If
        
        WriteToDebugLogFile("   UpdateScrollOwnership: called from " & CallbackOwner.Name & " on " & CallbackOwner.parent.name)
        If Not CallbackOwner Is Nothing	Then
            ' Update scroll registry and queue redraw
            If ScrollRegistry Is Nothing Then Set ScrollRegistry = New Collection
            
            If ScrollRegistry.Exists(CStr(hWnd)) Then
                Dim reg As clsScrollRegistry
                Set reg = ScrollRegistry(CStr(hWnd))
                reg.LastScrollPos = newPos
            End If

            QueueRedraw hWnd
        Else
            WriteToDebugLogFile("UpdateScrollOwnership: from something not in ucDictionary")
            
        End If
    End If
    
    errorHandler:
        ' just skip if there is an error 
        If Err.Number <> 0 Then Debug.Print "UpdateScrollOwnership error " & Err.Description
End Sub

Public Function Canvas_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If uMsg = WM_MOUSEWHEEL Then
        
        Dim newPos As Long
        newPos = GetScrollPos(hwnd, SB_VERT) ' Get current vertical scroll position

        UpdateScrollOwnership hwnd, newPos

        If Not CallbackOwner Is Nothing Then
            WriteToDebugLogFile("Canvas_WindowProc: calling HandleMouseScroll in " & CallbackOwner.name & " on " & CallbackOwner.parent.name)
            CallByName CallbackOwner, "HandleMouseScroll", vbMethod, wParam
        End If

        Exit Function
    End If

    Canvas_WindowProc = CallWindowProc(OriginalCanvasProc, hwnd, uMsg, wParam, lParam)
End Function

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

Private ScrollRegistry As Collection

Public Sub RegisterScrollOwner(hWnd As LongPtr)
    
    WriteToDebugLogFile("RegisterScrollOwner hWnd=" & hWnd)

    If ucDictionary.Exists(hWnd) Then
        WriteToDebugLogFile("   RegisterScrollOwner: called from " & ucDictionary(hWnd).Name & " on " & ucDictionary(hWnd).parent.name)
    End If
    
    Dim state As clsScrollRegistry
    state.hWnd = hWnd
    state.IsActive = True
    state.LastScrollPos = 0
    
    If ScrollRegistry Is Nothing Then Set ScrollRegistry = New Collection
    
    ScrollRegistry.Add state, CStr(hWnd)
End Sub

Private RedrawQueue As Collection

Public Sub QueueRedraw(hWnd As LongPtr)
    
    If RedrawQueue Is Nothing Then Set RedrawQueue = New Collection
    
    If Not RedrawQueue.Exists(CStr(hWnd)) Then
        Dim req As New clsRedrawRequest
        req.hWnd = hWnd
        req.LastRequested = Timer
        RedrawQueue.Add req, CStr(hWnd)
    Else
        RedrawQueue(CStr(hWnd)).LastRequested = Timer
    End If
    
    If ucDictionary.Exists(hWnd) Then
        WriteToDebugLogFile("   QueueRedraw: called from " & ucDictionary(hWnd).Name & " on " & ucDictionary(hWnd).parent.name)
    Else
        WriteToDebugLogFile("   QueueRedraw: called from something not in ucDictionary ")
    End If
    
End Sub

Public Sub FlushRedraws()
    
    If RedrawQueue.Count = 0 Then Exit Sub
    
    WriteToDebugLogFile("Flushing the Redraws - there are " & RedrawQueue.Count & " requiring attention")
    
    Dim req As clsRedrawRequest
    Dim hWnds As String = "|"    ' holds the handles redraw during this pass
    Dim now As Double
    Dim countRedrawItems As Integer = 1
    
    now = Timer ' Capture current timestamp once for consistency

    For Each req In RedrawQueue
        ' check that enough times has passed since last redraw, and the handle hasn't been refreshed
        ' during this loop, refresh it
        If ucDictionary.Exists(req.hWnd) Then
            WriteToDebugLogFile("   FlushRedraws, request #" & countRedrawItems & " RedrawQueue loop control to redraw: " & ucDictionary(req.hWnd).Name & " on " & ucDictionary(req.hWnd).parent.name & " tag: " & ucDictionary(req.hWnd).Tag)
        End If
        
        If now - req.LastRequested > 0.05 And InStr(hWnds, "|" & req.hWnd & "|") = 0 Then
            If req.Region <> 0 Then
                RedrawWindow req.hWnd, req.Region, 0, RDW_INVALIDATE
            Else
                InvalidateRect req.hWnd, ByVal 0&, True
            End If
            
            hWnds = hWnds & CStr(req.hWnd) & "|"
            WriteToDebugLogFile("   FlushRedraws, request #" & countRedrawItems & " RedrawQueue called RedrawWindow")
            
            ' if this flush is for a custom button, then tell it the redraw via flushredraws is done.
            If ucDictionary.Exists(req.hWnd) Then
                On Error Resume Next
                CallByName ucDictionary(req.hWnd), "FlushRedrawComplete", vbMethod
                If Err = 0 Then WriteToDebugLogFile("   FlushRedraws: request #" & countRedrawItems & " this is a button, so call its FlushRedrawComplete ")
                On Error GoTo 0
            End If
            countRedrawItems += 1
        End If
    Next

    RedrawQueue.Clear
    WriteToDebugLogFile("FlushRedraws - done ")
    
End Sub