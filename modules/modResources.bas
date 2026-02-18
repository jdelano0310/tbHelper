Attribute VB_Name = "modResources"

Public Enum ResourceID
    resNone
    resBinoculars
    resDelete
    resDownload
    resError
    resFolder
    resInfo
    resLogHistory
    resLogHistoryPanelIcon
    resMessageBox
    resQuestion
    resRevert
    resRevertPanelIcon
    resSuccess
    resWarning
    resFolderClosed
    resFolderOpen
    resBlackFolder
End Enum

Private mImages As New Dictionary(Of ResourceID, LongPtr)
Private mStdPics As New Dictionary(Of ResourceID, StdPicture)

Public Sub LoadImages()
    If IsCodeRunningInTheIDE Then
        'WriteToDebugLogFile "LoadImages - IDE"
        LoadImagesFromDisk
    Else
        'WriteToDebugLogFile "LoadImages - EXE"
        LoadImagesFromResources
    End If
End Sub

Private Sub AddImage(id As ResourceID, pic As LongPtr)
    mImages.Add id, pic
End Sub

Private Sub LoadImagesFromDisk()
    'WriteToDebugLogFile "    LoadImagesFromDisk"
    Dim base As String
    base = App.Path & "\"

    AddImage resBinoculars, LoadGDIPlusBitmapFromFile(base & "binoculars.png")
    AddImage resDelete, LoadGDIPlusBitmapFromFile(base & "delete.png")
    AddImage resDownload, LoadGDIPlusBitmapFromFile(base & "download.png")
    AddImage resError, LoadGDIPlusBitmapFromFile(base & "error.png")
    AddImage resFolder, LoadGDIPlusBitmapFromFile(base & "folder.png")
    AddImage resInfo, LoadGDIPlusBitmapFromFile(base & "info.png")
    AddImage resLogHistory, LoadGDIPlusBitmapFromFile(base & "logHistory.png")
    AddImage resLogHistoryPanelIcon, LoadGDIPlusBitmapFromFile(base & "logHistorypanel icon.png")
    AddImage resMessageBox, LoadGDIPlusBitmapFromFile(base & "messagebox.png")
    AddImage resQuestion, LoadGDIPlusBitmapFromFile(base & "question.png")
    AddImage resRevert, LoadGDIPlusBitmapFromFile(base & "revert.png")
    AddImage resRevertPanelIcon, LoadGDIPlusBitmapFromFile(base & "revert panel icon.png")
    AddImage resSuccess, LoadGDIPlusBitmapFromFile(base & "success.png")
    AddImage resWarning, LoadGDIPlusBitmapFromFile(base & "warning.png")
    AddImage resFolderClosed, LoadGDIPlusBitmapFromFile(base & "folder_closed.ico")
    AddImage resFolderOpen, LoadGDIPlusBitmapFromFile(base & "folder_open.ico")
    AddImage resBlackFolder, LoadGDIPlusBitmapFromFile(base & "black_folder_open.ico")
End Sub

Private Function LoadGDIPlusBitmapFromFile(path As String) As LongPtr
    Dim bmp As LongPtr
    If GdipLoadImageFromFile(StrPtr(path), bmp) = 0 Then
        LoadGDIPlusBitmapFromFile = bmp
    End If
End Function

Private Sub LoadImagesFromResources()
    'WriteToDebugLogFile "    LoadImagesFromResources"

    On Error GoTo LoadError

    AddImage resBinoculars, LoadGDIPlusBitmapFromResource("BINOCULARS.PNG")
    AddImage resDelete, LoadGDIPlusBitmapFromResource("DELETE.PNG")
    AddImage resDownload, LoadGDIPlusBitmapFromResource("DOWNLOAD.PNG")
    AddImage resError, LoadGDIPlusBitmapFromResource("ERROR.PNG")
    AddImage resFolder, LoadGDIPlusBitmapFromResource("FOLDER.PNG")
    AddImage resInfo, LoadGDIPlusBitmapFromResource("INFO.PNG")
    AddImage resLogHistory, LoadGDIPlusBitmapFromResource("LOGHISTORY.PNG")
    AddImage resLogHistoryPanelIcon, LoadGDIPlusBitmapFromResource("LOGHISTORYPANEL ICON.PNG")
    AddImage resMessageBox, LoadGDIPlusBitmapFromResource("MESSAGEBOX.PNG")
    AddImage resQuestion, LoadGDIPlusBitmapFromResource("QUESTION.PNG")
    AddImage resRevert, LoadGDIPlusBitmapFromResource("REVERT.PNG")
    AddImage resRevertPanelIcon, LoadGDIPlusBitmapFromResource("REVERT PANEL ICON.PNG")
    AddImage resSuccess, LoadGDIPlusBitmapFromResource("SUCCESS.PNG")
    AddImage resWarning, LoadGDIPlusBitmapFromResource("WARNING.PNG")
    AddImage resFolderClosed, LoadGDIPlusBitmapFromResource("FOLDER_CLOSED.ICO")
    AddImage resFolderOpen, LoadGDIPlusBitmapFromResource("FOLDER_OPEN.ICO")
    AddImage resBlackFolder, LoadGDIPlusBitmapFromResource("BLACK_FOLDER_OPEN.ICO")
    Exit Sub

LoadError:
    'WriteToDebugLogFile "    LoadImagesFromResources: Error: " & Err.Description
End Sub

Public Function GetImagePtr(id As ResourceID) As LongPtr
    If mImages.Exists(id) Then
        GetImagePtr = mImages(id)
    End If
End Function

Public Function GetImageStd(id As ResourceID) As StdPicture
    If mStdPics.Exists(id) Then
        Set GetImageStd = mStdPics(id)
    ElseIf mImages.Exists(id) Then
        Dim pic As StdPicture
        Set pic = GpBitmapToStdPicture(mImages(id))
        If Not pic Is Nothing Then
            mStdPics.Add id, pic
            Set GetImageStd = pic
        End If
    End If
End Function

Private Function GetResourceTypePtr(resName As String) As LongPtr
    
    'WriteToDebugLogFile "GetResourceTypePtr: " & resName
    Dim ext As String
    Dim resourceValue As LongPtr = 0
    
    ext = LCase$(Mid$(resName, InStrRev(resName, ".") + 1))

    Select Case ext
        Case "png"
            resourceValue = StrPtr("PNG")
        Case "ico"
            resourceValue = 14 ' RT_GROUP_ICON

    End Select
    
    'WriteToDebugLogFile "GetResourceTypePtr: resourceValue " & resourceValue
    GetResourceTypePtr = resourceValue
    
End Function

Public Function LoadGDIPlusBitmapFromResource(resName As String) As LongPtr
    Dim hRes As LongPtr
    Dim hData As LongPtr
    Dim pData As LongPtr
    Dim pImage As LongPtr
    Dim size As Long
    Dim hGlobal As LongPtr
    Dim pGlobal As LongPtr
    Dim pStream As IStream
    Dim gpBmp As LongPtr
    Dim ext As String

    'WriteToDebugLogFile "LoadGDIPlusBitmapFromResource: " & resName

    ext = LCase$(Mid$(resName, InStrRev(resName, ".") + 1))

    If ext = "png" Then
        'WriteToDebugLogFile "         PNG"
        hRes = FindResourceW(0, StrPtr(resName), StrPtr("PNG"))
    ElseIf ext = "ico" Then
        'WriteToDebugLogFile "         ICO"
        hRes = FindResourceW(0, StrPtr(resName), StrPtr("ICO")) ' RT_GROUP_ICON
    Else
        'WriteToDebugLogFile "         **** something else **** "
        Exit Function
    End If
    If hRes = 0 Then Exit Function
    
    hData = LoadResource(0, hRes)
    If hData = 0 Then Exit Function
    

    pData = LockResource(hData)
    If pData = 0 Then Exit Function

    size = SizeofResource(0, hRes)
    If size = 0 Then Exit Function

    ' --- try to locate actual image bytes inside the resource block ---
    Dim header(0 To 7) As Byte
    CopyMemory VarPtr(header(0)), pData, 8
    'WriteToDebugLogFile "         Header: " & _
        Hex(header(0)) & " " & Hex(header(1)) & " " & Hex(header(2)) & " " & Hex(header(3)) & " " & _
        Hex(header(4)) & " " & Hex(header(5)) & " " & Hex(header(6)) & " " & Hex(header(7))

    pImage = pData

    Dim i As Long
    Dim found As Boolean

    ' PNG magic: 89 50 4E 47 0D 0A 1A 0A
    ' ICO magic: 00 00 01 00
    If ext = "png" Then
        If Not (header(0) = &H89 And header(1) = &H50 And header(2) = &H4E And header(3) = &H47) Then
            For i = 0 To size - 8
                CopyMemory VarPtr(header(0)), pData + i, 8
                If header(0) = &H89 And header(1) = &H50 And header(2) = &H4E And header(3) = &H47 Then
                    pImage = pData + i
                    found = True
                    'WriteToDebugLogFile "         PNG signature found at offset " & i
                    Exit For
                End If
            Next
        Else
            found = True
            'WriteToDebugLogFile "         PNG signature at offset 0"
        End If
    ElseIf ext = "ico" Then
        If Not (header(0) = &H0 And header(1) = &H0 And header(2) = &H1 And header(3) = &H0) Then
            For i = 0 To size - 4
                CopyMemory VarPtr(header(0)), pData + i, 4
                If header(0) = &H0 And header(1) = &H0 And header(2) = &H1 And header(3) = &H0 Then
                    pImage = pData + i
                    found = True
                    'WriteToDebugLogFile "         ICO signature found at offset " & i
                    Exit For
                End If
            Next
        Else
            found = True
            'WriteToDebugLogFile "         ICO signature at offset 0"
        End If
    End If

    If Not found Then
        'WriteToDebugLogFile "         *** No PNG/ICO signature found in resource block"
        Exit Function
    End If

    ' --- now copy from pImage instead of pData ---
    hGlobal = GlobalAlloc(&H2, size)
    If hGlobal = 0 Then Exit Function

    pGlobal = GlobalLock(hGlobal)
    If pGlobal = 0 Then
        GlobalFree hGlobal
        Exit Function
    End If

    CopyMemory pGlobal, pImage, size
    GlobalUnlock hGlobal

    If WinDevLib.wdAPI.CreateStreamOnHGlobal(hGlobal, True, pStream) <> 0 Then
        GlobalFree hGlobal
        Exit Function
    End If

    Dim status As Long
    status = WinDevLib.wdGDIP.GdipCreateBitmapFromStream(pStream, gpBmp)
    'WriteToDebugLogFile "         GdipCreateBitmapFromStream status: " & status

    If status = 0 Then
        LoadGDIPlusBitmapFromResource = gpBmp
    End If

    Set pStream = Nothing
End Function

Private Function GpBitmapToStdPicture(gpBmp As LongPtr) As StdPicture
    Dim hBmp As LongPtr
    Dim status As Long
    Dim pic As StdPicture
    Dim pd As PICTDESC

    'WriteToDebugLogFile "  GpBitmapToStdPicture: gpBmp " & gpBmp
    
    If gpBmp = 0 Then Exit Function

    status = GdipCreateHBITMAPFromBitmap(gpBmp, hBmp, &HFFFFFFFF)
    If status <> 0 Or hBmp = 0 Then Exit Function
    
    'WriteToDebugLogFile "  GpBitmapToStdPicture: passed  status <> 0 Or hBmp = 0"
    With pd
        .cbSizeofstruct = Len(pd)
        .picType = vbPicTypeBitmap
        .hImage = hBmp
        .hPalette = 0
    End With

    'WriteToDebugLogFile "  GpBitmapToStdPicture: calling OleCreatePictureIndirect "
    OleCreatePictureIndirect pd, IID_IPicture, True, pic
    Set GpBitmapToStdPicture = pic
End Function

' checking to make sure the resources are embedded in the EXE
Private Declare Function EnumResourceTypesW Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr) As Long

Private Declare Function EnumResourceNamesW Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr

Public Sub EnumerateAllResources()
    Dim hMod As LongPtr
    hMod = GetModuleHandleW(0)

    Debug.Print "Enumerating resources..."
    EnumResourceTypesW hMod, AddressOf EnumTypesCallback, 0
End Sub

Private Function EnumTypesCallback( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lParam As LongPtr) As Long

    Dim typeName As String

    If lpType < &H10000 Then
        typeName = "#" & CStr(lpType)
    Else
        typeName = StrFromPtrW(lpType)
    End If

    LogToFile "Resource Type: " & typeName
    
    EnumResourceNamesW hModule, lpType, AddressOf EnumNamesCallback, 0

    EnumTypesCallback = 1 ' continue enumeration
End Function

Private Function EnumNamesCallback( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lpName As LongPtr, _
    ByVal lParam As LongPtr) As Long

    Dim name As String

    If lpName < &H10000 Then
        name = "#" & CStr(lpName)
    Else
        name = StrFromPtrW(lpName)
    End If

    LogToFile "     Name: " & name

    EnumNamesCallback = 1 ' continue enumeration
End Function

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Public Function StringFromPtrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function

    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function

    Dim s As String
    s = String$(cch, vbNullChar)

    CopyMemory StrPtr(s), p, cch * 2

    StringFromPtrW = s
End Function

Private Function StrFromPtrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    StrFromPtrW = StringFromPtrW(p)
End Function

Private Sub LogToFile(ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\resource_dump.txt" For Append As #f
    Print #f, text
    Close #f
End Sub