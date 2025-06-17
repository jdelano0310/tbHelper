Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As SIZE) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function GetModuleFileName Lib "kernel32" _
Alias "GetModuleFileNameA" _
(ByVal hModule As Long, _
ByVal lpFileName As String, _
ByVal nSize As Long) As Long

Private Const WM_GETFONT As Long = &H31

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
