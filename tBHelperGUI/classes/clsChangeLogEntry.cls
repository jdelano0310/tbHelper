VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clsChangeLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsChangeLogEntry
Private m_Values As Collection
Private m_TypeText As String
Private m_TypeColor As Long
Private m_TypeTextColor As Long

' colors for the type of changle log item
Private Const COLOR_IMPORTANT = vbBlue
Private Const COLOR_IMPROVED = RGB(107, 142, 35)      ' other green  
Private Const COLOR_FIXED = RGB(56, 118, 29) ' green
Private Const COLOR_UPDATED = RGB(107, 45, 103)       ' Orange badge
Private Const COLOR_TIP = RGB(22, 83, 126)           ' blueish
Private Const COLOR_DEFAULT = &H808080       ' Gray badge  
Private Const COLOR_KNOWN = vbBlack
Private Const COLOR_ADDED = RGB(75, 0, 130)
Private Const COLOR_WARNING = RGB(153, 0, 0)

Private Const TEXT_COLOR_WHITE = &HFFFFFF
    
Public Property Get Values() As Collection
    Set Values = m_Values
End Property

Public Property Get TypeText() As String
    TypeText = m_TypeText
End Property

Public Property Let TypeText(ByVal s As String)
    m_TypeText = s
End Property

Public Property Get TypeColor() As Long
    TypeColor = m_TypeColor
End Property

Public Property Let TypeColor(ByVal c As Long)
    m_TypeColor = c
End Property

Public Property Get TypeTextColor() As Long
    TypeTextColor = m_TypeTextColor
End Property

Public Property Let TypeTextColor(ByVal c As Long)
    m_TypeTextColor = c
End Property

Public Sub Initialize(ByVal columnDefs As Collection, columnValues() As Variant)
    Set m_Values = New Collection
    Dim i As Long
    Dim value As String

    For i = LBound(columnValues) To UBound(columnValues)
        value = columnValues(i)
        Dim colName As String

        If i + 1 <= columnDefs.Count Then
            colName = columnDefs(i + 1).Name  ' Collections are 1-based
            m_Values.Add value, colName

            If UCase(colName) = "TYPE" Then
                m_TypeText = UCase(value)
                m_TypeTextColor = TEXT_COLOR_WHITE

                Select Case m_TypeText
                    Case "IMPORTANT": m_TypeColor = COLOR_IMPORTANT
                    Case "IMPROVED":  m_TypeColor = COLOR_IMPROVED
                    Case "FIXED":     m_TypeColor = COLOR_FIXED
                    Case "UPDATED":   m_TypeColor = COLOR_UPDATED
                    Case "TIP":       m_TypeColor = COLOR_TIP
                    Case "KNOWN ISSUE": m_TypeColor = COLOR_KNOWN
                    Case "ADDED":     m_TypeColor = COLOR_ADDED
                    Case "WARNING":   m_TypeColor = COLOR_WARNING
                    Case Else:        m_TypeColor = COLOR_DEFAULT
                End Select
            End If
        End If
    Next i
End Sub