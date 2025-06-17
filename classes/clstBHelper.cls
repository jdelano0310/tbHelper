VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clstBHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' add your member fields and procedures here
Private _installedtBVersion As Integer    ' this didn't fit in the change log items collection
Private _sessionLog As Collection         ' thinking write the log here instead of to the listbox on form1 

Public Property Let InstalledtBVersion(installedVersion As Integer)
    _installedtBVersion = installedVersion
End Property

Public Property Get InstalledtBVersion() As Integer

    InstalledtBVersion = _installedtBVersion
End Property

Public Property Let SessionLog(LogMessage As String)
    _sessionLog.Add(LogMessage)
End Property

Public Property Get SessionLog() As Collection

    Set SessionLog = _sessionLog
    
End Property