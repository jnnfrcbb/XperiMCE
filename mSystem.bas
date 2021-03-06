Attribute VB_Name = "mSystem"
Public Const ANYSIZE_ARRAY = 1
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Type LUID
    LowPart As Long
    HighPart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long


Public Sub EnableShutDown()
    
    StopPlayback
    
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    ' enable shutdown privilege for the current application
    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)

End Sub
