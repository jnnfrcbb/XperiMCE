Attribute VB_Name = "mProcess"
'Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long
Const SW_SHOWNORMAL = 1

Public Type PointAPI
    x As Long
    y As Long
End Type
Public MouseX As Long
Public MouseY As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer




Public Type PROCESSENTRY32
    dwSize              As Long ' The length in bytes of the structure
    cntUsage            As Long ' The number of references to the process
    th32ProcessID       As Long ' Identifier of the process
    th32DefaultHeapID   As Long ' Identifier of the default heap for the process
    th32ModuleID        As Long ' Identifier of the process's module
    cntThreads          As Long ' The number of threads started by the program
    th32ParentProcessID As Long ' The identifier of the process that created this process
    pcPriClassBase      As Long ' The base priority by any threads created by this class
    dwFlags             As Long
    szExeFile           As String * 260 ' The filename of the executable file for the process
End Type

'CreateToolhelp32Snapshot creates a snapshot of what is running on the computer the moment the function is called.
Public Declare Function CreateToolhelp32Snapshot Lib "Kernel32.dll" ( _
               ByVal dwFlags As Long, _
               ByVal th32ProcessID As Long) As Long

'Process32First retrieves information about the first process in the process list contained in a system snapshot.
Public Declare Function Process32First Lib "Kernel32.dll" ( _
               ByVal hSnapshot As Long, _
               lppe As PROCESSENTRY32) As Long

'Process32Next retrieves information about the next unread process in the process list contained in a system snapshot.
Public Declare Function Process32Next Lib "Kernel32.dll" ( _
               ByVal hSnapshot As Long, _
               lppe As PROCESSENTRY32) As Long
               
'CloseHandle closes a handle and the object associated with that handle.
Public Declare Function CloseHandle Lib "Kernel32.dll" ( _
               ByVal hObject As Long) As Long
               
'Include the process list in the snapshot
Public Const TH32CS_SNAPPROCESS As Long = &H2

Public bKey As Boolean

Public xmlFileDoc As New DOMDocument
Public xmlFileTypes As IXMLDOMNodeList
Public xmlFileType As IXMLDOMNode

Public bRemoteSupport As Boolean
Public bGameControllerSupport As Boolean
Public bControllerBypass As Boolean

Public Volume As cVolume
Public lVolumeCount As Long

Public lPower As Long

Public lMouseCount As Long
Public bMouseOut As Boolean
Public bMouseOnForm As Boolean
Public bMouseCharm As Boolean

Public sProcessPath As String
Public bProcess As Boolean

Public lNotify As Long
Public bNotify As Boolean

Public bStatusVisible As Boolean

Public bSounds As Boolean

Public bNewKey As Boolean

Public bFocus As Boolean

Public lGameProcessInterval As Long

Public fso As New FileSystemObject

Public Function KeyPress(KeyCode As Integer, Optional Shift As Integer = 0, Optional bForceKey As Boolean = False) As Boolean
    
    On Error Resume Next
    
    With frmMain
        
        If bForceKey = True Then
            bKey = True
        End If
        
        If bKey = True Then
        
            If Shift = 1 Then
            
                bKey = False
            
                Select Case KeyCode
            
                    Case vbKeyShift
                        
                        'Do Until KeyPress(vbKeyShift, 0, True) = True
                        '    DoEvents
                        'Loop
            
                    Case 51 'hash
                    
                        Do Until KeyPress(vbKeyShift, 0, True) = True
                            DoEvents
                        Loop
                                                
                End Select
            
                bKey = True
                
            Else
            
                Shift = 0
            
                bKey = False
                
                If .pPower(0).Visible = True Then
                
                    Do Until PowerKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                    
                ElseIf .pLockHolder.Visible = True Then
                
                    Do Until LockKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                    
                ElseIf .pVideoHolder.Visible = True And bMinVideo = False Then
        
                    Do Until VideoPlayingKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                              
                ElseIf .pCharmsHolder.Visible = True And bStatusVisible = False Then
                
                    Do Until CharmsKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                     
                ElseIf .pMenuHolder.Visible = True Then
                
                    Do Until MenuKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                
                ElseIf .pPlaylistHolder.Visible = True Then
                    
                    Do Until PlaylistKey(KeyCode) = True
                        DoEvents
                    Loop
                       
                ElseIf .pTileOptions.Visible = True Then
                
                    Do Until TileOptionsKey(KeyCode, Shift) = True
                        DoEvents
                    Loop
                
                Else
                
                    Select Case KeyCode
                    
                        Case vbKeyF1 'charms
                                                
                            If bDashLocked = False Then
                            
                                Do Until OpenCharms = True
                                    DoEvents
                                Loop
                            
                            End If
                            
                        'case vbKeyF2 'menu
                            
                        'case vbKeyF3 'search
                            
                        Case vbKeyF4 'lock
                                    
                            Do Until ShowLock = True
                                DoEvents
                            Loop
                
                        Case vbKeyF5 'shuffle
                        
                        Case vbKeyF6 'loop
                        
                        Case vbKeyControl
                        
                        Case Else
                        
                            Select Case CurrentScreen
                               
                                Case 0 'home
                                
                                    Do Until DashKey(KeyCode, Shift) = True
                                        DoEvents
                                    Loop
                                
                                Case 1 'app
                                
                                    Do Until AppKey(KeyCode, Shift) = True
                                        DoEvents
                                    Loop
                                
                                Case 2 'search
                                
                                    Do Until SearchKey(KeyCode, Shift) = True
                                        DoEvents
                                    Loop
                                
                                Case 3 'search results
                               
                            End Select
                    
                    End Select
                    
                End If
                   
                bKey = True
                
            End If

        End If

    End With

    KeyPress = True

End Function

Public Function LoadFileTypes() As Boolean
    
    On Error Resume Next

    xmlFileDoc.Load App.Path & "\filetypes.xml"
        
    Set xmlFileTypes = xmlFileDoc.selectNodes("//file")
        
    LoadFileTypes = True

End Function

Public Function OpenFile(sCommand As String, Optional sExtCommand As String, Optional bClearPlaylist As Boolean, Optional bAutoPlay As Boolean = True) As Boolean

    'On Error Resume Next

    Dim sExtension As String
    Dim sFileString As String
        
    sExtension = Mid(sCommand, Len(sCommand) - 2)

    For Each xmlFileType In xmlFileTypes
        
        DoEvents
        
        If sExtension = xmlFileType.Attributes.getNamedItem("extension").Text Then
            
            If xmlFileType.Attributes.getNamedItem("internal").Text = "yes" Then
            
                Select Case xmlFileType.Attributes.getNamedItem("type").Text
                
                    Case "audio"
                    
                        LoadAudio sCommand, , bAutoPlay
                    
                    Case "video"
                    
                        LoadWMPVideo sCommand, bAutoPlay
                        
                    Case "DVD"
                    
                        'load DVD
                    
                    
                End Select
            
            Else
            
                Select Case xmlFileType.Attributes.getNamedItem("type").Text
                
                    Case "shortcut"
                    
                        ShellExecute frmMain.hwnd, "Open", sCommand, vbNullString, "C:\", SW_SHOWNORMAL
                    
                    Case "executable"
                    
                        Shell sCommand, vbMaximizedFocus
                        
                    Case "file"
                    
                        If xmlFileType.Attributes.getNamedItem("hasQuotes").Text = "yes" Then
                        
                            sFileString = xmlFileType.Attributes.getNamedItem("hasQuotes").Text & " " & Chr(34) & sCommand & Chr(34)
                        Else
                        
                            sFileString = xmlFileType.Attributes.getNamedItem("hasQuotes").Text & " " & sCommand
                        
                        End If
                    
                        If sExtCommand <> "" Then
                        
                            sFileString = sFileString & " " & sExtCommand
                            
                        End If
                    
                        Shell sFileString
                    
                End Select
            
            End If
            
            'type
        
        End If
    
    Next
        
    OpenFile = True
    
End Function

Public Function processCommand(ByVal command As Long) As Boolean

    'On Error Resume Next
    
    Dim bProc As Boolean

    bProc = False

    If bRemoteSupport = True Then
    
        Select Case command
        
            Case 1 'back
            
                Do Until KeyPress(vbKeyBack) = True
                    DoEvents
                Loop
                
                
                bProc = True
        
            Case 8 'mute
        
                Do Until VolumeToggleMute = True
                    DoEvents
                Loop
            
                bProc = True
        
            Case 9 'volume down
                
                Do Until VolumeDown = True
                    DoEvents
                Loop
            
                bProc = True
                
            Case 10 'volume up
            
                Do Until VolumeUp = True
                    DoEvents
                Loop
            
                bProc = True
                
            Case 11 'next track
            
                Select Case PlayBack.Source
                
                    Case 0 'music
                    
                        Do Until PlaylistNext = True
                            DoEvents
                        Loop
                        
                    Case 1 'video
                    
                        Do Until SkipForwardLarge = True
                            DoEvents
                        Loop
                    
                End Select
            
                bProc = True
            
            Case 12 'previous track
            
                Select Case PlayBack.Source
                
                    Case 0 'music
                    
                        Do Until PlaylistBack = True
                            DoEvents
                        Loop
                        
                    Case 1 'video
                    
                        Do Until SkipBackLarge = True
                            DoEvents
                        Loop
                    
                End Select
            
                bProc = True
            
            Case 13 'stop
            
                Do Until StopPlayback = True
                    DoEvents
                Loop
            
                bProc = True
            
            Case 14, 46, 47 'play pause, play, pause
            
                Do Until PlayPause = True
                    DoEvents
                Loop
            
                bProc = True
            
            Case 19 'bass down
            
            Case 20 'bass boost
            
            Case 21 'bass up
            
            Case 22 'treble down
            
            Case 23 'treble up
            
            Case 48 'record
            
                'bProc = True
            
            Case 49 'fast forward
            
                Do Until SkipForwardSmall = True
                    DoEvents
                Loop
            
                bProc = True
            
            Case 50 'rewind
            
                Do Until SkipBackSmall = True
                    DoEvents
                Loop
            
                bProc = True
            
            Case 51 'channel up
                        
                Do Until KeyPress(vbKeyPageUp, 0) = True
                    DoEvents
                Loop
                
                bProc = True
                
            Case 52 'channel down
            
                Do Until KeyPress(vbKeyPageDown, 0) = True
                    DoEvents
                Loop
                
                bProc = True
                
            Case Else
                
        End Select
        
    End If
        
    processCommand = bProc

End Function

Public Function VolumeDown() As Boolean

    On Error Resume Next
    
    If (Volume.GetMasterVolumeLevelScalar * 100) < 2 Then
    
        Volume.SetMute 1
    
    Else
    
        Volume.SetMasterVolumeLevelScalar Volume.GetMasterVolumeLevelScalar - 0.02
    
    End If
    
    Do Until ShowVolume(CLng(Volume.GetMasterVolumeLevelScalar * 100)) = True
        DoEvents
    Loop
                
    VolumeDown = True
    
End Function

Public Function VolumeUp() As Boolean
    
    On Error Resume Next
    
    Volume.SetMute 0
            
    Volume.SetMasterVolumeLevelScalar Volume.GetMasterVolumeLevelScalar + 0.02

    Do Until ShowVolume(CLng(Volume.GetMasterVolumeLevelScalar * 100)) = True
        DoEvents
    Loop
                
    VolumeUp = True
    
End Function

Public Function VolumeToggleMute() As Boolean

    On Error Resume Next
    
    If Volume.GetMute = 0 Then
    
        Volume.SetMute 1
        
    Else
    
        Volume.SetMute 0
        
    End If

    Do Until ShowVolume(CLng(Volume.GetMasterVolumeLevelScalar * 100)) = True
        DoEvents
    Loop
                
    VolumeToggleMute = True
    
End Function

Public Function ShowVolume(lVolume As Long) As Boolean

    On Error Resume Next
    
    With frmMain
    
        Dim lVolIcon As Long
         
        Select Case (lVolume)
        
            Case 0
                lVolIcon = 0
        
            Case 1 To 6
                lVolIcon = 1
                
            Case 7 To 12
                lVolIcon = 2
            
            Case 13 To 18
                lVolIcon = 3
            
            Case 19 To 25
                lVolIcon = 4
            
            Case 26 To 31
                lVolIcon = 5
            
            Case 32 To 37
                lVolIcon = 6
            
            Case 38 To 43
                lVolIcon = 7
            
            Case 44 To 50
                lVolIcon = 8
            
            Case 51 To 56
                lVolIcon = 9
            
            Case 57 To 62
                lVolIcon = 10
            
            Case 62 To 68
                lVolIcon = 11
            
            Case 69 To 75
                lVolIcon = 12
            
            Case 76 To 81
                lVolIcon = 13
    
            Case 82 To 87
                lVolIcon = 14
            
            Case 88 To 96
                lVolIcon = 15
            
            Case 97 To 100
                lVolIcon = 16
            
        End Select
    
        .pVolumeIcon.Picture = LoadPictureGDIplus(App.Path & "\Images\Volume\" & lVolIcon & ".png")

        If Volume.GetMute = 1 Then
        
            .lblVolume.Caption = "X"
            
        Else
        
            .lblVolume.Caption = lVolume
        
        End If
        
        lVolumeCount = 0
        
        .pVolumeHolder.ZOrder 0
        
        .pVolumeHolder.Visible = True
        
        .tmrVolume.Enabled = True
        
    End With

    ShowVolume = True

End Function


Public Function HideCursor(bCursor As Boolean) As Boolean

    On Error Resume Next
    
    Select Case bCursor
        Case True
            Do Until ShowCursor(False) < 0
                DoEvents
            Loop
        Case False
            Do Until ShowCursor(True) >= 0
                DoEvents
            Loop
    End Select

    HideCursor = True

End Function

Public Function BeginCloseExe()

    On Error Resume Next
    
    With frmMain
        
        Dim i As Integer
        
        For i = 0 To .pPower.ubound
            .pPower(i).TransparencyPct = 75
            .pPower(i).BackStyleOpaque = False
            .pPower(i).BackColor = eDashProp(lRowSelected).TileColor
        Next
        
        lPower = 2
        
        .pPower(lPower).TransparencyPct = 0
        .pPower(lPower).BackStyleOpaque = True
        
        SetPowerCaption lPower
    
        'Select Case CurrentScreen
        '    Case 0
        '        Do Until ClearDash(True) = True
        '            DoEvents
        '        Loop
        'End Select
        
        .pPowerBack.Visible = Not .pTileOptions.Visible
        
        '.pTopBar.Visible = False
        
        Do Until CaptureScreen(.hDC) = True
            DoEvents
        Loop
        
        .pFocus.Left = .ScaleWidth
        
        For i = 0 To .pPower.ubound
            DoEvents
            .pPower(i).ZOrder 0
            .pPower(i).Visible = True
            .tmrWait.Enabled = True
            Do Until .tmrWait.Enabled = False
                DoEvents
            Loop
        Next
        
        .lblPowerOptions.Visible = True

    End With
                   
End Function

Public Function SetPowerCaption(Index As Long)

    On Error Resume Next
    
    Select Case Index
    
        Case 0 'close app
        
            frmMain.lblPowerOptions.Caption = "Shutdown"
        
        Case 1
        
            frmMain.lblPowerOptions.Caption = "Restart"
        
        Case 2
        
            frmMain.lblPowerOptions.Caption = "Exit"
        
        Case 3
        
            frmMain.lblPowerOptions.Caption = "Sleep"
        
        Case 4
        
            frmMain.lblPowerOptions.Caption = "Log Off"
        
    End Select
        
End Function

Public Function ProcessPower(Index As Long)

    On Error Resume Next
    
    SetPowerCaption lPower
                            
    Sound "yes"
                            
    Select Case Index
    
        Case 0 'shutdown
    
            EnableShutDown
        
            ExitWindowsEx 1, 1
                
        Case 1 'restart
    
            EnableShutDown
            
            ExitWindowsEx 2, 0
            
        Case 2 'close app
        
            Unload frmMain
            
            End
        
        Case 3 'sleep
    
            EnableShutDown
            
            SetSystemPowerState True, True
            
        Case 4 'log off
    
            EnableShutDown
        
            ExitWindowsEx 0, 1
            
    End Select
        

End Function

Public Function ClosePowerOptions() As Boolean

    On Error Resume Next
    
    With frmMain
        
        Dim i As Integer
        
        Sound "closeelement"
        
        .lblPowerOptions.Visible = False
        
        For i = 0 To .pPower.ubound
            DoEvents
            .pPower(i).Visible = False
            .tmrWait.Enabled = True
            Do Until .tmrWait.Enabled = False
                DoEvents
            Loop
        Next
        
        .pFocus.Left = 0
        
        .pTopBar.Visible = bShowTopBar
        
        Select Case CurrentScreen
            Case 0
                Do Until RestoreDash(True) = True
                    DoEvents
                Loop
        End Select
        
    End With
    
    ClosePowerOptions = True

End Function

Public Function PowerKey(KeyCode As Integer, Shift As Integer) As Boolean

    On Error Resume Next
    
    With frmMain

        Select Case KeyCode
        
            Case vbKeyLeft
            
                If lPower > 0 Then
                    
                    .pPower(lPower).TransparencyPct = 75
                    .pPower(lPower).BackStyleOpaque = False
                    
                    lPower = lPower - 1
                    
                    Sound "select"
                    
                    .pPower(lPower).TransparencyPct = 0
                    .pPower(lPower).BackStyleOpaque = True
                    
                    SetPowerCaption lPower
                    
                Else
                    
                    Sound "listend"
                    
                End If
            
            Case vbKeyRight
            
                If lPower < .pPower.ubound Then
                    
                    .pPower(lPower).TransparencyPct = 75
                    .pPower(lPower).BackStyleOpaque = False
                    
                    lPower = lPower + 1
                    
                    Sound "select"
                    
                    .pPower(lPower).TransparencyPct = 0
                    .pPower(lPower).BackStyleOpaque = True
                    
                    SetPowerCaption lPower
                    
                Else
                    
                    Sound "listend"
                    
                End If
            
            Case vbKeyBack
                
                ClosePowerOptions

            Case vbKeyReturn
        
                ProcessPower lPower
        
        End Select
            
    End With
    
    PowerKey = True
            
End Function

Public Function LoadGame(File As String, Name As String, Thumb As String, isEmulator As String, EmulatorPath As String, hasQuotes As String, SubCaption As String, Interval As String) As Boolean

    On Error Resume Next
    
    Dim sExecutePath As String
    Dim fso As New FileSystemObject
    Dim sSource As String
    
    'loadfile##file##name##thumb##isemulator##emulatorpath##hasquotes##subcaption

    sSource = "game##" & File & "##" & Name & "##" & Thumb & "##" & isEmulator & "##" & EmulatorPath & "##" & hasQuotes & "##" & SubCaption & "##" & Interval
        
    If isEmulator = "True" Then
    
        sProcessPath = fso.GetFileName(EmulatorPath)
    
        If hasQuotes = "True" Then

            sExecutePath = EmulatorPath & " " & Chr(34) & File & Chr(34)
            
        Else
        
            sExecutePath = EmulatorPath & " " & File
            
        End If
        
        Shell sExecutePath, vbNormalFocus
    
    Else
    
        If Mid(File, Len(File) - 2) = "lnk" Then
        
            sExecutePath = GetTarget(File)
        
            ShellExecute hwnd, "open", File, vbNullString, vbNullString, SW_SHOWNORMAL
            
        Else
        
            sExecutePath = File
           
            Shell sExecutePath, vbNormalFocus
    
        End If

       sProcessPath = fso.GetFileName(sExecutePath)
            
    End If
    
    frmMain.tmrPlayback.Enabled = False
    frmMain.tmrMouse.Enabled = False
    frmMain.tmrMouseOut.Enabled = False
     
    bControllerBypass = True
    
    HideCursor False
    
    lGameProcessInterval = CLng(Interval)
    
    frmMain.tmrCheckForProcess.Interval = lGameProcessInterval
        
    frmMain.tmrCheckForProcess.Enabled = True
    
    Select Case frmMain.WindowState
        Case vbNormal
            ResizeState = 0
        Case vbMinimized
            ResizeState = 1
        Case vbMaximized
            ResizeState = 2
    End Select
    
    frmMain.WindowState = vbMinimized

    If lGamesRow > -1 Then
    
        Do Until UpdateTile(lGamesRow, 0, sSource, vbNullString, "Last Played | " & Name, Thumb, True) = True
            DoEvents
        Loop
    
    End If
    
    LoadGame = True

End Function

Private Function GetTarget(strPath As String) As String
    
    On Error Resume Next
    
    'Gets target path from a shortcut file
    
    On Error GoTo Error_Loading
    
    Dim wshShell As Object
    
    Dim wshLink As Object
    
    Set wshShell = CreateObject("WScript.Shell")
    
    Set wshLink = wshShell.CreateShortcut(strPath)
    
    GetTarget = wshLink.TargetPath
    
    Set wshLink = Nothing
    
    Set wshShell = Nothing
    
    Exit Function

Error_Loading:
    
    GetTarget = "Error occured."

End Function

Public Function EndProcessCheck()

    On Error Resume Next
    
    sProcessPath = ""
    
    frmMain.tmrCheckForProcess.Enabled = False

    frmMain.WindowState = ResizeState
    
    bControllerBypass = False
    frmMain.tmrMouseOut.Enabled = bMouseOut
    frmMain.tmrMouse.Enabled = True
    
    Do Until ReselectDash = True
        DoEvents
    Loop
    
    frmMain.pFocus.SetFocus

End Function

Public Function GetTime(lH As Long, lM As Long) As String

    On Error Resume Next
    
    Dim sH As String
    Dim sM As String

    If lH < 10 Then
        sH = "0" & lH
    Else
        sH = lH
    End If

    If lM < 10 Then
        sM = "0" & lM
    Else
        sM = lM
    End If
    
    GetTime = sH & ":" & sM
    
End Function

Public Function IsProcessRunning(ByVal sProcess As String) As Boolean
    
Dim processInfo As PROCESSENTRY32   ' information about a process in that list
    Dim hSnapshot   As Long             ' handle to the snapshot of the process list
    Dim success     As Long             ' success of having gotten info on another process
    Dim retval      As Long             ' generic return value
    Dim exeName     As String           ' filename of the process
    
    ' First, make a snapshot of the current process list.
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    ' Get information about the first process in the list.
    processInfo.dwSize = Len(processInfo)
    success = Process32First(hSnapshot, processInfo)
    
    ' Make sure a handle was returned.
    If hSnapshot = -1 Then
        'no snapshot return to main
        IsProcessRunning = False
        Exit Function
    Else
        ' Loop for each process on the list.
        Do While success <> 0
            ' Extract the filename of the process (i.e., remove the empty space)
            exeName = Left(processInfo.szExeFile, InStr(processInfo.szExeFile, vbNullChar) - 1)
            
            ' check the process name
            
            If UCase(exeName) = UCase(sProcess) Then
                IsProcessRunning = True
                Exit Function
            End If
            
            ' Get information about the next process, if there is one.
            processInfo.dwSize = Len(processInfo)
            success = Process32Next(hSnapshot, processInfo)
        Loop
        
        ' Destroy the snapshot, now that we no longer need it.
        retval = CloseHandle(hSnapshot)
    End If


End Function
 
Public Sub TerminateProcess(app_exe As String)
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
        Process.Terminate
    Next
End Sub

Public Function InitSounds() As Boolean

    On Error Resume Next
    
    With frmMain
    
        Dim i As Long
        Dim sFile As String
        
        For i = 1 To 6
            
            Load .cSound(i)
            
            Select Case i
            
                Case 1 'close
                              
                    sFile = App.Path & "\sounds\close.wav"
        
                Case 2 ' focus
                              
                    sFile = App.Path & "\sounds\focus.wav"
        
                Case 3 'no
                              
                    sFile = App.Path & "\sounds\no.wav"
        
                Case 4 'notify
                              
                    sFile = App.Path & "\sounds\notify.wav"
        
                Case 5 'select
                              
                    sFile = App.Path & "\sounds\select.wav"
        
                Case 6 'yes
                              
                    sFile = App.Path & "\sounds\yes.wav"
        
            End Select
                
            .cSound(i).settings.autoStart = False
                           
            .cSound(i).URL = sFile
                    
        Next
    
    End With
    
    InitSounds = True
    
End Function

Public Function Sound(sSound As String)

    On Error Resume Next
    
    If bSounds = True Then
    
        Dim i As Long
        Dim bTemp As Boolean
        
        If PlayBack.State <> ePlaying Then
    
            Select Case sSound
            
                Case "openapp", "openelement"
                
                    i = 5
                
                Case "closeapp", "closeelement"
                
                    i = 1
               
                Case "yes"
                
                    i = 6
                
                Case "no"
                
                    i = 3
                
                Case "select"
                
                    i = 2
                    
                Case "notify"
                
                    i = 4
                    
            End Select
            
            If i <> -1 Then
            
                frmMain.cSound(i).Controls.play
                
            End If
            
        End If

    End If

End Function

Public Function CloseSounds() As Boolean

    On Error Resume Next
    
    With frmMain
    
        Dim i As Long
        
        For i = 1 To .cSound.ubound
            'cSound(i).Controls.stop
        Next
        
    End With
    
    CloseSounds = True
    
End Function
