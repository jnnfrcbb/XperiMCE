VERSION 5.00
Begin VB.UserControl PlayerControl 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   420
   Begin VB.Timer tmrPlaying 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   300
      Top             =   0
   End
   Begin VB.Timer tmrNetRadio 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -300
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PMP"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "PlayerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Enum ePlaySource
    eLocalFile = 0
    eNetRadio = 1
    eStreamingFile = 2
    eCD = 3
End Enum

Public Enum eStatus
    eError = 0
    eWaiting = 1
    eInitialised = 2
    eLoaded = 3
    ePlaying = 4
    ePaused = 5
    eStopped = 6
    eShutdown = 7
End Enum

Public Enum eSpeakers
    stereo = 0
    Surround = 1
End Enum

Event StreamingLoaded()

Event StreamUpdate(url As String, sName As String, sBPS As String, sGenre, sSong As String)

Event StreamingEnded()

Event PlayerShutdown()

Event LevelChange(Left As Long, Right As Long)

Event CDTitle(title As String)

Event CDArtist(artist As String)

Event CDTrackDetails(index As Long, TrackTitle As String, TrackArtist As String)

Private PlaySource As ePlaySource
    
Private sCurrentFile As String

Private bFileLoaded As Boolean

Public bInitialised As Boolean

Private eCurrentStatus As Long

Public sError As String

Private eSpeakerMode As eSpeakers

Private Speaker(3) As Long

Private lCurrentVolume As Double

Private lSystemVolume As Double

Private sStreamSong As String

Private sStreamName As String

Private sStreamBPS As String

Private sStreamGenre As String

Private bStreamProxyRequired As Boolean

Private sStreamProxyUserName As String

Private bStreamSave As Boolean

Private CDi() As BASS_CD_INFO
Private sCDDriveLetter As String
Private bCDLookup As Boolean
Private CDDrive As Long
Private sCDTitle As String
Private sCDArtist As String
Private sCDYear As String
Private sCDTrackName() As String
Private sCDTrackArtist() As String
Private sCDFreeDBID As String

Private fx() As Long
Private fxValue() As Long
Private fxHz() As Long

Private fxSide() As Long
Private fxSideValue() As Long
Private fxSideHz() As Long

Private fxCentre() As Long
Private fxCentreValue() As Long
Private fxCentreHz() As Long

Private fxRear() As Long
Private fxRearValue() As Long
Private fxRearHz() As Long

Private lEQChannelCount As Long
Private chan As Long

Public sURL As String

Public Property Get StreamSong() As String

    StreamSong = sStreamSong

End Property

Public Property Get StreamName() As String

    StreamName = sStreamName

End Property

Public Property Get StreamBPS() As String

    StreamBPS = sStreamBPS

End Property

Public Property Let StreamSong(sString As String)

    sStreamSong = sString

End Property

Public Property Let StreamName(sString As String)

    sStreamName = sString

End Property

Public Property Let StreamBPS(sString As String)

    sStreamBPS = sString

End Property


Public Property Let StreamProxy(ByVal bProxy As Boolean)

    bStreamProxyRequired = bProxy

    If bProxy = False Then
    
        Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, vbNullString)  ' disable proxy
    
    Else
        
        Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, VarPtr(proxy(0))) ' enable proxy
    
    End If

End Property

Public Property Let StreamProxyUserName(ByVal sProxy As String)

    sStreamProxyUserName = sProxy

End Property

Public Property Let StreamSave(ByVal bSave As Boolean)
    
    DoDownload = bSave

End Property


Public Property Get PlayerError() As String

    PlayerError = sError

End Property

Public Property Get PlayerStatus() As eStatus

    PlayerStatus = eCurrentStatus

End Property

Public Property Get PlayerVolume() As Double

    PlayerVolume = CLng((BASS_GetConfig(BASS_CONFIG_GVOL_STREAM) / 100))

End Property

Public Property Let PlayerVolume(ByVal lVol As Double)

    Dim NewVol As Double

    If lVol > 100 Then
    
        NewVol = 100
        
    ElseIf lVol < 0 Then
    
        NewVol = 0
        
    Else
        
        NewVol = lVol
        
    End If
    
    NewVol = lVol * 100
    
    Call BASS_SetConfig(BASS_CONFIG_GVOL_STREAM, NewVol)
    
    'Dim i As Integer
    
    'If lVol > 1 Then
    '    lVol = 1
    'ElseIf lVol < 0 Then
    '    lVol = 0
    'End If
    
    'If bInitialised = True Then
    '    For i = 0 To UBound(StreamHandle)
    '        BASS_ChannelSetAttribute StreamHandle(i), BASS_ATTRIB_VOL, NewVol
    '    Next
    'End If
    
    lCurrentVolume = NewVol

End Property

Public Property Get SystemVolume() As Long

    Dim lVol As Double
    
    lVol = BASS_GetVolume
    
    If lVol < 1 Then
        
        If lVol > 0.1 Then
    
            lVol = CLng(Mid(lVol, 3, 2))
            
        Else
        
            lVol = 0
            
        End If
        
    Else
        
        lVol = 100
        
    End If

    SystemVolume = lVol

End Property

Public Property Let SystemVolume(ByVal lVol As Long)

    Dim NewVol As Double

    If lVol > 1 Then
    
        NewVol = 1
        
    ElseIf lVol < 0 Then
    
        NewVol = 0
        
    Else
        
        NewVol = lVol
        
    End If
    
    Call BASS_SetVolume(NewVol)
    
    lSystemVolume = NewVol

End Property

Public Property Get FilePosition() As Double
    
    If bFileLoaded = True Then
    
        FilePosition = BASS_ChannelBytes2Seconds(StreamHandle(0), BASS_ChannelGetPosition(StreamHandle(0), 0))
            
    Else
    
        FilePosition = 0
        
    End If

End Property

Public Property Get FilePositionString() As String
    
    If bFileLoaded = True Then
    
        FilePositionString = GetTimeString(BASS_ChannelBytes2Seconds(StreamHandle(0), BASS_ChannelGetPosition(StreamHandle(0), 0)))
            
    Else
    
        FilePositionString = "00:00"
        
    End If

End Function

Public Property Let FilePosition(lNewPos As Double)
    
    Static i As Integer
    
    If bFileLoaded = True Then
    
        For i = 0 To UBound(StreamHandle)
        
            If BASS_ChannelSeconds2Bytes(StreamHandle(i), lNewPos) < 0 Then
            
                sError = "Error setting position"
                
                eCurrentStatus = eError
                
            End If
            
        Next i
    
    End If

End Property

Public Property Get FileDuration() As Double
    
    If bFileLoaded = True Then
    
        FileDuration = BASS_ChannelBytes2Seconds(StreamHandle(0), BASS_ChannelGetLength(StreamHandle(0), 0))
        
    Else
    
        FileDuration = 0
        
    End If

End Property

Public Property Get FileDurationString() As String
    
    If bFileLoaded = True Then
        
        FileDurationString = GetTimeString(BASS_ChannelBytes2Seconds(StreamHandle(0), BASS_ChannelGetLength(StreamHandle(0), 0)))
    
    Else
    
        FileDurationString = "00:00"
        
    End If

End Property

Private Function GetTimeString(dSeconds As Double) As String
    
    Dim lMinutes As Long
    Dim lSeconds As Long
    
    On Error Resume Next
    
    GetTimeString = dSeconds \ 60 & ":" & Format(dSeconds Mod 60, "00")

End Function

Public Function PlayerInitialise(eMode As ePlaySource) As Boolean

    SetControl Me

    If bFileLoaded = False Then
        
        ' check the correct BASS was loaded
        If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        
            sError = "An incorrect version of BASS.DLL was loaded"
            
            eCurrentStatus = eError
            
            bInitialised = False
            
            Exit Function
        
        End If
        
        'initialise on default device
        If (BASS_Init(-1, 44100, 0, UserControl.hwnd, 0) = 0) Then
            
            sError = "Can't initialise digital sound system"
            
            eCurrentStatus = eError
            
            bInitialised = False
            
            Exit Function
            
        End If
                    
        If SystemSpeakers < 4 Then
    
            Call BASS_Free
    
            ' Initialize output - default device, 44100hz, STEREO, 16 bits
        
            If BASS_Init(-1, 44100, BASS_DEVICE_NOSPEAKER, UserControl.hwnd, 0) = BASSFALSE Then
                
                sError = "Can't initialize digital sound system"
                
                eCurrentStatus = eError
                
                bInitialised = False
                
                Exit Function
            
            End If
    
            eSpeakerMode = stereo
            
        ElseIf SystemSpeakers >= 4 Then
        
            If eSpeakerMode = stereo Then
                
                Debug.Print "stereo playback"
                
                Call BASS_Free
        
                ' Initialize output - default device, 44100hz, STEREO, 16 bits
            
                If BASS_Init(-1, 44100, BASS_DEVICE_NOSPEAKER, UserControl.hwnd, 0) = BASSFALSE Then
                    
                    sError = "Can't initialize digital sound system"
                    
                    eCurrentStatus = eError
                    
                    bInitialised = False
                    
                    Exit Function
                
                End If
        
    
            Else
        
                Debug.Print "surround playback"
                
                If SystemSpeakers = 4 Then '3.1
                    Speaker(0) = BASS_SPEAKER_FRONT 'front speakers
                    Speaker(1) = -1 'no side speakers
                    Speaker(2) = BASS_SPEAKER_CENLFE 'centre/sub
                    Speaker(3) = -1 'no rear speakers
                ElseIf SystemSpeakers = 6 Then '5.1
                    Speaker(0) = BASS_SPEAKER_FRONT 'front speakers
                    Speaker(1) = BASS_SPEAKER_REAR 'rear/side speakers
                    Speaker(2) = BASS_SPEAKER_CENLFE 'centre/sub
                    Speaker(3) = -1 'no rear speakers
                ElseIf SystemSpeakers = 8 Then '7.1
                    Speaker(0) = BASS_SPEAKER_FRONT 'front speakers
                    Speaker(1) = BASS_SPEAKER_REAR 'rear/side speakers
                    Speaker(2) = BASS_SPEAKER_CENLFE 'centre/sub
                    Speaker(3) = BASS_SPEAKER_REAR2 'rear centre (7.1)
                End If
                        
                Call BASS_Free
            
                ' Initialize output - default device, 44100hz, SURROUND, 16 bits
        
                If BASS_Init(-1, 44100, BASS_DEVICE_SPEAKERS, UserControl.hwnd, 0) = BASSFALSE Then
                    
                    sError = "Can't initialize digital sound system"
                    
                    eCurrentStatus = eError
                    
                    bInitialised = False
                    
                    Exit Function
                
                End If
                
            End If
            
        End If
        
        sError = ""
        
        eCurrentStatus = eInitialised
        
        bInitialised = True
            
        If eMode = 0 Then
        
            Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 0) ' disable playlist processing
            
        ElseIf eMode = 1 Then
        
            Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 1) ' enable playlist processing
            Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) ' minimize automatic pre-buffering, so we can do it (and display it) instead
            Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, VarPtr(proxy(0)))  ' setup proxy server location

            Set WriteFile = New clsFileIo
            
            cthread = 0

        ElseIf eMode = 2 Then
            
            Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 1) ' enable playlist processing
            Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) ' minimize automatic pre-buffering, so we can do it (and display it) instead
            Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, VarPtr(proxy(0)))  ' setup proxy server location

            Set WriteFile = New clsFileIo
            
            cthread = 0
        
        ElseIf eMode = 3 Then

            'Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 0) ' disable playlist processing

        End If
        
        Call BASS_SetConfig(BASS_CONFIG_GVOL_STREAM, lCurrentVolume)
    
    ElseIf bFileLoaded = True Then
    
        PlayerShutdown
        
        PlayerInitialise eMode
        
    End If

    PlaySource = eMode

    PlayerInitialise = bInitialised

End Function

Public Function PlayerShutdown()

    Static i As Integer

    If PlaySource = eLocalFile Then
    
        If bFileLoaded = True Then
        
            For i = 0 To UBound(StreamHandle)
            
                Call BASS_StreamFree(StreamHandle(i))
                
            Next i
        
            bFileLoaded = False
        
        End If
    
    ElseIf PlaySource = eNetRadio Then
    
        sStreamName = ""
        sStreamBPS = ""
        sStreamGenre = ""
        sStreamSong = ""
        
        Call BASS_StreamFree(StreamHandle(0))
        
        bFileLoaded = False
        
        tmrNetRadio.Enabled = False
    
    ElseIf PlaySource = eCD Then
    
        Call BASS_CD_Release(CDDrive)
        
        bFileLoaded = False
        
    End If
    
    tmrPlaying.Enabled = False

    ' Close sound system and release everything
    Call BASS_Free
    
    eCurrentStatus = eShutdown
    
    sError = ""

End Function

Public Function FileLoad(sFile As String, eMode As ePlaySource, Optional CDTrack As Long) As Boolean

    Static i As Integer
    
    Dim threadid As Long
    
    Dim sType As String

    PlayerShutdown
    
    If PlayerInitialise(eMode) = True Then
        
        If eMode = eLocalFile Then
        
            Call BASS_PluginLoad(App.Path & "\bass_alac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bass_aac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\basswma.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bassflac.dll", 0)
        
            If eSpeakerMode = stereo Then
        
                ReDim StreamHandle(0)
                
                'StreamHandle(0) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
                
                StreamHandle(0) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_SAMPLE_FX)
                
            ElseIf eSpeakerMode = Surround Then
            
                ReDim StreamHandle(UBound(Speaker))
            
                For i = 0 To UBound(Speaker)
                
                    If Speaker(i) <> -1 Then
                    
                        StreamHandle(i) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, Speaker(i) Or BASS_SAMPLE_FX)
                        
                        If BASS_ErrorGetCode <> 0 Then
                            sError = BASS_ErrorGetCode
                        Else
                            sError = ""
                        End If
                        
                    Else
                    
                        StreamHandle(i) = -1
                        
                    End If
                    
                Next i
                
            End If
            
            If StreamHandle(0) = 0 Then
                
                sCurrentFile = ""
                
                bFileLoaded = False
                
                eCurrentStatus = eError
                
                FileLoad = False
                
                sError = "Can't open stream"
                    
            Else
                
                sCurrentFile = sFile
        
                bFileLoaded = True
                
                eCurrentStatus = eLoaded
                
                sError = ""
                
                FileLoad = True
                
                'FilePlay
                        
                'Do Until SyncStreams = True
                '    DoEvents
                'Loop
    
            End If
    
        ElseIf eMode = eNetRadio Then
        
            ReDim StreamHandle(0)
            
            Call BASS_PluginLoad(App.Path & "\bass_alac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bass_aac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\basswma.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bassflac.dll", 0)
            
            sURL = sFile
            
            If (cthread) Then   ' already connecting
            
                Call Beep
            
            Else
                
                Call CopyMemory(proxy(0), ByVal sStreamProxyUserName, Len(sStreamProxyUserName))   ' get proxy server
        
                ' open URL in a new thread (so that main thread is free)
                
                cthread = CreateThread(ByVal 0&, 0, AddressOf OpenURL, 0, 0, threadid)   ' threadid param required on win9x
            
            End If
            
            sCurrentFile = sFile
            
            bFileLoaded = True
            
            eCurrentStatus = ePlaying
            
            sError = ""
            
            SetTimer True
            
        ElseIf eMode = eStreamingFile Then
        
            ReDim StreamHandle(0)
            
            sType = Mid(sFile, InStrRev(sFile, ".") + 1, 3)
                        
            Call BASS_PluginLoad(App.Path & "\bass_alac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bass_aac.dll", 0)
            Call BASS_PluginLoad(App.Path & "\basswma.dll", 0)
            Call BASS_PluginLoad(App.Path & "\bassflac.dll", 0)
                
            sURL = sFile
            
            If (cthread) Then   ' already connecting
            
                Call Beep
            
            Else
                
                Call CopyMemory(proxy(0), ByVal sStreamProxyUserName, Len(sStreamProxyUserName))   ' get proxy server
        
                ' open URL in a new thread (so that main thread is free)
                
                cthread = CreateThread(ByVal 0&, 0, AddressOf OpenURL, 0, 0, threadid)   ' threadid param required on win9x
            
            End If
            
            sCurrentFile = sFile
            
            bFileLoaded = True
            
            eCurrentStatus = ePlaying
            
            sError = ""
            
            SetTimer True
            
        ElseIf eMode = eCD Then
        
            ReDim StreamHandle(0)
            
            CDLoad sFile, bCDLookup
            
            StreamHandle(0) = BASS_CD_StreamCreate(CDDrive, CDTrack, 0)

            Call BASS_ChannelPlay(StreamHandle(0), BASSFALSE)
            
            bFileLoaded = True
            
        End If
        
        PlaySource = eMode
    
    End If
    
    For i = 0 To lEQChannelCount
        SetEQ CLng(i)
    Next
                  
    tmrPlaying.Enabled = True

End Function

Public Function FilePause()

    Static i As Integer
    
    If bFileLoaded = True Then

        For i = 0 To UBound(StreamHandle)
    
            If StreamHandle(i) <> -1 Then
    
                Call BASS_ChannelPause(StreamHandle(i))
            
            End If
            
        Next i
            
        eCurrentStatus = ePaused
            
        'Do Until SyncStreams = True
        '    DoEvents
        'Loop
        
    End If
        
End Function

Public Function FilePlay()

    Dim i As Integer

    If bFileLoaded = True Then
    
        eCurrentStatus = ePlaying
            
        sError = ""
        
        'Do Until SyncStreams = True
        '    DoEvents
        'Loop
        
        For i = 0 To UBound(StreamHandle)
    
            If StreamHandle(i) <> -1 Then
        
                If (BASS_ChannelPlay(StreamHandle(i), BASSFALSE) = 0) Then
                
                    'eCurrentStatus = eError
                
                    sError = "Can't play stream (" & i & ")"
                    
                End If
                
            End If
            
        Next i
    
        eCurrentStatus = ePlaying
            
    End If
        

End Function

Public Function FileRestart()

    Static i As Integer

    If bFileLoaded = True Then
    
        eCurrentStatus = eStopped
        
        'Do Until SyncStreams = True
        '    DoEvents
        'Loop
        
        For i = 0 To UBound(StreamHandle)
    
            If StreamHandle(i) <> -1 Then
    
                Call BASS_ChannelPlay(StreamHandle(i), BASSTRUE)
                
            End If
        
        Next i
        
        eCurrentStatus = ePlaying
        
    End If

End Function

Public Function FileStop()

    Static i As Integer

    If bFileLoaded = True Then
    
        For i = 0 To UBound(StreamHandle)
    
            If StreamHandle(i) <> -1 Then
    
                Call BASS_ChannelStop(StreamHandle(i))
            
            End If
            
        Next i
        
        'Do Until SyncStreams = True
        '    DoEvents
        'Loop
        
        eCurrentStatus = eStopped
        
    End If

End Function

Public Property Get SystemSpeakers() As Long

    Dim BassInfo As BASS_INFO
    
    Call BASS_GetInfo(BassInfo)
    
    SystemSpeakers = BassInfo.speakers

End Property

Private Sub tmrNetRadio_Timer()

    Dim progress As Long
    
    progress = BASS_StreamGetFilePosition(StreamHandle(0), BASS_FILEPOS_BUFFER) * 100 / BASS_StreamGetFilePosition(StreamHandle(0), BASS_FILEPOS_END)    ' percentage of buffer filled
    
    If (progress >= 75 Or BASS_StreamGetFilePosition(StreamHandle(0), BASS_FILEPOS_CONNECTED) = 0) Then ' over 75% full (or end of download)

        ' get the broadcast name and bitrate
        Dim icyPtr As Long
        icyPtr = BASS_ChannelGetTags(StreamHandle(0), BASS_TAG_ICY)
        
        Call DoMeta
        
        If (icyPtr = 0) Then icyPtr = BASS_ChannelGetTags(StreamHandle(0), BASS_TAG_HTTP) ' no ICY tags, try HTTP
        
        If (icyPtr) Then
            
            Dim icyStr As String
            
            Do
                
                icyStr = VBStrFromAnsiPtr(icyPtr)
                
                Debug.Print "ICY-" & icyStr
                
                icyPtr = icyPtr + Len(icyStr) + 1
                sStreamName = IIf(LCase(Mid(icyStr, 1, 9)) = "icy-name:", Mid(icyStr, 10), sStreamName)
                sStreamBPS = IIf(LCase(Mid(icyStr, 1, 7)) = "icy-br:", "Bitrate: " & Mid(icyStr, 8), sStreamBPS)
                sStreamGenre = IIf(LCase(Mid(icyStr, 1, 10)) = "icy-Genre:", Mid(icyStr, 11), sStreamGenre)
                
                ' NOTE: you can get more ICY info like: icy-genre:, icy-url:... :)
            
            Loop While (icyStr <> "")
        
            Select Case LCase(Mid(sStreamName, 2))
            
                Case "absoluteradioir2"
                
                    sStreamName = "Absolute Radio"
        
                Case "absolutecrir2"
                
                    sStreamName = "Absolute Classic Rock"
        
                Case "absolute60sir"
                
                    sStreamName = "Absolute 60s"
        
                Case "absolute70sir"
                
                    sStreamName = "Absolute 70s"
        
                Case "absolute80sir2"
                
                    sStreamName = "Absolute 80s"
        
                Case "absolute90sir"
                
                    sStreamName = "Absolute 90s"
        
                Case "absolute00sir2"
                
                    sStreamName = "Absolute 00s"
        
            End Select
        
        End If

        ' get the stream title and set sync for subsequent titles
        Call DoMeta
    
        RaiseEvent StreamUpdate(sCurrentFile, sStreamName, sStreamBPS, sStreamGenre, sStreamSong)

        Call BASS_ChannelSetSync(StreamHandle(0), BASS_SYNC_META, 0, AddressOf MetaSync, 0)
        
        ' set sync for end of stream
        Call BASS_ChannelSetSync(StreamHandle(0), BASS_SYNC_END, 0, AddressOf EndSync, 0)
        
        ' play it!
        Call BASS_ChannelPlay(StreamHandle(0), BASSFALSE)
        
        RaiseEvent StreamingLoaded
        
        tmrNetRadio.Enabled = False ' finished prebuffering, stop monitoring
        
    Else
        
        sStreamName = "buffering... " & progress & "%"
        
    End If

    If PlaySource = eNetRadio Then
    
        RaiseEvent StreamUpdate(sCurrentFile, sStreamName, sStreamBPS, sStreamGenre, sStreamSong)
        
    End If
    
End Sub

Private Sub tmrPlaying_Timer()

    On Local Error Resume Next
    ' update levels
    Static updatecount As Long, levl As Long, levr As Long
    Dim level As Long
    
    level = BASS_ChannelGetLevel(StreamHandle(0))
    levl = levl - 1500
    If (levl < 0) Then levl = 0
    levr = levr - 1500
    If (levr < 0) Then levr = 0
    If (level <> -1) Then
        If (levl < LoWord(level)) Then levl = LoWord(level)
        If (levr < HiWord(level)) Then levr = HiWord(level)
    End If
    
    RaiseEvent LevelChange(levl, levr)
    
End Sub

Private Sub UserControl_Initialize()

bFileLoaded = False

lEQChannelCount = 0

lCurrentVolume = 10000

eSpeakerMode = stereo

ReDim StreamHandle(0)

ReDim fx(0)
ReDim fxHz(0)
ReDim fxValue(0)

ReDim fxSide(0)
ReDim fxSideHz(0)
ReDim fxSideValue(0)

ReDim fxCentre(0)
ReDim fxCentreHz(0)
ReDim fxCentreValue(0)

ReDim fxRear(0)
ReDim fxRearHz(0)
ReDim fxRearValue(0)

End Sub

Private Sub UserControl_InitProperties()

eCurrentStatus = eWaiting

SetControl Me

End Sub

Public Function SetTimer(bEnabled As Boolean)

    tmrNetRadio.Enabled = bEnabled

End Function

Public Function ClosePlayer()

FileStop

PlayerShutdown

RaiseEvent PlayerShutdown

End Function

Public Function EQChannelCount(lCount As Long)

    Static i As Integer

    lEQChannelCount = lCount - 1
    
    ReDim Preserve fx(lEQChannelCount)
    ReDim Preserve fxHz(lEQChannelCount)
    ReDim Preserve fxValue(lEQChannelCount)
    
    ReDim Preserve fxSide(lEQChannelCount)
    ReDim Preserve fxSideHz(lEQChannelCount)
    ReDim Preserve fxSideValue(lEQChannelCount)
    
    ReDim Preserve fxCentre(lEQChannelCount)
    ReDim Preserve fxCentreHz(lEQChannelCount)
    ReDim Preserve fxCentreValue(lEQChannelCount)
    
    ReDim Preserve fxRear(lEQChannelCount)
    ReDim Preserve fxRearHz(lEQChannelCount)
    ReDim Preserve fxRearValue(lEQChannelCount)
    
    For i = 0 To lEQChannelCount
        fx(i) = BASS_ChannelSetFX(StreamHandle(0), BASS_FX_DX8_PARAMEQ, 0)
        fxValue(i) = 0
        fxHz(i) = 0
        
        If eSpeakerMode = Surround Then
        
            If Speaker(1) <> -1 Then
                fxSide(i) = BASS_ChannelSetFX(StreamHandle(1), BASS_FX_DX8_PARAMEQ, 0)
                fxSideValue(i) = 0
                fxSideHz(i) = 0
            End If
            
            If Speaker(2) <> -1 Then
                fxCentre(i) = BASS_ChannelSetFX(StreamHandle(2), BASS_FX_DX8_PARAMEQ, 0)
                fxCentreValue(i) = 0
                fxCentreHz(i) = 0
            End If
            
            If Speaker(3) <> -1 Then
                fxRear(i) = BASS_ChannelSetFX(StreamHandle(3), BASS_FX_DX8_PARAMEQ, 0)
                fxRearValue(i) = 0
                fxRearHz(i) = 0
            End If
        End If
    Next
    
End Function

Public Function EQSetChannel(index As Long, lHz As Long, lValue As Long)
    
    Debug.Print "index: " & index & " | lHz: " & lHz & " | Value: " & lValue

    Select Case lHz
        Case Is < 60
            lHz = 60
        Case Is > 16000
            lHz = 16000
    End Select
    
    Select Case lValue
        Case Is < -15
            lValue = -15
        Case Is > 15
            lValue = 15
    End Select
    
    fxHz(index) = lHz
    fxValue(index) = lValue

    fxSideHz(index) = lHz
    fxSideValue(index) = lValue

    fxCentreHz(index) = lHz
    fxCentreValue(index) = lValue

    fxRearHz(index) = lHz
    fxRearValue(index) = lValue
    
    If bFileLoaded = True Then
        Call UpdateFX(index)
    End If
    
End Function

Private Function SetEQ(index As Long)

    Dim p As BASS_DX8_PARAMEQ
    
    fx(index) = BASS_ChannelSetFX(StreamHandle(0), BASS_FX_DX8_PARAMEQ, 0)
    p.fGain = fxValue(index)
    p.fBandwidth = 18
    p.fCenter = fxHz(index)
    Call BASS_FXSetParameters(fx(index), p)
    
    If eSpeakerMode = Surround Then
        If UBound(StreamHandle) > 0 Then
        
            If Speaker(1) <> -1 Then
                fxSide(index) = BASS_ChannelSetFX(StreamHandle(1), BASS_FX_DX8_PARAMEQ, 0)
                p.fGain = fxSideValue(index)
                p.fBandwidth = 18
                p.fCenter = fxSideHz(index)
                Call BASS_FXSetParameters(fxSide(index), p)
            End If
            
            If Speaker(2) <> -1 Then
                fxCentre(index) = BASS_ChannelSetFX(StreamHandle(2), BASS_FX_DX8_PARAMEQ, 0)
                p.fGain = fxCentreValue(index)
                p.fBandwidth = 18
                p.fCenter = fxCentreHz(index)
                Call BASS_FXSetParameters(fxCentre(index), p)
            End If
            
            If Speaker(3) <> -1 Then
                fxRear(index) = BASS_ChannelSetFX(StreamHandle(3), BASS_FX_DX8_PARAMEQ, 0)
                p.fGain = fxRearValue(index)
                p.fBandwidth = 18
                p.fCenter = fxRearHz(index)
                Call BASS_FXSetParameters(fxRear(index), p)
            End If
        End If
    End If
    
    If bFileLoaded = True Then
        Call UpdateFX(index)
    End If
    
End Function

Private Function UpdateFX(ByVal b As Long)

    Dim v As Integer
    Dim p As BASS_DX8_PARAMEQ
    
    If (b <= lEQChannelCount) Then
        v = fxValue(b)
        Call BASS_FXGetParameters(fx(b), p)
        p.fGain = v '10# - v
        Call BASS_FXSetParameters(fx(b), p)
        
        If eSpeakerMode = Surround Then
            If UBound(StreamHandle) > 0 Then
                v = fxSideValue(b)
                Call BASS_FXGetParameters(fxSide(b), p)
                p.fGain = v '10# - v
                Call BASS_FXSetParameters(fxSide(b), p)
                
                v = fxCentreValue(b)
                Call BASS_FXGetParameters(fxCentre(b), p)
                p.fGain = v '10# - v
                Call BASS_FXSetParameters(fxCentre(b), p)
                
                v = fxRearValue(b)
                Call BASS_FXGetParameters(fxRear(b), p)
                p.fGain = v '10# - v
                Call BASS_FXSetParameters(fxRear(b), p)
            End If
        End If
    End If
    
End Function

Public Function CDLoad(DriveLetter As String, Optional bCDDBLookup) As Boolean

    Static i As Integer
    
    bCDLookup = bCDDBLookup
    
    sCDDriveLetter = DriveLetter

    Call BASS_PluginLoad(App.Path & "\basscd.dll", 0)

    For i = 0 To 10
    
        ReDim Preserve CDi(i)
        
        If BASS_CD_GetInfo(i, CDi(i)) <> 0 Then

            If Chr$(65 + CDi(i).letter) = Mid(DriveLetter, 1, 1) Then
                       
                Exit For
                
            End If
            
        End If
    
    Next
    
    CDDrive = i
    
    sCDTitle = "Unknown"
    sCDArtist = "Unknown"
    
    RaiseEvent CDArtist(sCDArtist)
    RaiseEvent CDTitle(sCDTitle)
    
    ReDim sCDTrackName(CDTrackCount - 1)
    ReDim sCDTrackArtist(CDTrackCount - 1)
    
    For i = 0 To UBound(sCDTrackName)
        sCDTrackName(i) = "Track " & Format(i + 1, "00")
        sCDTrackArtist(i) = sCDArtist
    Next

    sCDFreeDBID = Mid(VBStrFromAnsiPtr(BASS_CD_GetID(CDDrive, BASS_CDID_CDDB)), 1, 8)
    
    If bCDDBLookup = True Then
    
        FreeDBInfo sCDFreeDBID
        
    End If
    
    CDLoad = True

End Function

Public Function CDLoadTrack(TrackNo As Long)

    If bFileLoaded = False Then

        FileLoad sCDDriveLetter, eCD, TrackNo

    Else

        Call BASS_CD_StreamSetTrack(StreamHandle(0), TrackNo)
        
    End If

End Function

Private Function FreeDBInfo(sID As String)

    Dim strURL As String
    Dim sRet As String
    Dim sLine() As String
    Dim lLine As Long
    Dim lLen As Long
    Dim sEnd As String
    Dim i As Long
    Dim strSplit() As String
    
    strURL = "http://www.freedb.org/freedb/misc/" & sID

    Dim objXML As Object
    Set objXML = CreateObject("Microsoft.XMLHTTP")
    objXML.Open "GET", strURL, False
    objXML.sEnd
    
    sRet = objXML.responseText
    
    If sRet <> "" Then
    
        'get cd title and artist
        
            lLine = InStr(1, sRet, "DTITLE")
            lLen = Len("DTITLE" & "=")
            sEnd = "DYEAR"
            strSplit = Split(Mid(sRet, lLine + lLen, InStr(lLine, sRet, sEnd) - lLine - Len(sEnd) - 2), " / ")
            
            sCDArtist = strSplit(0)
            sCDTitle = strSplit(1)
            
            RaiseEvent CDArtist(sCDArtist)
            RaiseEvent CDTitle(sCDTitle)
        
        'get track names
        
            sLine() = Split(sRet, vbCrLf, , vbTextCompare)
            On Error Resume Next
            For i = 0 To 99
    
                lLine = InStr(1, sRet, "TTITLE" & i)
                
                If lLine <> 0 Then
                
                    ReDim Preserve sCDTrackName(i)
                    ReDim Preserve sCDTrackArtist(i)
                    
                    lLen = Len("TTITLE" & i & "=")
                    sEnd = "TTITLE" & (i + 1)
                    strSplit() = Split(Mid(sRet, lLine + lLen, InStr(lLine, sRet, sEnd) - lLine - Len(sEnd) - 2), " / ")
                    
                    If UBound(strSplit) > 0 Then
                        sCDTrackName(i) = strSplit(1)
                        sCDTrackArtist(i) = strSplit(0)
                    Else
                        sCDTrackName(i) = strSplit(0)
                        sCDTrackArtist(i) = sCDArtist
                    End If
                    
                    RaiseEvent CDTrackDetails(i, CDTrackName(i), sCDTrackArtist(i))
                    
                Else
    
                    i = i - 1
                    lLine = InStr(1, sRet, "TTITLE" & i)
    
                    ReDim Preserve sCDTrackName(i)
                    ReDim Preserve sCDTrackArtist(i)
    
                    lLen = Len("TTITLE" & i & "=")
                    sEnd = "EXTD"
    
                    strSplit() = Split(Mid(sRet, lLine + lLen, InStr(lLine, sRet, sEnd) - lLine - Len(sEnd) - 6), " / ")
    
                    If UBound(strSplit) > 0 Then
                        sCDTrackName(i) = strSplit(1)
                        sCDTrackArtist(i) = strSplit(0)
                    Else
                        sCDTrackName(i) = strSplit(0)
                        sCDTrackArtist(i) = sCDArtist
                    End If
                
                    RaiseEvent CDTrackDetails(i, CDTrackName(i), sCDTrackArtist(i))
                
                    Exit For
                
                End If
                
            Next
            
    End If
        
End Function


Public Property Get CDTrackCount() As Long

    CDTrackCount = BASS_CD_GetTracks(CDDrive)

End Property

Public Property Get CDTrackSize(index As Long) As Long

    CDTrackSize = BASS_CD_GetTrackLength(CDDrive, index)

End Property

Public Property Get CDTrackLength(index As Long) As Long

    CDTrackLength = CDTrackSize(index) \ 176400

End Property

Public Property Get CDTotalSize() As Long

    Dim i As Long
    
    Dim lTemp As Long
    
    lTemp = 0
    
    For i = 0 To CDTrackCount - 1
    
        DoEvents
    
        lTemp = lTemp + CDTrackSize(i)
        
    Next
    
    CDTotalSize = lTemp + 2

End Property

Public Property Get CDTotalLength() As Long

    CDTotalLength = (CDTotalSize \ 176400) + 2

End Property

Public Property Get CDName() As String

    CDName = sCDTitle

End Property

Public Property Get CDArtist() As String

    CDArtist = sCDArtist

End Property

Public Property Get CDTrackName(index As Long) As String

    CDTrackName = sCDTrackName(index)

End Property

Public Property Get CDTrackArtist(index As Long) As String

    CDTrackArtist = sCDTrackArtist(index)

End Property

Public Property Get CDFreeDBID() As String

    CDFreeDBID = sCDFreeDBID
End Property

Public Function SetPosition(dPos As Double)

    Dim i As Integer
    
    FilePause

    Debug.Print "streams: " & UBound(StreamHandle)
    
    For i = 0 To UBound(StreamHandle)
    
        If StreamHandle(i) <> -1 Then
    
            BASS_ChannelSetPosition StreamHandle(i), BASS_ChannelSeconds2Bytes(StreamHandle(0), dPos), BASS_POS_BYTE
        
        End If
    
    Next
    
    FilePlay

End Function

Public Property Let PlaybackSpeakers(ByVal speakers As eSpeakers)

    eSpeakerMode = speakers

End Property

Public Property Get PlaybackSpeakers() As eSpeakers

    PlaybackSpeakers = eSpeakerMode

End Property
