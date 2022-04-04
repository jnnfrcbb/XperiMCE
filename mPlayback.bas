Attribute VB_Name = "mPlayback"
'Option Explicit

Public Type tPlaylistItem
    File As String
    Title As String
    Source As ePlaybackSources
    SubSource As Long
    Thumb As String
End Type
Public Playlist() As tPlaylistItem

Public Enum ePlaybackStates
    eWaiting = -1
    eStopped = 0
    ePlaying = 1
    ePaused = 2
    eBuffering = 3
    eFinished = 4
End Enum

Public Enum ePlaybackSources
    eNone = -1
    audio = 0
    video = 1
    dvd = 2
    Game = 5
    External = 6
    URL = 8
End Enum

Public Type ePlayback
    Source As ePlaybackSources
    SubSource As Long
    File As String
    Position As Double
    Duration As Double
    Shuffle As Boolean
    Repeat As Boolean
    Title As String
    Artist As String
    Album As String
    Track As Long
    Thumb As String
    BitRate As String
    BitrateVariable As Boolean
    SampleRate As String
    BitsPerSample As Long
    State As ePlaybackStates
    PlaylistIndex As Long
    PlaylistCount As Long
    Speakers As Long
    Scrobbled As Boolean
    NowPlayingShow As Boolean
    InfoString As String
    Channels As String
    Length As String
End Type
Public PlayBack As ePlayback

Public Type eMusicSettings
    Speakers As Long
    Volume As Double
    eqCount As Long
    eqHz() As Long
    eqValue() As Long
End Type
Public MusicSettings As eMusicSettings

Public lPlaylistHolderIndex As Long
Public lPlaylistHolderRow As Long

Public lPrevFx As Long

Public lOSD As Long
Public lOSDIndex As Long
Public bOSDAvailable As Boolean

Public bMinVideo As Boolean

Public Function LoadAudio(sFile As String, Optional sTitle As String = "", Optional bAutoPlay As Boolean = True) As Boolean
    
    Dim i As Long
'    Dim ID3Tag As New ID3ComTag
'    Dim MP3Inf As New MP3Info
    Dim fso As New FileSystemObject
    Dim bTemp As Boolean
    
    Dim Tags As New clsTags
    Dim SongTime As Long
        
    On Error GoTo errHandle
                
    With frmMain
    
        If .cPlayer.ubound > 0 Then
            For i = 1 To .cPlayer.ubound
                Unload .cPlayer(i)
            Next
        End If
        
        Load .cPlayer(1)
        
        .cPlayer(1).PlaybackSpeakers = MusicSettings.Speakers
        .cPlayer(1).PlayerVolume = MusicSettings.Volume
        
        .cPlayer(1).EQChannelCount MusicSettings.eqCount
        For i = 0 To 9
            .cPlayer(1).EQSetChannel i, MusicSettings.eqHz(i), MusicSettings.eqValue(i)
        Next
            
        Select Case Playlist(PlayBack.PlaylistIndex).SubSource
        
            Case 0 'local file
            
                bTemp = True
            
                .cPlayer(1).PlayerInitialise 0
            
                .cPlayer(1).FileLoad sFile, 0
                
                If bAutoPlay = True Then
                    
                    .cPlayer(1).FilePlay
                    
                End If
                    
                If Not Tags.Loaded(ttAutomatic) Then
                    Tags.Create
                End If
                        
                Tags.Load sFile, ttAutomatic
    
                PlayBack.Artist = Tags.GetTag("ARTIST", ttAutomatic)
                If PlayBack.Artist = "" Then
                    PlayBack.Artist = Tags.GetTag("ALBUMARTIST", ttAutomatic)
                End If
                PlayBack.Album = Tags.GetTag("ALBUM", ttAutomatic)
                PlayBack.Track = Tags.GetTag("TRACKNUMBER", ttAutomatic)
                PlayBack.Title = Tags.GetTag("TITLE", ttAutomatic)
                If PlayBack.Title = "" Then
                    PlayBack.Title = Playlist(PlayBack.PlaylistIndex).Title
                End If
                
                Select Case Mid(sFile, InStrRev(sFile, ".") + 1)
    
                    Case "mp3", "mp1", "mp2", "mpa", "mpc"
                    
                        Tags.GetMPEGAudioAttributes atMPEG
                        
                        PlayBack.Channels = Tags.GetAudioAttribute(aaChannels)
                        PlayBack.BitRate = Tags.MPGAttributesBitrate
                        PlayBack.BitrateVariable = IIf(Tags.MPGAttributesVBR, "True", "False")
                        PlayBack.SampleRate = Tags.MPGAttributesSampleRate
                        PlayBack.BitsPerSample = Tags.GetAudioAttribute(aaBitsPerSample)
                        SongTime = CLng(Tags.GetAudioAttribute(aaPlayTime))
                        PlayBack.Length = Tags.GetTimeFormat(SongTime)
                    
                    Case "flac", "ogg"
                    
                        Tags.GetFlacAudioAttributes atFlac
                        
                        PlayBack.Channels = Tags.FlacAttributesChannels
                        PlayBack.BitRate = Tags.FlacAttributesBitrate
                        PlayBack.BitrateVariable = False
                        PlayBack.SampleRate = Tags.FlacAttributesSampleRate
                        PlayBack.BitsPerSample = Tags.GetAudioAttribute(aaBitsPerSample)
                        SongTime = CLng(Tags.FlacAttributesPlayTime)
                        PlayBack.Length = Tags.GetTimeFormat(SongTime)
                        
                    Case "m4a"
                    
                        Tags.GetMP4AudioAttributes atMP4
                        
                        PlayBack.Channels = Tags.MP4AttributesChannels
                        PlayBack.BitRate = Tags.MP4AttributesBitrate
                        PlayBack.BitrateVariable = False
                        PlayBack.SampleRate = Tags.MP4AttributesSampleRate
                        PlayBack.BitsPerSample = Tags.GetAudioAttribute(aaBitsPerSample)
                        SongTime = CLng(Tags.FlacAttributesPlayTime)
                        PlayBack.Length = Tags.GetTimeFormat(SongTime)
                    
                End Select
    
                PlayBack.Thumb = Playlist(PlayBack.PlaylistIndex).Thumb
                PlayBack.Scrobbled = False
                PlayBack.Channels = MP3Inf.Channels

                If bLastFMLoggedIn = True Then
                    .cLastFM(1).UpdateNowPlaying PlayBack.Artist, PlayBack.Title, PlayBack.Album
                End If
                
            Case 1 'streaming file
            
                .cPlayer(1).PlayerInitialise 1
            
                .cPlayer(1).PlaybackSpeakers = MusicSettings.Speakers
            
                .cPlayer(1).FileLoad sFile, 1
                
            Case 2 'cd
            
                .cPlayer(1).PlayerInitialise 2
            
                .cPlayer(1).PlaybackSpeakers = MusicSettings.Speakers
            
                .cPlayer(1).CDLoad sFile, True
                
            Case 3 'netradio
            
                bTemp = False
            
                .cPlayer(1).PlayerInitialise 1
            
                .cPlayer(1).PlaybackSpeakers = MusicSettings.Speakers
            
                .cPlayer(1).FileLoad sFile, 1
            
                PlayBack.Title = sTitle
                PlayBack.Artist = "Streaming radio"
                PlayBack.Album = ""
                PlayBack.Thumb = Playlist(PlayBack.PlaylistIndex).Thumb
            
        End Select
                  
        If Tiles(.cRow(lMusicRow).TileTag(0)).Source <> "app##cMusicHome##10" Then
                    
            Do Until AddTile(lMusicRow, 0, 0, PlayBack.Title, "Now Playing", PlayBack.Thumb, "app##cMusicHome##10", 0, False) = True
                DoEvents
            Loop
            
        Else
        
            Do Until UpdateTile(lMusicRow, 0, "app##cMusicHome##10", PlayBack.Title, "Now Playing", PlayBack.Thumb, False) = True
                DoEvents
            Loop
            
        End If
        
        
        Do Until RemoteSend("NOWPLAYING##" & PlayBack.Title & "##" & PlayBack.Artist & "##" & PlayBack.Album & "##" & PlayBack.Thumb) = True
            DoEvents
        Loop
        
        If lRowSelected = lMusicRow Then
            If .cRow(lMusicRow).TileCurrent = 0 Then
                .cRow(lMusicRow).TileSelected(.cRow(lMusicRow).TileCurrent) = True
            End If
        End If
                     
        .tmrPlayback.Enabled = True
    
        'If bDashLocked = True Then
        
            If Len(PlayBack.Title) > 40 Then
                .lblLockArt(0).Caption = RTrim(Mid(PlayBack.Title, 1, 40)) & "..."
            Else
                .lblLockArt(0).Caption = PlayBack.Title
            End If
            
            .pLockArt.Picture = LoadPictureGDIplus(PlayBack.Thumb)
            
            .pLockArt.Visible = True
            .lblLockArt(0).Visible = True
            .lblLockArt(1).Visible = True
            
            SetLockBack .pLockArt.Picture, True, 5, 50
        
        'End If
                    
        PlayBack.SubSource = Playlist(PlayBack.PlaylistIndex).SubSource
        PlayBack.Speakers = .cPlayer(1).SystemSpeakers
        
        PlayBack.InfoString = ProcessAudioInfo(PlayBack.Speakers, CLng(PlayBack.Channels), CLng(PlayBack.BitRate), PlayBack.BitrateVariable, CLng(PlayBack.SampleRate))
        
        .cMusicHome.SendMessage "nowplayingstart##" & PlayBack.Title & "##" & PlayBack.Artist & "##" & PlayBack.Album & "##" & PlayBack.Thumb & "##" & PlayBack.PlaylistIndex & "##" & PlayBack.PlaylistCount & "##" & PlayBack.Shuffle & "##" & PlayBack.Repeat & "##" & PlayBack.BitRate & "##" & PlayBack.Speakers & "##" & PlayBack.SubSource & "##" & PlayBack.InfoString, Nothing
        
        .cMusicHome.SendMessage "nowplayingplaying", Nothing
        
        PlaybackNotify
        
    End With
    
    
    LoadAudio = True
    
errHandle:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
    
End Function

Public Function CloseAudio(Optional bPlaylist As Boolean = False) As Boolean

    On Error Resume Next
    
    Dim i As Long
    
    With frmMain
    
        .tmrPlayback.Enabled = False
    
        If .cPlayer.ubound > 0 Then
            
            .cPlayer(1).PlayerShutdown
            
            Unload .cPlayer(1)
            
        End If
        
        If bPlaylist = False Then
            
            Do Until DeleteTile(lMusicRow, 0, 0, False, False) = True
                DoEvents
            Loop
                        
            .cMusicHome.SendMessage "nowplayingend", Nothing
                
            .pLockArt.Visible = False
            Set .pLockArt.Picture = Nothing
            .lblLockArt(0).Visible = False
            .lblLockArt(1).Visible = False
                
            SetBackground LoadPictureGDIplus(.pLockHolder.Tag), .pLockHolder
            
        End If
        
        If lRowSelected = lMusicRow Then
            If .cRow(lMusicRow).TileCurrent = 0 Then
                .cRow(lMusicRow).TileSelected(.cRow(lMusicRow).TileCurrent) = True
            End If
        End If
                 
    End With
    
    CloseAudio = True

End Function

Public Function LoadWMPVideo(sFile As String, Optional bAutoPlay As Boolean = True) As Boolean

    On Error Resume Next
    
    With frmMain
    
        PlayBack.Source = 1
    
        If bMinVideo = True Then
        
            .pVideoHolder.Move .pMinVideo.Left, .pMinVideo.Top, .pMinVideo.Width, .pMinVideo.Height
                
            .pVideoHolder.BorderStyle = 1
                
            bOSDAvailable = False
                
        Else
        
            .pVideoHolder.Move 0, 0, .ScaleWidth, .ScaleHeight
                
            .pVideoHolder.BorderStyle = 0
            
            bOSDAvailable = True
                        
        End If
            
        .wmp.Move 0, 0, .pVideoHolder.Width, .pVideoHolder.Height
        
        .wmp.Visible = True
    
        .pVideoHolder.ZOrder 0
        .pVideoHolder.Visible = True
        
        .wmp.enableContextMenu = False
    
        .wmp.URL = sFile
        
        .tmrPlayback.Enabled = True
        
        .wmp.stretchToFit = True
        
        .wmp.Visible = True
        
        If bAutoPlay = True Then
        
            .wmp.Controls.play
            
        End If
        
        lPrevFx = 0
        
        .tmrOSD.Enabled = True
        
    End With

    LoadWMPVideo = True

End Function

Public Function PlayPause() As Boolean

    On Error Resume Next
    
    Sound "select"

    With frmMain
    
        Select Case PlayBack.State
                                
            Case 1 'playing
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(.cPlayer.ubound).FilePause
                        
                        .cMusicHome.SendMessage "nowplayingpause", Nothing
                    
                        PlaybackNotify
                    
                    Case 1 'video
                    
                        ShowOSD True

                        .wmp.Controls.pause
                    
                    Case 2 'DVD
                                               
                End Select
                
                Do Until RemoteSend("CONTROL##PAUSED") = True
                    DoEvents
                Loop
            
            Case 2 'paused
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(.cPlayer.ubound).FilePlay
                    
                        .cMusicHome.SendMessage "nowplayingplaying", Nothing
                    
                        PlaybackNotify
                    
                    Case 1 'video
                    
                        .wmp.Controls.play
                    
                    Case 2 'DVD
                                                                
                End Select
            
                Do Until RemoteSend("CONTROL##PLAYING") = True
                    DoEvents
                Loop
            
        End Select

    End With

    PlayPause = True
    
End Function

Public Function ResetPlayback() As Boolean

    On Error Resume Next
    
    With frmMain

    PlaylistLoad 0, False

'        Select Case PlayBack.Source
            
'            Case 0 'audio
            
'                .cPlayer(1).SetPosition 0
            
'                .cPlayer(1).FilePause
            
'                .cMusicHome.SendMessage "nowplayingpause", Nothing
                    
'            Case 1 'video
            
            
'                StopPlayback
            
'            Case 2 'DVD
            
'        End Select
    
    End With

    ResetPlayback = True

End Function

Public Function StopPlayback(Optional bPlaylist As Boolean = False) As Boolean

    On Error Resume Next
    
    PlayBack.File = ""
    
    Do Until RemoteSend("CONTROL##STOP") = True
        DoEvents
    Loop

    With frmMain
    
        .lblStatus.Caption = sWeatherStatus

        Select Case PlayBack.Source
            
            Case 0 'audio
            
                CloseAudio bPlaylist
            
            Case 1 'video
            
                CloseVideo
    
            Case 2 'DVD
                                        
        End Select

    End With

    PlayBack.Album = ""
    PlayBack.Artist = ""
    PlayBack.BitRate = ""
    PlayBack.BitrateVariable = False
    PlayBack.Duration = 0
    PlayBack.File = ""
    PlayBack.Position = 0
    PlayBack.SampleRate = ""
    PlayBack.Source = -1
    PlayBack.Speakers = 0
    PlayBack.State = eWaiting
    PlayBack.SubSource = 0
    PlayBack.Thumb = ""
    PlayBack.Title = ""
    PlayBack.Track = 0

    StopPlayback = True
    
End Function


Public Function SkipForwardSmall() As Boolean

    On Error Resume Next
    
    Sound "select"

    With frmMain
    
        Select Case PlayBack.State
        
            Case 1, 2 'playing,paused
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(1).SetPosition .cPlayer(1).FilePosition + 10
                        
                        If PlayBack.State = 2 Then
                        
                            .cPlayer(1).FilePause
                            
                        End If
                        
                        PlaybackNotify
                    
                    Case 1 'video
                    
                        ShowOSD True

                        .wmp.Controls.currentPosition = .wmp.Controls.currentPosition + 15
                    
                        If PlayBack.State = 2 Then
                        
                            .wmp.Controls.pause
                            
                        End If
                    
                    Case 2 'DVD
                                                                    
                End Select
            
        End Select

    End With

    SkipForwardSmall = True
    
End Function

Public Function SkipForwardLarge() As Boolean

    On Error Resume Next
    
    Sound "select"

    With frmMain
    
        Select Case PlayBack.State
        
            Case 1, 2 'playing,paused
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(1).SetPosition .cPlayer(1).FilePosition + 30
                    
                        If PlayBack.State = 2 Then
                        
                            .cPlayer(1).FilePause
                            
                        End If
                        
                        PlaybackNotify
                    
                    Case 1 'video
                    
                        ShowOSD True

                        .wmp.Controls.currentPosition = .wmp.Controls.currentPosition + 60
                    
                        If PlayBack.State = 2 Then
                        
                            .wmp.Controls.pause
                            
                        End If
                    
                    Case 2 'DVD
                                                                    
                End Select
            
        End Select

    End With

    SkipForwardLarge = True
    
End Function

Public Function SkipBackSmall() As Boolean

    On Error Resume Next
    
    Sound "select"

    With frmMain
    
        Select Case PlayBack.State
        
            Case 1, 2 'playing,paused
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(1).SetPosition .cPlayer(1).FilePosition - 10
                    
                        If PlayBack.State = 2 Then
                        
                            .cPlayer(1).FilePause
                            
                        End If
                        
                        PlaybackNotify
                    
                    Case 1 'video
                    
                        ShowOSD True

                        .wmp.Controls.currentPosition = .wmp.Controls.currentPosition - 15
                    
                        If PlayBack.State = 2 Then
                        
                            .wmp.Controls.pause
                            
                        End If
                    
                    Case 2 'DVD
                                                                    
                End Select
            
        End Select

    End With

    SkipBackSmall = True
    
End Function

Public Function SkipBackLarge() As Boolean

    On Error Resume Next
    
    Sound "select"

    With frmMain
    
        Select Case PlayBack.State
        
            Case 1, 2 'playing,paused
            
                Select Case PlayBack.Source
                    
                    Case 0 'audio
                    
                        .cPlayer(1).SetPosition .cPlayer(1).FilePosition - 30
                    
                        If PlayBack.State = 2 Then
                        
                            .cPlayer(1).FilePause
                            
                        End If
                        
                        PlaybackNotify
                        
                    Case 1 'video
                    
                        ShowOSD True

                        .wmp.Controls.currentPosition = .wmp.Controls.currentPosition - 60
                    
                        If PlayBack.State = 2 Then
                        
                            .wmp.Controls.pause
                            
                        End If
                    
                    Case 2 'DVD
                      
                End Select
            
        End Select
            
    End With

    SkipBackLarge = True
    
End Function


Public Function PlaybackRepeatToggle() As Boolean

    On Error Resume Next
    
    PlayBack.Repeat = Not PlayBack.Repeat
    
    If PlayBack.Repeat = True Then
        frmMain.pVidControl(4).BlendPct = 100
    Else
        frmMain.pVidControl(4).BlendPct = 0
    End If
    
End Function

Public Function PlaybackShuffleToggle() As Boolean

    On Error Resume Next
    
    PlayBack.Shuffle = Not PlayBack.Shuffle
    
    If PlayBack.Shuffle = True Then
        frmMain.pVidControl(3).BlendPct = 100
    Else
        frmMain.pVidControl(3).BlendPct = 0
    End If
    
End Function

Public Function PlaylistAdd(sFile As String, sTitle As String, eSource As ePlaybackSources, Optional lSubSource As Long, Optional sThumb As String) As Boolean

    On Error Resume Next
    
    Dim sSub As String

    With frmMain
    
        Dim lUbound As Long
    
        If Playlist(0).File <> "" Then
            ReDim Preserve Playlist(UBound(Playlist) + 1)
        End If
        
        lUbound = UBound(Playlist())
        
        Playlist(lUbound).File = sFile
        Playlist(lUbound).Title = sTitle
        Playlist(lUbound).Source = eSource
        Playlist(lUbound).SubSource = lSubSource
        Playlist(lUbound).Thumb = sThumb
    
        sSub = ProcessSubSource(Playlist(lUbound).Source, Playlist(lUbound).SubSource)
    
        .cPlaylist.TileSet lUbound, CStr(lUbound + 1) & " | " & Playlist(lUbound).Title, sSub, sThumb
        
        PlayBack.PlaylistCount = lUbound
        
        .cMusicHome.SendMessage "playlistupdated##" & PlayBack.PlaylistIndex & "##" & PlayBack.PlaylistCount, Nothing
        
    End With

    PlaylistAdd = True
    
End Function

Public Function PlaylistRemove(Index As Long) As Boolean

    On Error Resume Next
    
    PlaylistRemove = True
    
End Function

Public Function PlaylistClear() As Boolean

    On Error Resume Next
    
    Static i As Integer

    ReDim Playlist(0)

    frmMain.cPlaylist.ClearTiles
    
    If PlayBack.File <> "" Then
        StopPlayback True
    End If

    PlaylistClear = True
    
End Function

Public Function PlaylistBack() As Boolean

    On Error Resume Next
    
    If PlayBack.Shuffle = True Then
    
        PlaylistLoad RandomNumber(0, PlayBack.PlaylistCount)
    
    Else
            
        If PlayBack.Position > 10 Then
        
            PlaylistLoad PlayBack.PlaylistIndex
            
        Else
            
            If PlayBack.PlaylistIndex > 0 Then
                PlaylistLoad PlayBack.PlaylistIndex - 1
            Else
                If PlayBack.Repeat = True Then
                    PlaylistLoad UBound(Playlist)
                Else
                    StopPlayback
                End If
            End If
        
        End If
        
    End If
    
    PlaylistBack = True
    
End Function

Public Function PlaylistNext() As Boolean
    
    On Error Resume Next
    
    If PlayBack.Shuffle = True Then
    
        PlaylistLoad RandomNumber(0, PlayBack.PlaylistCount)
    
    Else
        
        If PlayBack.PlaylistIndex < UBound(Playlist) Then
            PlaylistLoad PlayBack.PlaylistIndex + 1
        Else
            If PlayBack.Repeat = True Then
                PlaylistLoad 0
            Else
                'StopPlayback
                ResetPlayback
            End If
        End If
    
    End If
    
    PlaylistNext = True
    
End Function

Public Function PlaylistLoad(Index As Long, Optional bAutoPlay As Boolean = True) As Boolean

    On Error Resume Next
    
    Dim i As Long

    If PlayBack.File <> "" Then
        StopPlayback True
    End If

    PlayBack.PlaylistIndex = Index
    
    PlayBack.Source = Playlist(Index).Source
    PlayBack.SubSource = Playlist(Index).SubSource
    PlayBack.File = Playlist(Index).File
    PlayBack.Title = Playlist(Index).Title
            
    Select Case Playlist(Index).Source
    
        Case 0 'audio
        
            If PlayBack.SubSource <> 3 Then
                OpenFile Playlist(Index).File, , False, bAutoPlay
            Else
                LoadAudio PlayBack.File, Playlist(Index).Title, bAutoPlay
            End If

        Case 1 'video
        
            OpenFile Playlist(Index).File, , False, bAutoPlay

        Case 2 'dvd
        
            OpenFile Playlist(Index).File, , False, bAutoPlay
            
        Case 3 'picture
                
    End Select

    Do Until RemoteSend("CONTROL##PLAYING") = True
        DoEvents
    Loop
      
    If Len(PlayBack.Title) <= 41 Then
        frmMain.lblVidDetails(1).Caption = PlayBack.Title
    Else
        frmMain.lblVidDetails(1).Caption = Trim(Mid(PlayBack.Title, 1, 38)) & "..."
    End If
    
    Do Until PlaylistIconSet = True
        DoEvents
    Loop
    
    PlaylistLoad = True

End Function

Public Function ShowOSD(bShow As Boolean)

    On Error Resume Next
    
    If bShow = True Then
    
        If frmMain.pPlaylistHolder.Visible = False Then
            If bMinVideo = True Then
                frmMain.pMiniOSD.ZOrder 0
            Else
                frmMain.pVidMenu.ZOrder 0
            End If
        End If
        
        If bMinVideo = True Then
            frmMain.pMiniOSD.Visible = True
        Else
            frmMain.pVidMenu.Visible = True
        End If
        
        frmMain.tmrShowOSD.Enabled = True
    
        lOSD = 0
    
        frmMain.tmrOSD.Enabled = True
    
    Else
    
        frmMain.pVidMenu.Visible = False
        
        frmMain.pMiniOSD.Visible = False
        
        frmMain.pVidMenu.Top = frmMain.ScaleHeight
    
    End If
    
End Function

Public Function VideoPlayingKey(KeyCode As Integer, Shift As Integer) As Boolean
    
    On Error Resume Next
    
    Select Case PlayBack.Source
    
        Case 1, 3
        
        Select Case KeyCode
        
            Case vbKeyShift, vbKeyMenu, 17, 68, 173, 174, 175, 93
            
                'nothing
            
            Case vbKeyBack
            
                ShowOSD False
            
            Case Else
            
                ShowOSD True
            
        End Select
            
        Select Case KeyCode
        
            Case vbKeyLeft
            
                Sound "select"
                    
                If lOSDIndex > 0 Then
                
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 75
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = False
                    
                    lOSDIndex = lOSDIndex - 1
                    
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 0
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = True
                
                Else
                
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 75
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = False
                    
                    lOSDIndex = frmMain.pVidControl.ubound
                    
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 0
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = True
                    
                End If
            
            Case vbKeyRight
            
                Sound "select"
                    
                If lOSDIndex < frmMain.pVidControl.ubound Then
                
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 75
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = False
                    
                    lOSDIndex = lOSDIndex + 1
                    
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 0
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = True
                    
                Else
                
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 75
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = False
                    
                    lOSDIndex = 0
                    
                    frmMain.pVidControl(lOSDIndex).TransparencyPct = 0
                    frmMain.pVidControl(lOSDIndex).BackStyleOpaque = True
                    
                End If
            
            Case vbKeyUp
            
            Case vbKeyDown
            
            Case vbKeyReturn
            
                Do Until pVidControlClick(CInt(lOSDIndex)) = True
                    DoEvents
                Loop
            
            Case vbKeyBack
            
                StopPlayback
            
            Case vbKeyMenu, 93
            
                Do Until ShowPlaylist = True
                    DoEvents
                Loop
            
            Case vbKeySpace
            
                PlayPause
        
            Case 66 'rewind
            
                If Shift = 3 Then
                    SkipBackSmall
                End If
                
            Case 70 'fastforward
            
                If Shift = 3 Then
                    SkipForwardSmall
                End If
    
        End Select
    
    End Select

    VideoPlayingKey = True

End Function

Public Function CloseVideo() As Boolean
    
    On Error Resume Next
    
    Sound "closeelement"
    
    frmMain.tmrPlayback.Enabled = False
    frmMain.lblStatus.Caption = ""
    frmMain.pVidPosition(1).Width = 0
    frmMain.lblVidDetails(0).Caption = ""
    frmMain.lblVidDetails(1).Caption = ""
    ShowOSD False
    bOSDAvailable = False
    frmMain.pVideoHolder.Visible = False
        
    Select Case PlayBack.Source
    
        Case 1 'video
            
            frmMain.wmp.Close
            
        Case 2 'dvd
        
    End Select

    CloseVideo = True

End Function

Public Function pVidControlClick(Index As Integer) As Boolean

    On Error Resume Next
    
    lOSD = 0

    Sound "yes"

    Select Case Index
    
        Case 0 'repeat
        
            PlaybackRepeatToggle
            
        Case 1 'shuffle
                
            PlaybackShuffleToggle
                
        Case 2 'skip back large
        
            SkipBackLarge
        
        Case 3 'playpause

            PlayPause
            
        Case 4 'skip forward large
        
            SkipForwardLarge
        
        
        Case 5 'playlist
        
            Do Until ShowPlaylist = True
                DoEvents
            Loop
            
        Case 6 'minimize
        
            Do Until MinRestoreVideo = True
                DoEvents
            Loop
            
    End Select
    
    pVidControlClick = True
    
End Function

Public Function ShowPlaylist() As Boolean

    On Error Resume Next
    
    Dim i As Long
    
    Sound "openelement"
            
    With frmMain
    
        Select Case CurrentScreen
            
            Case 0
                .cRow(lRowSelected).Deselect
            Case 2
                '.cKeyboard
            Case 3
                .cSearch.Deselect
        End Select
    
        Do Until CaptureScreen(frmMain.pPlaylistHolder.hDC) = True
            DoEvents
        Loop
        
        .pPlaylistHolder.Left = 0
        .pPlaylistHolder.ZOrder 0
            
        Do Until PlaylistIconSet = True
            DoEvents
        Loop
    
        .cPlaylist.TileSelected(PlayBack.PlaylistIndex) = True
        
        .pPlaylistHolder.Visible = True
        
    End With

    ShowPlaylist = True

End Function

Public Function ClosePlaylist()

    On Error Resume Next
    
    With frmMain
    
        Sound "closeelement"
    
        .pPlaylistHolder.Visible = False
              
        Select Case CurrentScreen
            
            Case 0
                .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
            Case 2
                '.cKeyboard
            Case 3
                .cSearch.Deselect
        End Select
        
    End With

End Function

Public Function PlaylistKey(KeyCode As Integer) As Boolean

    On Error Resume Next
    
    With frmMain
    
        Select Case KeyCode
        
            Case vbKeyLeft
            
                .cPlaylist.TileSelectLeft
            
            Case vbKeyRight
            
                .cPlaylist.TileSelectRight
        
            Case vbKeyBack, vbKeyMenu, 93
            
                ClosePlaylist
            
            Case vbKeyReturn
            
                Sound "yes"
            
                PlayBack.PlaylistIndex = .cPlaylist.TileCurrent
                
                Do Until PlaylistLoad(PlayBack.PlaylistIndex) = True
                    DoEvents
                Loop
            
                Select Case PlayBack.Source
                    Case 0
                    Case Else
                        ClosePlaylist
                End Select
                
            Case vbKeyShift
    
        End Select

    End With

    PlaylistKey = True

End Function

Public Function ConvertPlayerStatus(lStatus As Long) As Long

    On Error Resume Next
    
    Select Case lStatus
        Case 0 'eError
            ConvertPlayerStatus = 0
        Case 1 'eWaiting
            ConvertPlayerStatus = -1
        Case 2 'eInitialised
            ConvertPlayerStatus = -1
        Case 3 'eLoaded
            ConvertPlayerStatus = -1
        Case 4 'ePlaying
            ConvertPlayerStatus = 1
        Case 5 'ePaused
            ConvertPlayerStatus = 2
        Case 6 'eStopped
            ConvertPlayerStatus = 0
        Case 7 'eShutdown
            ConvertPlayerStatus = 0
    End Select

End Function

Private Function ArtFilename(sArtist As String, sAlbum As String) As String

    On Error Resume Next

    Dim sFilename As String
    
    sFilename = sAlbum & " - " & sArtist
    
    sFilename = ReplaceInvalidCharacters(sFilename)
    
    ArtFilename = App.Path & "\AlbumArt\" & sFilename & ".jpg"

End Function

Private Function ReplaceInvalidCharacters(ByVal sInput As String) As String
  
    On Error Resume Next

    sInput = Replace(sInput, Chr(92), "") '/
    
    sInput = Replace(sInput, Chr(47), "") '\
    
    sInput = Replace(sInput, Chr(58), "") ':
    
    sInput = Replace(sInput, Chr(42), "") '*
    
    sInput = Replace(sInput, Chr(63), "") '?
    
    sInput = Replace(sInput, Chr(34), "") '"
    
    sInput = Replace(sInput, Chr(60), "") '<
    
    sInput = Replace(sInput, Chr(62), "") '>
    
    sInput = Replace(sInput, Chr(124), "") '|
    
    sInput = Replace(sInput, Chr(46), "") '.
    
    sInput = Replace(sInput, Chr(32), Chr(43)) 'space
    
    ReplaceInvalidCharacters = sInput

End Function

Public Function PlaylistIconSet() As Boolean

    With frmMain
    
        Dim i As Long
        
        
        For i = 0 To .cPlaylist.TileCount
            DoEvents
            .cPlaylist.TileIconClear CLng(i)
        Next
        .cPlaylist.TileIconSet PlayBack.PlaylistIndex, App.Path & "\Images\osd\playlist_current.png"
        
        
    End With
    
    PlaylistIconSet = True

End Function

Public Function InitPlayback() As Boolean

    ReDim Playlist(0)
    
    sYTQuality = "hd1080"
    
    bMinVideo = False
    
    PlayBack.NowPlayingShow = True
    PlayBack.Shuffle = False
    PlayBack.Repeat = False
    PlayBack.Source = -1
    
    Set Volume = New cVolume
        
    InitPlayback = True

End Function


Public Function PlaybackNotify() As Boolean

    With frmMain
    
        If bDashLoaded = True And bDashLocked = False And PlayBack.NowPlayingShow = True Then
            
            bStatusVisible = True
            
            lNotify = 0
                        
            '.pCharmsHolder.Top = 0 - .pCharmsHolder.Height + .sCharmSection.Height
            '.pCharmsHolder.ZOrder 0
            '.pCharmsHolder.Visible = True
            
            .pNotifyHolder.ZOrder 0
            .pNotifyHolder.Visible = True
            .pNotifyIcon.Picture = LoadPictureGDIplus(PlayBack.Thumb) '.pNotifyDefault(2).Picture
                    
            .tmrNotify.Enabled = True
        
        End If

    End With

End Function

Public Function ProcessAudioInfo(Speakers As Long, Optional Channels As Long = 0, Optional BitRate As Long = 0, Optional Variable As Boolean = False, Optional SampleRate As Long = 0) As String

    Dim sSpeakers As String
    Dim sChannels As String
    Dim sBitRate As String
    Dim sSampleRate As String
    
    Select Case Speakers
        Case 2
            sSpeakers = "Speakers: 2"
        Case 3
            sSpeakers = "Speakers: 2.1"
        Case 4
            sSpeakers = "Speakers: 3.1"
        Case 5
            sSpeakers = "Speakers: 4.1"
        Case 6
            sSpeakers = "Speakers: 5.1"
        Case 8
            sSpeakers = "Speakers: 7.1"
    End Select
    
    Select Case Channels
        Case 0
            sChannels = ""
        Case Else
            sChannels = " | Channels: " & Channels
    End Select
    
    Select Case BitRate
        Case 0
            sBitRate = ""
        Case Else
            If Variable = True Then
                sBitRate = " | Bitrate: VBR"
            Else
                sBitRate = " | Bitrate: " & Mid(CStr(BitRate), 1, 3) & " Kbps"
            End If
    End Select
    
    Select Case SampleRate
        Case 0
            sSampleRate = ""
        Case Else
            sSampleRate = " | Sample rate: " & Mid(CStr(SampleRate / 1000), 1, 4) & " kHz"
    End Select
    
    ProcessAudioInfo = sSpeakers & sChannels & sBitRate & sSampleRate & " | V: " & frmMain.cPlayer(1).PlayerVolume & "%"

End Function


Public Function TruncatePlaylist() As Boolean

    ReDim Preserve Playlist(PlayBack.PlaylistIndex)
    
    PlayBack.PlaylistCount = UBound(Playlist)

    TruncatePlaylist = True
    
End Function

