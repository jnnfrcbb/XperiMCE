Attribute VB_Name = "mMessaging"
Public Function CentralMessage(sMessage As String, Optional pImage As GDIpImage = Nothing)

    On Error Resume Next

    With frmMain
    
        For Each Control In .Controls
        
            Control.SendMessage sMessage, pImage
            
        Next
    
    End With

End Function

Public Function ProcessMessage(cSource As Control, sMessage As String, pImage As GDIpImage) As Boolean

    On Error Resume Next
    
    With frmMain
    
        Dim strSplit() As String
        Dim sSource As String
        Static i As Integer
        Dim lTemp As Long
        
        strSplit = Split(sMessage, "##")
        
        Select Case LCase(strSplit(0))
        
            Case "sound"
            
                Sound strSplit(1)
        
            Case "goup"
            
                OpenCharms
            
            Case "leavecontrol"
            
                CloseApp cSource
            
            Case "loadfile"
        
                Do Until PlaylistClear = True
                    DoEvents
                Loop
                
                Select Case cSource.Name
                
                    Case "cMusicHome", "cMusicArtists", "cMusicAlbums"
                    
                        'loadfile##file##title##artist##thumb
                    
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 0, 0, strSplit(4)) = True
                            DoEvents
                        Loop
                
                        Do Until PlaylistLoad(0) = True
                            DoEvents
                        Loop
                        
                    Case "cFilms"
                        
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 1, 0, strSplit(3)) = True
                            DoEvents
                        Loop
                        
                        Do Until PlaylistLoad(0) = True
                            DoEvents
                        Loop
                        
                        sSource = "file##" & strSplit(1) & "##" & strSplit(2) & "##1##0"
                        
                        If lVideoRow > -1 Then
                            
                            Do Until UpdateTile(lVideoRow, 0, sSource, vbNullString, "Last Watched | " & strSplit(2), strSplit(3), True, 0) = True
                                DoEvents
                            Loop
                        
                        End If
                
                    Case "cTV"
                           
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 1, 1, strSplit(3)) = True
                            DoEvents
                        Loop
                        
                        Do Until PlaylistLoad(0) = True
                            DoEvents
                        Loop
                        
                        sSource = "file##" & strSplit(1) & "##" & strSplit(2) & "##1##1"
                        
                        If lVideoRow > -1 Then
                
                            Do Until UpdateTile(lVideoRow, 0, sSource, vbNullString, "Last Watched | " & strSplit(2), strSplit(3), True, 1) = True
                                DoEvents
                            Loop
                        
                        End If
    
                    Case "cGames"
                            
                        'loadfile##file##name##thumb##isemulator##hasquotes##subcaption
                        
                        Do Until LoadGame(strSplit(1), strSplit(2), strSplit(3), strSplit(4), strSplit(5), strSplit(6), strSplit(7), strSplit(8)) = True
                            DoEvents
                        Loop
                        
                    Case "cDevices"
                                
                        'loadfile##type##path##name##thumb
                        
                        PlaylistClear
                        
                        PlaylistAdd strSplit(3), strSplit(4), CLng(strSplit(1)), CLng(strSplit(2)), strSplit(5)
                
                        PlaylistLoad (0)
                    
                End Select
        
            Case "queuefile"
            
                Select Case cSource.Name
                
                    Case "cMusicHome", "cMusicArtists", "cMusicAlbums"
                    
                        'queuefile##file##title##artist##thumb
                    
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 0, 0, strSplit(4)) = True
                            DoEvents
                        Loop
                
                    Case "cFilms"
                    
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 1, 0, strSplit(3)) = True
                            DoEvents
                        Loop
                        
                    Case "cTV"
                                
                        Do Until PlaylistAdd(strSplit(1), strSplit(2), 1, 1, strSplit(3)) = True
                            DoEvents
                        Loop
        
                End Select
                    
            Case "loadradio"
            
                Do Until PlaylistClear = True
                    DoEvents
                Loop
                
                Do Until PlaylistAdd(strSplit(1), strSplit(2), 0, 3, strSplit(3)) = True
                    DoEvents
                Loop
                
                Do Until PlaylistLoad(0) = True
                    DoEvents
                Loop
                
            Case "pin"
                
                Select Case cSource.Name
                
                    Case "cFilms"
                            
                        'pin##film##file##title#thumb
                              
                        sSource = "file##" & strSplit(2) & "##" & strSplit(3) & "##1##0"
                    
                        If bPinReplace = True Then
                    
                            Do Until UpdateTile(lRowSelected, .cRow(lRowSelected).TileCurrent, sSource, vbNullString, "Pinned | " & strSplit(3), strSplit(4), True, 0) = True
                                DoEvents
                            Loop
                        
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
                        
                        Else
                        
                            If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                                lTemp = .cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent + 1)
                            Else
                                lTemp = -1
                            End If
                            
                            Do Until AddTile(lRowSelected, .cRow(lRowSelected).TileCurrent + 1, lTemp, vbNullString, "Pinned | " & strSplit(3), strSplit(4), sSource, 0) = True
                                DoEvents
                            Loop
                            
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent + 1) = True
                        
                        End If
                
                    Case "cTV"
                    
                        'pin##tvshow##title##folder##thumb
                              
                        sSource = strSplit(1) & "##" & strSplit(3)
                            
                        If bPinReplace = True Then
                        
                            Do Until UpdateTile(lRowSelected, .cRow(lRowSelected).TileCurrent, sSource, vbNullString, "Pinned | " & strSplit(2), strSplit(4), True, 1) = True
                                DoEvents
                            Loop
                            
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
                        
                        Else
                            
                            If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                                lTemp = .cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent + 1)
                            Else
                                lTemp = -1
                            End If
                            
                            Do Until AddTile(lRowSelected, .cRow(lRowSelected).TileCurrent + 1, lTemp, vbNullString, "Pinned | " & strSplit(2), strSplit(4), sSource, 1) = True
                                DoEvents
                            Loop
                            
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent + 1) = True
                        
                        End If
            
                    Case "cGames"
                    
                        'loadfile##file##name##thumb##isemulator##emulatorpath##hasquotes##subcaption
                    
                        sSource = "game##" & strSplit(1) & "##" & strSplit(2) & "##" & strSplit(3) & "##" & strSplit(4) & "##" & strSplit(5) & "##" & strSplit(6) & "##" & strSplit(7)
    
                        If bPinReplace = True Then
                        
                            Do Until UpdateTile(lRowSelected, .cRow(lRowSelected).TileCurrent, sSource, vbNullString, "Pinned | " & strSplit(2), strSplit(3), True, 0) = True
                                DoEvents
                            Loop
                            
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
                        
                        Else
                            
                            If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                                lTemp = .cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent + 1)
                            Else
                                lTemp = -1
                            End If
                            
                            Do Until AddTile(lRowSelected, .cRow(lRowSelected).TileCurrent + 1, lTemp, vbNullString, "Pinned | " & strSplit(2), strSplit(3), sSource, 0) = True
                                DoEvents
                            Loop
                            
                            .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent + 1) = True
                        
                        End If
                                            
                    Case "cMusicHome"
                    
                        Select Case strSplit(1)
                        
                            Case "radio"
                            
                                sSource = "radio##" & strSplit(2) & "##" & strSplit(3) & "##" & strSplit(4)
            
                                If bPinReplace = True Then
                                
                                    Do Until UpdateTile(lRowSelected, .cRow(lRowSelected).TileCurrent, sSource, vbNullString, "Pinned | " & strSplit(2), strSplit(4), True, 0) = True
                                        DoEvents
                                    Loop
                                    
                                    .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
                                
                                Else
                                    
                                    If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                                        lTemp = .cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent + 1)
                                    Else
                                        lTemp = -1
                                    End If
                                    
                                    Do Until AddTile(lRowSelected, .cRow(lRowSelected).TileCurrent + 1, lTemp, vbNullString, "Pinned | " & strSplit(2), strSplit(4), sSource, 0) = True
                                        DoEvents
                                    Loop
                                    
                                    .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent + 1) = True
                                
                                End If
                                
                            Case "musicartist"
                                    
                                sSource = "musicartist##" & strSplit(2) & "##" & strSplit(3) & "##"
            
                                If bPinReplace = True Then
                                
                                    Do Until UpdateTile(lRowSelected, .cRow(lRowSelected).TileCurrent, sSource, vbNullString, "Pinned | " & strSplit(2), strSplit(3), True, 0) = True
                                        DoEvents
                                    Loop
                                    
                                    .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
                                
                                Else
                                    
                                    If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                                        lTemp = .cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent + 1)
                                    Else
                                        lTemp = -1
                                    End If
                                    
                                    Do Until AddTile(lRowSelected, .cRow(lRowSelected).TileCurrent + 1, lTemp, vbNullString, "Pinned | " & strSplit(2), strSplit(3), sSource, 0) = True
                                        DoEvents
                                    Loop
                                    
                                    .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent + 1) = True
                                
                                End If
                                
                            Case "smartmix"
                            
                            
                            Case "album"
                            
                            
                        End Select
            
                End Select
                
            Case "notify"
            
                If UBound(strSplit()) = 1 Then
                    Notify strSplit(1), cSource
                ElseIf UBound(strSplit()) = 1 Then
                    Notify strSplit(1), cSource, CLng(strSplit(2))
                End If
                
            Case "playpause"
            
                PlayPause
            
            Case "playlistback"
            
                PlaylistBack
            
            Case "playlistforward"
            
                PlaylistNext
            
            Case "truncateplaylist"
            
                TruncatePlaylist
            
            Case "toggleshuffle"
            
                PlaybackShuffleToggle
            
            Case "togglerepeat"
                
                PlaybackRepeatToggle
                
            Case "nowplayingshow"
            
                PlayBack.NowPlayingShow = True
                
            Case "nowplayinghide"
            
                PlayBack.NowPlayingShow = False
                
            Case "showplaylist"
            
                Do Until ShowPlaylist = True
                    DoEvents
                Loop
                
            Case "sync"
            
                Select Case strSplit(1)
                
                    Case "status"
                    
                        'sync##status#file
                        
                        MessageRunningApp "sync##status##" & SyncStatus(strSplit(2))
            
                    Case "add"
                    
                        'sync##add##file##title##artist
            
                        SyncAdd strSplit(2), strSplit(3), strSplit(4), strSplit(5)
            
                    Case "remove"
                    
                        'sync##remove##file##title##artist
            
                        SyncRemove strSplit(2), strSplit(3), strSplit(4), strSplit(5)
            
                    Case "clear"
                    
                        'sync##clear
                        
                        SyncClear
            
                End Select
                
            Case "schedule"
                
                If strSplit(2) <> "" Then
                    
                    If CLng(strSplit(1)) > UBound(Events()) Then
                    
                        ReDim Preserve Events(CLng(strSplit(1)))
                        
                    End If
                    
                    Events(UBound(Events())).Name = strSplit(2)
                    Events(UBound(Events())).Location = strSplit(3)
                    Events(UBound(Events())).Start = GetEventTime(strSplit(4))
                    Events(UBound(Events())).End = GetEventTime(strSplit(5))
                    Events(UBound(Events())).Color = strSplit(6)
                    Events(UBound(Events())).Notified = False
                    Events(UBound(Events())).Thumb = strSplit(7)
                    
                    .cNotificationWidget.ItemSet Events(UBound(Events())).Name, Events(UBound(Events())).Start & " - " & Events(UBound(Events())).End, Events(UBound(Events())).Location, Events(UBound(Events())).Thumb
                    
                    'Do Until UpdateTile(lCalendarTile(0), lCalendarTile(1), "app##cCalendar##1", (CLng(strSplit(1)) + 1) & " events", "Schedule", "/images/dash/home/1.png", False) = True
                    '    DoEvents
                    'Loop
                    
                    'If (lRowSelected = lCalendarTile(0)) And (.cRow(lCalendarTile(0)).TileCurrent = lCalendarTile(1) And bDashLoaded = True) Then
                    '    .cRow(lCalendarTile(0)).TileSelected(lCalendarTile(1)) = True
                    'End If
            
                End If
                
            Case "clearschedule"
            
                .cNotificationWidget.ClearItems
                
            Case "deviceinserted"
            
                Notify "Device inserted"
                
            Case "deviceremoved"
            
                Notify "Device removed"
                
            Case "lovetrack"
            
                If PlayBack.Source = audio Then
            
                    lTemp = .cLastFM(1).LoveTrack(strSplit(1), strSplit(2))
                    If LCase(strSplit(1)) = LCase(PlayBack.Artist) And LCase(strSplit(2)) = LCase(PlayBack.Title) Then
                        If lTemp = 0 Then
                            .cMusicHome.SendMessage "nowplayingloved"
                        Else
                            .cMusicHome.SendMessage "togglelovefailed"
                        End If
                    End If
                
                End If
                
            Case "unlovetrack"
            
                If PlayBack.Source = audio Then
            
                    lTemp = .cLastFM(1).UnloveTrack(strSplit(1), strSplit(2))
                    If LCase(strSplit(1)) = LCase(PlayBack.Artist) And LCase(strSplit(2)) = LCase(PlayBack.Title) Then
                        If lTemp = 0 Then
                            .cMusicHome.SendMessage "nowplayingunloved"
                        Else
                            .cMusicHome.SendMessage "togglelovefailed"
                        End If
                    End If
                End If
                
        End Select
        
    End With

    ProcessMessage = True

End Function

Public Function MessageRunningApp(sMessage As String, Optional pImage As GDIpImage = Nothing) As Boolean

    On Error Resume Next
    
    Dim ctl As Control
    
    For Each ctl In frmMain.Controls
        DoEvents
        If LCase(ctl.Name) = LCase(CurrentApp.Name) Then
            ctl.SendMessage sMessage, pImage
            Exit For
        End If
    Next

    MessageRunningApp = True

End Function

Public Function Notify(sNotification As String, Optional Source As Control, Optional lRetain As Long = 0, Optional PlaySound As Boolean = False, Optional Rumble As Boolean = True, Optional pIcon As GDIpImage = Nothing) As Boolean

    On Error Resume Next
    
    With frmMain

        If bDashLoaded = True And bDashLocked = False And .pVideoHolder.Visible = False Then
            
            bNotify = True
            
            bStatusVisible = True
            
            lNotify = 0
            
            If CountControllers > 0 And Rumble = True Then
                .tmrRumble.Enabled = True
            End If
            
            If Not pIcon Is Nothing Then
                .pNotifyIcon.Picture = pIcon
            Else
                .pNotifyIcon.Picture = .pNotifyDefault(0).Picture
            End If
            
            .pNotifyHolder.ZOrder 0
            .pNotifyHolder.Visible = True
            
            .lblStatus.Caption = "Notification | " & sNotification
            .lblStatus.Caption = sNotification
            
            'If PlaySound = True Then
                Sound "notify"
            'End If
        
            Do Until RemoteSend("NOTIFY##" & .lblStatus.Caption) = True
                DoEvents
            Loop
        
            .tmrNotify.Enabled = True
        
            If lRetain = 1 Then
            
            End If
        
        End If
        
    End With

    Notify = True

End Function

