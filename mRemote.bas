Attribute VB_Name = "mRemote"
Public sRemoteQueue() As String
Public bRemoteSendReady As Boolean
Public bRemoteConnected As Boolean

Public lClientTimeout As Long

Public sWeatherString As String

Public Function RemoteSend(sString As String) As Boolean

    With frmMain
    
        If bRemoteConnected = True Then
            
            .lstRemoteQueue.AddItem sString & "#|#"
            
            sString = vbNull
        
        End If
        
    End With
    
    RemoteSend = True

End Function

Public Function RemoteConnected() As Boolean

    Dim sServerAddress As String

    Dim dr As Drive

    With frmMain
    
        Notify "Remote Connection Established", , , True, , .pServiceIcon(1).Picture
        
        bRemoteConnected = True
    
        bRemoteSendReady = True
        
        SetServiceIcon 1, sOn
        
        SetClientTimeout False
    
        .tmrRemote.Enabled = True
        
        ServerReady
    
        'If bNotify = True Then
        '    Do Until bNotify = False
        '        DoEvents
        '    Loop
        'End If
        
        sServerAddress = "\" & fso.GetDrive(Mid(App.Path, 1, 3)).VolumeName & "##" & Mid(App.Path, 3) & "##" & Mid(App.Path, 1, 1)
        
        Do Until RemoteSend("APP##" & sServerAddress) = True
            DoEvents
        Loop
        
        Do Until RemoteSend("STATUS##" & .lblStatus.Caption) = True
            DoEvents
        Loop
            
        Do Until RemoteSend(sWeatherString) = True
            DoEvents
        Loop
        
        If PlayBack.File <> "" Then
            
            Select Case PlayBack.State
                Case ePlaying
                    Do Until RemoteSend("CONTROL##PLAYING") = True
                        DoEvents
                    Loop
                Case Else
                    Do Until RemoteSend("CONTROL##PAUSED") = True
                        DoEvents
                    Loop
            End Select
            
            Do Until RemoteSend("NOWPLAYING##" & PlayBack.Title & "##" & PlayBack.Artist & "##" & PlayBack.Album & "##" & PlayBack.Thumb) = True
                DoEvents
            Loop
            
        End If
        
        ServerReady
        
    End With
    
    RemoteConnected = True

End Function

Public Function RemoteCommand(sCommand As String) As Boolean

    Dim strCommand() As String
    Dim strSplit() As String
    Dim sSource As String
    Dim i As Long
    
    strCommand() = Split(sCommand, "#|#")
    
    For i = 0 To UBound(strCommand())
            
        If strCommand(i) <> "" Then
            
            strSplit() = Split(strCommand(i), "##")
            
            If strSplit(0) <> "" Then
                
                Select Case LCase(strSplit(0))
                
                    Case "control"
                    
                        Select Case LCase(strSplit(1))
                        
                            Case "playpause"
                            
                                PlayPause
                            
                            Case "previous"
                            
                                PlaylistBack
                            
                            Case "next"
                            
                                PlaylistNext
                            
                            Case "mute"
                            
                                Do Until VolumeToggleMute = True
                                    DoEvents
                                Loop
                        
                        End Select
                        
                    Case "music"
                    
                        Select Case LCase(strSplit(1))
                        
                            Case "load"
                            
                                Do Until (PlaylistClear) = True
                                    DoEvents
                                Loop
                                
                                Do Until PlaylistAdd(strSplit(2), strSplit(3), 0, 0, strSplit(4)) = True
                                    DoEvents
                                Loop
                            
                                Do Until PlaylistLoad(0) = True
                                    DoEvents
                                Loop
                                
                            Case "queue"
                                    
                                Do Until PlaylistAdd(strSplit(2), strSplit(3), 0, 0, strSplit(4)) = True
                                    DoEvents
                                Loop
                
                            Case "radio"
                            
                                Do Until (PlaylistClear) = True
                                    DoEvents
                                Loop
                                
                                Do Until PlaylistAdd(strSplit(2), strSplit(3), 0, 3, strSplit(4)) = True
                                    DoEvents
                                Loop
                            
                                Do Until PlaylistLoad(0) = True
                                    DoEvents
                                Loop
                                
                        End Select
                        
                    Case "video"
                    
                        Select Case LCase(strSplit(1))
                        
                                   
                        End Select
                        
                    Case "playlist"
                    
                        Select Case LCase(strSplit(1))
                        
                            Case "clear"
                            
                                PlaylistClear
                        
                            Case "loaditem"
                            
                                PlaylistLoad CLng(strSplit(2))
                            
                        End Select
                        
                    Case "ping"
                    
                        'Do Until RemoteSend("PING##")
                        '    DoEvents
                        'Loop
                        
                End Select
                
            End If

        End If

    Next
           
    ServerReady
           
End Function

Public Function SetClientTimeout(bEnabled As Boolean)

    lClientTimeout = 0

    frmMain.tmrClientTimeout.Enabled = bEnabled
    
End Function

Public Function ServerReady() As Boolean

    Do Until RemoteSend("READY##")
        DoEvents
    Loop
             
    ServerReady = True
    
End Function
