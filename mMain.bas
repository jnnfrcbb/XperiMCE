Attribute VB_Name = "mMain"

Public Enum eCurrentScreen
    Home = 0
    RunningApp = 1
    Search = 2
    SearchResults = 3
End Enum
Public CurrentScreen As eCurrentScreen
Public PreviousScreen As eCurrentScreen

Public ResizeState As Long

Public lCharmSelected As Long


Public Function InitMain() As Boolean

    With frmMain
    
        .ResizeOld = .ScaleWidth
       
        frmSplash.Show
        
        frmSplash.ZOrder 0
        
        bDashLoaded = False
    
        Dim cControl As Control
    
        For Each cControl In .Controls
            
            If TypeOf cControl Is Label Then
            
                cControl.UseMnemonic = False
            
            End If
        
        Next
            
        lWeatherTile(0) = -1
        lNewsTile(0) = -1
        lSportTile(0) = -1
        lCalendarTile(0) = -1
        
        .pFocus.Left = 0
       
        .pMenuHolder.Left = -.pMenuHolder.Width

        'PreventMonitorSleeping
    
    End With
    
    InitMain = True

End Function


Public Function OpenCharms() As Boolean

    On Error Resume Next
    
    Sound "openelement"
    
    If bDashLocked = False And bAppDirect = False Then
        
        Static i As Integer
    
        With frmMain
        
            For i = 0 To .pCharm.UBound
                .pCharm(i).TransparencyPct = 75
            Next
            
            .pCharm(3).TransparencyPct = 0
            
            lCharmSelected = 3
                    
            .pCharmsHolder.Visible = True
            
            .tmrCharmsOpen.Enabled = True
            
            Do Until .tmrCharmsOpen.Enabled = False
                DoEvents
            Loop
            
            .pCharmsHolder.ZOrder 0
            
        End With
    
    End If
    
    OpenCharms = True

End Function

Public Function CloseCharms() As Boolean

    On Error Resume Next
    
    Sound "closeelement"

    With frmMain
    
        .tmrCharmsClose.Enabled = True
        
        .pFocus.ZOrder 0
        
        .pVideoHolder.ZOrder 0
        
        Do Until .tmrCharmsClose.Enabled = False
            DoEvents
        Loop
        
        Select Case CurrentScreen
            Case 0
                .cRow(lRowSelected).TileSelected(.cRow(lRowSelected).TileCurrent) = True
            Case 1 'running app
                MessageRunningApp "returndown"
            Case 2 'search
            Case 3 'search results
                .cSearch.TileSelected(.cSearch.TileCurrent) = True
        End Select
                    
        .pCharmsHolder.Visible = False
            
    End With
    
    CloseCharms = True

End Function

Public Function CharmsKey(KeyCode As Integer, Shift As Integer) As Boolean

    On Error Resume Next

    Static i As Integer
    
    With frmMain
    
        Select Case KeyCode
        
            Case vbKeyLeft
            
                If lCharmSelected > 0 Then
                
                    frmMain.pCharm(lCharmSelected).TransparencyPct = 75
                    
                    lCharmSelected = lCharmSelected - 1
                    
                    Sound "select"
                    
                    frmMain.pCharm(lCharmSelected).TransparencyPct = 0
                
                Else
                
                    Sound "listend"
                
                End If
            
            Case vbKeyRight
            
                If lCharmSelected < frmMain.pCharm.UBound Then
                
                    frmMain.pCharm(lCharmSelected).TransparencyPct = 75
                    
                    lCharmSelected = lCharmSelected + 1
                    
                    Sound "select"
                    
                    frmMain.pCharm(lCharmSelected).TransparencyPct = 0
                
                Else
                
                    Sound "listend"
                
                End If
            
            Case vbKeyReturn
            
                Do Until ProcessCharm(lCharmSelected) = True
                    DoEvents
                Loop
                
            Case vbKeyDown, vbKeyBack, vbKeyF1
    
                CloseCharms
            
        End Select

    End With

    CharmsKey = True

End Function

Public Function ProcessCharm(lCharm As Long) As Boolean

    On Error Resume Next
    
    Sound "yes"

    Dim ctl As Control
            
    With frmMain

        Do Until CloseCharms = True
            DoEvents
        Loop
        
        Select Case lCharm
        
            Case 0 'settings
            
            Case 1 'power
            
                BeginCloseExe
                
            Case 2 'lock
            
                Do Until ShowLock = True
                    DoEvents
                Loop
            
            Case 3 'show dash
            
                If CurrentScreen = RunningApp Then
                    For Each ctl In .Controls
                        DoEvents
                        If ctl.Name = CurrentApp.Name Then
                            CloseApp ctl
                            Exit For
                        End If
                    Next
                End If
                
                Do Until RestoreDash(True) = True
                    DoEvents
                Loop
                
            Case 4 'playlist
            
                Do Until ShowPlaylist = True
                    DoEvents
                Loop
            
            Case 5 'search
            
                If CurrentScreen = RunningApp Then
                    For Each ctl In .Controls
                        DoEvents
                        If ctl.Name = CurrentApp.Name Then
                            CloseApp ctl
                            Exit For
                        End If
                    Next
                End If
                
                Do Until OpenSearch = True
                    DoEvents
                Loop
            
            Case 6 'minrestore
            
                frmMain.Visible = False
                
                If frmMain.WindowState = vbNormal Then
                    frmMain.WindowState = vbMaximized
                    xmlSettingsDoc.selectSingleNode("//window").Attributes.getNamedItem("fullscreen").Text = "yes"
                    xmlSettingsDoc.save App.Path & "\settings.xml"
                ElseIf frmMain.WindowState = vbMaximized Then
                    frmMain.WindowState = vbNormal
                    xmlSettingsDoc.selectSingleNode("//window").Attributes.getNamedItem("fullscreen").Text = "no"
                    xmlSettingsDoc.save App.Path & "\settings.xml"
                End If
                
                frmMain.Visible = True
                
            Case Else
            
        End Select
        
    End With
    
    ProcessCharm = True

End Function


Public Function Loading(bLoading As Boolean)
    
    'On Error Resume Next

    Dim i As Integer

    If bLoading = True Then
        For i = 0 To frmMain.pLoading.UBound
            frmMain.pLoading(i).Animate lvicAniCmdStart
        Next
    Else
        For i = 0 To frmMain.pLoading.UBound
            frmMain.pLoading(i).Animate lvicAniCmdStop
        Next
    End If
    
    If bLoading = False Then
        frmMain.pSearchIcon.Visible = eDashProp(lRowSelected).Search
        frmMain.pSideFadeIcon(1).Visible = eDashProp(lRowSelected).Search
    ElseIf bLoading = True Then
        frmMain.pSearchIcon.Visible = False
        frmMain.pSideFadeIcon(1).Visible = False
    End If

    frmMain.pMenuIcon.Visible = Not bLoading
    
    For i = 0 To frmMain.pLoading.UBound
        frmMain.pLoading(i).Visible = bLoading
    Next
        
End Function

Public Function MinRestoreVideo() As Boolean

    On Error Resume Next
    
    With frmMain
    
        If PlayBack.Source = video Or PlayBack.Source = audio Then
        
            If bMinVideo = True Then
            
                .pVideoHolder.Move 0, 0, .ScaleWidth, .ScaleHeight
                
                .pVideoHolder.BorderStyle = 0
                
                bOSDAvailable = True
                
                bMinVideo = False
                
            Else
                
                ShowOSD False
            
                .pVideoHolder.Move .pMinVideo.Left, .pMinVideo.Top, .pMinVideo.Width, .pMinVideo.Height
                
                .pVideoHolder.BorderStyle = 1
                
                bOSDAvailable = True
                
                bMinVideo = True
                
            End If
        
            Select Case PlayBack.Source
            
                Case video
                
                    .wmp.Move 0, 0, .pVideoHolder.Width, .pVideoHolder.Height
                
                Case audio
                
            End Select
            
            .pVideoHolder.ZOrder 0
            
        End If
    
    End With

    MinRestoreVideo = True

End Function
