Attribute VB_Name = "mDash"
'Option Explicit

Public Type eDashProperties
    Index As Long
    Title As String
    FullMain As Boolean
    TileColor As OLE_COLOR
    TileColorRGB As String
    TileBorder As Boolean
    Search As Boolean
End Type

Public Type tTile
    Index As Long
    Title As String
    SubTitle As String
    Thumb As String
    Source As String
    SubSource As Long
    Locked As Boolean
End Type

Public Type tPinItems
    Index As Long
    Title As String
    Thumb As String
    App As String
End Type

Public Tiles() As tTile
Public eDashProp() As eDashProperties

Public PinItems() As tPinItems

Public bShowWelcome As Boolean

Public bDashLoaded As Boolean

Public lMenu As Long
Public lMenuSelect As Long

Public lWeatherTile(1) As Long
Public lNewsTile(1) As Long
Public lSportTile(1) As Long
Public lCalendarTile(1) As Long

Public lMusicRow As Long
Public lVideoRow As Long
Public lGamesRow As Long

Public bWeatherBack As Boolean
Public bUniversalBack As Boolean
Public bMenuLeft As Boolean

Public lTileOption As Long

Public lCharmCount As Long

Public lPinRow As Long
Public lPinTile As Long
Public bPinLocked As Boolean
Public bPinReplace As Boolean

Public bDashLocked As Boolean

Public bShowTopBar As Boolean
Public bBottomBarTrans As Boolean

Public Enum eServiceState
    sOff = 0
    sPreparing = 1
    sOn = 2
End Enum

Public sWeatherStatus As String

Public sAlarmSource As String
Public sAlarmThumb As String

Public lRowSelected As Long
Public lTileSelected As Long

Public Const cRowTitleSelected = vbWhite
Public Const cRowTitleDeselected = &HCDCDCD

Public RowLoaded() As Boolean
Public lScroll() As Long
Public Const lScrollSpeed As Long = 6

Public Function InitDashboard() As Boolean
    
    Loading True
    
    ' on error resume next

    Dim i As Long
    Dim r As Long
    Dim t As Long
    Dim sCaption As String
    Dim sSubCaption As String
    Dim sThumb As String
    Dim strSplit() As String
                 
    With frmMain
               
        ReDim RowLoaded(0)
               
        ReDim Tiles(t)
                        
        For i = 0 To UBound(eDashProp())
                
            DoEvents
                
            Do Until LoadRow(i, eDashProp(i).Title) = True
                DoEvents
            Loop
            
            r = 0
            
            For Each xmlSetting In xmlSettingsDoc.selectNodes("//dashboard/dash[" & eDashProp(i).Index & "]/item")
                
                DoEvents
                
                If xmlSetting.Attributes.getNamedItem("visible").Text = "yes" Then
                    
                    ReDim Preserve Tiles(t)
                    
                    'get tile locked
                        If xmlSetting.Attributes.getNamedItem("locked").Text = "yes" Then
                            Tiles(t).Locked = True
                        Else
                            Tiles(t).Locked = False
                        End If
                        
                    'get caption
                        If xmlSetting.Attributes.getNamedItem("showtitle").Text = "yes" Then
                            If xmlSetting.Attributes.getNamedItem("title").Text <> "" Then
                                sCaption = xmlSetting.Attributes.getNamedItem("title").Text
                            Else
                                sCaption = ""
                            End If
                        Else
                            sCaption = ""
                        End If
                        
                    'get subcaption
                        If xmlSetting.Attributes.getNamedItem("subtitle").Text <> "" Then
                            sSubCaption = xmlSetting.Attributes.getNamedItem("subtitle").Text
                        Else
                            sSubCaption = ""
                        End If
                    
                    'get thumb
                        If xmlSetting.Attributes.getNamedItem("thumb").Text <> "" Then
                            If Mid(xmlSetting.Attributes.getNamedItem("thumb").Text, 1, 1) = "/" Then
                                sThumb = App.Path & xmlSetting.Attributes.getNamedItem("thumb").Text
                            Else
                                sThumb = xmlSetting.Attributes.getNamedItem("thumb").Text
                            End If
                        Else
                            sThumb = ""
                        End If
                
                    .cRow(i).TileSet r, sCaption, sSubCaption, sThumb, , CStr(t)
                    
                    Tiles(t).Index = CLng(xmlSetting.Attributes.getNamedItem("index").Text)
                    Tiles(t).Title = sCaption
                    Tiles(t).Thumb = sThumb
                    Tiles(t).Source = xmlSetting.Attributes.getNamedItem("source").Text
                    
                    strSplit() = Split(Tiles(t).Source, "##")
                    
                    Select Case strSplit(0)
                    
                        Case "app"
                        
                            Select Case strSplit(1)
                            
                                Case "cWeather"
                                
                                    lWeatherTile(0) = i
                                    lWeatherTile(1) = r
                                    
                                Case "cNews"
                                
                                    lNewsTile(0) = i
                                    lNewsTile(1) = r
                                
                                Case "cSport"
                                
                                    lSportTile(0) = i
                                    lSportTile(1) = r
                                    
                                Case "cCalendar"
                                
                                    lCalendarTile(0) = i
                                    lCalendarTile(1) = r
                                
                            End Select
                        
                        Case "file", "ytvideo"
                        
                            Tiles(t).SubSource = CLng(strSplit(4))
                            
                        Case "live"
                        
                            Select Case strSplit(1)
                            
                            End Select
                        
                    End Select
                    
                    If .cRow(i).Top > 0 And .cRow(i).Top < .ScaleHeight Then
                        
                        .tmrWait.Enabled = True
                        
                        Do Until .tmrWait.Enabled = False
                            DoEvents
                        Loop
                        
                    End If
                    
                    r = r + 1
                
                    t = t + 1
                
                End If
                
            Next
            
            If r > 1 Then
                .lblRow(i).Caption = .lblRow(i).Caption & " | " & r & " Tiles"
            Else
                .lblRow(i).Caption = .lblRow(i).Caption & " | " & r & " Tile"
            End If
            
            bDashLocked = .pLockHolder.Visible
            
            If i = 1 Then
            
                InitDashboard = True

            End If
                
        Next
    
        .cPlaylist.TileColorSelected = eDashProp(0).TileColor
    
        Loading False
                
    End With
    
    bDashLoaded = True

End Function

Private Function RandomInteger(lowerBound As Integer, upperBound As Integer) As Integer

    ' on error resume next

    Randomize
        
    RandomInteger = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)

End Function

Public Function DashKey(KeyCode As Integer, Shift As Integer) As Boolean

    Dim bTemp As Boolean
    Dim strSplit() As String

    ' on error resume next

    With frmMain
    
        Select Case KeyCode
        
            Case vbKeyLeft
            
                If .cRow(lRowSelected).TileCurrent > 0 Then
                    .cRow(lRowSelected).TileSelectLeft
                Else
                    If bMenuLeft = True Then
                        Do Until OpenMenu = True
                            DoEvents
                        Loop
                    Else
                        Sound "listend"
                    End If
                End If
            
            Case vbKeyRight
            
                .cRow(lRowSelected).TileSelectRight
            
            Case vbKeyUp
            
                Do Until SelectRow(lRowSelected - 1) = True
                    DoEvents
                Loop
            
            Case vbKeyDown
            
                Do Until SelectRow(lRowSelected + 1) = True
                    DoEvents
                Loop
            
            Case vbKeyBack
            
                Do Until ShowTileOptions = True
                    DoEvents
                Loop
            
            Case vbKeyReturn
            
                'If .cRow(lRowSelected).TileCurrent < .cRow(lRowSelected).TileCount Then
                
                    Sound "yes"
                
                    Do Until ProcessTile(CLng(.cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent))) = True
                        DoEvents
                    Loop
                    
                'Else
                
                '    Do Until InitPinSelect(False) = True
                '        DoEvents
                '    Loop
                
                'End If
            
            Case vbKeyShift
            
                If eDashProp(lRowSelected).Search = True Then
                
                    Do Until OpenSearch = True
                        DoEvents
                    Loop
                    
                End If

            Case vbKeyMenu
              
                Do Until OpenMenu = True
                    DoEvents
                Loop
                                            
        End Select

    End With
    
    DashKey = True

End Function


Public Function OpenSearch() As Boolean

    ' on error resume next
    
    Sound "openelement"

    Dim i As Integer

    With frmMain

        Do Until ClearDash(True) = True
            DoEvents
        Loop
        
        .pTopBar.Visible = True
        
        .lblRow(lRowSelected).ForeColor = cRowTitleDeselected
    
        .pAppLogo.Visible = False
        .sSearchBack.Width = 0
        .sSearchBack.Visible = True
        
        .pSideFadeBox(1).Visible = True
        .pSideFadeIcon(1).Visible = True
        .pSearchHolder.Visible = True
        
        .cKeyboard.MainKeyDeselectedBackColor = &H4B4B4B
        .cKeyboard.FunctionKeyDeselectedBackColor = &H3C3C3C
        .cKeyboard.MainKeySelectedBackColor = eDashProp(lRowSelected).TileColor
        .cKeyboard.FunctionKeySelectedBackColor = eDashProp(lRowSelected).TileColor
        
        .pSearchIcon.TransparencyPct = 0
    
        .tmrOpenSearch.Enabled = True
        
        Do Until .tmrOpenSearch.Enabled = False
            DoEvents
        Loop
        
        .pSearchIcon.Visible = True
        
        .lblSearch.ForeColor = cRowTitleDeselected
        
        .lblSearch.Caption = "Type to search " & LCase(eDashProp(lRowSelected).Title) & "..."
        
        .lblSearch.Visible = True
        
        .cKeyboard.Visible = True
        
        Do Until .cKeyboard.InitialFocus = True
            DoEvents
        Loop
        
        PreviousScreen = CurrentScreen
        CurrentScreen = Search
                
    End With
        
    OpenSearch = True

End Function

Public Function OpenMenu() As Boolean

    ' on error resume next
    
    Static i As Integer
    
    With frmMain
        
        .lblMenu(0).BackStyle = 1
        .lblMenu(0).BackColor = eDashProp(lRowSelected).TileColor
        .lblMenu(lMenu).ForeColor = vbWhite
        
        If .lblMenu.UBound > 0 Then
            For i = 1 To .lblMenu.UBound
                .lblMenu(i).BackStyle = 0
                .lblMenu(i).BackColor = eDashProp(lRowSelected).TileColor
                .lblMenu(i).ForeColor = &HE0E0E0
            Next
        End If
        
        lMenu = 0
        
        Select Case CurrentScreen
            Case 0 'home
                lMenuSelect = .cRow(lRowSelected).TileCurrent
                .cRow(lRowSelected).Deselect
                .lblRow(lRowSelected).ForeColor = cRowTitleDeselected
            Case 2
            Case 3
                .cSearch.Deselect
        End Select
        
        Sound "openelement"
        
        .pSideFadeBox(0).Visible = True
        
        .pMenuHolder.Width = 1
        .pMenuHolder.Left = .pSideFade.Width
        .pMenuHolder.Visible = True
                
        .tmrOpenMenu.Enabled = True
        
        Do Until .tmrOpenMenu.Enabled = False
            DoEvents
        Loop
        
    End With
    
    OpenMenu = True
    
End Function

Public Function CloseMenu() As Boolean

    ' on error resume next
    
    With frmMain
    
        Do Until ReselectDash = True
            DoEvents
        Loop

        Sound "closeelement"
    
        .tmrCloseMenu.Enabled = True
        
        Do Until .tmrCloseMenu.Enabled = False
            DoEvents
        Loop
        
        .pMenuHolder.Visible = False
        .pSideFadeBox(0).Visible = False
        
    End With
    
    CloseMenu = True

End Function

Public Function MenuSelect(Index As Long) As Boolean

    ' on error resume next
    
    Do Until CloseMenu = True
        DoEvents
    Loop
                
    Select Case Index
    
        Case 0 'home
            
            Select Case CurrentScreen
            
                Case 2 'search
                    
                    Do Until CloseSearch = True
                        DoEvents
                    Loop
            
                Case 3 'search results
            
                    Do Until CloseSearchResults(True) = True
                        DoEvents
                    Loop
                            
            End Select
        
        Case 1 'playlist

            Do Until ShowPlaylist = True
                DoEvents
            Loop
                    
        Case 2 'devices
        
            Do Until OpenApp(frmMain.cDevices) = True
                DoEvents
            Loop
        
        Case 3 'lock
        
            Do Until ShowLock = True
                DoEvents
            Loop
        
        Case 4 'exit
     
            BeginCloseExe
     
    End Select
    
    MenuSelect = True

End Function

Public Function MenuKey(KeyCode As Integer, Shift As Integer) As Boolean

    ' on error resume next
    
    With frmMain
    
        Select Case KeyCode
        
            Case vbKeyUp
            
                If lMenu > 0 Then
                    
                    .lblMenu(lMenu).BackStyle = 0
                    .lblMenu(lMenu).ForeColor = &HE0E0E0
                    
                    lMenu = lMenu - 1
                    
                    Sound "select"
                    
                    .lblMenu(lMenu).BackStyle = 1
                    .lblMenu(lMenu).BackColor = eDashProp(lRowSelected).TileColor
                    .lblMenu(lMenu).ForeColor = vbWhite
                
                Else
                
                    Sound "listend"
                
                End If
            
            Case vbKeyDown
            
                If lMenu < .lblMenu.UBound Then
            
                    .lblMenu(lMenu).BackStyle = 0
                    .lblMenu(lMenu).ForeColor = &HE0E0E0
                    
                    lMenu = lMenu + 1
                    
                    Sound "select"
                    
                    .lblMenu(lMenu).BackStyle = 1
                    .lblMenu(lMenu).BackColor = eDashProp(lRowSelected).TileColor
                    .lblMenu(lMenu).ForeColor = vbWhite
                
                Else
                
                    Sound "listend"
                
                End If
                
            Case vbKeyReturn
        
                Do Until MenuSelect(lMenu) = True
                    DoEvents
                Loop
        
            Case vbKeyMenu, 93, vbKeyBack, vbKeyRight
            
                Do Until CloseMenu = True
                    DoEvents
                Loop
            
        End Select
            
    End With
    
    MenuKey = True

End Function

Public Function CloseSearch() As Boolean
    
    ' on error resume next
    
    Static i As Integer
    
    With frmMain
    
        Sound "closeelement"
    
        Do Until .cKeyboard.CloseControl = True
            DoEvents
        Loop
    
        .cKeyboard.Visible = False
        
        .lblSearch.Visible = False
        
        .tmrCloseSearch.Enabled = True
        
        Do Until .tmrCloseSearch.Enabled = False
            DoEvents
        Loop
                
        .pSearchIcon.TransparencyPct = 60
        
        .sSearchBack.Visible = False
        .pSideFadeBox(1).Visible = False
        .pSearchHolder.Visible = False
        .lblAppTitle.Visible = True
        
        UpdateTopBar
        
        .lblRow(lRowSelected).ForeColor = cRowTitleSelected

        UpdateTopBar

        Do Until RestoreDash(True) = True
            DoEvents
        Loop
        
        CurrentScreen = Home
        
    End With
        
    CloseSearch = True

End Function

Public Function SearchKey(KeyCode As Integer, Shift As Integer) As Boolean
    
    ' on error resume next
    
    With frmMain

        Select Case KeyCode
        
            Case vbKeyBack
            
                If .cKeyboard.CurrentString = "" Then
                    Do Until CloseSearch = True
                        DoEvents
                    Loop
                Else
                    .cKeyboard.KeyPress KeyCode, Shift
                End If
        
            Case vbKeyShift
        
                Do Until CloseSearch = True
                    DoEvents
                Loop
                
            Case vbKeyMenu, 93
              
                Do Until OpenMenu = True
                    DoEvents
                Loop
                 
            Case Else
            
                Sound "select"
            
                .cKeyboard.KeyPress KeyCode, Shift
            
        End Select
    
    End With

    SearchKey = True

End Function

Public Function UpdateTile(lDash As Long, lTile As Long, sSource As String, sTitle As String, sSubTitle As String, sThumb As String, bSave As Boolean, Optional lSubSource As Long = 0)
    
     On Error Resume Next
    
    If bDashLoaded = True Then
        
        Dim sString As String
        Dim bTitle As Boolean
    
        Dim lRowIndex As Long
        Dim lTileTag As Long
        Dim lTileIndex As Long
        
        lTileTag = CLng(frmMain.cRow(lDash).TileTag(lTile))
    
        lRowIndex = eDashProp(lDash).Index
    
        lTileIndex = Tiles(lTileTag).Index
        
        If Mid(sThumb, 1, 1) = "/" Or Mid(sThumb, 1, 1) = "\" Then
            sThumb = App.Path & sThumb
        End If
        
        With xmlSettingsDoc
    
            Tiles(lTileTag).Source = sSource
            Tiles(lTileTag).Thumb = sThumb
            Tiles(lTileTag).Title = sTitle
        
            If bSave = True Then
            
                sString = "//dashboard/dash[" & lRowIndex & "]/item[" & lTileIndex & "]"
            
                .selectSingleNode(sString).Attributes.getNamedItem("title").Text = sTitle
                .selectSingleNode(sString).Attributes.getNamedItem("subtitle").Text = sSubTitle
                .selectSingleNode(sString).Attributes.getNamedItem("thumb").Text = sThumb
                .selectSingleNode(sString).Attributes.getNamedItem("source").Text = sSource
                If sTitle <> vbNullString Then
                    xmlNodeOld.Attributes.getNamedItem("showtitle").Text = "yes"
                Else
                    xmlNodeOld.Attributes.getNamedItem("showtitle").Text = "no"
                End If
        
                .save App.Path & "\settings.xml"
                
            End If
                
            frmMain.cRow(lDash).TileSet lTile, sTitle, sSubTitle, sThumb, , CStr(lTileTag)
    
        End With
        
    End If

    UpdateTile = True
    
End Function

Public Function ProcessTile(Index As Long) As Boolean

    ' on error resume next
    
    Dim strSplit() As String
    Dim ctl As Control
    Dim i As Integer
    
    If Tiles(Index).Source <> "" Then
        
        strSplit() = Split(Tiles(Index).Source, "##")
        
        Select Case LCase(strSplit(0))
        
            Case "app"
            
                For Each Control In frmMain.Controls
                    DoEvents
                    If Control.Name = strSplit(1) Then
                        Set ctl = Control
                        Exit For
                    End If
                Next
            
                If UBound(strSplit) = 3 Then
                
                    Do Until OpenApp(ctl, strSplit(3), Tiles(Index).Thumb) = True
                        DoEvents
                    Loop
        
                Else
                
                    Do Until OpenApp(ctl, , Tiles(Index).Thumb) = True
                        DoEvents
                    Loop
                
                End If
               
            Case "audio"
                                           
            Case "video"
                        
            Case "game"
            
                Do Until LoadGame(strSplit(1), strSplit(2), strSplit(3), strSplit(4), strSplit(5), strSplit(6), strSplit(7), strSplit(8)) = True
                    DoEvents
                Loop

            Case "service"
            
            Case "file", "ytvideo"
            
                'filename##title##type##subsource
                
                Do Until PlaylistClear = True
                    DoEvents
                Loop
                
                Do Until PlaylistAdd(strSplit(1), strSplit(2), CLng(strSplit(3)), CLng(strSplit(4)), Tiles(Index).Thumb) = True
                    DoEvents
                Loop
                
                Do Until PlaylistLoad(0) = True
                    DoEvents
                Loop
                        
            Case "tvshow"
            
                Do Until OpenApp(frmMain.cTV, "initshow##" & strSplit(1)) = True
                    DoEvents
                Loop
                
            Case "musicartist"
            
                Do Until OpenApp(frmMain.cMusicHome, "initartist##" & strSplit(1)) = True
                    DoEvents
                Loop
            
            Case "radio"
                
                Do Until PlaylistClear = True
                    DoEvents
                Loop
                
                Do Until PlaylistAdd(strSplit(2), strSplit(1), audio, 3, Tiles(Index).Thumb) = True
                    DoEvents
                Loop
                
                Do Until PlaylistLoad(0) = True
                    DoEvents
                Loop
                                
            Case "external"
                       
                frmMain.tmrPlayback.Enabled = False
                frmMain.tmrMouse.Enabled = False
                frmMain.tmrMouseOut.Enabled = False
                 
                bControllerBypass = True
                
                HideCursor False
                
                sProcessPath = strSplit(1)
                
                Select Case frmMain.WindowState
                    Case vbNormal
                        ResizeState = 0
                    Case vbMinimized
                        ResizeState = 1
                    Case vbMaximized
                        ResizeState = 2
                End Select
                
                'frmMain.WindowState = vbMinimized

                If Shell(strSplit(1), vbMaximizedFocus) <> 0 Then
            
                    frmMain.tmrCheckForProcess.Enabled = True
                
                End If
                
        End Select
    
    End If
    
    ProcessTile = True
                                                  
End Function

Public Function ProcessSubSource(eSource As ePlaybackSources, lSubSource As Long) As String

    ' on error resume next
    
    Dim sSub As String

    Select Case eSource
    
        Case 0 'audio
        
            sSub = "Audio"

        Case 1 'video
        
            Select Case lSubSource
            
                Case 0
                
                    sSub = "Film"
                
                Case 1
                
                    sSub = "TV Episode"
                
            End Select
        
        Case 2 'dvd
        
            sSub = "DVD"
        
        Case 5 'game
        
            Select Case lSubSource
            
                Case 0
                
                    sSub = "PC Game"
                    
                Case 1
                
                    sSub = "N64 Game"
                    
                Case 2
                
                    sSub = "PS2 Game"
                    
            End Select
        
        Case 6 'external
        
            sSub = "External"
        
    End Select

    ProcessSubSource = sSub

End Function

Public Function ClearDash(bAnimate As Boolean) As Boolean

    ' on error resume next
    
    Dim i As Integer

    With frmMain
    
        For i = 0 To .cRow.UBound
            .lblRow(i).Visible = False
            .cRow(i).Visible = False
            If bAnimate = True Then
                If (.cRow(i).Top < .ScaleHeight And .cRow(i).Top > -.cRow(i).Height) Then
                    .tmrWait.Enabled = True
                    Do Until .tmrWait.Enabled = False
                        DoEvents
                    Loop
                End If
            End If
        Next
        
        .lblHomeTime.Visible = False
        
    End With

    ClearDash = True

End Function

Public Function RestoreDash(bAnimate As Boolean) As Boolean

    ' on error resume next
    
    Dim i As Integer

    With frmMain
    
        .lblHomeTime.Visible = True
        
        For i = 0 To .cRow.UBound
            .lblRow(i).Visible = True
            If i = lRowSelected Then
                .cRow(i).Visible = True
            End If
            If bAnimate = True Then
                If (.cRow(i).Top < .ScaleHeight And .cRow(i).Top > -.cRow(i).Height) Then
                    .tmrWait.Enabled = True
                    Do Until .tmrWait.Enabled = False
                        DoEvents
                    Loop
                End If
            End If
        Next
        
    End With

    RestoreDash = True

End Function

Public Function TileOptionsKey(KeyCode As Integer, Shift As Integer) As Boolean

    ' on error resume next
    
    Dim lLeft As Long
    
    If bPinLocked = True Then
        lLeft = 0
    Else
        lLeft = 0
    End If

    With frmMain
    
        If .pTilePinSelectHolder.Visible = True Then
        
            Do Until PinSelectKey(KeyCode, Shift) = True
                DoEvents
            Loop
            
        Else
    
            Select Case KeyCode
            
                Case vbKeyLeft
                
                    If lTileOption > lLeft Then
                        
                        .pTile(lTileOption).TransparencyPct = 50
                        .pTile(lTileOption).BackStyleOpaque = False
                        
                        If lTileOption = 4 And bPinLocked = True Then
                        
                            lTileOption = 1
                        
                        Else
                        
                            lTileOption = lTileOption - 1
                        
                        End If
                        
                        Sound "select"
                        
                        .pTile(lTileOption).TransparencyPct = 0
                        .pTile(lTileOption).BackColor = eDashProp(lRowSelected).TileColor
                        .pTile(lTileOption).BackStyleOpaque = True
                        
                        SetTileOptionsCaption lTileOption
                        
                    Else
                        
                        Sound "listend"
                        
                    End If
                
                Case vbKeyRight
                
                    If lTileOption < .pTile.UBound Then
                        
                        .pTile(lTileOption).TransparencyPct = 75
                        .pTile(lTileOption).BackStyleOpaque = False
                        
                        If lTileOption = 1 And bPinLocked = True Then
                        
                            lTileOption = 4
                        
                        Else
                        
                            lTileOption = lTileOption + 1
                        
                        End If
                        
                        Sound "select"
                        
                        .pTile(lTileOption).TransparencyPct = 0
                        .pTile(lTileOption).BackColor = eDashProp(lRowSelected).TileColor
                        .pTile(lTileOption).BackStyleOpaque = True
                        
                        SetTileOptionsCaption lTileOption
                        
                    Else
                        
                        Sound "listend"
                        
                    End If
                
                Case vbKeyBack
                    
                    CloseTileOptions
    
                Case vbKeyReturn
            
                    ProcessTileOption lTileOption
            
            End Select
            
        End If
            
    End With
    
    TileOptionsKey = True

End Function


Public Function SetTileOptionsCaption(Index As Long)

    ' on error resume next
    
    Select Case Index
    
        Case 0 'close app
        
            frmMain.lblTileOptions.Caption = "Power options"
        
        Case 1 'close app
        
            frmMain.lblTileOptions.Caption = "Lock screen"
        
        Case 2
        
            frmMain.lblTileOptions.Caption = "Delete tile"
        
        Case 3
        
            frmMain.lblTileOptions.Caption = "Edit tile"
        
        Case 4
        
            frmMain.lblTileOptions.Caption = "Add tile"
        
    End Select
        
End Function

Public Function ProcessTileOption(Index As Long) As Boolean

    ' on error resume next
    
    SetTileOptionsCaption lTileOption
    
    'CloseTileOptions
                            
    Dim lTemp As Long
                            
    With frmMain
                                
        Sound "yes"
            
        Select Case Index
        
            Case 0 'power
        
                BeginCloseExe
            
            Case 1 'lock
                    
                Do Until ShowLock = True
                    DoEvents
                Loop
                
                CloseTileOptions False
        
            Case 2 'delete tile
            
                lTemp = CLng(.cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent))
            
                Do Until DeleteTile(lRowSelected, .cRow(lRowSelected).TileCurrent, lTemp, False) = True
                    DoEvents
                Loop
                            
            Case 3 'edit tile
            
                Do Until InitPinSelect(True) = True
                    DoEvents
                Loop
            
            Case 4 'add tile
            
                Do Until InitPinSelect(False) = True
                    DoEvents
                Loop
                        
        End Select
        
    End With
    
    ProcessTileOption = True

End Function

Public Function InitPinSelect(bReplace As Boolean) As Boolean
    
    ' on error resume next
    
    Dim lTemp As Long
    
    With frmMain
    
        bPinReplace = bReplace
    
        .cPinItems.ClearTiles
        
        .cPinItems.TileColorSelected = eDashProp(lRowSelected).TileColor
        
        For lTemp = 0 To UBound(PinItems())
    
            .cPinItems.TileSet lTemp, , PinItems(lTemp).Title, PinItems(lTemp).Thumb, , CStr(PinItems(lTemp).Index)
    
        Next
    
        .cPinItems.TileSelected(0) = True
        
        .cRow(lRowSelected).Deselect
        
        Do Until CaptureScreen(.pTilePinSelectHolder.hDC) = True
            DoEvents
        Loop
        
        .pTilePinSelectHolder.Left = 0
        .pTilePinSelectHolder.ZOrder 0
        .pTilePinSelectHolder.Visible = True
        
    End With
    
    InitPinSelect = True

End Function

Public Function PinSelectKey(KeyCode As Integer, Shift As Integer) As Boolean

    ' on error resume next
    
    Dim sTemp() As String
    Dim ctl As Control

    With frmMain
    
        Select Case KeyCode
        
            Case vbKeyLeft
            
                .cPinItems.TileSelectLeft
                
                .lblTilePinOptions(1).Caption = PinItems(.cPinItems.TileCurrent).Title
            
            Case vbKeyRight
            
                .cPinItems.TileSelectRight
            
                .lblTilePinOptions(1).Caption = PinItems(.cPinItems.TileCurrent).Title
            
            Case vbKeyBack
            
                .pTilePinSelectHolder.Visible = False
            
            Case vbKeyReturn
            
                sTemp = Split(PinItems(.cPinItems.TileCurrent).App, "##")
            
                For Each Control In frmMain.Controls
                    DoEvents
                    If Control.Name = sTemp(0) Then
                        Set ctl = Control
                        Exit For
                    End If
                Next
                            
                ctl.Left = 0
                
                ctl.ZOrder 0
                
                ctl.Visible = True
                
                .pTileOptions.Visible = False
                .pTilePinSelectHolder.Visible = False
                
                If UBound(sTemp()) < 2 Then
                
                    ctl.SendMessage ("pin")
            
                Else
                
                    ctl.SendMessage (sTemp(2))
            
                End If
                
                CurrentApp.Name = sTemp(0)
            
                PreviousScreen = CurrentScreen
                CurrentScreen = RunningApp
            
                frmMain.pFocus.SetFocus
            
        End Select
        
    End With
    
    PinSelectKey = True

End Function

Public Function CloseTileOptions(Optional bSound As Boolean = True) As Boolean

    ' on error resume next
    
    With frmMain
        
        .pTileOptions.Visible = False
    
        If bSound = True Then
    
            Sound "openelement"
        
        End If
    
        Select Case CurrentScreen
            Case 0
                Do Until RestoreDash(True) = True
                    DoEvents
                Loop
        End Select
        
    End With
    
    CloseTileOptions = True

End Function

Public Function ShowTileOptions() As Boolean

    ' on error resume next
    
    Sound "closeelement"

    With frmMain
        
        Dim i As Integer
        
        bPinLocked = Tiles(.cRow(lRowSelected).TileTag(.cRow(lRowSelected).TileCurrent)).Locked
        
        If bPinLocked = False Then

            .pTile(0).TransparencyPct = 0
            .pTile(0).BackColor = eDashProp(lRowSelected).TileColor
            .pTile(0).BackStyleOpaque = True
            
            For i = 1 To .pTile.UBound
                .pTile(i).TransparencyPct = 50
                .pTile(i).BackStyleOpaque = False
            Next
            
            lTileOption = 0
        
        Else
            
            For i = 2 To 3
                .pTile(i).TransparencyPct = 90
                .pTile(i).BackStyleOpaque = False
            Next
            
            .pTile(0).TransparencyPct = 0
            .pTile(0).BackColor = eDashProp(lRowSelected).TileColor
            .pTile(0).BackStyleOpaque = True
            
            .pTile(1).TransparencyPct = 50
            .pTile(1).BackStyleOpaque = False
        
            .pTile(4).TransparencyPct = 50
            .pTile(4).BackStyleOpaque = False
        
            lTileOption = 0
        
        End If
        
        lPinRow = lRowSelected
        lPinTile = .cRow(lTileSelected).TileCurrent
        
        SetTileOptionsCaption lTileOption
        
        Do Until CaptureScreen(frmMain.pTileOptions.hDC) = True
            DoEvents
        Loop
        
        .pTileOptions.ZOrder 0
        .pTileOptions.Left = 0 '.pSideFade.Width
        .pTileOptions.Visible = True

    End With
    
    ShowTileOptions = True

End Function

Public Function DeleteTile(Row As Long, Tile As Long, TileIndex As Long, bHide As Boolean, Optional bSave As Boolean = True) As Boolean

    ' on error resume next
    
    Loading True

    With frmMain
    
        Dim sTemp As String
        Dim lTemp As Long
        Dim i As Long
        Dim lXMLIndex As Long
        Dim bEnd As Boolean
    
        Dim xmlRowNode As IXMLDOMNode
        Dim xmlnode As IXMLDOMNode
    
        If Tile < .cRow(Row).TileCount Then
            bEnd = False
        Else
            bEnd = True
        End If
        
        lXMLIndex = Tiles(TileIndex).Index
        
        Do Until .cRow(Row).TileDelete(Tile) = True
            DoEvents
        Loop
        
        If Row = lRowSelected Then
            If Tile = 0 Then
                .cRow(Row).TileSelected(0) = True
            ElseIf Tile >= .cRow(Row).TileCount Then
                .cRow(Row).TileSelected(.cRow(Row).TileCount) = True
            Else
                .cRow(Row).TileSelected(Tile) = True
            End If
        End If
        
        If bHide = True Then
            
            Set xmlnode = xmlSettingsDoc.selectSingleNode("//dashboard/dash[" & eDashProp(Row).Index & "]/item[" & lXMLIndex & "]").Attributes.getNamedItem("visible")
        
            xmlnode.Text = "no"
        
        Else
        
            Set xmlRowNode = xmlSettingsDoc.selectSingleNode("//dashboard/dash[" & eDashProp(Row).Index & "]")
        
            Set xmlnode = xmlRowNode.selectSingleNode("item[" & lXMLIndex & "]")
        
            xmlRowNode.removeChild xmlnode
            
            i = 0
            
            If bEnd = False Then
                For Each xmlnode In xmlRowNode.selectNodes("item")
                    If i >= lXMLIndex Then
                        xmlnode.Attributes.getNamedItem("index").Text = CStr(CLng(xmlnode.Attributes.getNamedItem("index").Text) - 1)
                    End If
                    i = i + 1
                Next
            End If
        
        End If
        
        If bSave = True Then
        
            xmlSettingsDoc.save App.Path & "\settings.xml"
        
        End If
    
        Do Until ScanTiles = True
            DoEvents
        Loop
        
        CloseTileOptions

    End With
    
    Loading False
    
    DeleteTile = True

End Function

Public Function AddTile(Row As Long, Tile As Long, TileIndex As Long, Caption As String, SubCaption As String, Thumb As String, Source As String, Optional SubSource As Long = 0, Optional bSave As Boolean = True) As Boolean

    On Error Resume Next
    
    Loading True

    Dim lUpper As Long

    Dim xmlnode As IXMLDOMNode
    Dim xmlNewEntry As IXMLDOMElement
    
    Dim xmlNodeNew As IXMLDOMNode
    Dim xmlNodeOld As IXMLDOMNode
    
    Dim lXMLIndex As Long

    Dim t As Long
    Dim c As Long
    
    With frmMain
    
        If TileIndex <> -1 Then
    
            lXMLIndex = Tiles(TileIndex).Index
            
        Else
        
            lXMLIndex = -1
            
        End If
    
        .cRow(Row).TileInsert Tile, Caption, SubCaption, Thumb
        
        Dim sTemp(4) As String
        
        sTemp(0) = "//dashboard/dash[" & eDashProp(Row).Index & "]"
        
        Set xmlnode = xmlSettingsDoc.selectSingleNode(sTemp(0))
        
        Set xmlNewEntry = xmlSettingsDoc.createElement("item")
        xmlnode.appendChild xmlNewEntry
        
        xmlNewEntry.setAttribute "index", ""
        xmlNewEntry.setAttribute "visible", ""
        xmlNewEntry.setAttribute "locked", ""
        xmlNewEntry.setAttribute "title", ""
        xmlNewEntry.setAttribute "subtitle", ""
        xmlNewEntry.setAttribute "showtitle", ""
        xmlNewEntry.setAttribute "thumb", ""
        xmlNewEntry.setAttribute "ispin", ""
        xmlNewEntry.setAttribute "source", ""
        
        lUpper = xmlnode.childNodes.Length - 2
        c = xmlnode.childNodes.Length - 2
        
        If lXMLIndex <> -1 Then
             
            For t = lXMLIndex To lUpper
                 
                sTemp(1) = "item[" & c & "]"
                sTemp(2) = "item[" & c - 1 & "]"
                
                sTemp(3) = sTemp(0) & "/" & sTemp(1)
                sTemp(4) = sTemp(0) & "/" & sTemp(2)
                
                Set xmlNodeOld = xmlSettingsDoc.selectSingleNode(sTemp(3))
                Set xmlNodeNew = xmlSettingsDoc.selectSingleNode(sTemp(4))
             
                xmlNodeOld.Attributes.getNamedItem("index").Text = CLng(xmlNodeNew.Attributes.getNamedItem("index").Text) + 1
                xmlNodeOld.Attributes.getNamedItem("visible").Text = xmlNodeNew.Attributes.getNamedItem("visible").Text
                xmlNodeOld.Attributes.getNamedItem("locked").Text = xmlNodeNew.Attributes.getNamedItem("locked").Text
                xmlNodeOld.Attributes.getNamedItem("title").Text = xmlNodeNew.Attributes.getNamedItem("title").Text
                xmlNodeOld.Attributes.getNamedItem("subtitle").Text = xmlNodeNew.Attributes.getNamedItem("subtitle").Text
                xmlNodeOld.Attributes.getNamedItem("showtitle").Text = xmlNodeNew.Attributes.getNamedItem("showtitle").Text
                xmlNodeOld.Attributes.getNamedItem("thumb").Text = xmlNodeNew.Attributes.getNamedItem("thumb").Text
                xmlNodeOld.Attributes.getNamedItem("ispin").Text = xmlNodeNew.Attributes.getNamedItem("ispin").Text
                xmlNodeOld.Attributes.getNamedItem("source").Text = xmlNodeNew.Attributes.getNamedItem("source").Text
            
                c = c - 1
             
             Next
             
            sTemp(1) = sTemp(0) & "/item[" & lXMLIndex & "]"
        
        Else
        
            sTemp(1) = sTemp(0) & "/item[" & lUpper & "]"
            
            lXMLIndex = lUpper
        
        End If
        
        Set xmlNodeOld = xmlSettingsDoc.selectSingleNode(sTemp(1))
        xmlNodeOld.Attributes.getNamedItem("index").Text = lXMLIndex
        xmlNodeOld.Attributes.getNamedItem("visible").Text = "yes"
        xmlNodeOld.Attributes.getNamedItem("locked").Text = "no"
        xmlNodeOld.Attributes.getNamedItem("title").Text = Caption
        xmlNodeOld.Attributes.getNamedItem("subtitle").Text = SubCaption
        If Caption <> vbNullString Then
            xmlNodeOld.Attributes.getNamedItem("showtitle").Text = "yes"
        Else
            xmlNodeOld.Attributes.getNamedItem("showtitle").Text = "no"
        End If
        xmlNodeOld.Attributes.getNamedItem("thumb").Text = Thumb
        xmlNodeOld.Attributes.getNamedItem("ispin").Text = "yes"
        xmlNodeOld.Attributes.getNamedItem("source").Text = Source
        
        If bSave = True Then
            xmlSettingsDoc.save App.Path & "\settings.xml"
        End If
    
        Do Until ScanTiles = True
            DoEvents
        Loop
    
    End With

    Loading False

    AddTile = True

End Function

Public Function ScanTiles() As Boolean
    
    ' on error resume next
    
    With frmMain
    
        Dim i As Long
        Dim r As Long
        Dim t As Long
        Dim sCaption As String
        Dim sSubCaption As String
        Dim sThumb As String
        Dim strSplit() As String
        
        Dim xmlDash As IXMLDOMNode
                     
        r = 0
        t = 0
        i = 0
        
        ReDim Tiles(t)
        
        For Each xmlDash In xmlSettingsDoc.selectNodes("//dashboard/dash")
        
            If xmlDash.Attributes.getNamedItem("visible").Text = "yes" Then
            
                r = 0
            
                For Each xmlSetting In xmlDash.selectNodes("item")
                    
                    DoEvents
                    
                    If xmlSetting.Attributes.getNamedItem("visible").Text = "yes" Then
                        
                        ReDim Preserve Tiles(t)
                        
                        'get tile locked
                            If xmlSetting.Attributes.getNamedItem("locked").Text = "yes" Then
                                Tiles(t).Locked = True
                            Else
                                Tiles(t).Locked = False
                            End If
                            
                        'get caption
                            If xmlSetting.Attributes.getNamedItem("showtitle").Text = "yes" Then
                                If xmlSetting.Attributes.getNamedItem("title").Text <> "" Then
                                    sCaption = xmlSetting.Attributes.getNamedItem("title").Text
                                Else
                                    sCaption = ""
                                End If
                            Else
                                sCaption = ""
                            End If
                    
                        'get subcaption
                            If xmlSetting.Attributes.getNamedItem("subtitle").Text <> "" Then
                                sSubCaption = xmlSetting.Attributes.getNamedItem("subtitle").Text
                            Else
                                sSubCaption = ""
                            End If
                        
                        'get thumb
                            If xmlSetting.Attributes.getNamedItem("thumb").Text <> "" Then
                                If Mid(xmlSetting.Attributes.getNamedItem("thumb").Text, 1, 1) = "/" Then
                                    sThumb = App.Path & xmlSetting.Attributes.getNamedItem("thumb").Text
                                Else
                                    sThumb = xmlSetting.Attributes.getNamedItem("thumb").Text
                                End If
                            Else
                                sThumb = ""
                            End If
                    
                        Tiles(t).Index = CLng(xmlSetting.Attributes.getNamedItem("index").Text)
                        Tiles(t).Title = sCaption
                        Tiles(t).Thumb = sThumb
                        Tiles(t).Source = xmlSetting.Attributes.getNamedItem("source").Text
        
                        .cRow(i).TileTag(r) = CStr(t)
                        
                        strSplit() = Split(Tiles(t).Source, "##")
                        
                        If UBound(strSplit()) > -1 Then
                                   
                            Select Case strSplit(0)
                            
                                Case "app"
                                
                                    Select Case strSplit(1)
                                    
                                        Case "cWeather"
                                        
                                            lWeatherTile(0) = i
                                            lWeatherTile(1) = r
                                    
                                        Case "cNews"
                                        
                                            lNewsTile(0) = i
                                            lNewsTile(1) = r
                                        
                                        Case "cSport"
                                        
                                            lSportTile(0) = i
                                            lSportTile(1) = r
                                                
                                        Case "cCalendar"
                                        
                                            lCalendarTile(0) = i
                                            lCalendarTile(1) = r
                                
                                    End Select
                                
                                Case "file", "ytvideo"
                                
                                    Tiles(t).SubSource = CLng(strSplit(4))
                                    
                                Case "live"
                                
                                    Select Case strSplit(1)
                                                                            
                                    End Select
                                
                            End Select
                            
                        End If
                        
                        r = r + 1
                    
                        t = t + 1
                    
                    End If
                    
                Next
                
                i = i + 1
                
            End If
    
        Next
        
    End With

    ScanTiles = True

End Function

Public Function ShowLock(Optional bHideHome As Boolean) As Boolean

    ' on error resume next
    
    With frmMain
    
        'Sound "closeelement"
    
        Dim i As Long
        
        bDashLocked = True
                
        If bHideHome <> True Then
            
            'For i = 0 To .cRow.UBound
            '    .cRow(i).Visible = False
            '    .lblRow(i).Visible = False
            '    If .cRow(i).Top > -.cRow(i).Height Or .cRow(i).Top < .ScaleHeight Then
            '        .tmrWait.Enabled = True
            '        Do Until .tmrWait.Enabled = False
            '            DoEvents
            '        Loop
            '    End If
            'Next
    
            Do Until ClearDash(True) = True
                DoEvents
            Loop
    
            .pTopBar.Visible = False
    
            .tmrWait.Enabled = True
            Do Until .tmrWait.Enabled = False
                DoEvents
            Loop
        
        End If
        
        If .lblLockWeather.Caption = "" Then
            .lblLockTime.Top = (.lblLockWeather.Top + .lblLockWeather.Height) - .lblLockTime.Height
        Else
            .lblLockTime.Top = .lblLockWeather.Top - .lblLockTime.Height
        End If
        
        .pLockHolder.ZOrder 0
        .pLockHolder.Visible = True
        .pFocus.Visible = False
        
        .lblLockTime.Visible = True
        '.pLockIcon.Visible = True
        If .cNotificationWidget.Visible = False Then
            '.lblLockNotifCount.Visible = True
            .pLockNotifGif.Animate lvicAniCmdStart
            '.pLockNotifGif.Visible = True
        End If
        .cNotificationWidget.ZOrder 0
        
        PreviousScreen = CurrentScreen
        CurrentScreen = LockScreen

        If frmMain.pLockArt.Visible = True Then
            
            SetLockBack frmMain.pLockArt.Picture, True, 25, 50
            
        Else
        
            'SetLockBack LoadPictureGDIplus(.pLockHolder.Tag), False
        
            SetBackground LoadPictureGDIplus(.pLockHolder.Tag), .pLockHolder, 0, 0
            
        End If
        
        .pLockHolder.SetFocus

    End With

    ShowLock = True

End Function

Public Function CloseLock() As Boolean

    ' on error resume next
    
    With frmMain
    
        Sound "openelement"
        
        Dim i As Long
        
        .lblLockTime.Visible = False
        .pLockIcon.Visible = False
        .lblLockNotifCount.Visible = False
        .pLockNotifGif.Visible = False
        .pLockNotifGif.Animate lvicAniCmdStop
        '.tmrWait.Enabled = True
        'Do Until .tmrWait.Enabled = False
        '    DoEvents
        'Loop
        
        .pFocus.Visible = True
        
        .pLockHolder.Visible = False
        Set .pLockHolder.Picture = Nothing
        
        .tmrWait.Enabled = True
        Do Until .tmrWait.Enabled = False
            DoEvents
        Loop
        
        .pTopBar.Visible = bShowTopBar
        
        .tmrWait.Enabled = True
        Do Until .tmrWait.Enabled = False
            DoEvents
        Loop
        
        'For i = 0 To .cRow.UBound
        '    .cRow(i).Visible = True
        '    .lblRow(i).Visible = True
        '    If .cRow(i).Top > -.cRow(i).Height Or .cRow(i).Top < .ScaleHeight Then
        '        .tmrWait.Enabled = True
        '        Do Until .tmrWait.Enabled = False
        '            DoEvents
        '        Loop
        '    End If
        'Next
        
        Do Until RestoreDash(True) = True
            DoEvents
        Loop
        
        .pFocus.ZOrder 0
        
        CurrentScreen = PreviousScreen
        PreviousScreen = LockScreen

        .pFocus.SetFocus

        ClosePin

        bDashLocked = False

    End With
    
    CloseLock = True
    
End Function

Public Function ReselectDash() As Boolean

    ' on error resume next
    
    With frmMain
    
        Select Case CurrentScreen
            Case 0 'home
                .cRow(lRowSelected).TileSelected(lMenuSelect) = True
                .lblRow(lRowSelected).ForeColor = cRowTitleSelected
            Case 2 'search
                .cKeyboard.GiveFocus
            Case 3 'search results
                .cSearch.TileSelected(.cSearch.TileCurrent) = True
        End Select
            
    End With
        
    ReselectDash = True

End Function
        
Public Function CloseSearchResults(Optional bAnimate As Boolean = False) As Boolean

    ' on error resume next
    
    CloseSearchResults = True

End Function

Public Function UpdateTopBar() As Boolean
   
    ' on error resume next

    If bShowTopBar = True Then
        
        With frmMain
                    
            If lRowSelected = 0 Then
            
                .pTopBar.Visible = True
                
            Else
            
                .pTopBar.Visible = False
            
            End If
            
            .pAppLogo.Visible = .pTopBar.Visible
            .lblAppTitle.Visible = False '.pTopBar.Visible
            .lblHomeTime.Visible = .pTopBar.Visible
                    
        End With
        
    End If
    
    UpdateTopBar = True

End Function

Public Function SetServiceIcon(Index As Integer, State As eServiceState) As Boolean

    With frmMain
    
        Select Case State
            
            Case 0
            
                .pServiceIcon(Index).BlendPct = 100
                .pServiceIcon(Index).GrayScale = lvicCCIR709
            
            Case 1
            
                .pServiceIcon(Index).BlendPct = 50
                .pServiceIcon(Index).GrayScale = lvicNoGrayScale
            
            Case 2
                
                .pServiceIcon(Index).BlendPct = 0
                .pServiceIcon(Index).GrayScale = lvicNoGrayScale
            
        End Select

    End With

End Function

Public Function SetBackground(pImage As GDIpImage, cTarget As Control, Optional lBlur As Long = 5, Optional lDarken As Long = 50) As Boolean

    Dim pEffects As New GDIpEffects
    Dim pAttribs As New GDIpEffects
    Dim lDestX As Long
    Dim lDestY As Long
    Dim lDestWidth As Long
    Dim lDestHeight As Long
    Dim lSourceX As Long
    Dim lSourceY As Long
    Dim lSourceWidth As Long
    Dim lSourceHeight As Long
    
    lDestX = 0
    lDestY = 0
    lDestWidth = cTarget.Width
    lDestHeight = cTarget.Height
    lSourceX = 0
    lSourceY = 0
    lSourceWidth = pImage.Width
    lSourceHeight = pImage.Height
    
    pEffects.CreateBlurEffect lBlur, False
    
    pAttribs.BlendColor = vbBlack
    pAttribs.BlendPct = lDarken
    
    pImage.Render cTarget.hDC, lDestX, lDestY, lDestWidth, lDestHeight, lSourceX, lSourceY, lSourceWidth, lSourceHeight, 0, pAttribs.AttributesHandle, , pEffects.EffectsHandle(lvicBlurFX)

    cTarget.Refresh

    SetBackground = True
    
End Function


Public Function SetLockBack(pImage As GDIpImage, Optional bStretch As Boolean = False, Optional lBlur As Long = 0, Optional lDarken As Long = 0) As Boolean

    Dim pEffects As New GDIpEffects
    Dim pAttribs As New GDIpEffects
    Dim lDestX As Long
    Dim lDestY As Long
    Dim lDestWidth As Long
    Dim lDestHeight As Long
    Dim lSourceX As Long
    Dim lSourceY As Long
    Dim lSourceWidth As Long
    Dim lSourceHeight As Long
    
    lDestX = 0
    lDestY = 0
    lDestWidth = frmMain.pLockHolder.Width
    lDestHeight = frmMain.pLockHolder.Height
    lSourceX = 0
    lSourceWidth = pImage.Width
    
    If bStretch = True Then
        lSourceY = pImage.Height * 0.21875
        lSourceHeight = pImage.Height * 0.5625
    Else
        lSourceY = pImage.Height
        lSourceHeight = pImage.Height
    End If
    
    pEffects.CreateBlurEffect lBlur, False
    
    pAttribs.BlendColor = vbBlack
    pAttribs.BlendPct = lDarken
    
    pImage.Render frmMain.pLockHolder.hDC, lDestX, lDestY, lDestWidth, lDestHeight, lSourceX, lSourceY, lSourceWidth, lSourceHeight, 0, pAttribs.AttributesHandle, , pEffects.EffectsHandle(lvicBlurFX)

    frmMain.pLockHolder.Refresh

    SetLockBack = True
    
End Function


Public Function CalculateScrollUp(lRow As Long, lNewRow As Long, lOldRow As Long) As Long

    ' on error resume next
    
    Dim t As Long
        
    With frmMain
        
        If lRow = lNewRow Then
        
            t = .lblRow(lRow).Top - (.pRowSpace.Top - .lblRow(lRow).Height)
            
        ElseIf lRow < lNewRow Then
        
            t = .lblRow(lRow).Top - (.pRowSpace.Top - (.lblRow(lRow).Height * ((lNewRow - lRow) + 1))) + (.lblRow(lRow).Height / 2)
        
        ElseIf lRow > lNewRow Then
        
            t = .lblRow(lRow).Top - (.cRow(lNewRow).Top + .cRow(lNewRow).Height + (.lblRow(lRow).Height * (lRow - lNewRow)))
            
        End If
        
    End With
    
    CalculateScrollUp = t

End Function

Public Function CalculateScrollDown(lRow As Long, lNewRow As Long, lOldRow As Long) As Long

    ' on error resume next
    
    Dim t As Long
    
    t = 0
        
    With frmMain
            
        If lRow = lNewRow Then
        
            t = (.pRowSpace.Top - .lblRow(lRow).Height) - .lblRow(lRow).Top
            
        ElseIf lRow < lNewRow Then
        
            t = (.pRowSpace.Top - (.lblRow(lRow).Height * ((lNewRow - lRow) + 1))) - .lblRow(lRow).Top - (.lblRow(lRow).Height / 2)
        
        ElseIf lRow > lNewRow Then
        
            t = (.cRow(lNewRow).Top + .cRow(lNewRow).Height + (.lblRow(lRow).Height * (lRow - lNewRow))) - .lblRow(lRow).Top
            
        End If
        
    End With
    
    CalculateScrollDown = t

End Function

Public Function LoadRow(Index As Long, Title As String) As Boolean

    ' on error resume next
        
    With frmMain
        
        If Index > .cRow.UBound Then
        
            ReDim Preserve RowLoaded(Index)
            ReDim Preserve lScroll(Index)
        
            Load .cRow(Index)
            Load .lblRow(Index)
        
            .cRow(Index).Top = .pRowSpace.Top
                        
            If Index < lRowSelected Then
            
                .lblRow(Index).Top = .lblRow(Index + 1).Top - .lblRow(Index).Height
                
            ElseIf Index = lRowSelected Then
            
                .lblRow(Index).Top = .pRowSpace.Top
                        
            ElseIf Index = lRowSelected + 1 Then
            
                .lblRow(Index).Top = .cRow(lRowSelected).Top + .cRow(lRowSelected).Height + (.lblRow(Index).Height / 2)
                
            ElseIf Index > lRowSelected + 1 Then
                
                .lblRow(Index).Top = .lblRow(Index - 1).Top + .lblRow(Index - 1).Height
                
            End If
            
            .cRow(Index).TileColorSelected = eDashProp(Index).TileColor
                    
            '.lblRow(Index).Left = .pRowSpace.Left + .pRowSpace.Width '(.pRowSpace.Width / 2)
            
            .lblRow(Index).ZOrder 0
            
            .lblRow(Index).Caption = Title
        
            .lblRow(Index).Visible = True
        
        Else
        
            .cRow(Index).TileColorSelected = eDashProp(Index).TileColor
        
            .cRow(Index).Top = .pRowSpace.Top
                        
            .lblRow(Index).Top = .cRow(Index).Top - .lblRow(Index).Height

            .cRow(Index).ClearTiles
        
            .lblRow(Index).Caption = Title
        
            .lblRow(Index).ZOrder 0
            
        End If
        
        RowLoaded(Index) = False
        
        .pMenuHolder.ZOrder 0
        
    End With
    
    LoadRow = True

End Function

Public Function LoadRowTiles(Index As Integer) As Boolean

    ' on error resume next
    
    'Loading True
    
    Dim i As Long
    Dim lTag As Long
    
    With frmMain
        
        For i = 0 To .cRow(Index).TileCount
            
            DoEvents
            
            lTag = CLng(.cRow(Index).TileTag(i))
            
            'Do Until .crow(Index).TileSet(i, , , , , CStr(lTag)) <> 0
            '    DoEvents
            'Loop
            
        Next
        
        RowLoaded(Index) = True

    End With

    LoadRowTiles = True

    'Loading False

End Function

Public Function SelectRow(lNewRow As Integer, Optional lNewTile As Long = 0) As Boolean

    On Error Resume Next
    
    Dim lOldRow As Long
    Dim r As Long
    
    lOldRow = lRowSelected
    
    With frmMain
        
        If lNewRow <> lOldRow Then
            
            If lNewRow >= 0 And lNewRow <= .cRow.UBound Then
            
                If lNewRow > 0 Then
                    .lblAppTitle.Visible = False
                End If
            
                If RowLoaded(lNewRow) = True Then
                
                    .cRow(lRowSelected).Deselect
                    
                    .cRow(lRowSelected).Visible = False
                    
                    .lblRow(lRowSelected).ForeColor = cRowTitleDeselected
                        
                    '.lblRow(lRowSelected).Left = .pRowSpace.Left + .pRowSpace.Width ' (.pRowSpace.Width / 2)
                        
                    lRowSelected = lNewRow
                    
                    If lRowSelected < lOldRow Then
                    
                        For r = 0 To .lblRow.UBound
                            DoEvents
                            lScroll(r) = CalculateScrollDown(r, CLng(lNewRow), lOldRow) / lScrollSpeed
                        Next
                    
                        .tmrScrollRowsDown.Enabled = True
                        
                        Do Until .tmrScrollRowsDown.Enabled = False
                            DoEvents
                        Loop
                    
                    ElseIf lRowSelected > lOldRow Then
                    
                        For r = 0 To .lblRow.UBound
                            DoEvents
                            lScroll(r) = CalculateScrollUp(r, CLng(lNewRow), lOldRow) / lScrollSpeed
                        Next
                    
                        .tmrScrollRowsUp.Enabled = True
                        
                        Do Until .tmrScrollRowsUp.Enabled = False
                            DoEvents
                        Loop
                
                    End If
                        
                    .lblRow(lRowSelected).ForeColor = cRowTitleSelected
                    
                    .lblRow(lRowSelected).Left = .pRowSpace.Left '+ (.pRowSpace.Width / 2)
                    
                    .cRow(lRowSelected).Visible = True
                        
                    .cRow(lRowSelected).TileSelected(lNewTile) = True
                        
                Else
                
                    Do Until LoadRowTiles(lNewRow) = True
                        DoEvents
                    Loop
                    
                    Do Until SelectRow(lNewRow, lNewTile) = True
                        DoEvents
                    Loop
                    
                End If
                
                If lRowSelected = 0 Then
                    .lblAppTitle.Visible = True
                End If
                
                .pSearchIcon.Visible = eDashProp(lRowSelected).Search
                
                Sound "select"
              
            End If
                
        Else
        
            .lblRow(lRowSelected).ForeColor = cRowTitleSelected
                    
            .cRow(lRowSelected).TileSelected(lNewTile) = True
            
            .cRow(lRowSelected).Visible = True
            
            .lblRow(lRowSelected).Visible = True
            
                        
        End If
          
    End With
    
    SelectRow = True
                    
End Function


Public Function ClearTiles() As Boolean

    With frmMain
    
        ' on error resume next
        
        Static i As Integer
    
        Do Until .cRow(0).ClearTiles = True
            DoEvents
        Loop
        
        .lblRow(0).Caption = ""
        
        If .cRow.UBound > 0 Then
            For i = 1 To .cRow.UBound
                Unload .cRow(i)
                Unload .lblRow(i)
            Next
        End If
    
    End With
    
    ClearTiles = True
    
End Function
