Attribute VB_Name = "mSettings"
Option Explicit

Public xmlSettingsDoc As New DOMDocument
Public xmlSettings As IXMLDOMNodeList
Public xmlSetting As IXMLDOMNode

Public Type tUser
    Name As String
    Type As String
    Cover As String
    Photo As String
    Auth As String
End Type
Public User As tUser

Public sBackground As String

Public bSettingFullScreen As Boolean

Public Function LoadSettings() As Boolean

    On Error Resume Next

    Dim strArray() As String
    Static i As Integer

    ReDim eDashProp(0)

    With frmMain
    
        'xmlSettingsDoc.preserveWhiteSpace = True

        xmlSettingsDoc.Load App.Path & "\settings.xml"
        
        xmlSettingsDoc.save App.Path & "\settings_backup.xml"
    
        'window
        
            'fullscreen
        
                If xmlSettingsDoc.selectSingleNode("//window").Attributes.getNamedItem("fullscreen").Text = "yes" Then
                
                    frmSplash.WindowState = vbMaximized
                
                    frmMain.WindowState = vbMaximized
                
                    bSettingFullScreen = True
                    
                Else
                
                    frmSplash.WindowState = vbNormal
                
                    frmMain.WindowState = vbNormal
                    
                    bSettingFullScreen = False
                    
                End If
                
        'process
        
            'appcommand
            
                 If xmlSettingsDoc.selectSingleNode("//process").Attributes.getNamedItem("appcommand").Text = "yes" Then
                    
                    bRemoteSupport = True
                
                Else
                
                    bRemoteSupport = False
                
                End If
                
            'gamecontroller
            
                 If xmlSettingsDoc.selectSingleNode("//process").Attributes.getNamedItem("gamecontroller").Text = "yes" Then
                    
                    bGameControllerSupport = True
                
                Else
                
                    bGameControllerSupport = False
                
                End If
        
            'mouseout
            
                 If xmlSettingsDoc.selectSingleNode("//process").Attributes.getNamedItem("mouseout").Text = "yes" Then
                    
                    bMouseOut = True
                
                Else
                
                    bMouseOut = False
                
                End If
            
            'mousecharm
            
                 If xmlSettingsDoc.selectSingleNode("//process").Attributes.getNamedItem("mousecharm").Text = "yes" Then
                    
                    bMouseCharm = True
                
                Else
                
                    bMouseCharm = False
                
                End If
                
            'sounds
                       
                 If xmlSettingsDoc.selectSingleNode("//process").Attributes.getNamedItem("sounds").Text = "yes" Then
                    
                    bSounds = True
                
                Else
                
                    bSounds = False
                
                End If
                     
        'dash
        
            'top bar
                
                If xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("topbar").Text = "yes" Then
                    bShowTopBar = True
                Else
                    bShowTopBar = False
                    .pMenuHolder.Top = 0
                    .pMenuHolder.Height = .ScaleHeight
                End If
                
                strArray() = Split(xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("highlightcolor").Text, ",")
                .pTopBar.BackColor = RGB(CInt(strArray(0)), CInt(strArray(1)), CInt(strArray(2)))
                For i = 0 To .lblMenu.UBound
                    .lblMenu(i).BackStyle = 1
                    .lblMenu(i).BackColor = .pTopBar.BackColor
                    .lblMenu(1).BackStyle = 0
                Next
                .cKeyboard.FunctionKeySelectedBackColor = .pTopBar.BackColor
                .cKeyboard.MainKeySelectedBackColor = .pTopBar.BackColor
                        
            'background
            
                strArray() = Split(xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("backcolor").Text, ",")
                frmSplash.BackColor = RGB(CInt(strArray(0)), CInt(strArray(1)), CInt(strArray(2)))
                .BackColor = frmSplash.BackColor
                .pFocus.BackColor = frmSplash.BackColor
                .pLockHolder.BackColor = frmSplash.BackColor
                
                If xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("background").Text = "yes" Then
                    sBackground = xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("image").Text
                Else
                    sBackground = vbNullString
                End If
                
                If xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("weatherback").Text = "yes" Then
                    bWeatherBack = True
                Else
                    bWeatherBack = False
                End If
                
                If xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("universalback").Text = "yes" Then
                    bUniversalBack = True
                Else
                    bUniversalBack = False
                End If
                
            'menu
        
                If xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("menuleft").Text = "yes" Then
                    bMenuLeft = True
                    CentralMessage "menulefton"
                Else
                    bMenuLeft = False
                    CentralMessage "menuleftoff"
                End If

           'pin
           
               sLockPin = xmlSettingsDoc.selectSingleNode("//dashboard").Attributes.getNamedItem("pin").Text
              
            'properties
            
                i = 0
                
                lVideoRow = -1
                lMusicRow = -1
                lGamesRow = -1
            
                For Each xmlSetting In xmlSettingsDoc.selectNodes("//dashboard/dash")
                
                    If xmlSetting.Attributes.getNamedItem("visible").Text = "yes" Then
                    
                        ReDim Preserve eDashProp(i)
                    
                        'dash index
                            eDashProp(i).Index = CLng(xmlSetting.Attributes.getNamedItem("index").Text)
                    
                        'dash title
                            eDashProp(i).Title = xmlSetting.Attributes.getNamedItem("title").Text
                        
                            Select Case eDashProp(i).Title
                                Case "Music"
                                    lMusicRow = i
                                Case "Videos"
                                    lVideoRow = i
                                Case "Games"
                                    lGamesRow = i
                            End Select
                        
                        'fullmain
                            If xmlSetting.Attributes.getNamedItem("fullmain").Text = "yes" Then
                                eDashProp(i).FullMain = True
                            Else
                                eDashProp(i).FullMain = False
                            End If
                        
                        'tileborder
                            If xmlSetting.Attributes.getNamedItem("tileborder").Text = "yes" Then
                                eDashProp(i).TileBorder = True
                            Else
                                eDashProp(i).TileBorder = False
                            End If
                        
                        'tile color
                        
                            strArray() = Split(xmlSetting.Attributes.getNamedItem("tilecolor").Text, ",")
                            
                            eDashProp(i).TileColor = RGB(CInt(strArray(0)), CInt(strArray(1)), CInt(strArray(2)))
                            
                            eDashProp(i).TileColorRGB = xmlSetting.Attributes.getNamedItem("tilecolor").Text
                        
                        'saerch
                            If xmlSetting.Attributes.getNamedItem("search").Text = "yes" Then
                                eDashProp(i).Search = True
                            Else
                                eDashProp(i).Search = False
                            End If
                            
                        i = i + 1
                        
                    End If
                    
                Next
        
            'alarm
            
                sAlarmSource = xmlSettingsDoc.selectSingleNode("//alarm").Attributes.getNamedItem("source").Text
                sAlarmThumb = xmlSettingsDoc.selectSingleNode("//alarm").Attributes.getNamedItem("thumb").Text
        
            'apps
        
                'i = 0
                
                'ReDim AppList(0)
                'ReDim AppsRunning(0)
            
                'For Each xmlSetting In xmlSettingsDoc.selectNodes("//apps/app")
                
                '    ReDim Preserve AppList(i)
                '    AppList(i).Name = xmlSetting.Attributes.getNamedItem("name").Text
                '    AppList(i).Caption = xmlSetting.Attributes.getNamedItem("caption").Text
                '    AppList(i).Title = xmlSetting.Attributes.getNamedItem("title").Text
                    
                '    If Mid(xmlSetting.Attributes.getNamedItem("thumb").Text, 1, 1) = "/" Then
                '        AppList(i).Thumb = App.Path & xmlSetting.Attributes.getNamedItem("thumb").Text
                '    Else
                '        AppList(i).Thumb = xmlSetting.Attributes.getNamedItem("thumb").Text
                '    End If
                    
                '    i = i + 1
                
                'Next
                
            'pin items
            
                i = 0
            
                For Each xmlSetting In xmlSettingsDoc.selectNodes("//pinitems/item")
                
                    ReDim Preserve PinItems(i)
                    
                    PinItems(i).Index = CLng(xmlSetting.Attributes.getNamedItem("index").Text)
                    PinItems(i).Title = xmlSetting.Attributes.getNamedItem("title").Text
                    If Mid(xmlSetting.Attributes.getNamedItem("thumb").Text, 1, 1) = "/" Then
                        PinItems(i).Thumb = App.Path & xmlSetting.Attributes.getNamedItem("thumb").Text
                    Else
                        PinItems(i).Thumb = xmlSetting.Attributes.getNamedItem("thumb").Text
                    End If
                    PinItems(i).App = xmlSetting.Attributes.getNamedItem("app").Text
                
                    i = i + 1
                
                Next
                
        'music
        
            'speakers
            
                    If xmlSettingsDoc.selectSingleNode("//music/surround").Text = "surround" Then
                        MusicSettings.Speakers = 1
                    Else
                        MusicSettings.Speakers = 0
                    End If
                    
                'volume
            
                    MusicSettings.Volume = CDbl(xmlSettingsDoc.selectSingleNode("//music/playbackVolume").Text)
                    
                'eq
                    
                    MusicSettings.eqCount = CLng(xmlSettingsDoc.selectSingleNode("//music/eq").Attributes.getNamedItem("count").Text)
                    ReDim MusicSettings.eqHz(MusicSettings.eqCount)
                    ReDim MusicSettings.eqValue(MusicSettings.eqCount)
                    
                    i = 0
                    For Each xmlSetting In xmlSettingsDoc.selectNodes("//music/eq/setting")
                        MusicSettings.eqHz(i) = xmlSetting.Attributes.getNamedItem("Hz").Text
                        MusicSettings.eqValue(i) = xmlSetting.Attributes.getNamedItem("value").Text
                        .scrEQ(i).Value = MusicSettings.eqValue(i)
                        .scrEQ(i).Tag = MusicSettings.eqHz(i)
                        .lblEQValue(i).Caption = MusicSettings.eqValue(i)
    
                        i = i + 1
                    Next
                                        
        'video
            
            'root folders
            
                .cTV.SendMessage "root##" & xmlSettingsDoc.selectSingleNode("//video/tv").Attributes.getNamedItem("folder").Text
                
                .cFilms.SendMessage "root##" & xmlSettingsDoc.selectSingleNode("//video/film").Attributes.getNamedItem("folder").Text
            
            'player
            
                .wmp.uiMode = "none"
                
            'controls
            
            'menu
            
        'user
        
            'local
            
                Dim sUser(3) As String
                
                sUser(0) = xmlSettingsDoc.selectSingleNode("//user").Attributes.getNamedItem("name").Text
                sUser(1) = xmlSettingsDoc.selectSingleNode("//user").Attributes.getNamedItem("type").Text
                sUser(2) = xmlSettingsDoc.selectSingleNode("//user").Attributes.getNamedItem("cover").Text
                sUser(3) = xmlSettingsDoc.selectSingleNode("//user").Attributes.getNamedItem("photo").Text
                
                Do Until UpdateUser(sUser(0), sUser(1), sUser(2), sUser(3)) = True
                    DoEvents
                Loop
                
            'google
            
                If xmlSettingsDoc.selectSingleNode("//user/google").Attributes.getNamedItem("login").Text = "yes" Then
                    
                    bGoogleEnabled = True
                
                    sGoogleUser = xmlSettingsDoc.selectSingleNode("//user/google").Attributes.getNamedItem("user").Text
        
                    sGooglePass = xmlSettingsDoc.selectSingleNode("//user/google").Attributes.getNamedItem("pass").Text
                    
                    refreshToken = xmlSettingsDoc.selectSingleNode("//user/google").Attributes.getNamedItem("refreshToken").Text
    
                Else
                    
                    bGoogleEnabled = False
                    
                End If
                
            'lastfm
            
                If xmlSettingsDoc.selectSingleNode("//user/lastfm").Attributes.getNamedItem("login").Text = "yes" Then
                    
                    bLastFMEnabled = True
                    
                    sLastFMUser = xmlSettingsDoc.selectSingleNode("//user/lastfm").Attributes.getNamedItem("user").Text
        
                    sLastFMPass = xmlSettingsDoc.selectSingleNode("//user/lastfm").Attributes.getNamedItem("pass").Text
        
                Else
    
                    bLastFMEnabled = False
                
                End If
                
    End With
    
    LoadSettings = True

End Function

Public Function UpdateUser(sName As String, sType As String, sCover As String, sPhoto As String, Optional sAuth = "") As Boolean

    Dim sMessage As String
    Dim pPhoto As GDIpImage
    Dim fso As New FileSystemObject

    User.Name = sName
    User.Type = sType
    User.Cover = sCover
    User.Photo = sPhoto
    User.Auth = sAuth
    
    With frmMain
    
        .lblMenuHeader(0).Caption = User.Name
        .lblMenuHeader(1).Caption = User.Type  'User
    
        sMessage = "userloggedin##" & User.Auth & "##" & User.Name & "##" & User.Type

        If fso.FileExists(User.Photo) Then
'            Set pPhoto = LoadPictureGDIplus(User.Photo)
'           CentralMessage sMessage, pPhoto
'           .pUser.Picture = pPhoto
'           .lblMenuHeader(0).Left = (.pUser.Left * 2) + .pUser.Width
'           .lblMenuHeader(1).Left = .lblMenuHeader(0).Left
        End If

        
        If fso.FileExists(User.Cover) Then
'            Set pPhoto = LoadPictureGDIplus(User.Cover)
'            CentralMessage "usercover", pPhoto

'           .pUserCover.FastRedraw = False
'           .pUserCover.Picture = pPhoto
'           .pUserCover.Effects.CreateBlurEffect 25, False
'           .pUserCover.FastRedraw = True
        End If
        

    End With

    UpdateUser = True

End Function
