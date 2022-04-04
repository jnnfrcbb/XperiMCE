Attribute VB_Name = "mYouTube"

Public sYTQuality As String
Public bYTContinue As String

Public YTPlaying As Boolean

Public Function YtMakeCall(sCommand As String, sArgs As String)
    
    On Error Resume Next
    
    Dim sString As String
    
    Dim ret

    sString = "<invoke name=""" & sCommand & """ returntype=""xml""><arguments><string>""" & sArgs & """</string><number>0</number><string>" & sYTQuality & "</string></arguments></invoke>"

    ret = frmMain.swf(1).CallFunction(sString)
    
    bYTContinue = True

End Function


Public Function YtClose()
    
    On Error Resume Next
    
    Dim sString As String
    
    Dim Point As PointAPI
    
    GetCursorPos Point
    
    With frmMain
    
    If .swf.UBound > 0 Then
    
        YTStop
        
        .pVideoHolder.Visible = False
        
        GetCursorPos Point
        MouseX = Point.x
        MouseY = Point.y
        
        PlayBack.State = eWaiting
                
    End If
    
    End With

End Function

Public Function YTPause()

    On Error Resume Next
    
    Dim sString As String

    sString = "<invoke name=""pauseVideo"" returntype=""xml""><arguments></arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString
    
    PlayBack.State = ePaused
    
    YTPlaying = False
    
End Function

Public Function YTPlay()

    On Error Resume Next
    
    Dim sString As String

    sString = "<invoke name=""playVideo"" returntype=""xml""><arguments></arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString
    
    YTPlaying = True
    
    PlayBack.State = ePlaying
    
End Function

Public Function YTStop()

    On Error Resume Next
    
    Dim sString As String

    With frmMain

    If .swf.UBound > 0 Then
    
        bOSDAvailable = False
        
        .tmrPlayback.Enabled = False
    
        .tmrYT.Enabled = False
        
        .pVidPosition(2).Visible = False

        sString = "<invoke name=""stopVideo"" returntype=""xml""><arguments></arguments></invoke>"
        
        .swf(1).CallFunction sString
        
        sString = "<invoke name=""destroy"" returntype=""xml""><arguments></arguments></invoke>"
        
        .swf(1).CallFunction sString
        
        Unload .swf(1)
        
        YTPlaying = False
        
        PlayBack.State = eStopped
        
    End If
    
    End With

End Function

Public Function YTSkipTo(Position As Double)

    On Error Resume Next
    
    Dim sString As String

    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>" & Position & "</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString

End Function

Public Function YTSkipBackSmall()

    On Error Resume Next
    
    Dim lPos As Double
    Dim sString As String

    lPos = PlayBack.Position - 10
        
    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>" & lPos & "</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString

End Function

Public Function YTSkipForwardSmall()

    On Error Resume Next
    
    Dim sString As String

    Dim lPos As Double

    If PlayBack.Position > PlayBack.Duration - 10 Then
    
        lPos = PlayBack.Duration
        
    Else
    
        lPos = PlayBack.Duration + 10
        
    End If
    
    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>" & lPos & "</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString

End Function

Public Function YTSkipBackLarge()

    On Error Resume Next
    
    Dim lPos As Double
    Dim sString As String

    lPos = PlayBack.Position - 30
        
    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>" & lPos & "</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString

End Function

Public Function YTSkipForwardLarge()

    On Error Resume Next
    
    Dim sString As String

    Dim lPos As Double

    If PlayBack.Position > PlayBack.Duration - 30 Then
    
        lPos = PlayBack.Duration
        
    Else
    
        lPos = PlayBack.Duration + 30
        
    End If
    
    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>" & lPos & "</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).CallFunction sString

End Function

Public Function YTSkipToStart()

    On Error Resume Next
    
    Dim sString As String

    sString = "<invoke name=""seekTo"" returntype=""xml""><arguments><number>0</number></arguments><arguments>true</arguments></invoke>"
        
    frmMain.swf(1).play
        
    frmMain.swf(1).CallFunction sString

End Function


Public Function YTIncreaseQuality()

    On Error Resume Next
    
    Dim sString As String

    Dim bTemp As Boolean
    
    bTemp = False

    'pQuality.Visible = True
    
    'pQuality.ZOrder 0

    sString = "<invoke name=""getPlaybackQuality"" returntype=""xml""><arguments></arguments></invoke>"

    sYTQuality = frmMain.swf(1).CallFunction(sString)
    
    sYTQuality = Replace(sYTQuality, "<string>", "")
    
    sYTQuality = Replace(sYTQuality, "</string>", "")
    
    Select Case sYTQuality
    
        Case "small"
        
            sYTQuality = "medium"
            
            'lblQuality = "Quality" & vbCr & "Medium"
            
            bTemp = True
            
        Case "medium"
        
            sYTQuality = "large"
            
            'lblQuality = "Quality" & vbCr & "Large"
        
            bTemp = True
            
        Case "large"
        
            sYTQuality = "hd720"
            
            'lblQuality = "Quality" & vbCr & "720p"
        
            bTemp = True
            
        Case "hd720"
        
            sYTQuality = "hd1080"
            
            'lblQuality = "Quality" & vbCr & "1080p"
            
            bTemp = True
        
    End Select
    
    If bTemp = True Then
        
        YTPause
        
        sString = "<invoke name=""setPlaybackQuality"" returntype=""xml""><arguments><string>" & sYTQuality & "</string></arguments></invoke>"
        
        frmMain.swf(1).play
            
        frmMain.swf(1).CallFunction sString
        
        YTPlay
    
    End If
      
End Function

Public Function YTDecreaseQuality()

    On Error Resume Next
    
    Dim sString As String

    Dim bTemp As Boolean
    
    bTemp = False

    sString = "<invoke name=""getPlaybackQuality"" returntype=""xml""><arguments></arguments></invoke>"
    
    sYTQuality = frmMain.swf(1).CallFunction(sString)
    
    sYTQuality = Replace(sYTQuality, "<string>", "")
    
    sYTQuality = Replace(sYTQuality, "</string>", "")
    
    Select Case sYTQuality
    
        Case "small"
        
            'lblQuality = "Quality" & vbCr & "Small"
    
        Case "medium"
        
            sYTQuality = "small"
            
            'lblQuality = "Quality" & vbCr & "Small"
    
            bTemp = True
    
        Case "large"
        
            sYTQuality = "medium"
        
            'lblQuality = "Quality" & vbCr & "Medium"
        
            bTemp = True
    
        Case "hd720"
        
            sYTQuality = "large"
        
            'lblQuality = "Quality" & vbCr & "Large"
        
            bTemp = True
    
        Case "hd1080"
        
            sYTQuality = "hd720"
            
            'lblQuality = "Quality" & vbCr & "720p"
        
            bTemp = True
    
    End Select
    
    If bTemp = True Then
        
        YTPause
        
        sString = "<invoke name=""setPlaybackQuality"" returntype=""xml""><arguments><string>" & sYTQuality & "</string></arguments></invoke>"
            
        frmMain.swf(1).play
            
        frmMain.swf(1).CallFunction sString
        
        YTPlay
    
    End If
    
End Function

Public Function YTGetState() As Double

    On Error Resume Next

    Dim sString As String
    
    sString = "<invoke name=""getPlayerState"" returntype=""xml""><arguments></arguments></invoke>"

    YTGetState = GetNumber(frmMain.swf(1).CallFunction(sString))

End Function

Public Function YTGetBuffered() As Double

    On Error Resume Next

    Dim sString As String
    
    sString = "<invoke name=""getVideoLoadedFraction"" returntype=""xml""><arguments></arguments></invoke>"

    YTGetBuffered = GetNumber(frmMain.swf(1).CallFunction(sString))

End Function
Public Function GetNumber(sString As String) As Double

    On Error Resume Next
    
    Dim dNumber As Double
    
    Dim lStringStart As Long
    
    Dim lStringEnd As Long
    
    lStringStart = InStr(1, sString, "<number>") + Len("<number>")
    
    lStringEnd = InStr(1, sString, "</number>") - lStringStart
    
    dNumber = CDbl(Mid(sString, lStringStart, lStringEnd))
    
    GetNumber = dNumber

End Function

Public Function GetTimeString(dSeconds As Double) As String
    
    On Error Resume Next
    
    Dim lMinutes As Long
    Dim lSeconds As Long
    
    On Error Resume Next
    
    GetTimeString = dSeconds \ 60 & ":" & Format(dSeconds Mod 60, "00")

End Function

