Attribute VB_Name = "mNetflix"
Public Function LoadNetflix(sMovieID As String) As Boolean

    Dim sURL As String
    Dim strSplit() As String
    
    With frmMain
        
        sURL = "http://www.netflix.com/WiPlayer?movieid=[movieid]&trkid=[trkid]&tctx=[tctx]"
    
        strSplit() = Split(sMovieID, "++")
        
        sURL = Replace(sURL, "[movieid]", strSplit(0))
    
        sURL = Replace(sURL, "[trkid]", strSplit(1))
        
        sURL = Replace(sURL, "[tctx]", strSplit(2))
        
        bOSDAvailable = False
                
        If bMinVideo = True Then
        
            .pVideoHolder.Move .pMinVideo.Left, .pMinVideo.Top, .pMinVideo.Width, .pMinVideo.Height
                
            .pVideoHolder.BorderStyle = 1
                
        Else
        
            .pVideoHolder.Move 0, 0, .ScaleWidth, .ScaleHeight
                
            .pVideoHolder.BorderStyle = 0
            
        End If
        
        .pVideoHolder.ZOrder 0
        
        Load .wPlayer(1)
        
        .wPlayer(1).Silent = True
        .wPlayer(1).fullScreen = True
        .wPlayer(1).TheaterMode = True
        
        .wPlayer(1).Navigate App.Path & "\blank.htm"
                
        .wPlayer(1).Move 0, 0, .pVideoHolder.Width, .pVideoHolder.Height
            
        .pVideoHolder.Visible = True
            
        .wPlayer(1).Visible = True
            
        .wPlayer(1).Navigate sURL
        
        PlayBack.State = ePlaying
        
    End With

    LoadNetflix = True
        
End Function

Public Function PreCloseNetflix() As Boolean

    With frmMain
        
        .wPlayer(1).Visible = False
        .pVideoHolder.Visible = False
    
        PlayBack.State = eWaiting
        
        .pFocus.SetFocus
        
    End With
    
    PreCloseNetflix = True
    
End Function

Public Function CloseNetflix() As Boolean

    Sound "closeelement"

    With frmMain
    
        If .wPlayer.UBound > 0 Then
            Unload .wPlayer(1)
        End If
        
        .pFocus.SetFocus
        
        .pVideoHolder.Visible = False
        
    End With
    
    CloseNetflix = True

End Function
