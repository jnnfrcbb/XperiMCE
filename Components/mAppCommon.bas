Attribute VB_Name = "mAppCommon"

Public Function UpdateTopBar(cCtl As UserControl) As Boolean
   
    On Error Resume Next

    Dim pEffects As New GDIpEffects
    
    With cCtl
        
        If lRowSelected = 0 Then
        
            pEffects.CreateBlurEffect 15, False
        
            .pBack.Picture.Render .pTopBar.hDC, 0, 0, .pTopBar.Width, .pTopBar.Height, 0, 0, .pTopBar.Width, .pTopBar.Height, , , , pEffects.EffectsHandle(lvicBlurFX)
        
        Else
        
            .pTopBar.Picture = Nothing
        
        End If
        
        .pTopBar.Refresh
        
    End With
    
    UpdateTopBar = True

End Function

