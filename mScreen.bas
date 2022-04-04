Attribute VB_Name = "mScreen"
'Declares
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
        
Public Declare Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long

Public Const ES_DISPLAY_REQUIRED As Long = &H2
Public Const ES_CONTINUOUS As Long = &H80000000

        
 
Public Function CaptureScreen(iTargetDC As Long) As Boolean
    Dim w As Long
    Dim h As Long
    
    'Get screen size
    w = CLng(Screen.Width / Screen.TwipsPerPixelX)
    h = CLng(Screen.Height / Screen.TwipsPerPixelY)
    
    'Capture screen
    
    'Dim pImage As GDIpImage
    
    BitBlt iTargetDC, 0, 0, w, h, GetWindowDC(frmMain.hwnd), 0, 0, vbSrcCopy
    
    BitBlt iTargetDC, 0, 0, w, h, GetWindowDC(frmMain.hwnd), 0, 0, vbSrcCopy
    
    'Set pImage = LoadPictureGDIplus(frmMain.pPlaylistHolder.Image)
    
    'Dim pEffects As New GDIpEffects
    'Dim pAttribs As New GDIpEffects
    'Dim lDestX As Long
    'Dim lDestY As Long
    'Dim lDestWidth As Long
    'Dim lDestHeight As Long
    'Dim lSourceX As Long
    'Dim lSourceY As Long
    'Dim lSourceWidth As Long
    'Dim lSourceHeight As Long
   '
   ' lDestX = 0
   ' lDestY = 0
   ' lDestWidth = frmMain.pPlaylistHolder.Width
   ' lDestHeight = frmMain.pPlaylistHolder.Height
   ' lSourceX = 0
   ' lSourceY = pImage.Height * 0.21875
   ' lSourceWidth = pImage.Width
   ' lSourceHeight = pImage.Height * 0.5625
   '
   ' pEffects.CreateBlurEffect 5, False
   '
   ' pAttribs.BlendColor = vbBlack
   ' pAttribs.BlendPct = 100

   ' pImage.Render frmMain.pPlaylistHolder.hDC, lDestX, lDestY, lDestWidth, lDestHeight, lSourceX, lSourceY, lSourceWidth, lSourceHeight, 0, pAttribs.AttributesHandle, , pEffects.EffectsHandle(lvicBlurFX)

    'frmMain.pPlaylistHolder.Refresh
    
    CaptureScreen = True
    
End Function


Public Function PreventMonitorSleeping()

    SetThreadExecutionState (ES_DISPLAY_REQUIRED Or ES_CONTINUOUS)

End Function

Public Function AllowMonitorSleeping()

    SetThreadExecutionState (ES_CONTINUOUS)

End Function
