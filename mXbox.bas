Attribute VB_Name = "mXbox"
Public Type xboxC
    ButtonA As Long
    ButtonB As Long
    ButtonX As Long
    ButtonY As Long
    ButtonLBumper As Long
    ButtonRBumper As Long
    ButtonLThumb As Long
    ButtonRThumb As Long
    ButtonBack As Long
    ButtonStart As Long
    DPadUp As Long
    DPadDown As Long
    DPadLeft As Long
    DPadRight As Long
    TriggerLeft As Long
    TriggerRight As Long
    StickL_X As Long
    StickL_Y As Long
    StickR_X As Long
    StickR_Y As Long
End Type

Public lControllerCount As Long
Public lXboxCount As Long
    
Public Function xBoxControllerRumble(lController As Long, LeftRumblePC As Long, RightRumblePC As Long)

    On Error Resume Next
    
    Dim lTemp As Long
    

    GPad_Rumble lController, (65535 / 100) * LeftRumblePC, (65535 / 100) * RightRumblePC

    

End Function

Public Function xboxController(lController As Long) As xboxC

    On Error Resume Next
    
    Dim cTemp As xboxC

    cTemp.ButtonA = GPad_ButtonA(lController)
    cTemp.ButtonB = GPad_ButtonB(lController)
    cTemp.ButtonX = GPad_ButtonX(lController)
    cTemp.ButtonY = GPad_ButtonY(lController)
    cTemp.ButtonLBumper = GPad_ButtonLBumper(lController)
    cTemp.ButtonRBumper = GPad_ButtonRBumper(lController)
    cTemp.ButtonLThumb = GPad_ButtonLThumb(lController)
    cTemp.ButtonRThumb = GPad_ButtonRThumb(lController)
    cTemp.ButtonBack = GPad_ButtonBack(lController)
    cTemp.ButtonStart = GPad_ButtonStart(lController)
    cTemp.DPadUp = GPad_ButtonDPadUp(lController)
    cTemp.DPadDown = GPad_ButtonDPadDown(lController)
    cTemp.DPadLeft = GPad_ButtonDPadLeft(lController)
    cTemp.DPadRight = GPad_ButtonDPadRight(lController)
    cTemp.TriggerLeft = GPad_ButtonLeftTrigger(lController)
    cTemp.TriggerRight = GPad_ButtonRightTrigger(lController)
    cTemp.StickL_X = GPad_LStickX(lController)
    cTemp.StickL_Y = GPad_LStickY(lController)
    cTemp.StickR_X = GPad_RStickX(lController)
    cTemp.StickR_Y = GPad_RStickY(lController)
    
    xboxController = cTemp

End Function

Public Function CountControllers() As Long

    On Error Resume Next
    
    Dim lCount As Long
    
    lCount = 0

    If GPad_Poll(0) = 1 Then

        lCount = lCount + 1
        
    End If
    
    'If GPad_Poll(1) = 1 Then
    
    '    lCount = lCount + 1
        
    'End If
    
    'If GPad_Poll(2) = 1 Then
    
    '    lCount = lCount + 1
        
    'End If
    
    'If GPad_Poll(3) = 1 Then
    
    '    lCount = lCount + 1
        
    'End If
    
    CountControllers = lCount

End Function



