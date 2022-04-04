Attribute VB_Name = "mLockPin"
Public sLockPin As String
Public sLockPinEntry As String

Public Function LockKey(KeyCode As Integer, Shift As Integer) As Boolean

    With frmMain
    
        If .lblLockPin(3).Visible = True Then
        
            Select Case KeyCode
            
                Case 48 To 57, 65 To 90
                
                    If Shift = 0 Then
                        
                        EnterPin LCase(Chr(KeyCode))
                    
                    Else
                
                        EnterPin UCase(Chr(KeyCode))
                        
                    End If
                    
                Case 96 To 105
                
                    EnterPin Chr(KeyCode - 48)
                
                Case vbKeyEscape
                
                    ClosePin
                
                Case vbKeyBack
                
                    DeletePin
                
            End Select
        
        Else
        
            If sLockPin <> "" Then
    
                BeginPin
        
            Else
            
                Select Case KeyCode
                
                    Case vbKeyBack
                    
                        .cNotificationWidget.Visible = Not .cNotificationWidget.Visible
                
                        If .cNotificationWidget.Visible = True Then
                        
                            .lblLockNotifCount.Visible = False
                        
                            .pLockNotifGif.Visible = False
                            
                            .pLockNotifGif.Animate lvicAniCmdStop
                        
                        Else
                        
                            .lblLockNotifCount.Caption = .cNotificationWidget.itemCount ' & "!"
                        
                            '.lblLockNotifCount.Visible = True
                            
                            .pLockNotifGif.Animate lvicAniCmdStart
                            
                            '.pLockNotifGif.Visible = True
                            
                        End If
                
                    Case vbKeyReturn
                    
                        CloseLock
                        
                End Select
            
            End If
            
        End If
        
    End With

    LockKey = True

End Function


Public Function BeginPin()

    Dim i As Integer
    
    frmMain.cNotificationWidget.Visible = False
    
    For i = 0 To 3
        
        DoEvents
    
        frmMain.lblLockPin(i).ZOrder 0
        frmMain.lblLockPin(i).Visible = True
        frmMain.pLockPin(i).Visible = True
        
    Next
    
End Function

Public Function ClosePin()

    Dim i As Integer
    
    For i = 0 To 3
        
        DoEvents
        
        frmMain.lblLockPin(i).Caption = ""
        frmMain.lblLockPin(i).ZOrder 0
        frmMain.lblLockPin(i).Visible = False
        frmMain.pLockPin(i).Visible = False
        
        ClearPin
        
    Next

    frmMain.cNotificationWidget.Visible = True
        
End Function

Public Function EnterPin(sNumber As String)

    Dim l As Long
    l = Len(sLockPinEntry)
    
    frmMain.lblLockPin(l).Caption = "*" 'sNumber

    sLockPinEntry = sLockPinEntry & sNumber
    
    If sLockPinEntry = sLockPin Then
    
        CloseLock
    
    End If

End Function

Public Function DeletePin()

    If Len(sLockPinEntry) > 0 Then

        Dim l As Long
        l = Len(sLockPinEntry)
        
        frmMain.lblLockPin(l - 1).Caption = ""

        sLockPinEntry = Mid(sLockPinEntry, 1, Len(sLockPinEntry) - 1)
    
    Else
    
        ClosePin
        
    End If

End Function


Public Function ClearPin()

    sLockPinEntry = ""

End Function
