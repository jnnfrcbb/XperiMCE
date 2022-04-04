Attribute VB_Name = "mApp"
Public Type eApp
    Name As String
    Thumb As String
    Title As String
    Caption As String
End Type
Public AppsRunning() As eApp
Public CurrentApp As eApp
Public AppList() As eApp

Public Function OpenApp(cApp As Control, Optional sCommand As String, Optional Thumb As String = "", Optional Caption As String = "") As Boolean

    On Error Resume Next

    Dim i As Long
    Dim lApp As Long
    
    'Sound "openapp"
    
    Do Until ClearDash(True) = True
        DoEvents
    Loop
    
    If bUniversalBack = True Then
        'cApp.SendMessage "backgroundchanged", frmMain.pBack.Picture
    End If
    
    cApp.Left = 0

    cApp.ZOrder 0
    
    cApp.Visible = True

    CurrentApp.Name = cApp.Name
    
    'If AppRunning(cApp.Name) = False Then
        
    '    For i = 0 To UBound(AppList())
    '        If cApp.Name = AppList(i).Name Then
    '            lApp = i
    '            Exit For
    '        End If
    '    Next
        
    '    AppListInsert 0, cApp.Name, AppList(lApp).Title, AppList(lApp).Thumb, AppList(lApp).Caption
    
        If sCommand <> "" Then
    
            Do Until cApp.SendMessage(sCommand) = True
                DoEvents
            Loop
    
        Else
    
            Do Until cApp.SendMessage("initcontrol##") = True
                DoEvents
            Loop
            
        End If
        
    'Else
    
    '    For i = 0 To UBound(AppsRunning())
    '        If AppsRunning(i).Name = cApp.Name Then
    '            Exit For
    '        End If
    '    Next
    
    '    AppListMove i, 0
        
    'End If
        
    PreviousScreen = CurrentScreen
    CurrentScreen = RunningApp
    
    frmMain.pFocus.SetFocus

    OpenApp = True

End Function

Public Function CloseApp(cApp As Control, Optional bForceClose As Boolean, Optional sCommand As String) As Boolean

    On Error Resume Next
    
    If bAppDirect = True Then
    
        Unload frmMain
        
        End
    
    Else
        
        Dim i As Long
        
        'remove app from running list
            'for i = 0 To UBound(AppsRunning())
            '    If LCase(AppsRunning(i).Name) = LCase(cApp.Name) Then
            '        AppListDelete i
            '        Exit For
            '    End If
            '    'AppListDelete i
            'Next
    
        'move app off screen and hide
            cApp.Left = frmMain.ScaleWidth
            cApp.ZOrder 1
            cApp.Visible = False
        
        'restore dash
            Do Until RestoreDash(True) = True
                DoEvents
            Loop
    
        'update key process
            PreviousScreen = CurrentScreen
            CurrentScreen = Home
        
        'ensure focus
            frmMain.pFocus.SetFocus
        
        If bForceClose = True Then
            If sCommand <> "" Then
                cApp.SendMessage ("closecontrol##" & sCommand)
            Else
                cApp.SendMessage ("closecontrol")
            End If
        End If
        
    End If
    
    CloseApp = True

End Function

Public Function MinimizeApp(cApp As Control) As Boolean

    On Error Resume Next
    
    cApp.Left = frmMain.ScaleWidth
    
    cApp.ZOrder 1
    
    cApp.Visible = False
    
    Do Until RestoreDash(True) = True
        DoEvents
    Loop

    PreviousScreen = CurrentScreen
    CurrentScreen = Home
    
    frmMain.pFocus.SetFocus
    
    MinimizeApp = True

End Function

Public Function AppRunning(AppName As String) As Boolean

    On Error Resume Next
    
    Dim bTemp As Boolean
    Dim i As Long
    
    bTemp = False
    
    For i = 0 To UBound(AppsRunning())
        DoEvents
        If LCase(AppsRunning(i).Name) = LCase(AppName) Then
            bTemp = True
            Exit For
        End If
    Next
        
    AppRunning = bTemp

End Function

Public Function AppListInsert(Index As Long, Name As String, Title As String, Thumb As String, Caption As String) As Long

    On Error Resume Next
    
    Dim i As Long
    Dim c As Long
    Dim r As Long
    
    On Error GoTo errHandle
    
    r = 1
    
    If AppsRunning(Index).Name <> "" Then
        
        c = UBound(AppsRunning()) + 1
    
        ReDim Preserve AppsRunning(c)
    
        For i = Index + 1 To UBound(AppsRunning())
        
            AppsRunning(c).Name = AppsRunning(c - 1).Name
            AppsRunning(c).Title = AppsRunning(c - 1).Title
            AppsRunning(c).Thumb = AppsRunning(c - 1).Thumb
            AppsRunning(c).Caption = AppsRunning(c - 1).Caption
            
            c = c - 1
        
        Next
        
    End If
    
    AppsRunning(Index).Name = Name
    AppsRunning(Index).Title = Title
    AppsRunning(Index).Thumb = Thumb
    AppsRunning(Index).Caption = Caption
    
    AppListInsert = r
    
errHandle:
    
    If Err.Number <> 0 Then
    
        r = -1
        
        Err.Clear
        
        Resume Next

    End If

End Function

Public Function AppListDelete(Index As Long) As Long

    On Error Resume Next
    
    Dim i As Long
    Dim r As Long
    
    On Error GoTo errHandle
    
    r = 1
    
    If UBound(AppsRunning()) > 0 Then
        
        For i = Index To UBound(AppsRunning()) - 1
        
            AppsRunning(i).Name = AppsRunning(i + 1).Name
            AppsRunning(i).Title = AppsRunning(i + 1).Title
            AppsRunning(i).Thumb = AppsRunning(i + 1).Thumb
            AppsRunning(i).Caption = AppsRunning(i + 1).Caption
        
        Next
    
        ReDim Preserve AppsRunning(UBound(AppsRunning()) - 1)
        
    Else
    
        AppListClear
        
        r = -1
        
    End If

    AppListDelete = r

errHandle:

    If Err.Number <> 0 Then
    
        r = -1
        
        Err.Clear
        
        Resume Next

    End If

End Function

Public Function AppListMove(Index As Long, NewIndex As Long) As Long

    On Error Resume Next
    
    Dim i As Long
    Dim sTemp(4) As String
    Dim r As Long
    
    On Error GoTo errHandle
    
    r = 1
    
    If Index <= UBound(AppsRunning()) And NewIndex < UBound(AppsRunning()) Then
        
        sTemp(0) = AppsRunning(Index).Name
        sTemp(1) = AppsRunning(Index).Title
        sTemp(2) = AppsRunning(Index).Thumb
        sTemp(3) = AppsRunning(Index).Caption
        
        AppListDelete Index
        
        AppListInsert NewIndex, sTemp(0), sTemp(1), sTemp(2), sTemp(3)
        
    Else
    
        r = -1

    End If

    AppListMove = r

errHandle:

    If Err.Number <> 0 Then
    
        r = -1
        
        Err.Clear
        
        Resume Next

    End If

End Function

Public Function AppListClear() As Long

    On Error Resume Next
    
    ReDim AppsRunning(0)

    AppListClear = 1

End Function

Public Function AppListShuffle() As Long

    On Error Resume Next
    
    Dim r As Long
    Dim i As Long
    Dim c As Long

    On Error GoTo errHandle
    
    r = 1
    
    For c = 0 To 5
    
        For i = 0 To UBound(AppsRunning())
        
            AppListMove i, RandomNumber(0, UBound(AppsRunning()))
        
        Next
    
    Next
    
    AppListShuffle = r
    
errHandle:

    If Err.Number <> 0 Then
    
        r = -1
        
        Err.Clear
        
        Resume Next

    End If

    
End Function

Public Function RandomNumber(lowerBound As Long, upperBound As Long) As Long

    On Error Resume Next
    
    Randomize

    RandomNumber = CLng(Int((upperBound - lowerBound + 1) * Rnd + lowerBound))
    
End Function

Public Function AppKey(KeyCode As Integer, Shift As Integer) As Boolean

    On Error Resume Next
    
    Dim ctl As Control
    For Each ctl In frmMain.Controls
        DoEvents
        If LCase(ctl.Name) = LCase(CurrentApp.Name) Then
            ctl.SendMessage "keypress##" & KeyCode & "##" & Shift
            Exit For
        End If
    Next

    AppKey = True

End Function
