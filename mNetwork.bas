Attribute VB_Name = "mNetwork"
Option Explicit

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

Public apiKey As String
Public authCode As String
Public clientID As String
Public clientSecret As String
Public accessToken As String
Public refreshToken As String
Public scope As String

Public bNetConnection As Boolean

Public bLastFmResponded As Boolean

Public sGoogleUser As String
Public sGooglePass As String
Public sGoogleName As String
Public sGooglePhoto As String
Public sGoogleCover As String
Public bGoogleLoggedIn As Boolean
Public bGoogleEnabled As Boolean

Public sLastFMUser As String
Public sLastFMPass As String
Public bLastFMLoggedIn As Boolean
Public bLastFMEnabled As Boolean

Public bLoggedIn As Boolean

Function CheckNetConnection() As Boolean
    
    On Error Resume Next

    Dim aux As String * 255
    Dim r As Long
    r = InternetGetConnectedStateEx(r, aux, 254, 0)
    
    If r = 1 Then
        CheckNetConnection = True
        SetServiceIcon 0, sOn
    Else
        CheckNetConnection = False
    End If
    
End Function

Public Function GoogleLogin(sUser As String, sPass As String) As Boolean

    On Error Resume Next

    With frmMain

        Do Until GoogleLogout = True
            DoEvents
        Loop
    
        Static i As Integer
        
        apiKey = "AIzaSyCzWjq1GgyjxrGVDy8xZas_hpBaM6Xk7Y8"
        
        authCode = "key=" & apiKey
        
        clientID = "313651221193.apps.googleusercontent.com"
        
        clientSecret = "zz1-RRp3-f2tIMMyH_LuyzPx"
        
        scope = "https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/calendar"
    
        'prepare login
            
            If .cGoogle.UBound > 0 Then
            
                For i = 1 To .cGoogle.UBound
                
                    DoEvents
                
                    Unload .cGoogle(i)
                    
                Next
                
            End If

        If bGoogleEnabled = True Then
                    
                Load .cGoogle(1)
                
            'give credentials
            
                .cGoogle(1).GiveAppCredentials apiKey, clientID, clientSecret, scope
    
            'login
            
                If refreshToken <> "" Then
                    
                    Do Until .cGoogle(1).LoginUser(sUser, sPass, refreshToken) <> 0
                        DoEvents
                    Loop
                
                Else
                    
                    Do Until .cGoogle(1).LoginUser(sUser, sPass) <> 0
                        DoEvents
                    Loop
                
                End If
            
                sGoogleUser = sUser
                sGooglePass = sPass
            
            End If
        
    End With

    GoogleLogin = True

End Function

Public Function GoogleLogout() As Boolean

    With frmMain

        CentralMessage "userloggedout", Nothing

        .lblMenuHeader(0).Caption = User.Name
        .lblMenuHeader(1).Caption = User.Type
        
        .lblMenuHeader(0).Left = .pUser.Left
        .lblMenuHeader(1).Left = .lblMenuHeader(0).Left
        
        .pUser.Picture = Nothing
            
    End With

    GoogleLogout = True

End Function

Public Function LastFMLogin(sUser As String, sPass As String)

    On Error Resume Next

    With frmMain
    
        Static i As Integer
        
        If .cLastFM.UBound > 0 Then
        
            For i = 1 To .cLastFM.UBound
            
                DoEvents
            
                Unload .cLastFM(i)
                
            Next
            
        End If
        
        If bLastFMEnabled = True Then
                
            Load .cLastFM(1)
            
            .cLastFM(1).AuthenticateLastFM sUser, sPass, "PGxUl6QQqXfD1eKljTYPFITS6woAZ2ba"
    
            Do Until bLastFmResponded = True
            
                DoEvents
                
            Loop
    
            sLastFMUser = sUser
            sLastFMPass = sPass
            
        End If

    End With

    LastFMLogin = True

End Function

Public Function CheckNetwork() As Boolean
    
    bNetConnection = CheckNetConnection
    
    'login to services
    
        Do Until GoogleLogout = True
            DoEvents
        Loop
    
        If bNetConnection = True Then
        
            Do Until GoogleLogin(sGoogleUser, sGooglePass) = True
                DoEvents
            Loop
            
            Do Until LastFMLogin(sLastFMUser, sLastFMPass) = True
                DoEvents
            Loop
            
        End If
   
    CheckNetwork = True

End Function
