Attribute VB_Name = "mInitialise"
Option Explicit

Public bAppDirect As Boolean

Public Sub Main()

    On Error Resume Next

    Load frmMain

End Sub

Public Function AppDirect(cApp As Control, Optional sArg As String = "") As Boolean

    With frmMain
                           
        .Width = frmSplash.Width
                
        .pServiceIcon(0).Visible = True
        
        .pServiceIcon(1).Visible = True
        
        .pServiceIcon(2).Visible = True
        
        .pServiceIcon(3).Visible = True
    
        .pServiceIcon(4).Visible = True
    
        .lblDateTime.Visible = True
            
        Unload frmSplash
                    
        .Visible = True
        
        .SetFocus
    
        Do Until OpenApp(cApp, sArg) = True
            DoEvents
        Loop
        
        bAppDirect = True
        
        bKey = True
        
        .pFocus.SetFocus
    
    End With

    AppDirect = True

End Function
