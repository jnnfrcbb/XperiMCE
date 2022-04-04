VERSION 5.00
Begin VB.Form frmScroll 
   Caption         =   "Form1"
   ClientHeight    =   1850
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   2980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1850
   ScaleWidth      =   2980
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrScrollRowsDown_Timer()

    On Error Resume Next
    
    Static i As Integer
    Dim r As Integer
    
    i = i + 1
    
    For r = 0 To lblRow.UBound
            
        lblRow(r).Top = lblRow(r).Top + lScroll(r)
    
    Next

    If i = lScrollSpeed Then
    
        i = 0
    
        tmrScrollRowsDown.Enabled = False
    
    End If
    
End Sub

Private Sub tmrScrollRowsUp_Timer()

    On Error Resume Next
    
    Static i As Integer
    Dim r As Integer
    
    i = i + 1
    
    For r = 0 To lblRow.UBound
            
        lblRow(r).Top = lblRow(r).Top - lScroll(r)
    
    Next

    If i = lScrollSpeed Then
    
        i = 0
    
        tmrScrollRowsUp.Enabled = False
    
    End If
    
End Sub

Private Sub Form_Load()

End Sub
