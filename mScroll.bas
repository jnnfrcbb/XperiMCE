Attribute VB_Name = "mScroll"
Option Explicit

' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

' Position Checking
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hwnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetAncestor Lib "user32.dll" ( _
                ByVal hwnd As Long, _
                ByVal gaFlags As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157
Private Const GA_ROOT = 2

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Check Messages
' ================================================
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(Lwnd)
      If fFrm Is Nothing Then
        ' it's not a form
        If Not IsOver(Lwnd, Xpos, Ypos) And IsOver(GetAncestor(Lwnd, GA_ROOT), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetAncestor(Lwnd, GA_ROOT)).MouseWheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.hwnd, Xpos, Ypos) Then MouseWheel MouseKeys, Rotation, Xpos, Ypos
      End If
  End Select
  
  WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hwnd As Long)
  On Error Resume Next
  SetProp hwnd, "PrevWndProc", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hwnd As Long)
  On Error Resume Next
  SetWindowLong hwnd, GWL_WNDPROC, GetProp(hwnd, "PrevWndProc")
  RemoveProp hwnd, "PrevWndProc"
End Sub

' Window Checks
' ================================================
Public Function IsOver(ByVal hwnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hwnd, rectCtl
  With rectCtl
    IsOver = (lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom)
  End With
End Function

Private Function GetForm(ByVal hwnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hwnd = hwnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  picBox.Cls
  picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
End Sub

''''''''

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
        Dim ctl As Control, cContainerCtl As Control
    Dim bHandled As Boolean
    Dim bOver As Boolean
    
    If frmMain.WindowState <> 1 Then
        
        If Rotation < 0 Then
        
            KeyPress vbKeyDown
            
        Else
        
            KeyPress vbKeyUp
        
        End If
        
    End If
        
End Sub
