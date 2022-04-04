VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAppCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AppCommandConstants
   APPCOMMAND_BROWSER_BACKWARD = 1
   APPCOMMAND_BROWSER_FORWARD = 2
   APPCOMMAND_BROWSER_REFRESH = 3
   APPCOMMAND_BROWSER_STOP = 4
   APPCOMMAND_BROWSER_SEARCH = 5
   APPCOMMAND_BROWSER_FAVORITES = 6
   APPCOMMAND_BROWSER_HOME = 7
   APPCOMMAND_VOLUME_MUTE = 8
   APPCOMMAND_VOLUME_DOWN = 9
   APPCOMMAND_VOLUME_UP = 10
   APPCOMMAND_MEDIA_NEXTTRACK = 11
   APPCOMMAND_MEDIA_PREVIOUSTRACK = 12
   APPCOMMAND_MEDIA_STOP = 13
   APPCOMMAND_MEDIA_PLAY_PAUSE = 14
   APPCOMMAND_LAUNCH_MAIL = 15
   APPCOMMAND_LAUNCH_MEDIA_SELECT = 16
   APPCOMMAND_LAUNCH_APP1 = 17
   APPCOMMAND_LAUNCH_APP2 = 18
   APPCOMMAND_BASS_DOWN = 19
   APPCOMMAND_BASS_BOOST = 20
   APPCOMMAND_BASS_UP = 21
   APPCOMMAND_TREBLE_DOWN = 22
   APPCOMMAND_TREBLE_UP = 23
   APPCOMMAND_MICROPHONE_VOLUME_MUTE = 24
   APPCOMMAND_MICROPHONE_VOLUME_DOWN = 25
   APPCOMMAND_MICROPHONE_VOLUME_UP = 26
   APPCOMMAND_HELP = 27
   APPCOMMAND_FIND = 28
   APPCOMMAND_NEW = 29
   APPCOMMAND_OPEN = 30
   APPCOMMAND_CLOSE = 31
   APPCOMMAND_SAVE = 32
   APPCOMMAND_PRINT = 33
   APPCOMMAND_UNDO = 34
   APPCOMMAND_REDO = 35
   APPCOMMAND_COPY = 36
   APPCOMMAND_CUT = 37
   APPCOMMAND_PASTE = 38
   APPCOMMAND_REPLY_TO_MAIL = 39
   APPCOMMAND_FORWARD_MAIL = 40
   APPCOMMAND_SEND_MAIL = 41
   APPCOMMAND_SPELL_CHECK = 42
   APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43
   APPCOMMAND_MIC_ON_OFF_TOGGLE = 44
   APPCOMMAND_CORRECTION_LIST = 45
   APPCOMMAND_MEDIA_PLAY = 46
   APPCOMMAND_MEDIA_PAUSE = 47
   APPCOMMAND_MEDIA_RECORD = 48
   APPCOMMAND_MEDIA_FAST_FORWARD = 49
   APPCOMMAND_MEDIA_REWIND = 50
   APPCOMMAND_MEDIA_CHANNEL_UP = 51
   APPCOMMAND_MEDIA_CHANNEL_DOWN = 52
End Enum

Public Enum AppCommandDeviceConstants
   FAPPCOMMAND_MOUSE = &H8000&
   FAPPCOMMAND_KEY = 0
   FAPPCOMMAND_OEM = &H1000&
End Enum

Public Enum AppCommandKeyStateConstants
   MK_LBUTTON = &H1
   MK_RBUTTON = &H2
   MK_SHIFT = &H4
   MK_CONTROL = &H8
   MK_MBUTTON = &H10
   MK_XBUTTON1 = &H20
   MK_XBUTTON2 = &H40
End Enum

Private Const FAPPCOMMAND_MASK As Long = &HF000&

Private Const WM_APPCOMMAND As Long = &H319&

Private m_hWnd As Long

Public Event AppCommand( _
      ByVal command As AppCommandConstants, _
      ByVal fromDevice As AppCommandDeviceConstants, _
      ByVal keyState As AppCommandKeyStateConstants, _
      ByRef processed As Boolean _
      )

Implements ISubclass

Public Sub Attach(ByVal hwnd As Long)
   Detach
   ' Start subclassing for the message.
   ' Will throw error if hWnd is not valid.
   AttachMessage Me, hwnd, WM_APPCOMMAND
   m_hWnd = hwnd
End Sub

Public Sub Detach()
   ' Stop subclassing if we are subclassing:
   If Not (m_hWnd = 0) Then
      DetachMessage Me, m_hWnd, WM_APPCOMMAND
      m_hWnd = 0
   End If
End Sub

Private Sub Class_Initialize()
   ' intentionally blank
End Sub

Private Sub Class_Terminate()
   '
   Detach
   '
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   ' intentionally blank
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
      On Error Resume Next
      
   Select Case iMsg
   Case WM_APPCOMMAND
      
      Dim cmd As Long
      ' app command is the hiword of the message with the
      ' device details in the highest 4 bits excluded:
      cmd = (lParam And &HFFF0000) / &H10000
      
      Dim fromDevice As Long
      ' the device is derived from the highest 4 bits:
      fromDevice = (lParam And &H70000000) / &H10000
      If (lParam And &H80000000) = &H80000000 Then
         fromDevice = fromDevice Or &H8000&
      End If
            
      Dim keys As Long
      ' the key details are in the loword:
      keys = lParam And &HFFFF&
      
      Dim processed As Boolean
                        
      RaiseEvent AppCommand(cmd, fromDevice, keys, processed)
      
      If (processed) Then
         ' tell windows we've used it:
         ISubclass_WindowProc = 1
      Else
         ' pass on to next handler:
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
      
   End Select
   
End Function




