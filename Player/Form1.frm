VERSION 5.00
Object = "{709AD0CF-BEBB-4454-BA1E-61793F4CB639}#204.0#0"; "cPlayer.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldEQ 
      Height          =   2955
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   120
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   5212
      _Version        =   327682
      Orientation     =   1
      Min             =   -15
      Max             =   15
      TickStyle       =   2
      TickFrequency   =   5
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   2940
      TabIndex        =   16
      Top             =   1860
      Width           =   1215
   End
   Begin cPlayer.PlayerControl Player 
      Left            =   4140
      Top             =   60
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Radio"
      Height          =   495
      Left            =   60
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Skip +"
      Height          =   495
      Left            =   5820
      TabIndex        =   14
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load MP3"
      Height          =   495
      Left            =   4620
      TabIndex        =   13
      Top             =   3660
      Width           =   1215
   End
   Begin VB.ListBox lstTracks 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Timer tmrpos 
      Interval        =   100
      Left            =   900
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play / Pause"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   1755
   End
   Begin VB.PictureBox prLevel 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   0
      Left            =   3300
      ScaleHeight     =   825
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtTrack 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Text            =   "1"
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load CD"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.PictureBox prLevel 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   1
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   60
      Width           =   315
   End
   Begin ComctlLib.Slider sldEQ 
      Height          =   2955
      Index           =   1
      Left            =   5460
      TabIndex        =   18
      Top             =   120
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   5212
      _Version        =   327682
      Orientation     =   1
      Min             =   -15
      Max             =   15
      TickStyle       =   2
      TickFrequency   =   5
   End
   Begin ComctlLib.Slider sldEQ 
      Height          =   2955
      Index           =   2
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   5212
      _Version        =   327682
      Orientation     =   1
      Min             =   -15
      Max             =   15
      TickStyle       =   2
      TickFrequency   =   5
   End
   Begin VB.Label lblEQ 
      Alignment       =   2  'Center
      Caption         =   "13000"
      Height          =   435
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label lblEQ 
      Alignment       =   2  'Center
      Caption         =   "7000"
      Height          =   435
      Index           =   1
      Left            =   5460
      TabIndex        =   11
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label lblEQ 
      Alignment       =   2  'Center
      Caption         =   "1000"
      Height          =   435
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label lblCDTitle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   4455
   End
   Begin VB.Label lblCDArtist 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblPos 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label lblCD 
      Caption         =   " "
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Do Until Player.CDLoad("E:\", True) = True
        DoEvents
    Loop
   
    lblCD.Caption = Player.CDTrackCount & " tracks"
    
    'lblCD.Caption = lblCD.Caption & vbCr & Player.CDFreeDBID
    
    Static i As Integer
    
    For i = 0 To Player.CDTrackCount - 1
    
        lstTracks.AddItem Player.CDTrackArtist(CLng(i)) & " - " & Player.cdtrackname(CLng(i))
        
    Next
    
End Sub

Private Sub Command2_Click()

    Player.EQChannelCount 3
    
    Player.PlaybackSpeakers = stereo
    
    Player.FileLoad "C:\1.flac", eLocalFile
    
    tmrpos.Enabled = True
    
    Player.Fileplay
    
End Sub

Private Sub Command3_Click()

    Player.filepause

End Sub

Private Sub Command4_Click()

    Player.SetPosition Player.FilePosition + 5

End Sub

Private Sub Command5_Click()

    Player.FileLoad "http://network.absoluteradio.co.uk/core/audio/mp3/live.pls?service=vrbb", eNetRadio

    'Player.FileLoad "http://www.bbc.co.uk/radio/listen/live/r6_aaclca.pls", eNetRadio

    Player.StreamSave = True

End Sub

Private Sub Command6_Click()

    Player.ClosePlayer

    Player.PlayerShutdown

End Sub

Private Sub Form_Load()

    Player.PlayerVolume = 25

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Player.ClosePlayer
    
End Sub

Private Sub lstTracks_DblClick()

    Player.CDLoadTrack lstTracks.ListIndex

End Sub

Private Sub Player_CDArtist(Artist As String)

    lblCDArtist.Caption = Artist

End Sub

Private Sub Player_CDTitle(Title As String)

    lblCDTitle.Caption = Title

End Sub

Private Sub Player_CDTrackDetails(Index As Long, TrackTitle As String, TrackArtist As String)

    'lstTracks.List(Index) = TrackArtist & " - " & TrackTitle

End Sub

Private Sub Player_LevelChange(Left As Long, Right As Long)

'    prLevel(0).Value = Left
'    prLevel(1).Value = Right

End Sub

Private Sub Player_StreamUpdate(url As String, sName As String, sBPS As String, sGenre As Variant, sSong As String)

    Debug.Print sName
    Debug.Print sSong

End Sub

Private Sub sldEQ_Change(Index As Integer)

    Player.EQSetChannel CLng(Index), CLng(lblEQ(Index).Caption), sldEQ(Index).Value

End Sub

Private Sub sldEQ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Player.EQSetChannel CLng(Index), CLng(lblEQ(Index).Caption), sldEQ(Index).Value

End Sub

Private Sub sldEQ_Scroll(Index As Integer)

    Player.EQSetChannel CLng(Index), CLng(lblEQ(Index).Caption), sldEQ(Index).Value

End Sub

Private Sub tmrpos_Timer()

    lblPos.Caption = Player.FilePositionString & " / " & Player.FileDurationString

End Sub
