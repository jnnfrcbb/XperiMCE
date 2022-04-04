VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{44D09AE3-1847-41E9-B1EF-890580211EC2}#1.0#0"; "AlphaImage.ocx"
Object = "{709AD0CF-BEBB-4454-BA1E-61793F4CB639}#204.0#0"; "cPlayer.ocx"
Object = "{C9D0B3A0-231C-418A-A820-59266F495E43}#4.0#0"; "prjLastFM.ocx"
Object = "{454BC717-09D7-4B22-B733-CAA2F3634B19}#5.0#0"; "prjKeyboard.ocx"
Object = "{2D421F81-4B63-41A8-BC5D-5E19ABF7C41B}#4.0#0"; "prjSport.ocx"
Object = "{E4B0F05F-76C4-4968-BE36-D5CA6B801F54}#16.1#0"; "prjVideos.ocx"
Object = "{78A4D0CF-E0D8-409C-AB9D-FF89D3D7F609}#11.0#0"; "cGoogleAuth.ocx"
Object = "{79A8D05F-8453-4684-8B83-13A4A8B0D380}#13.0#0"; "prjHoriSccroll.ocx"
Object = "{49C056E5-3380-4DF0-97D3-F0D03BF1760B}#4.2#0"; "prjWeather.ocx"
Object = "{0943EB94-7787-4E0B-B1A0-A163ECE12205}#2.1#0"; "prjNews.ocx"
Object = "{BDE8178E-60D8-46EB-A0CE-2C74B3A127C9}#4.1#0"; "prjGames.ocx"
Object = "{9372AE8B-CB9C-414C-8540-7C3D898886C8}#10.2#0"; "prjMusic.ocx"
Object = "{F5A336BA-2331-41FF-85E1-00221C89CD19}#17.0#0"; "prjNetwork.ocx"
Object = "{A53FDAAA-172D-4F90-918B-696A40D6849E}#4.0#0"; "prjCalendar.ocx"
Object = "{CA238D15-DCEB-4658-B233-9024AA2D001A}#3.0#0"; "prjNotifications.ocx"
Object = "{334117A3-CDAB-4790-BADF-A373EEE9BF61}#2.0#0"; "prjGallery.ocx"
Object = "{8B999C6C-9E18-43AD-8952-CBFBFC31565E}#3.1#0"; "prjDevices.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "XperiMCE"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Segoe UI Light"
      Size            =   21.75
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   960
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox pVidMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H001E1E1E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1400
      Left            =   4050
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   6300
      Begin VB.Timer tmrShowOSD 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   8760
         Top             =   600
      End
      Begin VB.Timer tmrOSD 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   8760
         Top             =   120
      End
      Begin VB.PictureBox pVidPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   90
         Index           =   0
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   418
         TabIndex        =   27
         Top             =   900
         Width           =   6300
         Begin VB.PictureBox pVidPosition 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDCDCD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   -15
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   0
            TabIndex        =   28
            Top             =   -15
            Width           =   8
         End
         Begin VB.PictureBox pVidPosition 
            Appearance      =   0  'Flat
            BackColor       =   &H00616161&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   -15
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   0
            TabIndex        =   38
            Top             =   -15
            Visible         =   0   'False
            Width           =   8
         End
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPlayPause 
         Height          =   320
         Index           =   2
         Left            =   6360
         Top             =   180
         Visible         =   0   'False
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":8D25A
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":9095A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPlayPause 
         Height          =   320
         Index           =   1
         Left            =   5880
         Top             =   600
         Visible         =   0   'False
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":90972
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":916EE
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPlayPause 
         Height          =   320
         Index           =   0
         Left            =   5940
         Top             =   900
         Visible         =   0   'False
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":91706
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":9253F
      End
      Begin VB.Label lblVidDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Video title"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   280
         Index           =   1
         Left            =   180
         TabIndex        =   30
         Top             =   1040
         Width           =   830
      End
      Begin VB.Label lblVidDetails 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Video Time"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   280
         Index           =   0
         Left            =   5200
         TabIndex        =   29
         Top             =   1040
         Width           =   940
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   5
         Left            =   4500
         Tag             =   "playlist"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":92557
         Blend           =   65280
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":93300
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   0
         Left            =   0
         Tag             =   "repeat"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":93318
         Blend           =   65280
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":94840
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   1
         Left            =   900
         Tag             =   "shuffle"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":94858
         Blend           =   65280
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":961CE
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   4
         Left            =   3660
         Tag             =   "skipforwardlarge"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":961E6
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":970F6
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   2
         Left            =   1800
         Tag             =   "skipback"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":9710E
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":98015
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   3
         Left            =   2700
         Tag             =   "playpause"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":9802D
         Frame           =   12
         BackColor       =   16288796
         Attr            =   1538
         Effects         =   "frmMain.frx":98E66
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVidControl 
         Height          =   900
         Index           =   6
         Left            =   5400
         Tag             =   "vidmin"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":98E7E
         Blend           =   65280
         Settings        =   19200
         BackColor       =   16288796
         Attr            =   514
         Effects         =   "frmMain.frx":9AAB6
      End
      Begin VB.Shape sVidControl 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   900
         Left            =   0
         Top             =   0
         Width           =   6300
      End
   End
   Begin VB.PictureBox pVideoHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   14000
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   14400
      Begin VB.Timer tmrYT 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   60
         Top             =   1920
      End
      Begin VB.Timer tmrMouse 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2940
         Top             =   4140
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmp 
         Height          =   3300
         Left            =   9480
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   3900
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   6879
         _cy             =   5821
      End
   End
   Begin VB.Timer tmrClientTimeout 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7060
      Top             =   1440
   End
   Begin VB.PictureBox pPlaylistHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   14400
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   14400
      Begin prjHoriSccroll.cItems cPlaylist 
         Height          =   8100
         Left            =   0
         Top             =   2350
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
         Transparent     =   -1  'True
      End
      Begin VB.Label lblPlaylist 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Playlist"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   1215
         TabIndex        =   35
         Top             =   1720
         Width           =   7995
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPlaylistHolderBack 
         Height          =   8100
         Left            =   0
         Top             =   0
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
         Image           =   "frmMain.frx":9AACE
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":9C06C
      End
   End
   Begin VB.Timer tmrRumble 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9420
      Top             =   7560
   End
   Begin VB.PictureBox pFocus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   0
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   0
      Width           =   14400
      Begin VB.PictureBox pEQHolder 
         Height          =   3135
         Index           =   0
         Left            =   3420
         ScaleHeight     =   3075
         ScaleWidth      =   7515
         TabIndex        =   98
         Top             =   2100
         Visible         =   0   'False
         Width           =   7575
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   9
            Left            =   6840
            Max             =   -15
            Min             =   15
            TabIndex        =   117
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   8
            Left            =   6120
            Max             =   -15
            Min             =   15
            TabIndex        =   115
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   7
            Left            =   5400
            Max             =   -15
            Min             =   15
            TabIndex        =   113
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   6
            Left            =   4680
            Max             =   -15
            Min             =   15
            TabIndex        =   111
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   5
            Left            =   3960
            Max             =   -15
            Min             =   15
            TabIndex        =   109
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   4
            Left            =   3240
            Max             =   -15
            Min             =   15
            TabIndex        =   107
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   3
            Left            =   2520
            Max             =   -15
            Min             =   15
            TabIndex        =   105
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   2
            Left            =   1800
            Max             =   -15
            Min             =   15
            TabIndex        =   103
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   1
            Left            =   1080
            Max             =   -15
            Min             =   15
            TabIndex        =   101
            Top             =   180
            Width           =   315
         End
         Begin VB.VScrollBar scrEQ 
            Height          =   1935
            Index           =   0
            Left            =   360
            Max             =   -15
            Min             =   15
            TabIndex        =   99
            Top             =   180
            Width           =   315
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   6720
            TabIndex        =   128
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   6000
            TabIndex        =   127
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   5280
            TabIndex        =   126
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   4560
            TabIndex        =   125
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3840
            TabIndex        =   124
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   123
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2400
            TabIndex        =   122
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   121
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   120
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQValue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   119
            Tag             =   "32"
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "16 kHz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   6720
            TabIndex        =   118
            Tag             =   "16000"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8 kHz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   6000
            TabIndex        =   116
            Tag             =   "8000"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4 kHz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   5280
            TabIndex        =   114
            Tag             =   "4000"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2 kHz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   4560
            TabIndex        =   112
            Tag             =   "2000"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1 kHz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3840
            TabIndex        =   110
            Tag             =   "1000"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500 Hz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   108
            Tag             =   "500"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "250 Hz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2400
            TabIndex        =   106
            Tag             =   "250"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "125 Hz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   104
            Tag             =   "125"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "64 Hz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   102
            Tag             =   "64"
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label lblEQ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "32 Hz"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Tag             =   "32"
            Top             =   2160
            Width           =   555
         End
      End
      Begin VB.PictureBox pRowSpace 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   1200
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   39
         Top             =   2558
         Visible         =   0   'False
         Width           =   975
         Begin VB.PictureBox pRowTitleSpace 
            Height          =   75
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   555
            TabIndex        =   41
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox pMenuHolder 
         BackColor       =   &H004B4B4B&
         BorderStyle     =   0  'None
         Height          =   81000
         Left            =   900
         ScaleHeight     =   5400
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   240
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   3600
         Begin VB.Timer tmrCloseMenu 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1920
            Top             =   5520
         End
         Begin VB.Timer tmrOpenMenu 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   600
            Top             =   5340
         End
         Begin VB.PictureBox pMenuWidth 
            Height          =   190
            Left            =   0
            ScaleHeight     =   135
            ScaleWidth      =   3540
            TabIndex        =   74
            Top             =   720
            Visible         =   0   'False
            Width           =   3600
         End
         Begin VB.Line lnMenu 
            BorderColor     =   &H00616161&
            Index           =   1
            X1              =   30
            X2              =   330
            Y1              =   720
            Y2              =   720
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pServiceIcon 
            Height          =   320
            Index           =   4
            Left            =   2520
            Top             =   7500
            Visible         =   0   'False
            Width           =   320
            _ExtentX        =   556
            _ExtentY        =   556
            Image           =   "frmMain.frx":9C084
            Blend           =   1677721600
            Settings        =   16777216
            Render          =   4
            Frame           =   516
            Attr            =   513
            Effects         =   "frmMain.frx":9FA71
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pServiceIcon 
            Height          =   320
            Index           =   2
            Left            =   1680
            Top             =   7500
            Visible         =   0   'False
            Width           =   320
            _ExtentX        =   556
            _ExtentY        =   556
            Image           =   "frmMain.frx":9FA89
            Blend           =   1677721600
            Settings        =   16777216
            Render          =   4
            Frame           =   516
            Attr            =   513
            Effects         =   "frmMain.frx":A9422
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pServiceIcon 
            Height          =   320
            Index           =   3
            Left            =   2100
            Top             =   7500
            Visible         =   0   'False
            Width           =   320
            _ExtentX        =   556
            _ExtentY        =   556
            Image           =   "frmMain.frx":A943A
            Blend           =   1677721600
            Settings        =   16777216
            Render          =   4
            Frame           =   516
            Attr            =   513
            Effects         =   "frmMain.frx":AC00A
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pServiceIcon 
            Height          =   320
            Index           =   0
            Left            =   840
            Top             =   7500
            Visible         =   0   'False
            Width           =   320
            _ExtentX        =   556
            _ExtentY        =   556
            Image           =   "frmMain.frx":AC022
            Blend           =   1677721600
            Settings        =   16777216
            Render          =   4
            Frame           =   516
            Attr            =   513
            Effects         =   "frmMain.frx":B09BF
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pServiceIcon 
            Height          =   320
            Index           =   1
            Left            =   1260
            Top             =   7500
            Visible         =   0   'False
            Width           =   320
            _ExtentX        =   556
            _ExtentY        =   556
            Image           =   "frmMain.frx":B09D7
            Blend           =   1677721600
            Settings        =   16777216
            Render          =   4
            Frame           =   516
            Attr            =   513
            Effects         =   "frmMain.frx":B1509
         End
         Begin VB.Label lblMenu 
            BackColor       =   &H00787878&
            BackStyle       =   0  'Transparent
            Caption         =   "  Shutdown"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   14.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   440
            Index           =   4
            Left            =   0
            TabIndex        =   81
            Top             =   3300
            Width           =   3450
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMenuHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "User Email"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   11.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   380
            Index           =   1
            Left            =   920
            TabIndex        =   80
            Top             =   470
            Width           =   2540
            WordWrap        =   -1  'True
         End
         Begin VB.Line lnMenu 
            BorderColor     =   &H00616161&
            Index           =   0
            X1              =   30
            X2              =   330
            Y1              =   258
            Y2              =   258
         End
         Begin VB.Label lblMenuHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   500
            Index           =   0
            Left            =   920
            TabIndex        =   79
            Top             =   120
            Width           =   2540
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMenu 
            BackColor       =   &H00787878&
            BackStyle       =   0  'Transparent
            Caption         =   "  Playlist"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   14.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   440
            Index           =   1
            Left            =   0
            TabIndex        =   78
            Top             =   1440
            Width           =   3450
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMenu 
            BackColor       =   &H00787878&
            BackStyle       =   0  'Transparent
            Caption         =   "  Lock"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   14.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   440
            Index           =   3
            Left            =   0
            TabIndex        =   77
            Top             =   2820
            Width           =   3450
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMenu 
            BackColor       =   &H00787878&
            BackStyle       =   0  'Transparent
            Caption         =   "  Home"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   14.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   440
            Index           =   0
            Left            =   0
            TabIndex        =   76
            Top             =   960
            Width           =   3450
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMenu 
            BackColor       =   &H00787878&
            BackStyle       =   0  'Transparent
            Caption         =   "  Device Manager"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   14.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   440
            Index           =   2
            Left            =   0
            TabIndex        =   75
            Top             =   1920
            Width           =   3450
            WordWrap        =   -1  'True
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pUser 
            Height          =   620
            Left            =   150
            Top             =   150
            Width           =   620
            _ExtentX        =   1085
            _ExtentY        =   1085
            Image           =   "frmMain.frx":B1521
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B28D1
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pUserCover 
            Height          =   900
            Left            =   0
            Top             =   0
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   1588
            Blend           =   1677721600
            Settings        =   1048576
            Render          =   4
            Frame           =   12
            BackColor       =   3947580
            Attr            =   513
            Effects         =   "frmMain.frx":B28E9
         End
      End
      Begin VB.PictureBox pNotifyHolder 
         BackColor       =   &H00EBEBEB&
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   0
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   960
         TabIndex        =   91
         Top             =   7200
         Visible         =   0   'False
         Width           =   14400
         Begin VB.Timer tmrRemote 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   10200
            Top             =   240
         End
         Begin VB.Timer tmrNotify 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9840
            Top             =   240
         End
         Begin VB.Timer tmrCheckForProcess 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   10560
            Top             =   240
         End
         Begin VB.Timer tmrMouseOut 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   12360
            Top             =   240
         End
         Begin VB.Timer tmrDateTime 
            Interval        =   100
            Left            =   11640
            Top             =   240
         End
         Begin VB.Timer tmrWait 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   12000
            Top             =   240
         End
         Begin VB.Timer tmrStartup 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   12720
            Top             =   240
         End
         Begin VB.Timer tmrXbox 
            Enabled         =   0   'False
            Interval        =   110
            Left            =   11280
            Top             =   240
         End
         Begin VB.Timer tmrPlayback 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   10920
            Top             =   240
         End
         Begin prjGoogleAuth.cGoogleAuth cGoogle 
            Height          =   440
            Index           =   0
            Left            =   7680
            TabIndex        =   92
            Top             =   180
            Visible         =   0   'False
            Width           =   440
            _ExtentX        =   767
            _ExtentY        =   767
         End
         Begin prjLastFM.cLastFM cLastFM 
            Height          =   380
            Index           =   0
            Left            =   8220
            TabIndex        =   93
            Top             =   180
            Visible         =   0   'False
            Width           =   380
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin cPlayer.PlayerControl cPlayer 
            Index           =   0
            Left            =   9300
            Top             =   180
            _ExtentX        =   767
            _ExtentY        =   767
         End
         Begin prjNetwork.cServer cServer 
            Height          =   500
            Left            =   8700
            TabIndex        =   94
            Top             =   180
            Visible         =   0   'False
            Width           =   500
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pNotifyDefault 
            Height          =   330
            Index           =   2
            Left            =   5520
            Top             =   420
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Image           =   "frmMain.frx":B2901
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B373A
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pNotifyDefault 
            Height          =   330
            Index           =   1
            Left            =   5100
            Top             =   420
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Image           =   "frmMain.frx":B3752
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B44CE
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pNotifyDefault 
            Height          =   330
            Index           =   0
            Left            =   4740
            Top             =   420
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Image           =   "frmMain.frx":B44E6
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B5488
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pNotifyIcon 
            Height          =   600
            Left            =   150
            Top             =   150
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            Image           =   "frmMain.frx":B54A0
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B6442
         End
         Begin VB.Label lblDateTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DateTime"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E1E1E&
            Height          =   380
            Left            =   12000
            TabIndex        =   96
            Top             =   230
            Visible         =   0   'False
            Width           =   1200
         End
         Begin WMPLibCtl.WindowsMediaPlayer cSound 
            Height          =   380
            Index           =   0
            Left            =   7260
            TabIndex        =   95
            Top             =   180
            Visible         =   0   'False
            Width           =   320
            URL             =   ""
            rate            =   1
            balance         =   0
            currentPosition =   0
            defaultFrame    =   ""
            playCount       =   1
            autoStart       =   -1  'True
            currentMarker   =   0
            invokeURLs      =   -1  'True
            baseURL         =   ""
            volume          =   50
            mute            =   0   'False
            uiMode          =   "full"
            stretchToFit    =   0   'False
            windowlessVideo =   0   'False
            enabled         =   -1  'True
            enableContextMenu=   -1  'True
            fullScreen      =   0   'False
            SAMIStyle       =   ""
            SAMILang        =   ""
            SAMIFilename    =   ""
            captioningID    =   ""
            enableErrorDialogs=   0   'False
            _cx             =   556
            _cy             =   661
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Notification text"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E1E1E&
            Height          =   500
            Left            =   1200
            TabIndex        =   97
            Top             =   230
            Width           =   10450
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00CDCDCD&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00CDCDCD&
            Height          =   900
            Left            =   0
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox pSideFadeBox 
         BackColor       =   &H003C3C3C&
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   0
         Left            =   0
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         Begin LaVolpeAlphaImg.AlphaImgCtl pSideFadeIcon 
            Height          =   750
            Index           =   0
            Left            =   80
            Top             =   100
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1323
            Image           =   "frmMain.frx":B645A
            Settings        =   -2147483648
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":B7535
         End
      End
      Begin VB.PictureBox pSideFadeBox 
         BackColor       =   &H003C3C3C&
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   1
         Left            =   0
         ScaleHeight     =   90
         ScaleMode       =   0  'User
         ScaleWidth      =   89.732
         TabIndex        =   87
         Top             =   1680
         Visible         =   0   'False
         Width           =   900
         Begin LaVolpeAlphaImg.AlphaImgCtl pSideFadeIcon 
            Height          =   500
            Index           =   1
            Left            =   228
            Top             =   210
            Width           =   500
            _ExtentX        =   873
            _ExtentY        =   873
            Image           =   "frmMain.frx":B754D
            Settings        =   8448
            Attr            =   513
            Effects         =   "frmMain.frx":B77BC
         End
      End
      Begin VB.PictureBox pTileOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004B4B4B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8100
         Left            =   14400
         ScaleHeight     =   540
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   960
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         Begin VB.PictureBox pTileOptionsHolder 
            BackColor       =   &H003C3C3C&
            BorderStyle     =   0  'None
            Height          =   1540
            Left            =   4950
            ScaleHeight     =   103
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   300
            TabIndex        =   82
            Top             =   3500
            Width           =   4500
            Begin LaVolpeAlphaImg.AlphaImgCtl pTile 
               Height          =   900
               Index           =   3
               Left            =   2700
               Tag             =   "edittile"
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1588
               Image           =   "frmMain.frx":B77D4
               Settings        =   19200
               Attr            =   513
               Effects         =   "frmMain.frx":B934C
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl pTile 
               Height          =   900
               Index           =   2
               Left            =   1800
               Tag             =   "deletetile"
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1588
               Image           =   "frmMain.frx":B9364
               Settings        =   19200
               Attr            =   513
               Effects         =   "frmMain.frx":BA6E0
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl pTile 
               Height          =   900
               Index           =   1
               Left            =   900
               Tag             =   "lockdash"
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1588
               Image           =   "frmMain.frx":BA6F8
               Settings        =   19200
               Attr            =   513
               Effects         =   "frmMain.frx":BBDA8
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl pTile 
               Height          =   900
               Index           =   0
               Left            =   0
               Tag             =   "power"
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1588
               Image           =   "frmMain.frx":BBDC0
               Settings        =   19200
               Attr            =   513
               Effects         =   "frmMain.frx":BDDBA
            End
            Begin VB.Label lblTileOptions 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Power Options"
               BeginProperty Font 
                  Name            =   "Segoe UI Light"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   300
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   680
               Left            =   0
               TabIndex        =   83
               Top             =   960
               Width           =   4370
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl pTile 
               Height          =   900
               Index           =   4
               Left            =   3600
               Tag             =   "addtile"
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1588
               Image           =   "frmMain.frx":BDDD2
               Settings        =   19200
               Attr            =   513
               Effects         =   "frmMain.frx":BEB99
            End
            Begin VB.Shape sTileOptions 
               BackColor       =   &H004B4B4B&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H004B4B4B&
               Height          =   610
               Index           =   0
               Left            =   0
               Top             =   910
               Width           =   5510
            End
         End
         Begin VB.PictureBox pSideFadeBox 
            BackColor       =   &H003C3C3C&
            BorderStyle     =   0  'None
            Height          =   900
            Index           =   2
            Left            =   0
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   72
            Top             =   1800
            Visible         =   0   'False
            Width           =   900
            Begin LaVolpeAlphaImg.AlphaImgCtl pSideFadeIcon 
               Height          =   700
               Index           =   2
               Left            =   110
               Top             =   95
               Width           =   700
               _ExtentX        =   1244
               _ExtentY        =   1244
               Image           =   "frmMain.frx":BEBB1
               Settings        =   8448
               Attr            =   513
               Effects         =   "frmMain.frx":C0732
            End
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pTileOptionsBack 
            Height          =   8100
            Left            =   0
            Top             =   0
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   14288
            Image           =   "frmMain.frx":C074A
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":C1CE8
         End
      End
      Begin VB.PictureBox pDashSettingsHolder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00181818&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   19200
         ScaleHeight     =   401
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   960
         TabIndex        =   6
         Tag             =   "225,12,0"
         Top             =   1080
         Visible         =   0   'False
         Width           =   14400
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00484642&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   11
            Left            =   10440
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   23
            Tag             =   "66,70,72"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   11
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C1D00
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C2880
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00A7A49E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   8
            Left            =   8100
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   22
            Tag             =   "158,164,167"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   8
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C2898
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C3418
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8232E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   5
            Left            =   5520
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   21
            Tag             =   "46,35,248"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   5
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C3430
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C3FB0
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   13
            Left            =   12000
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   20
            Tag             =   "0,0,0"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   13
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C3FC8
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00736958&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   10
            Left            =   9660
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   19
            Tag             =   "88,105,115"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   10
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C3FE0
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C4B60
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00303030&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   12
            Left            =   11220
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   18
            Tag             =   "48,48,48"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   12
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C4B78
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C56F8
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   9
            Left            =   8880
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   17
            Tag             =   "128,128,128"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   9
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C5710
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C6290
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   7
            Left            =   7320
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   16
            Tag             =   "255,255,255"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   7
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C62A8
               Settings        =   50
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C6E28
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00B2&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   6
            Left            =   6300
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   15
            Tag             =   "178,0,255"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   6
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C6E40
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C7AB2
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00D88200&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   4
            Left            =   4740
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   14
            Tag             =   "0,130,216"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   4
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C7ACA
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C864A
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00008A00&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   3
            Left            =   3960
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   13
            Tag             =   "0,138,0"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   3
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C8662
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C91E2
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0007ECF8&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   2
            Left            =   3180
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   12
            Tag             =   "248,236,7"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   2
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C91FA
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":C9D7A
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H001D54FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   1
            Left            =   2400
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   11
            Tag             =   "255,84,29"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   1
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":C9D92
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":CA912
            End
         End
         Begin VB.PictureBox pTileColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000CE1&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   0
            Left            =   1620
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   10
            Tag             =   "225,12,0"
            Top             =   1500
            Width           =   735
            Begin LaVolpeAlphaImg.AlphaImgCtl pDashSettingsColorSelected 
               Height          =   195
               Index           =   0
               Left            =   60
               Top             =   60
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   344
               Image           =   "frmMain.frx":CA92A
               Settings        =   100
               Render          =   4
               Attr            =   513
               Effects         =   "frmMain.frx":CB4AA
            End
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl pDashPinSetting 
            Height          =   735
            Left            =   1620
            Top             =   3300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Image           =   "frmMain.frx":CB4C2
            Settings        =   25
            Render          =   4
            Attr            =   513
            Effects         =   "frmMain.frx":CC134
         End
         Begin VB.Label lblDashSettingsTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Home dash settings"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   615
            Left            =   1200
            TabIndex        =   25
            Top             =   120
            Width           =   5175
         End
         Begin VB.Shape sDashSettingsSelect 
            BackColor       =   &H00000000&
            BorderColor     =   &H00FFFFFF&
            Height          =   795
            Left            =   1590
            Top             =   1470
            Width           =   795
         End
         Begin VB.Label lblTileSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   15.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   450
            Index           =   2
            Left            =   1410
            TabIndex        =   9
            Top             =   4320
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblTileSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allows pins"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   15.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   450
            Index           =   1
            Left            =   1410
            TabIndex        =   8
            Top             =   2520
            Width           =   1440
         End
         Begin VB.Label lblTileSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tile colour"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   15.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   450
            Index           =   0
            Left            =   1410
            TabIndex        =   7
            Top             =   720
            Width           =   1365
         End
      End
      Begin VB.PictureBox pSearchWidth 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   3840
         ScaleHeight     =   135
         ScaleWidth      =   2415
         TabIndex        =   67
         Top             =   300
         Visible         =   0   'False
         Width           =   2409
      End
      Begin VB.Timer tmrCharm 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7320
         Top             =   180
      End
      Begin VB.PictureBox pTilePinSelectHolder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00484642&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7545
         Left            =   14400
         ScaleHeight     =   503
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   960
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         Begin prjHoriSccroll.cItems cPinItems 
            Height          =   3735
            Left            =   0
            Top             =   2250
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   6588
            Transparent     =   -1  'True
         End
         Begin VB.Label lblTilePinOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Item to Pin"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   480
            Index           =   0
            Left            =   1200
            TabIndex        =   52
            Top             =   1725
            Width           =   2820
         End
         Begin VB.Label lblTilePinOptions 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin TV Show"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   480
            Index           =   1
            Left            =   11400
            TabIndex        =   51
            Top             =   6000
            Width           =   1995
         End
      End
      Begin prjWeather.cWeather cWeather 
         Height          =   8100
         Left            =   14400
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin VB.Timer tmrScrollRowsDown 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   3300
      End
      Begin VB.Timer tmrScrollRowsUp 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   2760
      End
      Begin VB.PictureBox pSize 
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   11100
         ScaleHeight     =   30
         ScaleWidth      =   30
         TabIndex        =   24
         Top             =   60
         Visible         =   0   'False
         Width           =   30
      End
      Begin prjSport.cSport cSport 
         Height          =   8100
         Left            =   14400
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjVideos.cFilms cFilms 
         Height          =   8100
         Left            =   14400
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjVideos.cTV cTV 
         Height          =   8100
         Left            =   14400
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjNews.cNews cNews 
         Height          =   8100
         Left            =   14400
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjGames.cGames cGames 
         Height          =   8100
         Left            =   14400
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjMusic.cMusicHome cMusicHome 
         Height          =   8100
         Left            =   14400
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjCalendar.cCalendar cCalendar 
         Height          =   8100
         Left            =   14400
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjGallery.cGallery cGallery 
         Height          =   8100
         Left            =   14400
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin prjDevices.cDevices cDevices 
         Height          =   8100
         Left            =   14400
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   14288
      End
      Begin VB.PictureBox pSearchHolder 
         BackColor       =   &H003C3C3C&
         BorderStyle     =   0  'None
         Height          =   5890
         Left            =   900
         ScaleHeight     =   393
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   840
         TabIndex        =   84
         Top             =   1680
         Visible         =   0   'False
         Width           =   12600
         Begin VB.Timer tmrOpenSearch 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   0
            Top             =   300
         End
         Begin VB.Timer tmrCloseSearch 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   0
            Top             =   0
         End
         Begin prjKeyboard.cKeyboard cKeyboard 
            Height          =   4400
            Left            =   300
            TabIndex        =   86
            Top             =   1200
            Visible         =   0   'False
            Width           =   12000
            _ExtentX        =   21167
            _ExtentY        =   7752
         End
         Begin prjHoriSccroll.cItems cSearch 
            Height          =   3740
            Left            =   -900
            Top             =   1658
            Visible         =   0   'False
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   6588
            Transparent     =   -1  'True
         End
         Begin VB.Label lblSearch 
            Appearance      =   0  'Flat
            BackColor       =   &H00181818&
            BackStyle       =   0  'Transparent
            Caption         =   "Type to search..."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00565656&
            Height          =   560
            Left            =   410
            TabIndex        =   85
            Top             =   200
            Visible         =   0   'False
            Width           =   11750
         End
         Begin VB.Shape sSearchBack 
            BackColor       =   &H00323232&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00444444&
            Height          =   540
            Left            =   300
            Top             =   180
            Visible         =   0   'False
            Width           =   12000
         End
         Begin VB.Shape pSearchBack 
            BackColor       =   &H004B4B4B&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H004B4B4B&
            Height          =   5000
            Left            =   0
            Top             =   900
            Width           =   12580
         End
      End
      Begin prjHoriSccroll.cItems cRow 
         Height          =   3735
         Index           =   0
         Left            =   0
         Top             =   2558
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   6588
         Transparent     =   -1  'True
         Begin LaVolpeAlphaImg.AlphaImgCtl pRowFade 
            Height          =   540
            Left            =   0
            Top             =   0
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   953
            Image           =   "frmMain.frx":CC14C
            Attr            =   513
            Effects         =   "frmMain.frx":D0AE9
         End
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pMenuIcon 
         Height          =   750
         Left            =   75
         Top             =   100
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "frmMain.frx":D0B01
         Settings        =   -2147468288
         Attr            =   513
         Effects         =   "frmMain.frx":D1BDC
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLoading 
         Height          =   500
         Index           =   0
         Left            =   200
         Top             =   200
         Visible         =   0   'False
         Width           =   500
         _ExtentX        =   873
         _ExtentY        =   873
         Image           =   "frmMain.frx":D1BF4
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":D81A4
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pSearchIcon 
         Height          =   495
         Left            =   228
         Top             =   1890
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Image           =   "frmMain.frx":D81BC
         Settings        =   15360
         Attr            =   513
         Effects         =   "frmMain.frx":D842B
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pSettingsIcon 
         Height          =   700
         Left            =   110
         Top             =   1895
         Visible         =   0   'False
         Width           =   700
         _ExtentX        =   1244
         _ExtentY        =   1244
         Image           =   "frmMain.frx":D8443
         Settings        =   15360
         Attr            =   1
         Effects         =   "frmMain.frx":D9FC4
      End
      Begin VB.Label lblAppTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "                       "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   740
         Left            =   1200
         TabIndex        =   70
         Top             =   500
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Label lblHomeTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   740
         Left            =   11600
         TabIndex        =   69
         Top             =   500
         Visible         =   0   'False
         Width           =   1600
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pAppLogo 
         Height          =   495
         Left            =   1200
         Top             =   135
         Visible         =   0   'False
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   873
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":D9FDC
      End
      Begin VB.Label lblRow 
         BackStyle       =   0  'Transparent
         Caption         =   "Test caption"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   750
         Index           =   0
         Left            =   1200
         TabIndex        =   40
         Top             =   1808
         Visible         =   0   'False
         Width           =   12000
      End
      Begin VB.Label lblWelcome 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hello"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   0
         TabIndex        =   2
         Top             =   3675
         Visible         =   0   'False
         Width           =   14400
         WordWrap        =   -1  'True
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pTopBar 
         Height          =   750
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   1323
         Render          =   4
         BackColor       =   3289650
         Attr            =   513
         Effects         =   "frmMain.frx":D9FF4
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pSideFade 
         Height          =   8100
         Left            =   0
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   14288
         Image           =   "frmMain.frx":DA00C
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":DABDD
      End
   End
   Begin VB.ListBox lstRemoteQueue 
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   12300
      TabIndex        =   57
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pVolumeHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00111111&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   13260
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   1005
      Begin VB.Timer tmrVolume 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Label lblVolume 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   0
         TabIndex        =   37
         Top             =   240
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pVolumeIcon 
         Height          =   855
         Left            =   75
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Image           =   "frmMain.frx":DABF5
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":DC903
      End
   End
   Begin VB.PictureBox pOsdPos 
      Height          =   195
      Left            =   60
      ScaleHeight     =   135
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   6165
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13380
      Top             =   120
   End
   Begin VB.PictureBox pLockHolder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00111111&
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   0
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14400
      Begin VB.Timer tmrCloseLockHolder 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   11160
         Top             =   4740
      End
      Begin prjNotifications.cNotificationWidget cNotificationWidget 
         Height          =   4050
         Left            =   3600
         TabIndex        =   60
         Top             =   2040
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   7144
      End
      Begin VB.Label lblLockNotifCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6965
         TabIndex        =   71
         Top             =   840
         Visible         =   0   'False
         Width           =   435
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockNotifGif 
         Height          =   675
         Left            =   6840
         Top             =   750
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1191
         Image           =   "frmMain.frx":DC91B
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":DCB08
      End
      Begin VB.Label lblLockPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00616161&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   36
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Index           =   3
         Left            =   8220
         TabIndex        =   65
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLockPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00616161&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   36
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Index           =   2
         Left            =   7215
         TabIndex        =   64
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLockPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00616161&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   36
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Index           =   1
         Left            =   6210
         TabIndex        =   63
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLockPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00616161&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   36
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Index           =   0
         Left            =   5190
         TabIndex        =   62
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockPin 
         Height          =   915
         Index           =   3
         Left            =   8220
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
         Image           =   "frmMain.frx":DCB20
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":DDCD4
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockPin 
         Height          =   915
         Index           =   2
         Left            =   7215
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
         Image           =   "frmMain.frx":DDCEC
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":DEEA0
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockPin 
         Height          =   915
         Index           =   1
         Left            =   6210
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
         Image           =   "frmMain.frx":DEEB8
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":E006C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockPin 
         Height          =   915
         Index           =   0
         Left            =   5190
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
         Image           =   "frmMain.frx":E0084
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":E1238
      End
      Begin VB.Label lblLockWeather 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   13725
         TabIndex        =   58
         Top             =   7140
         Width           =   75
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockIcon 
         Height          =   555
         Left            =   11760
         Top             =   6540
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         Image           =   "frmMain.frx":E1250
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":E2900
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockArt 
         Height          =   1695
         Left            =   600
         Top             =   5760
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2990
         Render          =   4
         Frame           =   20
         Border          =   6381921
         Attr            =   513
         Effects         =   "frmMain.frx":E2918
      End
      Begin VB.Label lblLockArt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   2505
         TabIndex        =   50
         Top             =   7140
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblLockArt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   2505
         TabIndex        =   46
         Top             =   6705
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblLockTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   12300
         TabIndex        =   54
         Top             =   6450
         Visible         =   0   'False
         Width           =   1500
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLockGradiant 
         Height          =   2020
         Left            =   0
         Top             =   6080
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   3572
         Image           =   "frmMain.frx":E2930
         Render          =   4
         Attr            =   513
         Effects         =   "frmMain.frx":E4B9E
      End
   End
   Begin VB.PictureBox pMinVideo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   10245
      ScaleHeight     =   2025
      ScaleWidth      =   3600
      TabIndex        =   42
      Top             =   4965
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox pMiniOSD 
      Appearance      =   0  'Flat
      BackColor       =   &H00111111&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   10245
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   238
      TabIndex        =   43
      Top             =   6990
      Visible         =   0   'False
      Width           =   3600
      Begin LaVolpeAlphaImg.AlphaImgCtl pMinVidControl 
         Height          =   315
         Index           =   3
         Left            =   1320
         Tag             =   "restore video"
         Top             =   105
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":E4BB6
         Attr            =   513
         Effects         =   "frmMain.frx":E6845
      End
      Begin VB.Label lblMinVidPos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 / 00:00"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   44
         Top             =   90
         Width           =   1755
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pMinVidControl 
         Height          =   315
         Index           =   2
         Left            =   900
         Tag             =   "playlistback"
         Top             =   105
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":E685D
         Attr            =   513
         Effects         =   "frmMain.frx":E7764
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pMinVidControl 
         Height          =   315
         Index           =   1
         Left            =   480
         Tag             =   "playlistnext"
         Top             =   105
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":E777C
         Attr            =   513
         Effects         =   "frmMain.frx":E868C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pMinVidControl 
         Height          =   315
         Index           =   0
         Left            =   60
         Tag             =   "playpause"
         Top             =   105
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Image           =   "frmMain.frx":E86A4
         Attr            =   513
         Effects         =   "frmMain.frx":E94DD
      End
   End
   Begin VB.PictureBox pPowerOptions 
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      Height          =   1540
      Left            =   4950
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   89
      Top             =   3500
      Width           =   4500
      Begin VB.Label lblPowerOptions 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   680
         Left            =   0
         TabIndex        =   90
         Top             =   960
         Visible         =   0   'False
         Width           =   4400
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPower 
         Height          =   900
         Index           =   4
         Left            =   3600
         Tag             =   "log off"
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":E94F5
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":EBABB
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPower 
         Height          =   900
         Index           =   3
         Left            =   2700
         Tag             =   "sleep"
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":EBAD3
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":EF92B
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPower 
         Height          =   900
         Index           =   2
         Left            =   1800
         Tag             =   "exit"
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":EF943
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":F29F6
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPower 
         Height          =   900
         Index           =   1
         Left            =   900
         Tag             =   "restart"
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":F2A0E
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":F5E8F
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pPower 
         Height          =   900
         Index           =   0
         Left            =   0
         Tag             =   "shutdown"
         Top             =   0
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1588
         Image           =   "frmMain.frx":F5EA7
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":F9512
      End
      Begin VB.Shape sPowerOptions 
         BackColor       =   &H004B4B4B&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H004B4B4B&
         Height          =   610
         Left            =   0
         Top             =   910
         Width           =   5510
      End
   End
   Begin VB.PictureBox pCharmsHolder 
      BackColor       =   &H00111111&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   14400
      Begin VB.Timer tmrCharmsOpen 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3060
         Top             =   120
      End
      Begin VB.Timer tmrCharmsClose 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2640
         Top             =   120
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   3
         Left            =   6900
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":F952A
         Attr            =   513
         Effects         =   "frmMain.frx":FBD8F
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   2
         Left            =   5940
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":FBDA7
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":FD457
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   1
         Left            =   5040
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":FD46F
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":FF469
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   4
         Left            =   7860
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":FF481
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":10022A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   5
         Left            =   8820
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":100242
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":101C0E
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   6
         Left            =   9780
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":101C26
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":102A32
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pCharm 
         Height          =   600
         Index           =   0
         Left            =   4140
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmMain.frx":102A4A
         Settings        =   19200
         Attr            =   513
         Effects         =   "frmMain.frx":104786
      End
   End
   Begin VB.Label lblShutdownTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   12300
      TabIndex        =   55
      Top             =   6900
      Visible         =   0   'False
      Width           =   1500
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pLockFade 
      Height          =   915
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   1614
      Image           =   "frmMain.frx":10479E
      Render          =   4
      Attr            =   513
      Effects         =   "frmMain.frx":105D3B
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pPowerBack 
      Height          =   8100
      Left            =   0
      Top             =   -8
      Visible         =   0   'False
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   14288
      Image           =   "frmMain.frx":105D53
      Settings        =   1048576
      Render          =   4
      Attr            =   513
      Effects         =   "frmMain.frx":1072F1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents m_cAppCommand As cAppCommand
Attribute m_cAppCommand.VB_VarHelpID = -1

Public ResizeOld As Double
Public ResizeNew As Double


Private Sub cCalendar_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub cCalendar_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next

    Do Until ProcessMessage(cCalendar, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cDevices_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cDevices, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cFilms_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub cFilms_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cFilms, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cGallery_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cGallery, sMessage, pImage) = True
        DoEvents
    Loop

End Sub


Private Sub cGames_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub cGames_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cGames, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cGoogle_Message(Index As Integer, sMessage As String)

    frmSplash.lblProgress.Caption = sMessage
    
End Sub

Private Sub cGoogle_UserLoggedIn(Index As Integer, Name As String, Photo As String, User As String, Cover As String, AuthorizationCode As String, NewRefreshToken As String)

    On Error Resume Next
    
    Dim sMessage As String
    Dim pPhoto As GDIpImage
    
    authCode = AuthorizationCode

    refreshToken = NewRefreshToken
    
    xmlSettingsDoc.selectSingleNode("//user/google").Attributes.getNamedItem("refreshToken").Text = refreshToken
    
    xmlSettingsDoc.save App.Path & "\settings.xml"
    
    sGoogleName = Name
    sGooglePhoto = Photo
    sGoogleCover = Cover
    
    Set pPhoto = LoadPictureGDIplus(Photo)
    sMessage = "userloggedin##" & authCode & "##" & Name & "##" & "Google account" 'User
    CentralMessage sMessage, pPhoto
    pUser.Picture = pPhoto
    
    Set pPhoto = LoadPictureGDIplus(Cover)
    CentralMessage "usercover", pPhoto
    
    lblMenuHeader(0).Caption = Name
    lblMenuHeader(1).Caption = "Google account" 'User

    pUserCover.FastRedraw = False
    pUserCover.Picture = pPhoto
    pUserCover.Effects.CreateBlurEffect 25, False
    pUserCover.FastRedraw = True

    lblMenuHeader(0).Left = (pUser.Left * 2) + pUser.Width
    lblMenuHeader(1).Left = lblMenuHeader(0).Left
            
    SetServiceIcon 3, sOn
    
    bLoggedIn = True
    
    Notify "Google Logged In | " & Name, cGoogle(Index), , , False, pServiceIcon(3).Picture
    
    lblMenu(4).ForeColor = &HE0E0E0
    
    'cCalendar.SendMessage "getschedule"
            
End Sub

Private Sub cGoogle_UserLoggedOut(Index As Integer, User As String, AuthorizationCode As String)

    On Error Resume Next
    
    authCode = AuthorizationCode

    CentralMessage "userloggedout##" & authCode

    SetServiceIcon 3, sOff
    
    lblMenuHeader(0).Left = pUser.Left
    lblMenuHeader(1).Left = lblMenuHeader(0).Left
    
    pUser.Picture = Nothing
            
    bLoggedIn = False
    
    Notify "Google Logged Out: | " & User, cGoogle(Index), , , False, pServiceIcon(3).Picture
    
    Unload cGoogle(Index)
    
    'lblMenu(3).ForeColor = &H808080
    lblMenu(4).ForeColor = &H808080
    
End Sub

Private Sub cKeyboard_GoUp()

    On Error Resume Next
    
    Do Until CloseSearch = True
        DoEvents
    Loop

End Sub

Private Sub cKeyboard_SearchCommand(SearchString As String)

    On Error Resume Next
    
End Sub

Private Sub cKeyboard_StringChange(sSearchString As String)

    On Error Resume Next
    
    lblSearch.ForeColor = cRowTitleSelected

    lblSearch.Caption = sSearchString

End Sub

Private Sub cLastFM_Authenticated(Index As Integer)

    Notify "Last.FM Logged In | " & sLastFMUser, cLastFM(Index), , , False, pServiceIcon(4).Picture

End Sub

Private Sub cLastFM_AuthenticationFailed(Index As Integer)
    
    On Error Resume Next

    bLastFMLoggedIn = False

    bLastFmResponded = True
    
    Unload cLastFM(Index)
    
    Notify "Last.FM Logged Out", cLastFM(Index), , , False, pServiceIcon(4).Picture
    
    cMusicHome.SendMessage "LastFM##logout"
    
    SetServiceIcon 4, sOff
    
End Sub

Private Sub cLastFM_AuthenticationStarted(Index As Integer)

    On Error Resume Next
    
    SetServiceIcon 4, sPreparing
    
End Sub

Private Sub cLastFM_GotFocus(Index As Integer)

    On Error Resume Next

    pFocus.SetFocus

End Sub

Private Sub cLastFM_NowPlayingUpdated(Index As Integer)

    On Error Resume Next

    SetServiceIcon 4, sOn

End Sub

Private Sub cLastFM_Scrobbled(Index As Integer)

    On Error Resume Next

    Notify "Scrobbled | " & PlayBack.Artist & " | " & PlayBack.Title, cLastFM(Index), , , False, pServiceIcon(4).Picture

    cMusicHome.SendMessage "nowplayingscrobbled"

    SetServiceIcon 4, sOn

End Sub

Private Sub cLastFM_UserLoggedIn(Index As Integer, sUser As String, sKey As String)

    On Error Resume Next

    bLastFMLoggedIn = True

    bLastFmResponded = True
    
    Notify "Last.FM Logged In | " & sUser, cLastFM, , , False, pServiceIcon(4).Picture
    
    cMusicHome.SendMessage "LastFM##login##" & sUser
    
    SetServiceIcon 4, sOn

End Sub

Private Sub cMusicHome_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub cMusicHome_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cMusicHome, sMessage, pImage) = True
        DoEvents
    Loop
    
End Sub


Private Sub cNews_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next

    Do Until ProcessMessage(cNews, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cNews_Publish(sTitle As String, sThumb As String, sLink As String)

    On Error Resume Next
    
    If lNewsTile(0) <> -1 And CurrentScreen = Home Then
        
        'cRow(lNewsTile(0)).TileSet lNewsTile(1), sTitle, "BBC News", sThumb, , CStr(lNewsTile(1))
    
        Do Until UpdateTile(lNewsTile(0), lNewsTile(1), "app##cNews##4", sTitle, "BBC News", sThumb, False) = True
            DoEvents
        Loop
            
        If (lRowSelected = lNewsTile(0)) And (cRow(lNewsTile(0)).TileCurrent = lNewsTile(1) And bDashLoaded = True And pMenuHolder.Visible = False) Then
            cRow(lNewsTile(0)).TileSelected(lNewsTile(1)) = True
        End If
        
    End If
    
End Sub


Private Sub cNotificationWidget_Click()

    Do Until CloseLock = True
        DoEvents
    Loop

End Sub

Private Sub cNotificationWidget_GotFocus()

    pFocus.SetFocus

End Sub


Private Sub Command1_Click()

    Notify "test"

End Sub

Private Sub cPinItems_TileClick(TileIndex As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        If TileIndex <> cPinItems.TileCurrent Then
        
            cPinItems.TileSelected(CLng(TileIndex)) = True
            
        Else
                
            KeyPress vbKeyReturn, 0
            
        End If

    End If

    pFocus.SetFocus

End Sub

Private Sub cPinItems_TileSelectFalse()

    On Error Resume Next
    
    Sound "listend"

End Sub

Private Sub cPinItems_TileSelectTrue()

    On Error Resume Next
    
    Sound "select"

End Sub

Private Sub cPlaylist_TileClick(TileIndex As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        If TileIndex <> cPlaylist.TileCurrent Then
        
            cPlaylist.TileSelected(CLng(TileIndex)) = True
            
        Else
                
            KeyPress vbKeyReturn, 0
            
        End If

    End If

    pFocus.SetFocus

End Sub


Private Sub cRow_GotFocus(Index As Integer)

    On Error Resume Next

    pFocus.SetFocus

End Sub

Private Sub cPlaylist_TileSelectFalse()

    On Error Resume Next
    
    Sound "listend"

End Sub

Private Sub cPlaylist_TileSelectTrue()

    On Error Resume Next
    
    Sound "select"

End Sub

Private Sub cRow_TileSelectFalse(Index As Integer)

    On Error Resume Next
    
    Sound "listend"

End Sub

Private Sub cRow_TileSelectTrue(Index As Integer)

    On Error Resume Next
    
    Sound "select"

End Sub

Private Sub cSearch_TileSelectFalse()

    On Error Resume Next
    
    Sound "listend"

End Sub

Private Sub cSearch_TileSelectTrue()

    On Error Resume Next
    
    Sound "select"

End Sub

Private Sub cServer_ConnectionClientReady()

    pServiceIcon(1).Animate lvicAniCmdStop
    
    SetClientTimeout False
    
    bRemoteSendReady = True

End Sub

Private Sub cServer_ConnectionClosed()
    
    bRemoteConnected = False
    
    bRemoteSendReady = False
    
    SetClientTimeout False
    
    Notify "Remote Connection Closed", , , True, , pServiceIcon(1).Picture

    SetServiceIcon 1, sOff

    tmrRemote.Enabled = False

    lstRemoteQueue.Clear

End Sub

Private Sub cServer_ConnectionDataArrival(bytesTotal As Long, sData As String)

    pServiceIcon(1).Animate lvicAniCmdStart
    
    SetClientTimeout False
    
    RemoteCommand sData

End Sub

Private Sub cServer_ConnectionDataComplete()

    pServiceIcon(1).Animate lvicAniCmdStop

    bRemoteSendReady = True

End Sub

Private Sub cServer_ConnectionDataSending(bytesSent As Long, bytesRemaining As Long)

    pServiceIcon(1).Animate lvicAniCmdStart

End Sub

Private Sub cServer_ConnectionEstablished()

    Do Until RemoteConnected = True
        DoEvents
    Loop

End Sub

Private Sub cServer_ConnectionRequested(requestID As Long)

    SetServiceIcon 1, sPreparing

End Sub

Private Sub cSport_GotFocus()

    On Error Resume Next

    pFocus.SetFocus

End Sub

Private Sub cSport_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next

    Do Until ProcessMessage(cSport, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cSport_Publish(sTitle As String, sThumb As String, sLink As String)

    On Error Resume Next

    If lSportTile(0) <> -1 And CurrentScreen = Home Then
    
        'cRow(lSportTile(0)).TileSet lSportTile(1), sTitle, "BBC Sport", sThumb, , CStr(lSportTile(1))
    
        Do Until UpdateTile(lSportTile(0), lSportTile(1), "app##cSport##5", sTitle, "BBC Sport", sThumb, False) = True
            DoEvents
        Loop
            
        If (lRowSelected = lSportTile(0)) And (cRow(lSportTile(0)).TileCurrent = lSportTile(1) And bDashLoaded = True) Then
            cRow(lSportTile(0)).TileSelected(lSportTile(1)) = True
        End If
        
    End If

End Sub

Private Sub cTV_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub cTV_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next
    
    Do Until ProcessMessage(cTV, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cWeather_GotFocus()

    On Error Resume Next

    pFocus.SetFocus

End Sub

Private Sub cWeather_Message(sMessage As String, pImage As LaVolpeAlphaImg.GDIpImage)

    On Error Resume Next

    Do Until ProcessMessage(cWeather, sMessage, pImage) = True
        DoEvents
    Loop

End Sub

Private Sub cWeather_Publish(sTitle As String, sThumb As String, sLink As String)

    'On Error Resume Next
    
    If InStr(1, LCase(sTitle), "unknown") = 0 Then
            
        Dim strSplit() As String
        
        strSplit() = Split(sThumb, "##")
        
        sWeatherString = "WEATHER##" & sTitle & "##" & strSplit(0) & "##" & strSplit(1)
                    
        Do Until RemoteSend(sWeatherString) = True
            DoEvents
        Loop
        
        Do Until Notify(sTitle, cWeather, , True, , LoadPictureGDIplus(strSplit(0))) = True
            DoEvents
        Loop
        
        sWeatherStatus = sTitle
        
        lblLockWeather.Caption = sTitle
    
        If bWeatherBack = True Then
            SetBackground LoadPictureGDIplus(strSplit(1)), pFocus, 25, 50
        End If
        
        If PlayBack.Source <> audio Then
        
            SetBackground LoadPictureGDIplus(strSplit(1)), pLockHolder, 0, 0
            
        End If
        
        'SetLockBack LoadPictureGDIplus(strSplit(1)), False
        pLockHolder.Tag = strSplit(1)
    
        If lWeatherTile(0) <> -1 Then
            
            Do Until UpdateTile(lWeatherTile(0), lWeatherTile(1), "app##cWeather##3", sTitle, "Weather for " & sLink, strSplit(0), False) = True
                DoEvents
            Loop
            
            If (lRowSelected = lWeatherTile(0)) And (cRow(lWeatherTile(0)).TileCurrent = lWeatherTile(1) And bDashLoaded = True) Then
                cRow(lWeatherTile(0)).TileSelected(lWeatherTile(1)) = True
            End If
            
        End If
            
    End If
        
End Sub


Private Sub Form_Resize()

    On Error Resume Next

    ResizeForm

End Sub

Private Function ResizeForm()

    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
    
        On Error Resume Next
    
        Dim lPC As Double
        
        With Me
        
            ResizeNew = .ScaleWidth
            
            lPC = (ResizeNew / ResizeOld) * 100
            
            Dim cCont As Control
            
            For Each cCont In .Controls
            
                ResizeControl cCont, lPC
                
            Next cCont
            
            'END FUNCTION
            
            ResizeOld = .ScaleWidth
        
        End With
        
    End If

End Function

Private Sub ResizeControl(cControl As Control, lPC As Double)

    On Error Resume Next

    If TypeOf cControl Is Timer Then
    
        'do nothing
    
    ElseIf TypeOf cControl Is Line Then
    
        cControl.X1 = cControl.X1 / 100 * lPC
        
        cControl.X2 = cControl.X2 / 100 * lPC
        
        cControl.Y1 = cControl.Y1 / 100 * lPC
        
        cControl.Y2 = cControl.Y2 / 100 * lPC
    
    ElseIf TypeOf cControl Is PlayerControl Then
    
        'do nothing
        
    Else
    
        cControl.Width = cControl.Width / 100 * lPC
        
        cControl.Height = cControl.Height / 100 * lPC
        
        cControl.Left = cControl.Left / 100 * lPC
        
        cControl.Top = cControl.Top / 100 * lPC
        
        If TypeOf cControl Is Label Then
        
            cControl.FontSize = cControl.FontSize / 100 * lPC
        
        ElseIf TypeOf cControl Is TextBox Then
        
            cControl.FontSize = cControl.FontSize / 100 * lPC
            
        End If
        
    End If

End Sub

Private Sub cNews_GotFocus()

    On Error Resume Next

    pFocus.SetFocus

End Sub


Private Sub Form_GotFocus()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub Form_Load()
   
    'On Error Resume Next

    Do Until InitMain = True
        DoEvents
    Loop

    'appcommand
        
        Set m_cAppCommand = New cAppCommand
        m_cAppCommand.Attach frmMain.hwnd
            
        bRemoteSupport = True
        
    'settings
    
        Do Until LoadSettings = True
            DoEvents
        Loop
        
    'checknetwork
    
        Do Until CheckNetwork = True
            DoEvents
        Loop
    
    'filetypes
    
        Do Until LoadFileTypes = True
            DoEvents
        Loop
        
    'playback
        
        Do Until InitPlayback = True
            DoEvents
        Loop
        
    'dash
    
        lRowSelected = 0
        lTileSelected = 0
        bShowWelcome = True
        
    'sounds
    
        Do Until InitSounds = True
            DoEvents
        Loop
        
    'controllers
    
        If bGameControllerSupport = True Then

            tmrXbox.Enabled = True
            
        End If
        
    'sync
        
        ReDim SyncItems(0)
        
    'ready to show main screen
    
        'Startup
        
            Select Case LCase(command())
        
                Case "music"
                    
                    Do Until AppDirect(cMusicHome) = True
                        DoEvents
                    Loop
                    
                Case "artists"
                    
                    Do Until AppDirect(cMusicHome, "initartists") = True
                        DoEvents
                    Loop
                    
                Case "albums"
                    
                    Do Until AppDirect(cMusicHome, "initalbums") = True
                        DoEvents
                    Loop
                    
                    
                Case "radio"
                    
                    Do Until AppDirect(cMusicHome, "initradio") = True
                        DoEvents
                    Loop
                    
                Case "smartmixes"
                    
                    Do Until AppDirect(cMusicHome, "initmmartmixes") = True
                        DoEvents
                    Loop
                    
                Case "films"
                
                    Do Until AppDirect(cFilms) = True
                        DoEvents
                    Loop
                    
                Case "tv"
                
                    Do Until AppDirect(cTV) = True
                        DoEvents
                    Loop
                    
                Case "games"
                
                    Do Until AppDirect(cGames) = True
                        DoEvents
                    Loop
                    
                Case "news"
                
                    Do Until AppDirect(cNews) = True
                        DoEvents
                    Loop
                    
                Case Else
                
                    tmrStartup.Enabled = True
                    
                    Do Until tmrStartup.Enabled = False
                        DoEvents
                    Loop
            
            End Select
            
        'central message network
            
            If bNetConnection = True Then
                CentralMessage "netconnected"
            Else
                CentralMessage "netdisconnected"
            End If
            
        'calendar
        
            ReDim Events(0)
            Events(0).Start = "99:99"
            Events(0).End = "99:99"
        
        'hook mousewheel
        
            'Call WheelHook(Me.hwnd)
            
        'keep screen awake
                    
            SetThreadExecutionState (ES_DISPLAY_REQUIRED Or ES_CONTINUOUS)

        'begin mousecount
            
            lMouseCount = 0
    
            tmrMouse.Enabled = True
    
            tmrMouseOut.Enabled = bMouseOut
                
End Sub

Private Sub Form_Terminate()
    
    On Error Resume Next
    
    Call WheelUnHook(Me.hwnd)
    
    StopPlayback
    
    Do Until HideCursor(False) = True
        DoEvents
    Loop

    End

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Call WheelUnHook(Me.hwnd)
    
    cServer.ConnectionClose
    
    AllowMonitorSleeping
    
    StopPlayback
    
    CloseSounds
    
    frmMain.tmrMouse.Enabled = False
    
    Do Until HideCursor(False) = True
        DoEvents
    Loop

    Cancel = 0

End Sub


Private Sub lblDateTime_Change()

    On Error Resume Next
    
    lblLockTime.Caption = lblDateTime.Caption
    lblShutdownTime.Caption = lblDateTime.Caption

End Sub

Private Sub lblLockTime_Change()

    pLockIcon.Left = lblLockTime.Left - pLockIcon.Width - (pLockIcon.Width / 4)
    
    'lblLockNotifCount.Left = pLockIcon.Left - lblLockNotifCount.Width - (pLockIcon.Width / 4)

End Sub

Private Sub lblLockTime_Click()

    On Error Resume Next
    
    pLockHolder.SetFocus

End Sub

Private Sub lblPowerOptions_Click()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub lblStatus_Change()

    'RemoteSend "STATUS##" & lblStatus.Caption

End Sub

Private Sub lblTilePinOptions_Click(Index As Integer)

    On Error Resume Next
    
    pFocus.SetFocus

End Sub



Private Sub pCharm_Click(Index As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        Do Until ProcessCharm(CLng(Index)) = True
            DoEvents
        Loop
        
    End If

    'pFocus.SetFocus

End Sub

Private Sub pCharm_MouseEnter(Index As Integer)

    On Error Resume Next
    
    pCharm(lCharmSelected).TransparencyPct = 75
    
    pCharm(Index).TransparencyPct = 0
    
    lCharmSelected = Index
    
End Sub

Private Sub pExit_Click()

    On Error Resume Next
    
    If bFocus = True Then
    
        BeginCloseExe
        
    End If

End Sub


Private Sub pFocus_GotFocus()

    On Error Resume Next

    bFocus = True

    'pFocus.BorderStyle = 1

End Sub

Private Sub pFocus_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    Do Until KeyPress(KeyCode, Shift) = True
        DoEvents
    Loop
    
End Sub


Private Sub pFocus_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    
        Case 4 'Ctrl+D
        
            Do Until KeyPress(vbKeyMenu, 0, True) = True
                DoEvents
            Loop
            
        Case 35 'hash
            
            Do Until KeyPress(vbKeyShift, 0) = True
                DoEvents
            Loop
        
        Case 42 'star
            
            Do Until KeyPress(vbKeyShift, 0) = True
                DoEvents
            Loop
        
    End Select

End Sub

Private Sub pFocus_LostFocus()

    On Error Resume Next

    bFocus = False

    'pFocus.BorderStyle = 0

End Sub

Private Sub Picture2_Click()

End Sub

Private Sub pLockHolder_Click()

    Do Until CloseLock = True
        DoEvents
    Loop
       
End Sub

Private Sub pLockHolder_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    KeyPress KeyCode, Shift

End Sub

Private Sub pMinVidControl_Click(Index As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        Select Case Index
        
            Case 0
            
                PlayPause
            
            Case 1
            
                PlaylistNext
            
            Case 2
            
                PlaylistBack
                
            Case 3
            
                Do Until MinRestoreVideo = True
                    DoEvents
                Loop
            
        End Select
        
    End If
    
    pFocus.SetFocus

End Sub

Private Sub pPower_Click(Index As Integer)
    
    On Error Resume Next
    
    'If bFocus = True Then

    '    ProcessPower CLng(Index)

    'End If

End Sub

Private Sub pPower_MouseEnter(Index As Integer)

    'On Error Resume Next
    
    'pPower(lPower).TransparencyPct = 75
    
    'SetPowerCaption CLng(Index)
    
    'pPower(Index).TransparencyPct = 0
    
    'lPower = Index
    
End Sub


Private Sub pTilePinSelectHolder_Click()

    On Error Resume Next
    
    pFocus.SetFocus

End Sub

Private Sub pTilePinSelectHolderBack_Click()

    On Error Resume Next
    
    pFocus.SetFocus
    
    pTilePinSelectHolder.Visible = False

End Sub

Private Sub pVidControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    If bFocus = True Then
    
    Do Until pVidControlClick(Index) = True
        DoEvents
    Loop
        
    End If
        
    pFocus.SetFocus

End Sub


Private Sub pVidMenu_GotFocus()

    On Error Resume Next
    
    lOSD = 0

    ClosePlaylist
    
    pFocus.SetFocus

End Sub

Private Sub pVidPosition_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    wmp.Controls.currentPosition = (wmp.currentMedia.Duration * (x / pVidPosition(0).Width))
        
    lOSD = 0

    pFocus.SetFocus

End Sub

Private Sub Timer1_Timer()

    

End Sub



Private Sub scrEQ_Change(Index As Integer)

    If cPlayer.UBound > 0 Then
    
        cPlayer(1).EQSetChannel CLng(Index), CLng(scrEQ(Index).Tag), CLng(scrEQ(Index).Value)
    
        lblEQValue(Index).Caption = scrEQ(Index).Value
    
    End If

End Sub

Private Sub scrEQ_Scroll(Index As Integer)

    If cPlayer.UBound > 0 Then
    
        cPlayer(1).EQSetChannel CLng(Index), CLng(scrEQ(Index).Tag), CLng(scrEQ(Index).Value)
    
        lblEQValue(Index).Caption = scrEQ(Index).Value
    
    End If

End Sub

Private Sub SysInfo_ConfigChangeCancelled()

End Sub

Private Sub tmrCharm_Timer()

    On Error Resume Next
    
    lCharmCount = lCharmCount + 1
    
    If lCharmCount = 1 And pCharmsHolder.Visible = False Then
    
        OpenCharms
        
        lCharmCount = 0
        
        tmrCharm.Enabled = False
    
    End If

End Sub

Private Sub tmrCharmsClose_Timer()

    On Error Resume Next

    Static i As Integer
    
    pFocus.Top = pFocus.Top - (pCharmsHolder.Height / 4)
    
    i = i + 1
    
    If i = 4 Then
    
    pFocus.Top = 0
    
        i = 0
    
        tmrCharmsClose.Enabled = False
        
    End If
End Sub

Private Sub tmrCharmsOpen_Timer()

    On Error Resume Next

    Static i As Integer
    
    pFocus.Top = pFocus.Top + (pCharmsHolder.Height / 4)
        
    i = i + 1
    
    If i = 4 Then
    
        i = 0
    
        tmrCharmsOpen.Enabled = False
        
    End If

End Sub

Private Sub tmrCheckForProcess_Timer()

    On Error Resume Next
    
    If IsProcessRunning(sProcessPath) = False Then
    
        If bProcess = False Then
        
            EndProcessCheck
            
        Else
        
            bProcess = False
            
        End If
    
    Else
    
        bProcess = True
        
    End If

End Sub

Private Sub tmrClientTimeout_Timer()

    lClientTimeout = lClientTimeout + 1
    
    If lClientTimeout = 10 Then
    
        'cServer.ConnectionClose
        
        SetClientTimeout False
        
    End If

End Sub

Private Sub tmrCloseLockHolder_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    pLockHolder.Top = pLockHolder.Top - (pLockHolder.Height / 5)
    
    If i = 5 Then
    
        i = 0
        
        pLockHolder.Visible = False
        pLockHolder.Top = 0
        Set pLockHolder.Picture = Nothing
        
        tmrCloseLockHolder.Enabled = False
    
    End If
    
End Sub

Private Sub tmrCloseSearch_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    sSearchBack.Width = sSearchBack.Width - pSearchWidth.Width
    
    If i = 5 Then
    
        i = 0
        
        tmrCloseSearch.Enabled = False
        
    End If
    
End Sub

Private Sub tmrDateTime_Timer()

    On Error Resume Next

    Dim sDate As String
    Dim sTime As String
    Dim sDateTime As String
    Dim i As Long
    
    sDate = WeekdayName(Weekday(Now), , vbSunday) & ", " & MonthName(Month(Now)) & " " & Day(Now)

    Select Case Day(Now)
    
        Case 1, 21, 31
        
            sDate = sDate & "st "
        
        Case 2, 22
        
            sDate = sDate & "nd "
            
        Case 3, 23
        
            sDate = sDate & "rd "
            
        Case 4 To 20, 24 To 30
    
            sDate = sDate & "th "
            
    End Select
    
    sDate = sDate & Year(Now) & " "
    
    sTime = Mid(Time, 1, 5)
    
    'sDateTime = sDate & "| " & sTime
    
    sDateTime = sTime

    lblDateTime.Caption = sDateTime
    
    lblHomeTime.Caption = sDateTime
    
    'For i = 0 To UBound(Events())
        
    '    DoEvents
        
    '    If QueryNotifyEvent(i) = True Then
            
    '        Notify "Event | " & Events(i).Name & " | " & Events(i).Start & " - " & Events(i).End, , , True
            
    '        Events(i).Notified = True
        
    '    End If
    
    'Next

End Sub

Private Sub tmrMouse_Timer()

    On Error Resume Next
    
    Dim Point As PointAPI
    Dim bTemp As Boolean
    Dim bOnScreen As Boolean
    
    GetCursorPos Point
    
    'is cursor on form
        If (Point.x > ((Me.Left / 15) - 1) And Point.x < ((Me.Left / 15) + Me.ScaleWidth) + 1) And (Point.y > ((Me.Top / 15) - 1) And Point.y < ((Me.Top / 15) + Me.ScaleHeight) + 1) Then
            bMouseOnForm = True
        Else
            bMouseOnForm = False
        End If
        
    If bMouseOnForm = True And bFocus = True Then 'ignore functions if mouse is not on form or form is not focused
        
        'is cursor on video player
            If Point.x > pVideoHolder.Left And Point.x < pVideoHolder.Width Then
                If Point.y > pVideoHolder.Top And Point.y < pVideoHolder.Height Then
                    bOnScreen = True
                Else
                    bOnScreen = False
                End If
            Else
                bOnScreen = False
            End If
    
        'charm mouse functions
            If Point.y > ((Me.Top / 15) - 1) And Point.y < ((Me.Top / 15) + (pTopBar.Height / 4)) And pCharmsHolder.Visible = False And pVideoHolder.Visible = False And bMouseCharm = True Then
                If tmrCharm.Enabled = False Then
                    tmrCharm.Enabled = True
                End If
            Else
                lCharmCount = 0
                tmrCharm.Enabled = False
                If pCharmsHolder.Visible = True Then
                    If Point.y > ((Me.Top / 15) + pCharmsHolder.Height) Then
                        If (GetAsyncKeyState(vbKeyLButton) And &H8000) < 0 Then
                            CloseCharms
                        End If
                    End If
                End If
            End If
    
        'menu mouse functions
            If pMenuHolder.Visible = True Then
                If (Point.x > ((Me.Left / 15) + pMenuHolder.Width) And Point.x < ((Me.Left / 15) + Me.ScaleWidth)) And ((Point.y > (Me.Top / 15)) And (Point.y < ((Me.Top / 15) + Me.ScaleHeight))) Then
                    If (GetAsyncKeyState(vbKeyLButton) And &H8000) < 0 Then
                        Do Until CloseMenu = True
                            DoEvents
                        Loop
                    End If
                End If
            End If
            
        'playlist mouse functions
            If pPlaylistHolder.Visible = True Then
                If Not (Point.x > ((Me.Left / 15) + pPlaylistHolder.Left) And Point.x < ((Me.Left / 15) + (pPlaylistHolder.Left + pPlaylistHolder.Width)) And Point.y > ((Me.Top / 15) + pPlaylistHolder.Top) And Point.y < ((Me.Top / 15) + (pPlaylistHolder.Top + pPlaylistHolder.Height))) Then
                    If (GetAsyncKeyState(vbKeyLButton) And &H8000) < 0 Then
                        ClosePlaylist
                    End If
                End If
            End If
        
        'hide mouse?
            
            If Point.x < MouseX - 5 Then
            
                bTemp = True
                
            ElseIf Point.x > MouseX + 5 Then
                
                bTemp = True
                    
            ElseIf Point.y < MouseY - 5 Then
                    
                bTemp = True
                        
            ElseIf Point.y > MouseY + 5 Then
                        
                bTemp = True
                        
            Else
                        
                bTemp = False
                            
            End If
                    
            If bTemp = True Then
            
                HideCursor False
            
                lMouseCount = 0
                
                tmrMouseOut.Enabled = bMouseOut
                
                If bOnScreen = True Then
                
                    ShowOSD bOSDAvailable
                
                End If
               
            End If
        
    End If

    MouseX = Point.x
    MouseY = Point.y

    SetThreadExecutionState (ES_DISPLAY_REQUIRED Or ES_CONTINUOUS)

End Sub



Private Sub tmrMouseOut_Timer()

    On Error Resume Next
    
    If bMouseOnForm = True Then

        lMouseCount = lMouseCount + 1
    
    Else
    
        lMouseCount = 0
        
    End If
    
    If lMouseCount = 5 Then
    
        HideCursor True
        
        tmrMouseOut.Enabled = False
    
    End If

End Sub

Private Sub tmrNotify_Timer()

    On Error Resume Next
    
    lNotify = lNotify + 1
    
    If lNotify = 5 Then
    
        lNotify = 0
        
        bNotify = False
        bStatusVisible = False
        
        pNotifyHolder.Visible = False
        
        lblStatus.Caption = sWeatherStatus
        
        Do Until RemoteSend("NOTIFY##" & lblStatus.Caption) = True
            DoEvents
        Loop
        
        tmrNotify.Enabled = False
        
    End If

End Sub

Private Sub tmrOpenMenu_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    pMenuHolder.Visible = True
    
    pMenuHolder.Width = pMenuHolder.Width + (pMenuWidth.Width / 5)
    
    If i = 5 Then
    
        i = 0
        
        pMenuHolder.Left = pSideFade.Width '0
        
        pMenuHolder.Width = pMenuWidth.Width
        
        tmrOpenMenu.Enabled = False
        
    End If
    
End Sub

Private Sub tmrCloseMenu_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    pMenuHolder.Width = pMenuHolder.Width - (pMenuWidth.Width / 5)
    
    If i = 5 Then
    
        i = 0
        
        pMenuHolder.Width = 1
        
        pMenuHolder.Visible = False
        
        tmrCloseMenu.Enabled = False
        
    End If
    
End Sub

Private Sub tmrOpenSearch_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    sSearchBack.Width = sSearchBack.Width + pSearchWidth.Width
    
    If i = 5 Then
    
        i = 0
        
        tmrOpenSearch.Enabled = False
        
    End If
    
End Sub

Private Sub tmrOSD_Timer()

    On Error Resume Next
    
    lOSD = lOSD + 1
    
    Static i As Integer
    
    
    If pPlaylistHolder.Visible = False Then
        If bMinVideo = True Then
            pMiniOSD.ZOrder 0
        Else
            pVidMenu.ZOrder 0
        End If
    End If
    
    If lOSD = 250 Then
    
        pVidMenu.Top = pVidMenu.Top + (pVidMenu.Height / 5)
        
    ElseIf lOSD = 251 Then
    
        pVidMenu.Top = pVidMenu.Top + (pVidMenu.Height / 5)
        
    ElseIf lOSD = 252 Then
    
        pVidMenu.Top = pVidMenu.Top + (pVidMenu.Height / 5)
        
    ElseIf lOSD = 253 Then
    
        pVidMenu.Top = pVidMenu.Top + (pVidMenu.Height / 5)
        
    ElseIf lOSD = 254 Then
    
        pVidMenu.Top = pVidMenu.Top + (pVidMenu.Height / 5)
        
    ElseIf lOSD = 255 Then
    
        lOSDIndex = 0
    
        pMiniOSD.Visible = False
    
        pVidMenu.Visible = False
    
        tmrOSD.Enabled = False
    
        For i = 0 To frmMain.pVidControl.UBound
            frmMain.pVidControl(i).TransparencyPct = 75
            frmMain.pVidControl(i).BackStyleOpaque = False
        Next
    
        frmMain.pVidControl(3).TransparencyPct = 0
        frmMain.pVidControl(3).BackStyleOpaque = True
        
        lOSDIndex = 3
        
    End If

End Sub

Private Sub tmrPlayback_Timer()

    On Error Resume Next
    
    Dim sPosString As String
    Dim sStatusOld As String

    sStatusOld = sPosString

    Select Case PlayBack.Source
    
        Case 0 'audio
        
            Select Case Playlist(PlayBack.PlaylistIndex).SubSource
            
                Case 0 'local file
                
                    PlayBack.Position = cPlayer(1).FilePosition
                    PlayBack.Duration = cPlayer(1).FileDuration
                
                    cMusicHome.SendMessage "nowplayingposition##" & CStr(PlayBack.Position) & "##" & CStr(PlayBack.Duration) & "##" & PlayBack.Shuffle & "##" & PlayBack.Repeat, Nothing
                    
                    sPosString = cPlayer(1).FilePositionString & " / " & cPlayer(1).FileDurationString
                
                    lblLockArt(1).Caption = sPosString
                                        
                    PlayBack.State = ConvertPlayerStatus(cPlayer(1).PlayerStatus)
                    
                    If PlayBack.State = 0 Then 'stopped
                        sPosString = "Stopped | " & sPosString
                    ElseIf PlayBack.State = 1 Then 'playing
                        sPosString = "Playing | " & sPosString
                    ElseIf PlayBack.State = 2 Then 'paused
                        sPosString = "Paused | " & sPosString
                    ElseIf PlayBack.State = 3 Then 'paused
                        sPosString = "Buffering | " & sPosString
                    End If
                    
                    If bNotify = False Then
                        If PlayBack.NowPlayingShow = True Then
                            lblStatus.Caption = sPosString & " | " & PlayBack.Title
                            If PlayBack.Artist <> "" Then
                                lblStatus.Caption = lblStatus.Caption & " | " & PlayBack.Artist
                            End If
                        Else
                            lblStatus.Caption = ""
                        End If
                    End If
                    
                    If CLng(PlayBack.Position) = 30 Then
                        If bLastFMLoggedIn = True And PlayBack.Scrobbled = False Then
                            PlayBack.Scrobbled = True
                            cLastFM(1).Scrobble PlayBack.Artist, PlayBack.Title, PlayBack.Album
                        End If
                    End If
                    
                    cRow(lMusicRow).TileSubCaption(0) = "Now Playing | " & cPlayer(1).FilePositionString & " / " & cPlayer(1).FileDurationString
                    
                    cRow(lMusicRow).TileCaption(0) = PlayBack.Title
                    
                    If PlayBack.Position >= PlayBack.Duration - 0.1 Then
                        PlaylistNext
                    End If
                    
                Case 1 'streaming file
                
                Case 2 'cd
                
                Case 3 'net radio
                
                    PlayBack.Album = Trim(cPlayer(1).StreamSong)
                    If InStr(1, LCase(cPlayer(1).StreamBPS), "http") = 0 Then
                        PlayBack.Artist = Trim(cPlayer(1).StreamBPS) & " kbps"
                    Else
                        PlayBack.Artist = Trim(cPlayer(1).StreamBPS)
                    End If
                    'PlayBack.Title = Trim(cPlayer(1).StreamName)
                    
                    sPosString = "Streaming | " & PlayBack.Artist ' & " | " & PlayBack.Title
                    If bNotify = False And PlayBack.NowPlayingShow = True Then
                        lblStatus.Caption = sPosString
                    End If
                    
                    lblLockArt(0).Caption = PlayBack.Title
                    lblLockArt(1).Caption = sPosString
                    
                    cRow(lMusicRow).TileCaption(0) = PlayBack.Title
                    cRow(lMusicRow).TileSubCaption(0) = "Now " & sPosString
                    
                    cMusicHome.SendMessage "nowplayingupdate##" & PlayBack.Title & "##" & PlayBack.Artist & "##" & PlayBack.Album & "##" & PlayBack.Thumb & "##3", Nothing
               
                    If cPlayer(1).StreamSong <> "" Then
                        'cNowPlaying.SetNowPlaying CurrentArt, CurrentFile, Player(1).StreamName, CurrentAlbum, Player(1).StreamSong, CurrentBitrate, StatusString(Player(1).PlayerStatus), CurrentSpeakerCount, True, , , , RadioRecord
                    Else
                        'cNowPlaying.SetNowPlaying CurrentArt, CurrentFile, Player(1).StreamSong, CurrentAlbum, Player(1).StreamName, CurrentBitrate, StatusString(Player(1).PlayerStatus), CurrentSpeakerCount, True, , , , RadioRecord
                    End If
                        
            End Select
        
            
        Case 1 'video
        
            PlayBack.Duration = wmp.currentMedia.Duration
            PlayBack.Position = wmp.Controls.currentPosition
            
            If PlayBack.Position > 0 Then
                sPosString = wmp.Controls.currentPositionString & " / " & wmp.currentMedia.durationString
                If PlayBack.Position >= PlayBack.Duration - 0.1 Then
                    PlaylistNext
                End If
            Else
                sPosString = "00:00 / " & wmp.currentMedia.durationString
            End If
            
            lblVidDetails(0).Caption = sPosString
            
            If PlayBack.State = 0 Then 'stopped
                sPosString = "Stopped | " & sPosString
            ElseIf PlayBack.State = 1 Then 'playing
                sPosString = "Playing | " & sPosString
            ElseIf PlayBack.State = 2 Then 'paused
                sPosString = "Paused | " & sPosString
            ElseIf PlayBack.State = 3 Then 'paused
                sPosString = "Buffering | " & sPosString
            End If
            
            If bNotify = False Then
                lblStatus.Caption = sPosString & " | " & PlayBack.Title
            End If
            
            lblMinVidPos.Caption = lblVidDetails(0).Caption
            
            pMinVidControl(0).Picture = pVidControl(3).Picture
            
            If PlayBack.Duration > 0 Then
                pVidPosition(1).Width = (pVidPosition(0).Width * (PlayBack.Position / PlayBack.Duration))
            End If

        Case 4 'game
        
        Case 5 'external
        
    End Select
    
    If sPosString <> sStatusOld Then

        Do Until RemoteSend("PLAYBACK##" & sPosString) = True
            DoEvents
        Loop
        
    End If

End Sub

Private Sub tmrProcess_Timer()

    On Error Resume Next

    Static i As Integer
    
    i = i + 1

    If i = 100 Then
    
        If CheckNetConnection = True Then
        
            SetServiceIcon 0, sOn
        
            If bNetConnection = False Then
            
                CentralMessage "NetConnected"
                
                GoogleLogin sGoogleUser, sGooglePass
                
                LastFMLogin sLastFMUser, sLastFMPass
            
            End If
            
            bNetConnection = True
        
        Else
        
            If bNetConnection = True Then
                
                bNetConnection = False
            
                SetServiceIcon 0, sOff
                
                Notify "Connection lost", , , True, False, pServiceIcon(0).Picture
                
                If frmMain.cGoogle.UBound > 0 Then
                    Unload frmMain.cGoogle(1)
                End If
                
                If frmMain.cLastFM.UBound > 0 Then
                    Unload frmMain.cLastFM(1)
                End If
                
                SetServiceIcon 0, sOff
                SetServiceIcon 3, sOff
                SetServiceIcon 4, sOff
                            
                If lWeatherTile(0) <> -1 Then
                                
                    Do Until UpdateTile(lWeatherTile(1), lWeatherTile(1), "app##cNews##4", "", "Weather", "\Images\Dash\Home\0.png", False)
                        DoEvents
                    Loop
                
                End If
                
                If lNewsTile(0) <> -1 Then
                    
                    Do Until UpdateTile(lNewsTile(0), lNewsTile(1), "app##cNews##4", "", "BBC News", "\Images\Dash\Home\10.png", False)
                        DoEvents
                    Loop
                    
                End If
                
                If lSportTile(0) <> -1 Then
                    
                    Do Until UpdateTile(lSportTile(0), lSportTile(1), "app##cNews##4", "", "BBC Sport", "\Images\Dash\Home\11.png", False)
                        DoEvents
                    Loop
                
                End If
                            
                CentralMessage "NetDisconnected"
            
            End If
        
        End If
        
        i = 0
        
    End If
    
End Sub


Private Sub tmrRemote_Timer()

    If bRemoteSendReady = False Then
        
        Do Until bRemoteSendReady = True
            DoEvents
        Loop
        
    End If

    If lstRemoteQueue.ListCount > 0 Then
        
        bRemoteSendReady = False
                
        SetClientTimeout True
                
        cServer.ConnectionSend lstRemoteQueue.List(0)
        
        lstRemoteQueue.RemoveItem 0
    
    End If
    
End Sub

Private Sub tmrRumble_Timer()

    Static i As Integer
    
    i = i + 1
    
    If i = 1 Then
    
        If bDashLoaded = True Then
    
            xBoxControllerRumble 0, 100, 100
        
        End If
        
    ElseIf i = 2 Then
    
        xBoxControllerRumble 0, 0, 0
        
        i = 0
        
        tmrRumble.Enabled = False
        
    End If

End Sub

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


Private Sub tmrShowOSD_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    If pVidMenu.Top > pOsdPos.Top Then
    
        pVidMenu.Top = pVidMenu.Top - (pVidMenu.Height / 5)
        
    Else
    
        tmrShowOSD.Enabled = False
        
    End If
    
End Sub


Private Sub tmrStartup_Timer()

    'On Error Resume Next

    Static i As Integer
    
    i = i + 1
    
    If i = 1 Then
    
        Me.Width = frmSplash.Width
        
    ElseIf i = 9 Then
                    
        Me.Visible = True
        
        Me.SetFocus
    
        If sBackground <> vbNullString Then
            SetBackground LoadPictureGDIplus(sBackground), pFocus, 25, 50
            SetBackground LoadPictureGDIplus(sBackground), pLockHolder, 25, 50
        End If
    
    End If
    
    If bShowWelcome = True Then
    
        If i = 10 Then
        
            lblWelcome.ForeColor = &H404040
            
            lblWelcome.Visible = True
            
            frmSplash.Visible = False
        
        ElseIf i = 11 Then
        
            lblWelcome.ForeColor = &H808080
            
        ElseIf i = 12 Then
        
            lblWelcome.ForeColor = &HC0C0C0
            
        ElseIf i = 13 Then
        
            lblWelcome.ForeColor = &HE0E0E0
            
        ElseIf i = 14 Then
        
            lblWelcome.ForeColor = vbWhite
        
        ElseIf i = 20 Then
        
            lblWelcome.ForeColor = &HE0E0E0
        
        ElseIf i = 21 Then
        
            lblWelcome.ForeColor = &HC0C0C0
        
        ElseIf i = 22 Then
        
            lblWelcome.ForeColor = &H808080
        
        ElseIf i = 23 Then
        
            lblWelcome.ForeColor = &H404040
            
        ElseIf i = 24 Then
            
            lblWelcome.Visible = False
                    
        ElseIf i = 27 Then
        
            If InStr(1, LCase(command), "/l") Then
            
                Do Until ShowLock(True) = True
                    DoEvents
                Loop
                
            End If
            
            If InStr(1, LCase(command), "/a") Then
            
                Dim strSplit() As String
                Dim ctl As Control
    
                strSplit() = Split(sAlarmSource, "##")
        
                Do Until PlaylistClear = True
                    DoEvents
                Loop
                
                Do Until PlaylistAdd(strSplit(2), strSplit(1), audio, 3, sAlarmThumb) = True
                    DoEvents
                Loop
                
                Do Until PlaylistLoad(0) = True
                    DoEvents
                Loop
                
            End If
            
            pTopBar.Visible = bShowTopBar
            
        ElseIf i = 28 Then
                    
            pSideFade.Visible = True
                      
        ElseIf i = 29 Then
                
            lblAppTitle.Visible = True
        
            pAppLogo.Visible = True
        
            'pMenuIcon.Visible = True
                
            'pSearchIcon.Visible = True
                  
        ElseIf i = 30 Then
        
            Do Until InitDashboard = True
                DoEvents
            Loop
            
        ElseIf i = 31 Then
        
            pServiceIcon(0).Visible = True
            
        ElseIf i = 32 Then
        
            pServiceIcon(1).Visible = True
            
        ElseIf i = 33 Then
        
            pServiceIcon(2).Visible = True
            
        ElseIf i = 34 Then
        
            pServiceIcon(3).Visible = True
        
        ElseIf i = 35 Then
        
            pServiceIcon(4).Visible = True
        
            lblDateTime.Visible = True
            
            lblHomeTime.Visible = True
                    
            'init dash items
        
                If bNetConnection = True Then
            
                    'weather
                    
                        cWeather.RequestPublish
            
                    'news
                
                        cNews.RequestPublish
                        
                    'sport
                    
                        cSport.RequestPublish
            
                End If
        
            'startup complete
                       
                tmrProcess.Enabled = True
        
                bShowWelcome = False
                
                lRowSelected = 0
            
                SelectRow CInt(lRowSelected)
                    
                cRow(lRowSelected).TileSelected(0) = True
    
                CurrentScreen = Home
    
                bKey = True
                
                Unload frmSplash
        
                i = 0
                
                pFocus.SetFocus
                
                tmrStartup.Enabled = False
            
        End If
        
    Else
        
        If i = 1 Then
        
            Do Until InitDashboard = True
                DoEvents
            Loop
        
            pServiceIcon(0).Visible = True
        
        ElseIf i = 2 Then
        
            pServiceIcon(1).Visible = True
            
        ElseIf i = 3 Then
        
            pServiceIcon(2).Visible = True
            
        ElseIf i = 4 Then
        
            pServiceIcon(3).Visible = True
            
        ElseIf i = 5 Then
        
            pServiceIcon(4).Visible = True
        
            'dash initialised
        
                bKey = True
            
                i = 0
                
                pFocus.SetFocus
                
                tmrStartup.Enabled = False
        End If
    
    End If

End Sub

Private Sub tmrVolume_Timer()

    On Error Resume Next
    
    lVolumeCount = lVolumeCount + 1
    
    If lVolumeCount = 30 Then
    
        pVolumeHolder.Visible = False
        
        tmrVolume.Enabled = False
        
    End If

End Sub

Private Sub tmrWait_Timer()

    On Error Resume Next

    Static i As Integer
    
    i = i + 1
    
    If i = 2 Then
    
        i = 0
        
        tmrWait.Enabled = False
        
    End If
    
End Sub


Private Sub tmrXbox_Timer()

    On Error Resume Next
    
    If bControllerBypass = False Then
        
        If CountControllers > 0 Then
        
            If CountControllers > lXboxCount Then
            
                'tell controls that an xbox controller is connected
                
                lXboxCount = CountControllers
                
                
            SetServiceIcon 2, sOn
                
                Notify "Xbox Controller Added | " & lXboxCount & " Controllers", tmrXbox, , True, , pServiceIcon(2).Picture
        
            End If
            
            If bKey = True Then
                
                If xboxController(0).ButtonA = 1 Then KeyPress vbKeyReturn, 0
            
                If xboxController(0).ButtonB = 1 Then KeyPress vbKeyBack, 0
                
                If xboxController(0).ButtonY = 1 Then KeyPress vbKeyControl, 0
                
                If xboxController(0).ButtonX = 1 Then KeyPress vbKeyMenu, 0
                
                If xboxController(0).DPadDown Then VolumeDown
                
                If xboxController(0).DPadUp Then VolumeUp
                    
                If xboxController(0).DPadLeft Then PlaylistBack
                
                If xboxController(0).DPadRight Then PlaylistNext
                
                If xboxController(0).ButtonLBumper = 1 Then StopPlayback
                
                If xboxController(0).ButtonRBumper = 1 Then PlayPause
                
                If xboxController(0).StickL_Y > 15000 Then
                    KeyPress vbKeyUp, 0
                ElseIf xboxController(0).StickL_Y < -15000 Then
                    KeyPress vbKeyDown, 0
                End If
                
                If xboxController(0).StickL_X > 15000 Then
                    KeyPress vbKeyRight, 0
                ElseIf xboxController(0).StickL_X < -15000 Then
                    KeyPress vbKeyLeft, 0
                End If
                
                If xboxController(0).StickR_Y > 15000 Then
                
                ElseIf xboxController(0).StickR_Y < -15000 Then
                
                End If
                
                If xboxController(0).StickR_X > 15000 Then
                
                ElseIf xboxController(0).StickR_X < -15000 Then
                
                End If
                
                If xboxController(0).ButtonLThumb = 1 Then
                
                End If
                
                If xboxController(0).ButtonRThumb = 1 Then
                
                End If
                
                If xboxController(0).TriggerRight > 225 Then
                    SkipForwardLarge
                ElseIf xboxController(0).TriggerRight > 50 Then
                    SkipForwardSmall
                End If
                
                If xboxController(0).TriggerLeft > 225 Then
                    SkipBackLarge
                ElseIf xboxController(0).TriggerLeft > 50 Then
                    SkipBackSmall
                End If
                
                If xboxController(0).ButtonBack = 1 Then
                
                    KeyPress vbKeyF1, 0
                
                End If
                
                If xboxController(0).ButtonStart = 1 Then
                    Do Until MinRestoreVideo = True
                        DoEvents
                    Loop
                End If
                
            End If
        
        Else
        
            If CountControllers < lXboxCount Then
            
                'tell controls that an xbox controller has been disconnected
                
                lXboxCount = CountControllers
            
                If lXboxCount = 0 Then
                    
                    SetServiceIcon 2, sOff
                    
                    Notify "Xbox Controller Removed | " & lXboxCount & " Controllers", tmrXbox, , True, , pServiceIcon(2).Picture
        
                End If
                
                
            End If
            
        End If
        
    Else
    
        If CountControllers > 0 Then
                
            If xboxController(0).ButtonBack = 1 And bDashLocked = False Then
                    
                'kill running game
                    
                TerminateProcess sProcessPath
                
                EndProcessCheck
                    
            End If
            
        End If
    
    End If
        
End Sub

Private Sub tmrYT_Timer()

    On Error Resume Next
    
    Dim lRet As Double
    Dim lBuf As Double

'    lRet = YTGetState

'    lBuf = YTGetBuffered

    Select Case lRet
        
        Case -1 'unstarted
        
            PlayBack.State = eWaiting
        
        Case 0 'ended
        
            PlayBack.State = eFinished
        
        Case 1 'playing
        
            pVidControl(0).Picture = pPlayPause(1).Picture
            
            PlayBack.State = ePlaying
        
        Case 2 'paused
        
            pVidControl(0).Picture = pPlayPause(0).Picture
        
            PlayBack.State = ePaused
        
        Case 3 'buffering
        
        PlayBack.State = eBuffering
        
            pVidControl(0).Picture = pPlayPause(2).Picture
        
        Case 5 'cued
    
    End Select

    pVidPosition(2).Width = (pVidPosition(0).Width / 100) * CDbl(Mid(CStr((lBuf * 100)), 1, 4))

End Sub

Private Sub wmp_GotFocus()

    On Error Resume Next
    
    lOSD = 0

    ClosePlaylist

    pFocus.SetFocus

End Sub

Private Sub wmp_KeyDown(ByVal nKeyCode As Integer, ByVal nShiftState As Integer)

    On Error Resume Next
    
    KeyPress nKeyCode, 0

End Sub

Private Sub wmp_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)

    On Error Resume Next
    
    If fX <> lPrevFx Then

        ShowOSD True
        
        lPrevFx = fX
        
    End If

End Sub

Private Sub wmp_PlayStateChange(ByVal NewState As Long)

    On Error Resume Next
    
    Select Case NewState
    
        Case 0
        
        Case 1 'stopped
        
            PlayBack.State = 0
        
            pVidControl(3).Picture = pPlayPause(0).Picture
        
        Case 2 'paused
        
            PlayBack.State = 2
        
            pVidControl(3).Picture = pPlayPause(0).Picture
        
        Case 3 'playing
        
            PlayBack.State = 1
            
            pVidControl(3).Picture = pPlayPause(1).Picture
        
        Case 4 'waiting
        
            pVidControl(3).Picture = pPlayPause(0).Picture
        
            PlayBack.State = -1
        
        Case 6 'buffering
        
            PlayBack.State = 3
        
            pVidControl(3).Picture = pPlayPause(1).Picture
        
        Case 8 'playback ended

            PlaylistNext

            pVidControl(3).Picture = pPlayPause(0).Picture
        
        Case 10 'ready
        
            'ready
            
    End Select

End Sub

Private Sub cRow_TileClick(Index As Integer, TileIndex As Integer)
    
    On Error Resume Next
    
    If bFocus = True Then
    
        Dim bTemp As Boolean
        
        bTemp = True
    
        If Index <> lRowSelected Then
            Do Until SelectRow(CLng(Index)) = True
                DoEvents
            Loop
            bTemp = False
        End If
    
        If TileIndex <> cRow(Index).TileCurrent Then
        
            cRow(Index).TileSelected(CLng(TileIndex)) = True
            
        Else
                
            If bTemp = True Then
            
                KeyPress vbKeyReturn, 0
    
            End If
            
        End If
        
    End If
    
    pFocus.SetFocus

End Sub

Private Sub lblMenu_Click(Index As Integer)

    On Error Resume Next
    
    If bFocus = True Then
    
        MenuSelect CLng(Index)

    End If

    pFocus.SetFocus

End Sub

Private Sub lblRow_Click(Index As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        SelectRow (Index)
        
        lRowSelected = Index

    End If

    pFocus.SetFocus

End Sub

Private Sub pMenuIcon_Click()

    On Error Resume Next
    
    If bFocus = True Then
        
        If pMenuHolder.Visible = False Then
            OpenMenu
        Else
            CloseMenu
        End If

    End If

    pFocus.SetFocus

End Sub

Private Sub pSearchIcon_Click()

    On Error Resume Next
    
    If bFocus = True Then
        
        If CurrentScreen <> 2 Then
            Do Until OpenSearch = True
                DoEvents
            Loop
        ElseIf CurrentScreen = 3 Then
            Do Until CloseSearchResults = True
                DoEvents
            Loop
        Else
            Do Until CloseSearch = True
                DoEvents
            Loop
        End If
    
    End If

    pFocus.SetFocus

End Sub


Private Sub cSearch_TileClick(Index As Integer)

    On Error Resume Next
    
    If bFocus = True Then
        
        If Index <> cSearch.TileCurrent Then
        
            cSearch.TileSelected(CLng(Index)) = True
            
        Else
                
            KeyPress vbKeyReturn, 0
            
        End If

    End If

    pFocus.SetFocus

End Sub




'''''''''''''''''''''''''''
'''''BEGIN APP COMMAND'''''
'''''''''''''''''''''''''''

Public Sub m_cAppCommand_AppCommand(ByVal command As AppCommandConstants, ByVal fromDevice As AppCommandDeviceConstants, ByVal keyState As AppCommandKeyStateConstants, ByRef processed As Boolean)
   
    On Error Resume Next
        
   ' process and determine whether to pass on to Windows or not:
   processed = processCommand(command)
   
End Sub


'''''''''''''''''''''''''
'''''END APP COMMAND'''''
'''''''''''''''''''''''''



