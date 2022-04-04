VERSION 5.00
Object = "{44D09AE3-1847-41E9-B1EF-890580211EC2}#1.0#0"; "AlphaImage.ocx"
Object = "{454BC717-09D7-4B22-B733-CAA2F3634B19}#5.0#0"; "prjKeyboard.ocx"
Object = "{79A8D05F-8453-4684-8B83-13A4A8B0D380}#13.0#0"; "prjHoriSccroll.ocx"
Begin VB.Form frmSandpit 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   16740
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   28
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   909
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1674
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pSideFadeBox 
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   1
      Left            =   0
      ScaleHeight     =   90
      ScaleMode       =   0  'User
      ScaleWidth      =   89.732
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   900
      Begin LaVolpeAlphaImg.AlphaImgCtl pSideFadeIcon 
         Height          =   500
         Index           =   1
         Left            =   228
         Top             =   210
         Width           =   500
         _ExtentX        =   864
         _ExtentY        =   864
         Image           =   "frmSandpit.frx":0000
         Settings        =   8448
         Attr            =   513
         Effects         =   "frmSandpit.frx":026F
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLoading 
         Height          =   530
         Index           =   1
         Left            =   186
         Top             =   185
         Visible         =   0   'False
         Width           =   532
         _ExtentX        =   917
         _ExtentY        =   917
         Image           =   "frmSandpit.frx":0287
         Render          =   4
         Attr            =   513
         Effects         =   "frmSandpit.frx":6837
      End
   End
   Begin VB.PictureBox pSearchHolder 
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      Height          =   5890
      Left            =   2940
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1260
      TabIndex        =   5
      Top             =   3060
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
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   7743
      End
      Begin prjHoriSccroll.cItems cSearch 
         Height          =   3740
         Left            =   -900
         Top             =   1680
         Visible         =   0   'False
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   6597
         Transparent     =   -1  'True
      End
      Begin VB.Label lblSearchNoResults 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No results found"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   20.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   550
         Left            =   0
         TabIndex        =   8
         Top             =   3180
         Visible         =   0   'False
         Width           =   12600
         WordWrap        =   -1  'True
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
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00181818&
         BackStyle       =   0  'Transparent
         Caption         =   "Type to search..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   16
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00565656&
         Height          =   560
         Left            =   410
         TabIndex        =   7
         Top             =   200
         Visible         =   0   'False
         Width           =   11750
      End
   End
   Begin VB.PictureBox pMenuWidth 
      Height          =   190
      Left            =   5880
      ScaleHeight     =   150
      ScaleWidth      =   3560
      TabIndex        =   2
      Top             =   850
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox pSideFadeBox 
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   0
      Left            =   0
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   1
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
         Image           =   "frmSandpit.frx":684F
         Settings        =   -2147483648
         Render          =   4
         Attr            =   513
         Effects         =   "frmSandpit.frx":792A
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pUser 
      Height          =   620
      Left            =   6030
      Top             =   280
      Width           =   620
      _ExtentX        =   1094
      _ExtentY        =   1094
      Image           =   "frmSandpit.frx":7942
      Render          =   4
      Attr            =   513
      Effects         =   "frmSandpit.frx":8CF2
   End
   Begin VB.Label lblMenuHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   500
      Index           =   0
      Left            =   6800
      TabIndex        =   4
      Top             =   250
      Width           =   2540
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMenuHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "User Email"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.5
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   380
      Index           =   1
      Left            =   6800
      TabIndex        =   3
      Top             =   600
      Width           =   2540
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AppTitle"
      ForeColor       =   &H00E0E0E0&
      Height          =   740
      Left            =   1200
      TabIndex        =   0
      Top             =   500
      Visible         =   0   'False
      Width           =   2070
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pSearchIcon 
      Height          =   495
      Left            =   228
      Top             =   1110
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   864
      Image           =   "frmSandpit.frx":8D0A
      Settings        =   15360
      Attr            =   513
      Effects         =   "frmSandpit.frx":8F79
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pLoading 
      Height          =   530
      Index           =   0
      Left            =   185
      Top             =   1085
      Visible         =   0   'False
      Width           =   530
      _ExtentX        =   917
      _ExtentY        =   917
      Image           =   "frmSandpit.frx":8F91
      Render          =   4
      Attr            =   513
      Effects         =   "frmSandpit.frx":F541
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pMenuIcon 
      Height          =   750
      Left            =   75
      Top             =   100
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Image           =   "frmSandpit.frx":F559
      Settings        =   -2147468288
      Attr            =   513
      Effects         =   "frmSandpit.frx":10634
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pUserCover 
      Height          =   900
      Left            =   5880
      Top             =   130
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   1588
      Blend           =   1677721600
      Settings        =   1048576
      Render          =   4
      Frame           =   12
      BackColor       =   3947580
      Attr            =   513
      Effects         =   "frmSandpit.frx":1064C
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pSideFade 
      Height          =   8100
      Left            =   0
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   14288
      Image           =   "frmSandpit.frx":10664
      Render          =   4
      Attr            =   513
      Effects         =   "frmSandpit.frx":11235
   End
End
Attribute VB_Name = "frmSandpit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pPowerBack_Click()

End Sub
