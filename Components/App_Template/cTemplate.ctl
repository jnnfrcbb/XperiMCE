VERSION 5.00
Object = "{44D09AE3-1847-41E9-B1EF-890580211EC2}#1.0#0"; "AlphaImage.ocx"
Object = "{79A8D05F-8453-4684-8B83-13A4A8B0D380}#13.0#0"; "prjHoriSccroll.ocx"
Object = "{454BC717-09D7-4B22-B733-CAA2F3634B19}#5.0#0"; "prjKeyboard.ocx"
Begin VB.UserControl cTemplate 
   BackColor       =   &H001E1E1E&
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   16
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   810
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1440
   Begin VB.PictureBox pSearchHolder 
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      Height          =   5890
      Left            =   900
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1260
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   12600
      Begin VB.Timer tmrCloseSearch 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer tmrOpenSearch 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   300
      End
      Begin prjKeyboard.cKeyboard cKeyboard 
         Height          =   4400
         Left            =   300
         TabIndex        =   21
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
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00181818&
         BackStyle       =   0  'Transparent
         Caption         =   "Type to search..."
         ForeColor       =   &H00565656&
         Height          =   560
         Left            =   410
         TabIndex        =   23
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
         TabIndex        =   22
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
      TabIndex        =   19
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
         Image           =   "cTemplate.ctx":0000
         Settings        =   -2147483648
         Render          =   4
         Attr            =   513
         Effects         =   "cTemplate.ctx":10DB
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
      TabIndex        =   18
      Top             =   1680
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
         Image           =   "cTemplate.ctx":10F3
         Settings        =   8448
         Attr            =   513
         Effects         =   "cTemplate.ctx":1362
      End
   End
   Begin VB.PictureBox pMenuHolder 
      BackColor       =   &H004B4B4B&
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   3720
      ScaleHeight     =   810
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
      Begin VB.PictureBox pMenuWidth 
         Height          =   190
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   3560
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Timer tmrCloseMenu 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   5580
      End
      Begin VB.Timer tmrOpenMenu 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1260
         Top             =   5460
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
         Left            =   920
         TabIndex        =   17
         Top             =   470
         Width           =   2540
         WordWrap        =   -1  'True
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
         Left            =   920
         TabIndex        =   16
         Top             =   120
         Width           =   2540
         WordWrap        =   -1  'True
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pUser 
         Height          =   620
         Left            =   150
         Top             =   150
         Width           =   620
         _ExtentX        =   1094
         _ExtentY        =   1094
         Image           =   "cTemplate.ctx":137A
         Render          =   4
         Attr            =   513
         Effects         =   "cTemplate.ctx":272A
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Home"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
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
         TabIndex        =   11
         Top             =   1020
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Playlist"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
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
         TabIndex        =   10
         Top             =   1980
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Subscriptions"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   440
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   2940
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Settings"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   440
         Index           =   5
         Left            =   0
         TabIndex        =   8
         Top             =   3780
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Exit"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   440
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   4260
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Search"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
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
         TabIndex        =   6
         Top             =   1500
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00787878&
         BackStyle       =   0  'Transparent
         Caption         =   "  Accounts"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.5
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
         TabIndex        =   5
         Top             =   2460
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Line lnMenu 
         BorderColor     =   &H00616161&
         Index           =   0
         X1              =   20
         X2              =   220
         Y1              =   258
         Y2              =   258
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
         Effects         =   "cTemplate.ctx":2742
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   340
      ScaleWidth      =   340
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox pSearchWidth 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3840
      ScaleHeight     =   140
      ScaleWidth      =   2410
      TabIndex        =   12
      Top             =   300
      Visible         =   0   'False
      Width           =   2409
   End
   Begin VB.PictureBox pLoadingCover 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7545
      Left            =   14400
      ScaleHeight     =   755
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14400
      Begin LaVolpeAlphaImg.AlphaImgCtl pNoNetwork 
         Height          =   675
         Left            =   6900
         Top             =   6720
         Width           =   675
         _ExtentX        =   1199
         _ExtentY        =   1199
         Image           =   "cTemplate.ctx":275A
         Attr            =   513
         Effects         =   "cTemplate.ctx":3690
      End
      Begin VB.Label lblNetwork 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Network Connection"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   60
         TabIndex        =   3
         Top             =   6180
         Width           =   14400
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl pLoadingLogo 
         Height          =   6750
         Left            =   3825
         Top             =   0
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   11906
         Image           =   "cTemplate.ctx":36A8
         Attr            =   513
         Effects         =   "cTemplate.ctx":47DD
      End
   End
   Begin VB.PictureBox pRowSpace 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1200
      ScaleHeight     =   1010
      ScaleWidth      =   940
      TabIndex        =   2
      Top             =   2558
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmrPublish 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10620
      Top             =   180
   End
   Begin VB.Timer tmrStartup 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   180
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9600
      Top             =   180
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   10080
      Top             =   180
   End
   Begin VB.Timer tmrHideTopBar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   540
      Top             =   540
   End
   Begin VB.Timer tmrShowTopBar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   60
   End
   Begin VB.Timer tmrScrollRowsUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrScrollRowsDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   720
   End
   Begin prjHoriSccroll.cItems cRow 
      Height          =   3735
      Index           =   0
      Left            =   0
      Top             =   2558
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   6579
      Transparent     =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pMenuIcon 
      Height          =   750
      Left            =   75
      Top             =   100
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Image           =   "cTemplate.ctx":47F5
      Settings        =   -2147468288
      Attr            =   513
      Effects         =   "cTemplate.ctx":58D0
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pLoading 
      Height          =   500
      Index           =   0
      Left            =   200
      Top             =   200
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   882
      _ExtentY        =   882
      Image           =   "cTemplate.ctx":58E8
      Render          =   4
      Attr            =   513
      Effects         =   "cTemplate.ctx":BE98
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pSearchIcon 
      Height          =   495
      Left            =   228
      Top             =   1890
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   864
      Image           =   "cTemplate.ctx":BEB0
      Settings        =   15360
      Attr            =   513
      Effects         =   "cTemplate.ctx":C11F
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AppTitle"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   28
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   740
      Left            =   1200
      TabIndex        =   14
      Top             =   180
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   750
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   1808
      Width           =   12000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pTopBar 
      Height          =   1080
      Left            =   0
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   1905
      Render          =   4
      Attr            =   513
      Effects         =   "cTemplate.ctx":C137
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pSideFade 
      Height          =   8100
      Left            =   0
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   14288
      Image           =   "cTemplate.ctx":C14F
      Render          =   4
      Attr            =   513
      Effects         =   "cTemplate.ctx":CD20
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pBack 
      Height          =   8100
      Left            =   0
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   14288
      Settings        =   1048576
      Render          =   4
      Attr            =   513
      Effects         =   "cTemplate.ctx":CD38
   End
End
Attribute VB_Name = "cTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB OpenUrl"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hOpen As Long, ByVal sURL As String, ByVal sHeaders As String, _
ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Dim bDownloaded As Boolean
Dim hOpen               As Long
Dim hOpenUrl            As Long
Dim sURL                As String
Dim bDoLoop             As Boolean
Dim bRet                As Boolean
Dim sReadBuffer         As String * 2048
Dim lNumberOfBytesRead  As Long
Dim sBuffer             As String
Dim fnum As Long
Private sResponse As String

Private ResizeOld As Double
Private ResizeNew As Double

Event Publish(sTitle As String, sThumb As String, sLink As String)
Event Message(sMessage As String, pImage As GDIpImage)

Private ScreenSelected As Long
Private ScreenPrev As Long
Private lRowSelected As Long
Private bKey As Boolean

Private Type ePubItem
    Title As String
    Thumb As String
    Link As String
End Type
Private ePublishList() As ePubItem
Private lPublishItem As Long

Private bPinRequest As Boolean

Private lMenu As Long
Private lMenuSelect As Long
Private bMenuLeft As Boolean

Private Const cRowTitleSelected = vbWhite
Private Const cRowTitleDeselected = &H808080

Private Const cSearchSelected = &HE0E0E0
Private Const cSearchDeselected = &H565656

Private RowLoaded() As Boolean

Private bSyncResponse As Boolean
Private lSync As Long

Private bNetwork As Boolean
Private authCode As String
Private bLoggedIn As Boolean

Private fso As New FileSystemObject

Private Sub cKeyboard_GoUp()

    On Error Resume Next
    
    Do Until CloseSearch = True
        DoEvents
    Loop
    
End Sub

Private Sub cKeyboard_SearchCommand(SearchString As String)

    On Error Resume Next
    
    Do Until DoSearch(lblSearch.Caption) = True
        DoEvents
    Loop

End Sub

Private Sub cKeyboard_StringChange(sSearchString As String)

    On Error Resume Next
    
    lblSearch.ForeColor = cSearchSelected

    lblSearch.Caption = sSearchString

End Sub


Private Sub cRow_Click(index As Integer)

    On Error Resume Next
    
    If index <> lRowSelected Then
        Do Until SelectRow(CLng(index)) = True
            DoEvents
        Loop
    End If

End Sub

Private Function UpdateTopBar() As Boolean
   
    On Error Resume Next

    If lRowSelected = 0 Then
    
        pTopBar.Visible = True
        
    Else
    
        pTopBar.Visible = False
    
    End If
    
    lblAppTitle.Visible = pTopBar.Visible
    
    UpdateTopBar = True

End Function

Private Sub pBack_Click()

End Sub

Private Sub pLoading_Click(index As Integer)

End Sub

Private Sub pNoNetwork_Click()

    On Error Resume Next
    
    CloseControl

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

End Sub

Private Sub tmrPublish_Timer()

    On Error Resume Next

    Static i As Integer
    
    i = i + 1
    
    If i = 15 Then
    
        If lPublishItem = lPublishCount Then
            lPublishItem = 0
        Else
            lPublishItem = lPublishItem + 1
        End If
    
        RaiseEvent Publish(ePublishList(lPublishItem).Title, LoadPictureGDIplus(ePublishList(lPublishItem).Thumb, False, True), "")
        
        i = 0
        
    End If

End Sub

Private Sub tmrScrollRowsDown_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    cRow(lRowSelected).Top = (cRow(lRowSelected).Top + pRowSpace.Height)

    Do Until PositionRows = True
        DoEvents
    Loop

    If i = 5 Then
    
        i = 0
    
        tmrScrollRowsDown.Enabled = False
    
    End If
    
End Sub

Private Sub tmrScrollRowsUp_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    cRow(lRowSelected).Top = (cRow(lRowSelected).Top - pRowSpace.Height)

    Do Until PositionRows = True
        DoEvents
    Loop

    If i = 5 Then
    
        i = 0
    
        tmrScrollRowsUp.Enabled = False
    
    End If
    
End Sub

Private Sub tmrShutdown_Timer()

    On Error Resume Next

    Static i As Integer
    
    i = i + 1
        
    Do Until QuickWait(25) = True
        DoEvents
    Loop
    
    If i = 1 Then
    
    ElseIf i = 2 Then
    
        pTopBar.Visible = False
        
    ElseIf i = 3 Then
    
        i = 0
    
        tmrShutdown.Enabled = False
    
    End If
    
End Sub

Private Sub tmrStartup_Timer()

    On Error Resume Next

    Static i As Integer
    
    Do Until QuickWait(25) = True
        DoEvents
    Loop
    
    i = i + 1
    
    If i = 5 Then
    
        pTopBar.Visible = True
    
    ElseIf i = 6 Then
    
        'pSearchIcon.Visible = True
    
    ElseIf i = 7 Then
    
        cRow(0).Visible = True
    
        lblRow(0).Visible = True
    
    ElseIf i = 8 Then
    
        i = 0
        
        tmrStartup.Enabled = False
    
    End If

End Sub

Private Sub tmrWait_Timer()

    On Error Resume Next
    
    Static i As Integer
    
    i = i + 1
    
    If i = 1 Then
    
        i = 0
        
        tmrWait.Enabled = False
        
    End If

End Sub

Private Function QuickWait(iInt As Integer) As Boolean

    On Error Resume Next
    
    tmrWait.Interval = iInt
    
    tmrWait.Enabled = True
    
    Do Until tmrWait.Enabled = False
        DoEvents
    Loop
    
    QuickWait = True

End Function

Private Sub UserControl_Initialize()

    On Error Resume Next

    ResizeOld = UserControl.ScaleWidth

    Dim cControl As Control

    For Each cControl In UserControl.Controls
        
        If TypeOf cControl Is Label Then
        
            cControl.UseMnemonic = False
        
        End If
    
    Next
      
    pMenuHolder.Left = -pMenuHolder.Width

    cRow(0).TileColorSelected = pTopBar.BackColor
    cRow(0).TileColorDeselected = &H303030
    cSearch.TileColorSelected = cRow(0).TileColorSelected
    cSearch.TileColorDeselected = cRow(0).TileColorDeselected
      
    cKeyboard.MainKeySelectedBackColor = pTopBar.BackColor
    cKeyboard.FunctionKeySelectedBackColor = pTopBar.BackColor
      
    ReDim RowLoaded(0)
    
    pLoadingCover.Left = 0
    pLoadingCover.ZOrder 0
    pLoadingCover.Visible = True
      
    cRow(0).Transparent = True

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    ResizeForm

End Sub

Private Function ResizeForm()

    On Error Resume Next

    Dim lPC As Double
    
    With UserControl
    
        ResizeNew = .ScaleWidth
        
        lPC = (ResizeNew / ResizeOld) * 100
        
        Dim cCont As Control
        
        For Each cCont In .Controls
        
            ResizeControl cCont, lPC
            
        Next cCont
        
        'END FUNCTION
        
        ResizeOld = .ScaleWidth
    
    End With

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

Public Function SendMessage(sMessage As String, Optional pImage As GDIpImage) As Boolean

    On Error Resume Next

    Dim strSplit() As String
    
    strSplit = Split(sMessage, "##")
    
    Select Case LCase(strSplit(0))
    
        Case "keypress"
        
            Do Until KeyPress(CInt(strSplit(1)), CInt(strSplit(2))) = True
                DoEvents
            Loop
    
        Case "initcontrol"
        
            Do Until InitControl = True
                DoEvents
            Loop
        
        Case "closecontrol"
        
            CloseControl
        
        Case "returndown"
        
            Do Until ReturnDown = True
                DoEvents
            Loop
            
        Case "requestpublish"
        
            Do Until RequestPublish = True
                DoEvents
            Loop
            
        Case "netconnected"
        
            If bNetwork = False Then
            
                RequestPublish
            
            End If
            
            bNetwork = True
        
        Case "netdisconnected"
        
            tmrPublish.Enabled = False
        
            bNetwork = False
            
            pLoadingCover.Left = 0
            pLoadingCover.ZOrder 0
            pLoadingCover.Visible = True
            
            ScreenSelected = -99
            
        Case "pin"
        
            Do Until InitPin = True
                DoEvents
            Loop
            
            bPinRequest = True
        
        Case "userloggedin"
        
            authCode = strSplit(1)
        
            lblMenuHeader(0).Caption = strSplit(2)
            lblMenuHeader(1).Caption = strSplit(3)
            
            lblMenuHeader(0).Left = (pUser.Left * 2) + pUser.Width
            lblMenuHeader(1).Left = lblMenuHeader(0).Left
            
            pUser.Picture = pImage
            
            lblMenu(2).ForeColor = &HE0E0E0
            lblMenu(3).ForeColor = &HE0E0E0
            lblMenu(4).ForeColor = &HE0E0E0
            
            bLoggedIn = True
            
        Case "userloggedout"
        
            authCode = strSplit(1)
            
            lblMenuHeader(0).Caption = "User not logged in"
            lblMenuHeader(1).Caption = ""
            
            lblMenuHeader(0).Left = pUser.Left
            lblMenuHeader(1).Left = lblMenuHeader(0).Left
            
            pUser.Picture = Nothing
            
            lblMenu(2).ForeColor = &H808080
            lblMenu(3).ForeColor = &H808080
            lblMenu(4).ForeColor = &H808080

            bLoggedIn = False
        
        Case "usercover"
        
            pUserCover.FastRedraw = False
            pUserCover.Picture = pImage
            pUserCover.Effects.CreateBlurEffect 25, False
            pUserCover.FastRedraw = True

        Case "authcode"
        
            authCode = strSplit(1)
            
        Case "menulefton"
        
            bMenuLeft = True
        
        Case "menuleftoff"
        
            bMenuLeft = False
            
        Case "sync"
                   
            Select Case strSplit(1)
            
                Case "status"
                
                    lSync = CLng(strSplit(2))
                
            End Select
            
            bSyncResponse = True
             
        Case "backgroundchanged"
        
            pBack.Picture = pImage
                  
    End Select
    
    SendMessage = True
    
End Function

Private Function KeyPress(KeyCode As Integer, Shift As Integer) As Boolean

    On Error Resume Next
    
    Dim i As Integer
    Dim lItem As Long
    
    Debug.Print ScreenSelected

    If bKey = True Then
    
        bKey = False

        If pMenuHolder.Visible = True Then
        
            Select Case KeyCode
            
                Case vbKeyUp
                
                    If lMenu > 0 Then
                        
                        lblMenu(lMenu).BackStyle = 0
                        lblMenu(lMenu).ForeColor = &HE0E0E0
                        
                        If lMenu = 5 Then
                            If bLoggedIn = True Then
                                lMenu = lMenu - 1
                            Else
                                lMenu = 1
                            End If
                        Else
                            lMenu = lMenu - 1
                        End If
                        
                        RaiseEvent Message("sound##select", Nothing)
                        
                        lblMenu(lMenu).BackStyle = 1
                        lblMenu(lMenu).ForeColor = vbWhite
                        
                    Else
                        
                        RaiseEvent Message("sound##listend", Nothing)
                    
                    End If
                
                Case vbKeyDown
                
                    If lMenu < lblMenu.UBound Then
                
                        lblMenu(lMenu).BackStyle = 0
                        lblMenu(lMenu).ForeColor = &HE0E0E0
                        
                        If lMenu = 1 Then
                            If bLoggedIn = True Then
                                lMenu = lMenu + 1
                            Else
                                lMenu = 5
                            End If
                        Else
                            lMenu = lMenu + 1
                        End If
                        
                        RaiseEvent Message("sound##select", Nothing)
                        
                        lblMenu(lMenu).BackStyle = 1
                        lblMenu(lMenu).ForeColor = vbWhite
                    
                    Else
                        
                        RaiseEvent Message("sound##listend", Nothing)
                    
                    End If
                    
                Case vbKeyReturn
            
                    Do Until MenuSelect(lMenu) = True
                        DoEvents
                    Loop
            
                Case vbKeyMenu, 93, vbKeyBack, vbKeyRight
                
                    Do Until CloseMenu = True
                        DoEvents
                    Loop
                
            End Select
        
        Else
    
            Select Case ScreenSelected
            
                Case -99 'no network
                
                    Select Case KeyCode
                    
                        Case vbKeyBack
                        
                            CloseControl
                    
                    End Select
            
                Case -1 'search
                
                    Select Case KeyCode
                    
                        Case vbKeyBack
                        
                            If cKeyboard.CurrentString = "" Then
                                Do Until CloseSearch = True
                                    DoEvents
                                Loop
                            Else
                                cKeyboard.KeyPress KeyCode, Shift
                            End If
                    
                        Case vbKeyShift
                    
                            Do Until CloseSearch = True
                                DoEvents
                            Loop
                            
                        Case vbKeyMenu, 93
                        
                            Do Until OpenMenu = True
                                DoEvents
                            Loop
                            
                        Case Else
                        
                            cKeyboard.KeyPress KeyCode, Shift
                    
                    End Select
                
                Case 0 'app home
                
                    Select Case KeyCode
                    
                        Case vbKeyLeft
                        
                            If cRow(lRowSelected).TileCurrent > 0 Then
                            
                                Do Until cRow(lRowSelected).TileSelectLeft = True
                                    DoEvents
                                Loop
                            
                            Else
                                
                                If bMenuLeft = True Then
                                
                                    Do Until OpenMenu = True
                                        DoEvents
                                    Loop
                                    
                                Else
                                
                                    RaiseEvent Message("sound##listend", Nothing)
                            
                                End If
                        
                            End If
                                
                        Case vbKeyRight
                    
                            Do Until cRow(lRowSelected).TileSelectRight = True
                                DoEvents
                            Loop
                            
                        Case vbKeyUp
                        
                            Do Until SelectRow(lRowSelected - 1) = True
                                DoEvents
                            Loop
                        
                        Case vbKeyDown
                        
                            Do Until SelectRow(lRowSelected + 1) = True
                                DoEvents
                            Loop
                                                            
                        Case vbKeyReturn
                    
                            RaiseEvent Message("sound##yes", Nothing)
                            
                            lItem = CLng(cRow(lRowSelected).TileTag(cRow(lRowSelected).TileCurrent))
                    
                            'Do Until LoadFilm(lItem, True) = True
                            '    DoEvents
                            'Loop
                            
                        Case vbKeyBack
                        
                            CloseControl
                            
                        Case vbKeyShift
                        
                            Do Until OpenSearch = True
                                DoEvents
                            Loop
                        
                        Case vbKeyMenu, 93
                          
                            Do Until OpenMenu = True
                                DoEvents
                            Loop
                                                                    
                    End Select
                                                
                Case 97 'settings screen
                    
                Case 98 'search results
                    
                    Select Case KeyCode
                        
                        Case vbKeyLeft
                        
                            If lblSearchNoResults.Visible = False Then
        
                                If cSearch.TileCurrent > 0 Then
                                    cSearch.TileSelectLeft
                                Else
                                    
                                    If bMenuLeft = True Then
                                    
                                        Do Until OpenMenu = True
                                            DoEvents
                                        Loop
                                        
                                    Else
                                    
                                        RaiseEvent Message("sound##listend", Nothing)
                                    
                                    End If
                        
                                End If
                                
                            Else
                                
                                RaiseEvent Message("sound##no", Nothing)
    
                            End If
                        
                        Case vbKeyRight
                        
                            If lblSearchNoResults.Visible = False Then
            
                                cSearch.TileSelectRight
                        
                            Else
                        
                                RaiseEvent Message("sound##no", Nothing)
    
                            End If
                        
                        Case vbKeyReturn
                        
                            If lblSearchNoResults.Visible = False Then
        
                                lItem = CLng(cSearch.TileTag(cSearch.TileCurrent))
                        
                                RaiseEvent Message("sound##yes", Nothing)
                            
                                pSearchHolder.Visible = False
    
                                'Do Until LoadFilm(lItem, True) = True
                                '    DoEvents
                                'Loop
                                
                            Else
                                
                                RaiseEvent Message("sound##no", Nothing)
    
                            End If
                                
                        Case vbKeyShift, vbKeyBack
                    
                            Do Until CloseSearchResults(False) = True
                                DoEvents
                            Loop
                            
                        Case vbKeyMenu, 93
                        
                            Do Until OpenMenu = True
                                DoEvents
                            Loop
                                
                    End Select
            
                Case 99 'pin request
                
                    Select Case KeyCode
                    
                        Case vbKeyLeft
                        
                            Do Until cRow(lRowSelected).TileSelectLeft = True
                                DoEvents
                            Loop
                            
                        Case vbKeyRight
                        
                            Do Until cRow(lRowSelected).TileSelectRight = True
                                DoEvents
                            Loop
                            
                        Case vbKeyUp
                    
                            Do Until SelectRow(lRowSelected - 1) = True
                                DoEvents
                            Loop
                    
                        Case vbKeyDown
                        
                            Do Until SelectRow(lRowSelected + 1) = True
                                DoEvents
                            Loop
                            
                        Case vbKeyReturn
                    
                            RaiseEvent Message("sound##yes", Nothing)
    
                            lItem = CLng(cRow(lRowSelected).TileTag(cRow(lRowSelected).TileCurrent))
                    
                            'RaiseEvent Message("pin##tvshow##" & Shows(lItem).Title & "##" & Shows(lItem).Folder & "##" & Shows(lItem).Thumb, Nothing)
                            
                            CloseControl
                            
                        Case vbKeyBack
                        
                            CloseControl
                                                                 
                    End Select
                                                        
            End Select
            
        End If
        
    End If
    
    bKey = True
    
    KeyPress = True

End Function

Private Function InitControl() As Boolean

    On Error Resume Next
    
    If bNetwork = True Then
    
        pLoadingCover.Visible = False
    
        Static i As Long
        
        ReDim RowLoaded(0)
      
        Loading True
        
        bPinRequest = False
        
        tmrStartup.Enabled = True
            
        Do Until tmrStartup.Enabled = False
            DoEvents
        Loop
            
        'Do Until LoadFilms(True) = True
        '    DoEvents
        'Loop
            
        ScreenSelected = 0
        
        ScreenPrev = -1
        
        cRow(0).TileSelected(0) = True
        
    Else
    
        ScreenSelected = 99
        
        pLoadingCover.ZOrder 0
        pLoadingCover.Visible = True
        
    End If
        
    bKey = True
    
    Loading False
    
    InitControl = True

End Function

Private Function InitPin() As Boolean

    On Error Resume Next
    
    bPinRequest = True
    
    pMenuIcon.Visible = False
    
    pSearchIcon.Visible = False

    pTopBar.Visible = True
    
    cRow(0).Visible = True
    
    lblRow(0).Visible = True

    'Do Until LoadTV("G:\Video\TV", , True) = True
    '    DoEvents
    'Loop
        
    cRow(0).TileSelected(0) = True
        
    ScreenSelected = 99

    bKey = True
    
    InitPin = True

End Function

Private Function CloseControl() As Boolean

    On Error Resume Next

    Dim i As Integer
    
    RaiseEvent Message("sound##closeapp", Nothing)
                    
    tmrShutdown.Enabled = True
    
    For i = 0 To cRow.UBound
        cRow(i).Visible = False
        lblRow(i).Visible = False
        Do Until QuickWait(25) = True
            DoEvents
        Loop
    Next
    
    Do Until tmrShutdown.Enabled = False
        DoEvents
    Loop

    RaiseEvent Message("LeaveControl", Nothing)

    cRow(0).Top = pRowSpace.Top
    
    cRow(0).ClearTiles
    lblRow(0).Caption = ""
    
    If cRow.UBound > 0 Then
        For i = 1 To cRow.UBound
            Unload cRow(i)
            Unload lblRow(i)
        Next
    End If
    
    CloseControl = True

    pMenuHolder.Left = -pMenuHolder.Width

    lRowSelected = 0
End Function

Public Function Loading(bLoading As Boolean)
    
    On Error Resume Next

    Dim i As Integer

    pSearchIcon.Visible = Not bLoading
    pSideFadeIcon(1).Visible = Not bLoading
    
    For i = 0 To pMenuIcon.UBound
        pMenuIcon(i).Visible = Not bLoading
    Next

    If bLoading = True Then
        For i = 0 To pLoading.UBound
            pLoading(i).Animate lvicAniCmdStart
            pLoading(i).Visible = True
        Next
    Else
        For i = 0 To pLoading.UBound
            pLoading(i).Animate lvicAniCmdStop
            pLoading(i).Visible = False
        Next
    End If
        
End Function

Public Function RequestPublish() As Boolean

    On Error Resume Next
    
    lPublishItem = 0
    
    RaiseEvent Publish(ePublishList(lPublishItem).Title, LoadPictureGDIplus(ePublishList(lPublishItem).Thumb, False, True), ePublishList(lPublishItem).Link)

    tmrPublish.Enabled = True

End Function

Private Function ReturnDown() As Boolean

    On Error Resume Next

    Select Case ScreenSelected
        
        Case 0
        
            cRow(0).TileSelected(cRow(0).TileCurrent) = True
        
        Case 1
    
    End Select


    ReturnDown = True

End Function

Private Function PositionRows() As Boolean

    On Error Resume Next

    lblRow(lRowSelected).Top = cRow(lRowSelected).Top - lblRow(lRowSelected).Height
    
    If lRowSelected > 0 Then
        
        cRow(lRowSelected - 1).Top = cRow(lRowSelected).Top - (pRowSpace.Height * 5)
    
        lblRow(lRowSelected - 1).Top = cRow(lRowSelected - 1).Top - lblRow(lRowSelected - 1).Height
        
    End If
    
    If lRowSelected > 1 Then
        
        cRow(lRowSelected - 2).Top = cRow(lRowSelected - 1).Top - (pRowSpace.Height * 5)
    
        lblRow(lRowSelected - 2).Top = cRow(lRowSelected - 2).Top - lblRow(lRowSelected - 2).Height

    End If
    
    If lRowSelected < cRow.UBound Then
        
        cRow(lRowSelected + 1).Top = cRow(lRowSelected).Top + (pRowSpace.Height * 5)
        
        lblRow(lRowSelected + 1).Top = cRow(lRowSelected + 1).Top - lblRow(lRowSelected + 1).Height

    End If

    If lRowSelected < cRow.UBound - 1 Then
        
        cRow(lRowSelected + 2).Top = cRow(lRowSelected + 1).Top + (pRowSpace.Height * 5)
        
        lblRow(lRowSelected + 2).Top = cRow(lRowSelected + 2).Top - lblRow(lRowSelected + 2).Height

    End If
    
    PositionRows = True

End Function

Private Function LoadRow(index As Long, Title As String) As Boolean

    On Error Resume Next
    
    If index > cRow.UBound Then
    
        ReDim Preserve RowLoaded(index)
    
        Load cRow(index)
        Load lblRow(index)
    
        cRow(index).Top = cRow(index - 1).Top + (pRowSpace.Height * 5)
        lblRow(index).Top = cRow(index).Top - lblRow(index).Height
        
        cRow(index).TileColorSelected = cRow(index - 1).TileColorSelected
        cRow(index).TileColorDeselected = cRow(index - 1).TileColorDeselected
        lblRow(index).ForeColor = cRowTitleDeselected
    
        cRow(index).ZOrder 0
        lblRow(index).ZOrder 0
        
        lblRow(index).Caption = Title
    
        cRow(index).Visible = True
        lblRow(index).Visible = True
    
    Else
    
        cRow(index).ClearTiles
    
        lblRow(index).Caption = Title
    
    End If
    
    RowLoaded(index) = False
    
    pMenuHolder.ZOrder 0
    
    LoadRow = True

End Function

Private Function LoadRowTiles(index As Long) As Boolean

    On Error Resume Next
    
    Loading True
    
    Dim i As Long
    Dim lTag As Long
    
    For i = 0 To cRow(index).TileCount
        
        DoEvents
        
        lTag = CLng(cRow(index).TileTag(i))
        
        'Do Until cRow(Index).TileSet(i, , , , , CStr(lTag)) <> 0
        '    DoEvents
        'Loop
        
    Next
    
    RowLoaded(index) = True

    LoadRowTiles = True

    Loading False

End Function

Private Function SelectRow(lNewRow As Long, Optional lNewTile As Long = 0) As Boolean

    On Error Resume Next
    
    If lNewRow >= 0 And lNewRow <= cRow.UBound Then
    
        If RowLoaded(lNewRow) = True Then
        
            cRow(lRowSelected).Deselect
            
            lblRow(lRowSelected).ForeColor = cRowTitleDeselected
                
            RaiseEvent Message("sound##select", Nothing)
                            
            If lNewRow < lRowSelected Then
            
                tmrScrollRowsDown.Enabled = True
                
                Do Until tmrScrollRowsDown.Enabled = False
                    DoEvents
                Loop
            
            ElseIf lNewRow > lRowSelected Then
            
                tmrScrollRowsUp.Enabled = True
                
                Do Until tmrScrollRowsUp.Enabled = False
                    DoEvents
                Loop
        
            End If
                
            RaiseEvent Message("sound##select", Nothing)
                            
            lRowSelected = lNewRow
            
            lblRow(lRowSelected).ForeColor = cRowTitleSelected
            
            cRow(lRowSelected).TileSelected(lNewTile) = True
                
            UpdateTopBar
    
        Else
        
            Do Until LoadRowTiles(lNewRow) = True
                DoEvents
            Loop
            
            Do Until SelectRow(lNewRow) = True
                DoEvents
            Loop
            
        End If
        
    Else
    
        RaiseEvent Message("sound##listend", Nothing)
    
    End If
    
    SelectRow = True
                    
End Function

Private Function OpenSearch() As Boolean

    On Error Resume Next
    
    Static i As Integer

        RaiseEvent Message("sound##openelement", Nothing)
                            
        For i = 0 To cRow.UBound
            DoEvents
            cRow(i).Visible = False
            lblRow(i).Visible = False
            If (cRow(i).Top + cRow(i).Height < UserControl.ScaleHeight) And (cRow(i).Top > -cRow(i).Height) Then
                Do Until QuickWait(50) = True
                    DoEvents
                Loop
            End If
        Next
    
        lblRow(lRowSelected).ForeColor = cRowTitleDeselected
    
        sSearchBack.Width = 0
        sSearchBack.Visible = True
        
        pSideFadeBox(1).Visible = True
        pSearchHolder.Visible = True
        
        pSearchIcon.TransparencyPct = 0
    
        tmrOpenSearch.Enabled = True
        
        Do Until tmrOpenSearch.Enabled = False
            DoEvents
        Loop
        
        pSearchIcon.Visible = True
        
        lblSearch.ForeColor = cSearchDeselected
        
        lblSearch.Caption = "Type to search..."
        
        lblSearch.Visible = True
        
        cKeyboard.Visible = True
        
        Do Until cKeyboard.InitialFocus = True
            DoEvents
        Loop
        
        ScreenPrev = ScreenSelected
        
        ScreenSelected = -1
        
    OpenSearch = True

End Function

Private Function CloseSearch() As Boolean
    
    On Error Resume Next
    
    Static i As Integer
    
        RaiseEvent Message("sound##closeelement", Nothing)
                            
        Do Until cKeyboard.CloseControl = True
            DoEvents
        Loop
    
        cKeyboard.Visible = False
        
        cSearch.Visible = False
        
        lblSearchNoResults.Visible = False
    
        cSearch.ClearTiles
        
        lblSearch.Visible = False
        
        tmrCloseSearch.Enabled = True
        
        Do Until tmrCloseSearch.Enabled = False
            DoEvents
        Loop
                
        pSearchIcon.TransparencyPct = 60
        
        sSearchBack.Visible = False
        pSideFadeBox(1).Visible = False
        pSearchHolder.Visible = False
        
        UpdateTopBar
        
        lblRow(lRowSelected).ForeColor = cRowTitleSelected
        
        For i = 0 To cRow.UBound
            DoEvents
            cRow(i).Visible = True
            lblRow(i).Visible = True
            If (cRow(i).Top + cRow(i).Height < UserControl.ScaleHeight) And (cRow(i).Top > -cRow(i).Height) Then
                Do Until QuickWait(50) = True
                    DoEvents
                Loop
            End If
        Next
        
        ScreenSelected = ScreenPrev
       
        ScreenPrev = -1
        
    CloseSearch = True

End Function

Private Function DoSearch(SearchString As String) As Boolean

    On Error Resume Next
    
    Loading True

    Dim i As Integer
    Dim lCount As Long
    
    RaiseEvent Message("sound##yes", Nothing)
                            
    lblSearchNoResults.Visible = False
    
    lblSearch.ForeColor = &H979797
    
    Do Until cKeyboard.CloseControl = True
        DoEvents
    Loop
    
    cKeyboard.Visible = False
    
    cSearch.Visible = True

    lCount = -1

    'For i = 0 To UBound(Film())
    
    '    DoEvents
        
        'If InStr(1, LCase(Film(i).Title), LCase(SearchString)) <> 0 Then
        
        '    lCount = lCount + 1
        
        '    cSearch.TileSet lCount, Film(i).Title, Film(i).Length & " | " & Film(i).Year & " | R: " & Film(i).Rating, Film(i).Thumb, True, CStr(i)
        
        'End If
    
    'Next
    
    If lCount > -1 Then
    
        cSearch.TileSelected(0) = True
        
    Else
    
        lblSearchNoResults.Visible = True
    
    End If
    
    ScreenSelected = 98
    
    Loading False
    
    DoSearch = True
    
End Function

Private Function CloseSearchResults(bHome As Boolean) As Boolean

    On Error Resume Next
    
    RaiseEvent Message("sound##closeelement", Nothing)

    cSearch.ClearTiles
    
    lblSearchNoResults.Visible = False

    cSearch.Visible = False
    
    If bHome = False Then
        
        cKeyboard.Visible = True
        
        Do Until cKeyboard.InitialFocus = True
            DoEvents
        Loop
    
        lblSearch.ForeColor = cSearchDeselected

        ScreenSelected = -1
        
    Else
    
        Do Until CloseSearch = True
            DoEvents
        Loop
    
    End If
    
    ScreenPrev = 0
    
    CloseSearchResults = True
    
End Function

Private Function OpenMenu() As Boolean

    On Error Resume Next
    
    Static i As Integer
    
    lblMenu(0).BackStyle = 1
    lblMenu(lMenu).ForeColor = vbWhite
    
    If lblMenu.UBound > 0 Then
        For i = 1 To lblMenu.UBound
            lblMenu(i).BackStyle = 0
            lblMenu(lMenu).ForeColor = &HE0E0E0
        Next
    End If
    
    lMenu = 0
    
    Select Case ScreenSelected
        Case 0 'home
            lMenuSelect = cRow(lRowSelected).TileCurrent
            cRow(lRowSelected).Deselect
            lblRow(lRowSelected).ForeColor = cRowTitleDeselected
        Case 97 'related
            lMenuSelect = cRelated.TileCurrent
            cRelated.Deselect
        Case 98 'search
            lMenuSelect = cSearch.TileCurrent
            cSearch.Deselect
    End Select
    
    RaiseEvent Message("sound##openelement", Nothing)
                        
    pSideFadeBox(0).Visible = True
    
    pMenuHolder.Width = 1
    pMenuHolder.Left = pSideFade.Width
    pMenuHolder.Visible = True
    pMenuHolder.ZOrder 0
                
    tmrOpenMenu.Enabled = True
    
    Do Until tmrOpenMenu.Enabled = False
        DoEvents
    Loop
        
    OpenMenu = True
    
End Function

Private Function CloseMenu() As Boolean

    On Error Resume Next
    
    Select Case ScreenSelected
        Case 0 'home
            cRow(lRowSelected).TileSelected(lMenuSelect) = True
            lblRow(lRowSelected).ForeColor = cRowTitleSelected
        Case 97 'related
            cRelated.TileSelected(lMenuSelect) = True
        Case 98 'search
            cSearch.TileSelected(lMenuSelect) = True
    End Select

    RaiseEvent Message("sound##closeelement", Nothing)
                            
    tmrCloseMenu.Enabled = True
    
    Do Until tmrCloseMenu.Enabled = False
        DoEvents
    Loop

    pMenuHolder.Visible = False
    pSideFadeBox(0).Visible = False
    
    CloseMenu = True

End Function

Private Function MenuSelect(index As Long) As Boolean

    On Error Resume Next
    
    Do Until CloseMenu = True
        DoEvents
    Loop
                
    Select Case index
    
        Case 0 'home
            
            Select Case ScreenSelected
            
                Case -1
                    
                    Do Until CloseSearch = True
                        DoEvents
                    Loop
            
                Case 0
                
                Case 97
                
                Case 98
            
                    Do Until CloseSearchResults(True) = True
                        DoEvents
                    Loop
            
            End Select
        
        Case 1 'search
        
            Do Until OpenSearch = True
                DoEvents
            Loop
                
        Case 2 'recommended
        
            RaiseEvent Message("sound##select", Nothing)
                            
        Case 3 'history
        
            RaiseEvent Message("sound##select", Nothing)
                            
        Case 4 'subscriptions
        
            RaiseEvent Message("sound##select", Nothing)
                            
        Case 5 'settings
        
            RaiseEvent Message("sound##select", Nothing)
                            
        Case 6 'exit
        
            CloseControl
        
    End Select
    
    MenuSelect = True

End Function

Private Function ClearTiles() As Boolean

    On Error Resume Next
    
    Static i As Integer

    Do Until cRow(0).ClearTiles = True
        DoEvents
    Loop
    
    lblRow(0).Caption = ""
    
    If cRow.UBound > 0 Then
        For i = 1 To cRow.UBound
            Unload cRow(i)
            Unload lblRow(i)
        Next
    End If
    
    ClearTiles = True
    
End Function


Private Sub cRow_TileClick(index As Integer, TileIndex As Integer)

    On Error Resume Next
    
    Dim bTemp As Boolean
    
    bTemp = True

    If index <> lRowSelected Then
        Do Until SelectRow(CLng(index)) = True
            DoEvents
        Loop
        bTemp = False
    End If

    If TileIndex <> cRow(index).TileCurrent Then
    
        cRow(index).TileSelected(CLng(TileIndex)) = True
        
    Else
            
        If bTemp = True Then
        
            KeyPress vbKeyReturn, 0

        End If
        
    End If
    
End Sub

Private Sub lblMenu_Click(index As Integer)

    On Error Resume Next
    
    MenuSelect CLng(index)

End Sub

Private Sub lblRow_Click(index As Integer)

    On Error Resume Next
    
    SelectRow (index)
    
    lRowSelected = index

End Sub

Private Sub pMenuIcon_Click()

    On Error Resume Next
    
    If pMenuHolder.Visible = False Then
        OpenMenu
    Else
        CloseMenu
    End If

End Sub

Private Sub pSearchIcon_Click()

    On Error Resume Next
    
    If ScreenSelected <> -1 Then
        Do Until OpenSearch = True
            DoEvents
        Loop
    Else
        Do Until CloseSearch = True
            DoEvents
        Loop
    End If

End Sub

Private Sub cSearch_TileClick(index As Integer)

    On Error Resume Next
    
    If TileIndex <> cSearch.TileCurrent Then
    
        cSearch.TileSelected(CLng(TileIndex)) = True
        
    Else
            
        KeyPress vbKeyReturn, 0
        
    End If

End Sub

Private Function SendGET(sURL As String) As String
    
    On Error Resume Next
    
    sBuffer = ""
    
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
 
    bDoLoop = True
    
    While bDoLoop
        
        DoEvents
    
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    
    Wend
     
    SendGET = sBuffer

End Function

Private Function SendPOST(sSiteURL As String, sHeader As String, Optional sParams As String) As String

    On Error Resume Next
    
    sResponse = ""
    
    Inet.Tag = "0"
    
    Inet.Execute sSiteURL, "POST", sParams, sHeader
    
    Do Until Inet.Tag = "1"
    
        DoEvents
        
    Loop
    
    SendPOST = sResponse
    
End Function

Private Sub cRow_TileDelete(index As Integer)

    On Error Resume Next
    
    RaiseEvent Message("sound##tiledelete", Nothing)

End Sub

Private Sub cRow_TileInsert(index As Integer)

    On Error Resume Next
    
    RaiseEvent Message("sound##tileinsert", Nothing)
    
End Sub

Private Sub cRow_TileSelectFalse(index As Integer)

    On Error Resume Next
    
    RaiseEvent Message("sound##listend", Nothing)

End Sub

Private Sub cRow_TileSelectTrue(index As Integer)

    On Error Resume Next
    
    RaiseEvent Message("sound##select", Nothing)

End Sub

Private Sub cSearch_TileSelectFalse()

    On Error Resume Next
    
    RaiseEvent Message("sound##listend", Nothing)

End Sub

Private Sub cSearch_TileSelectTrue()

    On Error Resume Next
    
    RaiseEvent Message("sound##select", Nothing)

End Sub

Private Function CloseThing() As Boolean

    Dim i As Integer
    
    RaiseEvent Message("sound##closeelement", Nothing)
    
    lblAppTitle.Caption = "TITLE"
    
    Do Until QuickWait(25) = True
        DoEvents
    Loop
    
    If bDirect = False Then
        
        Select Case ScreenPrev
        
            Case 0
            
                For i = 0 To cRow.UBound
                    cRow(i).Visible = True
                    lblRow(i).Visible = True
                Next
        
            Case 98
            
                lblSearch.Visible = True
                cSearch.Visible = True
                sSearchBack.Visible = True
            
        End Select
    
    End If

    ScreenSelected = ScreenPrev
    
    CloseThing = True
    
End Function
