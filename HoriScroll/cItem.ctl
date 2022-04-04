VERSION 5.00
Object = "{44D09AE3-1847-41E9-B1EF-890580211EC2}#1.0#0"; "AlphaImage.ocx"
Begin VB.UserControl cItem 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00646464&
   CanGetFocus     =   0   'False
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI Light"
      Size            =   14.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00979797&
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   2550
      UseMnemonic     =   0   'False
      Width           =   5610
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pItemFade 
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2672
      Image           =   "cItem.ctx":0000
      Render          =   4
      Attr            =   513
      Effects         =   "cItem.ctx":0BCE
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pIcon 
      Height          =   555
      Left            =   4080
      Top             =   180
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      Render          =   4
      Attr            =   513
      Effects         =   "cItem.ctx":0BE6
   End
   Begin VB.Label pSpace 
      Caption         =   "Label1"
      Height          =   75
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label pDefault 
      Caption         =   "Label1"
      Height          =   3300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Shape sBorder 
      BorderColor     =   &H00444444&
      BorderWidth     =   2
      Height          =   2550
      Left            =   0
      Top             =   60
      Width           =   4770
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pTrans 
      Height          =   555
      Left            =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      Image           =   "cItem.ctx":0BFE
      Render          =   4
      Attr            =   513
      Effects         =   "cItem.ctx":17CC
   End
   Begin VB.Label lblSubCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F1F1F1&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1980
      UseMnemonic     =   0   'False
      Width           =   4650
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pImage 
      Height          =   1785
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   3149
      Render          =   4
      Border          =   13091254
      BackColor       =   3158064
      Attr            =   1
      Effects         =   "cItem.ctx":17E4
   End
End
Attribute VB_Name = "cItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dWidth As Single

Private bSelected As Boolean

Private Const cCaptionTrue = vbWhite
Private Const cCaptionFalse = &H979797
Private cTileColorSelected As OLE_COLOR
Private cTileColorDeselected As OLE_COLOR

Private sCaption As String
Private sSubcaption As String
Private sImage As String

Private bFromSelection As Boolean
Private bHasCaption As Boolean

Private sOldHeight As Single
Private sOldWidth As Single

Event Click()

Public Function TileSet(Optional sItemCaption As String = "", Optional sItemSubCaption As String = "", Optional sItemImage As String = "", Optional bItemBlankImage As Boolean = False, Optional sIcon As String = "") As Single

    On Error GoTo errHandle
    
    Dim pItemImage As GDIpImage
    Dim pResizeImage As New GDIpImage
    Dim SaveOptions As SAVESTRUCT
    
    sCaption = sItemCaption
    sSubcaption = sItemSubCaption
    sImage = sItemImage
    
    'AsyncAbortDownloads
    
    If sItemImage <> "" Then
        If IsURL(sItemImage) = False Then
            If Dir(sItemImage, vbHidden) <> "" Then
                Set pItemImage = LoadPictureGDIplus(sItemImage, , False)
            End If
        Else
            On Error GoTo errHandle
            Set pItemImage = LoadPictureGDIplus(sItemImage)
        End If
    Else
        Set pItemImage = Nothing
    End If
    
    'hide item objects
        
        lblCaption.Visible = False
        lblSubCaption.Visible = False
        pImage.Visible = False
        pTrans.Visible = False
    
    'clear item objects
    
        dWidth = 0
    
        lblCaption.Caption = ""
        
        Set pImage.Picture = Nothing

    'set captions
    
        If sItemCaption <> "" Then
            lblCaption.Caption = sItemCaption
            lblCaption.Visible = True
            pTrans.Visible = True
        Else
            pTrans.Visible = False
        End If
        
        If sItemSubCaption <> "" Then
            lblSubCaption.Caption = sItemSubCaption
            lblSubCaption.Visible = True
        End If
        
    'handle image
    
        If Not pItemImage Is Nothing Then
        
            dWidth = CalculateWidth(pItemImage.Height, pItemImage.Width)
            
            SaveOptions.Width = dWidth * 1.25
            SaveOptions.Height = pImage.Height * 1.25
            
            Call SavePictureGDIplus(pItemImage, pResizeImage, , SaveOptions)
            
            pImage.Picture = pResizeImage
            
            pImage.Visible = True
            
        Else
        
            dWidth = pDefault.Width
            
        End If
        
        If pItemImage Is Nothing Then
        
            dWidth = pDefault.Width
                
        End If

    'set, postion, and show main image
    
        pImage.Height = UserControl.ScaleHeight
        
        pImage.Width = dWidth
        
    're-position and show caption
            
        lblCaption.Width = dWidth - (lblCaption.Left * 2)
            
        lblCaption.Top = (pImage.Top + pImage.Height) - (lblCaption.Height + (sBorder.BorderWidth * 2))
        
        pTrans.Move 0, lblCaption.Top, dWidth, (lblCaption.Height + (sBorder.BorderWidth * 2))
        
        lblSubCaption.Move lblCaption.Left, UserControl.ScaleHeight + pSpace.Height, lblCaption.Width
    
    'position border
    
        PositionBorder
        
        'sBorder.Visible = False
            
    'set new control size
            
        pItemFade.Move pImage.Left, pImage.Top, pImage.Width, pImage.Height
    
        bFromSelection = False
            
        UserControl.Width = (dWidth * 15)
    
    'set, position, and show icon
    
        Do Until TileIconSet(sIcon) = True
            DoEvents
        Loop
        
    'return
        
        TileSet = dWidth

    'tidy up

        Set pItemImage = Nothing
        Set pResizeImage = Nothing
                
errHandle:
    
    If Err.Number <> 0 Then
    
        'an error has raised
    
        'return error as negative for processing
        
            TileSet = -Err.Number
    
        'clear error
    
            Err.Clear
            
            Resume Next
    
    End If

End Function

Private Function CalculateWidth(lOrigHeight As Single, lOrigWidth As Single) As Single

    Dim dRatio As Single
    
    dRatio = lOrigWidth / lOrigHeight
    
    CalculateWidth = CSng(UserControl.ScaleHeight * dRatio)

End Function

Private Function PositionBorder() As Boolean
    
    If sBorder.BorderWidth > 1 Then
        
        sBorder.Top = 0 '(sBorder.BorderWidth / 2)
        
        sBorder.Left = sBorder.Top
        
        sBorder.Height = UserControl.ScaleHeight ' - (sBorder.BorderWidth / 2)
    
        sBorder.Width = dWidth ' - (sBorder.BorderWidth / 2)
    
    Else
    
        sBorder.Height = UserControl.ScaleHeight
    
        sBorder.Width = dWidth
    
    End If
    
    PositionBorder = True

End Function

Public Property Let TileSelected(ByVal bSelect As Boolean)

    bFromSelection = True

    If bSelect = True Then
    
        UserControl.BackColor = cTileColorSelected
        lblCaption.ForeColor = cCaptionTrue
        
        pItemFade.Visible = False
    
    ElseIf bSelect = False Then
    
        UserControl.BackColor = cTileColorDeselected
        lblCaption.ForeColor = cCaptionFalse
    
        pItemFade.Visible = True
    
    End If

    sBorder.BorderColor = UserControl.BackColor

    bSelected = bSelect
    
    PropertyChanged TileSelected

End Property

Public Property Get TileSelected() As Boolean

    TileSelected = bSelected

End Property

Private Sub lblCaption_Click()

    RaiseEvent Click

End Sub

Private Sub lblSubCaption_Click()

    RaiseEvent Click

End Sub

Private Sub pDefault_Click()

    RaiseEvent Click

End Sub

Private Sub pImage_Click()

    RaiseEvent Click

End Sub

Private Sub pItemFade_Click()

    RaiseEvent Click

End Sub

Private Sub pSpace_Click()

    RaiseEvent Click

End Sub

Private Sub pTrans_Click()

    RaiseEvent Click

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()

    TileSelected = False
    
    sOldHeight = UserControl.ScaleHeight
    sOldWidth = UserControl.ScaleWidth
        
    bFromSelection = False

End Sub

Private Function ProcessResize() As Boolean

    Dim dPC As Double
    Dim cCtl As Control

    With UserControl
    
        If bFromSelection = False Then
        
            If .ScaleHeight <> sOldHeight Then 'And .ScaleWidth <> sOldWidth Then
            
                dPC = (.ScaleHeight / sOldHeight)
            
                For Each cCtl In .Controls
                
                    Do Until ResizeControl(cCtl, dPC) = True
                        DoEvents
                    Loop
                
                Next
            
            End If
           
        End If
    
        sOldHeight = .ScaleHeight
        sOldWidth = .ScaleWidth
              
        bFromSelection = False
                  
    End With

    ProcessResize = True

End Function

Private Function ResizeControl(cCtl As Control, dPC As Double) As Boolean

    Dim dTop As Integer
    Dim dLeft As Integer
    Dim dWidth As Integer
    Dim dHeight As Integer

    If TypeOf cCtl Is Line Then
    
        'nothing
        
    Else
    
        dHeight = cCtl.Height * dPC
        dWidth = cCtl.Width * dPC
        dTop = cCtl.Top * dPC
        dLeft = cCtl.Left * dPC
    
        cCtl.Move dLeft, dTop, dWidth, dHeight
    
    End If
    
    If TypeOf cCtl Is Shape Then
    
        cCtl.BorderWidth = cCtl.BorderWidth * dPC
    
    End If
    
    If TypeOf cCtl Is Label Or TypeOf cCtl Is TextBox Then
    
        cCtl.Font.Size = cCtl.Font.Size * dPC
    
    End If
    
    ResizeControl = True

End Function


Private Sub UserControl_InitProperties()

    'sOldHeight = UserControl.ScaleHeight
    'sOldWidth = UserControl.ScaleWidth
    
    'bFromSelection = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'sOldHeight = UserControl.ScaleHeight
    'sOldWidth = UserControl.ScaleWidth
    
    'bFromSelection = False
    
End Sub

Private Sub UserControl_Resize()

    Do Until ProcessResize = True
        DoEvents
    Loop

End Sub


Public Property Let TileColorDeselected(ByVal cColor As OLE_COLOR)

    cTileColorDeselected = cColor
    
    If bSelected = False Then
        UserControl.BackColor = cColor
    End If
    
    TileSelected = bSelected

End Property

Public Property Let TileColorSelected(ByVal cColor As OLE_COLOR)

    cTileColorSelected = cColor

    If bSelected = True Then
        UserControl.BackColor = cColor
    End If
    
    TileSelected = bSelected

End Property

Public Function ClearTile(Optional bImagesOnly As Boolean = False) As Boolean

    Set pImage.Picture = Nothing
        If bImagesOnly = False Then
        lblCaption.Caption = ""
        lblSubCaption.Caption = ""
        TileSelected = False
    End If

End Function

Private Function IsURL(sPath As String) As Boolean

    Dim bTemp As Boolean
    
    bTemp = False
        
    If Mid(sPath, 1, 3) = "htt" Or Mid(sPath, 1, 3) = "www" Then
        bTemp = True
    End If
    
    IsURL = bTemp

End Function

Public Function TileResetImage(sItemImage As String) As Boolean

    If sItemImage <> "" Then
        If IsURL(sItemImage) = False Then
            If Dir(sItemImage) <> "" Then
                Set pImage.Picture = LoadPictureGDIplus(sItemImage, , False)
            End If
        Else
            On Error GoTo errHandle
            Set pImage.Picture = LoadPictureGDIplus(sItemImage, , False)
        End If
    Else
        Set pImage.Picture = Nothing
    End If
    
    TileResetImage = True

errHandle:
    
    If Err.Number <> 0 Then
    
        'an error has raised
        
            Debug.Print Err.Number
    
        'clear error
    
            Err.Clear
            
            Resume Next
    
    End If
    
End Function

Public Property Let TileCaption(sNewCaption As String)

    Dim lOldHeight As Double
    
    lOldHeight = lblCaption.Height

    lblCaption.Caption = sNewCaption
    
    'If lblCaption.Height <> lOldHeight Then
    
        lblCaption.Top = (pImage.Top + pImage.Height) - (lblCaption.Height + (sBorder.BorderWidth * 2))
        
        pTrans.Move 0, lblCaption.Top, UserControl.ScaleWidth, (lblCaption.Height + (sBorder.BorderWidth * 2))
        
        'lblSubCaption.Move lblCaption.Left, UserControl.ScaleHeight + pSpace.Height, lblCaption.Width
    
    'End If

End Property

Public Property Let TileSubCaption(sNewCaption As String)

    lblSubCaption.Caption = sNewCaption
    
    lblSubCaption.Visible = True
    
End Property

Public Property Let Font(sFont As String)

    lblCaption.Font.Name = sFont
    lblSubCaption.Font.Name = sFont

End Property

Private Sub UserControl_Show()
    
    UserControl.Extender.ZOrder (vbSendToBack)

End Sub

Public Function TileIconClear() As Boolean

        Set pIcon.Picture = Nothing
    
        pIcon.Visible = False
    
        TileIconClear = True
    
End Function

Public Function TileIconSet(sIcon As String) As Boolean

    If sIcon <> "" Then
    
        pIcon.Left = UserControl.ScaleWidth - pIcon.Width - pIcon.Top
        
        pIcon.Picture = LoadPictureGDIplus(sIcon)
        
        pIcon.Visible = True
        
    Else
    
        Do Until TileIconClear = True
            DoEvents
        Loop
    
    End If
    
    TileIconSet = True
    
End Function

