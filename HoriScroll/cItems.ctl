VERSION 5.00
Begin VB.UserControl cItems 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00181818&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   ControlContainer=   -1  'True
   MaskColor       =   &H00181818&
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1440
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   5520
      Top             =   1980
   End
   Begin VB.Timer tmrSelect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7140
      Top             =   1980
   End
   Begin VB.Timer tmrLeftStop 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6060
      Top             =   2460
   End
   Begin VB.Timer tmrRightStop 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6600
      Top             =   2460
   End
   Begin VB.Timer tmrScrollRight 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   1500
   End
   Begin VB.Timer tmrScrollLeft 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6060
      Top             =   1500
   End
   Begin prjHoriSccroll.cItem cItem 
      Height          =   3300
      Index           =   0
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
      _extentx        =   5292
      _extenty        =   5821
   End
   Begin VB.Label pItemSpace 
      Caption         =   "Label1"
      Height          =   3150
      Left            =   4620
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Label pItemRight 
      Caption         =   "Label1"
      Height          =   1275
      Left            =   13200
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label pItemSelected 
      Caption         =   "Label1"
      Height          =   3750
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label pItemDefault 
      Caption         =   "Label1"
      Height          =   3300
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "cItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type tTile
    Width As Single
    Caption As String
    SubCaption As String
    ImageLocation As String
    IconLocation As String
    Loaded As Boolean
    Selected As Boolean
    Tag As String
    ImageLoaded As Boolean
    BlankImage As Boolean
End Type
Private Tiles() As tTile

Private sItemWidth() As Single
Private bMultiSelect As Boolean

Private ResizeOld As Double
Private ResizeNew As Double

Private lScroll As Single
Private lScrollItem As Long

Private cTileColorSelected As OLE_COLOR
Private cTileColorDeselected As OLE_COLOR

Private lRightStop As Long
Private lLeftStop As Long

Private lSelect As Long

Private lPrevSelected As Long
Private lNewSelect As Long

Private Const lScrollTick As Long = 5

Event TileClick(TileIndex As Integer)
Event Click()
Event TileSelectTrue()
Event TileSelectFalse()
Event TileDelete()
Event TileInsert()

Private Sub cItem_Click(Index As Integer)

    RaiseEvent TileClick(Index)

End Sub

Private Sub tmrRightStop_Timer()

    lRightStop = lRightStop + 1

    Select Case lRightStop

        Case 2
            
            PrepareLeftScroll lScrollItem

        Case 3 To 7

            If ((cItem(lScrollItem).Left + cItem(lScrollItem).Width) >= pItemRight.Left) Then
        
                cItem(lScrollItem).Left = cItem(lScrollItem).Left - lScroll
            
            End If
            
            PositionLeftScroll
        
        Case 8
        
            SelectItemTrue lNewSelect
        
            lRightStop = 0
                    
            tmrRightStop.Enabled = False
    
    End Select
        
End Sub

Private Sub tmrLeftStop_Timer()

    lLeftStop = lLeftStop + 1

    Select Case lLeftStop

        Case 2
            
            PrepareRightScroll lScrollItem

        Case 3 To 7

            If cItem(lScrollItem).Left < pItemSelected.Left Then
        
                cItem(lScrollItem).Left = cItem(lScrollItem).Left + lScroll
            
            End If
            
            PositionRightScroll
        
        Case 8
        
            SelectItemTrue lNewSelect
        
            lLeftStop = 0
        
            tmrLeftStop.Enabled = False
    
    End Select
        

End Sub

Private Sub tmrScrollLeft_Timer()

    Static i As Integer
    
    i = i + 1
    
    cItem(lScrollItem).Left = cItem(lScrollItem).Left - lScroll
    
    Do Until PositionLeftScroll = True
        DoEvents
    Loop
    
    If i = lScrollTick Or cItem(lScrollItem).Left > pItemDefault.Left Then
    
        i = 0
        
        tmrScrollLeft.Enabled = False
        
    End If

End Sub

Private Sub tmrScrollRight_Timer()

    Static i As Integer
    
    i = i + 1
    
    cItem(lScrollItem).Left = cItem(lScrollItem).Left + lScroll
    
    Do Until PositionRightScroll = True
        DoEvents
    Loop
    
    If i = lScrollTick Or (cItem(lScrollItem).Left + cItem(lScrollItem).Width) <= pItemRight.Left Then
    
        i = 0
        
        tmrScrollRight.Enabled = False
        
    End If
End Sub

Private Sub tmrSelect_Timer()

    lSelect = lSelect + 1
    
    If lSelect = 5 Then
    
        SelectItemTrue lPrevSelected
    
        tmrSelect.Enabled = False
    
    End If

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()

    cTileColorSelected = &H4551CF
    
    cTileColorDeselected = &H646464 ' &H736958
    
    cItem(0).Left = pItemDefault.Left
    
    cItem(0).Top = pItemDefault.Top
    
    cItem(0).TileColorSelected = cTileColorSelected
    
    cItem(0).TileColorDeselected = cTileColorDeselected
    
    lPrevSelected = 0
    
    ReDim Tiles(0)

    ResizeOld = UserControl.ScaleWidth

End Sub

Public Function TileSet(Index As Long, Optional sCaption As String = "", Optional sSubcaption As String = "", Optional sImage As String = "", Optional bItemBlankImage As Boolean = False, Optional sItemTag As String = "", Optional sIcon As String = "") As Single

    On Error Resume Next

    Dim i As Long
    Dim sRes As Single
    Dim lLeft As Single
    Dim lWidth As Single
    Dim lPrevWidth As Long
    
    sRes = 0
    
    If Index > cItem.UBound Then
    
        For i = cItem.UBound + 1 To Index
    
            ReDim Preserve Tiles(i)
    
            Load cItem(i)
            
            cItem(i).Left = cItem(i - 1).Left + cItem(i - 1).Width + pItemSpace.Width
        
            cItem(i).Visible = True
        
        Next
        
        lLeft = cItem(cItem.UBound - 1).Left + cItem(cItem.UBound - 1).Width + pItemSpace.Width
        
    Else
    
        lPrevWidth = cItem(Index).Width
        lLeft = cItem(Index).Left
        
        cItem(Index).Visible = False
    
        If cItem(Index).TileSelected = True Then
            TileSelected(Index) = False
        End If
        
    End If
    
    cItem(Index).TileColorDeselected = cTileColorDeselected
    cItem(Index).TileColorSelected = cTileColorSelected
    
    sRes = cItem(Index).TileSet(sCaption, sSubcaption, sImage, bItemBlankImage, sIcon)

    If sRes = 0 Then
        
        lWidth = pItemDefault.Width
        
    ElseIf sRes > 0 Then
        
        lWidth = sRes
            
    End If
    
    Tiles(Index).Caption = sCaption
    Tiles(Index).SubCaption = sSubcaption
    Tiles(Index).ImageLocation = sImage
    Tiles(Index).Tag = sItemTag
    Tiles(Index).ImageLoaded = True
    Tiles(Index).BlankImage = bItemBlankImage
    
    cItem(Index).Move lLeft, pItemDefault.Top, lWidth, pItemDefault.Height
    
    cItem(Index).Visible = True
    Tiles(Index).Loaded = True
    
    Do Until RepositionTiles(Index, lPrevWidth) = True
        DoEvents
    Loop
            
    TileSet = sRes

End Function

Public Property Let TileSelected(Index As Long, bSelect As Boolean)

    lRightStop = 0
            
    lLeftStop = 0
    
    lNewSelect = Index
    
    Dim i As Long
        
    If Index >= cItem.LBound And Index <= cItem.UBound Then
    
        If bMultiSelect = False Then
        
            Deselect
            
        End If

        If Index < lPrevSelected Then
    
            'tmrLeftStop.Enabled = True
        
            If cItem(Index).Left < pItemDefault.Left Then
            
                Do Until PrepareRightScroll(Index) = True
                    DoEvents
                Loop
                
                tmrScrollRight.Enabled = True
            
                'Do Until tmrScrollRight.Enabled = False
                '    DoEvents
                'Loop
            
            End If
            
                lLeftStop = 0
    
                tmrLeftStop.Enabled = True
    
        ElseIf Index > lPrevSelected Then
    
            'tmrRightStop.Enabled = True
    
            If (cItem(Index).Left + cItem(Index).Width) >= pItemRight.Left Then  'UserControl.ScaleWidth Then 'pItemRight.Left Then
            
                Do Until PrepareLeftScroll(Index) = True
                    DoEvents
                Loop
                
                tmrScrollLeft.Enabled = True
            
                'Do Until tmrScrollLeft.Enabled = False
                '    DoEvents
                'Loop
            
            End If
            
                lRightStop = 0
            
                tmrRightStop.Enabled = True
    
        Else
        
            SelectItemTrue Index
        
        End If

        If bSelect = True Then

            'lSelect = 0
        
            'tmrSelect.Enabled = True

            'SelectItemTrue Index
        
            cItem(Index).TileSelected = True
        
        ElseIf bSelect = False Then
        
            SelectItemFalse Index
    
        End If
    
        Tiles(Index).Selected = bSelect
    
    End If
    
    lPrevSelected = Index

    'Do Until SetTileImages = True
    '    DoEvents
    'Loop

End Property

Private Function SetTileImages() As Boolean

    Dim i As Long
    
    For i = 0 To cItem.UBound
        If (cItem(i).Left > (-cItem(i).Width - 100)) And (cItem(i).Left < (UserControl.ScaleWidth + 100)) Then
            If Tiles(i).ImageLoaded = False Then
                cItem(i).TileResetImage Tiles(i).ImageLocation
                Tiles(i).ImageLoaded = True
            End If
        Else
            If Tiles(i).ImageLoaded = True Then
                cItem(i).ClearTile True
                Tiles(i).ImageLoaded = False
            End If
        End If
    Next
    
    SetTileImages = True
    
End Function

Public Property Get TileSelected(Index As Long) As Boolean
    
    If Index >= cItem.LBound And Index <= cItem.UBound Then
        
        TileSelected = cItem(Index).TileSelected
        
    End If

End Property

Public Function Deselect()

    Static i As Long
    
    For i = 0 To cItem.UBound
        If Tiles(i).Selected = True Then
            SelectItemFalse i
            Tiles(i).Selected = False
        End If
    Next

End Function

Public Property Let Transparent(ByVal bTransparent As Boolean)

    If bTransparent = True Then
    
        UserControl.BackStyle = 0
    
    ElseIf bTransparent = False Then

        UserControl.BackStyle = 1
        
    End If
    
    PropertyChanged Transparent

End Property

Public Property Get Transparent() As Boolean

    If UserControl.BackStyle = 0 Then
    
        Transparent = True
        
    ElseIf UserControl.BackStyle = 1 Then
    
        Transparent = False
        
    End If
    
End Property

Public Property Let MultiSelect(ByVal bMulti As Boolean)

    bMultiSelect = bMulti
    
    PropertyChanged MultiSelect

End Property

Public Property Get MultiSelect() As Boolean

    MultiSelect = bMultiSelect

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    If PropBag.ReadProperty("Transparent", False) = True Then
    
        UserControl.BackStyle = 0
    
    Else
    
        UserControl.BackStyle = 1
    
    End If

End Sub
Private Sub UserControl_Resize()

    'On Error Resume Next

    ResizeForm

End Sub

Private Function ResizeForm()

    'On Error Resume Next

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

    'On Error Resume Next

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


Private Sub UserControl_Show()

    UserControl.Extender.ZOrder (vbSendToBack)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    If UserControl.BackStyle = 0 Then
        PropBag.WriteProperty "Transparent", True
    Else
        PropBag.WriteProperty "Transaprent", False
    End If

End Sub

Private Function SelectItemFalse(Index As Long) As Boolean
    
    cItem(Index).TileSelected = False
    
    Tiles(Index).Selected = False

    cItem(Index).Visible = False
    
    cItem(Index).Move cItem(Index).Left, pItemDefault.Top, cItem(Index).Width, pItemDefault.Height
    
    cItem(Index).Visible = True
    
    SelectItemFalse = True

End Function

Private Function SelectItemTrue(Index As Long) As Boolean

    cItem(Index).TileSelected = True
    
    Tiles(Index).Selected = True

    cItem(Index).Visible = False
    
    cItem(Index).Move cItem(Index).Left, pItemSelected.Top, cItem(Index).Width, pItemSelected.Height
    
    cItem(Index).Visible = True
    
    SelectItemTrue = True
    
End Function

Private Function PrepareLeftScroll(Index As Long) As Boolean
    
    If Index <= cItem.UBound Then
        
        lScroll = ((cItem(Index).Left - (pItemRight.Left - cItem(Index).Width)) / lScrollTick)
        
        lScrollItem = Index
        
    End If
    
    PrepareLeftScroll = True

End Function

Private Function PositionLeftScroll() As Boolean

    Dim i As Integer
    Dim c As Integer
    Dim lItem As Long

    If lScrollItem < cItem.UBound Then
        
        cItem(lScrollItem + 1).Left = (cItem(lScrollItem).Left + (cItem(lScrollItem).Width + pItemSpace.Width))

    End If

    If lScrollItem > 0 Then
    
        lItem = lScrollItem - 1
    
        For i = 0 To lScrollItem - 1
    
            DoEvents
    
            If (cItem(lItem).Left > -cItem(lItem).Width) Then
        
                cItem(lItem).Left = (cItem(lItem + 1).Left - (cItem(lItem).Width + pItemSpace.Width))
    
            End If
            
            lItem = lItem - 1
        
        Next

    End If
    
    PositionLeftScroll = True

End Function


Private Function PrepareRightScroll(Index As Long) As Boolean

    If Index > cItem.UBound Then

        Index = cItem.UBound
        
    End If

    lScroll = ((pItemDefault.Left + (-cItem(Index).Left)) / lScrollTick)

    lScrollItem = Index

    PrepareRightScroll = True

End Function

Private Function PositionRightScroll() As Boolean

    Static i As Integer
    
    If lScrollItem > 0 Then
    
        cItem(lScrollItem - 1).Left = cItem(lScrollItem).Left - ((cItem(lScrollItem - 1).Width + pItemSpace.Width))
    
    End If
    
    If lScrollItem < cItem.UBound Then
    
        For i = lScrollItem + 1 To cItem.UBound
        
            DoEvents
        
            If i <= cItem.UBound Then
            
                If cItem(i).Left < UserControl.ScaleWidth Then
                
                    cItem(i).Left = (cItem(i - 1).Left + (cItem(i - 1).Width + pItemSpace.Width))
                    
                End If
                
            End If

        Next
    
    End If

    PositionRightScroll = True

End Function

Public Property Get TileCount() As Long

    If Tiles(0).Loaded = True Then

        TileCount = cItem.UBound
        
    Else
        
        TileCount = -1
    
    End If

End Property

Public Property Let TileColorDeselected(ByVal cColor As OLE_COLOR)

    Static i As Integer
    
    For i = 0 To cItem.UBound
        cItem(i).TileColorDeselected = cColor
    Next

    cTileColorDeselected = cColor

End Property

Public Property Get TileColorDeselected() As OLE_COLOR

    TileColorDeselected = cTileColorDeselected

End Property

Public Property Let TileColorSelected(ByVal cColor As OLE_COLOR)

    Static i As Integer
    
    For i = 0 To cItem.UBound
        cItem(i).TileColorSelected = cColor
    Next
    cTileColorSelected = cColor

End Property

Public Property Get TileColorSelected() As OLE_COLOR

    TileColorSelected = cTileColorSelected

End Property
Public Property Get TileCurrent() As Long

    TileCurrent = lPrevSelected
    
End Property

Public Function TileSelectRight() As Boolean

    If lPrevSelected < cItem.UBound Then
        
        RaiseEvent TileSelectTrue

        TileSelected(lPrevSelected + 1) = True
        
    Else
    
        RaiseEvent TileSelectFalse
        
    End If
    
    TileSelectRight = True

End Function

Public Function TileSelectLeft() As Boolean

    If lPrevSelected > 0 Then
    
        RaiseEvent TileSelectTrue

        TileSelected(lPrevSelected - 1) = True
        
    Else
    
        RaiseEvent TileSelectFalse
        
    End If
    
    TileSelectLeft = True

End Function

Public Function TileJumpRight() As Boolean

    If lPrevSelected < cItem.UBound - 4 Then

        RaiseEvent TileSelectTrue
                            
        Do Until TileSelected(lPrevSelected + 5) = True
        
            DoEvents
        
        Loop
        
    Else
    
        RaiseEvent TileSelectFalse
                            
        Do Until TileSelected(cItem.UBound) = True
        
            DoEvents
        
        Loop
        
    End If
    
    TileJumpRight = True

End Function

Public Function TileJumpLeft() As Boolean

    If lPrevSelected > 4 Then

        RaiseEvent TileSelectTrue
                            
        Do Until TileSelected(lPrevSelected - 5) = True
        
            DoEvents
        
        Loop
        
    Else
    
        RaiseEvent TileSelectFalse
                            
        Do Until TileSelected(0) = True
        
            DoEvents
        
        Loop
        
    End If
    
    TileJumpLeft = True

End Function

Public Property Get TileTag(Index As Long) As String

    TileTag = Tiles(Index).Tag

End Property

Public Property Let TileTag(Index As Long, ByVal sTag As String)

    Tiles(Index).Tag = sTag

End Property

Public Function ClearTiles(Optional bImagesOnly As Boolean = False) As Boolean

    Deselect

    Dim i As Integer
            
    If bImagesOnly = True Then
    
        For i = 0 To cItem.UBound
            cItem(i).ClearTile True
            Tiles(i).ImageLoaded = False
        Next
    
    Else
    
        ReDim Tiles(0)
        
        Tiles(0).Loaded = False
    
        cItem(0).Visible = False
    
        cItem(0).ClearTile False
        
        If cItem.UBound > 0 Then
        
            For i = 1 To cItem.UBound
            
                Unload cItem(i)
            
            Next
            
        End If
        
        cItem(0).Left = pItemDefault.Left
        
        lPrevSelected = 0
        lNewSelect = 0
        lScroll = 0
        lScrollItem = 0
    
    End If

    ClearTiles = True

End Function

Private Function RepositionTiles(Index As Long, Optional lPrevWidth As Long = -1) As Boolean

    Dim i As Integer
    
    If lPrevWidth > -1 Then
        
        If cItem(Index).Width <> lPrevWidth Then
            If (cItem(Index).Left > -cItem(Index).Width) And (cItem(Index).Left < UserControl.ScaleWidth) Then
                If Index < cItem.UBound Then
                    For i = Index + 1 To cItem.UBound
                       DoEvents
                        If cItem(i).Left < UserControl.ScaleWidth Then
                            cItem(i).Left = cItem(i - 1).Left + cItem(i - 1).Width + pItemSpace.Width
                        Else
                            Exit For
                        End If
                    Next
               End If
            End If
        End If
        
    Else
    
        If (cItem(Index).Left > -cItem(Index).Width) And (cItem(Index).Left < UserControl.ScaleWidth) Then
            If Index < cItem.UBound Then
                For i = Index + 1 To cItem.UBound
                   DoEvents
                    If cItem(i).Left < UserControl.ScaleWidth Then
                        cItem(i).Left = cItem(i - 1).Left + cItem(i - 1).Width + pItemSpace.Width
                    Else
                        Exit For
                    End If
                Next
           End If
        End If
    
    End If
    
    RepositionTiles = True

End Function

Public Function TileDelete(Index As Long) As Boolean

    Dim i As Integer
    Dim bTemp As Boolean
    
    Dim lTemp As Long
    
    lTemp = lPrevSelected

    Deselect
    
    If Index < cItem.UBound Then
        
        For i = Index To cItem.UBound - 1
            
            Tiles(i).Width = Tiles(i + 1).Width
            Tiles(i).Caption = Tiles(i + 1).Caption
            Tiles(i).SubCaption = Tiles(i + 1).SubCaption
            Tiles(i).ImageLocation = Tiles(i + 1).ImageLocation
            Tiles(i).IconLocation = Tiles(i + 1).IconLocation
            Tiles(i).Loaded = Tiles(i + 1).Loaded
            Tiles(i).Selected = Tiles(i + 1).Selected
            Tiles(i).Tag = Tiles(i + 1).Tag
            Tiles(i).BlankImage = Tiles(i + 1).BlankImage
            
            
            TileSet CLng(i), Tiles(i).Caption, Tiles(i).SubCaption, Tiles(i).ImageLocation, Tiles(i).BlankImage, Tiles(i).Tag, Tiles(i).IconLocation
            
        Next
    
    End If
    
    RaiseEvent TileDelete
    
    'ReDim Preserve Tiles(UBound(Tiles()) - 1)
    
    Unload cItem(cItem.UBound)

    If Index < cItem.UBound Then
        Do Until RepositionTiles(Index) = True
            DoEvents
        Loop
    End If
    
    If lPrevSelected = Index Then
        TileSelected(lPrevSelected) = True
    ElseIf lPrevSelected < cItem.Count Then
        TileSelected(lPrevSelected - 1) = True
    Else
        TileSelected(cItem.Count) = True
    End If
        
    TileDelete = True
        
End Function

Public Function TileInsert(Index As Long, Optional sCaption As String = "", Optional sSubcaption As String = "", Optional sImage As String = "", Optional bItemBlankImage As Boolean = False, Optional sItemTag As String = "") As Single

    Dim i As Integer
    Dim t As Integer
    
    Dim lTemp As Long
    
    lTemp = lPrevSelected
    
    Deselect
    
    ReDim Preserve Tiles(UBound(Tiles()) + 1)
    
    t = UBound(Tiles())
    
    For i = Index + 1 To cItem.UBound + 1
    
        Tiles(t).Width = Tiles(t - 1).Width
        Tiles(t).Caption = Tiles(t - 1).Caption
        Tiles(t).SubCaption = Tiles(t - 1).SubCaption
        Tiles(t).ImageLocation = Tiles(t - 1).ImageLocation
        Tiles(t).IconLocation = Tiles(t - 1).IconLocation
        Tiles(t).Loaded = Tiles(t - 1).Loaded
        Tiles(t).Selected = Tiles(t - 1).Selected
        Tiles(t).Tag = Tiles(t - 1).Tag
            
        TileSet CLng(t), Tiles(t).Caption, Tiles(t).SubCaption, Tiles(t).ImageLocation, , Tiles(t).Tag
            
        t = t - 1
    
    Next

    RaiseEvent TileInsert

    TileSet Index, sCaption, sSubcaption, sImage, bItemBlankImage, sItemTag
    
    Do Until RepositionTiles(Index) = True
        DoEvents
    Loop
    
    If lTemp >= Index Then
    
        lTemp = lTemp + 1
        
    End If
    
    TileSelected(lTemp) = True
    
    
    TileInsert = True

End Function

Public Property Get TileCaption(Index As Long) As String

    TileCaption = Tiles(Index).Caption

End Property

Public Property Let TileCaption(Index As Long, sNewCaption As String)

    If Index >= 0 And Index <= cItem.UBound Then
    
        Tiles(Index).Caption = sNewCaption
        
        cItem(Index).TileCaption = sNewCaption

    End If

End Property


Public Property Get TileSubCaption(Index As Long) As String

    TileSubCaption = Tiles(Index).SubCaption

End Property

Public Property Let TileSubCaption(Index As Long, sNewCaption As String)

    If Index >= 0 And Index <= cItem.UBound Then
    
        Tiles(Index).SubCaption = sNewCaption
        
        cItem(Index).TileSubCaption = sNewCaption

    End If

End Property

Public Property Let Font(sFont As String)

    Dim i As Long
    For i = 0 To cItem.UBound
        cItem(i).Font = sFont
    Next

End Property

Public Function TileIconClear(Index As Long) As Boolean

        cItem(Index).TileIconClear
    
        TileIconClear = True
    
End Function

Public Function TileIconSet(Index As Long, sIcon As String) As Boolean

    cItem(Index).TileIconSet sIcon
    
    TileIconSet = True
    
End Function


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

Public Function TilesFolded(bFolded As Boolean) As Boolean

    Dim i As Integer
    Dim t As Integer
    
    If bFolded = True Then
        t = cItem.UBound
    Else
        t = 0
    End If

    For i = 0 To cItem.UBound
        cItem(t).Visible = Not bFolded
        Do Until QuickWait(25) = True
            DoEvents
        Loop
        If bFolded = True Then
            t = t - 1
        Else
            t = t + 1
        End If
    Next
    
    TilesFolded = True

End Function



