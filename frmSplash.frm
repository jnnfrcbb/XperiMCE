VERSION 5.00
Object = "{44D09AE3-1847-41E9-B1EF-890580211EC2}#1.0#0"; "AlphaImage.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing"
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
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   14400
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pProgressBack 
      Height          =   435
      Left            =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   767
      Image           =   "frmSplash.frx":8D25A
      Attr            =   513
      Effects         =   "frmSplash.frx":8DE29
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pLoading 
      Height          =   1215
      Left            =   6630
      Top             =   3450
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2143
      Image           =   "frmSplash.frx":8DE41
      Settings        =   19200
      Attr            =   1537
      Effects         =   "frmSplash.frx":943F1
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
      BackColor       =   0
      Attr            =   513
      Effects         =   "frmSplash.frx":94409
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ResizeOld As Double
Private ResizeNew As Double

Private Sub Form_Load()

    On Error Resume Next

    ResizeOld = Me.ScaleWidth

    Dim cControl As Control

    For Each cControl In Me.Controls
        
        If TypeOf cControl Is Label Then
        
            cControl.UseMnemonic = False
        
        End If
    
    Next

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    ResizeForm

End Sub

Private Function ResizeForm()

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

Private Sub Form_Activate()

    On Error Resume Next

    'frmSplash.pLoading.Left = (frmSplash.Width / 2) - (frmSplash.pLoading.Width / 2)

    'frmSplash.pLoading.Top = (frmSplash.Height / 2) - (frmSplash.pLoading.Height / 2)
        
    frmSplash.pLoading.Visible = True
    
End Sub

