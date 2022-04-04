VERSION 5.00
Object = "{79A8D05F-8453-4684-8B83-13A4A8B0D380}#12.0#0"; "prjHoriSccroll.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmTest"
   ClientHeight    =   8060
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   14720
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8060
   ScaleWidth      =   14720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2940
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin prjHoriSccroll.cItems cTiles 
      Height          =   3735
      Left            =   120
      Top             =   180
      Width           =   14415
      _ExtentX        =   25418
      _ExtentY        =   6579
      Transaprent     =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   555
      Left            =   6000
      TabIndex        =   2
      Top             =   1620
      Width           =   1515
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    cTiles.TileInsert 0, "Test insert", "Subcaption", "D:\10.jpg"

End Sub

Private Sub Command2_Click()

    cTiles.TileDelete 0

End Sub

Private Sub Command3_Click()

    cTiles.TileSubCaption(0) = "Test"

End Sub

Private Sub cTiles_TileClick(TileIndex As Integer)

    cTiles.TileSelected(CLng(TileIndex)) = True

End Sub

Private Sub Form_Load()

    cTiles.Transparent = True

    cTiles.TileSet 0, "Caption 1", "SubCaption", "D:\1.png" ', , , "D:\2.jpg"
    cTiles.TileSet 1, "Caption 2", "SubCaption", "D:\10.jpg" '"https://lh4.googleusercontent.com/-3peEbPjV6LM/AAAAAAAAAAI/AAAAAAAAAAA/H7rurEVe-Wg/s64-c/photo.jpg"
    cTiles.TileSet 2, "Caption 3", "SubCaption", "G:\Video\Films\10 Things I Hate About You\folder.jpg"
    cTiles.TileSet 3, "Caption 4", "SubCaption", "D:\1.jpg"

    cTiles.ZOrder 1
    
End Sub
