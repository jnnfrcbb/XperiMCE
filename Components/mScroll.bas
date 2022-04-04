Attribute VB_Name = "mScroll"

Private lRowSelected As Long
Private lTileSelected As Long

Private Const cRowTitleSelected = vbWhite
Private Const cRowTitleDeselected = &H808080

Private RowLoaded() As Boolean
Private lScroll() As Long
Private Const lScrollSpeed As Long = 6

Private Function CalculateScrollUp(lRow As Long, lNewRow As Long, lOldRow As Long) As Long

    ' on error resume next
    
    Dim t As Long
        
    If lRow = lNewRow Then
    
        t = lblRow(lRow).Top - (pRowSpace.Top - lblRow(lRow).Height)
        
    ElseIf lRow < lNewRow Then
    
        t = lblRow(lRow).Top - (pRowSpace.Top - (lblRow(lRow).Height * ((lNewRow - lRow) + 1))) + (lblRow(lRow).Height / 2)
    
    ElseIf lRow > lNewRow Then
    
        t = lblRow(lRow).Top - (cRow(lNewRow).Top + cRow(lNewRow).Height + (lblRow(lRow).Height * (lRow - lNewRow)))
        
    End If
    
    CalculateScrollUp = t

End Function

Private Function CalculateScrollDown(lRow As Long, lNewRow As Long, lOldRow As Long) As Long

    ' on error resume next
    
    Dim t As Long
    
    t = 0
    
    If lRow = lNewRow Then
    
        t = (pRowSpace.Top - lblRow(lRow).Height) - lblRow(lRow).Top
        
    ElseIf lRow < lNewRow Then
    
        t = (pRowSpace.Top - (lblRow(lRow).Height * ((lNewRow - lRow) + 1))) - lblRow(lRow).Top - (lblRow(lRow).Height / 2)
    
    ElseIf lRow > lNewRow Then
    
        t = (cRow(lNewRow).Top + cRow(lNewRow).Height + (lblRow(lRow).Height * (lRow - lNewRow))) - lblRow(lRow).Top
        
    End If
    
    CalculateScrollDown = t

End Function

Private Function LoadRow(Index As Long, Title As String) As Boolean

    ' on error resume next
    
    If Index > cRow.UBound Then
    
        ReDim Preserve RowLoaded(Index)
        ReDim Preserve lScroll(Index)
    
        Load cRow(Index)
        Load lblRow(Index)
    
        cRow(Index).Top = pRowSpace.Top
                    
        If Index < lRowSelected Then
        
            lblRow(Index).Top = lblRow(Index + 1).Top - lblRow(Index).Height
            
        ElseIf Index = lRowSelected Then
        
            lblRow(Index).Top = pRowSpace.Top
                    
        ElseIf Index = lRowSelected + 1 Then
        
            lblRow(Index).Top = cRow(lRowSelected).Top + cRow(lRowSelected).Height + lblRow(Index).Height
            
        ElseIf Index > lRowSelected + 1 Then
            
            lblRow(Index).Top = lblRow(Index - 1).Top + lblRow(Index - 1).Height
            
        End If
        
        cRow(Index).TileColorSelected = cRow(0).TileColorSelected
    
        lblRow(Index).ZOrder 0
        
        lblRow(Index).Caption = Title
    
        lblRow(Index).Visible = True
    
    End If
    
    RowLoaded(Index) = False
    
    pMenuHolder.ZOrder 0
    
    LoadRow = True

End Function

Private Function LoadRowTiles(Index As Integer) As Boolean

    ' on error resume next
    
    'Loading True
    
    Dim i As Long
    Dim lTag As Long

    For i = 0 To cRow(Index).TileCount
        
        DoEvents
        
        lTag = CLng(cRow(Index).TileTag(i))
        
        'Do Until .crow(Index).TileSet(i, , , , , CStr(lTag)) <> 0
        '    DoEvents
        'Loop
        
    Next
    
    RowLoaded(Index) = True

    LoadRowTiles = True

    'Loading False

End Function

Private Function SelectRow(lNewRow As Integer, Optional lNewTile As Long = 0) As Boolean

    On Error Resume Next
    
    Dim lOldRow As Long
    Dim r As Long
    
    Debug.Print "NewTile: " & lNewTile
    
    lOldRow = lRowSelected

    If lNewRow <> lOldRow Then
        
        If lNewRow >= 0 And lNewRow <= cRow.UBound Then
        
            If lNewRow > 0 Then
                lblAppTitle.Visible = False
            End If
        
            If RowLoaded(lNewRow) = True Then
            
                cRow(lRowSelected).Deselect
                
                cRow(lRowSelected).Visible = False
                
                lblRow(lRowSelected).ForeColor = cRowTitleDeselected
                    
                lRowSelected = lNewRow
                
                If lRowSelected < lOldRow Then
                
                    For r = 0 To lblRow.UBound
                        DoEvents
                        lScroll(r) = CalculateScrollDown(r, CLng(lNewRow), lOldRow) / lScrollSpeed
                    Next
                
                    tmrScrollRowsDown.Enabled = True
                    
                    Do Until tmrScrollRowsDown.Enabled = False
                        DoEvents
                    Loop
                
                ElseIf lRowSelected > lOldRow Then
                
                    For r = 0 To lblRow.UBound
                        DoEvents
                        lScroll(r) = CalculateScrollUp(r, CLng(lNewRow), lOldRow) / lScrollSpeed
                    Next
                
                    tmrScrollRowsUp.Enabled = True
                    
                    Do Until tmrScrollRowsUp.Enabled = False
                        DoEvents
                    Loop
            
                End If
                    
                lblRow(lRowSelected).ForeColor = cRowTitleSelected
                
                cRow(lRowSelected).Visible = True
                    
                cRow(lRowSelected).TileSelected(lNewTile) = True
                    
            Else
            
                Do Until LoadRowTiles(lNewRow) = True
                    DoEvents
                Loop
                
                Do Until SelectRow(lNewRow, lNewTile) = True
                    DoEvents
                Loop
                
            End If
            
            If lRowSelected = 0 Then
                lblAppTitle.Visible = True
            Else
                lblAppTitle.Visible = False
            End If
            
            RaiseEvent Message("sound##select", Nothing)
          
        End If
            
    Else
    
        lblRow(lRowSelected).ForeColor = cRowTitleSelected
                
        cRow(lRowSelected).TileSelected(lNewTile) = True
        
        cRow(lRowSelected).Visible = True
        
        lblRow(lRowSelected).Visible = True
        
                    
    End If
              
    SelectRow = True
                    
End Function


Private Function ClearTiles() As Boolean

    ' on error resume next
    
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


