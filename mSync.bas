Attribute VB_Name = "mSync"
Public Type tSyncItem
    File As String
    Title As String
    Artist As String
    Thumb As String
End Type
Public SyncItems() As tSyncItem

Public Function SyncAdd(sFile As String, sTitle As String, sThumb As String, Optional sArtist As String = "") As Boolean

    Dim sNotify As String

    If SyncItems(UBound(SyncItems())).File <> "" Then
        ReDim Preserve SyncItems(UBound(SyncItems()) + 1)
    End If
    
    SyncItems(UBound(SyncItems())).File = sFile
    SyncItems(UBound(SyncItems())).Title = sTitle
    SyncItems(UBound(SyncItems())).Artist = sArtist
    SyncItems(UBound(SyncItems())).Thumb = sThumb

    sNotify = "Sync | Added | " & sTitle
    
    If sArtist <> "" Then
        sNotify = sNotify & " | " & sArtist
    End If
    
    Notify sNotify
    
    SyncAdd = True

End Function

Public Function SyncRemove(sFile As String, sTitle As String, sThumb As String, Optional sArtist As String = "") As Boolean

    Dim sNotify As String
    Dim i As Long
    
    If UBound(SyncItems()) > 0 Then
        
        For i = Index To UBound(SyncItems()) - 1
        
            SyncItems(i).File = SyncItems(i + 1).File
            SyncItems(i).Title = SyncItems(i + 1).Title
            SyncItems(i).Artist = SyncItems(i + 1).Artist
            SyncItems(i).Thumb = SyncItems(i + 1).Thumb
        Next
    
        ReDim Preserve SyncItems(UBound(SyncItems()) - 1)
    
    Else
    
        SyncClear
    
    End If
    
    sNotify = "Sync | Removed | " & sTitle
    
    If sArtist <> "" Then
        sNotify = sNotify & " | " & sArtist
    End If

    Notify sNotify
    
    SyncRemove = True

End Function

Public Function SyncStatus(sFile As String) As Long

    Dim lTemp As Long
    Dim i As Long
    
    lTemp = 0
    
    For i = 0 To UBound(SyncItems())
        If LCase(SyncItems(i).File) = LCase(sFile) Then
            lTemp = 1
            Exit For
        End If
    Next
    
    SyncStatus = lTemp

End Function

Public Function SyncClear() As Boolean

    ReDim SyncItems(0)
    
    SyncClear = True

End Function
