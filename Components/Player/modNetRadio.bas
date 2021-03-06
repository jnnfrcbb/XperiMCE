Attribute VB_Name = "modNetRadio"
'/////////////////////////////////////////////////////////////////////////////////
' modNetRadio.bas - Copyright (c) 2002-2007 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                         [http://www.jobnik.org]
'                                                         [  jobnik@jobnik.org  ]
'
' * Save local copy is added by: Peter Hebels @ http://www.phsoft.nl
'                                             e-mail: info@phsoft.nl
'
' Other sources: frmNetRadio.frm & clsFileIo.cls
'
' BASS Internet radio example
' Originally translated from - netradio.c - Example of Ian Luck
'/////////////////////////////////////////////////////////////////////////////////

Option Explicit

Public chan As Long
Public TmpNameHold As String
Public TmpNameHold2 As String

Public proxy(100) As Byte ' proxy server

' SAVE LOCAL COPY
Public WriteFile As clsFileIo
Public FileIsOpen As Boolean, GotHeader As Boolean
Public DownloadStarted As Boolean, DoDownload As Boolean
Public DlOutput As String, SongNameUpdate As Boolean

' THREADING
Public cthread As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' MESSAGE BOX
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'FROM USERCONTROL

Public StreamHandle() As Long

Public uC As PlayerControl

Public Function SetControl(nuC As PlayerControl)

Set uC = nuC

End Function

' display error message
Public Sub Error_(ByVal es As String)

    uC.sError = es

End Sub

' update stream title from metadata
Sub DoMeta()
    Dim meta As Long
    Dim p As String, tmpMeta As String
    Dim sTemp As String
    meta = BASS_ChannelGetTags(StreamHandle(0), BASS_TAG_META)
    If meta = 0 Then Exit Sub
    tmpMeta = VBStrFromAnsiPtr(meta)
    If ((Mid(tmpMeta, 1, 13) = "StreamTitle='")) Then
        p = Mid(tmpMeta, 14)
        TmpNameHold = Mid(p, 1, InStr(p, ";") - 2)
        uC.StreamSong = TmpNameHold
        
        If TmpNameHold = TmpNameHold2 Then
            ' do noting
        Else
            TmpNameHold2 = TmpNameHold
            GotHeader = False
            DownloadStarted = False
        End If
        
        sTemp = Replace(uC.StreamName, " ", "_") & "_" & Replace(DateValue(Now), "/", "-") & "_" & Replace(Mid(TimeValue(Now), 1, 5), ":", "")

        If RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)) <> "" Then
            sTemp = sTemp & "_" & Replace(RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)), " ", "_")
        End If
        
        Debug.Print sTemp
        DlOutput = App.Path & "\" & sTemp & ".mp3"
        Debug.Print DlOutput
    End If
End Sub

Sub MetaSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Call DoMeta
End Sub

Sub EndSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)

        uC.StreamName = "not playing"
        uC.StreamBPS = ""
        uC.StreamSong = ""

End Sub

'Public Sub OpenURL(ByVal clkURL As Long)
Public Sub OpenURL()

        uC.SetTimer False
        'Call BASS_StreamFree(uc.streamhandle(0)) ' close old stream
        uC.StreamName = "connecting..."
        uC.StreamBPS = ""
        uC.StreamSong = ""
        
        StreamHandle(0) = BASS_StreamCreateURL(uC.sURL, 0, BASS_STREAM_BLOCK Or BASS_STREAM_STATUS Or BASS_STREAM_AUTOFREE, AddressOf SUBDOWNLOADPROC, 0)

        If StreamHandle(0) = 0 Then
            uC.StreamName = "not playing"
            Call Error_("Can't play the stream")
        Else
            uC.SetTimer (True)
        End If
done:
    Call CloseHandle(cthread)   ' close the thread
    cthread = 0
End Sub

' The following functions where added by Peter Hebels
Public Sub SUBDOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    If (buffer And length = 0) Then
        uC.StreamBPS = VBStrFromAnsiPtr(buffer) ' display connection status
        Exit Sub
    End If

    If (Not DoDownload) Then
        DownloadStarted = False
        Call WriteFile.CloseFile
        Exit Sub
    End If

    If (Trim(DlOutput) = "") Then Exit Sub

    If (Not DownloadStarted) Then
        DownloadStarted = True
        Call WriteFile.CloseFile
        DoMeta
        If (WriteFile.OpenFile(DlOutput)) Then
            SongNameUpdate = False
        Else
            
            SongNameUpdate = True
            
            GotHeader = False
        End If
    End If

    If (Not SongNameUpdate) Then
        If (length) Then
            Call WriteFile.WriteBytes(buffer, length)
        Else
            Call WriteFile.CloseFile
            GotHeader = False
        End If
    Else
        DownloadStarted = False
        Call WriteFile.CloseFile
        GotHeader = False
    End If
End Sub

Public Function RemoveSpecialChar(strFileName As String)
    Dim i As Byte
    Dim SpecialChar As Boolean
    Dim SelChar As String, OutFileName As String

    For i = 1 To Len(strFileName)
        SelChar = Mid(strFileName, i, 1)
        SpecialChar = InStr(":/\?*|<>" & Chr$(34), SelChar) > 0

        If (Not SpecialChar) Then
            OutFileName = OutFileName & SelChar
            SpecialChar = False
        Else
            OutFileName = OutFileName
            SpecialChar = False
        End If
    Next i

    RemoveSpecialChar = OutFileName
End Function
