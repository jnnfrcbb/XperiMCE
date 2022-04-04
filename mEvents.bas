Attribute VB_Name = "mEvents"

Private Type tEvent
    Name As String
    Location As String
    Start As String
    End As String
    CalendarID As String
    Color As String
    Notified As Boolean
    Thumb As String
End Type
Public Events() As tEvent

Public Function GetEventTime(sTime As String) As String

    GetEventTime = Mid(sTime, 9, 2) & " " & MonthName(Mid(sTime, 6, 2)) & ", " & Mid(sTime, 12, 5)

End Function

Public Function QueryNotifyEvent(Index As Long) As Boolean

    Dim bTemp As Boolean
    Dim strHour() As String
    
    bTemp = False
    
    strHour() = Split(Events(Index).Start, ":")
    
    If Hour(Now) >= CLng(strHour(0)) - 1 And Hour(Now) <= CLng(strHour(0)) And Events(Index).Notified = False Then
    
        bTemp = True
    
    End If
    
    QueryNotifyEvent = bTemp

End Function
