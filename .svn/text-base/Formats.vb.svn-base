Public Class Formats

    Public Shared Function FormatPostedTimeWithDayNumberSuffix(ByVal Entry As Date, Optional ByVal format As String = "d{0} o\f MMMM yyyy") As String
        Dim dateTime As DateTime = Date.Parse(Entry)
        'If Not DBNull.IsNull(dateTime) Then
        Dim formatTime As String = dateTime.ToString(format)
        Dim day As Integer = dateTime.Day
        Dim dayFormat As String = String.Empty
        Select Case day
            Case 1, 21, 31
                dayFormat = "st"
            Case 2, 22
                dayFormat = "nd"
            Case 3, 23
                dayFormat = "rd"
            Case Else
                dayFormat = "th"
        End Select
        Return String.Format(formatTime, dayFormat)
        'End If
    End Function '- See more at: http://www.sunblognuke.com/blog/format-date-time-with-suffix-like-1st-2nd-3rd-4th.aspx#sthash.Q7giobGO.dpu

End Class


