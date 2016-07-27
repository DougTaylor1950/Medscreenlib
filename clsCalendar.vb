Option Strict On
Imports System.Data.OleDb




''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Calendar
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''   Class based around Calendar table for identifying public holidays
''' </summary>
''' <remarks>
''' Adding a comment to change the file
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class Calendar
    Private Shared _dayStart As TimeSpan, _dayEnd As TimeSpan

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if the passed date is a weekday and not a holiday
  ''' </summary>
  ''' <param name="checkDate"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsWorkingDay(ByVal checkDate As Date) As Boolean
    Return IsWeekDay(checkDate) AndAlso _
           Not IsAHoliday(checkDate)
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if the passed date is a weekday, i.e. Mon - Fri
  ''' </summary>
  ''' <param name="checkDate"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsWeekDay(ByVal checkDate As Date) As Boolean
    Return checkDate.DayOfWeek <> DayOfWeek.Saturday AndAlso _
           checkDate.DayOfWeek <> DayOfWeek.Sunday
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if the passed date is marked as being a holiday in the passed location
  ''' </summary>
  ''' <param name="checkDate"></param>
  ''' <param name="location"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   Location is the ISO country ID, sometimes followed by a sub region, e.g. UK-NI
  '''   for Northern Ireland.  These are defined in the VGL library lib_Constants.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsAHoliday(ByVal checkDate As Date, ByVal location As String) As Boolean
    Dim cmdFindHoliday As New OleDbCommand("SELECT Day FROM Calendar " & _
            "WHERE trunc(Day) = trunc(?) AND EntryType = ? AND Location = ?  ", _
                                                 MedConnection.Connection)
    cmdFindHoliday.Parameters.Add("CheckDate", OleDbType.Date).Value = checkDate
    cmdFindHoliday.Parameters.Add("EventType", OleDbType.VarChar, 10).Value = "HOLIDAY"
    cmdFindHoliday.Parameters.Add("Location", OleDbType.VarChar, 10).Value = location
    Dim result As Object = Nothing
    Try
      If MedConnection.Open Then
        result = cmdFindHoliday.ExecuteScalar
      End If
    Catch ex As System.Exception
      Console.WriteLine(ex.Message)
    Finally
      MedConnection.Close()
    End Try
    Return Not result Is Nothing
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if the passed date is marked as being a holiday in any of the passed location
  ''' </summary>
  ''' <param name="checkDate"></param>
  ''' <param name="location"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   Location is the ISO country ID, sometimes followed by a sub region, e.g. UK-NI
  '''   for Northern Ireland.  These are defined in the VGL library lib_Constants.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsAHoliday(ByVal checkDate As Date, ByVal location() As String) As Boolean
    Dim locations As String = "" 'String.Join(", ", location)
    checkDate = Medscreen.TruncateDate(checkDate)
    Dim cmdFindHoliday As New OleDbCommand("SELECT Day FROM Calendar WHERE trunc(Day) = ? " & _
                                                 "AND EntryType = ? AND Location IN (", _
                                                 MedConnection.Connection)
    cmdFindHoliday.Parameters.Add("CheckDate", OleDbType.Date).Value = checkDate
    cmdFindHoliday.Parameters.Add("EventType", OleDbType.VarChar, 10).Value = "HOLIDAY"
        Dim k As Integer
        For k = 0 To location.Length - 1
            locations += ",?"
            cmdFindHoliday.Parameters.Add("LK" & k.ToString.Trim, OleDbType.VarChar).Value = location(k)
        Next
        If locations.Length > 0 Then
            locations = locations.Substring(1)
        End If
        cmdFindHoliday.CommandText &= locations & ")"
        'Console.WriteLine(cmdFindHoliday.CommandText)
        Dim result As Object = Nothing
        Try
            If MedConnection.Open Then
                result = cmdFindHoliday.ExecuteScalar
                Console.WriteLine(result)
            End If
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        Finally
            MedConnection.Close()
        End Try
        Return Not result Is Nothing
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if the passed date is marked as being a holiday for the UK
  ''' </summary>
  ''' <param name="checkDate"></param>
    ''' <returns></returns>
  ''' <remarks>
  '''   Location is the ISO country ID, sometimes followed by a sub region, e.g. UK-NI
  '''   for Northern Ireland.  These are defined in the VGL library lib_Constants.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsAHoliday(ByVal checkDate As Date) As Boolean
    Return IsAHoliday(checkDate, New String() {"UK", "UK-EN"})
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Adds the specified number of working days (+ve or -ve) to the passed date
  ''' </summary>
  ''' <param name="fromDate"></param>
  ''' <param name="days"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   If days is 0, the next working day is returned (or fromDate if it is a working day)
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function AddWorkingDays(ByVal fromDate As Date, ByVal days As Integer) As Date
    Dim oneDay As Integer = 1
    ' If we're taking days off, set oneDay to be negative
    If days < 0 Then oneDay = -1
    If days = 0 Then
      ' No days need adding, return the next working day
      fromDate = GetWorkingDay(fromDate, True)
    Else
      ' While there are still days to add, add (or subtract) one working day
      While days <> 0
        fromDate = AddWorkingDay(fromDate, oneDay)
        days -= oneDay
      End While
    End If
    Return fromDate
    End Function

    ''' <developer>CONCATENO\taylor</developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FromDate"></param>
    ''' <param name="toDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>20-Jun-2014 12:07</created><Author>CONCATENO\taylor</Author></revision></revisionHistory>
    Public Shared Function WorkingDaysBetween(ByVal FromDate As DateTime, ByVal toDate As DateTime) As Double
        Dim dblRet As Double = 0
        Dim NextDay As DateTime
        Dim LastDay As DateTime
        If toDate > FromDate Then
            Dim ts As TimeSpan = toDate.Subtract(FromDate)
            dblRet = ts.TotalDays
            NextDay = DateSerial(FromDate.Year, FromDate.Month, FromDate.Day).AddDays(1)
            LastDay = DateSerial(toDate.Year, toDate.Month, toDate.Day)
            If Not IsWorkingDay(toDate) Then        'If the put on hold date is a non working day then take away hours
                ts = NextDay.Subtract(FromDate)
                dblRet -= ts.TotalDays
            End If

            'step through each full day between taking 1 off if not a working day.
            While NextDay < LastDay
                If Not IsWorkingDay(NextDay) Then
                    dblRet -= 1
                End If
                NextDay = NextDay.AddDays(1)
            End While

            'finish up last day 
            If Not IsWorkingDay(toDate) Then
                ts = toDate.Subtract(LastDay)
                dblRet -= ts.TotalDays
            End If

        End If
        Return dblRet
    End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the nearest working date to fromDate, or fromDate itself if it's a working day
  ''' </summary>
  ''' <param name="fromDate"></param>
  ''' <param name="forward">True to search forward in time, False to search back</param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetWorkingDay(ByVal fromDate As Date, ByVal forward As Boolean) As Date
    If Not IsWorkingDay(fromDate) Then
      Dim oneDay As Integer = 1
      If Not forward Then oneDay = -1
      fromDate = AddWorkingDay(fromDate, oneDay)
    End If
    Return fromDate
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Adds or subtracts one working day to / from fromDate
  ''' </summary>
  ''' <param name="fromDate">Date to add to</param>
  ''' <param name="oneDay">-ve Integer to subtract, any other value to add</param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Shared Function AddWorkingDay(ByVal fromDate As Date, ByVal oneDay As Integer) As Date
    If oneDay < 0 Then oneDay = -1 Else oneDay = 1
    ' Add or subtract one day (oneDay should be passed as either 1 or -1
    fromDate = fromDate.AddDays(oneDay)
    ' If result is not a working day, repeat until it is
    While Not Calendar.IsWorkingDay(fromDate)
      fromDate = fromDate.AddDays(oneDay)
    End While
    ' Return the working day
    Return fromDate
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the date at one second before the end of the month.
  ''' </summary>
  ''' <param name="fromDate"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function EndOfMonth(ByVal fromDate As Date) As Date
    ' Return date one second back from start of next month to get end of this month
    Return StartOfMonth(fromDate.AddMonths(1)).AddSeconds(-1)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the data at midnight on the first of the month
  ''' </summary>
  ''' <param name="fromDate"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function StartOfMonth(ByVal fromDate As Date) As Date
    Return fromDate.Date.Subtract(TimeSpan.FromDays(fromDate.Day - 1))
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Indicates the amount of time between two dates that is in hours
  ''' </summary>
  ''' <param name="startTime"></param>
  ''' <param name="endTime"></param>
  ''' <param name="dayStart"></param>
  ''' <param name="dayEnd"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [08/01/2007]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function TimeInHours(ByVal startTime As Date, ByVal endTime As Date, ByVal dayStart As TimeSpan, ByVal dayEnd As TimeSpan) As TimeSpan
    Dim minutes As Double, nextDate As Date
    Do While startTime < endTime
      If IsInHours(startTime, dayStart, dayEnd) Then
        ' In hours will end either at day end, or and end time
        nextDate = MinDate(startTime.Date.Add(dayEnd), endTime)
        ' Add amount of time in hours
        minutes += nextDate.Subtract(startTime).TotalMinutes
      ElseIf startTime.TimeOfDay.TotalMinutes < dayStart.TotalMinutes Then
        ' Before day start, next working period may start at day start
        nextDate = startTime.Date.Add(dayStart)
      Else
        ' After day start, next working period may start at start of next day
        nextDate = startTime.Date.AddDays(1).Add(dayStart)
      End If
      startTime = nextDate
    Loop
    Return TimeSpan.FromMinutes(minutes)
  End Function

    Public Shared Function TimeInHours(ByVal startTime As Date, ByVal endTime As Date) As TimeSpan
        Return TimeInHours(startTime, endTime, Calendar.WorkDayStart, Calendar.WorkDayEnd)
    End Function

    Public Shared ReadOnly Property WorkDayStart() As TimeSpan
        Get
            If _dayStart.Equals(TimeSpan.Zero) Then
                Try
                    _dayStart = Glossary.Interval.Parse(New Glossary.ConfigItem("WORKHOURS_START").DefaultValue).ToTimeSpan()
                Catch ex As Exception
                    _dayStart = TimeSpan.FromHours(8.5)
                    Medscreen.LogError(ex, False, "Error fetching WORKHOURS_START config item.")
                End Try
            End If
        End Get
    End Property

    Public Shared ReadOnly Property WorkDayEnd() As TimeSpan
        Get
            If _dayEnd.Equals(TimeSpan.Zero) Then
                Try
                    _dayEnd = Glossary.Interval.Parse(New Glossary.ConfigItem("WORKHOURS_END").DefaultValue).ToTimeSpan()
                Catch ex As Exception
                    _dayEnd = TimeSpan.FromHours(8.5)
                    Medscreen.LogError(ex, False, "Error fetching WORKHOURS_END config item.")
                End Try
            End If
        End Get
    End Property



  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''  Indicates whether the specified date is in hours
  ''' </summary>
  ''' <param name="checkDate"></param>
  ''' <param name="dayStart">Start of day (usually 8am)</param>
  ''' <param name="dayEnd">End of day (usually 6pm)</param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [08/01/2007]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function IsInHours(ByVal checkDate As Date, ByVal dayStart As TimeSpan, ByVal dayEnd As TimeSpan) As Boolean
    Return checkDate.TimeOfDay.TotalMinutes >= dayStart.TotalMinutes AndAlso _
           checkDate.TimeOfDay.TotalMinutes < dayEnd.TotalMinutes AndAlso _
           Calendar.IsWorkingDay(checkDate)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the least of two dates
  ''' </summary>
  ''' <param name="date1"></param>
  ''' <param name="date2"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [08/01/2007]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function MinDate(ByVal date1 As Date, ByVal date2 As Date) As Date
    If date1.CompareTo(date2) > 0 Then
      Return date2
    Else
      Return date1
    End If
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the greater of two dates
  ''' </summary>
  ''' <param name="date1"></param>
  ''' <param name="date2"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [08/01/2007]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function MaxDate(ByVal date1 As Date, ByVal date2 As Date) As Date
    If date1.CompareTo(date2) > 0 Then
      Return date1
    Else
      Return date2
    End If
  End Function

End Class
