Public Class Utils
#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"
    Public Shared Function ConvertStringToInterval(ByVal inInterval As String) As Double
        Dim NewInterval As Double = 0
        If inInterval.Trim.Length > 0 Then
            Dim IntervalArray As String() = inInterval.Trim.Split(New Char() {" ", ":", "."})
            NewInterval = CInt(IntervalArray(0))
            Dim Hours As Integer = CInt(IntervalArray(1))
            NewInterval += Hours / 24
            Dim Mins As Integer = CInt(IntervalArray(2))
            NewInterval += Mins / (24 * 60)
            Dim Secs As Integer = CInt(IntervalArray(3))
            NewInterval += Secs / (24 * 60 * 60)
        End If
        Return NewInterval
    End Function

    Public Shared Function ConvertIntervalToString(ByVal inInterval As Double) As String
        Dim Days As Integer = Math.Truncate(inInterval)
        Dim dHours As Double = (inInterval - Days) * 24
        Dim Hours As Integer = Math.Truncate(dHours)
        Dim Mins As Integer = (dHours - Hours) * 60
        Dim strRet As String = ""
        If Days + Hours + Mins <> 0 Then
            strRet = Days.ToString.PadLeft(4, " ") & " " & Hours.ToString("00") & ":" & Mins.ToString("00") & ":00.00"
        End If
        Return strRet
    End Function

    Public Shared Function ConvertLatLongToCoordinates(ByVal inLatLong As String) As String
        Dim latLong As String() = inLatLong.Split(New Char() {","})
        Dim lat As Decimal = CDbl(latLong(0))
        Dim lng As Decimal = CDbl(latLong(1))
        Dim latC As Char = "N"
        If lat < 0 Then latC = "S"
        Dim lngC As Char = "E"
        If lng < 0 Then lngC = "W"
        lat = Math.Abs(lat)
        lng = Math.Abs(lng)
        Dim intpart As Integer = Math.Truncate(lat)
        Dim decpart As Integer = CInt((lat - intpart) * 60)
        Dim retStr As String = intpart.ToString("00") & decpart.ToString("00") & latC

        intpart = Math.Truncate(lng)
        decpart = CInt((lng - intpart) * 60)
        retStr += " " & intpart.ToString("000") & decpart.ToString("00") & lngC
        Return retStr
    End Function

    Public Shared Function ConvertCoordinatesToLatLong(ByVal Coordinates As String) As String
        Dim retString As String = ""
        If Coordinates.Trim.Length > 0 Then
            Dim coordArray As String() = Coordinates.Trim.Split(New Char() {" "})
            If coordArray.Length = 2 Then
                Dim istCoord As String = coordArray(0)
                Dim LastChar As Char = istCoord.Chars(istCoord.Length - 1)
                Dim CoordText As String = istCoord.Substring(0, istCoord.Length - 1)
                Dim inttext As String = ""
                Dim dectext As String = ""
                If CoordText.Length = 4 Then
                    inttext = CoordText.Substring(0, 2)
                    dectext = CoordText.Substring(2, 2)
                ElseIf CoordText.Length = 5 Then
                    inttext = CoordText.Substring(0, 3)
                    dectext = CoordText.Substring(3, 2)
                End If
                Dim Latitude As Double = CDbl(inttext) + CDbl(dectext) / 60
                If LastChar = "S" Then Latitude = Latitude * -1

                Dim scndCoord As String = coordArray(1)
                LastChar = scndCoord.Chars(scndCoord.Length - 1)
                CoordText = scndCoord.Substring(0, scndCoord.Length - 1)

                inttext = ""
                dectext = ""
                If CoordText.Length = 4 Then
                    inttext = CoordText.Substring(0, 2)
                    dectext = CoordText.Substring(2, 2)
                ElseIf CoordText.Length = 5 Then
                    inttext = CoordText.Substring(0, 3)
                    dectext = CoordText.Substring(3, 2)
                End If
                Dim Longitude As Double = CDbl(inttext) + CDbl(dectext) / 60
                If LastChar = "W" Then Longitude = Longitude * -1

                retString = Latitude.ToString("0.0000") & "," & Longitude.ToString("0.0000")

            End If
        End If
        Return retString
    End Function

    
#End Region


#Region "Procedures"

#End Region

#Region "Properties"

#End Region
#End Region

End Class