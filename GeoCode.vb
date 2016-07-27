Imports BenjaminSchroeter
Public Class GeoCode

    Private Shared Loc As BenjaminSchroeter.GeoNames.Location = Nothing

    Public Shared Sub ClearData()
        Loc = Nothing
    End Sub

    Public Shared Function GetLatitude(ByVal PostCode As String) As Double
        If Loc Is Nothing Then
            Loc = New BenjaminSchroeter.GeoNames.Location(PostCode)
        End If
        If Loc.Lat.Trim.Length > 0 Then
            Return Loc.Lat
        Else
            Return 200
        End If
    End Function

    <CLSCompliant(False)> _
    Public Shared Function Search(ByVal PlaceName As String, ByVal Country As String) As GeoNames.Geoname
        Dim str As GeoNames.Geoname = BenjaminSchroeter.GeoNames.GeoNamesOrgWebservice.FindPlaceByName(PlaceName, Country)

        Return str
    End Function

    Public Shared Function GetLongitude(ByVal PostCode As String) As Double
        If Loc Is Nothing Then
            Loc = New BenjaminSchroeter.GeoNames.Location(PostCode)
        End If
        If Loc.Lng.Trim.Length > 0 Then
            Return Loc.Lng
        Else
            Return 200
        End If
    End Function

    Public Overloads Shared Function GetLatLong(ByVal PostCode As String, ByVal CountryCode As Integer) As String
        Dim strRet As String = ""
        Dim cCode As String = CConnection.PackageStringList("Lib_country.GetCountryCodeFromID", CountryCode)
        If cCode.Trim.Length > 0 Then
            strRet = GetLatLong(PostCode, cCode)
        End If
        Return strRet
    End Function

    Public Overloads Shared Function GetLatLong(ByVal PostCode As String, ByVal CountryCode As String) As String
        If Loc Is Nothing Then
            Loc = New BenjaminSchroeter.GeoNames.Location(PostCode, CountryCode)
        End If
        If Loc.PostCode <> PostCode Then
            Loc = New BenjaminSchroeter.GeoNames.Location(PostCode, CountryCode)
        End If
        Dim retString As String = ""
        Dim fString As String = "{0:n4},{1:n4}"
        If Loc.Lat IsNot Nothing AndAlso Loc.Lat.Trim.Length > 0 AndAlso Loc.Lng IsNot Nothing AndAlso Loc.Lng.Trim.Length > 0 Then
            Dim lat As Double = CDbl(Loc.Lat)
            Dim lng As Double = CDbl(Loc.Lng)
            retString = lat.ToString("0.0000") & "," & lng.ToString("0.0000")
        End If
        Return retString
    End Function

    Public Shared Function ConvertCoordinates(ByVal inCoords As String) As String
        Dim strRet As String = ""
        If inCoords.Trim.Length = 0 Then
            Return strRet
            Exit Function
        End If


        Dim strTemp As String = inCoords.ToUpper.Trim
        If InStr(strTemp, "N") + InStr(strTemp, "S") + InStr(strTemp, "W") + InStr(strTemp, "E") = 0 AndAlso strTemp.Length > 5 Then  'Check for pure numeric
            Dim strElArray As String() = strTemp.Split(New Char() {" ", ","}) 'split into north/south East/west
            Dim ns As Double = CDbl(strElArray(0))
            Dim ew As Double = CDbl(strElArray(1))
            Dim degrees As Integer = Int(Math.Abs(ns))
            Dim minutes As Integer = (Math.Abs(ns) - degrees) * 60
            strRet += degrees.ToString("00") & minutes.ToString("00")
            If ns > 0 Then
                strRet += "N "
            Else
                strRet += "S "
            End If

            degrees = Int(Math.Abs(ew))
            minutes = (Math.Abs(ew) - degrees) * 60
            strRet += degrees.ToString("00") & minutes.ToString("00")
            If ew > 0 Then
                strRet += "E"
            Else
                strRet += "W"
            End If

        Else
            'Split into two bits
            If InStr(strTemp, ":") + InStr(strTemp, "°") + InStr(strTemp, ".") = 0 Then
                If strTemp.Length > 5 Then
                    strRet = strTemp.ToUpper.Trim
                Else 'postcode 
                    Dim strLat As String = CConnection.PackageStringList("Lib_officer.GetLatitude", strTemp.Trim.ToUpper)
                    Dim strLong As String = CConnection.PackageStringList("Lib_officer.GetLongitude", strTemp.Trim.ToUpper)
                    Dim ns As Double = CDbl(strLat)
                    Dim ew As Double = CDbl(strLong)
                    Dim degrees As Integer = Int(Math.Abs(ns))
                    Dim minutes As Integer = (Math.Abs(ns) - degrees) * 60
                    strRet += degrees.ToString("00") & minutes.ToString("00")
                    If ns > 0 Then
                        strRet += "N "
                    Else
                        strRet += "S "
                    End If

                    degrees = Int(Math.Abs(ew))
                    minutes = (Math.Abs(ew) - degrees) * 60
                    strRet += degrees.ToString("00") & minutes.ToString("00")
                    If ew > 0 Then
                        strRet += "E"
                    Else
                        strRet += "W"
                    End If
                End If
            Else 'other formats
                strTemp = Medscreen.ReplaceString(strTemp, "° ", "°")
                strTemp = Medscreen.ReplaceString(strTemp, "' ", "'")
                strTemp = Medscreen.ReplaceString(strTemp, ", ", ",")
                Dim elArray As String() = strTemp.Split(New Char() {",", " "})
                'we should now have two elements
                If elArray.Length = 2 Then
                    Dim startChar As Char
                    Dim intPos As Integer
                    Dim Element As String = elArray(0)
                    intPos = InStr(Element, "N") + InStr(Element, "S")
                    startChar = Mid(Element, intPos, 1)
                    If intPos = 1 Then
                        Element = Mid(Element, 2)
                    Else
                        Element = Mid(Element, 1, intPos - 1)
                    End If
                    'now split into elements
                    Dim degMinsSec As String() = Element.Split(New Char() {":", """", "'", "°", "."})
                    Dim degrees As Integer = degMinsSec(0)
                    Dim minutes As Integer = degMinsSec(1)
                    If InStr(Element, ".") = 0 Then
                        strRet = degrees.ToString("00").Trim & minutes.ToString("00").Trim & startChar & " "
                    Else
                        strRet = degrees.ToString("00").Trim & CDbl(CDbl("." & minutes.ToString("00")) * 60).ToString("00").Trim & startChar & " "
                    End If

                    Element = elArray(1)
                    intPos = InStr(Element, "E") + InStr(Element, "W")
                    startChar = Mid(Element, intPos, 1)
                    If intPos = 1 Then
                        Element = Mid(Element, 2)
                    Else
                        Element = Mid(Element, 1, intPos - 1)
                    End If
                    'now split into elements
                    degMinsSec = Element.Split(New Char() {":", """", "'", "°", "."})
                    degrees = degMinsSec(0)
                    minutes = degMinsSec(1)
                    If InStr(Element, ".") = 0 Then
                        strRet += degrees.ToString("000").Trim & minutes.ToString("00").Trim & startChar & " "
                    Else
                        strRet += degrees.ToString("000").Trim & CDbl(CDbl("." & minutes.ToString("00")) * 60).ToString("00").Trim & startChar & " "
                    End If
                End If
            End If

        End If

        Return strRet
    End Function

    Public Shared Function MapXML(ByVal location As String, Optional ByVal scale As Integer = 7) As String
        Dim strret As String = ""
        strret = "<PortLoc>" & vbCrLf & "<GEOLocation>" & location & "</GEOLocation>" & vbCrLf & _
                     "<MapScale>" & scale & "</MapScale>" & vbCrLf & _
                 "</PortLoc>"
        Return strret
    End Function


    Public Shared Function ConvertStringCoordinatesToNumeric(ByVal inCoords As String) As String
        Dim strRet As String = ""
        Dim Coords As String() = inCoords.Trim.Split(New Char() {" "})
        Dim sDegrees As String = Mid(Coords(0), 1, 2)
        Dim sminutes As String = Mid(Coords(0), 3, 2)

        Dim DirChar As String = ""
        If Coords(0).Length = 5 Then
            DirChar = Mid(Coords(0), 5, 1)
        End If

        Dim Lat As Double = CDbl(sDegrees) + sminutes / 60
        If DirChar = "S" Then Lat = Lat * -1

        If Coords(1).Length = 5 Then Coords(1) = "0" & Coords(1)

        sDegrees = Mid(Coords(1), 1, 3)
        sminutes = Mid(Coords(1), 4, 2)

        DirChar = Mid(Coords(1), 6, 1)
        Dim Lng As Double = CDbl(sDegrees) + sminutes / 60
        If DirChar = "W" Then Lng = Lng * -1

        strRet = Lat.ToString("00.0000") & "," & Lng.ToString("000.0000")

        Return strRet
    End Function

    Public Overloads Shared Function GetLatLong(ByVal PostCode As String) As String
        Return GetLatLong(PostCode, "GB")
    End Function

End Class

Public Class MapScale

    Private myMaxLat As Double = -181
    Private Property MaxLat() As Double
        Get
            Return myMaxLat
        End Get
        Set(ByVal value As Double)
            myMaxLat = value
        End Set
    End Property


    Private myPath As String = "path=color:0x0000ff|weight:5"
    Public Property Path() As String
        Get
            Return myPath
        End Get
        Set(ByVal value As String)
            myPath = value
        End Set
    End Property


    Private myMinLat As Double = 181
    Private Property MinLat() As Double
        Get
            Return myMinLat
        End Get
        Set(ByVal value As Double)
            myMinLat = value
        End Set
    End Property

    Private myMaxLong As Double = -181
    Private Property MaxLong() As Double
        Get
            Return myMaxLong
        End Get
        Set(ByVal value As Double)
            myMaxLong = value
        End Set
    End Property

    Private myMinLong As Double = 181
    Private Property MinLong() As Double
        Get
            Return myMinLong
        End Get
        Set(ByVal value As Double)
            myMinLong = value
        End Set
    End Property

    Private myLAstLocString As String = ""
    Public Property LastLocString() As String
        Get
            Return myLAstLocString
        End Get
        Set(ByVal value As String)
            myLAstLocString = value
        End Set
    End Property


    Public Sub New()
        MyBase.New()
    End Sub

    Public Overloads Sub AddLoc(ByVal Lat As Double, ByVal Lng As Double)
        Dim LocString As String = "|" & Lat.ToString("0.00") & "," & Lng.ToString("0.00")
        If LastLocString <> LocString Then
            myPath += LocString
            LastLocString = LocString
        End If

        If Lat > MaxLat Then MaxLat = Lat
        If Lat < MinLat Then MinLat = Lat
        If Lng > MaxLong Then MaxLong = Lng
        If Lng < MinLong Then MinLong = Lng
    End Sub

    Public Overloads Sub AddLoc(ByVal Latlng As String)
        If Latlng.Trim.Length = 0 Then Exit Sub
        Dim aLatLong As String() = Latlng.Split(New Char() {","})
        Dim Lat As Double = aLatLong(0)
        Dim Lng As Double = aLatLong(1)

        AddLoc(Lat, Lng)
    End Sub

    Function MapScale() As Integer
        Dim LatSpan As Double = MaxLat() - MinLat()
        Dim longspan As Double = MaxLong() - MinLong()
        If LatSpan < longspan Then
            LatSpan = longspan

        End If
        'Work down 
        If LatSpan = 0 And longspan = 0 Then
            Return 5
        ElseIf LatSpan = 0 And MinLat() > 180 Then
            Return 1
        ElseIf LatSpan > 150 Then
            Return 1
        ElseIf LatSpan > 75 Then
            Return 2

        ElseIf LatSpan > 50 Then
            Return 3

        ElseIf LatSpan > 20 Then
            Return 4

        ElseIf LatSpan > 10 Then
            Return 5
        ElseIf LatSpan > 8 Then
            Return 6
        ElseIf LatSpan > 0.85 Then
            Return 7
        ElseIf LatSpan > 0.7 Then
            Return 8
        ElseIf LatSpan > 0.6 Then
            Return 9
        Else
            Return 10
        End If
    End Function

    Public Function Centre() As String
        Dim LatCentre As Double = MinLat + (MaxLat - MinLat) / 2
        Dim LngCentre As Double = MinLong + (MaxLong - MinLong) / 2
        Return LatCentre.ToString("0.0000") & "," & LngCentre.ToString("0.0000")
    End Function

End Class
