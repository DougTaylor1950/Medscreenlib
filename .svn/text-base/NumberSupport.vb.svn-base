Public Class NumberSupport
    Private Shared sNumberText As String() = {"Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", _
                                              "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", _
                                              "Twenty", "Thirty", "Fourty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"}
    Public Shared Function HundredsTensUnits(ByVal TestValue As Integer, _
                                   Optional ByVal bUseAnd As Boolean = True) As String

        Dim CardinalNumber As Integer
        Dim sHundredsTensUnits As String = ""

        If TestValue > 99 Then
            CardinalNumber = TestValue \ 100
            sHundredsTensUnits = sNumberText(CardinalNumber) & " Hundred "
            TestValue = TestValue - (CardinalNumber * 100)
        End If

        If bUseAnd = True Then
            sHundredsTensUnits = sHundredsTensUnits & "and "
        End If

        If TestValue > 20 Then
            CardinalNumber = TestValue \ 10
            sHundredsTensUnits = sHundredsTensUnits & _
                                sNumberText(CardinalNumber + 18) & " "
            TestValue = TestValue - (CardinalNumber * 10)
        End If

        If TestValue > 0 Then
            sHundredsTensUnits = sHundredsTensUnits & sNumberText(TestValue) & " "
        End If
        Return sHundredsTensUnits
    End Function


    Public Shared Function NumberAsText(ByVal NumberIn As String, _
                              Optional ByVal AND_or_CHECK_or_DOLLAR As String = "") As String
        Dim cnt As Long
        Dim DecimalPoint As Long
        Dim CardinalNumber As Long
        Dim CommaAdjuster As Long
        Dim TestValue As Long
        Dim CentsString As String = ""
        Dim NumberSign As String = ""
        Dim WholePart As String = ""
        Dim BigWholePart As String = ""
        Dim DecimalPart As String = ""
        Dim tmp As String = ""
        Dim sStyle As String = ""
        Dim bUseAnd As Boolean
        Dim bUseCheck As Boolean
        Dim bUseDollars As Boolean


        '----------------------------------------
        'Begin setting conditions for formatting
        '----------------------------------------

        'Determine whether to apply special formatting.
        'If nothing passed, return routine result
        'converted only into its numeric equivalents,
        'with no additional format text.
        sStyle = LCase(AND_or_CHECK_or_DOLLAR)

        'User passed "AND": "and" will be added 
        'between hundredths and tens of dollars,
        'ie "Three Hundred and Forty Two"
        bUseAnd = sStyle = "and"

        'User passed "DOLLAR": "dollar(s)" and "cents"
        'appended to string,
        'ie "Three Hundred and Forty Two Dollars"
        bUseDollars = sStyle = "dollar"

        'User passed "CHECK" *or* "DOLLAR"
        'If "check", cent amount returned as a fraction /100
        'i.e. "Three Hundred Forty Two and 00/100"
        'If "dollar" was passed, "dollar(s)" and "cents"
        'appended instead.
        bUseCheck = (sStyle = "check") Or (sStyle = "dollar")



        '----------------------------------------
        'Begin validating the number, and breaking
        'into constituent parts
        '----------------------------------------

        'prepare to check for valid value in
        NumberIn = CStr(NumberIn).Trim

        If Not IsNumeric(NumberIn) Then

            'invalid entry - abort
            Return "Error - Number improperly formed"
            Exit Function

        Else

            'decimal check
            DecimalPoint = InStr(NumberIn.ToString, ".")

            If DecimalPoint > 0 Then

                'split the fractional and primary numbers
                DecimalPart = Mid(NumberIn, DecimalPoint + 1)
                WholePart = Left(NumberIn, DecimalPoint - 1)

            Else

                'assume the decimal is the last char
                DecimalPoint = Len(NumberIn) + 1
                WholePart = NumberIn

            End If

            If InStr(NumberIn, ",,") Or _
               InStr(NumberIn, ",.") Or _
               InStr(NumberIn, ".,") Or _
               InStr(DecimalPart, ",") Then

                Return "Error - Improper use of commas"
                Exit Function

            ElseIf InStr(NumberIn, ",") Then

                CommaAdjuster = 0
                WholePart = ""

                For cnt = DecimalPoint - 1 To 1 Step -1

                    If Not Mid(NumberIn, cnt, 1) Like "[,]" Then

                        WholePart = Mid$(NumberIn, cnt, 1) & WholePart

                    Else

                        CommaAdjuster = CommaAdjuster + 1

                        If (DecimalPoint - cnt - CommaAdjuster) Mod 3 Then

                            Return "Error - Improper use of commas"
                            Exit Function

                        End If 'If
                    End If  'If Not
                Next  'For cnt
            End If  'If InStr
        End If  'If Not


        If Left(WholePart, 1) Like "[+-]" Then
            NumberSign = IIf(Left(WholePart, 1) = "-", "Minus ", "Plus ")
            WholePart = Mid$(WholePart, 2)
        End If


        '----------------------------------------
        'Begin code to assure decimal portion of
        'check value is not inadvertently rounded
        '----------------------------------------

        '----------------------------------------
        'Final prep step - this assures number
        'within range of formatting code below
        '----------------------------------------

        '----------------------------------------
        'Begin creating the output string
        '----------------------------------------

        'Very Large values
        TestValue = Val(BigWholePart)

        If TestValue > 999999 Then
            CardinalNumber = TestValue \ 1000000
            tmp = HundredsTensUnits(CardinalNumber) & "Quadrillion "
            TestValue = TestValue - (CardinalNumber * 1000000)
        End If

        If TestValue > 999 Then
            CardinalNumber = TestValue \ 1000
            tmp = tmp & HundredsTensUnits(CardinalNumber) & "Trillion "
            TestValue = TestValue - (CardinalNumber * 1000)
        End If

        If TestValue > 0 Then
            tmp = tmp & HundredsTensUnits(TestValue) & "Billion "
        End If

        'Lesser values
        TestValue = Val(WholePart)

        If TestValue = 0 And BigWholePart = "" Then tmp = "Zero "

        If TestValue > 999999 Then
            CardinalNumber = TestValue \ 1000000
            tmp = tmp & HundredsTensUnits(CardinalNumber) & "Million "
            TestValue = TestValue - (CardinalNumber * 1000000)
        End If

        If TestValue > 999 Then
            CardinalNumber = TestValue \ 1000
            tmp = tmp & HundredsTensUnits(CardinalNumber) & "Thousand "
            TestValue = TestValue - (CardinalNumber * 1000)
        End If

        If TestValue > 0 Then
            If Val(WholePart) < 99 And BigWholePart = "" Then bUseAnd = False
            tmp = tmp & HundredsTensUnits(TestValue, bUseAnd)
        End If

        'If in dollar mode, assure the text is the correct plurality
        If bUseDollars = True Then

            CentsString = HundredsTensUnits(DecimalPart)

            If tmp = "One " Then
                tmp = tmp & "Dollar"
            Else
                tmp = tmp & "Dollars"
            End If

            If Len(CentsString) > 0 Then

                tmp = tmp & " and " & CentsString

                If CentsString = "One " Then
                    tmp = tmp & "Cent"
                Else
                    tmp = tmp & "Cents"
                End If

            End If

        ElseIf bUseCheck = True Then

            tmp = tmp & "and " & Left$(DecimalPart & "00", 2)
            tmp = tmp & "/100"

        Else

            If Len(DecimalPart) > 0 Then

                tmp = tmp & "Point"

                For cnt = 1 To Len(DecimalPart)
                    tmp = tmp & " " & sNumberText(Mid$(DecimalPart, cnt, 1))
                Next

            End If  'If DecimalPart
        End If   'If bUseDollars 


        'done!
        If InStr(tmp, "and") = 1 Then tmp = Mid(tmp, 4)
        Return NumberSign & tmp.Trim
    End Function


End Class