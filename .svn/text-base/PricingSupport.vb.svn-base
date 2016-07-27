Public Class PricingSupport

    Private Shared myChangeDate As Date = Today                 'Date on which the pricing is to change

    Public Shared Sub PriceHasChanged(ByVal discount As Double, ByVal ExistingScheme As String, _
    ByVal LineItem As String, ByVal AClient As Intranet.intranet.customerns.Client, _
    ByVal SchemeType As String, ByVal ds As MedscreenLib.Glossary.DiscountScheme, _
     ByVal DiscType As String)
        If discount = 0 And ExistingScheme.Trim.Length > 0 Then
            RemoveExistingScheme(ExistingScheme, AClient)
        Else

            ''Dim 
            'Dim DiscType As String
            'Dim myItemCode As String = LineItem
            ''do we want a scheme for this sample type or all samples
            'If MsgBox("Do you want to apply this scheme to all Sample types", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            '    DiscType = "TYPE"
            '    myItemCode = "SAMPLE"
            '    ds = FindDiscountScheme(discount, "", "SAMPLE")
            '    ExistingScheme = Me.pnlAO.Scheme & "," & Me.pnlCallOut.Scheme & "," & _
            '        Me.pnlExternal.Scheme & "," & Me.pnlFixedSite.Scheme & "," & Me.pnlRoutine.Scheme
            'Else
            '    DiscType = "CODE"
            '    ds = FindDiscountScheme(discount, myItemCode, "")

            'End If
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                RemoveExistingScheme(ExistingScheme, AClient)
                CreateCustomerScheme(ds, AClient)
                MedscreenLib.Medscreen.LogAction("Scheme added " & ds.FullDescription & " for customer - " & AClient.SMIDProfile)
                'End If
            Else
                MedscreenLib.Medscreen.LogAction("Scheme created " & SchemeType & "-" & discount & " for customer - " & AClient.SMIDProfile)
                CreateScheme(discount, SchemeType, ExistingScheme, LineItem, DiscType, AClient)
            End If
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' get the first number from the string 
    ''' </summary>
    ''' <param name="innum"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	13/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared Function GetNumber(ByVal innum As String) As String
        Dim i As Integer = 0
        Dim strRet As String = ""
        While i < innum.Length AndAlso Char.IsDigit(innum.Chars(i))
            strRet += innum.Chars(i)
            i += 1
        End While
        Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Remove existing schemes
    ''' </summary>
    ''' <param name="ExistingScheme"></param>
    ''' <param name="AClient"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	13/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub RemoveExistingScheme(ByVal ExistingScheme As String, ByVal AClient As Intranet.intranet.customerns.Client)
        If ExistingScheme.Trim.Length > 0 Then      'We need to get rid of existing schemes
            Dim Schemes As String() = ExistingScheme.Split(New Char() {","})
            Dim k As Integer
            For k = 0 To Schemes.Length - 1     'Could be more than one value 
                Dim cd As Intranet.intranet.customerns.CustomerDiscount
                Dim strDiscId As String = Schemes.GetValue(k)
                If strDiscId.Trim.Length <> 0 Then       'Check it is not a dummy string 
                    strDiscId = GetNumber(strDiscId)
                    cd = AClient.InvoiceInfo.DiscountSchemes.GetDiscountByScheme(strDiscId)
                    If Not cd Is Nothing Then       'We've found the item remove it 

                        cd.RemoveScheme(myChangeDate)                  'expire item
                        'remove from collection 
                        AClient.InvoiceInfo.DiscountSchemes.Refresh()
                    End If
                End If
            Next
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Decode the service in use
    ''' </summary>
    ''' <param name="SenderService"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	06/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function IdentifySenderService(ByVal SenderService As String) As String
        Dim strType As String = ""
        If SenderService = "ANALONLY" Then
            strType = "Analysis Only"
        ElseIf SenderService = "EXTERNAL" Then
            strType = "External Sample"
        ElseIf SenderService = "CALLOUT" Then
            strType = "Call-Out Sample"
        ElseIf SenderService = "FIXED" Then
            strType = "Fixed-Site Sample"
        ElseIf SenderService = "ROUTINE" Then
            strType = "Routine Sample"
        End If
        Return strType
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Identify which discount scheme the user thinks is best
    ''' </summary>
    ''' <param name="Discount"></param>
    ''' <param name="ItemCode"></param>
    ''' <param name="SampleType"></param>
    ''' <param name="AClient"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	06/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FindDiscountScheme(ByVal Discount As Double, ByVal ItemCode As String, _
    ByVal SampleType As String, ByVal AClient As Intranet.intranet.customerns.Client) As MedscreenLib.Glossary.DiscountScheme
        Dim ds As MedscreenLib.Glossary.DiscountScheme
        Dim From As Integer = 0
        ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
        If ds Is Nothing Then
            Return Nothing
        Else
            Dim strRet As String = ds.FullDescription(AClient.Identity)
            Dim intRet As MsgBoxResult = MsgBox("Do you want to use this scheme?" & vbCrLf & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            While intRet = MsgBoxResult.No And Not ds Is Nothing
                From += 1
                ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
                If Not ds Is Nothing Then                   'We may kick out with no scheme
                    strRet = ds.FullDescription(AClient.Identity)
                    intRet = MsgBox("Do you want to use this scheme?" & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                End If
            End While
            Return ds
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a scheme and add one entry 
    ''' </summary>
    ''' <param name="Discount"></param>
    ''' <param name="strType"></param>
    ''' <param name="ExistingScheme"></param>
    ''' <param name="ItemCode"></param>
    ''' <param name="DiscountType"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub CreateScheme(ByVal Discount As Double, ByVal strType As String, _
    ByVal ExistingScheme As String, ByVal ItemCode As String, ByVal DiscountType As String, _
    ByVal AClient As Intranet.intranet.customerns.Client)
        'We need to create a scheme
        If Discount = 0 Then Exit Sub 'No discount no need for a scheme
        Dim strSchemeTitle As String = Math.Abs(Discount).ToString("0.00") & "% "
        If Discount < 0 Then
            strSchemeTitle += "Surcharge on "
        Else
            strSchemeTitle += "Discount on "
        End If
        'Check for discount type 
        If DiscountType = "TYPE" Then
            strSchemeTitle += "All "
        End If
        strSchemeTitle += strType
        If MsgBox("create discount scheme for " & strSchemeTitle, MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim objNewScheme As New MedscreenLib.Glossary.DiscountScheme()
            objNewScheme.Description = strSchemeTitle
            objNewScheme.DiscountSchemeId = MedscreenLib.CConnection.NextID("DISCOUNTSCHEME_HEADER", "IDENTITY")
            objNewScheme.Insert()

            objNewScheme.Update()
            MedscreenLib.Glossary.Glossary.DiscountSchemes.Add(objNewScheme)
            Dim objNewEntry As New MedscreenLib.Glossary.DiscountSchemeEntry()
            objNewEntry.ApplyTo = ItemCode
            objNewEntry.Discount = Discount
            objNewEntry.DiscountSchemeId = objNewScheme.DiscountSchemeId
            objNewEntry.DiscountType = DiscountType
            objNewEntry.EntryNumber = 1
            objNewEntry.Insert()
            objNewEntry.Update()
            objNewScheme.Entries.Add(objNewEntry)
            PricingSupport.RemoveExistingScheme(ExistingScheme, AClient)
            CreateCustomerScheme(objNewScheme, AClient)
        End If



    End Sub

    Private Shared strContractAmendment As String = ""


    Public Shared Sub CreateCustomerScheme(ByVal Ds As MedscreenLib.Glossary.DiscountScheme, _
    ByVal AClient As Intranet.intranet.customerns.Client)
        Dim CDexist As Intranet.intranet.customerns.CustomerDiscount = AClient.InvoiceInfo.DiscountSchemes.GetDiscountByScheme(Ds.DiscountSchemeId)
        If CDexist Is Nothing Then
            Dim cd As New Intranet.intranet.customerns.CustomerDiscount() 'Scheme not in list 
            cd.CustomerID = AClient.Identity               'Initialise scheme set customer Id
            cd.DiscountScheme = Ds                          'Set scheme object
            cd.DiscountSchemeId = Ds.DiscountSchemeId       'Set scheme ID
            cd.Priority = AClient.InvoiceInfo.DiscountSchemes.Count + 1         'Set Priority to 1 + number of schemes
            Dim tmpContAmend As String = strContractAmendment
            Dim Remembered As Boolean
            tmpContAmend = MedscreenLib.Medscreen.GetContractAmendment(Remembered, tmpContAmend)
            'Dim objConAmmend As Object = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typString, "Contract Ammendment")
            'If Not objConAmmend Is Nothing Then
            cd.ContractAmendment = CStr(tmpContAmend)
            If Remembered Then strContractAmendment = tmpContAmend 'Remember the contract amendment 

            'End If
            cd.SetEndDateNull()
            cd.Startdate = myChangeDate
            cd.Insert()                                     'Insert and update scheme
            cd.Update()
            AClient.InvoiceInfo.DiscountSchemes.Add(cd) 'Add to customer's discount schemes
        Else 'if we have a discount we need to remove its end date
            CDexist.SetEndDateNull()
            CDexist.Startdate = myChangeDate
            CDexist.Update()

        End If

    End Sub

    Public Shared Function GetChangeDate() As Boolean
        Dim objChangeDate As Object = DateSerial(Today.Year, Today.Month, 1)
        If myChangeDate <> Today Then objChangeDate = myChangeDate
        objChangeDate = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typDate, "Price Change on", "Get Price change date", objChangeDate)
        If objChangeDate Is Nothing Then
            Return False
        Else
            myChangeDate = CDate(objChangeDate)
            Return True
        End If
    End Function

    Public Shared Property ChangeDate() As Date
        Get
            Return myChangeDate
        End Get
        Set(ByVal Value As Date)
            myChangeDate = Value
        End Set
    End Property
End Class
