Imports System.Xml.Linq
Imports System.Data.OleDb


Namespace Address

    Public Enum LabelFormat
        Customer
        SalesOrder
        VDIVessel
    End Enum

#Region "Address"


#Region "Address"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Address.Caddress
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' class representing a row in the address table
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action>Function to validate the Email address and phone/fax numbers  added</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action>Address type missing from fields collection, Fax field corrected</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class Caddress
        'Inherits MedscreenLib.CAddress

        Private myParent As AddressCollection
        Private intUsage As Integer = -1
        Private strCustomer As String = ""


        Private Shared strPCError As String
        Private myFields As TableFields = New TableFields("ADDRESS")


        Private objAddressID As IntegerField = New IntegerField("ADDRESS_ID", 0, True)
        Private objCrmID As StringField = New StringField("CRM_ID", "", 20)
        Private objAddrline1 As StringField = New StringField("ADDRLINE1", "", 60)
        Private objAddrline2 As StringField = New StringField("ADDRLINE2", "", 60)
        Private objAddrline3 As StringField = New StringField("ADDRLINE3", "", 60)
        Private objAddrline4 As StringField = New StringField("ADDRLINE4", "", 60)
        Private objCity As StringField = New StringField("CITY", "", 30)
        Private objDistrict As StringField = New StringField("DISTRICT", "", 30)
        Private objPostcode As StringField = New StringField("POSTCODE", "", 15)
        Private objCountry As IntegerField = New IntegerField("COUNTRY", 0)
        Private objPhone As StringField = New StringField("PHONE", "", 40)
        Private objFax As StringField = New StringField("FAX", "", 40)
        Private objEmail As StringField = New StringField("EMAIL", "", 100)
        Private objContact As StringField = New StringField("CONTACT", "", 40)
        Private objDeleted As BooleanField = New BooleanField("DELETED", "F")
        Private objAddressType As StringField = New StringField("ADDRESS_TYPE", "", 4)
        Private objModifiedON As TimeStampField = New TimeStampField("MODIFIED_ON")
        Private objModifiedBY As UserField = New UserField("MODIFIED_BY")

        Private strStatus As String = ""
        Private strCountry As String = ""


        Public Function ToXMLElement() As XElement
            Dim Element As New XElement("Address")
            Element.Add(New XElement("line1", AdrLine1))
            Element.Add(New XElement("line2", AdrLine2))
            Element.Add(New XElement("line3", AdrLine3))
            Element.Add(New XElement("line4", AdrLine4))
            Element.Add(New XElement("city", City))
            Element.Add(New XElement("district", District))
            Element.Add(New XElement("postcode", PostCode))
            Element.Add(New XElement("contact", Contact))
            Element.Add(New XElement("fax", Fax))
            Element.Add(New XElement("phone", Phone))

            Return Element
        End Function

        Public Function ToXMLElementFull() As XElement
            Dim Element As XElement = ToXMLElement()
            Return Element
        End Function

        Private Sub SetupFields()
            myFields.Add(objAddressID)
            myFields.Add(objCrmID)
            myFields.Add(objAddrline1)
            myFields.Add(objAddrline2)
            myFields.Add(objAddrline3)
            myFields.Add(objAddrline4)
            myFields.Add(objCity)
            myFields.Add(objDistrict)
            myFields.Add(objPostcode)
            myFields.Add(objCountry)
            myFields.Add(objPhone)
            myFields.Add(objFax)
            myFields.Add(objEmail)
            myFields.Add(objContact)
            myFields.Add(objDeleted)
            myFields.Add(objAddressID)
            myFields.Add(objModifiedON)
            myFields.Add(objModifiedBY)
            myFields.Add(objAddressType)    'Missing!

        End Sub


        '''<summary>
        ''' Constructor creates an address object
        ''' </summary>
        Public Sub New()
            SetupFields()
        End Sub

        '''<summary>
        ''' Constructor creates an address object
        ''' </summary>
        ''' <param name='ID'>ID of the address in the address table </param>
        Public Sub New(ByVal ID As Integer)
            SetupFields()
            Me.objAddressID.Value = ID
            Me.objAddressID.OldValue = ID
            Me.Load()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' See if a link value is in the addresslink table as the right type
        ''' </summary>
        ''' <param name="CustomerID">LinkValue</param>
        ''' <param name="LinkType">Type of addresslink</param>
        ''' <returns>True if Link exists</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2006]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function HasCustomerAddressLink(ByVal CustomerID As String, ByVal LinkType As Integer) As Boolean
            Dim objReturn As Object
            Dim oCmd As New OleDb.OleDbCommand("Select lib_address.GetCustomerAddressLink(?, ?, ?) from dual", MedConnection.Connection)
            Try
                If CConnection.ConnOpen Then
                    oCmd.Parameters.Add(CConnection.IntegerParameter("addrid", Me.AddressId))
                    oCmd.Parameters.Add(CConnection.StringParameter("LinkValue", CustomerID, 30))
                    oCmd.Parameters.Add(CConnection.IntegerParameter("LinkType", LinkType))
                    objReturn = oCmd.ExecuteScalar
                    If objReturn Is Nothing Or objReturn Is System.DBNull.Value Then
                        Return False
                    Else
                        Return True
                    End If
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Getting addresslink")
            Finally
                CConnection.SetConnClosed()
            End Try
        End Function

        Public Function Exists() As Boolean

        End Function

        Public Function GetAddressLink(ByVal CustomerID As String, ByVal AddressType As Integer) As Integer
            Dim paramColl As New Collection()
            paramColl.Add(CStr(Me.AddressId))
            paramColl.Add(CustomerID)
            paramColl.Add(CStr(AddressType))

            Dim objRet As String = CConnection.PackageStringList("lib_address.GetCustomerAddressLink", paramColl)
            If objRet.Trim.Length = 0 Then objRet = "0"
            Return CInt(objRet)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Set the addresslink
        ''' </summary>
        ''' <param name="CustomerID"></param>
        ''' <param name="AddressType"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>ressp
        ''' <history>
        ''' 	[taylor]	13/04/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Function SetAddressLink(ByVal CustomerID As String, ByVal AddressType As Integer) As Boolean
            Dim blnRet As Boolean = False
            Dim oCmd As New OleDb.OleDbCommand("lib_address.CreateCustomerAddressLink", MedConnection.Connection)
            Try
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.Parameters.Add(CConnection.IntegerParameter("AddressId", Me.AddressId))
                oCmd.Parameters.Add(CConnection.StringParameter("CustomerId", CustomerID, 30))
                oCmd.Parameters.Add(CConnection.IntegerParameter("AddressType", AddressType))
                If CConnection.ConnOpen Then
                    Dim intRet As Integer
                    intRet = oCmd.ExecuteNonQuery
                    blnRet = (intRet = 1)
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Setting Address link")
            Finally
                CConnection.SetConnClosed()
            End Try
            Return blnRet
        End Function



        '''<summary>
        ''' Creates an XML representation of the address object
        ''' </summary>
        ''' <returns> returns a string containg XML </returns>
        Public Overridable Function ToXML(Optional ByVal blnParent As Boolean = False) As String
            Dim strRet As String = ""
            If Me.AddressId > 0 Then
                Dim oCmd As New OleDb.OleDbCommand()
                'Dim oread As OleDb.OleDbDataReader
                Try
                    Dim OColl As New Collection()
                    OColl.Add(CStr(Me.AddressId))

#If R2004 Then
                    If blnParent Then
                        OColl.Add("Parent address ")
                    Else
                        OColl.Add(" ")
                    End If
#Else
#End If
                    strRet = CConnection.PackageStringList("Lib_Address.RowToXml", OColl)
                Catch ex As Exception
                    Medscreen.LogError(ex, , "addres ID " & Me.AddressId)
                Finally
                    'CConnection.SetConnClosed()             'Close connection
                    'If Not oread Is Nothing Then            'Try and close reader
                    '    If Not oread.IsClosed Then oread.Close()
                    'End If


                End Try

            End If
            'Dim strRet As String = "<Address>"

            'strRet += "<addressid>" & Medscreen.FixAmpersands(Me.AddressId) & "</addressid>"
            'If InStr(Me.AdrLine1, "<") > 0 Then
            '    Me.AdrLine1 = Me.AdrLine1.Replace("<", "[")
            '    Me.Update()
            'End If
            'If InStr(Me.AdrLine1, ">") > 0 Then
            '    Me.AdrLine1 = Me.AdrLine1.Replace(">", "]")
            '    Me.Update()
            'End If

            'strRet += "<line1>" & Medscreen.FixAmpersands(Me.AdrLine1) & "</line1>"
            'strRet += "<line2>" & Medscreen.FixAmpersands(Me.AdrLine2) & "</line2>"
            'strRet += "<line3>" & Medscreen.FixAmpersands(Me.AdrLine3) & "</line3>"
            'strRet += "<line4>" & Medscreen.FixAmpersands(Me.AdrLine4) & "</line4>"
            'strRet += "<city>" & Medscreen.FixAmpersands(Me.City) & "</city>"
            'strRet += "<district>" & Medscreen.FixAmpersands(Me.District) & "</district>"
            'strRet += "<postcode>" & Medscreen.FixAmpersands(Me.PostCode) & "</postcode>"
            'strRet += "<contact>" & Medscreen.FixAmpersands(Me.Contact) & "</contact>"
            'strRet += "<email>" & Medscreen.FixAmpersands(Me.Email) & "</email>"
            'If Me.Country > 0 Then
            '    strRet += "<phone> +" & CStr(Country) & " " & Medscreen.FixAmpersands(Me.Phone) & "</phone>"
            '    strRet += "<fax>  +" & CStr(Country) & " " & Medscreen.FixAmpersands(Me.Fax) & "</fax>"
            'Else
            '    strRet += "<phone> " & Medscreen.FixAmpersands(Me.Phone) & "</phone>"
            '    strRet += "<fax>  " & Medscreen.FixAmpersands(Me.Fax) & "</fax>"
            'End If
            'If Not Me.CountryName Is Nothing AndAlso Me.CountryName.Trim.Length = 0 Then
            '    If Me.Country > 0 Then
            '        Try
            '            Dim oCmd As New OleDb.OleDbCommand("Select country_name from country where country_id = " & Country, medConnection.Connection)
            '            CConnection.SetConnOpen()
            '            Me.CountryName = oCmd.ExecuteScalar
            '        Catch ex As Exception
            '        Finally
            '            CConnection.SetConnClosed()
            '        End Try
            '    End If
            'End If
            'strRet += "<countryid>" & Me.Country & "</countryid>"
            'strRet += "<country>" & Medscreen.FixAmpersands(Me.CountryName) & "</country>"


            'strRet += "</Address>" & vbCrLf
            Return strRet
        End Function
        '''<summary>
        ''' Status of the address this alllows an address 
        ''' to be removed or other possible settings
        ''' </summary>
        ''' 
        Public Property Status() As String
            Get
                Return strStatus
            End Get
            Set(ByVal Value As String)
                strStatus = Value
            End Set
        End Property

        '''<summary>
        ''' First line of the address
        ''' </summary>
        Public Property AdrLine1() As String
            Get
                Return Me.objAddrline1.Value
            End Get
            Set(ByVal Value As String)
                Me.objAddrline1.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Second line of the address
        ''' </summary>
        Public Property AdrLine2() As String
            Get
                Return Me.objAddrline2.Value
            End Get
            Set(ByVal Value As String)
                objAddrline2.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Third line of the address
        ''' </summary>
        Public Property AdrLine3() As String
            Get
                Return Me.objAddrline3.Value
            End Get
            Set(ByVal Value As String)
                objAddrline3.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Fourth line of the address
        ''' </summary>
        Public Property AdrLine4() As String
            Get
                Return Me.objAddrline4.Value
            End Get
            Set(ByVal Value As String)
                objAddrline4.Value = Value
            End Set
        End Property

        '''<summary>
        ''' City or postal town
        ''' </summary>
        Public Property City() As String
            Get
                Return Me.objCity.Value
            End Get
            Set(ByVal Value As String)
                objCity.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Contact associated with this address
        ''' </summary>
        Public Property Contact() As String
            Get
                Dim tmpContact As String = objContact.Value
                tmpContact = Medscreen.ReplaceString(tmpContact, "MR ", "Mr ")
                tmpContact = Medscreen.ReplaceString(tmpContact, "MRS ", "Mrs ")
                Return tmpContact
            End Get
            Set(ByVal Value As String)
                objContact.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Email address associated with the address
        ''' </summary>
        Public Property Email() As String
            Get
                Return Me.objEmail.Value
            End Get
            Set(ByVal Value As String)
                objEmail.Value = Value
            End Set
        End Property

        Property GeoCode() As String
            Get
                Return Me.objCrmID.Value
            End Get
            Set(ByVal value As String)
                objCrmID.Value = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Validate the phone number 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ValidatePhone() As Constants.PhoneNoErrors
            Return Caddress.ValidatePhoneNum(Me.Phone)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Check the validity of a fax number 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ValidateFax() As Constants.PhoneNoErrors
            Return Caddress.ValidatePhoneNum(Me.Fax)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Validate the Email address given
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ValidateEmail() As Constants.EmailErrors
            Return MedscreenLib.Medscreen.ValidateEmail(Me.Email)
            'Dim myReturn As Constants.EmailErrors = Constants.EmailErrors.None
            'Dim intPos As Integer
            'If Me.Email.Trim.Length > 0 Then 'We have something to work on 
            '    intPos = InStr(Me.Email, "@")       'See if we have an 'at
            '    If intPos = 0 Then                  'No hat
            '        myReturn = Constants.EmailErrors.NoAt
            '    Else                                'got an hat 
            '        Dim intpos2 As Integer = InStr(intPos, Me.Email, ".")
            '        If intpos2 = 0 Then             'Is there a '.' after the @
            '            myReturn = Constants.EmailErrors.NoDomain
            '        End If
            '    End If
            'End If
            'Return myReturn
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Check that a phone number contains only digits or spaces
        ''' </summary>
        ''' <param name="InPhoneNumber">Phone number to check</param>
        ''' <returns>
        '''A value from the enumeration</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ValidatePhoneNum(ByVal InPhoneNumber As String) As Constants.PhoneNoErrors
            Dim myRet As Constants.PhoneNoErrors = Constants.PhoneNoErrors.NoError

            'check for illegal characters only digits and space allowed
            Dim i As Integer
            Dim inPhoneNum As String = InPhoneNumber.Trim
            For i = 0 To inPhoneNum.Length - 1
                If Not (inPhoneNum.Chars(i) = " " Or _
                   Char.IsDigit(inPhoneNum.Chars(i)) Or _
                   inPhoneNum.Chars(i) = "-" Or inPhoneNum.Chars(i) = "(" Or inPhoneNum.Chars(i) = ")") Then
                    If inPhoneNum.Chars(i) = "+" AndAlso i = 0 Then
                        myRet = Constants.PhoneNoErrors.InTDialCodePresent
                        Exit For
                    Else
                        myRet = Constants.PhoneNoErrors.IllegalCharacterPresent
                        Exit For
                    End If
                End If
            Next

            Return myRet

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Translate an XML address into an address
        ''' </summary>
        ''' <param name="XMlIn"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub ReadXML(ByVal XMlIn As Xml.XmlElement)
            Dim anode As Xml.XmlNode
            For Each anode In XMlIn
                Debug.WriteLine(anode.Name & "-" & anode.InnerXml)
                If anode.Name.ToUpper = "ADDRESSID" Then
                    Me.AddressId = anode.InnerText
                ElseIf anode.Name.ToUpper = "LINE1" Then
                    Me.AdrLine1 = anode.InnerText
                ElseIf anode.Name.ToUpper = "LINE2" Then
                    Me.AdrLine2 = anode.InnerText
                ElseIf anode.Name.ToUpper = "LINE3" Then
                    Me.AdrLine3 = anode.InnerText
                ElseIf anode.Name.ToUpper = "LINE4" Then
                    Me.AdrLine4 = anode.InnerText
                ElseIf anode.Name.ToUpper = "CITY" Then
                    Me.City = anode.InnerText
                ElseIf anode.Name.ToUpper = "DISTRICT" Then
                    Me.District = anode.InnerText
                ElseIf anode.Name.ToUpper = "POSTCODE" Then
                    Me.PostCode = anode.InnerText
                ElseIf anode.Name.ToUpper = "CONTACT" Then
                    Me.Contact = anode.InnerText
                ElseIf anode.Name.ToUpper = "EMAIL" Then
                    Me.Email = anode.InnerText
                ElseIf anode.Name.ToUpper = "PHONE" Then
                    Me.Phone = anode.InnerText
                ElseIf anode.Name.ToUpper = "FAX" Then
                    Me.Fax = anode.InnerText
                ElseIf anode.Name.ToUpper = "COUNTRYID" Then
                    Me.Country = anode.InnerText

                End If
            Next

        End Sub


        '''<summary>
        ''' Fax number for the address
        ''' </summary>
        Public Property Fax() As String
            Get
                Return Me.objFax.Value
            End Get
            Set(ByVal Value As String)
                objFax.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Has this address been deleted 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [11/09/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Deleted() As Boolean
            Get
                Return Me.objDeleted.Value
            End Get
            Set(ByVal Value As Boolean)
                Me.objDeleted.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Country ID this is the International dialling code prefix for this 
        ''' country, exceptions are Carribbean islands where it is the prefix plus dialling code
        ''' </summary>
        ''' <seealso cref="Country"/>
        Public Property Country() As Integer
            Get
                Return Me.objCountry.Value
            End Get
            Set(ByVal Value As Integer)
                objCountry.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The name of the country as stored in the country table
        ''' </summary>
        Public Property CountryName() As String
            Get
                Return Me.strCountry
            End Get
            Set(ByVal Value As String)
                strCountry = Value
            End Set
        End Property

        '''<summary>
        ''' In britain this is the County in the US it is the state etc.
        ''' </summary>
        Public Property District() As String
            Get
                Return Me.objDistrict.Value
            End Get
            Set(ByVal Value As String)
                objDistrict.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The type of address for example Customer, Site Etc
        ''' </summary>
        Public Property AddressType() As String
            Get
                Return Me.objAddressType.Value
            End Get
            Set(ByVal Value As String)
                objAddressType.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The postal code, look at the IUPAC definitions
        ''' </summary>
        Public Property PostCode() As String
            Get
                Return Me.objPostcode.Value
            End Get
            Set(ByVal Value As String)
                objPostcode.Value = Value
            End Set
        End Property


        '''<summary>
        ''' The Id associated with the address, its primary key
        ''' </summary>
        Public Property AddressId() As Long
            Get
                Return Me.objAddressID.Value
            End Get
            Set(ByVal Value As Long)
                objAddressID.Value = Value
                objAddressID.OldValue = Value
            End Set
        End Property

        '''<summary>
        ''' Phone number for this address
        ''' </summary>
        Public Property Phone() As String
            Get
                Return Me.objPhone.Value
            End Get
            Set(ByVal Value As String)
                objPhone.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Return phone with Country ID
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property PhoneFormatted() As String
            Get
                If Me.Country > 0 AndAlso Me.Phone.Trim.Length > 0 Then
                    If Me.Country = 1002 Then
                        Return "+1" & " " & Me.Phone
                    Else
                        Return "+" & CStr(Me.Country) & " " & Me.Phone
                    End If

                Else
                    Return Me.Phone
                End If
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Fax number formatted with country code 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property FaxFormatted() As String
            Get
                If Me.Country > 0 AndAlso Me.Fax.Trim.Length > 0 Then
                    If Me.Country = 1002 Then
                        Return "+1" & " " & Me.Fax
                    Else
                        Return "+" & CStr(Me.Country) & " " & Me.Fax
                    End If

                Else
                    Return Me.Fax
                End If
            End Get
        End Property
        '''<summary>
        ''' Date address last modified (a timestamp field)
        ''' </summary>
        Public Property DateModified() As Date
            Get
                Return Me.objModifiedON.Value
            End Get
            Set(ByVal Value As Date)
                objModifiedON.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The user who last modified these data
        ''' </summary>
        Public Property ModifiedBy() As String
            Get
                Return Me.objModifiedBY.Value
            End Get
            Set(ByVal Value As String)
                objModifiedBY.Value = Value
            End Set
        End Property


        '''<summary>
        ''' Link to Goldmine, this is future development
        ''' </summary>
        Public Property GoldId() As String
            Get
                Return Me.objCrmID.Value
            End Get
            Set(ByVal Value As String)
                objCrmID.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Exposes the fields in the row as a dataset
        ''' </summary>
        Public Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property




        '''<summary>
        ''' display the address as a block
        ''' </summary>
        ''' <param name='strLineBreak'>
        '''     the line break character used eg &lt;BR&gt; or VBCRLF
        ''' </param>
        ''' <param name='UseContact'>
        '''     indicates whether the contact name will be displayed as well
        ''' </param>
        ''' <returns> A string containing the formattted address </returns>
        Public Function BlockAddress(ByVal strLineBreak As String, Optional ByVal UseContact As Boolean = False, Optional ByVal IncludePhone As Boolean = False) As String
            Dim strAddress As String = ""

            With Me
                If UseContact Then
                    If .Contact.Trim.Length > 0 Then strAddress += .Contact
                    If IncludePhone Then
                        If .Phone.Trim.Length > 0 Then strAddress += " (" & .Phone & ")"
                    End If
                    If strAddress.Trim.Length > 0 Then strAddress += strLineBreak
                End If

                If .AdrLine1.Length > 0 Then strAddress += .AdrLine1 & "," & strLineBreak
                If .AdrLine2.Length > 0 Then strAddress += .AdrLine2 & "," & strLineBreak
                If .AdrLine3.Length > 0 Then strAddress += .AdrLine3 & "," & strLineBreak
                If .AdrLine4.Length > 0 Then strAddress += .AdrLine4 & strLineBreak
                Select Case .Country
                    Case 45, 47, 49, 33, 65, 46, 30
                        If .Country = 65 Then
                            strAddress += "SINGAPORE "
                        End If
                        If .PostCode.Length > 0 Then strAddress += .PostCode & " "
                        If .City.Length > 0 Then
                            strAddress += .City & strLineBreak
                        Else
                            strAddress += strLineBreak
                        End If
                        If .Country = 45 Then
                            strAddress += "DENMARK" & strLineBreak
                        ElseIf .Country = 47 Then
                            strAddress += "NORWAY" & strLineBreak
                        ElseIf .Country = 49 Then
                            strAddress += "GERMANY" & strLineBreak
                        ElseIf .Country = 30 Then
                            strAddress += "GREECE" & strLineBreak
                        ElseIf .Country = 33 Then
                            strAddress += "FRANCE" & strLineBreak
                        ElseIf .Country = 46 Then
                            strAddress += "SWEDEN" & strLineBreak
                        ElseIf .Country = 65 Then
                            strAddress += "REPUBLIC OF SINGAPORE" & strLineBreak
                        End If
                    Case 27
                        If .City.Length > 0 Then strAddress += .City & strLineBreak
                        If .District.Length > 0 Then strAddress += .District & strLineBreak
                        If .PostCode.Length > 0 Then strAddress += .PostCode & " SOUTH AFRICA" & strLineBreak

                    Case 44
                        If .City.Length > 0 Then strAddress += .City & strLineBreak
                        If .District.Length > 0 Then strAddress += .District & strLineBreak
                        If .PostCode.Length > 0 Then strAddress += .PostCode & strLineBreak
                    Case Else
                        If .PostCode.Length > 0 Then strAddress += .PostCode & " "
                        If .City.Length > 0 Then
                            strAddress += .City & strLineBreak
                        Else
                            strAddress += strLineBreak
                        End If
                        strAddress += .CountryName & strLineBreak
                        'dr = myGlossaries.Countrylist.TOS_COUNTRY.Select("Country_ID = " & .Country)
                        'If dr.Length = 1 Then
                        '    cnt = dr(0)
                        '    straddress += UCase(cnt.COUNTRY_NAME) & strlinebreak
                        'End If
                End Select
            End With
            Return strAddress
        End Function

        '''<summary>
        ''' Check that a postcode conforms to the Royal Mail formats for UK postcodes
        ''' </summary>
        ''' <remarks>
        '''  This routine disregards leading and trailing spaces but there
        '''must only be one space between outcodes and incodes
        '''<para />
        '''<para>              Valid UK postcode formats </para>
        '''<para>               Outcode Incode Example</para>
        '''<para>               AN      NAA     B1 6AD</para>
        '''<para>               ANN     NAA     S31 2BD</para>
        '''<para>               AAN     NAA     SW5 8SG</para>
        '''<para>               ANA     NAA     W1A 4DJ</para>
        '''<para>               AANN    NAA     CB10 2BQ</para>
        '''<para>               AANA    NAA     EC2A 1HQ</para>
        '''<para/>
        '''<para>               Incode letters AA cannot be one of C,I,K,M,O or V.</para>
        ''' </remarks>
        ''' <example>
        '''             ' Usage:    If Not Valid_UKPostcode(Me!PostCode) Then  <para/>
        '''                       MsgBox "Invalid postcode format",vbInformation <para/>
        '''               End If
        '''</example>
        ''' <param name='sPostcode'>Postcode to be validated</param>
        ''' <returns>True (-1)  -  Postcode conforms to valid pattern
        '''               False (0)   -  Postcode has failed pattern matching
        ''' </returns>
        Public Shared Function IsValidUKPostcode(ByVal sPostcode As String) As Boolean
            ' Function: IsValidUKPostcode
            '
            ' Purpose:  Check that a postcode conforms to the Royal Mail formats for UK
            ' postcodes
            '
            '
            ' Colin Byrne (100551,2730@Compuserve.com)
            '
            Dim sOutCode As String
            Dim sInCode As String

            Dim bValid As Boolean
            Dim iSpace As Integer

            strPCError = ""
            ' Trim leading and trailing spaces
            sPostcode = Trim(sPostcode)

            iSpace = InStr(sPostcode, " ")

            bValid = True

            '  If there is no space in the string then it is not a full postcode
            If iSpace = 0 Or sPostcode = "" Then
                IsValidUKPostcode = False
                strPCError = "No space in Post Code"
                Exit Function
            End If

            '  Split post code into outcode and incodes
            sOutCode = Left$(sPostcode, iSpace - 1)
            sInCode = Mid$(sPostcode, iSpace + 1)

            '  Check incode is valid
            '  ... this will also test that the length is a valid 3 characters long
            bValid = MatchPattern(sInCode, "NAA")
            If Not bValid Then
                strPCError = "InCode (" & sInCode & ")is badly formed!"
                Exit Function
            End If

            If bValid Then
                '  Test second and third characters for invalid letters
                If InStr("CIKMOV", Mid$(sInCode, 2, 1)) > 0 Or InStr("CIKMOV", Mid$(sInCode, 3, 1)) > 0 Then
                    bValid = False
                    strPCError = "Illegal Character 'CIKMOV' in Incode(" & sInCode & ")"
                    Exit Function
                End If
            End If

            If bValid Then
                Select Case Len(sOutCode)
                    Case 0, 1
                        bValid = False
                    Case 2
                        bValid = MatchPattern(sOutCode, "AN")
                    Case 3
                        bValid = MatchPattern(sOutCode, "ANN") Or MatchPattern(sOutCode, "AAN") Or MatchPattern(sOutCode, "ANA")
                    Case 4
                        bValid = MatchPattern(sOutCode, "AANN") Or MatchPattern(sOutCode, "AANA")
                End Select
            End If

            ' If bValid is False by the time it gets here
            ' ...it has failed one of the above tests
            If Not bValid Then
                strPCError = "Outcode (" & sOutCode & ")is badly formed!"
            End If
            IsValidUKPostcode = bValid

        End Function

        '''<summary>
        ''' Indicates that there is an error in the post code
        ''' </summary>
        ''' <returns> String containg the error text</returns>
        Public Function PostCodeError() As String
            If Me.Country = 44 Then
                Return strPCError
            Else
                Return ""
            End If
        End Function

        '''<summary>
        ''' Checks a string against a supplied pattern
        ''' </summary>
        ''' <param name='sString'>Supplied String</param>
        ''' <param name='sPattern'>Pattern to check for</param>
        ''' <returns>True indicating pattern is present</returns>
        ''' 
        Public Shared Function MatchPattern(ByVal sString As String, ByVal sPattern As String) As Boolean

            Dim cPattern As String
            Dim cString As String

            Dim iPosition As Integer
            Dim bMatch As Boolean

            ' If the lengths don't match then it fails the test
            If Len(sString) <> Len(sPattern) Then
                MatchPattern = False
                Exit Function
            End If

            ' All strings to uppercase - ByVal ensures callers string is not affected
            sString = UCase(sString)
            sPattern = UCase(sPattern)

            ' Assume it matches until proven otherwise
            bMatch = True

            For iPosition = 1 To Len(sString)

                ' Take the characters at the current position from both strings
                cPattern = Mid$(sPattern, iPosition, 1)
                cString = Mid$(sString, iPosition, 1)

                ' See if the source character conforms to the pattern one
                Select Case cPattern
                    Case "N"                ' Numeric
                        If Not IsNumeric(cString) Then bMatch = False
                    Case "A"                ' Alphabetic
                        If Not (cString >= "A" And cString <= "Z") Then bMatch = False
                End Select

            Next iPosition

            MatchPattern = bMatch

        End Function

        '''<summary>
        ''' Constructor for an Address
        ''' </summary>
        ''' <param name='ID'> ID of the address</param>
        ''' <param name='Parent'> Collection in which the address is to be added</param>
        Public Sub New(ByVal ID As Integer, ByVal Parent As AddressCollection)
            SetupFields()
            Me.AddressId = ID
            Me.objAddressID.Value = ID
            Me.objAddressID.OldValue = ID
            Me.objAddressID.SetIsNull = False


        End Sub


        '''<summary>
        ''' Refresh the contents of the object against the database
        ''' </summary>
        ''' <returns>True succesful refresh, note this doesn't indicate that the data has changed
        ''' </returns>
        Public Function Refresh() As Boolean
            Dim ocmd As New OleDb.OleDbCommand()
            Dim oread As OleDb.OleDbDataReader = Nothing
            'Dim objTemp As Object
            If Me.objAddressID.IsNull Then Exit Function
            Try
                ocmd.Connection = MedscreenLib.MedConnection.Connection
                ocmd.CommandText = Fields.FullRowSelect & " where ADDRESS_ID = ? "  ' & Me.AddressId
                ocmd.Parameters.Add(Me.Fields.SetUpParameter(Me.objAddressID))
                'oConn.Open()
                MedscreenLib.CConnection.SetConnOpen()
                oread = ocmd.ExecuteReader
                If oread.Read Then
                    'DT = oread.GetSchemaTable()
                    Fields.readfields(oread)

                Else

                End If

            Catch ex As Exception

            Finally
                If Not oread Is Nothing Then
                    If Not oread.IsClosed Then oread.Close()
                End If
                MedscreenLib.CConnection.SetConnClosed()
                'oConn.Close()
            End Try

        End Function

        '''<summary>
        ''' Id in Goldmine if exists Note: this is for future developments
        ''' </summary>
        Public Property GoldmineID() As String
            Get
                Return Me.GoldId
            End Get
            Set(ByVal Value As String)
                GoldId = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write the address contents back to the database
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [17/10/2005]</date><Action>Modified to prevent writes to Zero</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Update() As Boolean
            'Check to see if we need to insert 

            If Me.AddressId <= 0 Then           'Check that this is a valid ID
                Return False
                Exit Function
            End If

            'See if country is uk and we have no geocode
            If Country = 44 And GeoCode.Trim.Length = 0 Then
                GeoCode = MedscreenLib.GeoCode.GetLatLong(PostCode)
                If GeoCode.Trim.Length = 0 Then
                    Dim ipos As Integer = InStr(PostCode, " ")
                    If ipos > 0 Then
                        Dim Outcode As String = Mid(PostCode, 1, ipos - 1)
                        GeoCode = MedscreenLib.GeoCode.GetLatLong(Outcode)
                    End If
                End If
            End If

            'Sort out insert date 
            If Not Me.Fields.Loaded Or Me.Fields.RowID = "" Then
                Me.DateModified = Now   'Set modified date
                If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then _
                    Me.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                Me.Fields.Insert(MedscreenLib.CConnection.DbConnection)
            End If

            If Me.Fields.Changed Then
                Me.DateModified = Now   'Set modified date
                If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then _
                    Me.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            End If
            Return Fields.Update(MedscreenLib.CConnection.DbConnection)
        End Function

        '''<summary>
        ''' Find from the address link table the customer associated with the address
        ''' </summary>
        ''' <returns>Customer ID as a string</returns>
        Public Function Customer() As String
            If strCustomer.Trim.Length > 0 Then
                Return strCustomer
            Else
                Dim strRet As String = ""
                Dim oRead As OleDb.OleDbDataReader = Nothing
                Dim blnMultiple As Boolean = False

                Try
                    Dim oCmd As New OleDb.OleDbCommand()
                    oCmd.CommandText = "Select distinct linkvalue from addresslink where address_id = " & Me.AddressId & _
                        " and LINKFIELD = 'Customer.Identity'"
                    oCmd.Connection = MedscreenLib.MedConnection.Connection
                    MedscreenLib.CConnection.SetConnOpen()
                    oRead = oCmd.ExecuteReader
                    If Not oRead Is Nothing Then
                        While oRead.Read
                            If strRet.Trim.Length = 0 Then
                                strRet = oRead.GetValue(0)
                            Else
                                blnMultiple = True
                            End If
                        End While
                    End If

                    If blnMultiple Then strRet = ""
                Catch ex As Exception
                Finally
                    If Not oRead Is Nothing Then
                        If Not oRead.IsClosed Then oRead.Close()
                    End If
                    MedscreenLib.CConnection.SetConnClosed()
                End Try
                strCustomer = strRet
                Return strRet
            End If
        End Function

        '''<summary>
        ''' Count the number of times this address can be found in the addresslink table
        ''' </summary>
        ''' <returns>Number of times used as integer</returns>
        Public Function Usage() As Integer
            If intUsage = -1 Then
                Try
                    Dim oCmd As New OleDb.OleDbCommand()
                    oCmd.CommandText = "Select count(*) from addresslink where address_id = " & Me.AddressId
                    oCmd.Connection = MedscreenLib.MedConnection.Connection
                    MedscreenLib.CConnection.SetConnOpen()
                    intUsage = oCmd.ExecuteScalar()
                Catch ex As Exception
                Finally
                    MedscreenLib.CConnection.SetConnClosed()
                End Try
            End If
            Return intUsage
        End Function

        ''' <summary>
        ''' Created by taylor on ANDREW at 23/01/2007 06:58:54
        '''     Load the address from the database
        ''' </summary>
        ''' <returns>
        '''     A System.Boolean value...
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>taylor</Author><date> 23/01/2007</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' 
        Public Function Load() As Boolean
            Return Me.Fields.Load(MedConnection.Connection, "address_id = " & Me.AddressId)
        End Function
    End Class

#End Region


#Region "AddressCollection"

    '''<summary>
    '''     The address collection is a collection of address objects, 
    '''     it actually represents the Sample Manager Address table
    ''' </summary>
    Public Class AddressCollection
        Inherits CollectionBase

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Enumeration describing the different ways of accessing an item form the collection
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Enum ItemType
            ''' <summary>Access by Address Id</summary>
            AddressId
            ''' <summary>Access by Index</summary>
            Index
        End Enum

        'Private oConn As New OleDb.OleDbConnection()
        Private oCmd As New OleDb.OleDbCommand()
        Private oRead As OleDb.OleDbDataReader
        'Private myDataset As Addresses = Nothing

        '''<summary>
        '''     overload of the simple property
        ''' </summary>
        ''' <param name='intAddressId'>Used to provide the index into the collection
        ''' </param>
        ''' <param name='intOpt'>Used to provide the index into the collection
        ''' </param>
        Default Public Overloads Property Item(ByVal intAddressId As Integer, _
            Optional ByVal intOpt As ItemType = ItemType.AddressId) As Caddress
            Get
                Dim anAddress As Caddress = Nothing
                If intOpt = ItemType.AddressId Then
                    For Each anAddress In List
                        If anAddress.AddressId = intAddressId Then
                            Exit For
                        End If
                        anAddress = Nothing
                    Next
                    If anAddress Is Nothing Then        'Okay didn't find the address try and load it 
                        anAddress = GetAddress(intAddressId)
                        If Not anAddress Is Nothing Then
                            Me.Add(anAddress)
                        End If
                    End If
                ElseIf intOpt = ItemType.Index Then         'Use default indexing 
                    anAddress = Me.List.Item(intAddressId)
                End If

                'Refresh the address prior to returning it to the user
                If Not anAddress Is Nothing Then anAddress.Refresh()
                Return anAddress
            End Get
            Set(ByVal Value As Caddress)

            End Set
        End Property

        '''<summary>
        '''     Loads the contents of the entire table into memory 
        '''</summary>
        '''<returns> True if the data is succesfully loaded</returns>
        Public Function Load() As Boolean
            Dim ocmd As New OleDb.OleDbCommand()
            Dim oread As OleDb.OleDbDataReader = Nothing
            'Dim objTemp As Object
            Dim oAddr As Caddress

            Try
                ocmd.Connection = MedscreenLib.MedConnection.Connection
                ocmd.CommandText = "Select a.rowid,a.*  from address a "

                If Not MedscreenLib.Glossary.Glossary.Countries Is Nothing Then

                End If
                'oConn.Open()
                MedscreenLib.CConnection.SetConnOpen()
                oread = ocmd.ExecuteReader
                While oread.Read
                    oAddr = New Caddress(oread.GetValue(oread.GetOrdinal("ADDRESS_ID")), Me)
                    oAddr.Fields.readfields(oread)
                    If Not MedscreenLib.Glossary.Glossary.Countries Is Nothing Then
                        Dim oCountry As Country = MedscreenLib.Glossary.Glossary.Countries.Item(oAddr.Country, 0)
                        If Not oCountry Is Nothing Then
                            oAddr.CountryName = oCountry.CountryName
                        End If
                    End If
                    Me.Add(oAddr)               ' Add to collection
                End While
            Catch ex As Exception

            Finally
                If Not oread Is Nothing Then
                    If Not oread.IsClosed Then oread.Close()
                End If
                MedscreenLib.CConnection.SetConnClosed()
                'oConn.Close()
            End Try

        End Function

        '''<summary>
        '''     add an address element into the collection
        ''' </summary>
        '''<param name='Address'>Address Object to add to collection</param>
        Public Sub Add(ByVal Address As Caddress)
            Me.List.Add(Address)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' See if there is a preexisting link 
        ''' </summary>
        ''' <param name="addressID"></param>
        ''' <param name="LinkField"></param>
        ''' <param name="LinkValue"></param>
        ''' <param name="TypeId"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Function HasLink(ByVal addressID As Integer, _
            ByVal LinkField As String, _
            ByVal LinkValue As String, _
            ByVal TypeId As Integer) As Boolean
            Dim objReturn As Object
            Dim oCmd As New OleDb.OleDbCommand("Select lib_address.GetAddressLink(?, ?, ?, ?) from dual", MedConnection.Connection)
            Try
                If CConnection.ConnOpen Then
                    oCmd.Parameters.Add(CConnection.IntegerParameter("addrid", addressID))
                    oCmd.Parameters.Add(CConnection.StringParameter("LinkField", LinkField, 50))
                    oCmd.Parameters.Add(CConnection.StringParameter("LinkValue", LinkValue, 30))
                    oCmd.Parameters.Add(CConnection.IntegerParameter("LinkType", TypeId))
                    objReturn = oCmd.ExecuteScalar
                    If objReturn Is Nothing Or objReturn Is System.DBNull.Value Then
                        Return False
                    Else
                        Return True
                    End If
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Getting addresslink")
            Finally
                CConnection.SetConnClosed()
            End Try

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create Link for address
        ''' </summary>
        ''' <param name="addressID"></param>
        ''' <param name="LinkField"></param>
        ''' <param name="LinkValue"></param>
        ''' <param name="TypeId"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Function CreateAddresslink(ByVal addressID As Integer, _
            ByVal LinkField As String, _
            ByVal LinkValue As String, _
            ByVal TypeId As Integer) As Boolean
            Dim blnRet As Boolean = False
            Dim oCmd As New OleDb.OleDbCommand("lib_address.CreateAddressLink", MedConnection.Connection)
            Try
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.Parameters.Add(CConnection.IntegerParameter("AddressId", addressID))
                oCmd.Parameters.Add(CConnection.StringParameter("inLinkValue", LinkValue, 30))
                oCmd.Parameters.Add(CConnection.StringParameter("inLinkField", LinkField, 50))
                oCmd.Parameters.Add(CConnection.IntegerParameter("AddressType", TypeId))
                If CConnection.ConnOpen Then
                    Dim intRet As Integer
                    intRet = oCmd.ExecuteNonQuery
                    blnRet = (intRet = 1)
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Setting Address link")
            Finally
                CConnection.SetConnClosed()
            End Try
            Return blnRet

        End Function

        Public Shared Function GetLinkedID(ByVal LinkValue As String, ByVal linkType As Integer, Optional ByVal LinkField As String = "Customer.Identity") As Integer
            Dim objReturn As Object
            Dim oCmd As New OleDb.OleDbCommand("Select lib_address.GetCustomerAddressId(?, ?, ?) from dual", MedConnection.Connection)
            Try
                If CConnection.ConnOpen Then
                    oCmd.Parameters.Add(CConnection.StringParameter("LinkValue", LinkValue, 30))
                    oCmd.Parameters.Add(CConnection.IntegerParameter("LinkType", linkType))
                    oCmd.Parameters.Add(CConnection.StringParameter("LinkField", LinkField, 50))
                    objReturn = oCmd.ExecuteScalar
                    If objReturn Is Nothing Or objReturn Is System.DBNull.Value Then
                        Return -1
                    Else
                        Return CInt(objReturn)
                    End If
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Getting addresslink")
            Finally
                CConnection.SetConnClosed()
            End Try


        End Function

        '''<summary>
        '''     sets up a link from the address to an alternate table such as the customer table 
        ''' </summary>
        ''' <param name='addressID'>Id of the address in the address table</param>
        ''' <param name='LinkField'>field in the alternate table that is linked in the form table.field</param>
        ''' <param name='LinkValue'>value of the primary key in that table</param>
        ''' <param name='TypeId'>Type of address link</param>
        ''' <returns> void</returns>
        ''' 
        Public Shared Function AddLink(ByVal addressID As Integer, _
            ByVal LinkField As String, _
            ByVal LinkValue As String, _
            ByVal TypeId As Integer) As Boolean

            Dim blnRet As Boolean = False
            If Not AddressCollection.HasLink(addressID, LinkField, LinkValue, TypeId) Then
                blnRet = AddressCollection.CreateAddresslink(addressID, LinkField, LinkValue, TypeId)
            Else
                blnRet = False
            End If
            Return blnRet

            '    oCmd.CommandText = "Select max(identity) + 1 from AddressLink"

            '    Dim intRet As Integer
            '    Dim intRes As Integer
            '    Try
            '        oCmd.Connection = MedscreenLib.medConnection.Connection      'Get Connection
            '        MedscreenLib.CConnection.SetConnOpen()                       'Open connection
            '        intRet = oCmd.ExecuteScalar                 'Get new id by fastest route possible

            '        oCmd.CommandText = "insert into AddressLink (" & _
            '          "IDENTITY,LinkField,LinkValue,ADDRESS_ID,Type,deleted) values(" & _
            '          intRet & ",'" & LinkField & "','" & LinkValue & "'," & _
            '          addressID & "," & TypeId & ",'F')"
            '        intRes = oCmd.ExecuteNonQuery


            '    Catch ex As Exception
            '        MedscreenLib.Medscreen.LogError(ex)
            ''MedscreenLib.medConnection.Connection = Nothing
            '    Finally
            '        MedscreenLib.CConnection.SetConnClosed()

            ''    End Try

        End Function

        '''<summary>
        ''' Produces a member by member copy of the address
        ''' </summary>
        ''' <param name='SourceID'>Id of the address that will be used to provide the copy</param>
        ''' <param name='Customer'>optional customer to create the address for</param>
        ''' <param name='adrLinkType'>optional type of link to create</param>
        ''' <returns> an address object</returns>
        Public Function Copy(ByVal SourceID As Integer, Optional ByVal Customer As String = "", Optional ByVal adrLinkType As Integer = 1) As Caddress

            Dim NewAddr As Caddress
            Dim SourceAddress As Caddress = Me.Item(SourceID, ItemType.AddressId)

            If SourceAddress Is Nothing Then
                Return Nothing
                Exit Function
            End If

            NewAddr = CreateAddress(SourceAddress.AdrLine1, SourceAddress.AddressType, Customer, adrLinkType)
            If Not NewAddr Is Nothing Then
                With SourceAddress
                    NewAddr.AdrLine2 = .AdrLine2
                    NewAddr.AdrLine3 = .AdrLine3
                    NewAddr.City = .City
                    NewAddr.Contact = .Contact
                    NewAddr.Country = .Country
                    NewAddr.CountryName = .CountryName
                    NewAddr.Email = .Email
                    NewAddr.Fax = .Fax
                    NewAddr.GoldmineID = .GoldmineID
                    NewAddr.Phone = .Phone
                    NewAddr.PostCode = .PostCode
                    NewAddr.Status = .Status
                    NewAddr.Update()
                    Me.Add(NewAddr)
                End With
            End If
            Return NewAddr

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get next available Address ID 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Function NextAddressId() As Integer
            Dim intRet As Integer
            Try
                intRet = MedscreenLib.CConnection.NextSequence("SEQ_ADDRESS")
            Catch ex As OleDb.OleDbException
                '    myTrans.Rollback()
                MedscreenLib.Medscreen.LogError(ex)
            Catch ex As Exception
                'myTrans.Rollback()
                MedscreenLib.Medscreen.LogError(ex)
            End Try
            Return intRet

        End Function

        '''<summary>
        ''' create a new address entity in the database
        ''' </summary>
        ''' <param name='LineOne'>First line of the address</param>
        ''' <param name='AdrType'>type of address see AddressType</param>
        ''' <param name='Customer'>optional customer to create the address for</param>
        ''' <param name='adrLinkType'>optional type of link to create</param>
        ''' <returns> an address object</returns>
        Public Function CreateAddress(Optional ByVal LineOne As String = "[new address]", _
            Optional ByVal AdrType As String = "CUST", _
            Optional ByVal Customer As String = "", _
        Optional ByVal adrLinkType As Integer = 1, Optional ByVal AddressId As Integer = -1) As Caddress

            'oCmd.CommandText = "Select LastVal from increments where major = 'SYNTAX' and minor = 'AddressID'"

            'Dim intRet As Integer
            'Dim intRes As Integer
            'Dim myTrans As OleDb.OleDbTransaction
            Dim intRet As Integer
            If AddressId = -1 Then
                intRet = NextAddressId()
            Else
                intRet = AddressId
            End If

            'Dim oInt As MedscreenLib.AddressInterface = New AddressInterface()
            'Dim intRet As Integer = oInt.CreateAddressID
            If intRet <> -1 Then
                Dim oAddr As New Caddress(intRet, Me)
                oAddr.Deleted = False
                oAddr.Update()
                oAddr.Fields.Loaded = True
                oAddr.AdrLine1 = LineOne                'Initialise address 
                oAddr.AddressType = AdrType             'Initialise address type
                oAddr.Country = 44                      'Initialise country to UK most common country
                oAddr.Update()                          'And save 
                Me.Add(oAddr)
                'oInt.Delete()

                If Customer.Trim.Length > 0 Then        ' Add link to customer
                    AddressCollection.AddLink(oAddr.AddressId, "Customer.Identity", Customer, adrLinkType)
                End If


                Return oAddr
            Else
                'oInt.Delete()
                Return Nothing
            End If
        End Function

        '''<summary>
        ''' attempt to get an address by its ID
        ''' </summary>
        ''' <param name='intID'>ID of the required address</param>
        ''' <returns> an address object or nothing if the address is non existant</returns>
        Public Function GetAddress(ByVal intID As Integer) As Caddress

            Dim tmpAddress As New Caddress(intID, Me)
            'Dim objTemp As Object

            Try
                tmpAddress.Refresh()
            Catch ex As Exception
            End Try
            Return tmpAddress

        End Function

        '''<summary>
        ''' constructor: create a new collection
        ''' </summary>
        Public Sub New()
            MyBase.New()
            'oConn.ConnectionString = support.ConnectString
        End Sub

        '''<summary>
        ''' convert data in the address table to XML
        ''' </summary>
        ''' <returns> XML representation of collection as a string </returns>
        Public Function ToXML() As String
            Dim strRet As String = "<Addresses xmlns=""http://tempuri.org/tempaddr.xsd"" > """

            Dim i As Integer
            Dim oAddr As Caddress

            If Me.Count = 0 Then Me.Load()

            For i = 0 To Count - 1
                oAddr = Me.Item(i, ItemType.Index)
                strRet += oAddr.ToXML
                Debug.WriteLine("Address " & i)
            Next

            strRet += "</Addresses>"
            Return strRet
        End Function
    End Class
#End Region
#End Region

#Region "DB Package"
    Public Class dbPackage

        Private Const C_LibAddress As String = "lib_Address"
        Private Shared _paramCustomer As New OleDbParameter("Customer", OleDbType.VarChar, 10)
        Private Shared _paramID As New OleDbParameter("Address_ID", OleDbType.Integer)
        Private Shared _paramFormat As New OleDbParameter("Address_ID", OleDbType.Integer)
        Private Shared _paramInvoice As New OleDbParameter("InvNum", OleDbType.VarChar, 20)
        Private Shared _paramDelimiter As New OleDbParameter("Delimiter", OleDbType.VarChar, 10)

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the address label for the specified address
        ''' </summary>
        ''' <param name="addressID" type="Int32"></param>
        ''' <param name="includeContact" type="Boolean"></param>
        ''' <returns>
        '''   A String containing formatted address, with lines separated by CRLF.
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetAddressLabel(ByVal addressID As Integer, ByVal includeContact As Boolean) As String
            Return GetAddressLabel(addressID, ControlChars.NewLine, includeContact)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the address label for the specified address
        ''' </summary>
        ''' <param name="addressID" type="Int32"></param>
        ''' <param name="delimiter" type="String"></param>
        ''' <param name="includeContact" type="Boolean"></param>
        ''' <returns>
        '''   A String containing formatted address, with lines separated by the specified delimiter.
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetAddressLabel(ByVal addressID As Integer, ByVal delimiter As String, ByVal includeContact As Boolean) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetAddressLabel", _
                                                New OleDbParameter() {_paramID, _paramDelimiter, _paramFormat}, _
                                                New Object() {addressID, delimiter, Math.Abs(CInt(includeContact))}))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the address label for the specified address
        ''' </summary>
        ''' <param name="addressID" type="Integer"></param>
        ''' <returns>
        '''   A String containing formatted address.
        ''' </returns>
        ''' <remarks>
        '''   Address is formatted including contact information, with lines separated
        '''   by CRLF.
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetAddressLabel(ByVal addressID As Integer) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetAddressLabel", _
                                                New OleDbParameter() {_paramID, _paramDelimiter}, _
                                                New Object() {addressID, ControlChars.NewLine}))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <returns>
        '''   String containing formatted address label with lines seperated by CRLF.
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceToLabel(ByVal invoiceNumber As String) As String
            Return GetInvoiceToLabel(invoiceNumber, ControlChars.NewLine)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <param name="delimiter" type="String">line separator</param>
        ''' <returns>
        '''   String containing formatted address label
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceToLabel(ByVal invoiceNumber As String, ByVal delimiter As String) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetInvoiceToLabel", _
                                                New OleDbParameter() {_paramInvoice, _paramDelimiter}, _
                                                New Object() {invoiceNumber, delimiter}))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice To address for the specified customer or vessel
        ''' </summary>
        ''' <param name="addressID" type="Int32"></param>
        ''' <param name="customerID" type="String">Customer or Vessel ID</param>
        ''' <param name="format" type="MedscreenLib.Address.LabelFormat"></param>
        ''' <returns>
        '''   String containing formatted address label
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceToLabel(ByVal addressID As Integer, ByVal customerID As String, ByVal format As LabelFormat) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetInvoiceToLabel", _
                                                New OleDbParameter() {_paramID, _paramCustomer, _paramFormat, _paramDelimiter}, _
                                                New Object() {addressID, customerID, CInt(format), ControlChars.NewLine}))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice Mail To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <returns>
        '''   String containing formatted address label, with lines separated by CRLF
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceMailToLabel(ByVal invoiceNumber As String) As String
            Return GetInvoiceMailToLabel(invoiceNumber, ControlChars.NewLine)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice Mail To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <param name="delimiter" type="String">line separator</param>
        ''' <returns>
        '''   String containing formatted address label
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceMailToLabel(ByVal invoiceNumber As String, ByVal delimiter As String) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetInvoiceMailToLabel", _
                                                New OleDbParameter() {_paramInvoice, _paramDelimiter}, _
                                                New Object() {invoiceNumber, delimiter}))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice Ship To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <returns>
        '''   String containing formatted address label, with lines separated by CRLF
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceShipToLabel(ByVal invoiceNumber As String) As String
            Return GetInvoiceShipToLabel(invoiceNumber, ControlChars.NewLine)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Formats the Invoice Ship To address for the specified invoice
        ''' </summary>
        ''' <param name="invoiceNumber" type="String"></param>
        ''' <param name="delimiter" type="String">line separator</param>
        ''' <returns>
        '''   String containing formatted address label
        ''' </returns>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetInvoiceShipToLabel(ByVal invoiceNumber As String, ByVal delimiter As String) As String
            Return ExecuteCommand(CreateCommand(C_LibAddress, "GetInvoiceShipToLabel", _
                                                New OleDbParameter() {_paramInvoice, _paramDelimiter}, _
                                                New Object() {invoiceNumber, delimiter}))
        End Function

        Private Shared Function CreateCommand(ByVal packageName As String, ByVal funcName As String, ByVal params() As OleDbParameter, ByVal values() As Object) As OleDbCommand
            Dim commandText As New System.Text.StringBuilder("SELECT " & packageName & "." & funcName & "(")
            Dim paramIndex As Integer, command As New OleDbCommand()
            For paramIndex = 0 To params.Length - 1
                commandText.Append("?,")
                command.Parameters.Add(params(paramIndex)).Value = values(paramIndex)
            Next
            commandText.Remove(commandText.Length - 1, 1)
            commandText.Append(") FROM Dual")
            command.CommandText = commandText.ToString
            command.Connection = MedConnection.Connection
            Return command
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Executes the command and clears it's parameters for next use
        ''' </summary>
        ''' <param name="cmd"></param>
        ''' <returns></returns>
        ''' <remarks>
        '''   Connection property is set to MedConnection.Connection if not already set
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Boughton]</Author><date> [22/11/2006]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Shared Function ExecuteCommand(ByVal cmd As OleDbCommand) As String
            Dim retVal As String = ""
            Try
                If cmd.Connection Is Nothing Then
                    cmd.Connection = MedConnection.Connection
                End If
                MedConnection.Open(cmd.Connection)
                retVal = cmd.ExecuteScalar.ToString
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                MedConnection.Close(cmd.Connection)
                cmd.Parameters.Clear()
            End Try
            Return retVal
        End Function

    End Class
#End Region

#Region "Country"

#Region "Country"

    '''<summary>
    ''' representation of a row in the country table within the Sample Manager Database
    ''' </summary>
    ''' <remarks>Countries are used extensively within the Sample Manager tables.<para/> 
    ''' They are particularly associated with ports and addresses<para/>
    ''' With Addresses they are used to determine the international dialling code for phone numbers
    ''' <para/>with ports they are used with volume discounts
    ''' </remarks>
    Public Class Country
#Region "Declaration"


        Private objCountryID As IntegerField = New IntegerField("COUNTRY_ID", 0, True)
        Private objCountryName As StringField = New StringField("COUNTRY_NAME", "", 30)
        Private objRegion As StringField = New StringField("REGION", "", 20)
        Private objCountryCode As StringField = New StringField("COUNTRY_CODE", "", 3)
        Private objCurrencyID As StringField = New StringField("CURRENCYID", "", 4)
        Private objContinetalArea As StringField = New StringField("CONTINENTAL_AREA", "", 20)

        Private myFields As TableFields = New TableFields("COUNTRY")
#End Region

        Private Sub SetupFields()
            myFields.Add(objCountryID)
            myFields.Add(objCountryName)
            myFields.Add(objRegion)
            myFields.Add(objCountryCode)
            myFields.Add(objCurrencyID)
            myFields.Add(Me.objContinetalArea)
        End Sub

        '''<summary>
        ''' Constructor
        ''' </summary>
        Public Sub New()
            SetupFields()
        End Sub

        '''<summary>
        ''' Returns the top level associated with the country
        ''' </summary>
        ''' <remarks>
        ''' This is used in conjunction with the port table to identify type D ports 
        ''' <para>or continents, howver these may be sub continents such as the far east or gulf
        ''' </para>
        ''' </remarks>
        Public Property ContinentalArea() As String
            Get
                Return Me.objContinetalArea.Value
            End Get
            Set(ByVal Value As String)
                Me.objContinetalArea.Value = Value
            End Set
        End Property

        '''<summary>
        ''' ID of the country in the table this is usually the International dialling code
        ''' </summary>
        Public Property CountryId() As Integer
            Get
                Return Me.objCountryID.Value
            End Get
            Set(ByVal Value As Integer)
                objCountryID.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Name of the country, this is the name as known from the UK
        ''' </summary>
        Public Property CountryName() As String
            Get
                Return Me.objCountryName.Value
            End Get
            Set(ByVal Value As String)
                objCountryName.Value = Value
            End Set
        End Property

        '''<summary>
        ''' IUPAC Country Code 
        ''' </summary>
        Public Property CountryCode() As String
            Get
                Return Me.objCountryCode.Value
            End Get
            Set(ByVal Value As String)
                objCountryCode.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Currency in use in the country using a 3 character code e.g. GBP
        ''' </summary>
        Public Property CurrencyCode() As String
            Get
                Return Me.objCurrencyID.Value
            End Get
            Set(ByVal Value As String)
                objCurrencyID.Value = Value
            End Set
        End Property

        '''<summary>
        ''' This may be used to indicate it is part of a particular grouping
        ''' </summary>
        Public Property Region() As String
            Get
                Return Me.objRegion.Value
            End Get
            Set(ByVal Value As String)
                objRegion.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Insert Country into database 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/11/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Insert() As Boolean
            If Not Me.myFields.Loaded Then
                Return Me.myFields.Insert(MedConnection.Connection)
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Update country information to database
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/11/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Update() As Boolean
            Return Me.myFields.Update(MedConnection.Connection)
        End Function

        '''<summary>
        ''' Fields in the row
        ''' </summary>
        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

        Public Function Load() As Boolean
            Dim ParamColl As New Collection()
            ParamColl.Add(CConnection.IntegerParameter("ID", Me.CountryId))
            Return myFields.Load(MedConnection.Connection, "Country_id = ?", ParamColl)
        End Function
    End Class
#End Region

#Region "CountryCollection"

    '''<summary>
    ''' Collection of Country objects
    ''' </summary>
    Public Class CountryCollection
        Inherits CollectionBase
#Region "Declarations"


        Private objCountryID As IntegerField = New IntegerField("COUNTRY_ID", 0, True)
        Private objCountryName As StringField = New StringField("COUNTRY_NAME", "", 30)
        Private objRegion As StringField = New StringField("REGION", "", 20)
        Private objCountryCode As StringField = New StringField("COUNTRY_CODE", "", 3)
        Private objCurrencyID As StringField = New StringField("CURRENCYID", "", 4)
        Private objContinetalArea As StringField = New StringField("CONTINENTAL_AREA", "", 20)


        Private myFields As TableFields = New TableFields("COUNTRY")
#End Region

        Public Enum indexBy
            Name
            Code
        End Enum

        Private Sub SetupFields()
            myFields.Add(objCountryID)
            myFields.Add(objCountryName)
            myFields.Add(objRegion)
            myFields.Add(objCountryCode)
            myFields.Add(objCurrencyID)
            myFields.Add(Me.objContinetalArea)
        End Sub

        '''<summary>
        ''' Index of a country object in the collection
        ''' </summary>
        ''' <param name='item'> Country to find index of </param>
        ''' <returns> Index of country </returns>
        Public Function IndexOf(ByVal item As Country) As Integer
            Return MyBase.List.IndexOf(item)
        End Function


        '''<summary>
        ''' Add a country to collection of countries
        ''' </summary>
        ''' <param name='Item'> Country To Add </param>
        ''' <returns>Index of added item</returns>
        Public Function Add(ByVal Item As Country) As Integer
            Me.List.Add(Item)
        End Function

        '''<summary>
        ''' Return a country Item from collection access by index
        ''' </summary>
        ''' <param name='index'> Index of item to get</param>
        Default Overloads Property Item(ByVal index As Integer) As Country
            Get
                Return CType(Me.List.Item(index), Country)
            End Get
            Set(ByVal Value As Country)
                Me.List.Item(index) = Value
            End Set
        End Property

        '''<summary>
        ''' Return a country Item from collection access by country Id
        ''' </summary>
        ''' <param name='index'> Country_id of item to get</param>
        Default Overloads Property Item(ByVal index As String, Optional ByVal indexon As indexBy = indexBy.Name) As Country
            Get
                Dim objCountry As Country = Nothing
                For Each objCountry In MyBase.List
                    If indexon = indexBy.Name Then
                        If objCountry.CountryName.ToUpper = index.ToUpper Then
                            Exit For
                        End If
                    ElseIf indexon = indexBy.Code Then
                        If objCountry.CountryCode.ToUpper = index.ToUpper Then
                            Exit For
                        End If
                    End If
                    objCountry = Nothing
                Next
                Return objCountry
            End Get
            Set(ByVal Value As Country)

            End Set
        End Property

        '''<summary>
        ''' Return a country Item from collection access by index, controlled by a second parameter
        ''' </summary>
        ''' <param name='Index'> Index of item to get <para> At the moment only 0 index by country id is available</para></param>
        ''' <param name='opt'> Option controlling what to get by</param>
        Default Overloads Property Item(ByVal Index As Integer, ByVal opt As Integer) As Country
            Get
                Dim i As Integer
                Dim oCountry As Country = Nothing
                For i = 0 To Me.Count - 1
                    oCountry = Item(i)
                    If opt = 0 Then
                        If oCountry.CountryId = Index Then
                            Exit For
                        End If
                    End If
                    oCountry = Nothing
                Next
                Return oCountry
            End Get
            Set(ByVal Value As Country)

            End Set
        End Property

        '''<summary>
        ''' Load the collection into memory
        ''' </summary>
        ''' <returns> True if load succesful</returns>
        Public Function Load() As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Try
                oCmd.Connection = MedscreenLib.MedConnection.Connection
                oCmd.CommandText = myFields.SelectString & " order by country_name"
                MedscreenLib.CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                Dim oCountry As Country
                oCountry = New Country()
                oCountry.CountryId = -1
                oCountry.CountryName = " -- "
                Me.Add(oCountry)
                While oRead.Read
                    oCountry = New Country()
                    oCountry.Fields.readfields(oRead)
                    Me.Add(oCountry)

                End While
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                MedscreenLib.CConnection.SetConnClosed()
            End Try
        End Function

        '''<summary>
        '''     Constructor
        ''' </summary>
        Public Sub New()
            MyBase.New()
            SetupFields()
        End Sub
    End Class
#End Region
#End Region

End Namespace