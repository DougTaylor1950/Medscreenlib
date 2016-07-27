Imports MedscreenLib
Namespace contacts

    ''' -----------------------------------------------------------------------------
    ''' Project	 : Intranet
    ''' Class	 : intranet.customerns.contacts.CustomerContact
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A contact for a customer
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class CustomerContact
        Inherits MedscreenLib.SMTable
#Region "Declarations"
        Private objContactId As IntegerField = New IntegerField("CONTACT_ID", -1, True)
        Private objCustomerId As StringField = New StringField("CUSTOMER_ID", "", 10)
        Private objAddressId As IntegerField = New IntegerField("ADDRESS_ID", -1)
        Private objContactType As StringPhraseField = New StringPhraseField("CONTACT_TYPE", "", 10, "CNTCT_TYP")
        Private objContactMethod As StringPhraseField = New StringPhraseField("CONTACT_METHOD", "", 10, "CNTCT_METH")
        Private objReportFormat As StringPhraseField = New StringPhraseField("REPORT_FORMAT", "", 10, "REP_FORMAT")
        Private objReportOptions As StringPhraseField = New StringPhraseField("REPORT_OPTIONS", "", 10, "REP_OPTION")
        Private objRemoved As BooleanField = New BooleanField("REMOVED", "F")
        Private objModifiedOn As TimeStampField = New TimeStampField("MODIFIED_ON")
        Private objModifiedBy As UserField = New UserField("MODIFIED_BY")
        '<Added code modified 05-Jul-2007 08:32 by LOGOS\Taylor> 
        Private objContactStatus As StringPhraseField = New StringPhraseField("CONTACT_STATUS", "", 10, "CONT_STAT")
        Private objOrderNumber As IntegerField = New IntegerField("ORDER_NUMBER", -1)
        '</Added code modified 05-Jul-2007 08:32 by LOGOS\Taylor> 
        '<Added code added 15-nov-2010>  Fields to be used by Secure email results
        Private objUserName As StringField = New StringField("USERNAME", "", 50)
        Private objPassword As StringField = New StringField("PASSWORD", "", 20)

        '</Added code added 15-nov-2010>

#Region "Enumerations"
        Public Enum StatusValues
            INACTIVE
            ACTIVE
            NONE
        End Enum

        Public Const GCST_STATUS_ACTIVE As String = "ACTIVE"
        Public Const GCST_STATUS_INACTIVE As String = "INACTIVE"


        Public Const GCST_PMContact = "PM"

#End Region

        Private myAddress As MedscreenLib.Address.Caddress
#End Region

#Region "Public Instance"
#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new customer contact
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new("CUSTOMER_CONTACTS")
            MyBase.Fields.Add(objContactId)
            MyBase.Fields.Add(objCustomerId)
            MyBase.Fields.Add(objAddressId)
            MyBase.Fields.Add(objContactType)
            MyBase.Fields.Add(objContactMethod)
            MyBase.Fields.Add(objReportFormat)
            MyBase.Fields.Add(objReportOptions)
            MyBase.Fields.Add(objRemoved)
            MyBase.Fields.Add(objModifiedOn)
            MyBase.Fields.Add(objModifiedBy)
            MyBase.Fields.Add(objContactStatus)
            MyBase.Fields.Add(objOrderNumber)
            MyBase.Fields.Add(objUserName)
            MyBase.Fields.Add(objPassword)
        End Sub
#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [15/11/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property UserName() As String
            Get
                Return Me.objUserName.Value
            End Get
            Set(ByVal Value As String)
                Me.objUserName.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [15/11/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Password() As String
            Get
                Return Me.objPassword.Value
            End Get
            Set(ByVal Value As String)
                Me.objPassword.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Priority of contact in the list
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' Each contact type has its own list
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	04/10/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property OrderNumber() As Integer
            Get
                Return Me.objOrderNumber.Value
            End Get
            Set(ByVal Value As Integer)
                objOrderNumber.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Id of contact 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ContactId() As Integer
            Get
                Return objContactId.Value
            End Get
            Set(ByVal Value As Integer)
                objContactId.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Work out who the contact is 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	23/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Contact() As String
            Get
                If Me.myAddress Is Nothing Then
                    If Me.AddressId > 0 Then
                        Me.myAddress = New MedscreenLib.Address.Caddress(Me.AddressId)
                    End If
                End If
                If Not Me.myAddress Is Nothing Then
                    Return Me.myAddress.Contact
                Else
                    Return ""
                End If
            End Get
            Set(ByVal Value As String)

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ID of customer
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CustomerId() As String
            Get
                Return objCustomerId.Value
            End Get
            Set(ByVal Value As String)
                objCustomerId.Value = Value
            End Set
        End Property

        Public Property Address() As MedscreenLib.Address.Caddress
            Get
                Return Me.myAddress
            End Get
            Set(ByVal Value As MedscreenLib.Address.Caddress)
                Me.myAddress = Value
                If Not Value Is Nothing Then
                    Me.AddressId = Value.AddressId
                End If
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ID of address
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property AddressId() As Integer
            Get
                Return objAddressId.Value
            End Get
            Set(ByVal Value As Integer)
                objAddressId.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' TYPE of contact linked to phrase CNTCT_TYP
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ContactType() As String
            Get
                Return objContactType.Value
            End Get
            Set(ByVal Value As String)
                objContactType.Value = Value
            End Set
        End Property

        Public ReadOnly Property ContactTypePhrase() As Glossary.Phrase
            Get
                Return objContactType.PhraseObject
            End Get
        End Property

        Public ReadOnly Property ContactTypeDescription() As String
            Get
                Return Me.objContactType.PhraseValue
            End Get

        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Method of contacting contact (phrase CNTCT_METH )
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ContactMethod() As String
            Get
                Return objContactMethod.Value
            End Get
            Set(ByVal Value As String)
                objContactMethod.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description of the contact method
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	05/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ContactMethodDescription() As String
            Get
                Return Me.objContactMethod.PhraseValue
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Status of the contact 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	05/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ContactStatus() As String
            Get
                If CStr(Me.objContactStatus.Value).Trim.Length = 0 Then
                    Return "ACTIVE"
                End If
                Return Me.objContactStatus.Value
            End Get
            Set(ByVal Value As String)
                objContactStatus.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Return the description of teh contact status 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	05/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ContactStatusDescription() As String
            Get
                Return Me.objContactStatus.PhraseValue
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Format of the report (phrase REP_FORMAT)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ReportFormat() As String
            Get
                Return objReportFormat.Value
            End Get
            Set(ByVal Value As String)
                objReportFormat.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Options associated with report (phrase REP_OPTION)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ReportOptions() As String
            Get
                Return objReportOptions.Value
            End Get
            Set(ByVal Value As String)
                objReportOptions.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Report Option description
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	28/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ReportOptionDescription() As String
            Get
                Return objReportOptions.PhraseValue
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Contact has been removed
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Removed() As Boolean
            Get
                Return objRemoved.Value
            End Get
            Set(ByVal Value As Boolean)
                objRemoved.Value = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Information last modified on 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ModifiedOn() As Date
            Get
                Return objModifiedOn.Value
            End Get
            Set(ByVal Value As Date)
                objModifiedOn.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Information last modified by 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ModifiedBy() As String
            Get
                Return objModifiedBy.Value
            End Get
            Set(ByVal Value As String)
                objModifiedBy.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose data fields
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Friend Property DataFields() As TableFields
            Get
                Return MyBase.Fields
            End Get
            Set(ByVal Value As TableFields)
                MyBase.Fields = Value
            End Set
        End Property

#End Region

#Region "Functions"

        ''' Update Contact info
        Public Function Update() As Boolean
            MyBase.Fields.Update(CConnection.DbConnection)
        End Function

        Public Function ToXML() As String
            Dim strRet As String = "<Contact><Details>" & CConnection.PackageStringList("Lib_contacts.ContactToXML", Me.ContactId)
            strRet += "</Details>"
            strRet += "<Address>"
            strRet += CConnection.PackageStringList("Lib_address.RowToXML", Me.AddressId)
            strRet += "</Address>"
            strRet += "<SMID>"
            strRet += CConnection.PackageStringList("Lib_Customer.GetSMIDProfile", Me.CustomerId)
            strRet += "</SMID>"
            strRet += "<SendLog>"
            Dim oColl As New Collection()
            oColl.Add(Me.ContactId)
            oColl.Add(Me.UserName)
            Dim strIds As String = CConnection.PackageStringList("lib_contacts.GetSendLogList", oColl)
            Dim strSendLog As String = ""
            If Not strIds Is Nothing AndAlso strIds.Trim.Length > 0 Then            'We have items in the log
                Dim ids As String() = strIds.Split(New Char() {","})
                Dim id As String
                For Each id In ids
                    If id.Trim.Length > 0 Then
                        Dim strLog = CConnection.PackageStringList("Lib_collection.GetSendLogRowXML", id)
                        If Not strLog Is Nothing Then strSendLog += strLog
                    End If
                Next
                strRet += strSendLog
            End If
            strRet += "</SendLog>"
            strRet += "</Contact>"
            Return strRet

        End Function
#End Region

#End Region
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : Intranet
    ''' Class	 : intranet.customerns.contacts.CustomerContactColl
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection of customer contacts
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class CustomerContactColl
        Inherits MedscreenLib.SMTableCollection
#Region "Declarations"
        Private objContactId As IntegerField = New IntegerField("CONTACT_ID", -1)
        Private objCustomerId As StringField = New StringField("CUSTOMER_ID", "", 10)
        Private objAddressId As IntegerField = New IntegerField("ADDRESS_ID", -1)
        Private objContactType As StringField = New StringField("CONTACT_TYPE", "", 10)
        Private objContactMethod As StringField = New StringField("CONTACT_METHOD", "", 10)
        Private objReportFormat As StringField = New StringField("REPORT_FORMAT", "", 10)
        Private objReportOptions As StringField = New StringField("REPORT_OPTIONS", "", 10)
        Private objRemoved As BooleanField = New BooleanField("REMOVED", "F")
        Private objModifiedOn As DateField = New DateField("MODIFIED_ON")
        Private objModifiedBy As StringField = New StringField("MODIFIED_BY", "", 10)
        '<Added code modified 05-Jul-2007 08:32 by LOGOS\Taylor> 
        Private objContactStatus As StringPhraseField = New StringPhraseField("CONTACT_STATUS", "", 10, "CONT_STAT")
        Private objOrderNumber As IntegerField = New IntegerField("ORDER_NUMBER", -1)

        '</Added code modified 05-Jul-2007 08:32 by LOGOS\Taylor> 
        '<Added code added 15-nov-2010>  Fields to be used by Secure email results
        Private objUserName As StringField = New StringField("USERNAME", "", 50)
        Private objPassword As StringField = New StringField("PASSWORD", "", 20)
        '</Added code added 15-nov-2010>
        Private myCustomer As String = ""
        Private myContactType As String = ""
#End Region

#Region "Public Instance"

#Region "Procedures"
        '''Create new collection 
        Public Sub New()
            MyBase.new("CUSTOMER_CONTACTS")
            MyBase.Fields.Add(objContactId)
            MyBase.Fields.Add(objCustomerId)
            MyBase.Fields.Add(objAddressId)
            MyBase.Fields.Add(objContactType)
            MyBase.Fields.Add(objContactMethod)
            MyBase.Fields.Add(objReportFormat)
            MyBase.Fields.Add(objReportOptions)
            MyBase.Fields.Add(objRemoved)
            MyBase.Fields.Add(objModifiedOn)
            MyBase.Fields.Add(objModifiedBy)
            MyBase.Fields.Add(objContactStatus)
            MyBase.Fields.Add(objOrderNumber)
            MyBase.Fields.Add(Me.objUserName)
            MyBase.Fields.Add(Me.objPassword)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection of contacts for a SMID-Profile
        ''' </summary>
        ''' <param name="CustomerID"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal CustomerID As String)
            MyClass.New()
            Me.myCustomer = CustomerID
        End Sub

        Public Sub New(ByVal customerId As String, ByVal ContactType As String)
            MyClass.New(customerId)
            Me.myContactType = ContactType
        End Sub
#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a contact by position 
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As Integer) As CustomerContact
            Get
                Return CType(MyBase.List.Item(index), CustomerContact)
            End Get
            Set(ByVal Value As CustomerContact)
                MyBase.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get the nth Contact of a type
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="Type"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	20/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As Integer, ByVal Type As String) As CustomerContact
            Get
                Dim intCount As Integer = index
                Dim objContact As CustomerContact
                For Each objContact In Me.List
                    If objContact.ContactType = Type Then
                        intCount -= 1
                        If intCount = 0 Then Exit For
                    End If
                    objContact = Nothing
                Next
                Return objContact
            End Get
            Set(ByVal Value As CustomerContact)

            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Customer ID
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	20/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CustomerId() As String
            Get
                Return Me.myCustomer
            End Get
            Set(ByVal Value As String)
                myCustomer = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a contact by contact ID
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As String) As CustomerContact
            Get
                Dim objContact As CustomerContact
                For Each objContact In Me.List
                    If objContact.ContactId = CInt(index) Then
                        Exit For
                    End If
                    objContact = Nothing
                Next
            End Get
            Set(ByVal Value As CustomerContact)
            End Set
        End Property

#End Region

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a new contact to the list 
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As CustomerContact) As Integer
            Return MyBase.List.Add(item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load a collection of contacts 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Me.List.Clear()
            Dim oRead As OleDb.OleDbDataReader
            Dim oCmd As New OleDb.OleDbCommand()
            Try ' Protecting 
                oCmd.Connection = CConnection.DbConnection
                'set up connect commandstring 
                Dim strSelect As String = MyBase.Fields.FullRowSelect()
                If Me.myCustomer.Trim.Length > 0 Then 'We are getting customer data 
                    strSelect += " where customer_id = ?"
                    oCmd.Parameters.Add(CConnection.StringParameter("CID", Me.myCustomer, 10))
                    'check to see if we have a contact type 
                    If Me.myContactType.Trim.Length > 0 Then
                        strSelect += " and contact_type = ? "
                        oCmd.Parameters.Add(CConnection.StringParameter("Ctype", Me.myContactType, 10))
                    End If
                    oCmd.CommandText += strSelect

                ElseIf Me.myContactType.Trim.Length > 0 Then
                    strSelect += " where contact_type = ? "
                    oCmd.Parameters.Add(CConnection.StringParameter("Ctype", Me.myContactType, 10))
                    oCmd.CommandText += strSelect

                End If




                If CConnection.ConnOpen Then            'Attempt to open reader
                    oRead = oCmd.ExecuteReader          'Get Date
                    While oRead.Read                    'Loop through data
                        Dim objContact As New CustomerContact()
                        objContact.DataFields.readfields(oRead)
                        Me.Add(objContact)
                    End While
                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "CustomerContactColl-Load-448")
            Finally
                CConnection.SetConnClosed()             'Close connection
                If Not oRead Is Nothing Then            'Try and close reader
                    If Not oRead.IsClosed Then oRead.Close()
                End If
            End Try

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new contact 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function CreateContact() As CustomerContact
            Dim objCustCont As New CustomerContact()
            Dim strContactId = CConnection.NextSequence("SEQ_CUSTOMER_CONTACTS")
            objCustCont.ContactId = strContactId
            objCustCont.DataFields.Insert(CConnection.DbConnection)
            Return objCustCont
        End Function


        Public Function CreateCustomerContact() As CustomerContact
            If Me.myCustomer.Trim.Length = 0 Then Exit Function
            Dim objCustCont As CustomerContact = Me.CreateContact
            objCustCont.CustomerId = Me.myCustomer
            Me.Add(objCustCont)
            objCustCont.Update()
            Return objCustCont

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get return string from database
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	02/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function ToXML() As String
            Dim strRet As String = CConnection.PackageStringList("lib_contacts.CustomerContactList", Me.CustomerId)
            Dim strReturn As String = ""
            If Not strRet Is Nothing AndAlso strRet.Trim.Length > 0 Then
                Dim strArray As String() = strRet.Split(New Char() {","})
                Dim i As Integer
                For i = 0 To strArray.Length - 1
                    Dim strId As String = strArray.GetValue(i)
                    If strId.Trim.Length > 0 Then
                        strReturn += CConnection.PackageStringList("lib_contacts.ContactToXML", strId)
                    End If
                Next
            End If
            Return strReturn
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose function to outside
        ''' </summary>
        ''' <param name="ClientId"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [30/06/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Shared Function ToXML(ByVal ClientId As String) As String
            Dim strRet As String = CConnection.PackageStringList("lib_contacts.CustomerContactList", ClientId)
            Dim strReturn As String = ""
            If Not strRet Is Nothing AndAlso strRet.Trim.Length > 0 Then
                Dim strArray As String() = strRet.Split(New Char() {","})
                Dim i As Integer
                If strArray.Length > 15 Then
                    strReturn += "<ROWSET><ROW>"
                    strReturn += "<CONTACT>Too many contacts to display in a timely fashion:- (" & strArray.Length & ") These contacts can be seen in the tree.</CONTACT>"
                    strReturn += "</ROW></ROWSET>"

                Else
                    For i = 0 To strArray.Length - 1
                        Dim strId As String = strArray.GetValue(i)
                        If strId.Trim.Length > 0 Then
                            strReturn += CConnection.PackageStringList("lib_contacts.ContactToXML", strId)
                        End If
                    Next
                End If

            End If
            Return strReturn
        End Function

        Public Overloads Shared Function ToXML(ByVal ClientId As String, ByVal _Type As String) As String
            Dim oColl As New Collection()
            oColl.Add(ClientId)
            oColl.Add(_Type)
            Dim strRet As String = CConnection.PackageStringList("lib_contacts.CustomerContactList", oColl)
            Dim strReturn As String = ""
            If Not strRet Is Nothing AndAlso strRet.Trim.Length > 0 Then
                Dim strArray As String() = strRet.Split(New Char() {","})
                Dim i As Integer
                For i = 0 To strArray.Length - 1
                    Dim strId As String = strArray.GetValue(i)
                    If strId.Trim.Length > 0 Then
                        strReturn += CConnection.PackageStringList("lib_contacts.ContactToXML", strId)
                    End If
                Next
            End If
            Return strReturn
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert contacts to XML
        ''' </summary>
        ''' <param name="CustId"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	22/12/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ToCustomerXML(ByVal CustId As String) As String
            Dim strRet As String = CConnection.PackageStringList("lib_contacts.CustomerContactList", CustId)
            Dim strReturn As String = ""
            If Not strRet Is Nothing AndAlso strRet.Trim.Length > 0 Then
                Dim strArray As String() = strRet.Split(New Char() {","})
                Dim i As Integer
                For i = 0 To strArray.Length - 1
                    Dim strId As String = strArray.GetValue(i)
                    If strId.Trim.Length > 0 Then
                        strReturn += CConnection.PackageStringList("lib_contacts.ContactToXML", strId)
                    End If
                Next
            End If
            Return strReturn

        End Function
#End Region
#End Region
    End Class
End Namespace