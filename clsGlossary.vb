'$Revision: 1.2 $
'$Author: taylor $
'$Date: 2005-09-02 17:04:22+01 $
'$Log: clsGlossary.vb,v $
'Revision 1.2  2005-09-02 17:04:22+01  taylor
'Added Exchange Rates
'
'Revision 1.1  2005-09-02 14:19:34+01  taylor
'Added myPhrase to class and Roles provided as shared properties
'
'Revision 1.0  2005-09-02 07:38:25+01  taylor
'Added to reporsitory
'
'
'<>
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Linq
Imports System.Windows.Forms
Imports MedscreenLib.Medscreen


Namespace Glossary

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Glossary.Glossary
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' an exposed class containg items in frequent use
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	26/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class Glossary

  Private Sub New()
    ' Prevent class being instantiated
  End Sub

  'Declarations of shared variables
#Region "Declarations"

#Region "General"

#Region "Personnel"

  Private Shared MyStaff As MedscreenLib.personnel.PersonnelList
  Private Shared mySMAddresses As MedscreenLib.Address.AddressCollection
  Private Shared myDepartments As System.Collections.Specialized.StringCollection
  Private Shared SMRoles As MedscreenLib.Glossary.SMRoles
  Private Shared SMRoleAssign As MedscreenLib.Glossary.RoleAssignments
#End Region

#Region "Financial"

  Private Shared myCurrencies As MedscreenLib.Glossary.Currencies
  Private Shared myExchangeRates As MedscreenLib.Glossary.ExchangeRates
  Private Shared myDiscountSchemes As MedscreenLib.Glossary.DiscountSchemes
  Private Shared myVolumeDiscountSchemes As MedscreenLib.Glossary.VolumeDiscountHeaders
  Private Shared objPrices As PriceList                         'Company Price List
#End Region

  Private Shared myCountries As MedscreenLib.Address.CountryCollection
  Private Shared objHolidays As MedscreenLib.Glossary.HolidayCollection
  Private Shared myLineItemGroups As MedscreenLib.Glossary.LineItemGroups
        'Private Shared oAddressType As MedscreenLib.Glossary.AddressTypes
  Private Shared objLineItems As MedscreenLib.Glossary.LineItems
  Private Shared ObjOptions As MedscreenLib.Glossary.OptionCollection
  Private Shared myMenus As MedscreenLib.CRMenuItems
      Private Shared mydefopt As MedscreenLib.Defaults.CustomerDefault
        'Private Shared Ogl As Groups
        Private Shared myPrinters As MedscreenLib.Glossary.PRINTERColl

        Public Enum LIMS
            SampleManager
            StarLIMS
            Labware
        End Enum

#End Region

#Region "Phrases"
  Private Shared ObjPhrases As MedscreenLib.Glossary.Phrases

  'Private Shared oAcc_Type As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oAccType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oConfMeth As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oPayInter As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oConfType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oCancStat As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oCollType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oServiceSold As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oTestType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oInvoicing As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oIndustry As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oPlatinum As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oFaxFormat As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oReportMethod As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oSwitch As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oVessType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oVessStatus As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oVessSubType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oCatType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oSendOption As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oDataAudit As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oInvVATType As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oCollStatus As MedscreenLib.Glossary.PhraseCollection
  'Private Shared oCollComments As MedscreenLib.Glossary.PhraseCollection
  Private Shared cancReasons As MedscreenLib.Glossary.PhraseCollection

#End Region

#End Region

#Region "Shared"



        'Public Shared Property OptionGroupList() As Groups
        '    Get
        '        If Ogl Is Nothing Then
        '            Ogl = New Groups()
        '            Ogl.ReadXml(Medscreen.LiveRoot & Medscreen.IniFiles & "CustOpt_Groups.xml")
        '        End If
        '        Return Ogl
        '    End Get
        '    Set(ByVal Value As Groups)
        '        Ogl = Value
        '    End Set
        'End Property

        '''' -----------------------------------------------------------------------------
        '''' <summary>
        '''' Groups of options 
        '''' </summary>
        '''' <param name="GroupName"></param>
        '''' <returns></returns>
        '''' <remarks>
        '''' </remarks>
        '''' <revisionHistory>
        '''' <revision><Author>[taylor]</Author><date> [06/06/2006]</date><Action></Action></revision>
        '''' </revisionHistory>
        '''' -----------------------------------------------------------------------------
        'Public Shared Function OptionGroups(ByVal GroupName As String) As Groups._OptionRow()
        '    Dim grrow As Groups.GROUPRow
        '    Dim myRet As Groups._OptionRow()
        '    For Each grrow In MedscreenLib.Glossary.Glossary.OptionGroupList.GROUP
        '        If grrow.Name.ToUpper = GroupName.ToUpper Then
        '            myRet = grrow.GetOptionRows
        '            Exit For
        '        End If
        '    Next
        '    Return myRet
        'End Function

        ''' -----------------------------------------------------------------------------
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a collection of printers defined in the printer table 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [16/01/2007]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Property Printers() As PRINTERColl
            Get
                If myPrinters Is Nothing Then           'If collection doesn't exist create it and load its contents
                    myPrinters = New PRINTERColl()
                    myPrinters.Load()
                End If
                Return myPrinters
            End Get
            Set(ByVal Value As PRINTERColl)
                myPrinters = Value
            End Set
        End Property



      Public Shared Property CustomerDefaults() As Defaults.CustomerDefault
          Get
              If mydefopt Is Nothing Then
                  mydefopt = New Defaults.CustomerDefault()
                  mydefopt.Load()


              End If
              Return mydefopt
          End Get
          Set(ByVal Value As Defaults.CustomerDefault)
              mydefopt = Value
          End Set
      End Property

      ''' <summary>
      '''      Implements access to a the country table within the Sample Manager Database
      ''' </summary>
      ''' <remarks>An instance to this object is only created when the first call to the property is made</remarks>
      ''' <seealso class='Addresses.Country' />
      Public Shared Property Countries() As MedscreenLib.Address.CountryCollection
          Get
              If myCountries Is Nothing Then
                  myCountries = New MedscreenLib.Address.CountryCollection()
                  myCountries.Load()
              End If
              Return myCountries
          End Get
          Set(ByVal Value As MedscreenLib.Address.CountryCollection)
              myCountries = Value
          End Set
      End Property

      ''' <summary>
      '''   implements access to the Address table within the Sample Manager Database
      '''   Returns a collection of Address objects
      ''' </summary>
      ''' <remarks>An instance to this object is only created when the first call to the property is made</remarks>
      Public Shared Property SMAddresses() As MedscreenLib.Address.AddressCollection
          Get
              If mySMAddresses Is Nothing Then
                  mySMAddresses = New MedscreenLib.Address.AddressCollection()
              End If
              Return mySMAddresses
          End Get
          Set(ByVal Value As MedscreenLib.Address.AddressCollection)
              mySMAddresses = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Holidays a collection of public holidays
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Holidays() As MedscreenLib.Glossary.HolidayCollection
          Get
              If objHolidays Is Nothing Then
                  objHolidays = New MedscreenLib.Glossary.HolidayCollection()
                  objHolidays.Load()
              End If
              Return objHolidays
          End Get
      End Property

   
#Region "Financial"
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Currencies supported
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property Currencies() As MedscreenLib.Glossary.Currencies
          Get

              If myCurrencies Is Nothing Then
                  myCurrencies = New MedscreenLib.Glossary.Currencies()
              End If
              Return myCurrencies
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.Currencies)
              myCurrencies = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Exchange rates for the currencies
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property ExchangeRates() As MedscreenLib.Glossary.ExchangeRates
          Get
              If myExchangeRates Is Nothing Then
                  myExchangeRates = New MedscreenLib.Glossary.ExchangeRates()

              End If
              Return myExchangeRates
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.ExchangeRates)
              myExchangeRates = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Groups of line items
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property LineItemGroups() As MedscreenLib.Glossary.LineItemGroups
          Get
              If myLineItemGroups Is Nothing Then
                  myLineItemGroups = New MedscreenLib.Glossary.LineItemGroups()
                  myLineItemGroups.Load()
              End If
              Return myLineItemGroups
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.LineItemGroups)
              myLineItemGroups = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Discount Schemes (simple)
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' Not volume discount schemes
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property DiscountSchemes() As MedscreenLib.Glossary.DiscountSchemes
          Get
              If myDiscountSchemes Is Nothing Then
                  myDiscountSchemes = New MedscreenLib.Glossary.DiscountSchemes()
                  myDiscountSchemes.Load()
              End If
              Return myDiscountSchemes
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.DiscountSchemes)
              myDiscountSchemes = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Volume discount schemes
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property VolumeDiscountSchemes() As MedscreenLib.Glossary.VolumeDiscountHeaders
          Get
              If myVolumeDiscountSchemes Is Nothing Then
                  myVolumeDiscountSchemes = New MedscreenLib.Glossary.VolumeDiscountHeaders()
                  myVolumeDiscountSchemes.Load()
              End If
              Return myVolumeDiscountSchemes
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.VolumeDiscountHeaders)
              myVolumeDiscountSchemes = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Current price list for all currencies
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property PriceList() As PriceList
          Get
              If objPrices Is Nothing Then
                  objPrices = New PriceList()
                  objPrices.Load()
              End If
              Return objPrices
          End Get
          Set(ByVal Value As PriceList)
              objPrices = Value
          End Set
      End Property


      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Line items 
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' A line item is any item that is in the price list, 
      ''' called a line item because it occupies a line in an invoice 
      ''' </remarks>
      ''' <history>
      ''' 	[taylor]	26/09/2005	Created
      ''' </history>
      ''' -----------------------------------------------------------------------------
      Public Shared Property LineItems() As MedscreenLib.Glossary.LineItems
          Get

              If objLineItems Is Nothing Then
                  objLineItems = New MedscreenLib.Glossary.LineItems()
                  objLineItems.Load()
              End If
              Return objLineItems
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.LineItems)
              objLineItems = Value
          End Set
      End Property

#End Region

#Region "Personnel"

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of personnel who can use Sample Manager
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property Personnel() As MedscreenLib.personnel.PersonnelList
          Get
              If MyStaff Is Nothing Then
                  MyStaff = New MedscreenLib.personnel.PersonnelList()
                  MyStaff.Load()
              End If
              Return MyStaff
          End Get
          Set(ByVal Value As MedscreenLib.personnel.PersonnelList)
              MyStaff = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' The current user using the application.
      ''' </summary>
      ''' <returns></returns>
      ''' <remarks>
      ''' This uses the current windows user to look up the personnel table for a match
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Function CurrentSMUser() As MedscreenLib.personnel.UserPerson

            Dim up As MedscreenLib.personnel.UserPerson
            Try
                up = Personnel.Item(System.Security.Principal.WindowsIdentity.GetCurrent.Name, _
                MedscreenLib.personnel.PersonnelList.FindIdBy.WindowsIdentity)
            Catch ex As Exception
                LogError(ex, , "Getting current user")
                Throw ex
            End Try
            Return up
        End Function



      Public Shared Function CurrentSMUserEmail() As String
          Dim up As MedscreenLib.personnel.UserPerson = CurrentSMUser()
          If Not up Is Nothing Then
              Return up.Email
          Else
              Return ""
          End If
      End Function
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of departments listed in the personnel table
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property Departments() As System.Collections.Specialized.StringCollection
          Get
              If myDepartments Is Nothing Then 'Not called so initialise variable and fill it 
                  myDepartments = New System.Collections.Specialized.StringCollection()
                  Dim oCmd As New OleDb.OleDbCommand()
                    Dim oRead As OleDb.OleDbDataReader = Nothing

                  Try
                      oCmd.Connection = MedscreenLib.medConnection.Connection
                      oCmd.CommandText = "Select distinct upper(department_id) from personnel"
                      If MedscreenLib.CConnection.ConnOpen Then
                          oRead = oCmd.ExecuteReader
                          While oRead.Read
              Dim strDepartment As String = oRead.GetString(0)
                              myDepartments.Add(strDepartment)
                          End While
                      End If
                  Catch ex As Exception
                  Finally
                      MedscreenLib.CConnection.SetConnClosed()
                      If Not oRead Is Nothing Then
                          If Not oRead.IsClosed Then oRead.Close()
                      End If
                  End Try
              End If
              Return myDepartments
          End Get
          Set(ByVal Value As System.Collections.Specialized.StringCollection)
              myDepartments = Value
          End Set
      End Property
#End Region

#Region "roles"
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A collection of Sample Manager Roles
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property SampleManagerRoles() As MedscreenLib.Glossary.SMRoles
          Get
              If SMRoles Is Nothing Then
                  SMRoles = New MedscreenLib.Glossary.SMRoles()
                  SMRoles.Load()
              End If
              Return SMRoles
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.SMRoles)
              SMRoles = Value
          End Set
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A collection of Sample Manager Role Assignments
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property SampleMangerAssignments() As MedscreenLib.Glossary.RoleAssignments
          Get
              If SMRoleAssign Is Nothing Then
                  SMRoleAssign = New MedscreenLib.Glossary.RoleAssignments()
                  SMRoleAssign.Load()
              End If
              Return SMRoleAssign
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.RoleAssignments)
              SMRoleAssign = Value
          End Set
      End Property
#Region "Functions "

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Information on the primary user in a role
      ''' </summary>
      ''' <param name="RoleName">Name of Role</param>
      ''' <param name="Info">What information to return</param>
      ''' <returns></returns>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Function UserInRole(ByVal RoleName As String, ByVal Info As String) As String

          Dim UP As MedscreenLib.personnel.UserPerson
            Dim strRet As String = ""
          If Not MedscreenLib.Glossary.Glossary.SampleMangerAssignments Is Nothing Then
              Dim strUser As String = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.TopRankedUser(RoleName)
              UP = Personnel.Item(strUser)
              If Not UP Is Nothing Then
                  Select Case Info
                      Case Constants.GCST_User_FullNetworkID
                            strRet = "CONCATENO\" & UP.WindowsID
                      Case Constants.GCST_User_NetworkID
                            strRet = UP.WindowsID
                      Case Constants.GCST_User_Identity
                            strRet = UP.Identity
                      Case Constants.GCST_User_Email
                            strRet = UP.Email
                      Case Constants.GCST_User_Manager
                            strRet = UP.Manager
                      Case Constants.GCST_User_Department
                            strRet = UP.Department
                      Case Constants.GCST_User_Location
                            strRet = UP.Location
                      Case Constants.GCST_User_Name
                            strRet = UP.Description
                      Case Else
                            strRet = UP.Identity
                  End Select
              End If
            End If
            Return strRet
      End Function

        Public Shared Function UserHasRole(ByVal RoleName As String, Optional ByVal Lims As LIMS = LIMS.SampleManager) As Boolean
            Dim UP As MedscreenLib.personnel.UserPerson = CurrentSMUser()
            Dim oColl As New Collection()

            If UP Is Nothing Then
                Return False
                Exit Function
            End If

            oColl.Add(RoleName)
            oColl.Add(UP.Identity)
            Dim strRet As String = CConnection.PackageStringList("Lib_utils.PositionInRole", oColl)
            If strRet IsNot Nothing AndAlso strRet.Trim.Length = 0 Then strRet = "0"
            Return (CInt(strRet) >= 0)

        End Function
#End Region

#End Region

      'Phrases these are various look up lists such as collection statuses
#Region "Phrases"

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A collection of possible phrases
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property PhraseList() As MedscreenLib.Glossary.Phrases
          Get
              If ObjPhrases Is Nothing Then
                  ObjPhrases = New MedscreenLib.Glossary.Phrases()
              End If
              Return ObjPhrases
          End Get
          Set(ByVal Value As MedscreenLib.Glossary.Phrases)
              ObjPhrases = Value
          End Set
      End Property

#Region "Customer"

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Types of Customer accounts 
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property CustomerAccountType() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("ACC_TYPE")
              'If oAcc_Type Is Nothing Then
              '  oAcc_Type = New MedscreenLib.Glossary.PhraseCollection("ACC_TYPE")
              '  oAcc_Type.Load()
              '  PhraseList.Add(oAcc_Type)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Exposes the phrase collection "Account Type"
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property AccountType() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("ACCTYPE")
              'If oAccType Is Nothing Then
              '  oAccType = New MedscreenLib.Glossary.PhraseCollection("ACCTYPE")
              '  oAccType.Load()
              '  PhraseList.Add(oAccType)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Industry Categories
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Industry() As MedscreenLib.Glossary.PhraseCollection
          Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("INDUSTRY")
              'If oIndustry Is Nothing Then
              '  oIndustry = New MedscreenLib.Glossary.PhraseCollection("INDUSTRY")
              '  oIndustry.Load()
              '  PhraseList.Add(oIndustry)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of Possible Report Methods
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property ReportMethod() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("REPORTMETH")
              'If oReportMethod Is Nothing Then
              '  oReportMethod = New MedscreenLib.Glossary.PhraseCollection("REPORTMETH")
              '  oReportMethod.Load()
              '  PhraseList.Add(oReportMethod)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Not used now
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Switch() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("SWITCH")
              'If oSwitch Is Nothing Then
              '  oSwitch = New MedscreenLib.Glossary.PhraseCollection("SWITCH")
              '  oSwitch.Load()
              '  PhraseList.Add(oSwitch)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of items where the data can be audited
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property DataAudits() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("DATAAUDITS")
              'If oDataAudit Is Nothing Then
              '  oDataAudit = New MedscreenLib.Glossary.PhraseCollection("DATAAUDITS")
              '  oDataAudit.Load()
              '  PhraseList.Add(oDataAudit)
              'End If
              Return myPhrase
          End Get

      End Property

#End Region

#Region "Vessels"
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Vessel Types
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property VesselTypeList() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("VESSTYPE")
              'If oVessType Is Nothing Then
              '  oVessType = New MedscreenLib.Glossary.PhraseCollection("VESSTYPE")
              '  oVessType.Load("phrase_text")
              '  PhraseList.Add(oVessType)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Vessel Statuses
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property VesselStatus() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("VESS_STAT")
              'If oVessStatus Is Nothing Then
              '  oVessStatus = New MedscreenLib.Glossary.PhraseCollection("VESS_STAT")
              '  oVessStatus.Load()
              '  PhraseList.Add(oVessStatus)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Vessel Sub Types
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property VesselSubTypeList() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("VESSUBTYPE")
              'If oVessSubType Is Nothing Then
              '  oVessSubType = New MedscreenLib.Glossary.PhraseCollection("VESSUBTYPE")
              '  oVessSubType.Load("Phrase_text")
              '  PhraseList.Add(oVessSubType)
              'End If
              Return myPhrase
          End Get

      End Property

#End Region ' Vessels

#Region "Invoicing"
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Payment intervals
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property PaymentIntervals() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("PAYINTER")
              'If oPayInter Is Nothing Then
              '  oPayInter = New MedscreenLib.Glossary.PhraseCollection("PAYINTER")
              '  oPayInter.Load()
              '  PhraseList.Add(oPayInter)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible Invoicing Methods
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Invoicing() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("INVOICING")
              'If oInvoicing Is Nothing Then
              '  oInvoicing = New MedscreenLib.Glossary.PhraseCollection("INVOICING")
              '  oInvoicing.Load()
              '  PhraseList.Add(oInvoicing)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible services that can be sold
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property ServicesSold() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("SERVICE")
              'If oServiceSold Is Nothing Then
              '  oServiceSold = New MedscreenLib.Glossary.PhraseCollection("SERVICE")
              '  oServiceSold.Load()
              '  PhraseList.Add(oServiceSold)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible states that the Platinum code can take
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Platinum() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("PLATINUM")
              'If oPlatinum Is Nothing Then
              '  oPlatinum = New MedscreenLib.Glossary.PhraseCollection("PLATINUM")
              '  oPlatinum.Load()
              '  PhraseList.Add(oPlatinum)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible VAT types that can be applied to a customer
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property InvoiceVATType() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("INVVATTYPE")
              'If oInvVATType Is Nothing Then
              '  oInvVATType = New MedscreenLib.Glossary.PhraseCollection("INVVATTYPE")
              '  oInvVATType.Load()
              '  PhraseList.Add(oInvVATType)
              'End If
              Return myPhrase
          End Get

      End Property
#End Region ' Invoicing

#Region "CustomerConfirmations"

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of Confirmation types
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property ConfTypeList() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("CONF_TYPE")
              'If oConfType Is Nothing Then
              '  oConfType = New MedscreenLib.Glossary.PhraseCollection("CONF_TYPE")
              '  oConfType.Load()
              '  PhraseList.Add(oConfType)
              'End If
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' A list of possible confirmation methods
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property ConfMethodList() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("CONF_METH")

              Return myPhrase
          End Get
      End Property
#End Region ' CustomerConfirmations

#Region "Programme mangement Phrases"
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Frequency that Programme Manager should contact customer
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [18/01/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property PMCustomerContactFrequency() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("PM_CALL")
              Return myPhrase
          End Get
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Type of Programme manager account 
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [18/01/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property PMAccountType() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("PM_TYPE")
              Return myPhrase
          End Get

      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Frequency of vessel testing 
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [18/01/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property PMVesselTestFrequency() As MedscreenLib.Glossary.PhraseCollection
          Get
              Dim myPhrase As PhraseCollection = PhraseList.Item("PM_CONT")
              Return myPhrase
          End Get
      End Property
#End Region

#Region "Sample Phrases"
        ''' <summary>
        '''     Status of Sample
        ''' </summary>
        ''' <value>
        '''     <para>
        '''         
        '''     </para>
        ''' </value>
        ''' <remarks>
        '''     
        ''' </remarks>
        Public Shared ReadOnly Property SampleMedscreenStatus() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("MEDSMPSTAT")
                Return myPhrase
            End Get

        End Property

        ''' <summary>
        '''     Overall result
        ''' </summary>
        ''' <value>
        '''     <para>
        '''         
        '''     </para>
        ''' </value>
        ''' <remarks>
        '''     
        ''' </remarks>
        Public Shared ReadOnly Property SampleOverallResult() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("OUTCOME")
                Return myPhrase
            End Get

        End Property

#End Region


#Region "Collections"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list of possible test types
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property TestTypeList() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("TEST_TYPE")
                'If oTestType Is Nothing Then
                '  oTestType = New MedscreenLib.Glossary.PhraseCollection("TEST_TYPE")
                '  oTestType.Load()
                '  PhraseList.Add(oTestType)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list of possible reasons for testing the donors
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property ReasonForTestList() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("OTHTST_RSN")
                'If oCollType Is Nothing Then
                '  oCollType = New MedscreenLib.Glossary.PhraseCollection("COLL_TYPE")
                '  oCollType.Load()
                '  PhraseList.Add(oCollType)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list of possible Cancellation Statuses
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property CancellationStatusList() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("CANCSTAT")
                'If oCancStat Is Nothing Then
                '  oCancStat = New MedscreenLib.Glossary.PhraseCollection("CANCSTAT")
                '  oCancStat.Load()
                '  PhraseList.Add(oCancStat)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Possible Fax Formats
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property FaxFormat() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("FAXFORMAT")
                'If oFaxFormat Is Nothing Then
                '  oFaxFormat = New MedscreenLib.Glossary.PhraseCollection("FAXFORMAT")
                '  oFaxFormat.Load()
                '  PhraseList.Add(oFaxFormat)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list of possible cancellation statuses
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property CollectionStatusList() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("COLL_STAT")
                'If oCollStatus Is Nothing Then
                '  oCollStatus = New MedscreenLib.Glossary.PhraseCollection("COLL_STAT")
                '  oCollStatus.Load()
                '  PhraseList.Add(oCollStatus)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list of possible Collections Comment types
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property CollectionCommentsList() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("COLLCOMTYP")
                'If oCollComments Is Nothing Then
                '  oCollComments = New MedscreenLib.Glossary.PhraseCollection("COLLCOMTYP")
                '  oCollComments.Load()
                '  PhraseList.Add(oCollComments)
                'End If
                Return myPhrase
            End Get

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Possible reasons for a cancellation
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Property CancellationReasons() As MedscreenLib.Glossary.PhraseCollection
            Get
                If cancReasons Is Nothing Then
                    cancReasons = New MedscreenLib.Glossary.PhraseCollection("CANC_COLL")
                    cancReasons.Load()
                End If
                Return cancReasons
            End Get
            Set(ByVal Value As MedscreenLib.Glossary.PhraseCollection)
                cancReasons = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Call out categories
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property CalloutCategory() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("CAT_TYPE")
                'If oCatType Is Nothing Then
                '  oCatType = New MedscreenLib.Glossary.PhraseCollection("CAT_TYPE")
                '  oCatType.Load()
                '  PhraseList.Add(oCatType)
                'End If
                Return myPhrase
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send options for collectors
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/10/2005]</date><Action>Added</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared ReadOnly Property SendOption() As MedscreenLib.Glossary.PhraseCollection
            Get
                Dim myPhrase As PhraseCollection = PhraseList.Item("SEND_OPT")
                'If oSendOption Is Nothing Then
                '  oSendOption = New MedscreenLib.Glossary.PhraseCollection("SEND_OPT")
                '  oSendOption.Load()
                '  PhraseList.Add(oSendOption)
                'End If
                Return myPhrase
            End Get
        End Property

#End Region ' Collections

#End Region ' Phrases


      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Options available 
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared ReadOnly Property Options() As MedscreenLib.Glossary.OptionCollection
          Get
              If ObjOptions Is Nothing Then
                  ObjOptions = New MedscreenLib.Glossary.OptionCollection()
                  ObjOptions.Load()
              End If
              Return ObjOptions
          End Get
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Menus available
      ''' </summary>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shared Property Menus() As MedscreenLib.CRMenuItems
          Get
              If myMenus Is Nothing Then
                  myMenus = New MedscreenLib.CRMenuItems()

                  'oConn.ConnectionString = ModTpanel.ConnectString(2)
                    myMenus.Load()

              End If
              Return myMenus
          End Get
          Set(ByVal Value As MedscreenLib.CRMenuItems)
              myMenus = Value
          End Set
      End Property

#End Region

End Class

'Role types 
#Region "Roles"
#Region "RoleHeader"

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Glossary.RoleHeader
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' This describes the possible values for a role
''' </summary>
''' <remarks>
''' Roles are a way of controlling what a person can do.  An individual can be in no roles, 
''' or many.
''' </remarks>
''' <history>
''' 	[taylor]	26/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class RoleHeader

#Region "declarations"


  Private objIdentity As New StringField("IDENTITY", "", 20, True)
  Private objGroupId As New StringField("GROUP_ID", "", 10)
  Private objDescription As New StringField("DESCRIPTION", "", 234)
  Private objModifiedOn As New DateField("MODIFIED_ON", DateField.ZeroDate)
  Private objModifiedBy As New StringField("MODIFIED_BY", "", 10)
    Private objModifiable As New BooleanField("MODIFIABLE", "T"c)
    Private objRemoveFlag As New BooleanField("REMOVEFLAG", "F"c)

  Private myFields As New TableFields("ROLE_HEADER")

#End Region

#Region "Public Instance"
#Region "Functions"
        Public Shared Function GetDescription(ByVal roleID As String) As String
            Dim strReturn As String = CConnection.PackageStringList("lib_utils.GetRoleDescription", roleID)
            If strReturn IsNot Nothing AndAlso strReturn.Trim.Length > 0 Then
                Return strReturn
            Else
                Return ""
            End If
        End Function
#End Region

#Region "Properties"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' The Id associated with this role
  ''' </summary>
  ''' <value>A string containging the RoleId</value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Property RoleID() As String
    Get
        Return objIdentity.Value.ToString
    End Get
    Set(ByVal Value As String)
      objIdentity.Value = Value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Description of the role
  ''' </summary>
  ''' <value>A string containing the description of the role</value>
  ''' <remarks>
  ''' The description is primarily for the GUI, it provides feedback to the users as to the intention of the role
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Property Description() As String
    Get
        Return Me.objDescription.Value.ToString
    End Get
    Set(ByVal Value As String)
      objDescription.Value = Value
    End Set
  End Property

#End Region

#Region "Procedures"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Constructor for a role header
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Sub New()
    SetupFields()
  End Sub
#End Region
#End Region

  Private Sub SetupFields()
    myFields.Add(objIdentity)
    myFields.Add(objGroupId)
    myFields.Add(objDescription)
    myFields.Add(objModifiedOn)
    myFields.Add(objModifiedBy)
    myFields.Add(objModifiable)
    myFields.Add(objRemoveFlag)
  End Sub


  Friend Property Fields() As TableFields
    Get
      Return myFields
    End Get
    Set(ByVal Value As TableFields)
      myFields = Value
    End Set
  End Property

End Class
#End Region

#Region "RoleHeaderCollection"
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Glossary.SMRoles
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of Role Headers
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	27/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class SMRoles
  Inherits CollectionBase

#Region "Declarations"


  Private objIdentity As New StringField("IDENTITY", "", 20, True)
  Private objGroupId As New StringField("GROUP_ID", "", 10)
  Private objDescription As New StringField("DESCRIPTION", "", 234)
  Private objModifiedOn As New DateField("MODIFIED_ON", DateField.ZeroDate)
  Private objModifiedBy As New StringField("MODIFIED_BY", "", 10)
    Private objModifiable As New BooleanField("MODIFIABLE", "T"c)
    Private objRemoveFlag As New BooleanField("REMOVEFLAG", "F"c)

  Private myFields As New TableFields("ROLE_HEADER")

  Private Sub SetupFields()
    myFields.Add(objIdentity)
    myFields.Add(objGroupId)
    myFields.Add(objDescription)
    myFields.Add(objModifiedOn)
    myFields.Add(objModifiedBy)
    myFields.Add(objModifiable)
    myFields.Add(objRemoveFlag)
  End Sub
#End Region

#Region "Public Instance"

#Region "Functions"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Add as Role to the collection
  ''' </summary>
  ''' <param name="Value">The item being added</param>
  ''' <returns>The index of the item added</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Function Add(ByVal Value As RoleHeader) As Integer
    Return MyBase.List.Add(Value)
  End Function

#End Region

#Region "Procedures"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Constructor for the collection
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Sub New()
    MyBase.New()
    SetupFields()
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Load the collection from the database
  ''' </summary>
  ''' <returns>TRUE if sucessful</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Function Load() As Boolean
    Dim oCmd As New OleDb.OleDbCommand()
    Dim blnReturn As Boolean = False
            Dim oRead As OleDb.OleDbDataReader = Nothing
    Dim objRole As RoleHeader

    Try
              oCmd.CommandText = myFields.FullRowSelect
      oCmd.Connection = MedscreenLib.medConnection.Connection
      CConnection.SetConnOpen()
      oRead = oCmd.ExecuteReader
      objRole = New RoleHeader()
      While oRead.Read
        objRole.Fields.ReadFields(oRead)
        Me.Add(objRole)
      End While
      blnReturn = True
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      If Not oRead Is Nothing Then
        If Not oRead.IsClosed Then oRead.Close()
              End If
              CConnection.SetConnClosed()
              oRead = Nothing
              oCmd = Nothing
    End Try
  End Function

#End Region

#Region "Properties"


  '''
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Retrieves an Item from the collection by its position
  ''' </summary>
  ''' <param name="index">Position of the item to get</param>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Default Public Overloads Property Item(ByVal index As Integer) As RoleHeader
    Get
      Return CType(MyBase.List.Item(index), RoleHeader)
    End Get
    Set(ByVal Value As RoleHeader)
      MyBase.List.Item(index) = Value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Retrieves an Item from the collection by its Role ID
  ''' </summary>
  ''' <param name="index">Role Id for the item to get</param>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Default Public Overloads Property Item(ByVal index As String) As RoleHeader
    Get
                Dim oRole As RoleHeader = Nothing
      Dim i As Integer
      For i = 0 To Me.Count - 1
        oRole = Item(i)
        If oRole.RoleID = index Then
          Exit For
        End If
        oRole = Nothing
      Next
      Return oRole
    End Get
    Set(ByVal Value As RoleHeader)

    End Set
  End Property


#End Region

#End Region

End Class

#End Region

#Region "RoleAssignment"
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Glossary.RoleAssignment
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' An assignment for a role
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	27/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class RoleAssignment

#Region "Declarations"

  Private objOperatorId As New StringField("OPERATOR_ID", "", 10)
  Private objRoleId As New StringField("ROLE_ID", "", 20)
  Private objGroupId As New StringField("GROUP_ID", "", 10)
  Private objRank As New IntegerField("RANK", 0)

  Private myFields As New TableFields("ROLE_ASSIGNMENT")
#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Constructor for the role assignment
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Sub New()
    SetupFields()
  End Sub

#End Region

#Region "Properties"
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' The Role ID for this role assignment
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Property Role() As String
    Get
        Return Me.objRoleId.Value.ToString
    End Get
    Set(ByVal Value As String)
      objRoleId.Value = Value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' The User who has this role, is the Primary key of the Personnel table
  ''' </summary>
  ''' <value>Personnel ID</value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Property User() As String
    Get
        Return Me.objOperatorId.Value.ToString
    End Get
    Set(ByVal Value As String)
      objOperatorId.Value = Value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' The rank this role holder has
  ''' </summary>
  ''' <value>The Rank as Integer</value>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Property Rank() As Integer
    Get
        Return CInt(Me.objRank.Value)
    End Get
    Set(ByVal Value As Integer)
      objRank.Value = Value
    End Set
  End Property
#End Region

#End Region
  Private Sub SetupFields()
    myFields.Add(objOperatorId)
    myFields.Add(objRoleId)
    myFields.Add(objGroupId)
    myFields.Add(objRank)
  End Sub

  Friend Property Fields() As TableFields
    Get
      Return myFields
    End Get
    Set(ByVal Value As TableFields)
      myFields = Value
    End Set
  End Property


End Class

#End Region

#Region "RoleAssignmentCollection"
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Glossary.RoleAssignments
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of Role Assignments
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	27/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class RoleAssignments
  Inherits CollectionBase

#Region "Shared"
    Private Shared _assignments As RoleAssignments

    Shared Sub New()
      _assignments = New RoleAssignments()
    End Sub

    Public Shared Function Current() As RoleAssignments
      If _assignments.Count = 0 Then
        _assignments.Load()
      End If
      Return _assignments
    End Function

#End Region

#Region "Declarations"

    Private objOperatorId As New StringField("OPERATOR_ID", "", 10)
    Private objRoleId As New StringField("ROLE_ID", "", 20)
    Private objGroupId As New StringField("GROUP_ID", "", 10)
    Private objRank As New IntegerField("RANK", 0)

    Private myFields As New TableFields("ROLE_ASSIGNMENT")
#End Region


#Region "Public Instance"

#Region "Functions"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a RoleAssignment to the collection
    ''' </summary>
    ''' <param name="value">The Role to add</param>
    ''' <returns>Position of Added Item</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal value As RoleAssignment) As Integer
      Return MyBase.List.Add(value)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load all Assignments form the database
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()
      Dim blnReturn As Boolean = False
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.CommandText = myFields.FullRowSelect
        oCmd.Connection = medConnection.Connection
        CConnection.SetConnOpen()
        oRead = oCmd.ExecuteReader
        While oRead.Read
          Dim objRoleAssign As New RoleAssignment()
          objRoleAssign.Fields.readfields(oRead)
          Me.Add(objRoleAssign)
        End While
        blnReturn = True
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        oRead = Nothing
        oCmd = Nothing
        CConnection.SetConnClosed()
      End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of Roles that the user has 
    ''' </summary>
    ''' <param name="User">User ID</param>
    ''' <returns>An Array of Role IDs</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UserRoles(ByVal User As String) As Collections.ArrayList
      Dim tmpArr As New ArrayList()
      Dim i As Integer
      Dim objRA As RoleAssignment

      For i = 0 To Me.Count - 1
        objRA = Me.Item(i)
        If objRA.User.ToUpper = User.ToUpper Then
          tmpArr.Add(objRA.Role)
        End If
      Next
      Return tmpArr
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Identify the highest ranked user in a role
    ''' </summary>
    ''' <param name="Role">Role to find highest ranked user</param>
    ''' <returns>User Id</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function TopRankedUser(ByVal Role As String) As String
      Dim i As Integer
            Dim objRA As RoleAssignment = Nothing

      For i = 0 To Me.Count - 1
        objRA = Me.Item(i)
        If objRA.Role.ToUpper = Role.ToUpper Then
          If objRA.Rank = 1 Then
            Exit For
          End If
        End If
        objRA = Nothing
      Next
      If Not objRA Is Nothing Then
        Return objRA.User
      Else
        Return ""
      End If

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns whether a user is in a particular Sample Manager role
    ''' </summary>
    ''' <param name="user">User (Personnel ID)</param>
    ''' <param name="role">Role (Role ID)</param>
    ''' <returns>TRUE if user is in Role</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UserHasRole(ByVal user As String, ByVal role As String) As Boolean
      Dim i As Integer
            Dim objRA As RoleAssignment = Nothing

      Me.List.Clear()                 'Ensure the list is refreshed
      Me.Load()                       'Re load list 
      For i = 0 To Me.Count - 1
        objRA = Me.Item(i)
        If objRA.Role.ToUpper = role.ToUpper Then
          If objRA.User.ToUpper = user.ToUpper Then
            Exit For
          End If
        End If
        objRA = Nothing
      Next
      If Not objRA Is Nothing Then
        Return True
      Else
        Return False
      End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holders of a particular role (As Array)
    ''' </summary>
    ''' <param name="Role">Role that list is to be returned for</param>
    ''' <returns>An array of Role Holders</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RoleHolders(ByVal Role As String) As RoleAssignments
      Dim tmpArr As New RoleAssignments()
      Dim i As Integer
      Dim objRA As RoleAssignment

      tmpArr.Add(New RoleAssignment())
      For i = 0 To Me.Count - 1
        objRA = Me.Item(i)
        If objRA.Role.ToUpper = Role.ToUpper Then
          tmpArr.Add(objRA)
        End If
      Next
      Return tmpArr
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Users who have a particular role
    ''' </summary>
    ''' <param name="Role">Role to get list for </param>
    ''' <returns>A collection of Role Assignments</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UserRoleHolders(ByVal Role As String) As RoleAssignments
      Dim tmpArr As New RoleAssignments()
      Dim i As Integer
      Dim objRA As RoleAssignment

      For i = 0 To Me.Count - 1
        objRA = Me.Item(i)
        If objRA.Role.ToUpper = Role.ToUpper Then
          tmpArr.Add(objRA)
        End If
      Next
      Return tmpArr
    End Function


#End Region

#Region "PRocedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor Create a role Assignment collection
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New()
      SetupFields()
    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Access a particular Item in the collection
    ''' </summary>
    ''' <param name="Index">Position of Item in collection</param>
    ''' <value>The Rolleassignment at this position</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Default Public Overloads Property Item(ByVal Index As Integer) As RoleAssignment
      Get
        Return CType(MyBase.List(Index), RoleAssignment)
      End Get
      Set(ByVal Value As RoleAssignment)
        MyBase.List(Index) = Value
      End Set
    End Property


#End Region

#End Region
    Private Sub SetupFields()
      myFields.Add(objOperatorId)
      myFields.Add(objRoleId)
      myFields.Add(objGroupId)
      myFields.Add(objRank)
    End Sub

End Class

#End Region
#End Region

'Phrases
#Region "Phrases"
#Region "Phrase"

'''<summary>
''' phrase types based on the SM Phrase table
''' </summary>
    Public Class Phrase
        'Implements IPhraseEntry
        Implements IComparer


#Region "declarations"

        Private myFields As TableFields
        Private objPhraseType As StringField
        Private objPhraseId As StringField
        Private objPhraseNum As StringField
        Private objPhraseText As StringField
        Private objStatus As StringField = New StringField("STATUS", "", 10)
        Private objGenText1 As StringField = New StringField("GENTEXT1", "", 2000)
        Private objGenText2 As StringField = New StringField("GENTEXT2", "", 80)
        Private objIcon As StringField = New StringField("ICON", "", 30)
        Private myFormatter As PhraseSupport

        Private Sub setupFields()
            myFields = New TableFields("PHRASE")

            objPhraseType = New StringField("PHRASE_TYPE", "", 10)
            myFields.Add(objPhraseType)
            objPhraseId = New StringField("PHRASE_ID", "", 10)
            myFields.Add(objPhraseId)
            objPhraseNum = New StringField("ORDER_NUM", "", 10)
            myFields.Add(objPhraseNum)
            objPhraseText = New StringField("PHRASE_TEXT", "", 234)
            myFields.Add(objPhraseText)
            'Lines below need uncommenting when the fields are added to the schema
            myFields.Add(objStatus)
            myFields.Add(objGenText1)
            myFields.Add(objGenText2)
            myFields.Add(objIcon)
        End Sub
#End Region

#Region "Public Instance"

#Region "Shared Functions"
        Public Shared Function GetPhraseText(ByVal phraseType As String, ByVal phraseID As String) As String
            Dim phraseText As String = ""
            Dim list As PhraseCollection = Glossary.PhraseList(phraseType)
            If Not list Is Nothing Then
                Dim phrase As Phrase = list.Item(phraseID)
                If Not phrase Is Nothing Then
                    phraseText = phrase.PhraseText
                End If
            End If
            Return phraseText
        End Function

        Public Overloads Shared Function Equals(ByVal objA As Object, ByVal objB As Object) As Boolean
            Return TypeOf objA Is Phrase AndAlso _
                   TypeOf objB Is Phrase AndAlso _
                   DirectCast(objA, Phrase).PhraseType = DirectCast(objB, Phrase).PhraseType AndAlso _
                   DirectCast(objA, Phrase).PhraseID = DirectCast(objB, Phrase).PhraseID
        End Function

#End Region

#Region "Functions"

        Public Overloads Function Equals(ByVal obj As Object) As Boolean
            Return Phrase.Equals(Me, obj)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Generates an XML representation of the phrase
        ''' </summary>
        ''' <returns>The phrase represented as XML</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function ToXml() As String
            Return myFields.ToXMLSchema()
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Returns the phrase text
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        '''   The phrase text is returned for use in combo boxes, where the phrase itself
        '''   can be added to the box, with the phrase text displayed for the user.
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overrides Function ToString() As String
            Return Me.PhraseText
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Indicates whether value matches PhraseID
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Boughton]</Author><date> [07/12/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function Equals(ByVal value As String) As Boolean
            Return value.Equals(Me.PhraseID)
        End Function


        Public Function Load() As Boolean
            Dim oColl As New Collection()
            oColl.Add(CConnection.StringParameter("PhraseType", Me.PhraseType, 10))
            oColl.Add(CConnection.StringParameter("PhraseID", Me.PhraseID, 10))
            Return myFields.Load(CConnection.DbConnection, "Phrase_Type = ? and Phrase_id = ?", oColl)
        End Function
#End Region

#Region "Procedures"
        ''' <summary>
        '''		Implements the constructor for the Phrase object.
        '''		calls private member SetupFields which populates the fields object.
        ''' </summary>
        Public Sub New()
            setupFields()
        End Sub

        ''' <summary>
        '''		Implements the constructor for the Phrase object.
        '''     overloaded to with a phrase type
        '''		calls private member SetupFields which populates the fields object.
        '''     then initialises the PhraseType member
        ''' </summary>
        Public Sub New(ByVal PhraseType As String)
            setupFields()
            Me.objPhraseType.Value = PhraseType
        End Sub

        Public Sub New(ByVal PhraseType As String, ByVal PhraseID As String)
            MyClass.New(PhraseType)
            Me.objPhraseId.Value = PhraseID
            Me.Load()
        End Sub
#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The phrase list this phrase belongs to 
        ''' </summary>
        ''' <value>Phrase list Id as string</value>
        ''' <remarks>
        ''' Phrases are organised in to lists of myPhrase, each list is composed of a number of myPhrase
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property PhraseType() As String
            Get
                Return objPhraseType.Value.ToString
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Text associated with this phrase
        ''' </summary>
        ''' <value>Returns the phrase dceriptive text</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property PhraseText() As String
            Get
                Return objPhraseText.Value.ToString
            End Get
            Set(ByVal Value As String)
                objPhraseText.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Phrase ID (10 character code the identifies this phrase)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property PhraseID() As String
            Get
                Return objPhraseId.Value.ToString
            End Get
            Set(ByVal Value As String)
                objPhraseId.Value = Value
            End Set
        End Property



        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Status of phrase
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [14/06/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property PhraseStatus() As String
            Get
                Return Me.objStatus.Value
            End Get
            Set(ByVal Value As String)
                Me.objStatus.Value = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Non specific text
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [14/06/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property GenText1() As String
            Get
                Return Me.objGenText1.Value
            End Get
            Set(ByVal Value As String)
                Me.objGenText1.Value = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Non specific text 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [14/06/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property GenText2() As String
            Get
                Return Me.objGenText2.Value
            End Get
            Set(ByVal Value As String)
                Me.objGenText2.Value = Value
            End Set
        End Property

        Public Property Icon() As String
            Get
                Return objIcon.Value
            End Get
            Set(ByVal Value As String)
                objIcon.Value = Value
            End Set
        End Property

        Protected Friend Sub SetPhraseID(ByVal value As String)
            objPhraseId.Value = value
        End Sub

        Protected Friend Sub SetPhraseText(ByVal value As String)
            objPhraseText.Value = value
        End Sub


        Public ReadOnly Property Formatter() As PhraseSupport
            Get
                'need to add checks that gentext1 is a valid xml phrase file
                If Me.myFormatter Is Nothing AndAlso InStr(GenText1, "<PhraseSupport") > 0 Then
                    Me.myFormatter = New PhraseSupport()
                    myFormatter.DeSerialiseFromXML(GenText1)
                ElseIf myFormatter Is Nothing Then
                    Me.myFormatter = New PhraseSupport()
                End If
                Return myFormatter
            End Get

        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Position in this phrase list
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property OrderNum() As String
            Get
                Return objPhraseNum.Value.ToString
            End Get
            Set(ByVal Value As String)
                objPhraseNum.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Exposes the database fields 
        ''' </summary>
        ''' <value>A collection of Fields</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

#End Region

#End Region

#Region "Private instance"

#Region "Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Look for search text in string, format will be SearchText=xxxxx;
        ''' </summary>
        ''' <param name="SearchText"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [15/06/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Function ParseText(ByVal SearchText As String) As String
            Dim AllText As String = GenText1 & GenText2
            Dim RetText As String = ""
            'Initialise searchText 
            SearchText = SearchText + "="
            Dim Ipos As Integer = InStr(AllText.ToUpper, SearchText.ToUpper)
            If Ipos > 0 Then 'we have found searchtext
                Dim tmpText As String = Mid(AllText, Ipos + SearchText.Length)
                Dim i As Integer = 0
                While i < tmpText.Length AndAlso tmpText.Chars(i) <> ";"
                    RetText += tmpText.Chars(i)
                    i += 1
                End While
            End If
            Return RetText
        End Function
#End Region
#End Region

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            If CType(x, Phrase).PhraseText.ToUpper > CType(y, Phrase).PhraseText.ToUpper Then
                Return 1
            ElseIf CType(x, Phrase).PhraseText.ToUpper = CType(y, Phrase).PhraseText.ToUpper Then
                Return 0
            Else
                Return -1
            End If
        End Function
    End Class
#End Region

#Region "PhraseCollection"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.PhraseCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of myPhrase that form a phrase list
    ''' </summary>
    ''' <remarks>
    ''' Phrases are organised in to phrase lists, all the phrase lists are stored together in the phrase table.
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class PhraseCollection
        Inherits System.Collections.Generic.List(Of Phrase)

#Region "Declarations"


        Private myFields As TableFields
        Private objPhraseType As StringField
        Private objPhraseId As StringField
        Private objPhraseNum As StringField
        Private objPhraseText As StringField
        Private objStatus As StringField = New StringField("STATUS", "", 10)
        Private objGenText1 As StringField = New StringField("GENTEXT1", "", 80)
        Private objGenText2 As StringField = New StringField("GENTEXT2", "", 80)
        Private objIcon As StringField = New StringField("ICON", "", 30)

        Private strPhraseType As String

#End Region

#Region "Enumerations"
        ''' <summary>
        '''     Enumeration to describe type of constructor
        ''' </summary>
        ''' <remarks>
        '''     
        ''' </remarks>
        Public Enum BuildBy
            ''' <summary>
            '''    Provide a table name, table must be in structure.txt document 
            ''' </summary>
            Table
            ''' <summary>
            '''     Use a PL/SQL function that returns a comma separated string of the form ID|Description,ID|Description
            ''' </summary>
            PLSQLFunction
            ''' <summary>
            '''     Use a SQL Query
            ''' </summary>

            Query

        End Enum

        Public Enum IndexBy
            PhraseId
            PhraseText
            OrderBY
        End Enum
#End Region 'Enumerations

#Region "Public Instance"

#Region "Functions"

        Public Function SortPhrase() As Boolean
            MyBase.Sort(New PhraseComparer)
        End Function

        Public Overloads Sub sort()
            SortPhrase()
        End Sub

        Public Function FindByGenText2(ByVal TextValue As String) As Phrase
            Dim pred As New PhraseFindClass(TextValue)
            Dim Ph As Phrase = Find(AddressOf pred.FindPhraseByGentext2)
            Return Ph
        End Function

      
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert phrase list to an untyped dataset
    ''' </summary>
    ''' <returns>untyped data set</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ToDataSet() As DataSet
      Dim tempStr As String = ToXml()
      Dim st As New IO.MemoryStream(CType(Medscreen.StringToByteArray(tempStr), Byte()))

      Dim Ds As New DataSet()

      Ds.ReadXml(st, XmlReadMode.InferSchema)

      st.Close()
      st = Nothing
      Return Ds
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clone a phrase collection so two can be used at once 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [25/10/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Clone() As PhraseCollection
      Dim CollRet As PhraseCollection = New PhraseCollection(Me.PhraseType)
      Me.Load()
      Return CollRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert phrase list to XML 
    ''' </summary>
    ''' <returns>Phrase list represented as XML in a string </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ToXml() As String
      Dim i As Integer
      Dim objItem As Phrase

            Dim strRet As String = "<" & strPhraseType & ">"
      For i = 0 To Me.Count - 1
        objItem = Me.Item(i)
        strRet += objItem.ToXml
      Next
      strRet += "</" & strPhraseType & ">"
      Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Write Phrase list out as XML to a file
    ''' </summary>
    ''' <param name="Filename"></param>
        ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub WriteXmlToFile(ByVal Filename As String)
      Try
        Dim w As IO.StreamWriter = New IO.StreamWriter(Filename, False)
        w.Write(Me.ToXml)
        w.Flush()
        w.Close()

      Catch ex As Exception
      End Try

      'Me.myFields.XmlToFile(Filename, True)
    End Sub

    
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find the phrase in the list 
    ''' </summary>
    ''' <param name="Phrase">Phrase entity to find</param>
    ''' <returns>Position in the list </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
        Public Overloads Function IndexOf(ByVal Phrase As String, Optional ByVal by As IndexBy = IndexBy.PhraseId) As Integer
            Dim i As Integer
            Dim objPhrase As Phrase
            Dim retI As Integer = -1
            For i = 0 To Me.Count - 1
                objPhrase = Me.Item(i)
                If by = IndexBy.PhraseId Then
                    If objPhrase.PhraseID = Phrase Then
                        retI = i
                        Exit For
                    End If
                Else
                    If objPhrase.PhraseText = Phrase Then
                        retI = i
                        Exit For
                    End If

                End If
            Next
            Return retI
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load the phrase list from the database
        ''' </summary>
        ''' <param name="OrderBy">Order by clause, default = Order Number</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub Load(Optional ByVal OrderBy As String = "order_num")
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim objPhrase As Phrase

            Try
                oCmd.CommandText = myFields.FullRowSelect & " where phrase_type = ? " & "order by " & OrderBy
                oCmd.Parameters.Add(CConnection.StringParameter("Phrasetype", strPhraseType, 10))
                oCmd.Connection = CConnection.DbConnection
                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                objPhrase = New Phrase(strPhraseType)
                Me.Add(objPhrase)
                While oRead.Read
                    objPhrase = New Phrase()
                    objPhrase.Fields.readfields(oRead)
                    Me.Add(objPhrase)
                End While
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
            End Try
        End Sub

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Phrase list for the identified phrase list type 
    ''' </summary>
    ''' <param name="PhraseType">Phrase List Type</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal PhraseType As String)
      MyBase.new()
      setupFields()
      Me.strPhraseType = PhraseType
    End Sub

        ''' <summary>
        ''' Created by taylor on ANDREW at 2/15/2007 14:42:07
        '''     Create a phrase list using the from parameter according to the method 
        '''     described by in BuildUsing Enumeration
        ''' </summary>
        ''' <param name="From" type="String">
        '''     <para>
        '''         Table or Pl/Sql function
        '''     </para>
        ''' </param>
        ''' <param name="BuildUsing" type="MedscreenLib.Glossary.PhraseCollection.BuildBy">
        '''     <para>
        '''         Way to build collection
        '''     </para>
        ''' </param>
        ''' <param name="WhereClause" type="String">
        '''     <para>
        '''         Optional where clause missing the word where, only used for table types
        '''     </para>
        ''' </param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>taylor</Author><date> 2/15/2007</date><Action></Action></revision>
        ''' </revisionHistory>
        '''     
        ''' 
        Public Sub New(ByVal From As String, ByVal BuildUsing As BuildBy, _
        Optional ByVal WhereClause As String = "", Optional ByVal SplitOn As String = ",", _
        Optional ByVal parameters As String = "")
            MyClass.new("") 'Call root constructor, with a blank phrase type
            Select Case BuildUsing

                Case BuildBy.PLSQLFunction  'USe a PL/SQL function
                    Dim oColl As New Collection()
                    Dim strCSV As String
                    If parameters.Trim.Length = 0 Then
                        strCSV = CConnection.PackageStringList(From, "")
                    Else                'Need to split parameters
                        Dim strParam As String() = parameters.Split(New Char() {","})
                        Dim i As Integer
                        For i = 0 To strParam.Length - 1
                            oColl.Add(strParam.GetValue(i))
                        Next
                        strCSV = CConnection.PackageStringList(From, oColl)

                    End If
                    If strCSV.Trim.Length = 0 Then Exit Sub
                    Dim strList As String() = strCSV.Split(New Char() {SplitOn})
                    Try
                        'Go through list building collection 
                        Dim i As Integer
                        For i = 0 To strList.Length - 1
                            Dim strElement As String = strList.GetValue(i)
                            If strElement.Trim.Length > 0 Then
                                Dim strElements As String() = strElement.Split(New Char() {"|"})
                                'Construct Phrase
                                Dim objPhrase As New Phrase()
                                objPhrase.PhraseID = strElements.GetValue(0)
                                objPhrase.PhraseText = strElements.GetValue(1)
                                objPhrase.OrderNum = i + 1
                                Me.Add(objPhrase) 'And put in collection 
                            End If
                        Next
                    Catch ex As Exception
                    End Try
                    'based on table name
                Case BuildBy.Table
                    Dim objTableStructure As New MedStructure.StructTable(From) ' Create a structure object
                    If objTableStructure.HasFieldFor(MedStructure.FieldUsage.RemoveFlag) Then
                        GetFromTableStruct(objTableStructure, WhereClause)
                    Else
                        GetFromTableStruct(objTableStructure, "")
                    End If
                Case BuildBy.Query
                    Dim oCmd As New OleDb.OleDbCommand()
                    ' Get the select command based on the fields required
                    oCmd.CommandText = From
                    'see if we have a where clause
                    If WhereClause.Trim.Length > 0 Then
                        If InStr(From.ToUpper, "WHERE ") = 0 Then
                            oCmd.CommandText = From & " where " & WhereClause
                        Else
                            oCmd.CommandText = From & " and " & WhereClause
                        End If
                    End If
                    'Now we can execute the query 
                    Dim oRead As OleDb.OleDbDataReader = Nothing
                    Dim i As Integer = 1

                    Try
                        oCmd.Connection = CConnection.DbConnection
                        If CConnection.ConnOpen Then            'Attempt to open reader
                            oRead = oCmd.ExecuteReader          'Get Date
                            While oRead.Read                    'Loop through data
                                'Construct Phrase
                                Dim objPhrase As New Phrase()
                                Dim strId As String = ""
                                Dim strDescription As String = ""

                                'Check data validity 
                                If Not oRead.IsDBNull(0) Then strId = oRead.GetValue(0)
                                If Not oRead.IsDBNull(1) Then strDescription = oRead.GetValue(1)
                                objPhrase.setPhraseID(strId)
                                objPhrase.setPhraseText(strDescription)
                                objPhrase.OrderNum = i
                                i += 1                          'Increment loop counter
                                Me.Add(objPhrase) 'And put in collection 

                            End While
                        End If
                    Catch ex As Exception
                        LogError(ex, , )
                    Finally
                        CConnection.SetConnClosed()             'Close connection
                        If Not oRead Is Nothing Then            'Try and close reader
                            If Not oRead.IsClosed Then oRead.Close()
                        End If
                    End Try

            End Select
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 2/16/2007 14:26:28
    '''     Create a new phrase collection based on a table structure
    ''' </summary>
    ''' <param name="TableStructure" type="MedscreenLib.MedStructure.StructTable">
    '''     <para>
    '''         Table structure 
    '''     </para>
    ''' </param>
    ''' <param name="WhereClause" type="String">
    '''     <para>
    '''         Where clause to use (optional)
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 2/16/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Public Sub New(ByVal TableStructure As MedStructure.StructTable, Optional ByVal WhereClause As String = "")
      GetFromTableStruct(TableStructure, WhereClause)
    End Sub

    Private Sub GetFromTableStruct(ByVal objTableStructure As MedStructure.StructTable, ByVal WhereClause As String)
            If objTableStructure Is Nothing Then Exit Sub 'Bog off we don't have a valid structure definition 
            ' Create a collection of fields we are interested in 
            Dim fields As New MedStructure.StructField.FieldCollection()

            If objTableStructure.PrimaryKeys.Length = 0 Then Exit Sub 'No primary keys
            fields.Add(objTableStructure.PrimaryKeys(0))
            Dim tmpFieldName As String = objTableStructure.PrimaryKeys(0).Name


            If objTableStructure.BrowseFields.Length = 0 Then Exit Sub ' No browse fields
            Dim intIndex As Integer = 0
            Dim blnFound As Boolean = False
            Dim browseIndex As Integer = 0              'Added to deal with MLP
            'Key fields can also be browse fields
            While intIndex < objTableStructure.BrowseFields.Length AndAlso Not blnFound
                If objTableStructure.BrowseFields(intIndex).Name <> tmpFieldName Then
                    fields.Add(objTableStructure.BrowseFields(intIndex))
                    browseIndex = fields.Count       'Added to deal with MLP
                    blnFound = True
                End If
                intIndex += 1
            End While

            Dim oCmd As OleDb.OleDbCommand
            ' Get the select command based on the fields required
            oCmd = objTableStructure.GetSelectCommand(fields, WhereClause)
            'Now we can execute the query 
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim i As Integer = 1

            browseIndex = objTableStructure.PrimaryKeys.Length
            Try
                oCmd.Connection = CConnection.DbConnection
                If CConnection.ConnOpen Then            'Attempt to open reader
                    oRead = oCmd.ExecuteReader          'Get Date
                    While oRead.Read                    'Loop through data
                        'Construct Phrase
                        Dim objPhrase As New Phrase()
                        Dim strId As String = ""
                        Dim strDescription As String = ""

                        'Check data validity 
                        If Not oRead.IsDBNull(0) Then strId = oRead.GetValue(0)
                        'Added to deal with MLP
                        If Not oRead.IsDBNull(browseIndex) Then strDescription = oRead.GetValue(browseIndex)
                        'If Not oRead.IsDBNull(1) Then strDescription = oRead.GetValue(1)
                        objPhrase.SetPhraseID(strId)
                        objPhrase.SetPhraseText(strDescription)
                        objPhrase.OrderNum = i
                        i += 1                          'Increment loop counter
                        Me.Add(objPhrase) 'And put in collection 

                    End While
                End If
            Catch ex As Exception
                LogError(ex, , )
            Finally
                CConnection.SetConnClosed()             'Close connection
                If Not oRead Is Nothing Then            'Try and close reader
                    If Not oRead.IsClosed Then oRead.Close()
                End If
            End Try


    End Sub
#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Phrase type for this phrase list
    ''' </summary>
    ''' <value>The phrase type associated with this phrase list</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseType() As String
      Get
        Return strPhraseType
      End Get

    End Property

    
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Phrase found by phrase ID
    ''' </summary>
    ''' <param name="Index">Position to retrieve from</param>
    ''' <value>Returns Phrase with the given ID</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal Index As String, Optional ByVal IndexBy As IndexBy = IndexBy.PhraseId) As Phrase
            Get
                Dim I As Integer
                Dim objPhrase As Phrase = Nothing

                For I = 0 To Count - 1
                    objPhrase = Me.Item(I)
                    If IndexBy = PhraseCollection.IndexBy.PhraseId Then
                        If objPhrase.PhraseID.Trim.ToUpper = Index.ToUpper Then
                            Exit For
                        End If
                    ElseIf IndexBy = PhraseCollection.IndexBy.OrderBY Then
                        If objPhrase.OrderNum.Trim.ToUpper = Index.Trim.ToUpper Then
                            Exit For
                        End If

                    End If
                    objPhrase = Nothing
                Next
                Return objPhrase
            End Get
            Set(ByVal Value As Phrase)

            End Set
        End Property

#End Region

#End Region
        Private Sub setupFields()
            myFields = New TableFields("PHRASE")

            objPhraseType = New StringField("PHRASE_TYPE", "", 10)
            myFields.Add(objPhraseType)
            objPhraseId = New StringField("PHRASE_ID", "", 10)
            myFields.Add(objPhraseId)
            objPhraseNum = New StringField("ORDER_NUM", "", 10)
            myFields.Add(objPhraseNum)
            objPhraseText = New StringField("PHRASE_TEXT", "", 234)
            myFields.Add(objPhraseText)
            'Lines below need uncommenting when the fields are added to the schema
            myFields.Add(objStatus)
            myFields.Add(objGenText1)
            myFields.Add(objGenText2)
            myFields.Add(objIcon)
        End Sub

        Private Class PhraseComparer
            Inherits System.Collections.Generic.Comparer(Of Phrase)

            Private myDirection As Integer

            Private col As Integer
            Private col2 As Integer = 0

            Public Sub New()
                col = 0
                myDirection = 1

            End Sub


            Public Sub New(ByVal iOption As Integer, ByVal idIrection As Integer)
                myDirection = idIrection
                col = iOption
            End Sub

            Public Overrides Function Compare(ByVal x As Phrase, ByVal y As Phrase) As Integer


                Dim result As Integer
                Try
                    'Convert input values to objects
                    
                    'check for dates 

                    result = [String].Compare(x.PhraseText, y.PhraseText)
                    'if equal and we are looking at a second column

                Catch ex As Exception
                    Return -1
                End Try
                'allow for changing direction
                Return result * myDirection
            End Function
        End Class

        Private Class PhraseFindClass
            Private surname As String
            Private forename As String
            Private Id As String
           

            Public Sub New(ByVal OfficerId As String)
                Id = OfficerId.ToUpper
            End Sub

            Public Function FindPhraseById(ByVal Item As Phrase) As Boolean
                Return Item.PhraseID.ToUpper = Id
            End Function

            Public Function FindPhraseByGentext2(ByVal Item As Phrase) As Boolean
                Return Item.GenText2.ToUpper = Id
            End Function

            
        End Class

    End Class

    Public Class FormulaTypes
        Inherits MedscreenLib.Glossary.PhraseCollection

        Private Shared myInstance As FormulaTypes

        Public Shared ReadOnly Property Info() As FormulaTypes
            Get
                If myInstance Is Nothing Then
                    myInstance = New FormulaTypes
                    myInstance.Load()
                End If
                Return myInstance
            End Get
        End Property

        Private Sub New()
            MyBase.New("RPTFLDNAME")

        End Sub
    End Class

#End Region

#Region "Collection Of Phrase lists"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.Phrases
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Phrase Lists
    ''' </summary>
    ''' <remarks>
    ''' This is a collection of Phrase lists
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class Phrases
        Inherits CollectionBase

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a phrase list to the collection
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As PhraseCollection) As Integer
            Return Me.List.Add(item)
        End Function

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construct a new list 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region "Properties"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a phrase list from the collection
        ''' </summary>
        ''' <param name="index">Phrase list to find</param>
        ''' <value>Returns a phrase list</value>
        ''' <remarks>
        '''   Phrases not already loaded are retrieved from the database and added to the collection
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public ReadOnly Property Item(ByVal index As String) As PhraseCollection
            Get
                Dim i As Integer
                Dim objPhrColl As PhraseCollection = Nothing
                ' Check to see if phrase is already loaded into our collection
                For i = 0 To Me.Count - 1
                    objPhrColl = CType(Me.List.Item(i), PhraseCollection)
                    If objPhrColl.PhraseType = index Then
                        Exit For
                    End If
                    objPhrColl = Nothing
                Next
                ' if not, load it
                If objPhrColl Is Nothing Then
                    objPhrColl = New PhraseCollection(index)
                    objPhrColl.Load()
                    Glossary.PhraseList.Add(objPhrColl)
                End If
                Return objPhrColl
            End Get
            'Set(ByVal Value As PhraseCollection)
            'End Set
        End Property

#End Region

#End Region

    End Class

#End Region ' Collection of phrase lists



#End Region

    'Currency type 
#Region "Currency"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.Currency
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Currency Entity, relates to the Currency table
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <seealso cref="ExchangeRate"/>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class Currency
#Region "Declarations"
        Private myFields As MedscreenLib.TableFields

        Private objCurrencyId As New StringField("CURRENCYID", "", 4, True)
        Private objDescription As New StringField("DESCRIPTION", "", 20)
        Private objSymbol As New StringField("SYMBOL", "", 3)
        Private objCharID As New StringField("SINGLECHARID", "", 1)
        Private objAsciiId As New IntegerField("ASCII_CODE", 0)
        Private objCulture As New StringField("CULTURE", "", 10)
#End Region

#Region "Public Instance"

#Region "Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Return the Currency in a form suitaable for use in HTML
        ''' </summary>
        ''' <returns>Currency in a form suitaable for use in HTML</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function HTMLCode() As String
            Select Case Me.CurrencyId
                Case "GBP"
                    Return "&#163;"
                Case "USD"
                    Return "$"
                Case "EUR"
                    Return "&#8364;"
                Case "SEK"
                    Return "K "
                Case Else
                    Return "&#163;"
            End Select
        End Function


#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' create a new Currency Entry
        ''' </summary>
        ''' <param name="CurrencyId">Id of the Currency</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal CurrencyId As String)
            myFields = New MedscreenLib.TableFields("Currency")
            myFields.Add(objCurrencyId)
            objCurrencyId.Value = CurrencyId
            objCurrencyId.OldValue = CurrencyId

            myFields.Add(objDescription)
            myFields.Add(objSymbol)
            myFields.Add(objCharID)
            myFields.Add(objAsciiId)
            myFields.Add(objCulture)
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Currency Id ISO representation of the Currency three character code
        ''' </summary>
        ''' <value>ISO representation of the Currency</value>
        ''' <remarks>
        ''' ISO representation of the Currency<para/>
        ''' <li>GBP - Pounds sterling</li>
        ''' <li>USD - US Dollors</li>
        ''' <li>EUR - Euros</li>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CurrencyId() As String
            Get
                Return Me.objCurrencyId.Value.ToString
            End Get
            Set(ByVal Value As String)
                objCurrencyId.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Culture for Currency used in formatting.
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	13/05/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Culture() As String
            Get
                Return Me.objCulture.Value
            End Get
            Set(ByVal Value As String)
                Me.objCulture.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description of the currency
        ''' </summary>
        ''' <value>Currency description</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Description() As String
            Get
                Return Me.objDescription.Value.ToString
            End Get
            Set(ByVal Value As String)
                objDescription.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Representation of the currency as a single character
        ''' </summary>
        ''' <value>single character Currency code</value>
        ''' <remarks>Currency represented as a single character <para/>
        ''' <li>S - Sterling</li>
        ''' <li>D - Dollars</li>
        ''' <li>E - Euros</li>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property SingleCharId() As Char
            Get
                Return Me.objCharID.Value.ToString.Chars(0)
            End Get
            Set(ByVal Value As Char)
                objCharID.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Ascii Code representation of the Currency
        ''' </summary>
        ''' <value>Ascii Code representation of the Currency</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property AsciiCode() As Integer
            Get
                Return CInt(Me.objAsciiId.Value)
            End Get
            Set(ByVal Value As Integer)
                objAsciiId.Value = Value
            End Set
        End Property

#End Region
#End Region

        Friend Property Fields() As TableFields
            Get
                Return Me.myFields
            End Get
            Set(ByVal Value As TableFields)
                objDescription.Value = Value.Item("DESCRIPTION").Value
                objDescription.OldValue = Value.Item("DESCRIPTION").OldValue
                objSymbol.Value = Value.Item("SYMBOL").Value
                objSymbol.OldValue = Value.Item("SYMBOL").OldValue
                objCharID.Value = Value.Item("SINGLECHARID").Value
                objCharID.OldValue = Value.Item("SINGLECHARID").OldValue
                objAsciiId.Value = Value.Item("ASCII_CODE").Value
                objAsciiId.OldValue = Value.Item("ASCII_CODE").OldValue
            End Set
        End Property

    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.Currencies
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Currency Items
    ''' </summary>
    ''' <remarks>
    ''' A very small collection as currently there are only three curencies in use
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class Currencies
        Inherits CollectionBase

#Region "Declarations"
        Private myFields As MedscreenLib.TableFields

        Private objCurrencyId As New StringField("CURRENCYID", "", 4, True)
        Private objDescription As New StringField("DESCRIPTION", "", 20)
        Private objSymbol As New StringField("SYMBOL", "", 3)
        Private objCharID As New StringField("SINGLECHARID", "", 1)
        Private objAsciiId As New IntegerField("ASCII_CODE", 0)
        Private objCulture As New StringField("CULTURE", "", 10)

#End Region

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Currencies Culture used for display purposes
        ''' </summary>
        ''' <param name="ID"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	24/07/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Culture(ByVal ID As String) As String
            Dim strReturn As String = "EN-GB"
            Dim objCurrency As Currency
            For Each objCurrency In Me.List
                If objCurrency.CurrencyId = ID Then
                    strReturn = objCurrency.Culture
                    Exit For
                End If
            Next
            Return strReturn
        End Function


        Public Function HTMLCODE(ByVal ID As String) As String
            Dim strReturn As String = "&#163"
            Dim objCurrency As Currency
            For Each objCurrency In Me.List
                If objCurrency.CurrencyId = ID Then
                    strReturn = objCurrency.HTMLCode
                    Exit For
                End If
            Next
            Return strReturn
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a currency Item to the list 
        ''' </summary>
        ''' <param name="item">Currency Item to add</param>
        ''' <returns>Position of added item in list</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As Currency) As Integer
            Return Me.List.Add(item)
        End Function

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construct and populate collection of currencies
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader
            Dim objCurr As Currency

            myFields = New MedscreenLib.TableFields("Currency")
            myFields.Add(objCurrencyId)
            myFields.Add(objDescription)
            myFields.Add(objSymbol)
            myFields.Add(objCharID)
            myFields.Add(objAsciiId)
            myFields.Add(objculture)

            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.FullRowSelect
            Try
                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Dim strID As String = oRead.GetValue(oRead.GetOrdinal("CURRENCYID"))
                    objCurr = New Currency(strID)
                    Me.Add(objCurr)
                    objCurr.Fields.readfields(oRead)
                End While
                oRead.Close()
            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()
            End Try
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieve Currency by its position in the List
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value>Currency Item</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As Integer) As Currency
            Get
                Return CType(Me.List.Item(index), Currency)
            End Get
            Set(ByVal Value As Currency)
                Me.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a currency by its 3 character ISO code
        ''' </summary>
        ''' <param name="Index">3 character code</param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads ReadOnly Property Item(ByVal Index As String) As Currency
            Get
                Dim i As Integer
                Dim objC As Currency = Nothing
                For i = 0 To Me.Count - 1
                    objC = Item(i)
                    If objC.CurrencyId = Index Then
                        Exit For
                    End If
                    objC = Nothing
                Next
                Return objC
            End Get
        End Property

#End Region
#End Region

    End Class
#End Region

    'Exchange Rates
#Region "ExchangeRates"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.ExchangeRate
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An exchange rate
    ''' </summary>
    ''' <remarks>
    ''' An exchange rate is for a particular currency for a period of time
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class ExchangeRate

#Region "Declarations"
        Private objMonth As New DateField("MONTHSTART", DateSerial(1800, 1, 1), True)
        Private objCurrId As New StringField("CURRENCYID", "", 4, True)
        Private objExchRate As New DoubleField("EXCHANGERATE", 0, )
        Private myFields As TableFields

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Exchange Rate Item
        ''' </summary>
        ''' <param name="Start">Date from this Exchange rate will start</param>
        ''' <param name="Currency">Currency it is for</param>
        ''' <param name="ExchRate">Actual Exchnage Rate</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Start As Date, ByVal Currency As String, ByVal ExchRate As Double)
            myFields = New TableFields("Exchange_rates")
            myFields.Add(objMonth)
            myFields.Add(objCurrId)
            myFields.Add(objExchRate)

            objMonth.Value = Start
            objMonth.OldValue = Start
            objCurrId.Value = Currency
            objCurrId.OldValue = Currency
            objExchRate.Value = ExchRate
            objExchRate.OldValue = ExchRate
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Start Date of the Exchange Rate
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property StartDate() As Date
            Get
                Return objMonth.Value
            End Get
            Set(ByVal Value As Date)
                objMonth.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Exchange Rate in force
        ''' </summary>
        ''' <value>Current exchange rate between this currency and the base currency (GBP)</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ExchangeRate() As Double
            Get
                Return Me.objExchRate.Value
            End Get
            Set(ByVal Value As Double)
                objExchRate.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Currency as three character code
        ''' </summary>
        ''' <value>Currency as three character code</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Currency() As String
            Get
                Return Me.objCurrId.Value
            End Get
            Set(ByVal Value As String)
                objCurrId.Value = Value
            End Set
        End Property
#End Region
#End Region

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.ExchangeRates
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Exchange Rates
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class ExchangeRates
        Inherits CollectionBase


#Region "Declarations"
        Private objMonth As New DateField("MONTHSTART", DateSerial(1800, 1, 1), True)
        Private objCurrId As New StringField("CURRENCYID", "", 4, True)
        Private objExchRate As New DoubleField("EXCHANGERATE", 0, )
        Private myFields As TableFields

#End Region

#Region "Public Instance"

#Region "Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add and Exchange rate item
        ''' </summary>
        ''' <param name="Item">Exchange Rate Item to add</param>
        ''' <returns>Position in the collection of added item
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal Item As ExchangeRate) As Integer
            Return Me.List.Add(Item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieve an exchange rate
        ''' </summary>
        ''' <param name="Currency">Currency</param>
        ''' <param name="Start"> date (any date whilst it is valid) </param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Item(ByVal Currency As String, Optional ByVal Start As Date = #1/1/1900#) As ExchangeRate
            Dim i As Integer
            'Dim tmpDate As Date
            Dim objER As ExchangeRate = Nothing

            If Start.Year = 1900 Then Start = Now
            Start = Start.AddDays(-(Start.Day - 1))
            Start = Date.Parse(Start.ToString("dd-MMM-yyyy"))

            For i = 0 To Me.Count - 1
                objER = Me.List.Item(i)
                If objER.Currency = Currency And _
                    Date.Compare(objER.StartDate, Start) = 0 Then
                    Exit For
                End If
                objER = Nothing
            Next

            Return objER
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert from one currency to another on a particular date
        ''' </summary>
        ''' <param name="InputCurrency">Currency to be converted from</param>
        ''' <param name="outputCurrency">Currency to be converted to</param>
        ''' <param name="InputValue">Amount to be converted</param>
        ''' <param name="at">Date when conversion takes place, defaults to null date (today)</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function Convert(ByVal InputCurrency As String, ByVal outputCurrency As String, _
        ByVal InputValue As Double, Optional ByVal at As Date = #1/1/1900#) As Double
            Dim oColl As New Collection()
            oColl.Add(InputValue)
            oColl.Add(InputCurrency)
            If outputCurrency Is Nothing Then outputCurrency = "GBP"
            oColl.Add(outputCurrency)

            If at.Year = 1900 Then
                oColl.Add(Now.ToString("dd-MMM-yy"))
            Else
                oColl.Add(at.ToString("dd-MMM-yy"))
            End If

            Dim strRet As String = CConnection.PackageStringList("LIB_CCTOOL.ConvertCurrency", oColl)
            If Not strRet Is Nothing Then
                Return CDbl(strRet)
            End If
            'Dim objXin As ExchangeRate
            'Dim objXOut As ExchangeRate

            'If InputCurrency = outputCurrency Then
            '    'Same currency no need to convert
            '    Return InputValue
            '    Exit Function
            'End If

            'objXin = Me.Item(InputCurrency, at)
            'objXOut = Me.Item(outputCurrency, at)

            'If objXin Is Nothing Or objXOut Is Nothing Then
            '    'An error has occured should raise an exception
            '    Return InputValue
            '    Exit Function
            'End If

            'Return (InputValue / objXin.ExchangeRate) * objXOut.ExchangeRate

        End Function

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new collection and Add contents
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim objX As ExchangeRate

            myFields = New TableFields("Exchange_rates")
            myFields.Add(objMonth)
            myFields.Add(objCurrId)
            myFields.Add(objExchRate)

            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.FullRowSelect
            Try
                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    myFields.readfields(oRead)
                    objX = New ExchangeRate(objMonth.Value, objCurrId.Value, objExchRate.Value)
                    Me.Add(objX)

                End While
                oRead.Close()
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
            End Try

        End Sub

#End Region

#Region "Properties"

#End Region
#End Region

    End Class
#End Region

    'Line items Sold and price list items
#Region "LineItems"


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.LineItemGroup
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A line item group
    ''' </summary>
    ''' <remarks>
    ''' Line item groups tracks the way line items are grouped together
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' <revisionHistory>
    ''' <revision author="Doug Taylor" date="27/09/2005">Commented</revision>
    '''</revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LineItemGroup

#Region "Declarations"
        Private objID As StringField = New StringField("ID", "", 10, True)
        Private objDescription As StringField = New StringField("DESCRIPTION", "", 40)
        Private objParentID As StringField = New StringField("PARENT_ID", "", 10)
        Private objCategory As StringField = New StringField("CATEGORY", "", 10)

        Private myFields As TableFields = New TableFields("LINEITEM_GROUPS")
        Private myParentIndex As Integer = -1
        Private myIndex As Integer = -1
        Private myHasChildren As Boolean = False

#End Region

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Parent of Item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Parent() As String
            Dim oLIG As LineItemGroup
            oLIG = MedscreenLib.Glossary.Glossary.LineItemGroups.Item(Me.Index)
            While oLIG.ParentIndex > -1
                oLIG = MedscreenLib.Glossary.Glossary.LineItemGroups.Item(oLIG.ParentID)
            End While
            Return oLIG.Parent
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Checks to see if the item has any parents
        ''' </summary>
        ''' <param name="ParentId">Parent to see if item is a descendent</param>
        ''' <returns>TRUE if item is a parent</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function HasParent(ByVal ParentId As String) As Boolean
            Dim oLIG As LineItemGroup
            oLIG = MedscreenLib.Glossary.Glossary.LineItemGroups.Item(Me.Index)
            While oLIG.ParentID <> ParentId And oLIG.ParentIndex > -1
                oLIG = MedscreenLib.Glossary.Glossary.LineItemGroups.Item(oLIG.ParentID)
            End While
            Return (oLIG.ParentID = ParentId)
        End Function

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor create group list
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()

            SetupFields()
        End Sub

#End Region

#Region "Properties"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Position of Parent in list
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ParentIndex() As Integer
            Get
                Return myParentIndex
            End Get
            Set(ByVal Value As Integer)
                myParentIndex = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Id of parent in list 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ParentID() As String
            Get
                Return Me.objParentID.Value
            End Get
            Set(ByVal Value As String)
                objParentID.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Category Group is in
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Category() As String
            Get
                Return Me.objCategory.Value
            End Get
            Set(ByVal Value As String)
                objCategory.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description of group
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Description() As String
            Get
                Return Me.objDescription.Value
            End Get
            Set(ByVal Value As String)
                objDescription.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Indicates whether there are children of this item
        ''' </summary>
        ''' <value>True If there are Children</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property HasChildren() As Boolean
            Get
                Return myHasChildren
            End Get
            Set(ByVal Value As Boolean)
                myHasChildren = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ID of this group
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ID() As String
            Get
                Return Me.objID.Value
            End Get
            Set(ByVal Value As String)
                objID.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Position of this group
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Index() As Integer
            Get
                Return myIndex
            End Get
            Set(ByVal Value As Integer)
                myIndex = Value
            End Set
        End Property

#End Region
#End Region

        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

        Private Sub SetupFields()
            myFields.Add(objID)
            myFields.Add(Me.objDescription)
            myFields.Add(Me.objParentID)
            myFields.Add(Me.objCategory)

        End Sub

    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.LineItemGroups
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Line Item Groups
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LineItemGroups
        Inherits CollectionBase

#Region "Declarations"
        Private objID As StringField = New StringField("ID", "", 10, True)
        Private objDescription As StringField = New StringField("DESCRIPTION", "", 40)
        Private objParentID As StringField = New StringField("PARENT_ID", "", 10)
        Private objCategory As StringField = New StringField("CATEGORY", "", 10)

        Private myFields As TableFields = New TableFields("LINEITEM_GROUPS")

#End Region

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a new Line Item Group object to the collection
        ''' </summary>
        ''' <param name="item">Line Item Group object to be added</param>
        ''' <returns>Poistion of Added Item in collection</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As LineItemGroup) As Integer
            Return MyBase.List.Add(item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load the Line Item Group collection from the database
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Dim strquery As String = "select id,description,category,parent_id " & _
                    "from lineitem_groups " & _
                    "start with parent_id is null " & _
                    "connect by  parent_id = prior id"
            'Dim oCmd As New OracleClient.OracleCommand()
            'Dim oRead As OracleClient.OracleDataReader
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim lastLGroup As LineItemGroup = Nothing



            Try
                'oCmd.Connection = support.DbOraConnection
                oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = strquery

                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Dim liGroup As New LineItemGroup()
                    liGroup.Fields.readfields(oRead)

                    If liGroup.ParentID.Trim.Length = 0 Then
                        liGroup.ParentIndex = -1
                    Else
                        lastLGroup = Me.Item(liGroup.ParentID)
                        liGroup.ParentIndex = lastLGroup.Index
                        lastLGroup.HasChildren = True
                    End If
                    liGroup.Index = Me.Add(liGroup)

                End While




            Catch ex As Exception
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()

            End Try
        End Function

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructe a new colelction of Line Item Groups
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()

            SetupFields()
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieve a Line Item group from collection, indexing by Position
        ''' </summary>
        ''' <param name="index">Position to retrieve</param>
        ''' <value>Line Item Group </value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As Integer) As LineItemGroup
            Get
                Return CType(MyBase.List.Item(index), LineItemGroup)
            End Get
            Set(ByVal Value As LineItemGroup)
                MyBase.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieve a Line Item group from collection, indexing by Line Item Group ID
        ''' </summary>
        ''' <param name="index">Group ID to retrieve</param>
        ''' <value>Line Item Group </value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As String) As LineItemGroup
            Get
                Dim i As Integer
                Dim liGroup As LineItemGroup = Nothing


                For i = 0 To Me.Count - 1
                    liGroup = Item(i)
                    If liGroup.ID = index Then
                        Exit For
                    End If
                    liGroup = Nothing
                Next

                Return liGroup
            End Get
            Set(ByVal Value As LineItemGroup)

            End Set
        End Property

#End Region
#End Region

        Private Sub SetupFields()
            myFields.Add(objID)
            myFields.Add(Me.objDescription)
            myFields.Add(Me.objParentID)
            myFields.Add(Me.objCategory)

        End Sub

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.LineItemPrice
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The price of a line item.  
    ''' </summary>
    ''' <remarks>
    ''' Line items will have many prices in their lifetime.  We have adopted the model which enables us 
    ''' to have both historical and future price list implementations.  The price list is also 
    ''' multi currency.  <para/>
    ''' For example if a price change is envisaged to go live on 28th July 2006, it is possible to create this
    ''' weeks in advance, the price will not be charged until its start date <see cref="LineItemPrice.StartDate"/> is passed.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LineItemPrice

#Region "Declaration"


        Private objItemcode As New StringField("ITEMCODE", "", 20, True)
        Private objPrice As New DoubleField("PRICE", 0.0)
        Private objStartdate As New DateField("STARTDATE", DateField.ZeroDate, True)
        Private objEnddate As New DateField("ENDDATE", DateField.ZeroDate)
        Private objCurrencyid As New StringField("CURRENCYID", "", 4, True)
        Private objPriceBook As New StringField("PRICEBOOK", "", 10, True)

        Private myFields As New TableFields("PRICES")
#End Region

#Region "Public Instance"

        Public Function ToXMLElement() As XElement
            Dim Element As New XElement("LineItemPrice")
            Element.Add(New XElement("Price", Price))
            Element.Add(New XElement("Description", Description))
            Element.Add(New XElement("LineItem", Lineitem))
            Element.Add(New XElement("EndDate", EndDate))
            Element.Add(New XElement("StartDate", StartDate))
            Element.Add(New XElement("Currency", Currency))
            Element.Add(New XElement("PriceBook", PriceBook))
            Element.Add(New XElement("LineItemGroup", LineItemGroup))
            Return Element
        End Function

        Public Function ToXMLElementFull() As XElement
            Dim Element As XElement = ToXMLElement()
            Return Element
        End Function


        Public Function Update() As Boolean
            Dim blnRet As Boolean = True
            Try
                myFields.GetRowId()
                If myFields.RowID.Trim.Length = 0 Then
                    blnRet = myFields.Insert(CConnection.DbConnection)
                Else
                    blnRet = myFields.Update(CConnection.DbConnection)
                End If
            Catch ex As Exception
                blnRet = False
            End Try
            Return blnRet
        End Function

        Public Function Insert() As Boolean
            Dim blnRet As Boolean = True
            objEnddate.SetNull()
            Try
                blnRet = myFields.Insert(CConnection.DbConnection)
            Catch ex As Exception

            End Try
            Return blnRet
        End Function
#Region "Functions"

#End Region

#Region "Procedures"
        '''<summary>Setup Fields</summary>
        Public Sub SetUpFields()
            myFields.Add(objItemcode)
            myFields.Add(objPrice)
            myFields.Add(objStartdate)
            myFields.Add(objEnddate)
            myFields.Add(objCurrencyid)
            myFields.Add(objPriceBook)
        End Sub

        '''<summary>Constructor, create a new Price instance</summary>
        Public Sub New()
            SetUpFields()
        End Sub

#End Region

#Region "Properties"

        Public Property PriceBook() As String
            Get
                Return objPriceBook.Value
            End Get
            Set(ByVal value As String)
                objPriceBook.Value = value
            End Set
        End Property


        Private myDescription As String
        Public Property Description() As String
            Get
                Return myDescription
            End Get
            Set(ByVal value As String)
                myDescription = value
            End Set
        End Property


        Private myLineItemGroup As String
        Public Property LineItemGroup() As String
            Get
                Return myLineItemGroup
            End Get
            Set(ByVal value As String)
                myLineItemGroup = value
            End Set
        End Property


        '''<summary>Indicates whether it is a fixed price or variable</summary>
        ''' <remarks>A variable Price is indicated by having a null value for the price in the database</remarks>
        Public ReadOnly Property VariablePrice() As Boolean
            Get
                Return (Me.objPrice.IsNull)
            End Get
        End Property

        '''<summary></summary>
        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

        '''<summary>Date this price becomes effective</summary>
        Public Property StartDate() As Date
            Get
                Return Me.objStartdate.Value
            End Get
            Set(ByVal Value As Date)
                objStartdate.Value = Value
            End Set
        End Property

        '''<summary>Date this Price Ends</summary>
        Public Property EndDate() As Date
            Get
                Return Me.objEnddate.Value
            End Get
            Set(ByVal Value As Date)
                objEnddate.Value = Value
            End Set
        End Property

        '''<summary>Currency as three character ISO code</summary>
        Public Property Currency() As String
            Get
                Return Me.objCurrencyid.Value
            End Get
            Set(ByVal Value As String)
                objCurrencyid.Value = Value
            End Set
        End Property

        '''<summary>Lineitem Code</summary>
        Public Property Lineitem() As String
            Get
                Return Me.objItemcode.Value
            End Get
            Set(ByVal Value As String)
                objItemcode.Value = Value
            End Set
        End Property

        '''<summary>Price for Line Item</summary>
        Public Property Price() As Double
            Get
                Return Me.objPrice.Value
            End Get
            Set(ByVal Value As Double)
                objPrice.Value = Value
            End Set
        End Property

#End Region

#End Region

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.PriceList
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The list of price 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class PriceList
        Inherits CollectionBase

#Region "Declarations"

        Private objItemcode As New StringField("ITEMCODE", "", 10)
        Private objPrice As New DoubleField("PRICE", 0.0)
        Private objStartdate As New DateField("STARTDATE", DateField.ZeroDate)
        Private objEnddate As New DateField("ENDDATE", DateField.ZeroDate)
        Private objCurrencyid As New StringField("CURRENCYID", "", 4)
        Private objPriceBook As New StringField("PRICEBOOK", "", 10, True)

        Private myFields As New TableFields("PRICES")
#End Region

#Region "Public Instance"

#Region "Functions"
        '''<summary>Load a collection of prices</summary>
        ''' <param name='WhereClause'>What gets loaded, default all values</param>
        ''' <returns>>TRUE if succesful</returns>
        Public Function Load(Optional ByVal WhereClause As String = "") As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim objLIP As LineItemPrice

            Try
                oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = myFields.FullRowSelect
                If WhereClause.Trim.Length > 0 Then
                    oCmd.CommandText += " where " & WhereClause
                End If
                oCmd.CommandText += " order by startdate desc,currencyid"
                CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                While oRead.Read
                    objLIP = New LineItemPrice()
                    objLIP.Fields.readfields(oRead)
                    Me.Add(objLIP)
                End While
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                oRead = Nothing
                oCmd = Nothing
                CConnection.SetConnClosed()
            End Try
        End Function

        '''<summary>Add a Price to the collection</summary>
        ''' <param name='value'>Item to add to collection</param>
        ''' <returns>Index of item added</returns>
        Public Function Add(ByVal value As LineItemPrice) As Integer
            Return MyBase.List.Add(value)
        End Function


#End Region


#Region "Procedures"

        '''<summary>Constructor create a new price item</summary>
        Public Sub New()
            MyBase.New()
            SetUpFields()
        End Sub

#End Region

#Region "Properties"

        '''<summary>Get Price by numeric index</summary>
        ''' <param name='index'>position from 0 based collection to return</param>
        Default Public Overloads Property Item(ByVal index As Integer) As LineItemPrice
            Get
                Return CType(MyBase.List.Item(index), LineItemPrice)
            End Get
            Set(ByVal Value As LineItemPrice)
                MyBase.List.Item(index) = Value
            End Set
        End Property


        '''<summary></summary>
        ''' <param name='ItemCode'>Price to recover</param>
        ''' <param name='inDate'>Date for which the price needs to be valid</param>
        ''' <param name='Currency'>Currency price is in, defaults to GBP</param>
        Default Public Overloads Property Item(ByVal ItemCode As String, _
            ByVal inDate As Date, Optional ByVal Currency As String = "GBP") As LineItemPrice
            Get
                Dim i As Integer
                Dim objLIP As LineItemPrice = Nothing

                For i = 0 To Count - 1
                    objLIP = Me.Item(i)
                    If objLIP.Lineitem.ToUpper = ItemCode.ToUpper Then
                        If objLIP.Currency.ToUpper = Currency.ToUpper Then
                            If (objLIP.StartDate <= inDate) And ((objLIP.EndDate >= inDate) Or _
                                (objLIP.EndDate = DateField.ZeroDate)) Then
                                Exit For
                            End If
                        End If
                    End If
                    objLIP = Nothing
                Next
                Return objLIP
            End Get
            Set(ByVal value As LineItemPrice)
            End Set
        End Property

#End Region
#End Region

        '''<summary></summary>
        Private Sub SetUpFields()
            myFields.Add(objItemcode)
            myFields.Add(objPrice)
            myFields.Add(objStartdate)
            myFields.Add(objEnddate)
            myFields.Add(objCurrencyid)
            myFields.Add(objPriceBook)
        End Sub

    End Class

    Public Class LineItemPriceList
        Inherits System.Collections.Generic.List(Of LineItemPrice)

#Region "Declarations"
        Private myPriceBook As String
        Private myCurrency As String
#End Region

#Region "Public Instance"

#Region "Functions"
        Public Function ToXMLElement() As XElement
            Dim Element As New XElement("LineItemPriceList")
            Element.Add(New XElement("PriceBook", myPriceBook))
            Element.Add(New XElement("Currency", myCurrency))
            For Each li As LineItemPrice In Me
                Element.Add(li.ToXMLElement)
            Next
            Return Element
        End Function

        Public Function ToXMLElementFull() As XElement
            Dim Element As New XElement("LineItemPriceList")
            For Each li As LineItemPrice In Me
                Element.Add(li.ToXMLElementFull)
            Next
            Return Element
        End Function


        Public Function Load() As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim objLIP As LineItemPrice

            Try
                oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = "Select p.ITEMCODE,PRICE,STARTDATE,ENDDATE,CURRENCYID,PRICEBOOK,l.type,l.description From Prices p, lineitems l "
                oCmd.CommandText += " where pricebook = ? and currencyid = ?  and enddate is null and l.itemcode = p.itemcode"
                oCmd.CommandText += " order by startdate desc,currencyid"

                oCmd.Parameters.Add(CConnection.StringParameter("PRICEBOOK", myPriceBook, 10))
                oCmd.Parameters.Add(CConnection.StringParameter("CURRENCYID", myCurrency, 10))
                CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                While oRead.Read
                    objLIP = New LineItemPrice()
                    objLIP.Fields.readfields(oRead)
                    Me.Add(objLIP)
                    If Not oRead.IsDBNull(6) Then objLIP.LineItemGroup = oRead.GetValue(6)
                    If Not oRead.IsDBNull(7) Then objLIP.Description = oRead.GetValue(7)
                End While
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                oRead = Nothing
                oCmd = Nothing
                CConnection.SetConnClosed()
            End Try
        End Function

        Public Function FindByLineItem(ByVal ItemCode As String) As LineItemPrice
            Dim pred As New PredicateClass(ItemCode)
            Dim optColor As LineItemPrice = Me.Find(AddressOf pred.FindLineItem)

            Return optColor
        End Function

        Public Function FindByGroup(ByVal Group As String) As LineItemPriceList
            Dim pred As New PredicateClass(Group)
            Dim optColor As System.Collections.Generic.List(Of LineItemPrice) = Me.FindAll(AddressOf pred.FindCategory)
            'Dim optColor As LineItemPriceList = Me.FindAll(AddressOf pred.FindCategory)
            Dim newlist As New LineItemPriceList()
            For Each li As LineItemPrice In optColor
                newlist.Add(li)
            Next
            Return newlist
        End Function
#End Region

#Region "Procedures"
        Public Sub New(ByVal PriceBook As String, ByVal Currency As String)
            MyBase.New()
            myPriceBook = PriceBook
            myCurrency = Currency
            Load()
        End Sub

        Public Sub New()
            MyBase.New()
        End Sub
#End Region

#Region "Properties"

#End Region
#End Region

        Private Class PredicateClass
            Private Id As String
            Public Sub New(ByVal code As String)
                Id = code
            End Sub

            Public Function FindLineItem(ByVal Item As LineItemPrice) As Boolean
                Return Item.Lineitem.ToUpper = Id.ToUpper And Item.EndDate = DateField.ZeroDate
            End Function

            Public Function FindCategory(ByVal Item As LineItemPrice) As Boolean
                Return Item.LineItemGroup.ToUpper = Id.ToUpper And Item.EndDate = DateField.ZeroDate
            End Function
        End Class


    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.LineItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Line items defined 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action>Added removeflag and reference fields</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LineItem

#Region "Declarations"


        Private myFields As TableFields

        Private objItemCode As StringField       '  VARCHAR2(10)
        Private objType As StringField           ' VARCHAR2(10)
        Private objDescription As StringField    'VARCHAR2(45)
        Private objService As StringField = New StringField("SERVICE", "", 10)
        Private objRemoveflag As BooleanField = New BooleanField("REMOVEFLAG", "")
        Private objReference As StringField = New StringField("REFERENCE", "", 10)
        Private objLocation As StringField = New StringField("LOCATION", "", 10)

        Private myPrice As LineItemPrice = Nothing

        Private myCurrency As String = ""
#End Region

#Region "Public Instance"

#Region "enumerations"

        '''<summary>Types of Line Items</summary>
        Public Enum LineItemTypes
            '''<summary>Additional screening charges</summary>
            ADDSCREEN
            '''<summary>Additional time on site</summary>
            ADDTIME
            '''<summary>Cancellation charges</summary>
            CANCEL
            '''<summary>Collection charges</summary>
            COLLECTION
            '''<summary>Non Item line items</summary>
            COMMENT
            '''<summary>Confirmation</summary>
            CONFIRM
            '''<summary>Credits</summary>
            CREDIT
            '''<summary>Deleivery charges</summary>
            DELIVERY
            '''<summary>Equipment</summary>
            EQUIPMENT
            '''<summary>Expenses</summary>
            EXPENSE
            '''<summary>Kits and equipment</summary>
            KITS
            '''<summary>Medical Review</summary>
            MED_REVIEW
            '''<summary>Minimum Charge</summary>
            MINCHARGE
            '''<summary>Kits and equipment</summary>
            PAIDKITS
            '''<summary>Retainers for services</summary>
            RETAINER
            '''<summary>Sample Charge</summary>
            SAMPLE
            '''<summary>Charge for Services</summary>
            SERVICES
            '''<summary>Additional Test</summary>
            TEST
        End Enum
#End Region

#Region "Function"


        '''<summary>Convert Line Item into XML</summary>
        Public Function toXML() As String
            Dim strRet As String = "<LineItem>"

            Dim i As Integer
            Dim objTf As TableField

            For i = 0 To myFields.Count - 1
                objTf = myFields.Item(i)
                strRet += objTf.ToXMLSchema
            Next
            strRet += "</LineItem>"
            Return strRet
        End Function

#End Region

#Region "Procedures"
        '''<summary>Create a new line item entity </summary>
        Public Sub New()
            SetupFields()
        End Sub

        Public Sub New(ByVal Item As String)
            MyClass.New()
            Me.Fields.Load(CConnection.DbConnection, "Itemcode = '" & Item & "'", , True)
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Reference for line item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Reference() As String
            Get
                Return Me.objReference.Value
            End Get
            Set(ByVal Value As String)
                objReference.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Location e.g warehouse that item is stored in.
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	13/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Location() As String
            Get
                Return Me.objLocation.Value
            End Get
            Set(ByVal Value As String)
                Me.objLocation.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Has this line item been removed 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Removed() As Boolean
            Get
                Return Me.objRemoveflag.Value
            End Get
            Set(ByVal Value As Boolean)
                objRemoveflag.Value = Value
            End Set
        End Property
        '''<summary></summary>
        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

        '''<summary>Item Code ID</summary>
        Public Property ItemCode() As String
            Get
                Return Me.objItemCode.Value
            End Get
            Set(ByVal Value As String)
                objItemCode.Value = Value
            End Set
        End Property

        '''<summary>Type of Line Item (string)</summary>
        Public Property Type() As String
            Get
                Return objType.Value
            End Get
            Set(ByVal Value As String)
                objType.Value = Value
            End Set
        End Property

        '''<summary>Description of Line Item </summary>
        Public Property Description() As String
            Get
                Return Me.objDescription.Value
            End Get
            Set(ByVal Value As String)
                Me.objDescription.Value = Value
            End Set
        End Property

        '''<summary>Service associated with Line Item </summary>
        Public Property Service() As String
            Get
                Return Me.objService.Value
            End Get
            Set(ByVal Value As String)
                objService.Value = Value
            End Set
        End Property

        '''<summary>Currency</summary>
        Public Property Currency() As String
            Get
                Return Me.myCurrency
            End Get
            Set(ByVal Value As String)
                myCurrency = Value
            End Set
        End Property

        '''<summary>Price of Line Item</summary>
        Public ReadOnly Property Price() As LineItemPrice
            Get
                If Me.ItemCode.Trim.Length > 0 Then
                    If Me.myPrice Is Nothing And Me.Currency.Trim.Length > 0 Then
                        myPrice = Glossary.PriceList.Item(Me.ItemCode, Now, Me.Currency)
                    End If
                End If
                Return myPrice
            End Get
        End Property

        '''<summary>Type of LineItem as an enumeration</summary>
        Public ReadOnly Property LineItemType() As LineItemTypes
            Get
                Dim RetPtype As LineItemTypes = LineItemTypes.SAMPLE
                If Me.Type.Trim.Length > 0 Then
                    Dim strPType As String = Type.ToUpper
                    RetPtype = [Enum].Parse(GetType(LineItemTypes), strPType)
                End If
                Return RetPtype
            End Get
        End Property

#End Region

#End Region


        '''<summary></summary>
        Private Sub SetupFields()
            myFields = New TableFields("LINEITEMS")

            objItemCode = New StringField("ITEMCODE", "", 10, True)
            myFields.Add(objItemCode)
            objType = New StringField("TYPE", "", 10)
            myFields.Add(objType)
            objDescription = New StringField("DESCRIPTION", "", 45)
            myFields.Add(objDescription)
            myFields.Add(Me.objService)
            myFields.Add(Me.objRemoveflag)
            myFields.Add(Me.objReference)
            myFields.Add(Me.objLocation)
        End Sub


    End Class

    '''<summary>Line Items sold to customers, a collection of <see cref="LineItem "/></summary>
    Public Class LineItems

        Inherits CollectionBase

#Region "Declarations"


        Private myFields As TableFields

        Private objItemCode As StringField       '  VARCHAR2(10)
        Private objType As StringField           ' VARCHAR2(10)
        Private objDescription As StringField    'VARCHAR2(45)
        Private objService As StringField = New StringField("SERVICE", "", 10)
        Private objRemoveflag As BooleanField = New BooleanField("REMOVEFLAG", "")
        Private objReference As StringField = New StringField("REFERENCE", "", 10)
        Private objLocation As StringField = New StringField("LOCATION", "", 10)

        Private blnLoaded As Boolean = False

#End Region

#Region "Public Instance"

#Region "Functions"

        '''<summary>Load these line items</summary>
        Public Function Load() As Boolean
            'Dim oConn As New OleDb.OleDbConnection()
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

            Dim objItem As LineItem

            Dim strQuery As String
            strQuery = myFields.FullRowSelect


            If blnLoaded = True Then Return blnLoaded : Exit Function
            Try
                'oConn.ConnectionString = modSupport.ConnectString
                oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = strQuery
                'oConn.Open()
                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    myFields.readfields(oRead)
                    objItem = New LineItem()
                    objItem.ItemCode = myFields.Item("ITEMCODE").Value
                    objItem.Fields.Fill(Me.myFields)
                    Me.Add(objItem)
                End While
                oRead.Close()
                blnLoaded = True
            Catch ex As Exception
                Medscreen.LogError(ex)

            Finally
                'oConn.Close()
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                oRead = Nothing
                CConnection.SetConnClosed()
            End Try
            Return blnLoaded
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a new line item to the collection
        ''' </summary>
        ''' <param name="Item">Item to add</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Function Add(ByVal Item As LineItem) As Integer
            Return MyBase.List.Add(Item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get all the line items in a group 
        ''' </summary>
        ''' <param name="GroupID">Id of the group to get</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Group(ByVal GroupID As String) As LineItems
            Dim retGroup As LineItems = New LineItems(False)

            Dim i As Integer
            Dim liG As LineItem

            For i = 0 To Me.Count - 1
                liG = Item(i)
                If InStr(liG.ItemCode, GroupID) = 1 Then
                    retGroup.Add(liG)
                End If

            Next
            Return retGroup
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert collection to XML
        ''' </summary>
        ''' <returns>Collection represented as XML</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToXml() As String
            Dim StrRet As String = "<LineItems>"

            Dim I As Integer
            Dim objItem As LineItem

            For I = 0 To Count - 1
                objItem = Item(I)
                StrRet += objItem.toXML
            Next
            StrRet += "</LineItems>"
            Return StrRet
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Call stored procedure to find all line items that match a type
        ''' </summary>
        ''' <param name="_Type"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/09/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function LineItemsByType(ByVal _Type As String) As String
            Dim strRet As String = ""
            Dim ReturnString As Object = Nothing

            Dim ocmd As New OleDb.OleDbCommand()
            ocmd.Connection = MedConnection.Connection
            Try
                CConnection.SetConnOpen()
                ocmd.CommandText = "Select lib_utils.lineitemsbytype(?) from dual"
                ocmd.Parameters.Add(CConnection.StringParameter("type", _Type, 10))
                ReturnString = ocmd.ExecuteScalar
                If ReturnString Is Nothing OrElse ReturnString Is DBNull.Value Then
                    ReturnString = ""
                End If
            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()
            End Try
            Return CStr(ReturnString)
        End Function

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection of Line Items
        ''' </summary>
        ''' <param name="bLoad">Load the collection as well</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(Optional ByVal bLoad As Boolean = True)
            MyBase.New()
            SetupFields()
            If bLoad Then Load()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get Line Item from collection by position
        ''' </summary>
        ''' <param name="Index">Position to get</param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal Index As Integer) As LineItem
            Get
                Return CType(MyBase.List.Item(Index), LineItem)
            End Get
            Set(ByVal Value As LineItem)
                MyBase.List.Item(Index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get line item by item code
        ''' </summary>
        ''' <param name="Index">Item code to get </param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads ReadOnly Property Item(ByVal Index As String) As LineItem
            Get
                Dim objItem As LineItem = Nothing
                Dim I As Integer

                For I = 0 To Me.Count - 1
                    objItem = MyBase.List.Item(I)
                    If objItem.ItemCode = Index.ToUpper Then
                        Exit For
                    End If
                    objItem = Nothing
                Next
                Return objItem

            End Get
        End Property

#End Region

#Region "Properties"

#End Region

#End Region

        Private Sub SetupFields()
            myFields = New TableFields("LINEITEMS")

            objItemCode = New StringField("ITEMCODE", "", 10, True)
            myFields.Add(objItemCode)
            objType = New StringField("TYPE", "", 10)
            myFields.Add(objType)
            objDescription = New StringField("DESCRIPTION", "", 45)
            myFields.Add(objDescription)
            myFields.Add(Me.objService)
            myFields.Add(Me.objRemoveflag)
            myFields.Add(Me.objReference)
            myFields.Add(Me.objLocation)
        End Sub



    End Class


#End Region

    'Address Types
#Region "AddressTypes"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.AddressType
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible values of Address Type
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class xAddressType
        Inherits SMTable

#Region "Declarations"

        Private objAddrTypeID As IntegerField = New IntegerField("ADDRESS_TYPE", -1, True)
        Private objDescription As StringField = New StringField("DESCRIPTION", "", 30)
        Private objAddressType As StringField = New StringField("ADDRESS_TYPE_CHAR", "", 4)
#End Region


#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new AddressType Entity
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("ADDRESS_TYPE")
            MyBase.Fields.Add(Me.objAddrTypeID)
            MyBase.Fields.Add(Me.objDescription)
            MyBase.Fields.Add(Me.objAddressType)

        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Address Type Id 
        ''' </summary>
        ''' <value>Returns the Address Type as an Integer</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property AddressTypeID() As Integer
            Get
                Return Me.objAddrTypeID.Value
            End Get
            Set(ByVal Value As Integer)
                objAddrTypeID.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Address type 
        ''' </summary>
        ''' <value>Address type as a string</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property AddressType() As String
            Get
                Return Me.objAddressType.Value
            End Get
            Set(ByVal Value As String)
                objAddressType.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description of this address type
        ''' </summary>
        ''' <value>Address Type as a string</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Description() As String
            Get
                Return Me.objDescription.Value
            End Get
            Set(ByVal Value As String)
                objDescription.Value = Value
            End Set
        End Property


#End Region

#End Region


        Friend Shadows Property Fields() As TableFields
            Get
                Return MyBase.Fields
            End Get
            Set(ByVal Value As TableFields)
                MyBase.Fields = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.AddressTypes
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Address Types
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class xAddressTypes
        Inherits MedscreenLib.SMTableCollection

#Region "Declarations"


        Private objAddrTypeID As IntegerField = New IntegerField("ADDRESS_TYPE", -1, True)
        Private objDescription As StringField = New StringField("DESCRIPTION", "", 30)
        Private objAddressType As StringField = New StringField("ADDRESS_TYPE_CHAR", "", 4)
#End Region

#Region "Public Instance"

#Region "Function"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get the position in the list of an address type
        ''' </summary>
        ''' <param name="Item">Address Type to find</param>
        ''' <returns>Position in the list of the address type</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function IndexOf(ByVal Item As String) As Integer
            Dim objADTY As xAddressType
            Dim I As Integer
            Dim Ret As Integer = -1
            For I = 0 To MyBase.Count - 1
                objADTY = Me.Item(I)
                If objADTY.AddressType = Item Then
                    Ret = I
                    Exit For
                End If
            Next
            Return Ret
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get the position of item in list
        ''' </summary>
        ''' <param name="Item">Find by Address type Id</param>
        ''' <returns>Position in the list</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function IndexOf(ByVal Item As Integer) As Integer
            Dim objADTY As xAddressType
            Dim I As Integer
            Dim Ret As Integer = -1
            For I = 0 To MyBase.Count - 1
                objADTY = Me.Item(I)
                If objADTY.AddressTypeID = Item Then
                    Ret = I
                    Exit For
                End If
            Next
            Return Ret
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an item to the list
        ''' </summary>
        ''' <param name="Item">Address Type Item </param>
        ''' <returns>Position of Item added to the list</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal Item As xAddressType) As Integer
            Return MyBase.List.Add(Item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load the list from the database
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

            Dim blnRet As Boolean = False

            Try
                oCmd.Connection = MedConnection.Connection
                CConnection.SetConnOpen()
                oCmd.CommandText = Me.Fields.FullRowSelect & " order by ADDRESS_TYPE"

                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Dim oAT As xAddressType = New xAddressType()
                    oAT.Fields.readfields(oRead)
                    Me.Add(oAT)
                End While
                blnRet = True
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                CConnection.SetConnClosed()
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
            End Try

            Return blnRet
        End Function
#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor for List 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("ADDRESS_TYPE")
            MyBase.Fields.Add(Me.objAddrTypeID)
            MyBase.Fields.Add(Me.objDescription)
            MyBase.Fields.Add(Me.objAddressType)
        End Sub

#End Region

#Region "Properties"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Address Type Item
        ''' </summary>
        ''' <param name="index">Position in List</param>
        ''' <value>Retrieves an address type Item by its position in the list </value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As Integer) As xAddressType
            Get
                Return CType(MyBase.List.Item(index), xAddressType)
            End Get
            Set(ByVal Value As xAddressType)
                MyBase.List.Item(index) = Value
            End Set
        End Property

#End Region

#End Region

    End Class

#End Region

#Region "Miscellaneous"

#Region "Holidays"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.Holiday
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Stores a holiday item 
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class Holiday

#Region "Declarations"
    Private mydate As Date
    Private myDescription As String
#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new holiday entity
    ''' </summary>
    ''' <param name="datIn">Date for the Holiday</param>
    ''' <param name="myDesc">Description for the Holiday</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal datIn As Date, ByVal myDesc As String)
      mydate = datIn
      myDescription = myDesc
    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Date of the Holiday
    ''' </summary>
    ''' <value>The date associated with this Holiday</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property HolidayDate() As Date
      Get
        Return mydate
      End Get
      Set(ByVal Value As Date)
        mydate = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Description of the Holiday
    ''' </summary>
    ''' <value>Description of the Holiday</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property HolidayDescription() As String
      Get
        Return myDescription
      End Get
      Set(ByVal Value As String)
        myDescription = Value
      End Set
    End Property

#End Region
#End Region


  End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.HolidayCollection
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A collection of Holiday Items
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class HolidayCollection
    Inherits Collections.CollectionBase

#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a Holiday Item to the collection
    ''' </summary>
    ''' <param name="objTest"></param>
    ''' <returns>Position in the collection</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal objTest As Holiday) As Integer
      Return Me.List.Add(objTest)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is the date supplied a Holiday
    ''' </summary>
    ''' <param name="datIn">Date to check</param>
    ''' <returns>TRUE if supplied date is a holiday</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function IsAHoliday(ByVal datIn As Date) As Boolean
      Dim hol As Holiday
      For Each hol In Me.List
        If Date.Compare(hol.HolidayDate, datIn) = 0 Then
          Return True
          Exit Function
        End If
      Next
      Return False
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load the collection of Holidays
    ''' </summary>
    ''' <returns>Void</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            'Dim oConn As New OleDb.OleDbConnection()
            Dim oReader As OleDb.OleDbDataReader = Nothing
            Dim datTemp As Date
            Dim strDesc As String
            Dim blnRet As Boolean = True
            Try

                oCmd.Connection = CConnection.DbConnection
                CConnection.SetConnOpen()
                oCmd.CommandText = "Select holiday,description from tos_holiday order by holiday"
                oReader = oCmd.ExecuteReader
                While oReader.Read
                    datTemp = oReader.GetDateTime(0)
                    If oReader.IsDBNull(1) Then
                        strDesc = ""
                    Else
                        strDesc = oReader.GetString(1)
                    End If
                    Me.Add(New Holiday(datTemp, strDesc))
                End While

                oCmd = Nothing

            Catch ex As Exception
                blnRet = False
            Finally
                CConnection.SetConnClosed()
                If Not oReader Is Nothing Then
                    If Not oReader.IsClosed Then oReader.Close()

                End If

            End Try
            Return blnRet
        End Function


#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Remove holiday from list
    ''' </summary>
    ''' <param name="index">Position of the item to remove</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Remove(ByVal index As Integer)
      ' Check to see if there is a widget at the supplied index.
      If index > Count - 1 Or index < 0 Then
        ' If no widget exists, a messagebox is shown and the operation is 
        ' cancelled.
        'System.Windows.Forms.MessageBox.Show("Index not valid!")
      Else
        ' Invokes the RemoveAt method of the List object.
        List.RemoveAt(index)
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Holiday Collection
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New()
    End Sub

#End Region

#Region "Properties"
    ' This line declares the Item property as ReadOnly, and 
    ' declares that it will return a Widget object.
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a Holiday from collection
    ''' </summary>
    ''' <param name="index">Position of Item to get</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads ReadOnly Property Item(ByVal index As Integer) As Holiday
      Get
        ' The appropriate item is retrieved from the List object and 
        ' explicitly cast to the Widget type, then returned to the 
        ' caller.
        Return CType(List.Item(index), Holiday)
      End Get
    End Property


#End Region
#End Region




  End Class
#End Region

#Region "Infomaker"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Defaults.InfomakerLink
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Class to manage access to information stored in the database on InfoMaker tables
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class InfomakerLink
    Inherits SMTable
    Private objIdentity As StringField = New StringField("IDENTITY", "", 30, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 234)
    Private objLibrary As StringField = New StringField("LIBRARY", "", 100)
    Private objReport As StringField = New StringField("REPORT", "", 50)
    Private objUserLibrary As StringField = New StringField("USER_LIBRARY", "", 20)
    Private objValidRoutine As StringField = New StringField("VALID_ROUTINE", "", 40)
    Private objCopies As IntegerField = New IntegerField("COPIES", 1)
    Private objPreview As BooleanField = New BooleanField("PREVIEW", "?")
    Private objGroupId As StringField = New StringField("GROUP_ID", "", 10)
    Private objModifiable As BooleanField = New BooleanField("MODIFIABLE", "?")
    Private objModifiedOn As TimeStampField = New TimeStampField("MODIFIED_ON")
    Private objModifiedBy As UserField = New UserField("MODIFIED_BY")
    Private objRemoveFlag As BooleanField = New BooleanField("REMOVEFLAG", "?")


    Private myParameters As InfomakerParameterList

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Parameteers for this report 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Parameters() As InfomakerParameterList
      Get
        If Me.myParameters Is Nothing Then
          Me.myParameters = New InfomakerParameterList(Me.ReportID)
          Me.myParameters.Load()
        End If
        Return Me.myParameters
      End Get
      Set(ByVal Value As InfomakerParameterList)
        Me.myParameters = Value
      End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new instance of link
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New("INFOMAKER_link")
      MyBase.Fields.Add(objIdentity)
      MyBase.Fields.Add(objDescription)
      MyBase.Fields.Add(objLibrary)
      MyBase.Fields.Add(objReport)
      MyBase.Fields.Add(objUserLibrary)
      MyBase.Fields.Add(objValidRoutine)
      MyBase.Fields.Add(objCopies)
      MyBase.Fields.Add(objPreview)
      MyBase.Fields.Add(objGroupId)
      MyBase.Fields.Add(objModifiable)
      MyBase.Fields.Add(objModifiedOn)
      MyBase.Fields.Add(objModifiedBy)
      MyBase.Fields.Add(objRemoveFlag)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Primary key of the report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReportID() As String
      Get
        Return Me.objIdentity.Value
      End Get
      Set(ByVal Value As String)
        Me.objIdentity.Value = Value
      End Set
    End Property

    Friend Property DataFields() As TableFields
      Get
        Return MyBase.Fields
      End Get
      Set(ByVal Value As TableFields)
        MyBase.Fields = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Public description of the report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Description() As String
      Get
        Return Me.objDescription.Value
      End Get
      Set(ByVal Value As String)
        Me.objDescription.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Library to use 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Library() As String
      Get
        Return Me.objLibrary.Value
      End Get
      Set(ByVal Value As String)
        Me.objLibrary.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Report that in the library
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Report() As String
      Get
        Return Me.objReport.Value
      End Get
      Set(ByVal Value As String)
        Me.objReport.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' VGL routine to call 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property UserLibrary() As String
      Get
        Return Me.objUserLibrary.Value
      End Get
      Set(ByVal Value As String)
        Me.objUserLibrary.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validation routine (VGL)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ValidRoutine() As String
      Get
        Return Me.objValidRoutine.Value
      End Get
      Set(ByVal Value As String)
        objValidRoutine.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default number of Copies
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Copies() As Integer
      Get
        Return Me.objCopies.Value
      End Get
      Set(ByVal Value As Integer)
        Me.objCopies.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Preview the report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Preview() As Boolean
      Get
        Return Me.objPreview.Value
      End Get
      Set(ByVal Value As Boolean)
        Me.objPreview.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Security Group 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property GroupId() As String
      Get
        Return Me.objGroupId.Value
      End Get
      Set(ByVal Value As String)
        Me.objGroupId.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is this report modifiable
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Modifiable() As Boolean
      Get
        Return Me.objModifiable.Value
      End Get
      Set(ByVal Value As Boolean)
        Me.objModifiable.Value = Value
      End Set
    End Property

    Public Function PrintReport() As Boolean

      Try
        Dim objRep As Object = CreateObject("LabSystems.Imprint")

        objRep.set_server("DSN=johnlive")
        objRep.set_username("live")
        objRep.set_database("ODBC")
        objRep.set_password("live")
        objRep.set_preview(True)
        objRep.set_rtime("\\db01\live_imprint\oleimprint.pbd")
        'objRep.is_server = "TNS=DB01-DEV.medscreen.local"
        'objRep.set_connectextra("PROXY_UID=DEV;PROXY_PWD=DEV;")
        objRep.login()
        If objRep.logged_in() Then
          objRep.set_library("\\db01\live_imprint\" & Me.Library)
          objRep.set_report(Me.Report)
          'If objRep.set_save("E:\temp\temp.psr", "PSREPORT") Then
          objRep.set_copies(Me.Copies)
          objRep.set_preview(Me.Preview)
          Dim objParam As InfomakerParameter
          Dim k As Integer
          For k = 0 To Me.Parameters.Count - 2
            objParam = Me.Parameters.Item(k)
            If objParam.PROMPT1 = "$INFOMAKER_TEMP" Then
              Dim obj As Object = Medscreen.GetParameter(Medscreen.MyTypes.typDate, objParam.DESCRIPTION, "", Now)
              Dim objDate As Date = CDate(obj)
              objRep.add_parameter(objDate.ToString("dd-MMM-yyyy"))
            Else
              Dim obj As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, objParam.DESCRIPTION, "", objParam.DEFAULTVALUE)
              objRep.add_parameter(CStr(obj))
            End If
          Next
          'objRep.add_parameter("C-1-FEB-2006-005520")
          objRep.add_parameter(Now.ToString("dd-MMM-yyyy HH:mm:ss"))
          objRep.print()
          'End If
        End If

      Catch ex As Exception

      End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Date Report last modified
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ModifiedOn() As Date
      Get
        Return Me.objModifiedOn.Value
      End Get
      Set(ByVal Value As Date)
        Me.objModifiedOn.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Who last modified the report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ModifiedBy() As String
      Get
        Return Me.objModifiedBy.Value
      End Get
      Set(ByVal Value As String)
        objModifiedBy.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Has this report been removed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property RemoveFlag() As Boolean
      Get
        Return objRemoveFlag.Value
      End Get
      Set(ByVal Value As Boolean)
        Me.objRemoveFlag.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load this report
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function DoLoad() As Boolean
      Return MyBase.Load(CConnection.DbConnection)
    End Function
  End Class

  Public Class InfomakerLinkList
    Inherits SMTableCollection
    Private objIdentity As StringField = New StringField("IDENTITY", "", 30, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 234)
    Private objLibrary As StringField = New StringField("LIBRARY", "", 100)
    Private objReport As StringField = New StringField("REPORT", "", 50)
    Private objUserLibrary As StringField = New StringField("USER_LIBRARY", "", 20)
    Private objValidRoutine As StringField = New StringField("VALID_ROUTINE", "", 40)
    Private objCopies As IntegerField = New IntegerField("COPIES", 1)
    Private objPreview As BooleanField = New BooleanField("PREVIEW", "?")
    Private objGroupId As StringField = New StringField("GROUP_ID", "", 10)
    Private objModifiable As BooleanField = New BooleanField("MODIFIABLE", "?")
    Private objModifiedOn As TimeStampField = New TimeStampField("MODIFIED_ON")
    Private objModifiedBy As UserField = New UserField("MODIFIED_BY")
    Private objRemoveFlag As BooleanField = New BooleanField("REMOVEFLAG", "?")

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new instance of link
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New("INFOMAKER_link")
      MyBase.Fields.Add(objIdentity)
      MyBase.Fields.Add(objDescription)
      MyBase.Fields.Add(objLibrary)
      MyBase.Fields.Add(objReport)
      MyBase.Fields.Add(objUserLibrary)
      MyBase.Fields.Add(objValidRoutine)
      MyBase.Fields.Add(objCopies)
      MyBase.Fields.Add(objPreview)
      MyBase.Fields.Add(objGroupId)
      MyBase.Fields.Add(objModifiable)
      MyBase.Fields.Add(objModifiedOn)
      MyBase.Fields.Add(objModifiedBy)
      MyBase.Fields.Add(objRemoveFlag)

    End Sub

    Default Public Overloads Property Item(ByVal index As Integer) As InfomakerLink
      Get
        Return CType(MyBase.List.Item(index), InfomakerLink)
      End Get
      Set(ByVal Value As InfomakerLink)
        MyBase.List.Item(index) = Value
      End Set
    End Property

    Default Public Overloads Property Item(ByVal index As String) As InfomakerLink
      Get
                Dim objIM As InfomakerLink = Nothing
        For Each objIM In MyBase.List
          If objIM.ReportID = index Then
            Exit For
          End If
          objIM = Nothing
        Next
        Return objIM
      End Get
      Set(ByVal Value As InfomakerLink)

      End Set
    End Property

    Public Function Add(ByVal Item As InfomakerLink) As Integer
      Return MyBase.List.Add(Item)
    End Function

    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()

      'Dim oConn As New OleDb.OleDbConnection()
            Dim oReader As OleDb.OleDbDataReader = Nothing
      Dim blnRet As Boolean = False

      Try

        oCmd.Connection = CConnection.DbConnection
        CConnection.SetConnOpen()
        oCmd.CommandText = MyBase.Fields.FullRowSelect & " where removeflag = 'F' order by description"
        oReader = oCmd.ExecuteReader
        While oReader.Read
          Dim objInfoReport As InfomakerLink = New InfomakerLink()
          objInfoReport.DataFields.readfields(oReader)
          Me.Add(objInfoReport)
        End While

        oCmd = Nothing
        blnRet = True
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        CConnection.SetConnClosed()
        If Not oReader Is Nothing Then
          If Not oReader.IsClosed Then oReader.Close()
        End If

      End Try
      Return blnRet

    End Function
  End Class

  Public Class InfomakerParameter
    Inherits SMTable
    Private objINFOMAKERLINK As StringField = New StringField("INFOMAKER_LINK", "", 30, True)
    Private objIDENTITY As StringField = New StringField("IDENTITY", "", 10, True)
    Private objORDER As StringField = New StringField("ORDER_NUMBER", "", 10)
    Private objDESCRIPTION As StringField = New StringField("DESCRIPTION", "", 30)
    Private objDEFAULTVALUE As StringField = New StringField("DEFAULT_VALUE", "", 30)
    Private objTYPE As StringField = New StringField("TYPE", "", 10)
    Private objPROMPT1 As StringField = New StringField("PROMPT1", "", 30)
    Private objPROMPT2 As StringField = New StringField("PROMPT2", "", 30)
    Private objMANDATORY As BooleanField = New BooleanField("MANDATORY", "?")

    Public Sub New()
      MyBase.New("INFOMAKER_PARAMETERS")
      MyBase.Fields.Add(Me.objINFOMAKERLINK)
      MyBase.Fields.Add(Me.objIDENTITY)
      MyBase.Fields.Add(Me.objORDER)
      MyBase.Fields.Add(Me.objDESCRIPTION)
      MyBase.Fields.Add(Me.objDEFAULTVALUE)
      MyBase.Fields.Add(Me.objTYPE)
      MyBase.Fields.Add(Me.objPROMPT1)
      MyBase.Fields.Add(Me.objPROMPT2)
      MyBase.Fields.Add(Me.objMANDATORY)

    End Sub


    Friend Property DataFields() As TableFields
      Get
        Return MyBase.Fields
      End Get
      Set(ByVal Value As TableFields)
        MyBase.Fields = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Primary key of the report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ParameterID() As String
      Get
        Return Me.objIDENTITY.Value
      End Get
      Set(ByVal Value As String)
        Me.objIDENTITY.Value = Value
      End Set
    End Property


    Public Property ReportID() As String
      Get
        Return Me.objINFOMAKERLINK.Value
      End Get
      Set(ByVal Value As String)
        Me.objINFOMAKERLINK.Value = Value
      End Set
    End Property



    Public Property ORDER() As String
      Get
        Return Me.objORDER.Value
      End Get
      Set(ByVal Value As String)
        Me.objORDER.Value = Value
      End Set
    End Property

    Public Property DESCRIPTION() As String
      Get
        Return Me.objDESCRIPTION.Value
      End Get
      Set(ByVal Value As String)
        Me.objDESCRIPTION.Value = Value
      End Set
    End Property

    Public Property DEFAULTVALUE() As String
      Get
        Return Me.objDEFAULTVALUE.Value
      End Get
      Set(ByVal Value As String)
        Me.objDEFAULTVALUE.Value = Value
      End Set
    End Property

    Public Property TYPE() As String
      Get
        Return Me.objTYPE.Value
      End Get
      Set(ByVal Value As String)
        Me.objTYPE.Value = Value
      End Set
    End Property

    Public Property PROMPT1() As String
      Get
        Return Me.objPROMPT1.Value
      End Get
      Set(ByVal Value As String)
        Me.objPROMPT1.Value = Value
      End Set
    End Property

    Public Property PROMPT2() As String
      Get
        Return Me.objPROMPT2.Value
      End Get
      Set(ByVal Value As String)
        Me.objPROMPT2.Value = Value
      End Set
    End Property

    Public Property MANDATORY() As Boolean
      Get
        Return Me.objMANDATORY.Value
      End Get
      Set(ByVal Value As Boolean)
        Me.objMANDATORY.Value = Value
      End Set
    End Property


  End Class

  Public Class InfomakerParameterList
    Inherits SMTableCollection
    Private objINFOMAKERLINK As StringField = New StringField("INFOMAKER_LINK", "", 30, True)
    Private objIDENTITY As StringField = New StringField("IDENTITY", "", 10, True)
    Private objORDER As StringField = New StringField("ORDER_NUMBER", "", 10)
    Private objDESCRIPTION As StringField = New StringField("DESCRIPTION", "", 30)
    Private objDEFAULTVALUE As StringField = New StringField("DEFAULT_VALUE", "", 30)
    Private objTYPE As StringField = New StringField("TYPE", "", 10)
    Private objPROMPT1 As StringField = New StringField("PROMPT1", "", 30)
    Private objPROMPT2 As StringField = New StringField("PROMPT2", "", 30)
    Private objMANDATORY As BooleanField = New BooleanField("MANDATORY", "?")
    Private myReportId As String

    Public Sub New(ByVal ReportID As String)
      MyBase.New("INFOMAKER_PARAMETERS")
      MyBase.Fields.Add(Me.objINFOMAKERLINK)
      MyBase.Fields.Add(Me.objIDENTITY)
      MyBase.Fields.Add(Me.objORDER)
      MyBase.Fields.Add(Me.objDESCRIPTION)
      MyBase.Fields.Add(Me.objDEFAULTVALUE)
      MyBase.Fields.Add(Me.objTYPE)
      MyBase.Fields.Add(Me.objPROMPT1)
      MyBase.Fields.Add(Me.objPROMPT2)
      MyBase.Fields.Add(Me.objMANDATORY)
      Me.objINFOMAKERLINK.Value = ReportID
      Me.objINFOMAKERLINK.OldValue = ReportID
      myReportId = ReportID

    End Sub

    Public Property ReportId() As String
      Get
        Return Me.myReportId
      End Get
      Set(ByVal Value As String)
        myReportId = Value
      End Set
    End Property

    Default Public Property Item(ByVal index As Integer) As InfomakerParameter
      Get
        Return CType(MyBase.List.Item(index), InfomakerParameter)
      End Get
      Set(ByVal Value As InfomakerParameter)
        MyBase.List.Item(index) = Value
      End Set
    End Property

    Public Function Add(ByVal Item As InfomakerParameter) As Integer
      Return MyBase.List.Add(Item)
    End Function

    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()

      'Dim oConn As New OleDb.OleDbConnection()
            Dim oReader As OleDb.OleDbDataReader = Nothing
      Dim blnRet As Boolean = False

      Try

        oCmd.Connection = CConnection.DbConnection
        CConnection.SetConnOpen()
        oCmd.CommandText = MyBase.Fields.FullRowSelect & " where INFOMAKER_LINK = '" & myReportId & "' order by order_number"
        oReader = oCmd.ExecuteReader
        While oReader.Read
          Dim objInfoReport As InfomakerParameter = New InfomakerParameter()
          objInfoReport.DataFields.readfields(oReader)
          Me.Add(objInfoReport)
        End While

        oCmd = Nothing
        blnRet = True
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        CConnection.SetConnClosed()
        If Not oReader Is Nothing Then
          If Not oReader.IsClosed Then oReader.Close()
        End If

      End Try
      Return blnRet

    End Function

  End Class
#End Region

#Region "Printers"
    <CLSCompliant(True)> _
    Public Class Printer
        Inherits MedscreenLib.DataBasedObject

        Private Shared m_printers As PrinterColl

        '''<summary>
        '''Default collection of available printers
        '''</summary>
        Shared ReadOnly Property Printers() As PrinterColl
            Get
                If m_printers Is Nothing Then
                    m_printers = New PrinterColl
                    m_printers.Load()
                End If
                Return m_printers
            End Get
        End Property

        Public Sub New()
            MyBase.New(MedStructure.GetTable("Printer").CreateDataTable.NewRow)
        End Sub

        Friend Sub New(ByVal printerData As DataRow)
            MyBase.New(printerData)
        End Sub

        Public Property Identity() As String
            Get
                Return Me.GetField("Identity").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Identity", Value)
            End Set
        End Property


        Public Property GroupId() As String
            Get
                Return Me.GetField("Group_ID").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Group_ID", Value)
            End Set
        End Property


        Public Property Description() As String
            Get
                Return Me.GetField("Description").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Description", Value)
            End Set
        End Property


        Public Property Location() As String
            Get
                Return Me.GetField("Location").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Location", Value)
            End Set
        End Property


        Public Property LogicalName() As String
            Get
                Return Me.GetField("Logical_Name").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Logical_Name", Value)
            End Set
        End Property


        Public Property DeviceType() As String
            Get
                Return Me.GetField("Device_Type").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Device_Type", Value)
            End Set
        End Property


        Public Property ModifiedOn() As Date
            Get
                If m_data.IsNull("Modified_On") Then
                    Return DateTime.MinValue
                Else
                    Return DirectCast(Me.GetField("Modified_On"), DateTime)
                End If
            End Get
            Set(ByVal Value As Date)
                Me.SetField("Modified_On", Value)
            End Set
        End Property


        Public Property ModifiedBy() As String
            Get
                Return Me.GetField("Modified_By").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Modified_By", Value)
            End Set
        End Property


        Public Property Modifiable() As Boolean
            Get
                Return Me.GetField("Modifiable").ToString = C_dbTRUE
            End Get
            Set(ByVal Value As Boolean)
                Me.SetField("Modifiable", YesNo(Value, C_dbTRUE, C_dbFALSE))
            End Set
        End Property


        Public Property Removeflag() As Boolean
            Get
                Return Me.GetField("RemoveFlag").ToString = C_dbTRUE
            End Get
            Set(ByVal Value As Boolean)
                Me.SetField("RemoveFlag", YesNo(Value, C_dbTRUE, C_dbFALSE))
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Name of the device as used by reporter
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[boughton]</Author><date> [06/06/2008]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DeviceName() As String
            Get
                Return Me.GetField("Device_Name").ToString.Trim
            End Get
            Set(ByVal Value As String)
                Me.SetField("Device_Name", Value)
            End Set
        End Property


        Public Property PrinterCodeType() As String
            Get
                Return Me.GetField("Printer_Code_Type").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Printer_Code_Type", Value)
            End Set
        End Property


        Public Property PrinterParameter() As String
            Get
                Return Me.GetField("Printer_Parameter").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Printer_Parameter", Value)
            End Set
        End Property


        Public Property GraphicsType() As String
            Get
                Return Me.GetField("Graphics_Type").ToString
            End Get
            Set(ByVal Value As String)
                Me.SetField("Graphics_Type", Value)
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Property for use in combo boxes
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[boughton]</Author><date> [06/06/2008]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Info() As String
            Get
                Return Me.DeviceName & " - " & Description
            End Get
        End Property

        Public ReadOnly Property Loaded() As Boolean
            Get
                Return m_data IsNot Nothing AndAlso m_data.RowState <> DataRowState.Detached
            End Get
        End Property

        Public Function Update() As Boolean
            'MyBase.Fields.Update(CConnection.DbConnection)
        End Function

    End Class

#Region "Printer Collection"
    ''' <summary>
    ''' A collection of printer objects 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[boughton]</Author><date> [06/06/2008]</date><Action></Action></revision>
    ''' </revisionHistory>
    <CLSCompliant(True)> _
    Public Class PrinterColl
        Inherits Collections.Generic.List(Of Printer)

        Private m_filterID As String
        Private m_filterName As String
        Private m_allowLoad As Boolean
        Private m_reporterPrinters As PrinterColl

        Public Sub New()
            MyBase.New()
            m_allowLoad = True
        End Sub

        Protected Sub New(ByVal list As Collections.Generic.List(Of Printer))
            MyBase.New(list)
        End Sub

        ''' <summary>
        ''' Get printer by printer identity 
        ''' </summary>
        ''' <param name="printerID"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[boughton]</Author><date> [06/06/2008]</date><Action></Action></revision>
        ''' </revisionHistory>
        Default Public Overloads ReadOnly Property Item(ByVal printerID As String) As Printer
            Get
                m_filterID = printerID
                SyncLock m_filterID
                    Return Me.Find(AddressOf Me.FilterByID)
                End SyncLock
            End Get
        End Property

        Public ReadOnly Property GetByName(ByVal deviceName As String) As Printer
            Get
                m_filterName = deviceName
                SyncLock m_filterName
                    Return Me.Find(AddressOf Me.FilterByName)
                End SyncLock
            End Get
        End Property

        Public ReadOnly Property ReporterPrinters() As PrinterColl
            Get
                ' Don't return a Reporter Printers collection if this already is one
                If Not Me.IsReporterPrinterCollection Then
                    ' Use of Reporter Printers collection is very common, so worth caching.
                    ' Will get re-loaded if parent collection is refreshed
                    If m_reporterPrinters Is Nothing Then
                        m_reporterPrinters = New PrinterColl(Me.FindAll(AddressOf Me.FilterDeviceName))
                    End If
                End If
                Return m_reporterPrinters
            End Get
        End Property

        Public ReadOnly Property IsReporterPrinterCollection() As Boolean
            Get
                Return Not m_allowLoad
            End Get
        End Property

        Protected Function FilterByID(ByVal prn As Printer) As Boolean
            Return prn.Identity = m_filterID
        End Function

        Protected Function FilterDeviceName(ByVal prn As Printer) As Boolean
            Return prn.DeviceName > ""
        End Function

        Protected Function FilterByName(ByVal prn As Printer) As Boolean
            Return prn.DeviceName = m_filterName
        End Function

        ''' <summary>
        ''' Load the collection of printers
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[boughton]</Author><date> [06/06/2008]</date><Action></Action></revision>
        ''' </revisionHistory>
        Public Function Load() As Boolean
            If Me.Count = 0 Then
                Me.Refresh()
            End If
            Return Me.Count > 0
        End Function

        Public Sub Refresh()
            If m_allowLoad Then
                ' Reset Reporter Printers collections to ensure it is also refreshed when required
                m_reporterPrinters = Nothing
                Me.Clear()
                Dim tableDef As MedStructure.StructTable = MedStructure.GetTable("Printer")
                Dim getPrinters As New OleDb.OleDbDataAdapter(tableDef.GetSelectCommand(tableDef.Fields, tableDef.GetFields(tableDef.GetFieldsWithUse(MedStructure.FieldUsage.RemoveFlag))))
                getPrinters.SelectCommand.Parameters(0).Value = C_dbFALSE
                Try
                    Dim printers As DataTable = tableDef.CreateDataTable
                    getPrinters.Fill(printers)
                    For Each printerData As DataRow In printers.Rows
                        Me.Add(New Printer(printerData))
                    Next
                Catch ex As Exception
                    Medscreen.LogError(ex)
                End Try
            End If
        End Sub

    End Class
#End Region
#End Region
#End Region


#Region "Customer"
  Public Class Customer

    Shared Function GetSMIDProfile(ByVal identity As String) As String
      Dim name As String = identity
      Try
        Dim cmdGetSmidProfile As New OleDb.OleDbCommand("SELECT SMID ||'-'|| Profile FROM Customer WHERE Identity =?", _
                                                        MedConnection.Connection)
        cmdGetSmidProfile.Parameters.Add("Customer_ID", OleDb.OleDbType.VarChar, 10)
        MedConnection.Open()
        cmdGetSmidProfile.Parameters(0).Value = identity
        name = CType(cmdGetSmidProfile.ExecuteScalar, String)
        cmdGetSmidProfile.Dispose()
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        MedConnection.Close()
      End Try
      Return name
    End Function

    Private Sub New()
      ' prevent instantiation
        End Sub

        Shared Function GetCompanyName(ByVal CustomerID As String) As String
            Return ""
        End Function
  End Class
#End Region
#Region "Component views"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : Intranet
  ''' Class	 : intranet.TestResult.ComponentView
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A map to Component View
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class ComponentView
    Inherits MedscreenLib.SMTable


#Region "Declarations"
    Private objAnalysis As StringField = New StringField("ANALYSIS", "", 10)
    Private objAnalysisVersion As StringField = New StringField("ANALYSIS_VERSION", "", 10)
    Private objName As StringField = New StringField("NAME", "", 40)
    Private objOrderNumber As StringField = New StringField("ORDER_NUMBER", "", 10)
    Private objResultType As StringField = New StringField("RESULT_TYPE", "", 1)
    Private objUnits As StringField = New StringField("UNITS", "", 10)
    Private objMinimum As DoubleField = New DoubleField("MINIMUM", 0.0)
    Private objMaximum As DoubleField = New DoubleField("MAXIMUM", 0.0)
    Private objTrueWord As StringField = New StringField("TRUE_WORD", "", 10)
    Private objFalseWord As StringField = New StringField("FALSE_WORD", "", 10)
    Private objAllowedCharacters As StringField = New StringField("ALLOWED_CHARACTERS", "", 26)
    Private objCalculation As StringField = New StringField("CALCULATION", "", 10)
    Private objPlaces As IntegerField = New IntegerField("PLACES", 0.0)
    Private objRep_Control As StringField = New StringField("REP_CONTROL", "", 3)
    Private objReplicates As StringField = New StringField("REPLICATES", "", 1)
    Private objSig_FigsNumber As IntegerField = New IntegerField("SIG_FIGS_NUMBER", 0.0)
    Private objSig_FigsRounding As IntegerField = New IntegerField("SIG_FIGS_ROUNDING", 0.0)
    Private objSig_FigsFilter As StringField = New StringField("SIG_FIGS_FILTER", "", 10)
    Private objMinimumPql As DoubleField = New DoubleField("MINIMUM_PQL", 0.0)
    Private objMaximumPql As DoubleField = New DoubleField("MAXIMUM_PQL", 0.0)
    Private objPqlCalculation As StringField = New StringField("PQL_CALCULATION", "", 10)
    Private objFormula As StringField = New StringField("FORMULA", "", 500)

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible analysis types
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum AnalysisType
      ''' <summary>Enzyme Linked test</summary>
      EA_EnzymeAssay
      ''' <summary>Radio Immuno Assay</summary>
      IA_ImmunoAssay
      ''' <summary>Gas Chromatography test</summary>
      GC_GasChromatography
      ''' <summary>Mass spec test</summary>
      MS_MassSpectrometry
      ''' <summary>Adulteration test</summary>
      AD_Adulteration
      ''' <summary>Medical Review</summary>
      MR_MRO
      ''' <summary></summary>
      QI
      QM
      TE
      UG
    End Enum

    Private Sub SetupFields()
      MyBase.Fields.Add(objAnalysis)
      MyBase.Fields.Add(objAnalysisVersion)
      MyBase.Fields.Add(objName)
      MyBase.Fields.Add(objOrderNumber)
      MyBase.Fields.Add(objResultType)
      MyBase.Fields.Add(objUnits)
      MyBase.Fields.Add(objMinimum)
      MyBase.Fields.Add(objMaximum)
      MyBase.Fields.Add(objTrueWord)
      MyBase.Fields.Add(objFalseWord)
      MyBase.Fields.Add(objAllowedCharacters)
      MyBase.Fields.Add(objCalculation)
      MyBase.Fields.Add(objPlaces)
      MyBase.Fields.Add(objRep_Control)
      MyBase.Fields.Add(objReplicates)
      MyBase.Fields.Add(objSig_FigsNumber)
      MyBase.Fields.Add(objSig_FigsRounding)
      MyBase.Fields.Add(objSig_FigsFilter)
      MyBase.Fields.Add(objMinimumPql)
      MyBase.Fields.Add(objMaximumPql)
      MyBase.Fields.Add(objPqlCalculation)
      MyBase.Fields.Add(objFormula)

    End Sub

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ?Create a new entity for the component view
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New("COMPONENT_VIEW")
      SetupFields()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new instance by analysis and component position 
    ''' </summary>
    ''' <param name="Analysis">Analysis </param>
    ''' <param name="orderNumber">Component Position </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Analysis As String, ByVal orderNumber As Integer)
      MyClass.New()
      Me.objAnalysis.Value = Analysis
      Me.objAnalysis.OldValue = Analysis
      Me.objOrderNumber.Value = orderNumber
      Me.objOrderNumber.OldValue = orderNumber
      Dim pcoll As New Collection()
      pcoll.Add(CConnection.StringParameter("Analysis", Me.Analysis, 10))
      pcoll.Add(CConnection.IntegerParameter("OrderNo", Me.OrderNumber))
      Me.Fields.Load(MedscreenLib.MedConnection.Connection, "analysis = ?  and order_number = ?", pcoll)
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Units that results are presented in 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Units() As String
      Get
        Return Me.objUnits.Value
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of result produced
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Result_type() As String
      Get
        Return Me.objResultType.Value
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Name of test 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
      Get
        Return objName.Value
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Position in Component list 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property OrderNumber() As Integer
      Get
        Return CInt(Me.objOrderNumber.Value)
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Analysis ID associated with these components
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Analysis() As String
      Get
        Return Me.objAnalysis.Value
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is this test a confirmatory one or not 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsConfirmatory() As Boolean
      Return (InStr("MS-GC-MR-", Mid(Me.Analysis, 1, 2)) > 0)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose datafields 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
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
#End Region
  End Class


  ''' -----------------------------------------------------------------------------
  ''' Project	 : Intranet
  ''' Class	 : intranet.TestResult.ComponentViewList
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A list of components associated with an anlysis and thier tests
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [28/10/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class ComponentViewList
    Inherits MedscreenLib.SMTableCollection


#Region "Declarations"
    Private objAnalysis As StringField = New StringField("ANALYSIS", "", 10)
    Private objAnalysisVersion As StringField = New StringField("ANALYSIS_VERSION", "", 10)
    Private objName As StringField = New StringField("NAME", "", 40)
    Private objOrderNumber As StringField = New StringField("ORDER_NUMBER", "", 10)
    Private objResultType As StringField = New StringField("RESULT_TYPE", "", 1)
    Private objUnits As StringField = New StringField("UNITS", "", 10)
    Private objMinimum As DoubleField = New DoubleField("MINIMUM", 0.0)
    Private objMaximum As DoubleField = New DoubleField("MAXIMUM", 0.0)
    Private objTrueWord As StringField = New StringField("TRUE_WORD", "", 10)
    Private objFalseWord As StringField = New StringField("FALSE_WORD", "", 10)
    Private objAllowedCharacters As StringField = New StringField("ALLOWED_CHARACTERS", "", 26)
    Private objCalculation As StringField = New StringField("CALCULATION", "", 10)
    Private objPlaces As IntegerField = New IntegerField("PLACES", 0.0)
    Private objRep_Control As StringField = New StringField("REP_CONTROL", "", 3)
    Private objReplicates As StringField = New StringField("REPLICATES", "", 1)
    Private objSig_FigsNumber As IntegerField = New IntegerField("SIG_FIGS_NUMBER", 0.0)
    Private objSig_FigsRounding As IntegerField = New IntegerField("SIG_FIGS_ROUNDING", 0.0)
    Private objSig_FigsFilter As StringField = New StringField("SIG_FIGS_FILTER", "", 10)
    Private objMinimumPql As DoubleField = New DoubleField("MINIMUM_PQL", 0.0)
    Private objMaximumPql As DoubleField = New DoubleField("MAXIMUM_PQL", 0.0)
    Private objPqlCalculation As StringField = New StringField("PQL_CALCULATION", "", 10)
    Private objFormula As StringField = New StringField("FORMULA", "", 500)

    Private strCust As String = ""
#End Region

    Private Sub SetupFields()
      MyBase.Fields.Add(objAnalysis)
      MyBase.Fields.Add(objAnalysisVersion)
      MyBase.Fields.Add(objName)
      MyBase.Fields.Add(objOrderNumber)
      MyBase.Fields.Add(objResultType)
      MyBase.Fields.Add(objUnits)
      MyBase.Fields.Add(objMinimum)
      MyBase.Fields.Add(objMaximum)
      MyBase.Fields.Add(objTrueWord)
      MyBase.Fields.Add(objFalseWord)
      MyBase.Fields.Add(objAllowedCharacters)
      MyBase.Fields.Add(objCalculation)
      MyBase.Fields.Add(objPlaces)
      MyBase.Fields.Add(objRep_Control)
      MyBase.Fields.Add(objReplicates)
      MyBase.Fields.Add(objSig_FigsNumber)
      MyBase.Fields.Add(objSig_FigsRounding)
      MyBase.Fields.Add(objSig_FigsFilter)
      MyBase.Fields.Add(objMinimumPql)
      MyBase.Fields.Add(objMaximumPql)
      MyBase.Fields.Add(objPqlCalculation)
      MyBase.Fields.Add(objFormula)

    End Sub


#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load a collection of components 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()                    'Create connection object 
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.Connection = MedConnection.Connection          ' set up command object 
        oCmd.CommandText = MyBase.Fields.FullRowSelect        'set up query 
        If Me.strCust.Trim.Length > 0 Then
          oCmd.CommandText += " where ANALYSIS = ? " '& strCust & "'"    'adjust if we have an analysis id
        End If
        Dim id_parameter As OleDb.OleDbParameter
        id_parameter = New OleDb.OleDbParameter("ID", strCust)
        id_parameter.DbType = DbType.String
        id_parameter.Size = 10
        oCmd.Parameters.Add(id_parameter)                'oConn.Open()
        If CConnection.ConnOpen Then                        ' If connection opened okay
          oRead = oCmd.ExecuteReader                      ' Do query 
          While oRead.Read
            Dim objCVL As New ComponentView()
            objCVL.DataFields.readfields(oRead)
            Me.Add(objCVL)
          End While
        End If

      Catch ex As Exception
        Medscreen.LogError(ex, , "Loading components ")
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        CConnection.SetConnClosed()

      End Try
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a component to the list 
    ''' </summary>
    ''' <param name="item"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal item As ComponentView) As Integer
      Return MyBase.List.Add(item)
    End Function
#End Region

#Region "Procedures"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new list from the component view
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New("COMPONENT_VIEW")
      Me.SetupFields()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new component view collection for a customer
    ''' </summary>
    ''' <param name="AnalysisID">Analysis to get list for</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal AnalysisID As String, Optional ByVal TableName As String = "COMPONENT_VIEW")
      MyClass.New()
      Me.strCust = AnalysisID
      Me.Fields.Table = TableName
    End Sub
#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get componentview from list 
    ''' </summary>
    ''' <param name="index">Position to get </param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As ComponentView
      Get
        Return CType(MyBase.List.Item(index), ComponentView)
      End Get
    End Property


    Default Public Overloads ReadOnly Property Item(ByVal Index As String, ByVal Pos As Integer) As ComponentView
      Get
        Dim objCV As ComponentView
        Dim i As Integer
        For i = 0 To Count - 1      'Go through list
          objCV = Me.Item(i)
          If objCV.Analysis = Index AndAlso objCV.OrderNumber = Pos Then
            Exit For
          End If
          objCV = Nothing         'null out item
        Next

        'sse if we've found it if not create it and load 

        objCV = New ComponentView(Index, Pos)
        MyBase.List.Add(objCV)
        Return objCV
      End Get
    End Property
#End Region
#End Region
  End Class


#End Region

#Region "Discount_schemes"
#Region "Non Volume Schemes"

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.DiscountScheme
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A simple (non volume) discount scheme entry
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class DiscountScheme

#Region "Declarations"
    Private objDiscountSchemeId As IntegerField = New IntegerField("IDENTITY", -1, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 80)
    Private objModifiedOn As TimeStampField = New TimeStampField("MODIFIED_ON")
    Private objModifiedBy As UserField = New UserField("MODIFIED_BY")
    Private objLastusedOn As DateField = New DateField("LASTUSED_ON", DateField.ZeroDate)
    Private objLastInvoice As StringField = New StringField("LASTINVOICE", "", 20)

    Private myFields As TableFields = New TableFields("DISCOUNTSCHEME_HEADER")

    Private myEntries As DiscountSchemeEntries


#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert the Discount Entry to XML
    ''' </summary>
    ''' <returns>A string representing the Discount Scheme as XML</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
        Public Function ToXML() As String
            Dim strRet As String = CConnection.PackageStringList("Lib_customer.DiscountSchemeRowXML", Me.DiscountSchemeId)
            'Dim strRet As String = "<DiscountScheme>"
            'strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
            'strRet += "<Description>" & Medscreen.FixXML(Me.Description) & "</Description>"
            'strRet += Me.Entries.ToXML
            'strRet += "</DiscountScheme>"
            Return strRet
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the discount scheme by reloading from the database
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Dim blnRet As Boolean = False
            Me.Entries.Clear()
            Me.Entries.Load()
            Return blnRet


        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Calculate Discount for this scheme, for a given Line Item or item type
        ''' </summary>
        ''' <param name="ItemCode">Line Item Code to discount on</param>
        ''' <param name="ItemType">Item Type to discount on</param>
        ''' <returns>Discount</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	27/09/2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Discount(ByVal ItemCode As String, ByVal ItemType As String) As Double
            Dim tmpDisc As Double = 0.0

            Dim ent As DiscountSchemeEntry
            Dim i As Integer


            For i = 0 To Me.Entries.Count - 1
                ent = Me.Entries.Item(i)
                If ent.DiscountType = "CODE" And ent.ApplyTo = ItemCode Then
                    tmpDisc = ent.Discount
                    Exit For
                ElseIf ent.DiscountType = "GROUP" And _
                    ((Mid(ItemCode, 1, 4) = ent.ApplyTo) Or _
                    (Mid(ItemCode, 1, 3) = ent.ApplyTo)) Then
                    tmpDisc = ent.Discount
                    Exit For
                ElseIf ent.DiscountType = "TYPE" And ent.ApplyTo = ItemType Then
                    tmpDisc = ent.Discount
                    Exit For

                End If
            Next


            Return tmpDisc
        End Function

        Public Function Insert() As Boolean
            Return Me.Fields.Insert(MedConnection.Connection)
        End Function

        Public Function Update() As Boolean
            Return Me.Fields.Update(MedConnection.Connection)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Provide a full description of the discount scheme 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function FullDescription(Optional ByVal Customer As String = "") As String
            Dim strRet As String = Me.Description & vbCrLf
            ' Add the individual elements of the scheme
            If Entries.Count > 1 Then
                strRet += vbCrLf & "This scheme discounts more than one item" & vbCrLf
            End If
            If Customer.Trim.Length > 0 Then
                Dim ocoll As New Collection()
                ocoll.Add(Customer)
                ocoll.Add(Me.DiscountSchemeId)
                Dim strResp As String = CConnection.PackageStringList("pckwarehouse.getdiscounttext", ocoll)
                strRet += strResp
                Return strRet
                Exit Function
            End If
            strRet += Entries.Description
            If Entries.Count > 1 Then
                strRet += vbCrLf & "Please do not use it unless you are sure - if so then please check all other prices," & vbCrLf _
                & " if you select NO you will be given the chance of creating a scheme later." & vbCrLf
            End If
            Return strRet
        End Function
#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Discount Scheme Entry
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      SetUpFields()
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Description of the Discount Scheme
    ''' </summary>
    ''' <value>Description of the Discount Scheme</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Description() As String
      Get
        Return Me.objDescription.Value
      End Get
      Set(ByVal Value As String)
        objDescription.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A discount Scheme ID
    ''' </summary>
    ''' <value>A Discount Scheme ID Entry</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objDiscountSchemeId.Value
      End Get
      Set(ByVal Value As Integer)
        objDiscountSchemeId.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A list of Discount Scheme entries
    ''' </summary>
    ''' <value>A list of Discount Scheme Entries</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Entries() As DiscountSchemeEntries
      Get
        If myEntries Is Nothing AndAlso Me.DiscountSchemeId > 0 Then
          myEntries = New DiscountSchemeEntries(Me.DiscountSchemeId)

        End If
        Return myEntries
      End Get
      Set(ByVal Value As DiscountSchemeEntries)
        Me.myEntries = Value
      End Set
    End Property

#End Region
#End Region

    Friend Property Fields() As TableFields
      Get
        Return myFields
      End Get
      Set(ByVal Value As TableFields)
        myFields = Value
      End Set
    End Property

    Private Sub SetUpFields()
      myFields.Add(objDiscountSchemeId)
      myFields.Add(objDescription)
      myFields.Add(objModifiedOn)
      myFields.Add(objModifiedBy)
      myFields.Add(objLastusedOn)
      myFields.Add(objLastInvoice)
    End Sub

  End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.DiscountSchemes
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A collection of Discount Schemes
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class DiscountSchemes
    Inherits CollectionBase

#Region "Declarations"
    Private objDiscountSchemeId As IntegerField = New IntegerField("IDENTITY", -1, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 80)
    Private objModifiedOn As DateField = New DateField("MODIFIED_ON", DateField.ZeroDate)
    Private objModifiedBy As StringField = New StringField("MODIFIED_BY", "", 10)
    Private objLastusedOn As DateField = New DateField("LASTUSED_ON", DateField.ZeroDate)
    Private objLastInvoice As StringField = New StringField("LASTINVOICE", "", 20)

    Private myFields As TableFields = New TableFields("DISCOUNTSCHEME_HEADER")

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Loads a discount Scheme collection from the Database
    ''' </summary>
    ''' <returns>TRUE if Succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean

      Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.Connection = MedConnection.Connection
        oCmd.CommandText = myFields.FullRowSelect
        CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
        While oRead.Read
          Dim ds As DiscountScheme
          ds = Me.Item(oRead.GetValue(oRead.GetOrdinal("IDENTITY")), 1)
          If ds Is Nothing Then
            ds = New DiscountScheme()
            Me.Add(ds)
          End If
          ds.Fields.readfields(oRead)
        End While
      Catch ex As Exception
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        CConnection.SetConnClosed()
      End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a discount scheme to the collection
    ''' </summary>
    ''' <param name="Item">Discount Scheme to add</param>
    ''' <returns>Position in the list of the added item</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal Item As DiscountScheme) As Integer
      Return MyBase.List.Add(Item)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find a discount scheme
    ''' </summary>
    ''' <param name="Discount">Discount amount to find</param>
    ''' <param name="ItemCode">ItemCode to look for </param>
    ''' <param name="SampleType">Type of discount to look for</param>
    ''' <param name="From">Starting point of search</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
        Public Function FindDiscountScheme(ByVal Discount As Double, ByVal ItemCode As String, ByVal SampleType As String, ByRef From As Integer) As DiscountScheme
            Dim oColl As New Collection()
            If ItemCode.Trim.Length = 0 Then
                oColl.Add(SampleType)
            Else
                oColl.Add(ItemCode)
            End If
            oColl.Add(Discount.ToString)
            Dim strDisc As String = CConnection.PackageStringList("PCKWAREHOUSE.FindDiscountSchemes", oColl)
            Dim ds As DiscountScheme = Nothing
            If strDisc Is Nothing OrElse strDisc.Trim.Length = 0 Then 'No scheme found
                Return ds
            Else
                Dim strDiscArray As String() = strDisc.Split(New Char() {","})      'Split discount schemes 
                If From >= strDiscArray.Length - 1 Then                              'Run through all schemes
                    Return ds
                    Exit Function
                End If
                ds = Me.Item(strDiscArray(From), 1)                                 'Get selected scheme
            End If
            'Dim i As Integer
            'For i = From To Me.Count - 1
            '  ds = Me.Item(i)
            '  Dim tmpDisc As Double = ds.Discount(ItemCode, SampleType)
            '  If Math.Abs(Discount - tmpDisc) < 0.05 Then
            '    From = i
            '    Exit For
            '  End If
            '  ds = Nothing
            'Next
            Return ds
        End Function

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Construct a collection of Discount Schemes
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.New()
      SetUpFields()
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieve an Item from the collection of Discount Schemes
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="opt"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal index As Integer, Optional ByVal opt As Integer = 0) As DiscountScheme
            Get
                Dim dsr As DiscountScheme = Nothing
                If opt = 0 Then
                    dsr = CType(MyBase.List.Item(index), DiscountScheme)
                ElseIf opt = 1 Then
                    Dim i As Integer
                    Dim ds As DiscountScheme = Nothing
                    For i = 0 To Me.Count - 1
                        ds = Me.Item(i)
                        If ds.DiscountSchemeId = index Then
                            Exit For
                        End If
                        ds = Nothing
                    Next
                    dsr = ds
                End If
                Return dsr
            End Get
            Set(ByVal Value As DiscountScheme)
                MyBase.List.Item(index) = Value
            End Set
    End Property

#End Region
#End Region

    Private Sub SetUpFields()
      myFields.Add(objDiscountSchemeId)
      myFields.Add(objDescription)
      myFields.Add(objModifiedOn)
      myFields.Add(objModifiedBy)
      myFields.Add(objLastusedOn)
      myFields.Add(objLastInvoice)
    End Sub

  End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.DiscountSchemeEntry
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Discount Scheme Entry
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class DiscountSchemeEntry

#Region "Declarations"
    Private objDiscountSchemeId As IntegerField = New IntegerField("SCHEME_ID", -1, True)
    Private objEntryNumber As IntegerField = New IntegerField("ENTRYNUMBER", -1, True)
    Private objType As StringField = New StringField("TYPE", "", 10)
        Private objApplyTo As StringField = New StringField("APPLYTO", "", 20)
    Private objDiscount As DoubleField = New DoubleField("DISCOUNT", 0.0)

    Private myfields As TableFields = New TableFields("DISCOUNTSCHEME_ENTRY")

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert the Discount Scheme Entry to XML
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML() As String
      Dim strRet As String = "<DiscountSchemeEntry>"
      strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
      strRet += "<EntryNo>" & Me.EntryNumber & "</EntryNo>"
      strRet += "<Type>" & Me.DiscountType & "</Type>"
      strRet += "<ApplyTo>" & Me.ApplyTo & "</ApplyTo>"
      strRet += "<Discount>" & Me.Discount & "</Discount>"

      strRet += "</DiscountSchemeEntry>"
      Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a description for the scheme 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Description() As String
      Dim strRet As String = "Scheme ID: " & Me.DiscountSchemeId & " "
      If Me.DiscountType = "CODE" Then
        Dim lineitem As LineItem = Glossary.LineItems(Me.ApplyTo)
        If Not lineitem Is Nothing Then
          strRet += Me.Discount & "% discount on " & lineitem.Service & " (" & Me.ApplyTo & ") " & lineitem.Description
          If lineitem.Removed Then strRet += " (Obsolete LineItem)"
        Else
          strRet += Me.Discount & "% discount on " & Me.ApplyTo
        End If
      Else
        strRet += Me.Discount & "% discount on " & Me.ApplyTo
      End If
      strRet += vbCrLf
      Return strRet
    End Function

    Public Function Insert() As Boolean
      Return Me.Fields.Insert(MedConnection.Connection)
    End Function

    Public Function Update() As Boolean
      Return Me.Fields.Update(MedConnection.Connection)
    End Function
#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Discount Scheme Entry
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      Setupfields()
    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' What the scheme applies to typically a line item but may be a line item group
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ApplyTo() As String
      Get
        Return objApplyTo.Value
      End Get
      Set(ByVal Value As String)
        objApplyTo.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Discount to apply
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Discount() As Double
      Get
        Return Me.objDiscount.Value
      End Get
      Set(ByVal Value As Double)
        objDiscount.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Discount Scheme ID
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objDiscountSchemeId.Value
      End Get
      Set(ByVal Value As Integer)
        objDiscountSchemeId.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of Discount 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountType() As String
      Get
        Return objType.Value
      End Get
      Set(ByVal Value As String)
        objType.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The order of this entry, lower numbers take priority over higher
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property EntryNumber() As Integer
      Get
        Return objEntryNumber.Value
      End Get
      Set(ByVal Value As Integer)
        objEntryNumber.Value = Value
      End Set
    End Property

#End Region
#End Region


    Friend Property Fields() As TableFields
      Get
        Return myfields
      End Get
      Set(ByVal Value As TableFields)
        myfields = Value
      End Set
    End Property

    Private Sub Setupfields()
      myfields.Add(objDiscountSchemeId)
      myfields.Add(objEntryNumber)
      myfields.Add(objType)
      myfields.Add(objApplyTo)
      myfields.Add(objDiscount)
    End Sub

  End Class


  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.DiscountSchemeEntries
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A collection of Discount Scheme Entries 
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[taylor]	27/09/2005	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Class DiscountSchemeEntries
    Inherits CollectionBase


#Region "Declarations"
    Private objDiscountSchemeId As IntegerField = New IntegerField("SCHEME_ID", -1, True)
    Private objEntryNumber As IntegerField = New IntegerField("ENTRYNUMBER", -1, True)
    Private objType As StringField = New StringField("TYPE", "", 10)
        Private objApplyTo As StringField = New StringField("APPLYTO", "", 20)
    Private objDiscount As DoubleField = New DoubleField("DISCOUNT", 0.0)

    Private myfields As TableFields = New TableFields("DISCOUNTSCHEME_ENTRY")

#End Region

#Region "Public enumerations"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Types of Discount Scheme Entry
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Enum DiscountSchemeEntryType
      ''' <summary>A line Item Code</summary>
      CODE
      ''' <summary>A line Item Group</summary>
      TYPE
    End Enum
#End Region


#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return a description of the elements  of the scheme 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Description() As String
            Dim strRet As String = ""
      Dim i As Integer
      Dim dse As DiscountSchemeEntry

      For i = 0 To Count - 1
        dse = Item(i)
        strRet += dse.Description
      Next
      Return strRet
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a discount scheme entry to the collection
    ''' </summary>
    ''' <param name="Item">Entry to add</param>
    ''' <returns>Position in list</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal Item As DiscountSchemeEntry) As Integer
      Return MyBase.List.Add(Item)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load the collection
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean

      Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.Connection = MedConnection.Connection
        oCmd.CommandText = myfields.FullRowSelect & " where SCHEME_ID = ? order by entrynumber,type"
        oCmd.Parameters.Add(CConnection.StringParameter("SchemeID", Me.DiscountSchemeId, 10))
        CConnection.SetConnOpen()
        oRead = oCmd.ExecuteReader
        While oRead.Read
          Dim ds As DiscountSchemeEntry = New DiscountSchemeEntry()
          ds.Fields.readfields(oRead)
          Me.Add(ds)
        End While
      Catch ex As Exception
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        CConnection.SetConnClosed()
      End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert Collection to XML
    ''' </summary>
    ''' <returns>Collection represented as an XML string</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML() As String
      Dim strRet As String = "<DiscountSchemeEntries>"
      Dim i As Integer
      Dim dse As DiscountSchemeEntry

      For i = 0 To Count - 1
        dse = Item(i)
        strRet += dse.ToXML & vbCrLf
      Next
      strRet += "</DiscountSchemeEntries>" & vbCrLf

      Return strRet
    End Function

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Discount scheme entry collection and load it
    ''' </summary>
    ''' <param name="SchemeId">Scheme Id for this collection</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal SchemeId As Integer)
      MyBase.New()
      Setupfields()
      Me.DiscountSchemeId = SchemeId
      Me.Load()

    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Id of the Discount Scheme in use
    ''' </summary>
    ''' <value>Discount Scheme ID</value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objDiscountSchemeId.Value
      End Get
      Set(ByVal Value As Integer)
        objDiscountSchemeId.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves the Discount Scheme Entry by position
    ''' </summary>
    ''' <param name="Index">Position in the Scheme</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal Index As Integer) As DiscountSchemeEntry
      Get
        Return CType(MyBase.List.Item(Index), DiscountSchemeEntry)
      End Get
      Set(ByVal Value As DiscountSchemeEntry)
        MyBase.List.Item(Index) = Value
      End Set
    End Property

#End Region
#End Region

    Friend Property Fields() As TableFields
      Get
        Return myfields
      End Get
      Set(ByVal Value As TableFields)
        myfields = Value
      End Set
    End Property

    Private Sub Setupfields()
      myfields.Add(objDiscountSchemeId)
      myfields.Add(objEntryNumber)
      myfields.Add(objType)
      myfields.Add(objApplyTo)
      myfields.Add(objDiscount)
    End Sub

  End Class

#End Region

#Region "Volume Schemes"

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.VolumeDiscountEntry
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A volume Discount Scheme Object
  ''' </summary>
  ''' <remarks>
  ''' Volume discount schemes implement discount banding, each entry is a band in the discount scheme.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class VolumeDiscountEntry

#Region "Declarations"
    Private myFields As New TableFields("VOLUMEDISCOUNT_ENTRY")
    Private objSchemeID As IntegerField = New IntegerField("SCHEME_ID", 0, True)
    Private objEntryNumber As IntegerField = New IntegerField("ENTRYNUMBER", 0)
    Private objMaxSamples As IntegerField = New IntegerField("MAX_SAMPLES", 0)
    Private objMinSamples As IntegerField = New IntegerField("MIN_SAMPLES", 0)
    Private objType As StringField = New StringField("TYPE", "", 10)
    Private objApplyTO As StringField = New StringField("APPLYTO", "", 30)
    Private objSampleDiscount As DoubleField = New DoubleField("SAMPLE_DISCOUNT", 0.0)
    Private objMinchargeDiscount As DoubleField = New DoubleField("MINCHARGE_DISCOUNT", 0.0)
    Private objAdditionalCharge As DoubleField = New DoubleField("ADDITIONAL_CHARGE", 0.0)

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert the Volume Discount Entry to XML
    ''' </summary>
    ''' <param name="CurrencyCode">Which Currency Code are we using (For Example)</param>
    ''' <returns>A string containing XML representation of object</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXml(ByVal CurrencyCode As String) As String
      Dim strRet As String = "<VolumeSchemeEntry>" & vbCrLf
      strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
      strRet += "<ApplyTo>" & Me.ApplyTo & "</ApplyTo>"
      strRet += "<EntryNo>" & Me.EntryNumber & "</EntryNo>"
      strRet += "<MinSamples>" & Me.MinSamples & "</MinSamples>"
      strRet += "<MaxSamples>" & Me.MaxSamples & "</MaxSamples>"
      strRet += "<MinChargeDiscount>" & Me.MinChargeDiscount.ToString("0.00") & "</MinChargeDiscount>"
      strRet += "<SampleDiscount>" & Me.SampleDiscount.ToString("0.00") & "</SampleDiscount>"
      strRet += "<AdditionalCharge>" & Medscreen.CurrencyCodeHTML(CurrencyCode) & Me.AdditionalCharge.ToString("0.00") & "</AdditionalCharge>"
      strRet += "<Type>" & Me.Type & "</Type>"

      strRet += "</VolumeSchemeEntry>" & vbCrLf
      Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Give an example of the discount scheme
    ''' </summary>
    ''' <param name="CurrencyCode">Currency that example will be in </param>
    ''' <returns>Returns an Example of teh discount scheme</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXmlExample(ByVal CurrencyCode As String, ByVal CustomerID As String) As String
      Dim strRet As String = "<VolumeSchemeEntry>" & vbCrLf
      Dim objPrice As LineItemPrice = MedscreenLib.Glossary.Glossary.PriceList.Item("APCOLS002", Now, CurrencyCode)
      Dim MinCharge As LineItemPrice = MedscreenLib.Glossary.Glossary.PriceList.Item("APCOLS007", Now, CurrencyCode)
            Dim minChargeDiscount As Object = Nothing
            Dim SampleDiscount As Object = Nothing

      Dim ocmd As New OleDb.OleDbCommand()
      ocmd.Connection = MedConnection.Connection
      Try
        CConnection.SetConnOpen()
        ocmd.CommandText = "Select pckwarehouse.getdiscount(?,?) from dual"
        ocmd.Parameters.Add(CConnection.StringParameter("CustId", CustomerID, 10))
        ocmd.Parameters.Add(CConnection.StringParameter("ItemCode", "APCOLS007", 10))
        minChargeDiscount = ocmd.ExecuteScalar
        If minChargeDiscount Is Nothing OrElse minChargeDiscount Is DBNull.Value Then
          minChargeDiscount = 1
        End If
        ocmd.Parameters(1).Value = "APCOLS002"
        SampleDiscount = ocmd.ExecuteScalar
        If SampleDiscount Is Nothing OrElse SampleDiscount Is DBNull.Value Then
          SampleDiscount = 1
        End If
      Catch ex As Exception
      Finally
        CConnection.SetConnClosed()
      End Try

            If objPrice Is Nothing Then Return "" : Exit Function

      Dim Price As Double = objPrice.Price * SampleDiscount
      strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
      strRet += "<EntryNo>" & Me.EntryNumber & "</EntryNo>"
      Dim ActualPrice As Double = Me.AdditionalCharge
      Dim ActMinCharge As Double = (MinCharge.Price * minChargeDiscount) * (1 - Me.MinChargeDiscount / 100) + Me.AdditionalCharge
      Dim actualPriceMin As Double = ActualPrice + (Price * (1 - Me.SampleDiscount / 100)) * Me.MinSamples
      Dim actualPriceMax As Double = ActualPrice + (Price * (1 - Me.SampleDiscount / 100)) * Me.MaxSamples

      If actualPriceMin < ActMinCharge Then actualPriceMin = ActMinCharge
      If actualPriceMax < ActMinCharge Then actualPriceMax = ActMinCharge

      strRet += "<Example>"

      Dim strSamples As String
      If Me.Type = "COUNTRY" Then         'Country based discount scheme
        Dim objCountry As MedscreenLib.Address.Country = MedscreenLib.Glossary.Glossary.Countries.Item(CInt(Me.ApplyTo), 0)
        strSamples = "Charge for samples in "
        If objCountry Is Nothing Then
          strSamples += Me.ApplyTo
        Else
          strSamples += objCountry.CountryName
        End If
        strRet += strSamples & " is " & Medscreen.CurrencyCodeHTML(CurrencyCode) & (Price * (1 - Me.SampleDiscount / 100)).ToString("0.00") & " per sample " & actualPriceMin & " Charged."
      ElseIf Me.Type = "COLL_TYPE" Then
        strSamples = "Charge for samples in " & Me.ApplyTo & " collections is "
        strRet += strSamples & " is " & Medscreen.CurrencyCodeHTML(CurrencyCode) & (Price * (1 - Me.SampleDiscount / 100)).ToString("0.00") & " per sample"
      ElseIf Me.Type = "PORT" Then
        strSamples = "Charge for samples collected at " & Me.ApplyTo & " is "
        strRet += strSamples & Medscreen.CurrencyCodeHTML(CurrencyCode) & (Price * (1 - Me.SampleDiscount / 100)).ToString("0.00") & " per sample"
      ElseIf Me.Type = "VESSEL" Then
        ocmd = New OleDb.OleDbCommand("Select name from vessel where vessel_id = '" & Me.ApplyTo.Trim & "'", MedConnection.Connection)
        Dim strVName As String = Me.ApplyTo
        Try
          If CConnection.ConnOpen Then
            strVName = ocmd.ExecuteScalar
          End If
        Catch
        Finally
          CConnection.SetConnClosed()
        End Try
        strSamples = "Charge for samples collected from " & strVName & " is "
        strRet += strSamples & Medscreen.CurrencyCodeHTML(CurrencyCode) & (Price * (1 - Me.SampleDiscount / 100)).ToString("0.00") & " per sample"
      Else
        If (Me.MaxSamples = Me.MinSamples) Or (Me.MaxSamples = 0) Then
          strSamples = "Charge for " & Me.MinSamples & " samples is "
        Else
          strSamples = "Charge for between " & Me.MinSamples & " and " & Me.MaxSamples & " samples is "
        End If
        If Me.MaxSamples > 0 Then
          strRet += strSamples & _
                Medscreen.CurrencyCodeHTML(CurrencyCode) & actualPriceMin.ToString("0.00")
          If Me.MaxSamples <> Me.MinSamples Then
            strRet += " - " & Medscreen.CurrencyCodeHTML(CurrencyCode) & actualPriceMax.ToString("0.00")
          End If
        Else
          strRet += strSamples & _
              Medscreen.CurrencyCodeHTML(CurrencyCode) & actualPriceMin.ToString("0.00") & " -  and then " & _
              Medscreen.CurrencyCodeHTML(CurrencyCode) & (Price * (1 - Me.SampleDiscount / 100)).ToString("0.00") & " per sample"
        End If
      End If

      strRet += "</Example>"

      strRet += "</VolumeSchemeEntry>" & vbCrLf
      Return strRet
    End Function

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Volume Discount Scheme Entity
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      SetupFields()
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Maximum number of samples for this discount band
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property MaxSamples() As Integer
      Get
        Return Me.objMaxSamples.Value
      End Get
      Set(ByVal Value As Integer)
        objMaxSamples.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Minimum no of samples for this discount band
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property MinSamples() As Integer
      Get
        Return Me.objMinSamples.Value
      End Get
      Set(ByVal Value As Integer)
        objMinSamples.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Minimum Charge Discount
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property MinChargeDiscount() As Double
      Get
        Return Me.objMinchargeDiscount.Value
      End Get
      Set(ByVal Value As Double)
        objMinchargeDiscount.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Discount per sample
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SampleDiscount() As Double
      Get
        Return Me.objSampleDiscount.Value
      End Get
      Set(ByVal Value As Double)
        Me.objSampleDiscount.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of Discount scheme
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Type() As String
      Get
        Return Me.objType.Value
      End Get
      Set(ByVal Value As String)
        objType.Value = Value
      End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Additional charge for this discount band
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AdditionalCharge() As Double
      Get
        Return Me.objAdditionalCharge.Value
      End Get
      Set(ByVal Value As Double)
        objAdditionalCharge.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' What this band applies to 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ApplyTo() As String
      Get
        Return Me.objApplyTO.Value
      End Get
      Set(ByVal Value As String)
        objApplyTO.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The scheme ID this band belongs to 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objSchemeID.Value
      End Get
      Set(ByVal Value As Integer)
        objSchemeID.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Band ID 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property EntryNumber() As Integer
      Get
        Return objEntryNumber.Value
      End Get
      Set(ByVal Value As Integer)
        objEntryNumber.Value = Value
      End Set
    End Property

#End Region
#End Region

    Friend Property Fields() As TableFields
      Get
        Return myFields
      End Get
      Set(ByVal Value As TableFields)
        myFields = Value
      End Set
    End Property

    Private Sub SetupFields()
      myFields.Add(objSchemeID)
      myFields.Add(objEntryNumber)
      myFields.Add(objMaxSamples)
      myFields.Add(objMinSamples)
      myFields.Add(objType)
      myFields.Add(objApplyTO)
      myFields.Add(objSampleDiscount)
      myFields.Add(objMinchargeDiscount)
      myFields.Add(objAdditionalCharge)

    End Sub

  End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.VolumeDiscountEntries
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A collection of Volume Discount Bands
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class VolumeDiscountEntries
    Inherits CollectionBase

#Region "Declarations"
    Private myFields As New TableFields("VOLUMEDISCOUNT_ENTRY")
    Private objSchemeID As IntegerField = New IntegerField("SCHEME_ID", 0, True)
    Private objEntryNumber As IntegerField = New IntegerField("ENTRYNUMBER", 0)
    Private objMaxSamples As IntegerField = New IntegerField("MAX_SAMPLES", 0)
    Private objMinSamples As IntegerField = New IntegerField("MIN_SAMPLES", 0)
    Private objType As StringField = New StringField("TYPE", "", 10)
    Private objApplyTO As StringField = New StringField("APPLYTO", "", 30)
    Private objSampleDiscount As DoubleField = New DoubleField("SAMPLE_DISCOUNT", 0.0)
    Private objMinchargeDiscount As DoubleField = New DoubleField("MINCHARGE_DISCOUNT", 0.0)
    Private objAdditionalCharge As DoubleField = New DoubleField("ADDITIONAL_CHARGE", 0.0)

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a new Scheme Entry to this collection 
    ''' </summary>
    ''' <param name="item">Scheme Entry to add</param>
    ''' <returns>Position of Added item in list </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal item As VolumeDiscountEntry) As Integer
      Return MyBase.List.Add(item)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load this collection of Discount Bands
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.Connection = MedConnection.Connection
        oCmd.CommandText = myFields.FullRowSelect & " where SCHEME_ID = ? order by entrynumber,type"
        oCmd.Parameters.Add(CConnection.StringParameter("SchemeID", Me.DiscountSchemeId, 10))

        CConnection.SetConnOpen()

        oRead = oCmd.ExecuteReader
        While oRead.Read
          Dim oVDH As VolumeDiscountEntry = New VolumeDiscountEntry()
          oVDH.Fields.readfields(oRead)
          Me.Add(oVDH)
        End While
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        CConnection.SetConnClosed()
      End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find out if line item type is in scheme 
    ''' </summary>
    ''' <param name="_type"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function TypeInString(ByVal _type As String) As Boolean
      Dim i As Integer
      Dim ds As VolumeDiscountEntry
      Dim blnReturn As Boolean = False

      For i = 0 To Count - 1
        ds = Item(i)
        If ds.Type = _type Then
          blnReturn = True
          Exit For
        End If

      Next
      Return blnReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert entire Scheme into XML, with examples as to the scheme
    ''' </summary>
    ''' <param name="CurrencyCode">Currency for Example</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXml(ByVal CurrencyCode As String) As String
      Dim strRet As String = "<VolumeDiscountEntries>"

      Dim i As Integer
      Dim ds As VolumeDiscountEntry

      For i = 0 To Count - 1
        ds = Item(i)
        strRet += ds.ToXml(CurrencyCode) & vbCrLf
      Next

      strRet += "</VolumeDiscountEntries>" & vbCrLf
      Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert Scheme to an Example of its action in XML
    ''' </summary>
    ''' <param name="CurrencyCode">Currency for Example</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXmlExample(ByVal CurrencyCode As String, ByVal CustomerID As String) As String
            'Dim strRet As String = "<VolumeDiscountEntries>"

            'Dim i As Integer
            'Dim ds As VolumeDiscountEntry

            'For i = 0 To Count - 1
            '  ds = Item(i)
            '  strRet += ds.ToXmlExample(CurrencyCode, CustomerID) & vbCrLf
            'Next

            'strRet += "</VolumeDiscountEntries>" & vbCrLf
            'Return strRet
            Dim oColl As New Collection()
            oColl.Add(CustomerID)
            oColl.Add(Me.DiscountSchemeId)
            Dim strRet As String = CConnection.PackageStringList("lib_customer.XMLExample", oColl)
            Return strRet
    End Function

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new collection of Banded Discount Scheme entries for a particular Scheme
    ''' </summary>
    ''' <param name="SchemeId">The Scheme Id for this collection</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal SchemeId As Integer)
      MyBase.New()
      SetupFields()
      Me.DiscountSchemeId = SchemeId
      Load()
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scheme Id for the collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objSchemeID.Value
      End Get
      Set(ByVal Value As Integer)
        objSchemeID.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieve an Item from the collection
    ''' </summary>
    ''' <param name="index">Index by position for the item to retrieve</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal index As Integer) As VolumeDiscountEntry
      Get
        Return CType(MyBase.List.Item(index), VolumeDiscountEntry)
      End Get
      Set(ByVal Value As VolumeDiscountEntry)
        MyBase.List.Item(index) = Value
      End Set
    End Property

#End Region
#End Region

    Private Sub SetupFields()
      myFields.Add(objSchemeID)
      myFields.Add(objEntryNumber)
      myFields.Add(objMaxSamples)
      myFields.Add(objMinSamples)
      myFields.Add(objType)
      myFields.Add(objApplyTO)
      myFields.Add(objSampleDiscount)
      myFields.Add(objMinchargeDiscount)
      myFields.Add(objAdditionalCharge)

    End Sub

  End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.VolumeDiscountHeader
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Scheme description for Volume (banded) scheme
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class VolumeDiscountHeader

#Region "Declarations"

    Private objIdentity As IntegerField = New IntegerField("IDENTITY", 1, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 80)
    Private objApplyAllEntries As BooleanField = New BooleanField("APPLY_ALL_ENTRIES", "?")
    Private objModifiedON As DateField = New DateField("MODIFIED_ON", DateField.ZeroDate)
    Private objModifiedBY As StringField = New StringField("MODIFIED_BY", "", 10)
    Private objLastusedON As DateField = New DateField("LASTUSED_ON", DateField.ZeroDate)
    Private objLastInvoice As StringField = New StringField("LASTINVOICE", "", 20)

    Private myFields As TableFields = New TableFields("VOLUMEDISCOUNT_HEADER")

    Private myEntries As VolumeDiscountEntries

#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert Scheme to XML calls Entries.ToXml
    ''' </summary>
    ''' <param name="CurrencyCode"></param>
    ''' <returns>An XML representation of the scheme</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXml(ByVal CurrencyCode As String) As String
      Dim strRet As String = "<VolumeDiscountHeader>"
      strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
      strRet += "<Description>" & Me.Description & "</Description>"
      strRet += "<ApplyAll>" & Me.ApplyAllEntries & "</ApplyAll>"
      strRet += Me.Entries.ToXml(CurrencyCode)

      strRet += "</VolumeDiscountHeader>" & vbCrLf

      Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find out if line item type is in scheme
    ''' </summary>
    ''' <param name="_type"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function TypeInScheme(ByVal _type As String) As Boolean
      Dim blnReturn As Boolean = False
      If Not Me.Entries Is Nothing Then
        blnReturn = Me.Entries.TypeInString(_type)
      End If
      Return blnReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Produce an example of this entire Scheme in XML format
    ''' </summary>
    ''' <param name="CurrencyCode"></param>
    ''' <returns>An Example of the scheme in XML</returns>
    ''' <remarks>
    '''     The code calls the Entries.ToXmlExample method for each entry in this scheme. 
        ''' <see cref="VolumeDiscountEntries.ToXmlExample"/>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXmlExample(ByVal CurrencyCode As String, ByVal CustomerID As String) As String
            Dim strRet As String = "<VolumeDiscountHeader>"
            strRet += "<SchemeID>" & Me.DiscountSchemeId & "</SchemeID>"
            strRet += "<Description>" & Me.Description & "</Description>"
            strRet += "<ApplyAll>" & Me.ApplyAllEntries & "</ApplyAll>"
      strRet += Me.Entries.ToXmlExample(CurrencyCode, CustomerID)
            strRet += "</VolumeDiscountHeader>" & vbCrLf

      Return strRet
    End Function

        ''' <developer>CONCATENO\Taylor</developer>
        ''' <summary>
        ''' Load scheme
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>19-Dec-2011 15:12</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Public Function Load() As Boolean
            Dim blnRet As Boolean = True
            Try
                Me.Fields.Load(CConnection.DbConnection)

            Catch ex As Exception
                blnRet = False
            End Try
            Return blnRet
        End Function

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Construct a Volume Discount Header Entity
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      SetupFields()
    End Sub


#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Apply all the Entries in the Scheme
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ApplyAllEntries() As Boolean
      Get
        Return Me.objApplyAllEntries.Value
      End Get
      Set(ByVal Value As Boolean)
        Me.objApplyAllEntries.Value = Value
      End Set
    End Property


        Private myCustomerId As String
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>07-Dec-2011 09:01</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Public Property Customerid() As String
            Get
                Return myCustomerId
            End Get
            Set(ByVal value As String)
                myCustomerId = value
            End Set
        End Property




        Private myFullDescription As String = ""
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>07-Dec-2011 08:53</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Public Property FullDescription() As String
            Get
                If myCustomerId.Trim.Length = 0 Then                  'No customer return description 
                    Return Description
                Else

                    If myFullDescription.Trim.Length = 0 Then
                        Dim ocoll As New Collection
                        ocoll.Add(Me.DiscountSchemeId)
                        ocoll.Add(Me.Customerid)
                        Dim strRet As String = CConnection.PackageStringList("pckwarehouse.CustVolDiscSchemeDescription", ocoll)
                        If strRet Is Nothing Then
                            Return Description
                        Else
                            myFullDescription = strRet
                            Return myFullDescription
                        End If
                    End If
                    Return myFullDescription
                    End If
            End Get
            Set(ByVal value As String)

            End Set
        End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Description of this Scheme, primarily for the GUI
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Description() As String
      Get
        Return Me.objDescription.Value
      End Get
      Set(ByVal Value As String)
        objDescription.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
        ''' A collection of Discount Scheme Entries <see cref="VolumeDiscountEntries"/>
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Entries() As VolumeDiscountEntries
      Get
        If myEntries Is Nothing AndAlso Me.DiscountSchemeId > 0 Then
          myEntries = New VolumeDiscountEntries(Me.DiscountSchemeId)

        End If
        Return myEntries
      End Get
      Set(ByVal Value As VolumeDiscountEntries)
        Me.myEntries = Value
      End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scheme ID for the this Discount Scheme
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Identity() As Integer
      Get
        Return objIdentity.Value
      End Get
      Set(ByVal Value As Integer)
        objIdentity.Value = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scheme ID (Property Identity is a mirror of this property)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property DiscountSchemeId() As Integer
      Get
        Return Me.objIdentity.Value
      End Get
      Set(ByVal Value As Integer)
        objIdentity.Value = Value
      End Set
    End Property


#End Region
#End Region

    Friend Property Fields() As TableFields
      Get
        Return myFields
      End Get
      Set(ByVal Value As TableFields)
        myFields = Value
      End Set
    End Property

    Private Sub SetupFields()
      myFields.Add(Me.objIdentity)
      myFields.Add(Me.objDescription)
      myFields.Add(Me.objApplyAllEntries)
      myFields.Add(Me.objModifiedON)
      myFields.Add(Me.objModifiedBY)
      myFields.Add(Me.objLastusedON)
      myFields.Add(Me.objLastInvoice)

    End Sub

    End Class


    Public Class CustomerVolumeDiscountHeaders
        Inherits VolumeDiscountHeaders
#Region "Declarations"
        Private myCustomerId As String = ""

#End Region

#Region "Public Instance"

#Region "Functions"
        'Public Overloads Function load() As Boolean
        '    Dim oCmd As New OleDb.OleDbCommand()
        '    Dim oRead As OleDb.OleDbDataReader = Nothing

        '    Try
        '        oCmd.Connection = MedConnection.Connection
        '        oCmd.CommandText = "Select * from volumediscount_headers"
        '        CConnection.SetConnOpen()

        '        oRead = oCmd.ExecuteReader
        '        While oRead.Read
        '            Dim oVDH As VolumeDiscountHeader = New VolumeDiscountHeader()
        '            oVDH.Fields.readfields(oRead)
        '            Me.Add(oVDH)
        '        End While
        '    Catch ex As Exception
        '        Medscreen.LogError(ex)
        '    Finally
        '        If Not oRead Is Nothing Then
        '            If Not oRead.IsClosed Then oRead.Close()
        '        End If
        '        CConnection.SetConnClosed()
        '    End Try


        'End Function
#End Region

#Region "Procedures"
        Public Sub New(ByVal CustomerID As String)
            MyBase.New()
            myCustomerId = CustomerID
            Load()
            For Each vd As VolumeDiscountHeader In MyBase.InnerList
                vd.Customerid = myCustomerId
            Next
        End Sub
#End Region

#Region "Properties"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>07-Dec-2011 09:35</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Public Property CustomerID() As String
            Get
                Return myCustomerId
                For Each vd As VolumeDiscountHeader In MyBase.List
                    vd.Customerid = myCustomerId
                Next

            End Get
            Set(ByVal value As String)
                myCustomerId = value
            End Set
        End Property
#End Region
#End Region


    End Class

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.VolumeDiscountHeaders
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' A collection of Volume Discount schemes
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class VolumeDiscountHeaders
    Inherits CollectionBase

#Region "Declarations"
    Private objIdentity As IntegerField = New IntegerField("IDENTITY", 1, True)
    Private objDescription As StringField = New StringField("DESCRIPTION", "", 80)
    Private objApplyAllEntries As BooleanField = New BooleanField("APPLY_ALL_ENTRIES", "?")
    Private objModifiedON As DateField = New DateField("MODIFIED_ON", DateField.ZeroDate)
    Private objModifiedBY As StringField = New StringField("MODIFIED_BY", "", 10)
    Private objLastusedON As DateField = New DateField("LASTUSED_ON", DateField.ZeroDate)
    Private objLastInvoice As StringField = New StringField("LASTINVOICE", "", 20)

    Private myFields As TableFields = New TableFields("VOLUMEDISCOUNT_HEADER")

#End Region

#Region "Public Enumerations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enumeration for the indexing of the Item code
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum IndexBy
      ''' -----------------------------------------------------------------------------
      ''' <summary>
      ''' Index on the scheme ID
      ''' </summary>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      SchemeID = 1
    End Enum
#End Region

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a new Discount Scheme to the collection 
    ''' </summary>
    ''' <param name="item">Discount Scheme to Add</param>
    ''' <returns>Position of Scheme in list</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal item As VolumeDiscountHeader) As Integer
      Return MyBase.List.Add(item)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load collection of schemes from database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean
      Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing

      Try
        oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = myFields.FullRowSelect
        CConnection.SetConnOpen()

        oRead = oCmd.ExecuteReader
        While oRead.Read
          Dim oVDH As VolumeDiscountHeader = New VolumeDiscountHeader()
          oVDH.Fields.readfields(oRead)
          Me.Add(oVDH)
        End While
      Catch ex As Exception
        Medscreen.LogError(ex)
      Finally
        If Not oRead Is Nothing Then
          If Not oRead.IsClosed Then oRead.Close()
        End If
        CConnection.SetConnClosed()
      End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert all schemes into XML
    ''' </summary>
    ''' <param name="CurrencyCode">Currency for examples produced</param>
    ''' <returns>An XML version of all schemes</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML(ByVal CurrencyCode As String) As String
      Dim strRet As String = "<VolumeDiscountHeaders>"

      Dim i As Integer
      Dim ds As VolumeDiscountHeader

      For i = 0 To Count - 1
        ds = Item(i)
        strRet += ds.ToXml(CurrencyCode) & vbCrLf
      Next

      strRet += "</VolumeDiscountHeaders>"
      Return strRet
    End Function
#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new collection of Volume discount 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      MyBase.new()
      SetupFields()
    End Sub

#End Region

#Region "Properties"






    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves an Volume Discount Scheme from the collection
    ''' </summary>
    ''' <param name="index">Index by position</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public Overloads Property Item(ByVal index As Integer) As VolumeDiscountHeader
      Get
        Return CType(MyBase.List.Item(index), VolumeDiscountHeader)
      End Get
      Set(ByVal Value As VolumeDiscountHeader)
        MyBase.List.Item(index) = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves an Volume Discount Scheme from the collection index by a particular item
    ''' </summary>
    ''' <param name="index">Item to index by</param>
    ''' <param name="opt">Indexing method</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Changed option to use enumeration</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public Overloads Property Item(ByVal index As Integer, ByVal opt As IndexBy) As VolumeDiscountHeader
      Get
                Dim oVDH As VolumeDiscountHeader = Nothing
        For Each oVDH In MyBase.List
          If opt = IndexBy.SchemeID Then
            If oVDH.Identity = index Then Exit For
          End If
          oVDH = Nothing
        Next
        Return oVDH
      End Get
      Set(ByVal Value As VolumeDiscountHeader)
        MyBase.List.Item(index) = Value
      End Set
    End Property

#End Region

#End Region

    Private Sub SetupFields()
      myFields.Add(Me.objIdentity)
      myFields.Add(Me.objDescription)
      myFields.Add(Me.objApplyAllEntries)
      myFields.Add(Me.objModifiedON)
      myFields.Add(Me.objModifiedBY)
      myFields.Add(Me.objLastusedON)
      myFields.Add(Me.objLastInvoice)

    End Sub


  End Class
#End Region
#End Region


#Region "Interval"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.Interval
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' SM Equivalent of TimeSpan
  ''' </summary>
  ''' <remarks>
  ''' Wraps around TimeSpan structure to allow conversion to / from SM Interval
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [25/05/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class Interval
    Private _timeSpan As TimeSpan

#Region "Shared"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an Interval based on the passed string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   An exception will be thrown if the format is not recognised.<p/>
    '''   Use IsInterval to determine if a value can be converted to an interval.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Parse(ByVal value As String) As Interval
      value = value.Trim
      ' Sample Manager format will allow "-  1 00", but TimeSpan expects to see "-1 00" so remove extra spaces
      If value.StartsWith("-") Then value = "-" & value.Substring(1).Trim
      ' Add time value if required
      Dim pos As Integer = value.LastIndexOf(" ")
      If pos = -1 Then
        Return New Interval(TimeSpan.Parse(value & ".0:0"))
      Else
        If value.IndexOf(":") = -1 Then value = value & ":0"
        Return New Interval(TimeSpan.Parse(value.Substring(0, pos).Replace(" ", "") & "." & value.Substring(pos + 1)))
      End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a value indicating whether the passed string is a valid SM interval
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    '''   Intervals are in the form "   d HH:mm:ss.ff" to "dddd HH:mm:ss.ff", leading 
    '''   spaces are optional.<p/>
    '''   Note that when writing to a field, the value must be right justified.
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function IsInterval(ByVal value As String) As Boolean
      Dim invalid As Boolean
      Try
        Interval.Parse(value)
      Catch
        invalid = True
      End Try
      Return Not invalid
    End Function

    Public Shared Function FromDays(ByVal days As Integer) As Interval
      Return New Interval(TimeSpan.FromDays(days))
    End Function

    Public Shared Function FromHours(ByVal hours As Integer) As Interval
      Return New Interval(TimeSpan.FromHours(hours))
    End Function

    Public Shared Function FromMinutes(ByVal minutes As Integer) As Interval
      Return New Interval(TimeSpan.FromMinutes(minutes))
    End Function

    Public Shared Function FromSeconds(ByVal seconds As Integer) As Interval
      Return New Interval(TimeSpan.FromSeconds(seconds))
    End Function

    Public Shared Function FromTicks(ByVal ticks As Long) As Interval
      Return New Interval(TimeSpan.FromTicks(ticks))
    End Function

    Public Shared Function FromMilliSeconds(ByVal milliSeconds As Integer) As Interval
      Return New Interval(TimeSpan.FromMilliseconds(milliSeconds))
    End Function

    Public Shared Function FromTimespan(ByVal span As TimeSpan) As Interval
      Return New Interval(span)
    End Function
#End Region

#Region "Instance"
    Public Sub New(ByVal span As TimeSpan)
      _timeSpan = span
    End Sub

    Public Sub New(ByVal days As Integer, ByVal hours As Integer, ByVal minutes As Integer, ByVal seconds As Integer, ByVal milliseconds As Integer)
      _timeSpan = New TimeSpan(days, hours, minutes, seconds, milliseconds)
    End Sub

    Public Sub New(ByVal hours As Integer, ByVal minutes As Integer, ByVal seconds As Integer, ByVal milliseconds As Integer)
      _timeSpan = New TimeSpan(hours, minutes, seconds, milliseconds)
    End Sub

    Public Sub New(ByVal ticks As Long)
      _timeSpan = New TimeSpan(ticks)
    End Sub

    Public ReadOnly Property Days() As Integer
      Get
        Return _timeSpan.Days
      End Get
    End Property

    Public ReadOnly Property Hours() As Integer
      Get
        Return _timeSpan.Hours
      End Get
    End Property

    Public ReadOnly Property Minutes() As Integer
      Get
        Return _timeSpan.Minutes
      End Get
    End Property

    Public ReadOnly Property Seconds() As Integer
      Get
        Return _timeSpan.Seconds
      End Get
    End Property

    Public ReadOnly Property MilliSeconds() As Integer
      Get
        Return _timeSpan.Milliseconds
      End Get
    End Property

    Public ReadOnly Property Ticks() As Long
      Get
        Return _timeSpan.Ticks
      End Get
    End Property

    Public ReadOnly Property TotalDays() As Double
      Get
        Return _timeSpan.TotalDays
      End Get
    End Property

    Public ReadOnly Property TotalHours() As Double
      Get
        Return _timeSpan.TotalHours
      End Get
    End Property

    Public ReadOnly Property TotalMinutes() As Double
      Get
        Return _timeSpan.TotalMinutes
      End Get
    End Property

    Public ReadOnly Property TotalSeconds() As Double
      Get
        Return _timeSpan.TotalSeconds
      End Get
    End Property

    Public ReadOnly Property TotalMilliSeconds() As Double
      Get
        Return _timeSpan.TotalMilliseconds
      End Get
    End Property

    Public Overrides Function ToString() As String
      Return Me.Days.ToString("0000") & " " & _
              Math.Abs(Me.Hours).ToString("00") & ":" & _
              Math.Abs(Me.Minutes).ToString("00") & ":" & _
              Math.Abs(Me.Seconds).ToString("00") & "." & _
              Math.Abs(Me.MilliSeconds).ToString("00")
    End Function

    Public Function ToTimeSpan() As TimeSpan
      Return _timeSpan
    End Function

    Public Overloads Function Equals(ByVal ticks As Long) As Boolean
      Return Me.Ticks = ticks
    End Function

    Public Overloads Function Equals(ByVal int As Interval) As Boolean
      Return Me.Equals(int.ToTimeSpan)
    End Function

    Public Overloads Function Equals(ByVal ts As TimeSpan) As Boolean
      Return _timeSpan.Equals(ts)
    End Function

#End Region

  End Class


#End Region

#Region "Support Classes"
  Public Class LIB_CUSTOPT

    Public Shared Function Option_Translate_DueOff(ByVal optValue As String) As String
      Return Options.Option_Translate_DueOff(optValue)
    End Function
  End Class
#End Region

#Region "Char validator"

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : Glossary.CharRangeValidator
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Class for validating strings against "Allowed Character" strings as used
  '''   in structure.txt and Analysis Component definitions.
  ''' </summary>
  ''' <remarks>
  '''   Validation strings are made up of individual allowable characters, plus ranges
  '''   in the form x..y where x is the first allowable character and y is the last.<p/>
  '''   For example 'A..Z0..9_-' is the validation string used for most identity fields,
  '''   allowing upper case chars, numbers, underscore and hyphen only.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class CharRangeValidator
    Inherits ReadOnlyCollectionBase

    Private _chars As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Create a new validator based on a validation string
    ''' </summary>
    ''' <param name="allowedChars"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal allowedChars As String)
      MyBase.new()
      If allowedChars Is Nothing Then Exit Sub
      Dim findrange As New Text.RegularExpressions.Regex(".\.\..")
      Dim rangeText As Text.RegularExpressions.Match
      Dim remainingChars As New Text.StringBuilder(allowedChars)
      Dim removedChars As Integer
      For Each rangeText In findrange.Matches(allowedChars)
        MyBase.InnerList.Add(New CharRange(rangeText.Value.Chars(0), rangeText.Value.Chars(rangeText.Length - 1)))
        ' Remove the range from the allowed characters string
        remainingChars.Remove(rangeText.Index - removedChars, rangeText.Length)
        removedChars = removedChars + rangeText.Length
      Next
      _chars = remainingChars.ToString
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   The characters in the validation string that were not part of any ranges
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   In the 'A..Z0..9_-' example, this property would return '_-'
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property RemainingChars() As String
      Get
        ' Char arrays can be converted to strings automatically, even with option strict on.
        Return _chars
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the range at the given index
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal index As Integer) As CharRange
      Get
        Return CType(MyBase.InnerList.Item(index), CharRange)
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed value complies with the validation string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   To comply, each character in the passed string must be inside one of the ranges
    '''   defined in the validation string, or be one of the additional characters in the
    '''   RemainingChars property.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidValue(ByVal value As String) As Boolean
      Dim c As Char, isValid As Boolean
      For Each c In value.ToCharArray
        If Not _chars Is Nothing Then
          ' Check to see if character is in list of allowed chars
          isValid = _chars.IndexOf(c) >= 0
        Else
          ' If no characters and no ranges specified, must be valid
          isValid = MyBase.Count = 0
        End If
        ' If not (yet) valid, check char ranges
        If Not isValid Then
          Dim range As CharRange
          For Each range In Me
            isValid = isValid Or range.IsInRange(c)
          Next
        End If
        If Not isValid Then Exit For
      Next
      Return isValid
    End Function

#Region "Char Range"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.CharRangeValidator.CharRange
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Class to validate a character against a character range
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CharRange

      Private _min As Char, _max As Char

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Create a new range with the specified minimum and maximum values
      ''' </summary>
      ''' <param name="min"></param>
      ''' <param name="max"></param>
      ''' <remarks>
      '''   No validation is performed to ensure that max is >= min, as many instances
      '''   of this are seen in structure.txt.
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Sub New(ByVal min As Char, ByVal max As Char)
        _min = min
        _max = max
      End Sub

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Indicates whether the passed character is inside the range
      ''' </summary>
      ''' <param name="value"></param>
      ''' <returns></returns>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Function IsInRange(ByVal value As Char) As Boolean
        Return value >= _min And value <= _max
      End Function

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Indicates whether all characters in the passed string are in the range
      ''' </summary>
      ''' <param name="value"></param>
      ''' <returns></returns>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Function IsInRange(ByVal value As String) As Boolean
        Dim inRange As Boolean = Not value Is Nothing
        If inRange Then
          Dim c As Char
          For Each c In value.ToCharArray
            inRange = inRange And IsInRange(c)
          Next
        End If
        Return inRange
      End Function

    End Class
#End Region

  End Class

#End Region

#Region "Config Items"
  Public Class ConfigItem
        Implements IDisposable

        Shared Function GetConfigValue(ByVal configID As String) As String
            Dim item As New ConfigItem(configID)
            Try
                Return item.DefaultValue
            Finally
                item.Dispose()
            End Try
        End Function

        Shared Function GetConfigValue(ByVal configID As String, ByVal userID As String) As String
            Dim item As New ConfigItem(configID)
            Try
                Return item.GetValue(userID)
            Finally
                item.Dispose()
            End Try
        End Function

        ''' <summary>
        '''   Attempt to retrieve the config value with respect to current user or group settings
        ''' </summary>
        ''' <param name="configID"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Shared Function GetCurrentConfigValue(ByVal configID As String) As String
            ' Default to current user
            Dim user As User = user.InnerUser
            ' If not found attempt to load from current Windows user
            If user Is Nothing Then
                user = user.GetCurrentUser
            End If
            ' Retrieve config value
            If user Is Nothing Then
                Return GetConfigValue(configID)
            Else
                Return GetConfigValue(configID, user.UserID)
            End If
        End Function


#Region "Instance"
    Private _getItem As OleDb.OleDbDataAdapter
    Private _configData As DataRow

    Public Sub New(ByVal Identity As String)
      _getItem = New OleDb.OleDbDataAdapter("SELECT * FROM Config_Header WHERE Identity = ?", MedConnection.Connection)
      _getItem.SelectCommand.Parameters.Add("Identity", OleDb.OleDbType.VarChar)
      _getItem.UpdateCommand = New OleDb.OleDbCommand("UPDATE Config_Header SET Value = ? WHERE Identity = ?", MedConnection.Connection)
      _getItem.UpdateCommand.Parameters.Add("Value", OleDb.OleDbType.VarChar).SourceColumn = "Value"
      _getItem.UpdateCommand.Parameters.Add("Identity", OleDb.OleDbType.VarChar).SourceColumn = "Identity"
      Me.Refresh(Identity)
    End Sub

    Private Sub Refresh(ByVal identity As String)
      _getItem.SelectCommand.Parameters(0).Value = identity
      Try
        If MedConnection.Open(_getItem.SelectCommand.Connection) Then
          Dim header As New DataTable("Config_Header")
          _getItem.Fill(header)
          MedConnection.Close(_getItem.SelectCommand.Connection)
          If header.Rows.Count > 0 Then
            _configData = header.Rows(0)
          Else
            ' Saves exceptions being thrown all over the place
            _configData = header.NewRow
          End If
        End If
      Catch
      Finally
        MedConnection.Close()
      End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Refresh the data in the Configuration Item object
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Refresh()
      Me.Refresh(Me.Identity)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Writes new value back to Config_Header table
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [05/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Update()
      If _configData.HasVersion(DataRowVersion.Proposed) Then
        Try
          MedConnection.Open(_getItem.SelectCommand.Connection)
          _getItem.Update(New DataRow() {_configData})
          _configData.AcceptChanges()
        Catch
        Finally
          MedConnection.Close(_getItem.SelectCommand.Connection)
        End Try
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the configuration item ID
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Identity() As String
      Get
        Return _configData!Identity.ToString
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the default configuration item value
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property DefaultValue() As String
      Get
        Return _configData!Value.ToString
      End Get
      Set(ByVal Value As String)
        If Me.Modifiable Then
          _configData!Value = Value
        End If
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether the configuration item may be modifed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Modifiable() As Boolean
      Get
        Return _configData!Modifiable.ToString = C_dbTRUE
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Description of configuration item
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Description() As String
      Get
        Return _configData!Description.ToString
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether configuration item may have user specific settings
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property UserModifiable() As Boolean
      Get
        Return _configData!User_Modifiable.ToString = C_dbTRUE
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Data Type configuration item value represents
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DataType() As String
      Get
        Return _configData!Data_Type.ToString
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   The configuration item category
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Category() As String
      Get
        Return _configData!Category.ToString
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the user specific configuration item value
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   If nothing is returned, user does not have a specific value, default should be used.
    '''   Note that this function returns the user setting, regardless of the UserModifiable
    '''   flag.  The GetValue function does take this flag into account.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
        Public Function UserValue(ByVal userID As String) As String
            Dim getValue As New OleDb.OleDbCommand("SELECT Value FROM Config_Item" & _
                                                   " WHERE Identity = ? AND Operator = ?", _
                                                   _getItem.SelectCommand.Connection)
            getValue.Parameters.Add("Config_ID", OleDb.OleDbType.VarChar, 10).Value = Me.Identity
            getValue.Parameters.Add("User_ID", OleDb.OleDbType.VarChar, 10).Value = userID

            Dim strRet As String = ""
            Try
                If MedConnection.Open(getValue.Connection) Then
                    Dim readValue As OleDb.OleDbDataReader = getValue.ExecuteReader
                    If readValue.Read Then
                        MedConnection.Close(getValue.Connection)
                        strRet = readValue.GetString(0)
                    End If
                End If
            Catch
            Finally
                MedConnection.Close()
            End Try
            Return strRet
        End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the configuration item value to use for the specified user
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   If no user specific value is stored, the default is returned
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [15/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetValue(ByVal userID As String) As String
            Dim value As String = ""
      If Me.UserModifiable Then
        value = Me.UserValue(userID)
      End If
      ' value will be nothing if not user modifiable, or no user value set
      If value Is Nothing Then
        value = Me.DefaultValue
      End If
      Return value
    End Function

    Public Sub SetUserValue(ByVal userID As String, ByVal value As String)
      If Not Me.UserModifiable Then
        Throw New InvalidOperationException("Cannot set value for non user-modifiable setting.")
      End If
      Dim setCommand As New OleDb.OleDbCommand()
      setCommand.Connection = _getItem.SelectCommand.Connection
      If Me.UserValue(userID) Is Nothing Then
        ' Insert value
        setCommand.CommandText = "INSERT INTO Config_Item (Value, Identity, Operator) VALUES (?, ?, ?)"
      Else
        ' Update value
        setCommand.CommandText = "UPDATE Config_Item SET Value = ? WHERE Identity = ? AND Operator = ?"
      End If
      setCommand.Parameters.Add(value)
      setCommand.Parameters.Add(Me.Identity)
      setCommand.Parameters.Add(userID)
      Try
        MedConnection.Open(setCommand.Connection)
        setCommand.ExecuteNonQuery()
      Catch
      Finally
        MedConnection.Close(setCommand.Connection)
      End Try
    End Sub

#End Region

        Private disposedValue As Boolean = False    ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    '' TODO: free other state (managed objects).
                End If
                _getItem.Dispose()
                '' TODO: free your own state (unmanaged objects).
                '' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

  End Class
#End Region


End Namespace

Namespace Defaults

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Defaults.CustomerDefault
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A class used to manage defaults for a customer type, not intended to be database based but to use XML files
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CustomerDefault
        Private myIndustryDefaults As New IndustryDefaultCollection()

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get default values 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(CustomerDefault))
            ' To read the file, create a FileStream object.
            Dim strFileName As String = Filename()
            If IO.File.Exists(strFileName) Then
                Try
                    Dim myFileStream As IO.FileStream = _
                    New IO.FileStream(strFileName, IO.FileMode.Open)
                    Dim Myobject As CustomerDefault = CType( _
                    mySerializer.Deserialize(myFileStream), CustomerDefault)

                    With Myobject
                        IndustryDefaults = .IndustryDefaults

                    End With
                    myFileStream.Close()
                    myFileStream = Nothing
                Catch ex As Exception
                End Try
            End If
        End Function

        Public Property IndustryDefaults() As IndustryDefaultCollection
            Get
                Return myIndustryDefaults
            End Get
            Set(ByVal Value As IndustryDefaultCollection)
                myIndustryDefaults = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Filename to be used 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Filename() As String
            Dim strFilename As String = Application.ExecutablePath

            Dim i As Integer = strFilename.Length

            While i > 0 And strFilename.Chars(i - 1) <> "\"
                i -= 1
            End While
            strFilename = Mid(strFilename, 1, i) & "Defaults"


            'IO.Directory.SetCurrentDirectory(Application.ExecutablePath)
            If Not IO.Directory.Exists(strFilename) Then

                IO.Directory.CreateDirectory(strFilename)
            End If
            strFilename += "\CustomerAccDefaults.XML"

            Return strFilename
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Save default values
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Save() As Boolean
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(CustomerDefault))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = Filename()
                Dim myWriter As IO.StreamWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
                myWriter.Flush()
                myWriter.Close()

            Catch ex As Exception
            End Try
        End Function

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Defaults.IndustryDefault
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Defaults for an industry
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class IndustryDefault
        Private strIndustry As String = ""
        Private strTestPanel As String = ""
        Private strBreath As String = ""
        Private strAlcohol As String = ""
        Private strCannabis As String = ""
        Private myOptionDefaults As New OptionDefaultCollection()
        Private strCollectionType As String = ""

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create Industry Default Item 
        ''' </summary>
        ''' <param name="Industry"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Industry As String)
            strIndustry = Industry
        End Sub

        Public Sub New()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default alcohol cutoff 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property AlcoholCutoff() As String
            Get
                Return Me.strAlcohol
            End Get
            Set(ByVal Value As String)
                strAlcohol = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Breath Alcohol cutoff 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property BreathAlcoholCutoff() As String
            Get
                Return Me.strBreath
            End Get
            Set(ByVal Value As String)
                strBreath = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default cannabis cutoff 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CannabisCutoff() As String
            Get
                Return Me.strCannabis

            End Get
            Set(ByVal Value As String)
                strCannabis = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Industry 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Industry() As String
            Get
                Return Me.strIndustry
            End Get
            Set(ByVal Value As String)
                strIndustry = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Typical collection type 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CollectionType() As String
            Get
                Return Me.strCollectionType
            End Get
            Set(ByVal Value As String)
                strCollectionType = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Typical testpanel
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property TestPanel() As String
            Get
                Return Me.strTestPanel
            End Get
            Set(ByVal Value As String)
                strTestPanel = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option defaults
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionDefaults() As OptionDefaultCollection
            Get
                Return Me.myOptionDefaults
            End Get
            Set(ByVal Value As OptionDefaultCollection)
                myOptionDefaults = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Defaults.IndustryDefaultCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection of industry defaults 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class IndustryDefaultCollection
        Inherits CollectionBase

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new collection 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()
        End Sub

        Default Public Overloads Property Item(ByVal index As Integer) As IndustryDefault
            Get
                Return CType(MyBase.List.Item(index), IndustryDefault)
            End Get
            Set(ByVal Value As IndustryDefault)
                MyBase.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get item by industry  code 
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Overloads Property Item(ByVal index As String) As IndustryDefault
            Get
                Dim oIndDef As IndustryDefault = Nothing
                For Each oIndDef In Me.List
                    If oIndDef.Industry = index Then
                        Exit For
                    End If
                    oIndDef = Nothing
                Next
                Return oIndDef
            End Get
            Set(ByVal Value As IndustryDefault)

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an item to collection 
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As IndustryDefault) As Integer
            Return MyBase.List.Add(item)
        End Function
    End Class

    Public Class OptionDefault
        Private strOptionID As String = ""
        Private strOptionValue As String = ""

        Public Sub New()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' default option value 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionId() As String
            Get
                Return Me.strOptionID
            End Get
            Set(ByVal Value As String)
                strOptionID = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default option value 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionValue() As String
            Get
                Return Me.strOptionValue
            End Get
            Set(ByVal Value As String)
                strOptionValue = Value
            End Set
        End Property
    End Class

    Public Class OptionDefaultCollection
        Inherits CollectionBase

        Public Function Add(ByVal item As OptionDefault) As Integer
            Return MyBase.List.Add(item)
        End Function

        Public Sub New()
            MyBase.new()
        End Sub


        Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As OptionDefault
            Get
                Return CType(MyBase.List.Item(index), OptionDefault)
            End Get
        End Property

        Default Public Overloads Property Item(ByVal index As String) As OptionDefault
            Get
                Dim objDef As OptionDefault = Nothing
                For Each objDef In Me.List
                    If objDef.OptionId = index Then
                        Exit For
                    End If
                    objDef = Nothing
                Next
                Return objDef
            End Get

            Set(ByVal Value As OptionDefault)

            End Set
        End Property

    End Class


End Namespace


Public Interface IPhraseEntry

  ReadOnly Property Identity() As String
  ReadOnly Property Text() As String

End Interface