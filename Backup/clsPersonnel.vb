Imports MedscreenLib
Imports MedscreenLib.Medscreen
Namespace personnel

    '''<summary>
    ''' 
    ''' </summary>
#Region "Person"

    '''<summary>
    ''' Describes a sample manager user
    ''' </summary>
    Public Class UserPerson

#Region "Declarations"


        'Private Instance As Personnel.PERSONNELRow
        Private objIdentity As StringField
        Private objGroupID As StringField
        Private objDefaultGroup As StringField
        Private objDescription As StringField
        Private objAuthority As StringField
        Private objLocationID As StringField
        Private objCostCode As StringField
        Private objModifiedON As DateField
        Private objModifiedBY As StringField
        Private objModifiable As StringField
        Private objRemoveFlag As StringField
        Private objEMAIL As StringField
        Private objWindowsIdentity As StringField
        Private objDepartmentID As StringField
        Private objManagerID As StringField
        Private objOptionAuthority As IntegerField = New IntegerField("OPTION_AUTHORITY", 1)

        Private myFields As TableFields
#End Region

#Region "code"

        '''<summary>
        ''' Authority Level associated with the User (Sample Manager)
        ''' </summary>
        Public Property Authority() As String
            Get
                Return Me.objAuthority.Value
            End Get
            Set(ByVal Value As String)
                Me.objAuthority.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Authority Level associated with the user (Medscreen)
        ''' </summary>
        Public Property OptionAuthority() As Integer
            Get
                Return Me.objOptionAuthority.Value
            End Get
            Set(ByVal Value As Integer)
                objOptionAuthority.Value = Value
            End Set
        End Property

        Private Sub SetUpFields()

            myFields = New TableFields("PERSONNEL")


            objIdentity = New StringField("IDENTITY", "", 10)
            myFields.Add(objIdentity)
            objGroupID = New StringField("GROUP_ID", "", 10)
            myFields.Add(objGroupID)
            objDefaultGroup = New StringField("DEFAULT_GROUP", "", 10)
            myFields.Add(objDefaultGroup)
            objDescription = New StringField("DESCRIPTION", "", 234)
            myFields.Add(objDescription)
            objAuthority = New StringField("AUTHORITY", "", 10)
            myFields.Add(objAuthority)
            objLocationID = New StringField("LOCATION_ID", "", 10)
            myFields.Add(objLocationID)
            objCostCode = New StringField("COST_CODE", "", 15)
            myFields.Add(objCostCode)
            objModifiedON = New DateField("MODIFIED_ON", DateField.ZeroDate)
            myFields.Add(objModifiedON)
            objModifiedBY = New StringField("MODIFIED_BY", "", 10)
            myFields.Add(objModifiedBY)
            objModifiable = New StringField("MODIFIABLE", "", 1)
            myFields.Add(objModifiable)
            objRemoveFlag = New StringField("REMOVEFLAG", "", 1)
            myFields.Add(objRemoveFlag)
            objEMAIL = New StringField("EMAIL", "", 40)
            myFields.Add(objEMAIL)
            objWindowsIdentity = New StringField("WINDOWSIDENTITY", "", 15)
            myFields.Add(objWindowsIdentity)
            objDepartmentID = New StringField("DEPARTMENT_ID", "", 15)
            myFields.Add(objDepartmentID)
            objManagerID = New StringField("MANAGER_ID", "", 10)
            myFields.Add(objManagerID)
            myFields.Add(objOptionAuthority)

        End Sub

        '''<summary>
        ''' Indicates that the user has been removed
        ''' </summary>
        ''' <returns>TRUE if user has been removed</returns>
        Public Function Removed() As Boolean
            If Me.objRemoveFlag.Value = "T" Then
                Return True
            Else
                Return False
            End If
        End Function

        '''<summary>
        ''' Create a new user entity
        ''' </summary>
        ''' <param name='UserId'>SM user identity</param>
        Public Sub New(ByVal UserId As String)
            'Instance = InstanceData
            SetUpFields()
            Me.objIdentity.Value = UserId
            Me.objIdentity.OldValue = UserId
        End Sub

        '''<summary>
        ''' menu (sample manager) to use
        ''' </summary>
        Public ReadOnly Property Location() As String
            Get
                Return Me.objLocationID.Value
            End Get
        End Property

        '''<summary>
        ''' User's manager is a pig's ear relationship back to personnel table
        ''' </summary>
        Public Property Manager() As String
            Get
                Return Me.objManagerID.Value
            End Get
            Set(ByVal Value As String)
                Me.objManagerID.Value = Value
            End Set
        End Property

        '''<summary>
        ''' User Employee ID
        ''' </summary>
        Public Property CostCode() As String
            Get
                Return Me.objCostCode.Value
            End Get
            Set(ByVal Value As String)
                Me.objCostCode.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Department the user works in 
        ''' </summary>
        Public Property Department() As String
            Get
                Return Me.objDepartmentID.Value
            End Get
            Set(ByVal Value As String)
                objDepartmentID.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Desciption of the user seems to be the full name
        ''' </summary>
        Public Property Description() As String
            Get
                Return Me.objDescription.Value
            End Get
            Set(ByVal Value As String)
                objDescription.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The user's Email address
        ''' </summary>
        Public Property Email() As String
            Get
                Return Me.objEMAIL.Value
            End Get
            Set(ByVal Value As String)
                Me.objEMAIL.Value = Value
            End Set
        End Property


        '''<summary>
        ''' The users Windows Identity
        ''' </summary>
        Public Property WindowsID() As String
            Get
                Return Me.objWindowsIdentity.Value
            End Get
            Set(ByVal Value As String)
                Me.objWindowsIdentity.Value = Value
            End Set
        End Property

        '''<summary>
        ''' The user's Sample Manager Identity
        ''' </summary>
        Public Property Identity() As String
            Get
                Return Me.objIdentity.Value
            End Get
            Set(ByVal Value As String)
                Me.objIdentity.Value = Value
            End Set
        End Property

        '''<summary>
        ''' Convert details to XML
        ''' </summary>
        Public Function ToXml() As String
            Return myFields.ToXMLSchema()
        End Function


        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property
#End Region
    End Class
#End Region


#Region "PersonelList"
    '''<summary>
    ''' A list of Sample Manager Users
    ''' </summary>
    Public Class PersonnelList

#Region "Declarations"


        Inherits CollectionBase

        Private objIdentity As StringField
        Private objGroupID As StringField
        Private objDefaultGroup As StringField
        Private objDescription As StringField
        Private objAuthority As StringField
        Private objLocationID As StringField
        Private objCostCode As StringField
        Private objModifiedON As DateField
        Private objModifiedBY As StringField
        Private objModifiable As StringField
        Private objRemoveFlag As StringField
        Private objEMAIL As StringField
        Private objWindowsIdentity As StringField
        Private objDepartmentID As StringField
        Private objManagerID As StringField
        Private objOptionAuthority As IntegerField = New IntegerField("OPTION_AUTHORITY", 1)

        Private myFields As TableFields
#End Region

#Region "Enumerations"

        '''<summary>
        ''' How to index the Item 
        ''' </summary>
        Public Enum FindIdBy
            '''<summary>Parameter represents a SM ID</summary>
            identity
            '''<summary>Parameter represents a Windows user name</summary>
            WindowsIdentity
            '''<summary>Parameter represents a Employee ID</summary>
            CostCode
            '''<summary>Parameter represents a department</summary>
            Department
        End Enum
#End Region

#Region "Code"


        Private Sub SetUpFields()

            myFields = New TableFields("PERSONNEL")


            objIdentity = New StringField("IDENTITY", "", 10)
            myFields.Add(objIdentity)
            objGroupID = New StringField("GROUP_ID", "", 10)
            myFields.Add(objGroupID)
            objDefaultGroup = New StringField("DEFAULT_GROUP", "", 10)
            myFields.Add(objDefaultGroup)
            objDescription = New StringField("DESCRIPTION", "", 234)
            myFields.Add(objDescription)
            objAuthority = New StringField("AUTHORITY", "", 10)
            myFields.Add(objAuthority)
            objLocationID = New StringField("LOCATION_ID", "", 10)
            myFields.Add(objLocationID)
            objCostCode = New StringField("COST_CODE", "", 15)
            myFields.Add(objCostCode)
            objModifiedON = New DateField("MODIFIED_ON", DateField.ZeroDate)
            myFields.Add(objModifiedON)
            objModifiedBY = New StringField("MODIFIED_BY", "", 10)
            myFields.Add(objModifiedBY)
            objModifiable = New StringField("MODIFIABLE", "", 1)
            myFields.Add(objModifiable)
            objRemoveFlag = New StringField("REMOVEFLAG", "", 1)
            myFields.Add(objRemoveFlag)
            objEMAIL = New StringField("EMAIL", "", 40)
            myFields.Add(objEMAIL)
            objWindowsIdentity = New StringField("WINDOWSIDENTITY", "", 15)
            myFields.Add(objWindowsIdentity)
            objDepartmentID = New StringField("DEPARTMENT_ID", "", 15)
            myFields.Add(objDepartmentID)
            objManagerID = New StringField("MANAGER_ID", "", 10)
            myFields.Add(objManagerID)
            myFields.Add(objOptionAuthority)

        End Sub

        '''<summary>
        ''' Create a new collection 
        ''' </summary>
        Public Sub New()
            MyBase.New()
            SetUpFields()
        End Sub

        '''<summary>
        ''' Add a UserPerson to the collection
        ''' </summary>
        ''' <param name='item'>UserPerson to add</param>
        ''' <returns>Void</returns>
        Public Function Add(ByVal item As UserPerson)
            Return Me.List.Add(item)
        End Function

        '''<summary>
        ''' Get a UserPerson
        ''' </summary>
        ''' <param name='Index'>Index into table by position</param>
        ''' <returns>A UserPerson</returns>
        Public Overloads Function Item(ByVal Index As Integer) As UserPerson
            Return CType(Me.List.Item(Index), UserPerson)
        End Function

        '''<summary>
        ''' Return a list of the people in a department
        ''' </summary>
        ''' <param name='Department'>Department to list</param>
        ''' <returns>A PersonnelList for the department</returns>
        Public Function DepartmentList(ByVal Department As String) As PersonnelList
            Dim newList As PersonnelList = New PersonnelList()

            Dim aPerson As UserPerson
            For Each aPerson In MyBase.List
                If aPerson.Department.ToUpper.Trim = Department.ToUpper AndAlso aPerson.Removed = False Then
                    newList.Add(aPerson)
                End If
            Next

            Return newList

        End Function

        '''<summary>
        ''' Get a userPerson object 
        ''' </summary>
        ''' <param name='SamId'></param>
        ''' <param name='Opt'>Which element of the entity is to used
        ''' <see cref="PersonnelList.FindIdBy"/>
        ''' </param>
        ''' <returns>UserPerson entity</returns>
        Public Overloads Function Item(ByVal SamId As String, Optional ByVal Opt As FindIdBy = FindIdBy.identity) As UserPerson
            Dim i As Integer
            Dim uP As UserPerson

            If Opt = FindIdBy.identity Then 'SM identity
                For i = 0 To Me.Count - 1
                    uP = Me.Item(i)
                    If uP.Identity.ToUpper = SamId.ToUpper Then
                        Exit For
                    End If
                    uP = Nothing
                Next
            ElseIf Opt = FindIdBy.WindowsIdentity Then 'Windows Login name 
                For i = 0 To Me.Count - 1
                    uP = Me.Item(i)
                    If uP.WindowsID.ToUpper = SamId.ToUpper Then
                        Exit For
                    End If
                    uP = Nothing
                Next
            ElseIf Opt = FindIdBy.CostCode Then 'Cost Code ID 
                For i = 0 To Me.Count - 1
                    uP = Me.Item(i)
                    If uP.CostCode.Trim.ToUpper = SamId.Trim.ToUpper Then
                        Exit For
                    End If
                    uP = Nothing
                Next
            End If
            Return uP
        End Function

        '''<summary>
        ''' Load the contents of the table
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        Public Function Load() As Boolean

            Dim ocmd As New OleDb.OleDbCommand(Me.myFields.SelectString)

            ocmd.Connection = MedscreenLib.CConnection.DbConnection
            Dim oread As OleDb.OleDbDataReader
            Dim objItem As UserPerson

            Try
                MedscreenLib.CConnection.SetConnOpen()
                oread = ocmd.ExecuteReader
                While oread.Read
                    objItem = New UserPerson(oread.GetString(0))
                    objItem.Fields.ReadFields(oread)
                    Me.Add(objItem)
                End While

            Catch ex As Exception
                LogError(ex)
            Finally
                If Not oread Is Nothing Then
                    If Not oread.IsClosed Then oread.Close()
                End If

                MedscreenLib.CConnection.SetConnClosed()

            End Try
        End Function

        '''<summary>
        ''' Convert details to a dataset
        ''' </summary>
        ''' <returns>A dataset</returns>
        Public Function ToDataSet() As DataSet
            Dim tempStr As String = ToXml()
            Dim st As New IO.MemoryStream(StringToByteArray(tempStr))

            Dim Ds As New DataSet()

            Ds.ReadXml(st, XmlReadMode.InferSchema)

            st.Close()
            st = Nothing
            Return Ds
        End Function

        '''<summary>
        ''' Convert collection to XML, calls UserPerson.ToXML 
        ''' </summary>
        ''' <returns>XML representation of collection</returns>
        Public Function ToXml() As String
            Dim i As Integer
            Dim objItem As UserPerson

            Dim strRet As String = "<?xml version='1.0'?><Personnel>"
            For i = 0 To Me.Count - 1
                objItem = Me.Item(i)
                strRet += objItem.ToXml
            Next
            strRet += "</Personnel>"
            Return strRet
        End Function

        '''<summary>
        ''' Write  the contents to a file
        ''' </summary>
        ''' <param name='Filename'>File the details are to be written to</param>
        ''' <returns>Void</returns>
        Public Function WriteXmlToFile(ByVal Filename As String)
            Try
                Dim w = New IO.StreamWriter(Filename, False)
                w.Write(Me.ToXml)
                w.Flush()
                w.Close()

            Catch ex As Exception
            End Try

            'Me.myFields.XmlToFile(Filename, True)
        End Function

        '''<summary>
        ''' Provide the users name
        ''' </summary>
        ''' <param name='SamId'>Sample Manager ID</param>
        ''' <returns>User's name</returns>
        Public Function GetPersonnelName(ByVal SamId As String) As String

            Dim UP As UserPerson = Me.Item(SamId)

            If UP Is Nothing Then
                Return "Unknown"
            Else
                Return UP.Description
            End If

        End Function

#End Region
    End Class

#End Region


End Namespace