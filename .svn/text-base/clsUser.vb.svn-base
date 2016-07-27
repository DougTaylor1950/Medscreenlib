Option Strict On

'''<summary>
''' Class dealing with user Logins
''' </summary>
<CLSCompliant(True)> _
Public Class User

#Region "Shared"

    '''<summary>Enumeration to aid selecting users</summary>
    Public Enum SearchType
        '''<summary>Select by Sample Manager ID</summary>
        Identity
        '''<summary>Select by Windows ID</summary>
        NetworkID
        NumericID
    End Enum


    Private Shared _userService As MedscreenLib.ts01.Service1
    Private Shared _frmLogin As frmLogin

    Private Shared _users As Hashtable
    Private Shared _user As User
    Private Shared _userTable As DataTable

    Shared Sub New()
        _users = New Hashtable()
        ' Initialise data commands
    End Sub

    '''<summary>
    ''' Get the user who is currently logged in to Windows
    ''' </summary>
    ''' <returns>A User Object</returns>
    Public Shared Property CurrentUser() As MedscreenLib.User
        Get
            If MedscreenLib.User.InnerUser Is Nothing AndAlso _
               MedscreenLib.User.LoginForm.ShowDialog = Windows.Forms.DialogResult.OK AndAlso _
               MedscreenLib.User.LoginForm.SelectedUser.IsValidated Then
                MedscreenLib.User.CurrentUser = MedscreenLib.User.LoginForm.SelectedUser
            End If
            Return MedscreenLib.User.InnerUser
        End Get
        Set(ByVal Value As MedscreenLib.User)
            ' Log out existing user
            If Not Value Is _user Then
                If Not InnerUser Is Nothing Then
                    If InnerUser.IsLoggedIn AndAlso Not InnerUser.LogOut() Then
                        Exit Property
                    End If
                End If
                ' Update
                InnerUser = Value
                ' Log in new user
                If Not InnerUser Is Nothing Then
                    ' Handlers added explicitly (rather than using WithEvents) because Login event was not being detected reliably
                    InnerUser.LogIn()
                End If
            End If
        End Set
    End Property



    Public Shared Function AuthenticateUser(ByVal domainAndUsername As String, ByVal pwd As String) As Boolean

        'Dim domainAndUsername As String = domain + "\" + username
        Dim path As String = "" ' Will contain path to AD user pofile, but we don't need it.
        Dim entry As New System.DirectoryServices.DirectoryEntry(path, domainAndUsername, pwd)

        Try

            ' Bind to the native AdsObject to force authentication.  Exception will be thrown here if validation fails.
            Return Not (entry.NativeObject Is Nothing)

            'Dim search As New System.DirectoryServices.DirectorySearcher(entry)

            'search.Filter = "(SAMAccountName=" + domainAndUsername.Substring(domainAndUsername.IndexOf("\") + 1) + ")"
            'search.PropertiesToLoad.Add("cn")
            'Dim result As System.DirectoryServices.SearchResult = search.FindOne()

            'If result Is Nothing Then
            '  Return False
            'End If

            '' Update the new path to the user in the directory.
            '_path = result.Path
            '_filterAttribute = CType(result.Properties("cn").Item(0), String)

        Catch ex As Exception
            Throw New Exception(String.Format("Error authenticating user.{0}{0}{1}", ControlChars.NewLine, ex.Message))
            Return False
        End Try

    End Function

    '''<summary>Exposes the login form</summary>
    Public Shared ReadOnly Property LoginForm() As frmLogin
        Get
            If _frmLogin Is Nothing Then _frmLogin = New frmLogin()
            Return _frmLogin
        End Get
    End Property

    Private Shared ReadOnly Property UserTable() As DataTable
        Get
            If _userTable Is Nothing Then
                _userTable = MedStructure.GetTable("Personnel").CreateDataTable
            End If
            Return _userTable
        End Get
    End Property

    '''<summary>Exposes User Service on TS01 (SOAP Service)</summary>
    Private Shared Function ValidationService() As MedscreenLib.ts01.Service1
        If _userService Is Nothing Then
            _userService = New MedscreenLib.ts01.Service1()
            _userService.Timeout = 5000
        End If
        Return _userService
    End Function

    '''<summary></summary>
    Private Shared Function WhereClause(ByVal CmdType As SearchType) As String
        Select Case CmdType
            Case SearchType.Identity
                Return "Identity = ?"
            Case SearchType.NetworkID
                Return "WindowsIdentity = ?"
            Case Else
                Return ""
        End Select
    End Function

    Public Shared Function GetUser(ByVal identity As String) As User
        Return GetUser(identity, SearchType.Identity)
    End Function

    '''<summary>
    ''' Get a user by a value and a SearchType <see cref="User.SearchType" />
    ''' </summary>
    ''' <param name='Identity'>Identity to search on</param>
    ''' <param name='Type'>Method to use when searching <see cref="User.SearchType" /></param>
    ''' <returns>A User Object</returns>
    Public Shared Function GetUser(ByVal identity As String, ByVal type As SearchType) As User
        If Type = SearchType.NetworkID Then
            identity = identity.ToLower
        Else
            identity = identity.ToUpper
        End If
        Dim myUser As User = Nothing
        Dim filter As String = MedscreenLib.User.WhereClause(Type)
        Dim userData() As DataRow = UserTable.Select(filter.Replace("?", "'" & identity & "'"))
        If userData.Length > 0 Then
            myUser = DirectCast(_users.Item(userData(0)!Identity), User)
        End If
        If myUser Is Nothing Then
            Dim daGetUser As New OleDb.OleDbDataAdapter("SELECT * FROM Personnel WHERE " & filter, MedConnection.Connection)
            daGetUser.SelectCommand.Parameters.Add("ID", OleDb.OleDbType.VarChar).Value = identity
            Try
                daGetUser.Fill(UserTable)
                userData = UserTable.Select(filter.Replace("?", "'" & identity & "'"))
                If userData.Length > 0 Then
                    myUser = New User(userData(0))
                    _users.Add(myUser.UserID, myUser)
                End If
            Catch ex As System.Exception
                Console.WriteLine(ex.Message)
            End Try
        End If
        Return myUser
    End Function

    '''<summary>
    ''' Get the user who is currently logged in to Windows
    ''' </summary>
    ''' <returns>A User Object</returns>
    Public Shared Function GetCurrentUser() As User
        '<Added code modified 16-Apr-2007 10:36 by LOGOS\taylor> 
        Dim strUserName As String = System.Security.Principal.WindowsIdentity.GetCurrent().Name

        'see if we are testuser if so login as taylor
        'If InStr(strUserName.ToUpper, "TEST") <> 0 Then strUserName = "logos\taylor"

        Dim tmpUser As MedscreenLib.User = GetUser(strUserName, SearchType.NetworkID)
        If tmpUser Is Nothing Then
            MsgBox(strUserName)
        End If
        Return tmpUser

        '</Added code modified 16-Apr-2007 10:36 by LOGOS\taylor> 
    End Function

    Public Overloads Shared Sub SetUser(ByVal UserID As String)
        Dim oCmd As New OleDb.OleDbCommand("Lib_User.SetUser", CConnection.DbConnection)
        Try
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.Parameters.Add(CConnection.StringParameter("UserId", UserID, 10))
            If CConnection.ConnOpen Then
                oCmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, True)
        End Try
    End Sub

    Public Overloads Sub SetUser()
        Dim oCmd As New OleDb.OleDbCommand("Lib_User.SetUser", CConnection.DbConnection)
        Try
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.Parameters.Add(CConnection.StringParameter("UserId", Me.UserID, 10))
            If CConnection.ConnOpen Then
                oCmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, True)
        End Try
    End Sub

    ''' <summary>
    '''   Get a list of all users in a given role
    ''' </summary>
    ''' <param name="roleIDs" type="String">
    '''   A single role, or comma delimited list of roles
    ''' </param>
    ''' <returns>
    '''   An array of MedscreenLib.User, one entry per user in role
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Function GetUsersInRoles(ByVal roleIDs As String) As User()
        Return GetUsersInRoles(roleIDs.Split(","c))
    End Function

    ''' <summary>
    '''   Get a list of all users in a given role
    ''' </summary>
    ''' <param name="roleIDs" type="String">
    ''' </param>
    ''' <returns>
    '''   An array of MedscreenLib.User, one entry per user in role
    ''' </returns>
    ''' <remarks>
    '''   Only one entry is returned per user, even if they are in more than one role
    ''' </remarks>
    Public Shared Function GetUsersInRoles(ByVal roleIDs() As String) As User()
        Dim roleID As String
        Dim users As New ArrayList()
        For Each roleID In roleIDs
            Dim roleHolder As Glossary.RoleAssignment
            For Each roleHolder In Glossary.RoleAssignments.Current.RoleHolders(roleID)
                If roleHolder.User > "" Then
                    Dim smUser As User = MedscreenLib.User.GetUser(roleHolder.User)
                    If Not smUser Is Nothing AndAlso _
                       users.IndexOf(smUser) < 0 Then
                        users.Add(smUser)
                    End If
                End If
            Next
        Next
        Return DirectCast(users.ToArray(GetType(User)), User())
    End Function

    Public Shared Function GetEmail(ByVal userid As String) As String
        Dim _user As User = GetUser(userid, SearchType.Identity)
        If _user IsNot Nothing Then
            Return _user.Email
        Else
            Return ""
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or Sets the current user, displaying Login Dialog if required
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	14/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property User() As MedscreenLib.User
        Get
            If MedscreenLib.User.InnerUser Is Nothing AndAlso _
         MedscreenLib.User.LoginForm.ShowDialog = Windows.Forms.DialogResult.OK AndAlso _
         MedscreenLib.User.LoginForm.SelectedUser.IsValidated Then
                MedscreenLib.User.User = MedscreenLib.User.LoginForm.SelectedUser
            End If
            Return MedscreenLib.User.InnerUser
        End Get
        Set(ByVal Value As MedscreenLib.User)
            ' Log out existing user
            If Not value Is _user Then
                If Not _user Is Nothing Then
                    If _user.IsLoggedIn AndAlso Not _user.LogOut() Then
                        Exit Property
                    End If
                End If
                ' Update
                _user = value
                ' Log in new user
                If Not _user Is Nothing Then
                    ' Handlers added explicitly (rather than using WithEvents) because Login event was not being detected reliably
                    _user.LogIn()
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the current user ID, displaying login form if necessary
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	23/10/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property CurrentUserID() As String
        Get
            If MedscreenLib.User.User Is Nothing Then
                Return ""
            Else
                Return MedscreenLib.User.User.UserID
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the current user, without displaying a Login dialog
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	14/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property InnerUser() As MedscreenLib.User
        Get
            Return _user
        End Get
        Private Set(ByVal value As MedscreenLib.User)
            _user = value
        End Set
    End Property

#End Region

#Region "Instance"
    Private m_LastMessage As String, m_UserData As DataRow, m_LoginInfo As UserLoginInfo
    Private m_Validated As Date = Constants.C_InitDate

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An event that fires before login
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">An <see cref="System.ComponentModel.CancelEventArgs" /> that contains the event data</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Event BeforeLogin(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)

    '''<summary>
    ''' An event that fires before logout
    ''' </summary>
    ''' <param name='sender'>The source of the event</param>
    ''' <param name='e'>An <see cref="System.ComponentModel.CancelEventArgs" /> that contains the event data</param>
    Public Shared Event BeforeLogout(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
    Public Shared Event AfterLogin As EventHandler
    Public Shared Event AfterLogout As EventHandler

    '''<summary>
    ''' User's Sample Manager ID
    ''' </summary>
    Private Sub New(ByVal userData As DataRow)
        m_UserData = userData
    End Sub


    '''<summary>
    ''' User's Sample Manager ID
    ''' </summary>
    Public ReadOnly Property UserID() As String
        Get
            Return m_UserData("Identity").ToString.Trim
        End Get
    End Property

    '''<summary>
    ''' User's Full Name
    ''' </summary>
    Public ReadOnly Property FullName() As String
        Get
            Return m_UserData("Description").ToString
        End Get
    End Property

    '''<summary>
    ''' User's Sample Manager Authority level
    ''' </summary>
    Public ReadOnly Property Authority() As Integer
        Get
            ' Authority is actually a text field, so trap conversion errors
            Try
                Return CInt(m_UserData("Authority"))
            Catch ex As System.Exception
                Return 0
            End Try
        End Get
    End Property

    '''<summary>
    ''' User's Department
    ''' </summary>
    Public ReadOnly Property Department() As String
        Get
            Return m_UserData("Department_ID").ToString
        End Get
    End Property

    '''<summary>
    ''' User's Email Address
    ''' </summary>
    Public ReadOnly Property Email() As String
        Get
            Return m_UserData("Email").ToString
        End Get
    End Property

    '''<summary>
    ''' User's Manager
    ''' </summary>
    Public ReadOnly Property Manager() As User
        Get
            Return GetUser(m_UserData("Manager_ID").ToString, SearchType.Identity)
        End Get
    End Property

    '''<summary>
    ''' User's Windows Identity
    ''' </summary>
    Public ReadOnly Property WindowsIdentity() As String
        Get
            Return m_UserData("WindowsIdentity").ToString
        End Get
    End Property

    '''<summary>
    ''' Get the user's login info
    ''' </summary>
    Public ReadOnly Property LoginInfo() As UserLoginInfo
        Get
            Return m_LoginInfo
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the current user ensuring they have been validated against Active Directory
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   CurrentUser can return a user even if they have not been validated,
    '''   this method will only return a user object if the user has been validated
    '''   against Active Directory (i.e. entered correct password), otherwise it will
    '''   return Nothing (though CurrentUser will still hold the current user object).
    '''   <para>
    '''   Users can be forced to re-validate themselves periodically by setting the
    '''   MedscreenLib.User.ValidationTimeout propery, or by calling the Validate(True)
    '''   method of the user object.
    '''   </para>
    '''   After sucessful validation the user object returned from this function becomes
    '''    the current user as returned by the CurrentUser property.
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	14/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property GetValidatedUser() As MedscreenLib.User
        Get
            Dim myUser As MedscreenLib.User = InnerUser
            If myUser Is Nothing Then
                myUser = CurrentUser ' Will display login dialog if required
                If myUser IsNot Nothing AndAlso Not myUser.IsValidated Then
                    ' User has had a chance to enter password and failed, return no user
                    myUser = Nothing
                End If
            ElseIf Not myUser.Validate Then ' Will display login dialog if required for validation
                ' Validation failed, return no user
                myUser = Nothing
            End If
            Return myUser
        End Get
    End Property


    '''<summary>
    ''' Log the user in to the current program
    ''' </summary>
    Public Function LogIn() As Boolean
        Dim e As New ComponentModel.CancelEventArgs()
        RaiseEvent BeforeLogin(Me, e)
        If Not e.Cancel Then
            If m_LoginInfo Is Nothing Then
                m_LoginInfo = New UserLoginInfo(MedscreenVariables.ApplicationName, Me.UserID)
            End If
            If m_LoginInfo.LogIn OrElse Me.IsLoggedIn Then
                RaiseEvent AfterLogin(Me, Nothing)
            End If
        End If
        Return Me.IsLoggedIn
    End Function

    '''<summary>
    ''' Log the user out of the current program
    ''' </summary>
    Public Function LogOut() As Boolean
        Dim loggedOut As Boolean = False
        If Not m_LoginInfo Is Nothing Then
            Dim e As New ComponentModel.CancelEventArgs()
            RaiseEvent BeforeLogout(Me, e)
            If Not e.Cancel Then
                loggedOut = m_LoginInfo.LogOut
                If loggedOut Then
                    RaiseEvent AfterLogout(Me, Nothing)
                End If
            End If
        End If
        Return loggedOut
    End Function

    '''<summary>
    ''' Is the User logged in
    ''' </summary>
    Public ReadOnly Property IsLoggedIn() As Boolean
        Get
            If m_LoginInfo Is Nothing Then
                m_LoginInfo = New UserLoginInfo(MedscreenVariables.ApplicationName, Me.UserID)
            End If
            Return m_LoginInfo.IsLoggedIn
        End Get
    End Property

    '''<summary>
    ''' See if the user has been validated
    ''' </summary>
    Public ReadOnly Property IsValidated() As Boolean
        Get
            Return m_Validated > Constants.C_InitDate
        End Get
    End Property

    '''<summary>
    '''     Time Last validated
    ''' </summary>
    Public ReadOnly Property LastValidated() As Date
        Get
            Return m_Validated
        End Get
    End Property

    ''' <summary>
    '''   Validates this user against Active Directory
    ''' </summary>
    ''' <returns>True if the user was validated against active directory</returns>
    ''' <remarks>
    '''  If user has already been validated, no re-validation takes place.
    '''</remarks>
    Public Function Validate() As Boolean
        Return Me.Validate(False)
    End Function


    ''' <summary>
    '''   Displays the login dialog set to this user
    ''' </summary>
    ''' <returns>True if the user was validated against active directory</returns>
    ''' <remarks>
    '''  If forcePwdEntry is True, the user must enter their password,
    '''  regardless of whether theyhave already been validated.
    ''' </remarks>
    Public Function Validate(ByVal forcePwdEntry As Boolean) As Boolean
        If forcePwdEntry OrElse Not Me.IsValidated Then
            LoginForm.SelectedUser = Me
            LoginForm.ShowDialog()
        End If
        Return Me.IsValidated
    End Function



    '''<summary>
    ''' Check the supplied password matches the users password
    ''' </summary>
    ''' <param name='Pwd'>Password to check</param>
    ''' <returns>TRUE if succesful</returns>
    Public Function Validate(ByVal pwd As String) As Boolean
        Try
            Dim valid As Boolean = AuthenticateUser(Me.WindowsIdentity, pwd)
            If valid Then
                m_LastMessage = ""
                m_Validated = DateTime.Now
            End If
            Return valid
        Catch ex As Exception
            m_LastMessage = ex.Message
            Return False
        End Try
    End Function

    '''<summary>
    ''' Returns last message received
    ''' </summary>
    Public ReadOnly Property LastMessage() As String
        Get
            Return m_LastMessage
        End Get
    End Property

    Public ReadOnly Property HasRole(ByVal roleID As String) As Boolean
        Get
            Return Glossary.RoleAssignments.Current.UserHasRole(Me.UserID, roleID)
        End Get
    End Property

#End Region

#Region "UserLoginInfo"
    '''<summary>
    ''' Class giving information about users using a program
    ''' </summary>
    <CLSCompliant(True)> _
  Public Class UserLoginInfo

#Region "Shared"

        Private m_daProgramUsage As OleDb.OleDbDataAdapter

        '''<summary>
        ''' Get program usage about a users
        ''' </summary>
        ''' <param name='Parameter'>Parameter to look for in Extra Info</param>
        ''' <param name='Value'>Looking for users with a parameter value like this</param>
        ''' <returns>Array of users</returns>
        Public Shared Function GetUserWithExtraInfo(ByVal Parameter As String, ByVal Value As String) As User()
            Dim myUsers() As User = Nothing
            Dim cmdFindUsers As New OleDb.OleDbCommand("SELECT User_ID FROM ProgramUsage " & _
                                                       "WHERE ExtraInfo LIKE '%" & Parameter & "=" & Value & "%' " & _
                                                       "AND LoggedIn = 'T'", MedConnection.Connection)
            Try
                CConnection.SetConnOpen()
                Dim rdrUsers As OleDb.OleDbDataReader = cmdFindUsers.ExecuteReader
                Dim userIDs As New ArrayList()
                While rdrUsers.Read
                    userIDs.Add(rdrUsers.GetString(0))
                End While
                rdrUsers.Close()
                If userIDs.Count > 0 Then
                    ReDim Preserve myUsers(userIDs.Count - 1)
                    Dim userCount As Integer, userID As String
                    For Each userID In userIDs
                        myUsers(userCount) = GetUser(userID, SearchType.Identity)
                    Next
                End If
            Catch ex As System.Exception
                Console.WriteLine(ex.Message)
            Finally
                CConnection.SetConnClosed()
            End Try
            Return myUsers
        End Function

        '''<summary>
        ''' Get program usage about a a program
        ''' </summary>
        ''' <param name='ProgramName'>Program to find user Information on</param>
        ''' <returns>Array of users</returns>
        Public Shared Function GetUsersInProgram(ByVal ProgramName As String) As User()
            Dim myUsers() As User = Nothing
            Dim cmdFindUsers As New OleDb.OleDbCommand("SELECT UserID FROM ProgramUsage " & _
                                                       "WHERE ProgramName = ?" & _
                                                       "AND LoggedIn = 'T'", MedConnection.Connection)
            cmdFindUsers.Parameters.Add("ProgramName", Data.OleDb.OleDbType.VarChar, 30).Value = ProgramName
            Try
                CConnection.SetConnOpen()
                Dim rdrUsers As OleDb.OleDbDataReader = cmdFindUsers.ExecuteReader
                Dim userIDs As New ArrayList()
                While rdrUsers.Read
                    userIDs.Add(rdrUsers.GetString(0))
                End While
                rdrUsers.Close()
                If userIDs.Count > 0 Then
                    ReDim Preserve myUsers(userIDs.Count - 1)
                    Dim userCount As Integer, userID As String
                    For Each userID In userIDs
                        myUsers(userCount) = GetUser(userID, SearchType.Identity)
                    Next
                End If
            Catch ex As System.Exception
                Console.WriteLine(ex.Message)
            Finally
                CConnection.SetConnClosed()
            End Try
            Return myUsers
        End Function

        '''<summary></summary>
        Private Function DataAdapter() As OleDb.OleDbDataAdapter
            If m_daProgramUsage Is Nothing Then
                m_daProgramUsage = New OleDb.OleDbDataAdapter()
                m_daProgramUsage.SelectCommand = New OleDb.OleDbCommand()
                With m_daProgramUsage.SelectCommand
                    .Connection = MedConnection.Connection
                    .CommandText = "SELECT * FROM ProgramUsage " & _
                                   "WHERE ProgramName = ? " & _
                                   "AND User_ID = ?"
                    .Parameters.Add("ProgramName", Data.OleDb.OleDbType.VarChar, 30)
                    .Parameters.Add("User_ID", Data.OleDb.OleDbType.VarChar, 10)
                    Medscreen.SetParameterSourceColumns(.Parameters)
                End With
                m_daProgramUsage.UpdateCommand = New OleDb.OleDbCommand()
                With m_daProgramUsage.UpdateCommand
                    .Connection = MedConnection.Connection
                    .CommandText = "UPDATE ProgramUsage " & _
                                   "SET ComputerName = ?, LoggedIn = ?, LastLogin = ?, ExtraInfo = ? " & _
                                   "WHERE ProgramName = ? AND User_ID = ?"
                    .Parameters.Add("ComputerName", Data.OleDb.OleDbType.VarChar, 10)
                    .Parameters.Add("LoggedIn", Data.OleDb.OleDbType.VarChar, 1)
                    .Parameters.Add("LastLogin", Data.OleDb.OleDbType.Date)
                    .Parameters.Add("ExtraInfo", Data.OleDb.OleDbType.VarChar, 40)
                    .Parameters.Add("ProgramName", Data.OleDb.OleDbType.VarChar, 30)
                    .Parameters.Add("User_ID", Data.OleDb.OleDbType.VarChar, 10)
                    Medscreen.SetParameterSourceColumns(.Parameters)
                End With
                m_daProgramUsage.InsertCommand = New OleDb.OleDbCommand()
                With m_daProgramUsage.InsertCommand
                    .Connection = MedConnection.Connection
                    .CommandText = "INSERT INTO ProgramUsage " & _
                                   "(ProgramName, User_ID, LoginGroup, ComputerName, " & _
                                   " LoggedIn, LastLogin, ExtraInfo) " & _
                                   "VALUES (?, ?, ?, ?, ?, ?, ?)"
                    .Parameters.Add("ProgramName", Data.OleDb.OleDbType.VarChar, 30)
                    .Parameters.Add("User_ID", Data.OleDb.OleDbType.VarChar, 10)
                    .Parameters.Add("LoginGroup", Data.OleDb.OleDbType.VarChar, 100)
                    .Parameters.Add("ComputerName", Data.OleDb.OleDbType.VarChar, 10)
                    .Parameters.Add("LoggedIn", Data.OleDb.OleDbType.VarChar, 1)
                    .Parameters.Add("LastLogin", Data.OleDb.OleDbType.Date)
                    .Parameters.Add("ExtraInfo", Data.OleDb.OleDbType.VarChar, 40)
                    Medscreen.SetParameterSourceColumns(.Parameters)
                End With
            End If
            Return m_daProgramUsage
        End Function

#End Region

#Region "Instance"
        Private m_ExtraInfo As New Collections.Specialized.NameValueCollection()
        Private m_LoginData As DataRow, m_ProgramUsage As DataTable

        '''<summary>Create a new instance
        ''' </summary>
        ''' <param name='ProgramName'>Program to find information for</param>
        ''' <param name='UserID'>User to obtain information on</param>
        Friend Sub New(ByVal ProgramName As String, ByVal UserID As String)
            Try
                If m_ProgramUsage Is Nothing Then m_ProgramUsage = New DataTable("ProgramUsage")
                DataAdapter.SelectCommand.Parameters("Programname").Value = ProgramName
                DataAdapter.SelectCommand.Parameters("User_ID").Value = UserID
                If CConnection.ConnOpen Then
                    DataAdapter.Fill(m_ProgramUsage)
                    If m_ProgramUsage.Rows.Count > 0 Then
                        m_LoginData = m_ProgramUsage.Rows(0)
                    Else
                        m_LoginData = m_ProgramUsage.NewRow
                        m_LoginData("User_ID") = UserID
                        m_LoginData("ProgramName") = ProgramName
                    End If
                End If
            Catch ex As System.Exception
                Console.WriteLine(ex.Message)
            Finally
                CConnection.SetConnClosed()
            End Try
        End Sub

        '''<summary>Program that we are trying to obtain information on
        ''' </summary>
        Public ReadOnly Property ProgramName() As String
            Get
                Return m_LoginData("ProgramName").ToString
            End Get
        End Property

        '''<summary>
        ''' Date and time of last login
        ''' </summary>
        Public ReadOnly Property LastLogin() As Date
            Get
                Return DateTime.Parse(m_LoginData("LastLogin").ToString)
            End Get
        End Property

        '''<summary>Computer this instance is running on</summary>
        Public ReadOnly Property ComputerName() As String
            Get
                Return m_LoginData("ComputerName").ToString
            End Get
        End Property

        '''<summary>Is the user this data set belongs to logged in</summary>
        Public ReadOnly Property IsLoggedIn() As Boolean
            Get
                Return m_LoginData("LoggedIn").ToString = Constants.C_dbTRUE
            End Get
        End Property

        '''<summary>Log User into program</summary>
        Public Function LogIn() As Boolean
            Dim loginResult As Boolean = False
            If Not Me.IsLoggedIn Then
                m_LoginData.BeginEdit()
                m_LoginData("ComputerName") = System.Windows.Forms.SystemInformation.ComputerName
                m_LoginData("LoggedIn") = Constants.C_dbTRUE
                m_LoginData("LastLogin") = DateTime.Now
                m_LoginData.EndEdit()
                loginResult = Me.UpdateTable
            End If
            Return loginResult
        End Function

        '''<summary>Log user out of program</summary>
        Public Function LogOut() As Boolean
            Dim logoutResult As Boolean = False
            If Me.IsLoggedIn And _
               Me.ComputerName = System.Windows.Forms.SystemInformation.ComputerName Then
                m_LoginData.BeginEdit()
                m_LoginData("LoggedIn") = Constants.C_dbFALSE
                m_LoginData("ExtraInfo") = ""
                m_LoginData.EndEdit()
                logoutResult = Me.UpdateTable()
            End If
            Return logoutResult
        End Function

        '''<summary>
        ''' Extra information to record
        ''' </summary>
        ''' <param name='Parameter'>Extra informatioin to record</param>
        Public Property ExtraInfo(ByVal Parameter As String) As String
            Get
                If m_ExtraInfo(Parameter) Is Nothing Then
                    Return ""
                Else
                    Return m_ExtraInfo(Parameter).ToString
                End If
            End Get
            Set(ByVal Value As String)
                If m_ExtraInfo(Parameter) Is Nothing Then
                    m_ExtraInfo.Add(Parameter, Value)
                Else
                    m_ExtraInfo(Parameter) = Value
                End If
                Dim index As Integer, info As String = ""
                For index = 0 To m_ExtraInfo.Count - 1
                    If m_ExtraInfo(index) > "" Then info &= ", " & m_ExtraInfo.Keys(index) & "=" & m_ExtraInfo(index)
                Next
                If info.Length > 2 Then info = info.Substring(2)
                m_LoginData.BeginEdit()
                m_LoginData("ExtraInfo") = info
                m_LoginData.EndEdit()
                Me.UpdateTable()
            End Set
        End Property

        '''<summary>Update information</summary>
        Private Function UpdateTable() As Boolean
            Dim updateOK As Boolean
            Try
                If CConnection.ConnOpen Then
                    If m_LoginData.RowState = DataRowState.Detached Then
                        m_ProgramUsage.Rows.Add(m_LoginData)
                    End If
                    If Not m_ProgramUsage.GetChanges Is Nothing Then
                        DataAdapter.Update(m_ProgramUsage.GetChanges)
                    End If
                    m_ProgramUsage.AcceptChanges()
                    updateOK = True
                End If
            Catch ex As System.Exception
                Console.WriteLine(ex.Message)
            Finally
                CConnection.SetConnClosed()
            End Try
            Return updateOK
        End Function

#End Region

    End Class
#End Region

#Region "UserException"
    <CLSCompliant(True)> _
  Public Class UserException
        Inherits MedscreenExceptions.MedscreenException

        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

    End Class
#End Region
End Class

#Region "ProgramUsage"

'''<summary>Class reflecting Program Usage table
''' </summary>
<CLSCompliant(True)> _
Public Class ProgramUsage
    Inherits SMTable

#Region "Declarations"


    Private objProgramName As StringField = New StringField("PROGRAMNAME", "", 30, True)
    Private objUserID As StringField = New StringField("USER_ID", "", 10, True)
    Private objLoginGroup As StringField = New StringField("LOGINGROUP", "", 100)
    Private objComputerName As StringField = New StringField("COMPUTERNAME", "", 10)
    Private objLastLogin As DateField = New DateField("LASTLOGIN")
    Private objLoggedIN As BooleanField = New BooleanField("LOGGEDIN", "F"c)
    Private objExtraInfo As StringField = New StringField("EXTRAINFO", "", 40)
#End Region

    '''<summary>
    ''' Creates a new Program Usage Object
    ''' </summary>
    Public Sub New()
        MyBase.New("PROGRAMUSAGE")
        MyBase.Fields.Add(objProgramName)
        MyBase.Fields.Add(objUserID)
        MyBase.Fields.Add(objLoginGroup)
        MyBase.Fields.Add(objComputerName)
        MyBase.Fields.Add(objLastLogin)
        MyBase.Fields.Add(objLoggedIN)
        MyBase.Fields.Add(objExtraInfo)

    End Sub

    '''<summary>
    ''' Name of the program in use
    ''' </summary>
    Public Property ProgramName() As String
        Get
            Return Me.objProgramName.StringValue
        End Get
        Set(ByVal Value As String)
            objProgramName.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Sample Manager User ID
    ''' </summary>
    Public Property UserId() As String
        Get
            Return Me.objUserID.StringValue
        End Get
        Set(ByVal Value As String)
            Me.objUserID.Value = Value
        End Set
    End Property

    '''<summary>
    ''' A parameter that may describe a Directory 
    ''' where information can be found or the group of programs 
    ''' this one belongs to
    ''' </summary>
    Public Property LoginGroup() As String
        Get
            Return Me.objLoginGroup.StringValue
        End Get
        Set(ByVal Value As String)
            Me.objLoginGroup.Value = Value

        End Set
    End Property

    '''<summary>TRUE = User is logged in to this program</summary>
    Public Property LoggedIn() As Boolean
        Get
            Return CBool(Me.objLoggedIN.Value)
        End Get
        Set(ByVal Value As Boolean)
            Me.objLoggedIN.Value = Value
        End Set
    End Property

    '''<summary>Date of last login</summary>
    Public Property LastLogin() As Date
        Get
            Return Me.objLastLogin.Value
        End Get
        Set(ByVal Value As Date)
            Me.objLastLogin.Value = Value
        End Set
    End Property

    '''<summary>Name of computer user is logged in to</summary>
    Public Property ComputerName() As String
        Get
            Return Me.objComputerName.StringValue
        End Get
        Set(ByVal Value As String)
            Me.objComputerName.Value = Value
        End Set
    End Property

    '''<summary>Additional information used in the management of the program</summary>
    Public Property ExtraInfo() As String
        Get
            Return Me.objExtraInfo.StringValue
        End Get
        Set(ByVal Value As String)
            Me.objExtraInfo.Value = Value
        End Set
    End Property

    '''<summary>Update information in the database</summary>
    ''' <returns>TRUE if succesful</returns>
    Public Function Update() As Boolean
        Return MyBase.DoUpdate
    End Function

    '''<summary></summary>
    Friend Property DataFields() As TableFields
        Get
            Return MyBase.Fields
        End Get
        Set(ByVal Value As TableFields)
            MyBase.Fields = Value
        End Set
    End Property

End Class

#End Region

#Region "ProgramUsageList"
'''<summary>A collection of ProgramUsage objects <see cref="ProgramUsage" /></summary>
<CLSCompliant(True)> _
Public Class ProgramUsageList
    Inherits SMTableCollection

#Region "Declarations"


    Private objProgramName As StringField = New StringField("PROGRAMNAME", "", 30, True)
    Private objUserID As StringField = New StringField("USER_ID", "", 10, True)
    Private objLoginGroup As StringField = New StringField("LOGINGROUP", "", 100)
    Private objComputerName As StringField = New StringField("COMPUTERNAME", "", 10)
    Private objLastLogin As DateField = New DateField("LASTLOGIN")
    Private objLoggedIN As BooleanField = New BooleanField("LOGGEDIN", "F"c)
    Private objExtraInfo As StringField = New StringField("EXTRAINFO", "", 40)

#End Region

    '''<summary>
    ''' Create a new collection of Program Usage objects
    ''' </summary>
    Public Sub New()
        MyBase.New("PROGRAMUSAGE")
        MyBase.Fields.Add(objProgramName)
        MyBase.Fields.Add(objUserID)
        MyBase.Fields.Add(objLoginGroup)
        MyBase.Fields.Add(objComputerName)
        MyBase.Fields.Add(objLastLogin)
        MyBase.Fields.Add(objLoggedIN)
        MyBase.Fields.Add(objExtraInfo)
    End Sub

    '''<summary>
    ''' Add a new item to list
    ''' </summary>
    ''' <param name='Item'>ProgramUsage object to add</param>
    ''' <returns>Position of added item in list</returns>
    Public Function Add(ByVal Item As ProgramUsage) As Integer
        Return MyBase.List.Add(Item)
    End Function

    '''<summary>
    ''' Get a collection of programs belonging to user
    ''' </summary>
    ''' <param name='User'>Sample Manager User Id to get the list of progams for</param>
    ''' <returns>TRUE if succesful</returns>
    Public Function LoadByUser(ByVal User As String) As Boolean
        Dim blnRet As Boolean = False
        Dim strQuery As String = MyBase.Fields.SelectString & " where user_id = '" & User.Trim.ToUpper & "'"

        Dim oCmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(strQuery, MedscreenLib.MedConnection.Connection)
        Dim oRead As OleDb.OleDbDataReader = Nothing

        Try
            If MedscreenLib.CConnection.ConnOpen Then
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Dim objPUsage As ProgramUsage = New ProgramUsage()
                    objPUsage.DataFields.readfields(oRead)
                    Me.Add(objPUsage)
                End While
            End If
            blnRet = True
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        Finally
            CConnection.SetConnClosed()
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
        End Try
        Return blnRet
    End Function
End Class

#End Region

