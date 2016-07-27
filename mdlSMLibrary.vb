Option Strict On
Imports System.Runtime.InteropServices


#Region "Database Connection"
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : MedConnection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''   Holds information about a database connection
''' </summary>
''' <remarks>
'''   Once initialised, connection information cannot be changed.  This ensures
'''   that the shared Instance property always returns the active connection and
'''   that the connection details have not been altered.<p/>
'''   For connections not automatically picked up from TNSNames.Ora files, new
'''   MedConnection objects may be created to add to the Connections collection.
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class MedConnection

#Region "Shared"
  Private Shared WithEvents _connection As OleDb.OleDbConnection
  Private Shared _medCon As MedConnection
  Private Shared _connectionType As Type
  Private Shared _innerState As Reflection.PropertyInfo

    Private Shared myErrors As String = ""
    Public Shared Property Errors() As String
        Get
            Return myErrors
        End Get
        Set(ByVal value As String)
            myErrors = value
        End Set
    End Property


  Public Shared Function GetRealState(ByVal dbconnection As Data.OleDb.OleDbConnection) As ConnectionState
    If _connectionType Is Nothing Then
      _connectionType = GetType(Data.OleDb.OleDbConnection)
      _innerState = _connectionType.GetProperty("StateInternal", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.NonPublic)
    End If
    If dbconnection Is Nothing Then
      Return ConnectionState.Broken
    ElseIf _innerState Is Nothing Then
      Return dbconnection.State
    Else
      Return DirectCast(_innerState.GetValue(dbconnection, Nothing), ConnectionState)
    End If
    End Function

    Public Shared Function GetRealState(ByVal dbconnection As Data.SqlClient.SqlConnection) As ConnectionState
        If _connectionType Is Nothing Then
            _connectionType = GetType(Data.OleDb.OleDbConnection)
            _innerState = _connectionType.GetProperty("StateInternal", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.NonPublic)
        End If
        If dbconnection Is Nothing Then
            Return ConnectionState.Broken
        ElseIf _innerState Is Nothing Then
            Return dbconnection.State
        Else
            Return DirectCast(_innerState.GetValue(dbconnection, Nothing), ConnectionState)
        End If
    End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a collection of all available connections
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   When first initialised, all available Oracle paths are scanned for TNSNames.ora
  '''   files and each of these is parsed in order to determine available connections.<p/>
  '''   The oracle locations are scanned in the order they appear in the registry.<p/>
  '''   Other connections may be added to the collection as required, but must have a unique name.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [03/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared ReadOnly Property Connections() As MedConnection.Collection
    Get
      Return MedConnection.Collection.Connections
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the OleDBConnection for the current MedConnection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   If no instance has been specified by setting the Instance property or
  '''   calling SetInstance, the first available connection is used.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared ReadOnly Property Connection() As OleDb.OleDbConnection
    Get
      If _connection Is Nothing Then
        Try
          If Not Instance Is Nothing Then
            _connection = New OleDb.OleDbConnection(Instance.ConnectionString)
            'Medscreen.LogAction("Connecting to " & Instance.FullName & " using " & Instance.ConnectionString)
          Else
            Medscreen.LogError("Connection requested before Instance set", False)
          End If
        Catch
        End Try
      End If
      Return _connection
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Clones the specified connection
  ''' </summary>
    ''' <returns>A new connection with a connection string matching that of the passed connection</returns>
  ''' <remarks>
  '''   Any cloned connections must be properly tidied up after use to prevent the maximum
  '''   number of connections to the database being exceeded.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function CloneConnection(ByVal instance As MedConnection) As OleDb.OleDbConnection
    Return New OleDb.OleDbConnection(instance.ConnectionString)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets or sets the MedConnection class for the current database connection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   When setting the instance, any existing connection is closed and dropped
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Property Instance() As MedConnection
    Get
      Return _medCon
    End Get
    Set(ByVal Value As MedConnection)
      MedConnection.Close(_connection)
      _medCon = Value
      _connection = Nothing
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the current instance to the one with the specified name
  ''' </summary>
  ''' <param name="InstanceName"></param>
  ''' <returns>False if name is not found.  Note that connection will revert to default.</returns>
  ''' <remarks>  
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function SetInstance(ByVal instanceName As String) As Boolean
    Instance = Connections.Item(instanceName)
    Return Not Instance Is Nothing
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the instance name as reported by the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  '''   Instance name is stored in the Config_Header table under "DB_INSTANCE"
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function ReportedInstance() As String
    Try
      Return New Glossary.ConfigItem("DB_INSTANCE").DefaultValue
        Catch ex As System.Exception
            Errors += ex.ToString
            Medscreen.LogError(ex)
            Return ""
    End Try
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Attempts to open a connection to the current instance
  ''' </summary>
  ''' <returns>True if connection is open upon return</returns>
  ''' <remarks>
  '''   The function will only return True if the connection is open and available (i.e. not executing).
  '''   It should never beassumed that a call to this function will set the connection open, the
  '''   return value should always be checked.
  '''   <para>By default the connection will timeout if not opened within 10 seconds</para>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function Open() As Boolean
    Return Open(MedConnection.Connection)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Attempts to open the passed database connection
  ''' </summary>
    ''' <returns>True if connection is open upon return</returns>
  ''' <remarks>
  '''   The function will only return True if the connection is open and available (i.e. not executing).
  '''   It should never beassumed that a call to this function will set the connection open, the
  '''   return value should always be checked.
  '''   <para>By default the connection will timeout if not opened within 10 seconds</para>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function Open(ByVal dbConnection As OleDb.OleDbConnection) As Boolean
    Return Open(dbConnection, 10000)
    End Function

    Public Shared Function Open(ByVal dbConnection As SqlClient.SqlConnection) As Boolean
        Return Open(dbConnection, 10000)
    End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Attempts to open the passed database connection
  ''' </summary>
    ''' <param name="dbConnection">The connection to open</param>
  ''' <param name="timout">Time to wait for connection to be available in milliseconds</param>
  ''' <returns>True if connection is open upon return</returns>
  ''' <remarks>
  '''   The function will only return True if the connection is open and available (i.e. not executing).
  '''   It should never beassumed that a call to this function will set the connection open, the
  '''   return value should always be checked.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function Open(ByVal dbConnection As OleDb.OleDbConnection, ByVal timout As Integer) As Boolean
    If dbConnection Is Nothing Then
      Return False
    Else
      Dim startTime As DateTime = DateTime.Now
      Dim realState As ConnectionState = GetRealState(dbConnection)
      If (realState And ConnectionState.Open) = 0 Then
        Try
          dbConnection.Open()
        Catch ex As System.Exception
                    Debug.WriteLine(ex.Message)
                    Errors += ex.ToString
        End Try
        realState = GetRealState(dbConnection)
      End If
      If realState <> ConnectionState.Open Then
        ' Action in progress, wait for it
        While realState <> ConnectionState.Open AndAlso startTime.AddMilliseconds(timout) >= Now
          Threading.Thread.Sleep(100)
          realState = GetRealState(dbConnection)
        End While
      End If
      Return realState = ConnectionState.Open
    End If
  End Function


    Public Shared Function Open(ByVal dbConnection As SqlClient.SqlConnection, ByVal timout As Integer) As Boolean
        If dbConnection Is Nothing Then
            Return False
        Else
            Dim startTime As DateTime = DateTime.Now
            Dim realState As ConnectionState = GetRealState(dbConnection)
            If (realState And ConnectionState.Open) = 0 Then
                Try
                    dbConnection.Open()
                Catch ex As System.Exception
                    Errors += ex.ToString
                    Debug.WriteLine(ex.Message)
                End Try
                realState = GetRealState(dbConnection)
            End If
            If realState <> ConnectionState.Open Then
                ' Action in progress, wait for it
                While realState <> ConnectionState.Open AndAlso startTime.AddMilliseconds(timout) >= Now
                    Threading.Thread.Sleep(100)
                    realState = GetRealState(dbConnection)
                End While
            End If
            Return realState = ConnectionState.Open
        End If
    End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Attempts to close the current connection to the current instance
  ''' </summary>
  ''' <returns>True if connection is closed upon return</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function Close() As Boolean
    Return Close(MedConnection.Connection)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Attempts to close the passes connection
  ''' </summary>
    ''' <param name="dbConnection">The connection to close</param>
  ''' <returns>True if connection is closed upon return</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function Close(ByVal dbConnection As OleDb.OleDbConnection) As Boolean
    If dbConnection Is Nothing Then
      Return False
    Else
      If dbConnection.State <> ConnectionState.Closed Then
        Try
          dbConnection.Close()
                Catch ex As System.Exception
                    Errors += ex.ToString
                    Debug.WriteLine(ex.Message)
        End Try
      End If
      Return dbConnection.State = ConnectionState.Closed
    End If
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Show a dialog allowing user to select database instance to use
  ''' </summary>
  ''' <remarks>
  '''   Dialog automatically displayes an entry for each member of the useDatabase enumeration.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function ShowDBDialog() As Boolean
    Dim dlgDBToUse As New frmCombo()
    dlgDBToUse.Text = "Select database instance"
    Dim oraCon As MedConnection
    For Each oraCon In MedConnection.Connections
      dlgDBToUse.cmbOptions.Items.Add(oraCon.TidyName)
    Next
    If dlgDBToUse.ShowDialog() = Windows.Forms.DialogResult.OK Then
      MedConnection.SetInstance(dlgDBToUse.cmbOptions.Text)
      Return True
    Else
      Return False
    End If
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the source column for each parameter to match the parameter name
  ''' </summary>
  ''' <param name="Parameters"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Sub SetParameterSourceColumns(ByVal Parameters As OleDb.OleDbParameterCollection)
    Dim myparameter As OleDb.OleDbParameter
    For Each myparameter In Parameters
      myparameter.SourceColumn = myparameter.ParameterName
    Next
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Log info messages from data connection
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="args"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Shared Sub OnInfoMessage(ByVal sender As Object, ByVal args As OleDb.OleDbInfoMessageEventArgs) Handles _connection.InfoMessage
    Dim err As OleDb.OleDbError
    For Each err In args.Errors
      Medscreen.LogError("The " & err.Source & " has received a SQLState " & err.SQLState & _
                         " error number " & err.NativeError & ":" & ControlChars.NewLine & err.Message, False)
    Next
  End Sub

#End Region

#Region "Instance"
  Private _name As String, _extra As String, _host As String, _SID As String, _pwd As String, _user As String, _provider As String
  Private _source As String, _path As String

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a new MedSconnection to be populated later
  ''' </summary>
  ''' <param name="name"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Sub New(ByVal name As String)
    If name Is Nothing Then
      Throw New ArgumentNullException("name")
    End If
    _provider = "MSDAORA.1"
    Dim dotPos As Integer = name.IndexOf(".")
    If dotPos > 0 Then
      _name = name.Substring(0, dotPos).ToUpper
      _extra = name.Substring(dotPos).ToUpper
    Else
      _name = name.Trim.ToUpper
    End If
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Create a new MedConnection for the specified Host and SID
  ''' </summary>
  ''' <param name="name"></param>
    ''' <param name="provider"/>
  ''' <param name="host"></param>
  ''' <param name="SID"></param>
  ''' <param name="user"></param>
  ''' <param name="password"></param>
  ''' <remarks>
  '''   Connection string is formatted as
  '''   Provider=(Provider);Password=(Password);User ID=(User);Data Source=(SID)-(Host)[.Extra]<p/>
  '''   Where Extra is any part of the name after a . (full stop) character.<p/>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub New(ByVal name As String, ByVal provider As String, ByVal host As String, ByVal SID As String, ByVal user As String, ByVal password As String)
    Me.New(name)
    _host = host
    _SID = SID
    _user = user
    _pwd = password
    If Not provider Is Nothing AndAlso provider.Trim > "" Then
      _provider = provider
    End If
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Create a new MedConnection for the specified data source
  ''' </summary>
  ''' <param name="name"></param>
    ''' <param name="provider"><b>N.B.<b/>If empty string or Nothing, defaults to MSDORA.1</b></param>
  ''' <param name="dataSource"></param>
  ''' <param name="user"></param>
  ''' <param name="password"></param>
  ''' <remarks>
  '''   Connection string is formatted as
  '''   Provider=(Provider);Password=(Password);User ID=(User);Data Source=(dataSource)<p/>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub New(ByVal name As String, ByVal provider As String, ByVal dataSource As String, ByVal user As String, ByVal password As String)
    Me.New(name)
    If Not provider Is Nothing AndAlso provider.Trim > "" Then
      _provider = provider
    End If
    _source = dataSource
    _user = user
    _pwd = password
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the name of the connection as defined in TNSNames.Ora
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Name() As String
    Get
      Return _name
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the name of the connection as defined in TNSNames.Ora
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property FullName() As String
    Get
      Return _name & _extra
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the name of the connection as defined in TNSNames.Ora
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property TidyName() As String
    Get
      Return _name & StrConv(_extra, VbStrConv.ProperCase)
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the name of the host (server) for this connection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Host() As String
    Get
      Return _host
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the host (server) to use for this connection
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  '''   Note that if DataSource has been specified in constructure, Host and SID properties are ignored.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Sub SetHost(ByVal value As String)
    _host = value
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the SID for this connection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property SID() As String
    Get
      Return _SID
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the SID (instance name) to use for this connection
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  '''   Note that if DataSource has been specified in constructure, Host and SID properties are ignored.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Sub SetSID(ByVal value As String)
    _SID = value
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the user name in use for this connection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   Defaults to name of SID, unless specified by setting the property first
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Property UserName() As String
    Get
      If _user Is Nothing Then
        Return _SID
      Else
        Return _user
      End If
    End Get
    Set(ByVal value As String)
      _user = value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the password to use for this connection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   Password defaults to SID unless specified using SetPassword
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public WriteOnly Property Password() As String
    Set(ByVal Value As String)
      _pwd = Value
    End Set
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the password to use for this connection
  ''' </summary>
    ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
    Private Function GetPassword() As String
        If _pwd Is Nothing Then
            Return Me.SID
        Else
            Return _pwd
        End If
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets the provider for this connection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Defaults to MSDAORA.1
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Provider() As String
        Get
            Return _provider
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the data source name
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   This is used in the connection string rather than the name so that names
    '''   connections can be created easily.<p/>
    '''   Return value is Host-SID
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DataSource() As String
        Get
            If _source Is Nothing OrElse _source.Trim = "" Then
                Return Me.Host & "-" & Me.SID & _extra
            Else
                Return _source
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the connection string to use when making this connection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Taylor]</Author><date> [19/10/2007]</date><Action>Made compliant with returning datasets</Action></revision>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property ConnectionString() As String
        Get
            Return "Provider=" & Me.Provider() & ";OLEDB.NET=true;PLSQLRSet=true" & _
                    ";Password=" & Me.GetPassword() & _
                    ";User ID=" & Me.UserName() & _
                    ";Data Source=" & Me.DataSource
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the network path to the host for this connection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Path is assumed to be \\Host\SID<p/>
    '''   Where DataSource has been specified rather than Host and SMID, Nothing is returned.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ServerPath() As String
        Get
            'If _source Is Nothing OrElse _source.Trim > "" Then
            '  Return "\\" & Me.Host & "\" & Me.SID
            'End If
            If _path Is Nothing Then
                Dim getPath As New OleDb.OleDbCommand("SELECT Value FROM Config_Header WHERE Identity = ?", _
                                                      New OleDb.OleDbConnection(Me.ConnectionString))
                getPath.Parameters.AddWithValue("Config_ID", "SYSTEM_ROOT_PATH")
                Try
                    MedConnection.Open(getPath.Connection)
                    _path = getPath.ExecuteScalar.ToString()
                    MedConnection.Close(getPath.Connection)
                Catch ex As Exception
                    Errors += ex.ToString
                    Medscreen.LogError(ex)
                End Try
                getPath.Connection.Dispose()
                getPath.Dispose()
            End If
            Return _path
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Mekes this connection the active one
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetActive()
        MedConnection.Instance = Me
    End Sub

#End Region

#Region "Collection"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : OracleInstances.MedConnection.Collection
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Collection of oracle connections
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [03/08/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class Collection
    Inherits ReadOnlyCollectionBase

    Private Shared _connections As MedConnection.Collection
    Private Shared trimChars As Char() = New Char() {" "c, "("c, ")"c, "="c}

    Shared Sub New()
      _connections = New MedConnection.Collection()
      ' Read all available TNSNames.Ora files to find available connections
      GetHomes()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the global list of available Oracle connections
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Connections() As MedConnection.Collection
      Get
        Return _connections
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Finds a list of all available oracle paths from the registry
    ''' </summary>
    ''' <remarks>
    '''   The TNSNames.Ora file for each path is passed and connections added to the
    '''   public shared collection accessible through the Connections property
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub GetHomes()
      'Debug.WriteLine("Opening registry...")
      Try
        Dim oraKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Oracle\All_Homes", False)
        Dim homeKeys() As String
        Dim paths As New ArrayList()
        homeKeys = oraKey.GetSubKeyNames
        Dim homeKeyName As String
        For Each homeKeyName In homeKeys
          Dim homeKey As Microsoft.Win32.RegistryKey = oraKey.OpenSubKey(homeKeyName, False)
          'Debug.WriteLine("Found oracle path " & homeKey.GetValue("Path").ToString)
          paths.Add(homeKey.GetValue("Path").ToString & "\Network\Admin\TNSNames.Ora")
          homeKey.Close()
        Next
        oraKey.Close()
        For Each homeKeyName In paths
          ParseFile(homeKeyName)
        Next
            Catch ex As Exception
                Errors += ex.ToString
                Medscreen.LogError(ex, , "Getting registry info")
      End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Parses a TNSNames.Ora file and creates MedConnection objects for each connection
    ''' </summary>
    ''' <param name="path"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [07/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub ParseFile(ByVal path As String)
      Try
        If IO.File.Exists(path) Then
          'Debug.WriteLine("Parsing " & path)
          Dim definitions As String
          With IO.File.OpenText(path)
            definitions = .ReadToEnd
            .Close()
          End With
          ' Look for line with 1st char as alphabetic - start of section
          ' Count all open brackets & close brackets until they're equal - end of section

          ' Set up regular expressions:
          '   Find definitions start points as Alpha char preceeded by new line
          Dim findStart As New System.Text.RegularExpressions.Regex("\n[A-Z]")
          '   Find either open or close brackets
          Dim findBracket As New Text.RegularExpressions.Regex("\(|\)")
          '   Find text between Open Bracket and Equals sign
          Dim getKey As New Text.RegularExpressions.Regex("\([A-Za-z0-9 _\-\.]+=")
          '   Find text between Equals sign and Close Bracket 
          Dim getVal As New Text.RegularExpressions.Regex("=[A-Za-z0-9 _\-\.]+\)")

          ' Find definition start points
          Dim defStart As Text.RegularExpressions.Match
          Dim lastDefEnd As Integer = -1
          For Each defStart In findStart.Matches(definitions)
            ' If this definition does not start after the end of the previous one, ignore it
            If defStart.Index > lastDefEnd Then
              ' Include in a try block to catch any invalid definitions
              Try
                ' The service name will be everything up to the first space
                Dim oraCon As New MedConnection(MedscreenLib.ExtractStringSection(definitions, defStart.Index - 1, ControlChars.NewLine, " ").Trim.ToUpper)
                ' Now count each bracket after the name until all brackets have been closed (bracketCount = 0)
                Dim bracketCount As Integer
                Dim bracketPos As Text.RegularExpressions.Match = findBracket.Match(definitions, defStart.Index)
                Do
                  If Not bracketPos Is Nothing Then
                    ' For open brackets, pick up the Key (between bracket and = sign) and see if it's useful
                    If bracketPos.Value = "(" Then
                      ' Open bracket found, increment bracket counter
                      bracketCount += 1
                      Dim key As String = getKey.Match(definitions, bracketPos.Index).Value.Trim(trimChars)
                      If Not key Is Nothing Then
                        ' If a key holds useful information, pick up is value (between = sign and close bracket)
                        Select Case key.ToUpper
                          Case "HOST"
                            oraCon.SetHost(getVal.Match(definitions, bracketPos.Index).Value.Trim(trimChars).ToUpper)
                          Case "SID"
                            oraCon.SetSID(getVal.Match(definitions, bracketPos.Index).Value.Trim(trimChars).ToUpper)
                        End Select
                      End If
                    Else
                      ' Close bracket found, decrement bracket counter
                      bracketCount -= 1
                    End If
                    ' Record the position in the definitions so we know where to start looking for the next definition
                    ' This is a safety check, in case file contains lines mid-definition with characters in position 1.
                    lastDefEnd = bracketPos.Index
                    ' Find next bracket
                    bracketPos = findBracket.Match(definitions, bracketPos.Index + 1)
                  End If
                Loop Until bracketCount = 0 Or bracketPos Is Nothing
                If oraCon.SID <> "" Then
                                    Debug.WriteLine("Found " & oraCon.Name & " Name=" & oraCon.FullName & " Host=" & oraCon.Host & " SID=" & oraCon.SID)
                  _connections.Add(oraCon)
                End If
              Catch
                Debug.WriteLine("Invalid definition found!")
              End Try
            End If
          Next
        End If
            Catch ex As Exception
                Errors += ex.ToString
                Medscreen.LogError(ex, , "Parse File")
      End Try
    End Sub

    Private Sub New()
      MyBase.New()
    End Sub

    Public Sub Add(ByVal connection As MedConnection)
      If (Not connection Is Nothing) AndAlso _
          Me.MatchingItem(connection.FullName) Is Nothing Then
        MyBase.InnerList.Add(connection)
      Else
        Debug.WriteLine("Duplicate connection " & connection.FullName)
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the connection with the matching fullName
    ''' </summary>
    ''' <param name="fullName"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function MatchingItem(ByVal fullName As String) As MedConnection
            Dim oraCon As MedConnection = Nothing

            If Not fullName Is Nothing Then fullName = fullName.ToUpper
            Dim fullNameLocal As String = fullName & ".MEDSCREEN.LOCAL"
            For Each oraCon In Me
                If oraCon.FullName = fullName OrElse oraCon.FullName = fullnameLocal Then
                    Exit For
                End If
                oraCon = Nothing
            Next
            Return oraCon
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the connection with the matching Fullname or Name
    ''' </summary>
    ''' <param name="connectionName"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal connectionName As String) As MedConnection
      Get
        If Not connectionName Is Nothing Then connectionName = connectionName.ToUpper
        Dim matchingCon As MedConnection = Me.MatchingItem(connectionName)
        If matchingCon Is Nothing Then
          Dim oraCon As MedConnection
          For Each oraCon In Me
            If oraCon.Name = connectionName Then
              ' Partial match, store but keep searching
              matchingCon = oraCon
            End If
          Next
        End If
        Return matchingCon
      End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As MedConnection
      Get
        Return CType(MyBase.InnerList.Item(index), MedConnection)
      End Get
    End Property

  End Class
#End Region

End Class
#End Region

#Region "Sys Utils"
Namespace SysUtils
  Public Module SysUtils

    '/// this is the windows api call that opens the sound file ( NOT MP3 files though )
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Int32, ByVal dwFlags As Int32) As Int32

    <Flags()> _
    Public Enum PlaySoundFlags As Integer
      SND_SYNC = 0
      SND_ASYNC = 1
      SND_FILENAME = &H20000
      SND_RESOURCE = &H40004
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Plays the specified sound file
    ''' </summary>
    ''' <param name="filename">Name of file to play</param>
    ''' <param name="async">True to play sound in background</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub PlaySoundFile(ByVal filename As String, ByVal async As Boolean)
      Dim flags As Integer = PlaySoundFlags.SND_FILENAME
      If async Then flags = flags Or PlaySoundFlags.SND_ASYNC
      PlaySound(filename, Nothing, flags)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Plays the sound from the specified resource
    ''' </summary>
    ''' <param name="resName">Name of resource to play</param>
    ''' <param name="appHandle">Handle to application owning resource</param>
    ''' <param name="async">True to play sound in background</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub PlaySoundRes(ByVal resName As String, ByVal appHandle As Integer, ByVal async As Boolean)
      Dim flags As Integer = PlaySoundFlags.SND_RESOURCE
      If async Then flags = flags Or PlaySoundFlags.SND_ASYNC
      PlaySound(resName, appHandle, flags)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Struct	 : SysUtils.SysUtils.SYSTEMTIME
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   System time structure used to pass to SetLocalTime etc
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Private Structure SYSTEMTIME
      Public year As Short
      Public month As Short
      Public dayOfWeek As Short
      Public day As Short
      Public hour As Short
      Public minute As Short
      Public second As Short
      Public milliseconds As Short
    End Structure

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Declaration or kernel32 SetLocalTime function
    ''' </summary>
    ''' <param name="time"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <DllImport("kernel32.dll")> _
    Private Function SetLocalTime(ByRef time As SYSTEMTIME) As Boolean
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets the system (local) time to the specified value
    ''' </summary>
    ''' <param name="newDate"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetSystemDate(ByVal newDate As DateTime)
      'Populate structure... 
      Dim st As SYSTEMTIME
      st.year = CType(newDate.Year, Short)
      st.month = CType(newDate.Month, Short)
      st.dayOfWeek = CType(newDate.DayOfWeek, Short)
      st.day = CType(newDate.Day, Short)
      st.hour = CType(newDate.Hour, Short)
      st.minute = CType(newDate.Minute, Short)
      st.second = CType(newDate.Second, Short)
      st.milliseconds = CType(newDate.Millisecond, Short)
      'Set the new time... 
      SetLocalTime(st)
    End Sub

  End Module
End Namespace
#End Region

#Region "Common"
Public Module Common

  Public Const C_dbTRUE As Char = "T"c, C_dbFALSE As Char = "F"c
  Public Const C_InitDate As Date = #1/1/1850#

  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Struct	 : Common.IDValuePair
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Structure for containing key - value pairs.
  ''' </summary>
  ''' <remarks>
  '''   This differs from the inbuild classes in that the ToString method returns the value
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Structure IDValuePair
    Private myID As String, myValue As String

    Public Sub New(ByVal ID As String, ByVal Value As String)
      myID = ID
      myValue = Value
    End Sub

    Public ReadOnly Property ID() As String
      Get
        Return myID
      End Get
    End Property

    Public ReadOnly Property Value() As String
      Get
        Return myValue
      End Get
    End Property

    Public Overrides Function ToString() As String
      Return myValue
    End Function

  End Structure

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the name of the entry (start) application 
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property ApplicationName() As String
    Get
      'Return System.Reflection.Assembly.GetCallingAssembly.GetName.Name
      Return System.Reflection.Assembly.GetEntryAssembly.GetName.Name
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Applies the specified discound to the specified price
  ''' </summary>
  ''' <param name="Price">Price to be discounted</param>
  ''' <param name="DiscountPercent">Percentage discount to apply</param>
  ''' <returns>The new price</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function ApplyDiscount(ByVal Price As Single, ByVal DiscountPercent As Single) As Single
    Return (Price * (1 - (DiscountPercent / 100)))
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a value indicating whether an item is null, or only contains spaces
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function IsBlank(ByVal Value As Object) As Boolean
    Return Value Is Nothing OrElse _
            Value.ToString.Trim.Length = 0
  End Function

  Public Function PackedDecimal(ByVal value As Long) As String
    Return value.ToString.PadLeft(10, " "c)
  End Function

  Public Function NVL(ByVal value As Object, ByVal nullValue As Integer) As Integer
    Return CType(GetNVL(value, nullValue), Integer)
  End Function

  Public Function NVL(ByVal value As Object, ByVal nullValue As Double) As Double
    Return CType(GetNVL(value, nullValue), Double)
  End Function

  Public Function NVL(ByVal value As Object, ByVal nullValue As String) As String
    Return CType(GetNVL(value, nullValue), String)
  End Function

  Public Function NVL(ByVal value As Object, ByVal nullValue As Date) As Date
    Return CType(GetNVL(value, nullValue), Date)
  End Function

  Private Function GetNVL(ByVal value As Object, ByVal nullValue As Object) As Object
    If value Is DBNull.Value Then
      Return nullValue
    Else
      Return value
    End If
  End Function

    ''' <summary>
    '''   Compares two values based on an equality operator
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Function CheckEquality(ByVal equalityOp As String, ByVal value1 As IComparable, ByVal value2 As IComparable) As Boolean
        Return (equalityOp = ">" AndAlso value1.CompareTo(value2) > 0) OrElse _
               (equalityOp = ">=" AndAlso value1.CompareTo(value2) >= 0) OrElse _
               (equalityOp = "<" AndAlso value1.CompareTo(value2) < 0) OrElse _
               (equalityOp = "<=" AndAlso value1.CompareTo(value2) <= 0) OrElse _
               (equalityOp = "=" AndAlso value1.CompareTo(value2) = 0) OrElse _
               (equalityOp = "<>" AndAlso value1.CompareTo(value2) <> 0)
    End Function
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Indicates whether all characters in string are digits (0-9)
  ''' </summary>
  ''' <param name="text"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Integer, isDigit As Boolean
    If Not text Is Nothing Then
      For i = 0 To text.Length - 1
        isDigit = Char.IsDigit(text.Chars(i))
        If Not isDigit Then Exit For
      Next
    End If
    Return isDigit
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Indicates whether all characters in string are alphabetic (A-Z, a-z)
  ''' </summary>
  ''' <param name="text"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function IsAlphabetic(ByVal text As String) As Boolean
    Dim i As Integer, isAlpha As Boolean
    If Not text Is Nothing Then
      For i = 0 To text.Length - 1
        isAlpha = Char.IsLetter(text.Chars(i))
        If Not isAlpha Then Exit For
      Next
    End If
    Return isAlpha
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns True if none of the characters are alphabetic
  ''' </summary>
  ''' <param name="text"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   This is NOT the same as <code>Not IsAlphabetic(text)</code>.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function IsNonAlphabetic(ByVal text As String) As Boolean
    Dim i As Integer, isAlpha As Boolean
    If Not text Is Nothing Then
      For i = 0 To text.Length - 1
        isAlpha = Char.IsLetter(text.Chars(i))
        If isAlpha Then Exit For
      Next
    End If
    Return Not isAlpha
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Extract a portion of a string between start and end characters
  ''' </summary>
  ''' <param name="expression"></param>
  ''' <param name="startPos"></param>
  ''' <param name="startWith"></param>
  ''' <param name="endWith"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' If starWith or endWith are blank, double quotes will be assumed.<p/>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [25/05/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function ExtractStringSection(ByVal expression As String, ByVal startPos As Integer, ByVal startWith As String, ByVal endWith As String) As String
    If startWith = "" Then startWith = """"
    If endWith = "" Then endWith = """"
    Dim start As Integer = expression.IndexOf(startWith, startPos)
    Dim finish As Integer = expression.IndexOf(endWith, start + 1)
    If start >= 0 And finish > start Then
      Return expression.Substring(start + 1, finish - (start + 1))
    Else
      Return ""
    End If
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Calls FormatTypeString for an enumeration member
  ''' </summary>
  ''' <param name="enumMember"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [06/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function EnumToString(ByVal enumMember As [Enum]) As String
    Return FormatTypeName(enumMember.ToString("G"))
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Formats a type name for display
  ''' </summary>
  ''' <param name="typeName"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   Intended for formatting Enumeration members, each new word (starting with
  '''   a number or capital letter) is preceded by a space.  Underscores are also
  '''   converted to spaces.
  '''   <example>This example show the output for a daft enumeration
  '''   <code>
  '''   Private Enum TestString
  '''     MyABCTest = 1
  '''     Testing = 2
  '''     JustMe = 4
  '''     Life_And_Soul = 8
  '''     Entry1 = 16
  '''     Entry10 = 32
  '''     List1Type = 64
  '''     List10Type = 128
  '''   End Enum
  ''' 
  '''   Private Sub TestEnumFormatting
  '''     Dim temp As TestString
  '''     For Each temp In [Enum].GetValues(GetType(TestString))
  '''       Console.WriteLine(temp.ToString &amp; &quot; = &quot; &amp; FormatTypeName(temp.ToString))
  '''     Next
  '''     temp = TestString.Entry10 Or TestString.JustMe Or TestString.MyABCTest
  '''     Console.WriteLine(temp.ToString &amp; &quot; = &quot; &amp; FormatTypeName(temp.ToString))
  '''   End Sub
  ''' 
  '''   </code>
  '''   The result is:
  '''   <code>
  '''   MyABCTest = My ABC Test
  '''   Testing = Testing
  '''   JustMe = Just Me
  '''   Life_And_Soul = Life And Soul
  '''   Entry1 = Entry 1
  '''   Entry10 = Entry 10
  '''   List1Type = List 1 Type
  '''   List10Type = List 10 Type
  '''   MyABCTest, JustMe, Entry10 = My ABC Test, Just Me, Entry 10
  '''   </code>
  '''   </example>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [06/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function FormatTypeName(ByVal typeName As String) As String
    Dim name As New Text.StringBuilder(typeName)
    ' First repalce all underscores with spaces
    name.Replace("_", " ")
    ' Next place a space before any upper case characters to separate words.
    ' Start at second character (index 1), no point adding a space at the start!
    Dim charIndex As Integer = 1
    Dim prevChar As Char = name.Chars(0)
    While charIndex < name.Length
      Dim thisChar As Char = name.Chars(charIndex)
      Dim nextChar As Char
      If charIndex < name.Length - 1 Then
        nextChar = name.Chars(charIndex + 1)
      Else
        nextChar = " "c
      End If
      If Not Char.IsWhiteSpace(prevChar) AndAlso _
         ((Char.IsDigit(thisChar) AndAlso _
            Not Char.IsDigit(prevChar)) OrElse _
          (Char.IsUpper(thisChar) AndAlso _
           (Not Char.IsUpper(nextChar) OrElse _
            Not Char.IsUpper(prevChar)))) Then
        name.Insert(charIndex, " ")
        charIndex += 1
      End If
      prevChar = thisChar
      charIndex += 1
    End While
    Return name.ToString
  End Function


  Public Function YesNo(ByVal value As Boolean) As String
    Return YesNo(value, "Yes", "No")
  End Function

  Public Function YesNo(ByVal value As Boolean, ByVal yes As String, ByVal no As String) As String
    If value Then
      Return yes
    Else
      Return no
    End If
  End Function
    Public Function ToBoolean(ByVal value As Boolean) As String
        Return YesNo(value, C_dbTRUE, C_dbFALSE)
    End Function

    Public Function ToBoolean(ByVal value As String) As Boolean
        Return value = C_dbTRUE
    End Function

End Module
#End Region

#Region "QuickLists"
Namespace QuickLists
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : QuickLists.QuickLists
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Contins methods for retrieving summary information from the database
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Module QuickLists

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of counties as IDValuePairs of CountryID and Country Name
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   The methods in this module are intended for backward compatibility with older
    '''   .NET applications.  Most now have classes that can be used for retieving the required information.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetCountryList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT Country_ID, Country_Name FROM Country", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(New IDValuePair(rdrCountries!Country_ID.ToString, rdrCountries!Country_Name.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of ports as IDValuePairs of Identity and Site_Name
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetPortList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT Identity, Site_Name FROM Port " & _
                                                  "WHERE RemoveFlag = 'F'", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(New IDValuePair(rdrCountries!Identity.ToString, rdrCountries!Description.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of vessels as IDValuePairs of Vessel_ID and Name
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetVesselList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT Name, Vessel_ID FROM Vessel " & _
                                                  "WHERE RemoveFlag = 'F'", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(New IDValuePair(rdrCountries!Vessel_ID.ToString, rdrCountries!Name.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a (string) list of Regions from the Country table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetRegionList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT DISTINCT Region FROM Country", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(rdrCountries!Region)
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of Line Item Groups as IDValuePairs of GroupID and Description
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetItemGroupList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT ID, Description FROM Lineitem_Groups", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(New IDValuePair(rdrCountries!ID.ToString, rdrCountries!Description.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Get a list of Line Item Types as strings
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetItemTypeList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT DISTINCT Type FROM Lineitems", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(rdrCountries!Type)
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of Line Items as IDValuePairs of ItemCode and Description
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetItemCodeList() As ArrayList
      Dim myList As New ArrayList()
      Dim cmdCountries As New OleDb.OleDbCommand("SELECT ItemCode, Description FROM Lineitems", MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrCountries As OleDb.OleDbDataReader = cmdCountries.ExecuteReader
        While rdrCountries.Read
          myList.Add(New IDValuePair(rdrCountries!ItemCode.ToString, rdrCountries!ItemCode.ToString & " - " & rdrCountries!Description.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the prase entries as IDValuePairs of Phrase_ID and Phrase_Text for the given phrase
    ''' </summary>
    ''' <param name="PhraseID"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetPhraseEntries(ByVal PhraseID As String) As ArrayList
      Dim myList As New ArrayList()
      Dim cmdPhrase As New OleDb.OleDbCommand("SELECT Phrase_ID, Phrase_Text " & _
                                              "FROM Phrase WHERE Phrase_Type = '" & PhraseID & "' " & _
                                              "ORDER BY Order_Num", _
                                              MedConnection.Connection)
      MedConnection.Open()
      Try
        Dim rdrPhrase As OleDb.OleDbDataReader = cmdPhrase.ExecuteReader
        While rdrPhrase.Read
          myList.Add(New IDValuePair(rdrPhrase!Phrase_ID.ToString, rdrPhrase!Phrase_Text.ToString))
        End While
      Catch
      Finally
        MedConnection.Close()
      End Try
      Return myList
    End Function

  End Module
End Namespace
#End Region

#Region "IndexedCollection"

Public Class IndexedCollection
  Inherits IndexedCollectionBase

  Public Sub New(ByVal index As String)
    MyBase.New(index)
  End Sub

  Public Sub Add(ByVal value As Object)
    MyBase.BaseAdd(value)
  End Sub

  Public Sub AddUnindexed(ByVal value As Object)
    MyBase.BaseAddUnindexed(value)
  End Sub

  Public Sub Remove(ByVal value As Object)
    MyBase.BaseRemove(value)
  End Sub

  Public Shadows Sub RemoveAt(ByVal index As Integer)
    MyBase.RemoveAt(index)
  End Sub

  Public ReadOnly Property IndexList() As IndexedCollection.Index.Collection
    Get
      Return MyBase.Indexes
    End Get
  End Property

End Class

#Region "Base"
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : IndexedCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''   An unsorted collection, accessible by index key or position
''' </summary>
''' <remarks>
'''   Each index contains unique values that reference only one item in the collection.
'''   Each group may reference more than one item in the collection.
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Boughton]</Author><date> [24/11/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public MustInherit Class IndexedCollectionBase
  Inherits CollectionBase

  Protected Event ValueInserted(ByVal sender As Object, ByVal e As InsertRemoveEventArgs)
  Protected Event ValueRemoved(ByVal sender As Object, ByVal e As InsertRemoveEventArgs)

  Private _indexes As Index.Collection, _primaryIndex As String

  Public Sub New()
    MyBase.New()
    _indexes = New Index.Collection()
  End Sub

  Public Sub New(ByVal index As String)
    Me.New()
    _indexes.Add(index)
    _primaryIndex = index
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the collection of indexes defined for this collection
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  '''   Each index relates to a property for the objects in the collection
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected ReadOnly Property Indexes() As Index.Collection
    Get
      Return _indexes
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Adds the specified object to the collection
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  '''   All indexes are automatically populated with appropriate values taken from the object.
  '''   The object must have properties to match the name of each defined for this collection.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overridable Sub BaseAdd(ByVal value As Object)
    MyBase.List.Add(value)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Adds the specified object to the collection, without updating indexes
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  '''   This is useful where the object is new and does not have it's details populated.<p/>
  '''   Note that calling ReIndex <b>will</b> update the index values for this object and so
  '''   should not be called until the object has been initialised.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [04/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overridable Sub BaseAddUnindexed(ByVal value As Object)
    MyBase.InnerList.Add(value)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Removes the specified object from the collection
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overridable Sub BaseRemove(ByVal value As Object)
    MyBase.List.Remove(value)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Removes the object with the specified key in the specified index
  ''' </summary>
  ''' <param name="indexName"></param>
  ''' <param name="key"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overridable Sub BaseRemove(ByVal indexName As String, ByVal key As Object)
    Dim keyIndex As Index = _indexes.Item(indexName)
    If Not keyIndex Is Nothing Then
      MyBase.List.RemoveAt(CInt(keyIndex.Item(key)))
    End If
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates index references to object when object inserted
  ''' </summary>
  ''' <param name="index"></param>
  ''' <param name="value"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overrides Sub OnInsertComplete(ByVal index As Integer, ByVal value As Object)
    MyBase.OnInsertComplete(index, value)
    Me.InsertIndexValue(value, index)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Removes index references to object when object removed
  ''' </summary>
  ''' <param name="index"></param>
  ''' <param name="value"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overrides Sub OnRemoveComplete(ByVal index As Integer, ByVal value As Object)
    MyBase.OnRemoveComplete(index, value)
    Me.RemoveIndexValue(value)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Clears all indexes when collection is cleared
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overrides Sub OnClearComplete()
    MyBase.OnClearComplete()
    Dim keyIndex As Index
    For Each keyIndex In _indexes
      keyIndex.Clear()
    Next
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Updates indexes when object changed
  ''' </summary>
  ''' <param name="index"></param>
  ''' <param name="oldValue"></param>
  ''' <param name="newValue"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Overrides Sub OnSetComplete(ByVal index As Integer, ByVal oldValue As Object, ByVal newValue As Object)
    MyBase.OnSetComplete(index, oldValue, newValue)
    Dim keyIndex As Index
    For Each keyIndex In _indexes
      keyIndex.Remove(CallByName(oldValue, keyIndex.IndexName, CallType.Get))
    Next
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a value in all indexes for the specified object
  ''' </summary>
  ''' <param name="value">The object being added to the collection</param>
  ''' <param name="index">The index of the object in the collection</param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Sub InsertIndexValue(ByVal value As Object, ByVal index As Integer)
    Dim keyIndex As Index
    For Each keyIndex In _indexes
      keyIndex.Add(CallByName(value, keyIndex.IndexName, CallType.Get), index)
    Next
    RaiseEvent ValueInserted(Me, New InsertRemoveEventArgs(value))
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Removes the references to the passed object from all indexes
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Sub RemoveIndexValue(ByVal value As Object)
    Dim keyIndex As Index
    For Each keyIndex In _indexes
      keyIndex.Remove(CallByName(value, keyIndex.IndexName, CallType.Get))
    Next
    RaiseEvent ValueRemoved(Me, New InsertRemoveEventArgs(value))
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the position of the object within the collection
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [21/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function IndexOf(ByVal value As Object) As Integer
    Return MyBase.List.IndexOf(value)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the item at the specified position within the collection
  ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Overridable Overloads ReadOnly Property ItemAt(ByVal index As Integer) As Object
    Get
      Return MyBase.List.Item(index)
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the item with the specified key in the primary index
  ''' </summary>
  ''' <param name="key"></param>
  ''' <value></value>
  ''' <remarks>
  '''   The primary index is the one specified in the constructor for this class.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Default Public Overridable Overloads ReadOnly Property Item(ByVal key As Object) As Object
    Get
      Try
        Return Me.Item(_primaryIndex, key)
      Catch
        Return Nothing
      End Try
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the item with the specified key in the specified index
  ''' </summary>
  ''' <param name="indexName"></param>
  ''' <param name="key"></param>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Default Public Overridable Overloads ReadOnly Property Item(ByVal indexName As String, ByVal key As Object) As Object
        Get
            Dim keyIndex As Index = Me.Indexes.Item(indexName)
            Dim RetObject As Object = Nothing
            If Not keyIndex Is Nothing Then
                Dim index As Object = keyIndex.Item(key)
                If TypeOf index Is Integer Then
                    RetObject = MyBase.List.Item(CInt(index))
                End If
            End If
            Return RetObject
        End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Rebuilds the values in the specified index
  ''' </summary>
  ''' <param name="indexName"></param>
  ''' <remarks>
  '''   Useful if index added after collection has been populated, or object added
  '''   using AddUnindexed.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Protected Sub BaseReIndex(ByVal indexName As String)
    Dim keyIndex As Index = Me.Indexes.Item(indexName)
    If Not keyIndex Is Nothing Then
      keyIndex.Clear()
      Dim pos As Integer
      For pos = 0 To Me.Count - 1
        keyIndex.Add(CallByName(Me.Item(pos), indexName, CallType.Get), pos)
      Next
    End If
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns a subset of the collection where the property matches the specified value
  ''' </summary>
  ''' <param name="propertyName"></param>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [04/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Overridable Function Subset(ByVal propertyName As String, ByVal value As Object) As Object()
    Dim items As New ArrayList()
    Dim item As Object
    For Each item In Me
      If CallByName(item, propertyName, CallType.Get).Equals(value) Then
        items.Add(item)
      End If
    Next
    Return items.ToArray()
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns an array of the distinct values for the specified property
  ''' </summary>
  ''' <param name="propertyName"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [04/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Overridable Function DistinctValues(ByVal propertyName As String) As Object()
    Dim items As New Hashtable()
    Dim item As Object
    For Each item In Me
      Dim value As Object = CallByName(item, propertyName, CallType.Get)
      If items(value) Is Nothing Then
        items.Add(value, value)
      End If
    Next
    Dim values(items.Count - 1) As Object
    items.Keys.CopyTo(values, 0)
    Return values
  End Function

  Public Overridable Function Values() As Object()
    Return Me.InnerList.ToArray()
  End Function

#Region "Index"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : IndexedCollection.Index
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   A single index for an IndexedCollection
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class Index
    Inherits Collections.DictionaryBase

    Private _indexName As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a collection index
    ''' </summary>
    ''' <param name="indexName"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub New(ByVal indexName As String)
      _indexName = indexName
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of this index
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   This is also the name of the property that will provide the key values for this index
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IndexName() As String
      Get
        Return _indexName
      End Get
    End Property

    Friend Sub Add(ByVal key As Object, ByVal position As Integer)
      MyBase.Dictionary.Add(key, position)
    End Sub

    Default Friend ReadOnly Property Item(ByVal key As Object) As Integer
      Get
        Dim position As Object = MyBase.Dictionary.Item(key)
        If position Is Nothing Then
          Return -1
        Else
          Return CInt(position)
        End If
      End Get
    End Property

    Friend Sub Remove(ByVal key As Object)
      MyBase.Dictionary.Remove(key)
    End Sub

#Region "Index Collection"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : IndexedCollection.Index.Collection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Holds a collection of indexes for an IndexedCollection
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class Collection
      Inherits DictionaryBase

      Friend Sub New()
        MyBase.New()
      End Sub

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Adds a new index to the collection
      ''' </summary>
      ''' <param name="indexName"></param>
      ''' <remarks>
      '''   The indexName must match a property defined on all objects within the
      '''   IndexedCollection.  The index is NOT automatically populated with values 
      '''   for objects already in the collection.  Call ReIndex on the IndexedCollection, 
      '''   passing in this index name to achieve this.
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Sub Add(ByVal indexName As String)
        If MyBase.Dictionary.Item(indexName) Is Nothing Then
          MyBase.Dictionary.Add(indexName, New Index(indexName))
        End If
      End Sub

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Removes the specified index from the collection
      ''' </summary>
      ''' <param name="indexName"></param>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Sub Remove(ByVal indexName As String)
        MyBase.Dictionary.Remove(indexName)
      End Sub

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Gets the CollectionIndex with the specified index name
      ''' </summary>
      ''' <param name="indexName"></param>
      ''' <value></value>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Default Public ReadOnly Property Item(ByVal indexName As String) As Index
        Get
          Try
            Return CType(MyBase.Dictionary(indexName), Index)
          Catch
            Return Nothing
          End Try
        End Get
      End Property

      ''' -----------------------------------------------------------------------------
      ''' <summary>
      '''   Enumerates through all indexes in the collection
      ''' </summary>
      ''' <returns></returns>
      ''' <remarks>
      ''' </remarks>
      ''' <revisionHistory>
      ''' <revision><Author>[BOUGHTON]</Author><date> [27/11/2006]</date><Action></Action></revision>
      ''' </revisionHistory>
      ''' -----------------------------------------------------------------------------
      Public Shadows Function GetEnumerator() As IEnumerator
        Return MyBase.Dictionary.Values.GetEnumerator
      End Function

    End Class

    'Public Class GroupCollection
    '  Inherits IndexCollection

    'End Class
#End Region

  End Class

#End Region

#Region "Group"
  'Public Class CollGroup

  'End Class
#End Region

#Region "Event Args"
  Protected Class InsertRemoveEventArgs
    Inherits EventArgs

    Private _value As Object

    Public Sub New(ByVal value As Object)
      _value = value
    End Sub

    Public ReadOnly Property Value() As Object
      Get
        Return _value
      End Get
    End Property

  End Class
#End Region

End Class

#End Region

#End Region


Public MustInherit Class DataBasedObject

    Public Event DataChanged As EventHandler

    Protected m_data As DataRow

    Protected Sub New()
    End Sub

    Protected Sub New(ByVal myData As DataRow)
        Me.SetData(myData)
    End Sub

    Protected Overridable Sub SetData(ByVal myData As DataRow)
        m_data = myData
    End Sub

    Protected Function GetField(ByVal columnName As String) As Object
        Try
            Return m_data(columnName)
        Catch ex As Exception

            Return DBNull.Value
        End Try
    End Function

    Protected Function GetString(ByVal columnName As String) As String
        Dim value As Object = Me.GetField(columnName)
        If value Is Nothing Then
            Return Nothing
        ElseIf TypeOf value Is String Then
            Return DirectCast(value, String)
        Else
            Return value.ToString
        End If
    End Function

    Protected Overridable Overloads Sub SetField(ByVal columnName As String, ByVal value As Object)
        Me.InnerSetField(columnName, value)
    End Sub

    Protected Sub InnerSetField(ByVal columnName As String, ByVal value As Object)
        m_data.BeginEdit()
        ' Treat empty Nullable types as Null
        If Object.Equals(value, Nothing) Then
            value = DBNull.Value
        End If
        m_data(columnName) = value
        m_data.EndEdit()
        RaiseEvent DataChanged(Me, EventArgs.Empty)
    End Sub

    Protected Overridable Overloads Sub SetField(ByVal columnName As String, ByVal value As Boolean)
        Me.InnerSetField(columnName, ToBoolean(value))
    End Sub

    Protected Overridable Overloads Sub SetField(ByVal columnName As String, ByVal value As DateTime)
        If value = DateTime.MinValue Then
            Me.InnerSetField(columnName, DBNull.Value)
        Else
            Me.InnerSetField(columnName, value)
        End If
    End Sub

    Public ReadOnly Property HasChanges() As Boolean
        Get
            Return m_data.RowState <> DataRowState.Unchanged
        End Get
    End Property

End Class
