Public Class SageConnection

    Private Shared mySageConnection As SqlClient.SqlConnection
    Private Shared _connectionType As Type
    Private Shared _innerState As Reflection.PropertyInfo


#Region "Shared"
#Region "Shared Properties"
    Public Shared Property SageConnection() As SqlClient.SqlConnection
        Get
            If mySageConnection Is Nothing Then
                mySageConnection = New SqlClient.SqlConnection("Persist Security Info=False;User ID=sage200reports;Password=sage200reports;Initial Catalog=Sage200_ConcatenoProcurement;Data Source=sql")
            End If
            Return mySageConnection
        End Get
        Set(ByVal Value As SqlClient.SqlConnection)
            mySageConnection = Value
        End Set
    End Property
#End Region

#Region "Shared Functions"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempt to open connection and return status
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Attempt to open connection and return status
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/04/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function ConnOpen() As Boolean
        Return Open(SageConnection)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find the real state of the connection
    ''' </summary>
    ''' <param name="dbconnection"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/04/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetRealState(ByVal dbconnection As Data.SqlClient.SqlConnection) As ConnectionState
        If _connectionType Is Nothing Then
            _connectionType = GetType(Data.SqlClient.SqlConnection)
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
    ''' Attempt to open connection and return status
    ''' </summary>
    ''' <param name="dbConnection"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/04/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Open(ByVal dbConnection As SqlClient.SqlConnection) As Boolean
        Return Open(dbConnection, 10000)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempt to open connection and return status
    ''' </summary>
    ''' <param name="dbConnection"></param>
    ''' <param name="timout"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/04/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Open(ByVal dbConnection As SqlClient.SqlConnection, ByVal timout As Integer) As Boolean
        If dbConnection Is Nothing Then
            Return False
        Else
            Dim startTime As DateTime = DateTime.Now
            Dim realState As ConnectionState = GetRealState(SageConnection)
            If (realState And ConnectionState.Open) = 0 Then
                Try
                    dbConnection.Open()
                Catch ex As System.Exception
                    Debug.WriteLine(ex.Message)
                End Try
                realState = GetRealState(SageConnection)
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

    Public Shared Function GetSageCompanyName(ByVal SageID As String) As String
        Dim oCmd As New SqlClient.SqlCommand("Select CustomerAccountName from SLCustomerAccount where CustomerAccountNumber = '" & SageID & "'", SageConnection)
        Dim oName As Object = Nothing
        Try
            If ConnOpen() Then
                oName = oCmd.ExecuteScalar
            End If
            If oName Is Nothing Then
                oName = "No Company name found"
            End If
        Catch ex As Exception
            oName = "Error"
            Medscreen.LogError(ex)
        Finally
            SageConnection.Close()
        End Try
        Return oName
    End Function

#End Region
#End Region

End Class
