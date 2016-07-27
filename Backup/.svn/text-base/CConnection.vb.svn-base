'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2006-06-06 17:51:29+01 $
'$Log: CConnection.vb,v $
'Revision 1.0  2006-06-06 17:51:29+01  taylor
'Initial revision
'
'Revision 1.0  2005-09-09 13:20:14+01  taylor
'Checked in after commenting
'
'<>
Option Strict On

'
'''<summary>
''' Class to provide access to a single data connector for use in all modules
''' </summary>
Public Class CConnection

    '''<summary>Which database to use</summary>
    Public Enum useDatabase
        '''<summary>Live database</summary>
        LIVE
        '''<summary>Development database</summary>
        DEV
        '''<summary>Test database</summary>
        TEST
        '''<summary>Test of migrated database</summary>
        MIGRATION
    End Enum

    Private Shared UsedDatabase As useDatabase = useDatabase.LIVE

    Private strConnect As String
    Private Shared intSchemaMajor As Integer = -1
    Private Shared intSchemaMinor As Integer = -1
    Private Shared intSchemaRelease As Integer = -1

    '''<summary>
    ''' Connection in use
    ''' </summary>
    Public Shared ReadOnly Property DbConnection() As OleDb.OleDbConnection
        Get
            Return MedConnection.Connection
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Compare a schema against the current schema value 
    ''' </summary>
    ''' <param name="SchemaVersion">Schema version to compare </param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 		Return values <0 indicate current schema is prior to sSchemeID
    '''     	Return values >0 indicate current schema is past sSchemaID
    '''     	Return value 0 indicates a match.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CompareSchema(ByVal SchemaVersion As String) As Integer
        Return Medscreen.CompareVersions(Medscreen.DataBaseVersion(), SchemaVersion)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main setting on version number
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property SchemaMajor() As Integer
        Get
            If intSchemaMajor = -1 Then
                SetSchemaIds()
            End If
            Return intSchemaMajor
        End Get
        Set(ByVal Value As Integer)
            intSchemaMajor = Value
        End Set
    End Property

    Private Shared Sub SetSchemaIds()
        Dim strReturn As String
        'If CConnection.SchemaMajor = -1 Then
        Dim oCmd As New OleDb.OleDbCommand("Select value from config_header where identity = 'SCHEMA'", CConnection.DbConnection)
        Try
            CConnection.SetConnOpen()
            strReturn = CStr(oCmd.ExecuteScalar)
        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()
        End Try
        Dim intPos As Integer
        intPos = InStr(strReturn, ".")
        If intPos > 0 Then
            intSchemaMajor = CInt(Mid(strReturn, 1, intPos - 1))
            strReturn = Mid(strReturn, intPos + 1)
            intPos = InStr(strReturn, ".")
            If intPos > 0 Then
                intSchemaMinor = CInt(Mid(strReturn, 1, intPos - 1))
                intSchemaRelease = CInt(Mid(strReturn, intPos + 1))
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Second place setting on version number
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property SchemaMinor() As Integer
        Get
            If intSchemaMinor = -1 Then
                SetSchemaIds()
            End If
            Return intSchemaMinor
        End Get
        Set(ByVal Value As Integer)
            intSchemaMinor = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Third place setting on schema version number
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property SchemaRelease() As Integer
        Get
            If intSchemaRelease = -1 Then
                SetSchemaIds()
            End If

            Return intSchemaRelease
        End Get
        Set(ByVal Value As Integer)
            intSchemaRelease = Value
        End Set
    End Property


    '''<summary> The 'Provider' for the connection string
    ''' </summary>
    Public Shared ReadOnly Property DBProvider() As String
        Get
            Return MedConnection.Instance.Provider
        End Get
    End Property

    '''<summary> The database in use
    ''' </summary>
    Public Shared ReadOnly Property DBDatabase() As String
        Get
            Return MedConnection.Instance.DataSource
        End Get
    End Property

    '''<summary> A human readable for the database in use
    ''' </summary>
    Public Shared Function DbDescription() As String
        If DatabaseInUse = useDatabase.DEV Then
            Return "Development"
        ElseIf DatabaseInUse = useDatabase.TEST Then
            Return "Test database"
        Else
            Return "Live Database"
        End If
    End Function

    '''<summary>The user name</summary>
    Public Shared Property DBUserName() As String
        Set(ByVal Value As String)
            MedConnection.Instance.UserName = Value
        End Set
        Get
            Return MedConnection.Instance.UserName
        End Get
    End Property

    '''<summary>The user's password</summary>
    Public Shared WriteOnly Property DBPassword() As String
        Set(ByVal Value As String)
            MedConnection.Instance.Password = Value
        End Set
    End Property

    '''<summary>The database Instance in use
    ''' </summary>
    ''' <param name='dbInUse'>Database in use</param>
    Public Overloads Shared ReadOnly Property DBInstance(ByVal dbInUse As useDatabase) As String
        Get
            Select Case dbInUse
                Case useDatabase.DEV, useDatabase.MIGRATION
                    Return "DEV"
                Case useDatabase.LIVE, useDatabase.TEST
                    Return "LIVE"

            End Select
        End Get
    End Property

    '''<summary>The database Instance in use
    ''' </summary>
    Public Overloads Shared ReadOnly Property DBInstance() As String
        Get
            Return MedConnection.Instance.SID
        End Get
    End Property

    '''<summary>Database Server in use
    ''' </summary>
    Public Overloads Shared ReadOnly Property DBServerName() As String
        Get
            Return MedConnection.Instance.Host
        End Get
    End Property

    '''<summary>Database Server in use
    ''' </summary>
    ''' <param name='dbInUse'>Database in use</param>
    Public Overloads Shared ReadOnly Property DBServerName(ByVal dbInUse As useDatabase) As String
        Get
            Return MedConnection.Connections(dbInUse.ToString).Host
        End Get
    End Property

    '''<summary>Server path in the form "\\ServerName\Instance"
    ''' </summary>
    Public Shared ReadOnly Property DBServerPath() As String
        Get
            Return MedConnection.Instance.ServerPath
        End Get
    End Property

    '''<summary>Server path in the form "\\ServerName\Instance"
    ''' </summary>
    ''' <param name='dbInUse'>Database in use</param>
    Public Shared ReadOnly Property DBServerPath(ByVal dbInUse As useDatabase) As String
        Get
            Return MedConnection.Connections(dbInUse.ToString).ServerPath
        End Get
    End Property


    '''<summary>Database in use
    ''' </summary>
    Public Shared Property DatabaseInUse() As useDatabase
        Get
            Return UsedDatabase
        End Get
        Set(ByVal Value As useDatabase)
            UsedDatabase = Value
            MedConnection.SetInstance(Value.ToString)
        End Set
    End Property

    '''<summary>Open a connection
    ''' </summary>
    ''' <returns>TRUE if Open is succesful</returns>
    Public Shared Function ConnOpen() As Boolean
        Return MedConnection.Open
    End Function

    '''<summary>Open the connection
    ''' </summary>
    Public Shared Sub SetConnOpen()
        ConnOpen()
    End Sub

    '''<summary>Close the connection
    ''' </summary>
    Public Shared Sub SetConnClosed()
        MedConnection.Close()
    End Sub

    Public Shared Function NextID(ByVal Table As String, ByVal IdField As String) As Integer
        Dim ocmd As New OleDb.OleDbCommand()

        'Build query to get sequence number 
        ocmd.CommandText = "Select max(" & IdField & ") from " & Table
        Dim intRet As Integer
        Dim intRes As Integer
        Dim myTrans As OleDb.OleDbTransaction
        Dim strID As String

        Try
            ocmd.Connection = DbConnection      'Get Connection
            SetConnOpen()                       'Open connection

            'Create a transaction ticket to handle these actions they need to be together
            'Take a lock on the datarow to prevent Phantom reads
            myTrans = DbConnection.BeginTransaction(IsolationLevel.ReadCommitted)
            ocmd.Transaction = myTrans
            intRet = CInt(ocmd.ExecuteScalar) + 1                'Get oldId by fastest route possible
        Catch
        Finally
            SetConnClosed()                       'Open connection

        End Try
        Return intRet
    End Function

    '''<summary>Get the next value for the supplied Sequence number
    ''' </summary>
    ''' <param name='Sequence'>Sequence generator to use</param>
    ''' <param name='Length'>Length of string to use, i.e. pad to zeros, default=10</param> 
    ''' <returns>Sequence number left padded with Zeros e.g. 0000012345 </returns>
    Public Shared Function NextSequence(ByVal Sequence As String, _
    Optional ByVal Length As Integer = 10) As String
        Dim ocmd As New OleDb.OleDbCommand()

        'Build query to get sequence number 
        ocmd.CommandText = "Select " & Sequence.Trim & ".nextval from dual"


        Dim intRet As Integer
        Dim intRes As Integer
        Dim myTrans As OleDb.OleDbTransaction
        Dim strID As String

        Try
            ocmd.Connection = DbConnection      'Get Connection
            SetConnOpen()                       'Open connection

            'Create a transaction ticket to handle these actions they need to be together
            'Take a lock on the datarow to prevent Phantom reads
            myTrans = DbConnection.BeginTransaction(IsolationLevel.ReadCommitted)
            ocmd.Transaction = myTrans
            intRet = CInt(ocmd.ExecuteScalar)                 'Get oldId by fastest route possible

            '                If intRes = 1 Then
            'I've update a row so Okay
            strID = CStr(intRet)
            strID = strID.PadLeft(Length)
            strID = strID.Replace(" ", "0")
            '       End If

            Return strID
        Catch ex As OleDb.OleDbException
            MedscreenLib.Medscreen.LogError(ex)
            Return Nothing
        Catch ex As Exception

            MedscreenLib.Medscreen.LogError(ex)
            Return Nothing
        Finally
            SetConnClosed()                       'Open connection

        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a comma separated list from server using defined package and library
    ''' </summary>
    ''' <param name="PackageRoutine"></param>
    ''' <param name="_Type"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function PackageStringList(ByVal PackageRoutine As String, ByVal _Type As String) As String
        Dim strRet As String = ""
        Dim ReturnString As Object

        Dim ocmd As New OleDb.OleDbCommand()
        ocmd.Connection = CConnection.DbConnection
        Try

            CConnection.SetConnOpen()
            If _Type.Trim.Length > 0 Then 'We have a single parameter
                ocmd.CommandText = "Select " & PackageRoutine.Trim & "(?) from dual"
                ocmd.Parameters.Add(CConnection.StringParameter("type", _Type, 400))
            Else 'Parameter less call
                ocmd.CommandText = "Select " & PackageRoutine.Trim & " from dual"
            End If
            Debug.WriteLine(ocmd.CommandText & "-" & Now.ToLongTimeString)
            ReturnString = ocmd.ExecuteScalar
            If ReturnString Is Nothing OrElse ReturnString Is DBNull.Value Then
                ReturnString = ""
            End If
        Catch ex As Exception
            Debug.WriteLine("Error in - : " & ocmd.CommandText & " - " & _Type)
            Medscreen.LogError(ex, , ocmd.CommandText)
        Finally
            CConnection.SetConnClosed()
        End Try
        Return CStr(ReturnString)
    End Function

    Public Overloads Shared Function PackageStringList(ByVal PackageRoutine As String, ByVal _Type As Integer) As String
        Dim strRet As String = ""
        Dim ReturnString As Object

        Dim ocmd As New OleDb.OleDbCommand()
        ocmd.Connection = CConnection.DbConnection
        Try

            CConnection.SetConnOpen()
            ocmd.CommandText = "Select " & PackageRoutine.Trim & "(?) from dual"
            ocmd.Parameters.Add(CConnection.IntegerParameter("type", _Type))
            Debug.WriteLine(ocmd.CommandText & "-" & Now.ToLongTimeString)
            ReturnString = ocmd.ExecuteScalar
            If ReturnString Is Nothing OrElse ReturnString Is DBNull.Value Then
                ReturnString = ""
            End If
        Catch ex As Exception
            Debug.WriteLine("Error in - : " & ocmd.CommandText)
            Medscreen.LogError(ex, , ocmd.CommandText)
        Finally
            CConnection.SetConnClosed()
        End Try
        Return CStr(ReturnString)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a Comma separated list by calling a PL/SQl Package, package.Function is supplied together with a collection of parameters.
    ''' Can also be used to return a string from a Package.
    ''' </summary>
    ''' <param name="PackageRoutine"></param>
    ''' <param name="Parameters"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [20/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function PackageStringList(ByVal PackageRoutine As String, ByVal Parameters As Collection) As String
        Dim strRet As String = ""
        Dim ReturnString As Object

        Dim ocmd As New OleDb.OleDbCommand()
        ocmd.Connection = CConnection.DbConnection
        Dim strQuery As String = "Select " & PackageRoutine.Trim & "("
        Dim strPa As String = ""
        Dim strPaValue As String            'Being built for error handler as routine is generic we want to know what it was doing when it goes belly up
        Dim i As Integer
        'Go through each parameter adding it to the query  as a ?
        For i = 1 To Parameters.Count
            If strPa.Length > 0 Then strPa += ","
            strPaValue += CStr(Parameters.Item(i)) & ","
            strPa += "?"
        Next
        strQuery += strPa & ") from dual"

        Try
            CConnection.SetConnOpen()
            ocmd.CommandText = strQuery
            'Generate a paremeter in the Oracle command for each parameter passed, hopefully they will be less than 20 characters
            For i = 1 To Parameters.Count
                Try     'Try converting to date if it exceptions use string
                    Dim DateString As String = CStr(Parameters.Item(i))

                    Dim tmpDate As Date = Date.Parse(DateString.Trim)
                    ocmd.Parameters.Add(CConnection.DateParameter("type" & CStr(i), tmpDate, True))
                Catch
                    ocmd.Parameters.Add(CConnection.StringParameter("type" & CStr(i), CStr(Parameters.Item(i)), 40))

                End Try
            Next
            Debug.WriteLine(strQuery & "- " & Now.ToLongTimeString)
            ReturnString = ocmd.ExecuteScalar   'Return a single value 

            If ReturnString Is Nothing OrElse ReturnString Is DBNull.Value Then
                ReturnString = ""
            End If

        Catch ex As Exception
            Medscreen.LogError(ex, , "Executing " & ocmd.CommandText)
        Finally
            CConnection.SetConnClosed()
        End Try
        Return CStr(ReturnString)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create a string parameter 
    ''' </summary>
    ''' <param name="Name">name of parameter</param>
    ''' <param name="Value">value of string </param>
    ''' <param name="length">size of string</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StringParameter(ByVal Name As String, ByVal Value As String, ByVal length As Integer) As OleDb.OleDbParameter
        Dim retParam As OleDb.OleDbParameter = New OleDb.OleDbParameter(Name, Value)
        If Value.Trim.Length = 0 Then retParam.Value = System.DBNull.Value
        retParam.DbType = DbType.String
        retParam.Size = length
        retParam.Direction = ParameterDirection.Input
        Return retParam
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a parameter to represent a boolean value 
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function BooleanParameter(ByVal Name As String, ByVal Value As Boolean) As OleDb.OleDbParameter
        If Value Then
            Return StringParameter(Name, "T", 1)
        Else
            Return StringParameter(Name, "F", 1)
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create date parameter
    ''' </summary>
    ''' <param name="Name">name of parameter</param>
    ''' <param name="Value">value of parameter</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function DateParameter(ByVal Name As String, ByVal Value As Date, Optional ByVal blnDateTime As Boolean = True) As OleDb.OleDbParameter
        Dim retParam As OleDb.OleDbParameter = New OleDb.OleDbParameter(Name, Value)
        If blnDateTime Then
            retParam.DbType = DbType.DateTime
            If Value <= DateField.ZeroDate Then
                retParam.Value = DBNull.Value
            End If
        Else
            retParam.DbType = DbType.String
            retParam.Size = 12
            If Value <= DateField.ZeroDate Then
                retParam.Value = DBNull.Value
            Else
                retParam.Value = Value.ToString("dd-MMM-yy")
            End If
        End If

        Return retParam
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create integer parameter
    ''' </summary>
    ''' <param name="Name">name of parameter</param>
    ''' <param name="Value">value of parameter</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function IntegerParameter(ByVal Name As String, ByVal Value As Integer) As OleDb.OleDbParameter
        Dim retParam As OleDb.OleDbParameter = New OleDb.OleDbParameter(Name, Value)
        retParam.DbType = DbType.Int32
        Return retParam
    End Function

End Class
