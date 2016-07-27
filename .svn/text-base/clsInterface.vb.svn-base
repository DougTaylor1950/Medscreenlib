Option Strict On

Imports System.Data.OleDb

Public MustInherit Class InterfaceHandler
    Implements IDisposable

#Region "Shared"
    Public Shared Event RequestExceptionThrown(ByVal sender As InterfaceRequestBase, ByVal e As InterfaceExceptionEventArgs)

    Public Const Request_StopInterface As String = "EXIT"

    Public Const Status_Request As String = "R"
    Public Const Status_Locked As String = "L"
    Public Const Status_Complete As String = "C"
    Public Const Status_Failed As String = "F"
    Public Const Status_Deleted As String = "X"
    Public Const Status_PassToVGL As String = "V"

    Protected Shared Sub RaiseExceptionEvent(ByVal sender As InterfaceRequestBase, ByVal e As InterfaceExceptionEventArgs)
        RaiseEvent RequestExceptionThrown(sender, e)
    End Sub
#End Region

#Region "Instance"
    Private _adapter As OleDbDataAdapter
    Private _interfaceTable As DataTable
    Private _finished As Boolean
    Private _requestType As Type
    Private _getRequestsCommand As OleDbCommand
    Private _refreshRequestCommand As OleDbCommand
    Private _disposed As Boolean
    Private _tableDef As MedStructure.StructTable
    Private _idSequence As String

    Sub New(ByVal tableName As String, ByVal requestType As Type, ByVal idSequence As String)
        ' TODO: Having Adapter shared causes problems creating requests in (very) quick succession
        _adapter = New OleDbDataAdapter()
        _tableDef = MedStructure.GetTable(tableName)
        Dim connection As OleDbConnection = MedConnection.CloneConnection(MedConnection.Instance)
        ' Get all requests to process
        _getRequestsCommand = _tableDef.GetSelectCommand(_tableDef.Fields, "WHERE Status = ?")
        _getRequestsCommand.Parameters(0).Value = Status_Request
        _getRequestsCommand.Connection = connection
        ' Get single request details
        _refreshRequestCommand = _tableDef.GetSelectCommand
        _refreshRequestCommand.Connection = connection
        ' Configure other commands on data adapter
        _adapter.InsertCommand = _tableDef.GetInsertCommand
        _adapter.InsertCommand.Connection = connection
        _adapter.UpdateCommand = _tableDef.GetUpdateCommand
        _adapter.UpdateCommand.Connection = connection
        _adapter.DeleteCommand = _tableDef.GetDeleteCommand
        _adapter.DeleteCommand.Connection = connection
        _interfaceTable = _tableDef.CreateDataTable
        _requestType = requestType
        If _requestType Is Nothing Then
            Throw New ArgumentNullException("requestType")
        ElseIf Not _requestType.IsSubclassOf(GetType(InterfaceRequestBase)) Then
            Throw New InvalidCastException(String.Format("Type {0} does not inerit from InterfaceRequestBase", _requestType.Name))
        End If
        If idSequence = "" Then
            _idSequence = MedStructure.Sequences.GetTableSequence(tableName)
        Else
            _idSequence = idSequence
        End If
    End Sub


    Protected Function GetNextRowID() As Integer
        Dim getID As New OleDbCommand("SELECT " & _idSequence & ".NextVal FROM Dual", MedConnection.Connection)
        Try
            If MedConnection.Open Then
                Return CType(getID.ExecuteScalar, Integer)
            End If
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            MedConnection.Close()
        End Try
    End Function


    Public Function ProcessRequests() As Boolean
        Try
            _interfaceTable.Clear()
            If RecordBackgroundLog() Then
                _adapter.SelectCommand = _getRequestsCommand
                MedConnection.Open(_adapter.SelectCommand.Connection)
                _adapter.Fill(_interfaceTable)
                MedConnection.Close(_adapter.SelectCommand.Connection)
                Dim requestData As DataRow
                If _interfaceTable.Rows.Count > 0 Then
                    For Each requestData In _interfaceTable.Rows
                        Dim request As InterfaceRequestBase = CreateRequest(requestData)
                        If Not request Is Nothing Then
                            Me.ExitRequested = request.IsExitRequest
                        End If
                        If Me.ExitRequested Then
                            Exit Function
                        Else
                            request.Process()
                        End If
                    Next
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try
    End Function

    Private Sub RefreshData(ByVal request As InterfaceRequestBase)
        'MedConnection.Open(_adapter.SelectCommand.Connection)
        _adapter.SelectCommand = _refreshRequestCommand
        _adapter.SelectCommand.Parameters(0).Value = request.ID
        _adapter.Fill(_interfaceTable)
        'MedConnection.Close(_adapter.SelectCommand.Connection)
    End Sub

    Private Function CreateRequest(ByVal row As DataRow) As InterfaceRequestBase
        Dim flags As Reflection.BindingFlags = Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Public
        ' Loop through all samples rows, creating an object for each one
        Dim ctorArgs As Object() = New Object() {Me, row}
        ' Create the request object of the required type
        Try
            Return DirectCast(Activator.CreateInstance(_requestType, flags, Nothing, ctorArgs, Nothing), InterfaceRequestBase)
        Catch ex As Exception
            Medscreen.LogError(ex)
            Return Nothing
        End Try
    End Function

    Protected Function NewRequest(ByVal requestCode As String) As InterfaceRequestBase
        Dim row As DataRow = _interfaceTable.NewRow
        row!RequestCode = requestCode
        row!Status = Status_Request
        Return Me.CreateRequest(row)
    End Function

    Protected Function UpdateRequest(ByVal requestData As DataRow) As Boolean
        If requestData.RowState = DataRowState.Detached Then
            ' Insert command will create value from sequence, but need a number
            ' because primary keys cannot be Null
            requestData(_interfaceTable.PrimaryKey(0).ColumnName) = Me.GetNextRowID
            _interfaceTable.Rows.Add(requestData)
        End If
        If requestData.RowState <> DataRowState.Unchanged Then
            _adapter.Update(New DataRow() {requestData})
            requestData.AcceptChanges()
            Return True
        Else
            Return False
        End If
    End Function

    Private Function RecordBackgroundLog() As Boolean
        Try
            Dim cmd As New OleDbCommand("SELECT MAX(Identity) FROM BackgroundLog WHERE Status = 'R' and ProgramName = ?", MedConnection.Connection)
            cmd.Parameters.Add("Name", OleDbType.VarChar, 20).Value = _interfaceTable.TableName
            If MedConnection.Open() Then
                Try
                    Dim id As Integer = CType(cmd.ExecuteScalar, Integer)
                    cmd.CommandText = "UPDATE BackgroundLog SET LastUpdate = SysDate WHERE Identity = ?"
                    cmd.Parameters.Clear()
                    cmd.Parameters.Add("ID", OleDbType.Integer).Value = id
                Catch
                    cmd.CommandText = "INSERT INTO BackgroundLog (Identity, Status, ProgramName, LastUpdate) SELECT MAX (Identity) + 1, 'R', ?, SysDate FROM BackgroundLog"
                End Try
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Medscreen.LogError(ex)
                End Try
            End If
        Catch ex As Exception  'System.Data.OleDb.OleDbException
            ' Connection unavailable
            Return False
        Finally
            MedConnection.Close()
        End Try
        Return True
    End Function

    ''' --------------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or Sets a value indicating whether interface should stop
    ''' </summary>
    ''' <value>
    ''' </value>
    ''' <remarks>
    '''   This can be set from code to prevent the interface handler from processing more
    '''   requests, or it can be set from the interface table using request code "EXIT".
    ''' </remarks>
    ''' --------------------------------------------------------------------------------
    Public Property ExitRequested() As Boolean
        Get
            Return _finished
        End Get
        Set(ByVal Value As Boolean)
            _finished = Value
        End Set
    End Property

    Public Sub Dispose() Implements IDisposable.Dispose
        _disposed = True
        _adapter.Dispose()
        _interfaceTable.Dispose()
        _refreshRequestCommand.Dispose()
        _getRequestsCommand.Dispose()
        _interfaceTable = Nothing
    End Sub

    Public ReadOnly Property IsDisposed() As Boolean
        Get
            Return _disposed
        End Get
    End Property
#End Region

    Public MustInherit Class InterfaceRequestBase
        Implements IDisposable

        Public Enum Direction
            ''' <summary>
            ''' Request to be posted into an interface table
            ''' </summary>
            Post
            ''' <summary>
            ''' Request from an interface table to be processed
            ''' </summary>
            Process
        End Enum

        Public Event ResponseReceived(ByVal sender As InterfaceRequestBase, ByVal e As EventArgs)
        Public Event RequestExceptionThrown(ByVal sender As InterfaceRequestBase, ByVal e As InterfaceExceptionEventArgs)
        Public Event Disposing As EventHandler

        Protected m_parent As InterfaceHandler
        Protected m_requestData As DataRow, m_type As Direction
        Private WithEvents m_timer As Timers.Timer

        Protected Sub New(ByVal parent As InterfaceHandler, ByVal requestData As DataRow)
            m_parent = parent
            m_requestData = requestData
            If requestData.RowState = DataRowState.Detached Then
                m_type = Direction.Post
            Else
                m_type = Direction.Process
                Me.Process()
            End If
        End Sub

        Public ReadOnly Property RequestDirection() As Direction
            Get
                Return m_type
            End Get
        End Property

        Overridable ReadOnly Property RequestCode() As String
            Get
                Return m_requestData!RequestCode.ToString
            End Get
        End Property

        Property Status() As String
            Get
                Return m_requestData!Status.ToString
            End Get
            Set(ByVal Value As String)
                m_requestData!status = Value
            End Set
        End Property

        Overridable ReadOnly Property IsExitRequest() As Boolean
            Get
                Return Me.RequestCode = InterfaceHandler.Request_StopInterface
            End Get
        End Property

        MustOverride ReadOnly Property ID() As Integer

        Public ReadOnly Property InterfaceTable() As String
            Get
                Return m_requestData.Table.TableName
            End Get
        End Property

        Protected Sub SetField(ByVal name As String, ByVal value As Object)
            Dim field As MedStructure.StructField = m_parent._tableDef.Fields(name)
            If Not field Is Nothing Then
                m_requestData.Item(name) = field.FormatValue(value)
            End If
        End Sub

        Public Function Update() As Boolean
            Dim sucess As Boolean
            sucess = m_parent.UpdateRequest(m_requestData)
            If sucess AndAlso Me.Status = Status_Request Then
                RunTimer()
            End If
            ' TODO: Could throw interface exception if update fails
        End Function

        Private Sub RunTimer()
            m_timer = New Timers.Timer()
            m_timer.Interval = 500
            m_timer.AutoReset = False
            m_timer.Start()
        End Sub

        Friend Sub Process()
            If Me.RequestDirection = Direction.Process Then
                Me.Status = Status_Locked
                Me.Update()
                If Not Me.IsExitRequest Then
                    Me.OnProcess()
                End If
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Override this sub to process interface requests
        ''' </summary>
        ''' <remarks>
        '''   If the interface is one-way only, there is no need to override this sub.
        '''   Any code to handle interface requests should go here.
        ''' </remarks>
        ''' <history>
        ''' 	[BOUGHTON]	19/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Protected Overridable Sub OnProcess()
        End Sub

        Private Sub timer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles m_timer.Elapsed
            m_parent.RefreshData(Me)
            If Me.HasResponse Then
                m_timer.Dispose()
                m_timer = Nothing
                Try
                    RaiseEvent ResponseReceived(Me, New EventArgs())
                Catch ex As Exception
                    ' Pass on the exception by raising an event
                    ' We can't re-throw the exception from the timer Elapsed Event as there will be nothing to catch it.
                    Dim eArgs As New InterfaceExceptionEventArgs(Me, ex)
                    ' Log error for tracability
                    Medscreen.LogError(ex)
                    ' Raise exception events both from request itself and handler
                    RaiseEvent RequestExceptionThrown(Me, eArgs)
                    InterfaceHandler.RaiseExceptionEvent(Me, eArgs)
                Finally
                    ' Set request status to deleted (X) and dispose
                    'Try
                    '  Me.Status = Status_Deleted
                    '  Me.Update()
                    Me.Dispose()
                    'Catch ex As Exception
                    '  Medscreen.LogError(ex)
                    'End Try
                End Try
            Else
                m_timer.Start()
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Indicates whether a posted request has been responded to
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        '''   Only applicable to Post Requests, not to Process Requests
        ''' </remarks>
        ''' <history>
        ''' 	[BOUGHTON]	19/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property HasResponse() As Boolean
            Get
                Return Me.RequestDirection = Direction.Post AndAlso _
                 (Me.IsComplete Or Me.Failed)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Indicates whether request was processed sucessfully
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	14/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property IsComplete() As Boolean
            Get
                Return Me.Status = Status_Complete
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Indicates whether request failed
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	14/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Failed() As Boolean
            Get
                Return Me.Status = Status_Failed
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not (m_parent Is Nothing OrElse m_parent.IsDisposed OrElse m_parent._interfaceTable Is Nothing) Then
                m_parent._interfaceTable.Rows.Remove(m_requestData)
            End If
            If Not m_timer Is Nothing Then
                m_timer.Dispose()
                m_timer = Nothing
            End If
            RaiseEvent Disposing(Me, Nothing)
        End Sub

    End Class

    <CLSCompliant(True)> _
  Public Class InterfaceExceptionEventArgs
        Inherits EventArgs

        Private _ex As Exception, _request As InterfaceRequestBase

        Public Sub New(ByVal request As InterfaceRequestBase, ByVal ex As Exception)
            _ex = New InterfaceException(request, ex)
            _request = request
        End Sub

        Public ReadOnly Property Exception() As Exception
            Get
                Return _ex
            End Get
        End Property

        Public ReadOnly Property Request() As InterfaceRequestBase
            Get
                Return _request
            End Get
        End Property

    End Class

    <CLSCompliant(True)> _
  Public Class InterfaceException
        Inherits Exception

        Protected Friend Sub New(ByVal request As InterfaceRequestBase, ByVal innerException As Exception)
            MyBase.New(String.Format("An error has occurred in Sample Manager processing an Interface Request.{0}" & _
                                     "The error occured in table: {1}, Row ID: {2}.{3}Please inform I.T. - this message has been logged.{4}", _
                                     ControlChars.NewLine, request.InterfaceTable, request.ID, ControlChars.NewLine, ControlChars.NewLine), _
                                     innerException)
        End Sub
    End Class

End Class

