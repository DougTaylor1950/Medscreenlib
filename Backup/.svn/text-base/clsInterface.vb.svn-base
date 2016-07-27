Option Strict On

Imports System.Data.OleDb

Public MustInherit Class InterfaceHandler
  Implements IDisposable

  Public Const Request_StopInterface As String = "EXIT"

  Public Const Status_Request As String = "R"
  Public Const Status_Locked As String = "L"
  Public Const Status_Complete As String = "C"
  Public Const Status_Failed As String = "F"
  Public Const Status_Deleted As String = "X"
  Public Const Status_PassToVGL As String = "V"

  Private _adapter As OleDbDataAdapter
  Private _interfaceTable As DataTable
  Private _finished As Boolean
  Private _requestType As Type
  Private _getRequestsCommand As OleDbCommand
  Private _refreshRequestCommand As OleDbCommand
  Private _disposed As Boolean
	Private _tableDef As MedStructure.StructTable
	Private _idSequence As String

  Shared Sub New()

  End Sub

  Sub New(ByVal tableName As String, ByVal requestType As Type, ByVal idSequence As String)
    _adapter = New OleDbDataAdapter()
    _tableDef = MedStructure.GetTable(tableName)
    ' Get all requests to process
    _getRequestsCommand = _tableDef.GetSelectCommand(_tableDef.Fields, "WHERE Status = ?")
    _getRequestsCommand.Parameters(0).Value = Status_Request
    ' Get single request details
		_refreshRequestCommand = _tableDef.GetSelectCommand
		_adapter.InsertCommand = _tableDef.GetInsertCommand
		_adapter.UpdateCommand = _tableDef.GetUpdateCommand
		_adapter.DeleteCommand = _tableDef.GetDeleteCommand
		_interfaceTable = _tableDef.CreateDataTable
		_requestType = requestType
		If _requestType Is Nothing Then
			Throw New ArgumentNullException("requestType")
		ElseIf Not _requestType.IsSubclassOf(GetType(InterfaceRequest)) Then
			Throw New InvalidCastException(String.Format("Type {0} does not inerit from InterfaceRequest", _requestType.Name))
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
						Dim request As InterfaceRequest = CreateRequest(requestData)
						If Not request Is Nothing Then
							Me.ExitRequested = request.IsExitRequest
						End If
						If ExitRequested Then
							Exit Function
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

	Private Sub RefreshData(ByVal request As InterfaceRequest)
		'MedConnection.Open(_adapter.SelectCommand.Connection)
		_adapter.SelectCommand = _refreshRequestCommand
		_adapter.SelectCommand.Parameters(0).Value = request.ID
		_adapter.Fill(_interfaceTable)
		'MedConnection.Close(_adapter.SelectCommand.Connection)
	End Sub

	Private Function CreateRequest(ByVal row As DataRow) As InterfaceRequest
		Dim flags As Reflection.BindingFlags = Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Public
		' Loop through all samples rows, creating an object for each one
		Dim ctorArgs As Object() = New Object() {Me, row}
		' Create the sample object, either of type MedSample, or something inheriting from it.
		Try
			Return DirectCast(Activator.CreateInstance(_requestType, flags, Nothing, ctorArgs, Nothing), InterfaceRequest)
		Catch ex As Exception
			Medscreen.LogError(ex)
		End Try
	End Function

	Protected Function NewRequest(ByVal requestCode As String) As InterfaceRequest
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
		If requestData.HasVersion(DataRowVersion.Proposed) Or _
		 Not requestData.HasVersion(DataRowVersion.Original) Then
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
		Catch ex As Exception		'System.Data.OleDb.OleDbException
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


	Public MustInherit Class InterfaceRequest
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

		Public Event ResponseReceived(ByVal sender As InterfaceRequest, ByVal e As EventArgs)

		Protected _parent As InterfaceHandler
		Protected _requestData As DataRow, _type As Direction
		Private WithEvents _timer As Timers.Timer

		Protected Sub New(ByVal parent As InterfaceHandler, ByVal requestData As DataRow)
			_parent = parent
			_requestData = requestData
			If requestData.RowState = DataRowState.Detached Then
				_type = Direction.Post
			Else
				_type = Direction.Process
				Me.Process()
			End If
		End Sub

		Public ReadOnly Property RequestDirection() As Direction
			Get
				Return _type
			End Get
		End Property

		Overridable ReadOnly Property RequestCode() As String
			Get
				Return _requestData!RequestCode.ToString
			End Get
		End Property

		Property Status() As String
			Get
				Return _requestData!Status.ToString
			End Get
			Set(ByVal Value As String)
				_requestData!status = Value
			End Set
		End Property

		Overridable ReadOnly Property IsExitRequest() As Boolean
			Get
				Return Me.RequestCode = InterfaceHandler.Request_StopInterface
			End Get
		End Property

		MustOverride ReadOnly Property ID() As Integer

		Protected Sub SetField(ByVal name As String, ByVal value As Object)
			Dim field As MedStructure.StructField = _parent._tableDef.Fields(name)
			If Not field Is Nothing Then
				_requestData.Item(name) = field.FormatValue(value)
			End If
		End Sub

		Public Sub Update()
			Try
				If _parent.UpdateRequest(_requestData) Then
					RunTimer()
				End If
			Catch ex As Exception
				Medscreen.LogError(ex)
			End Try
		End Sub

		Private Sub RunTimer()
			_timer = New Timers.Timer()
			_timer.Interval = 500
			_timer.AutoReset = False
			_timer.Start()
		End Sub

		Private Sub Process()
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

		Private Sub timer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles _timer.Elapsed
			_parent.RefreshData(Me)
			If Me.HasResponse Then
				_timer.Dispose()
				_timer = Nothing
				Try
					RaiseEvent ResponseReceived(Me, New EventArgs())
				Catch
				End Try
				Me.Status = Status_Deleted
				Me.Update()
				Me.Dispose()
			Else
				_timer.Start()
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
		Protected ReadOnly Property HasResponse() As Boolean
			Get
				Return Me.RequestDirection = Direction.Post AndAlso _
				 (Me.Status = Status_Complete Or Me.Status = Status_Failed)
			End Get
		End Property

		Public Sub Dispose() Implements IDisposable.Dispose
			If Not (_parent Is Nothing OrElse _parent.IsDisposed OrElse _parent._interfaceTable Is Nothing) Then
				_parent._interfaceTable.Rows.Remove(_requestData)
			End If
			If Not _timer Is Nothing Then
				_timer.Dispose()
			End If
		End Sub

	End Class

End Class

