
''' <summary>
''' A class to handle the BLOB_VALUES table
''' </summary>
''' <remarks></remarks>
''' <revisionHistory></revisionHistory>
''' <author></author>
Public Class BlobManagement
#Region "Constants"
    Private Const cstTable As String = "BLOB_VALUES"
    Private Const cstSequence As String = "SEQ_BLOB_VALUES"
#End Region

#Region "Declarations"
    Private myTableName As String = ""
    Private myBlobId As String = ""
    Private myisBinary As Boolean
    Private myFieldName As String = ""
    Private myBlob As Array
    Private myRowId As String = ""

#End Region

#Region "Instance"
#Region "Properties"

    ''' <summary>
    ''' Table in use
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property TableName() As String
        Get
            Return myTableName
        End Get
        Set(ByVal value As String)
            myTableName = value
        End Set
    End Property




    Public Property BLobId() As String
        Get
            Return myBlobId
        End Get
        Set(ByVal value As String)
            myBlobId = value
        End Set
    End Property


    Public Property IsBinary() As Boolean
        Get
            Return myisBinary
        End Get
        Set(ByVal value As Boolean)
            myisBinary = value
        End Set
    End Property



    Public Property FieldName() As String
        Get
            Return myFieldName
        End Get
        Set(ByVal value As String)
            myFieldName = value
        End Set
    End Property



    Public Property Blob() As Array
        Get
            Return myBlob
        End Get
        Set(ByVal value As Array)
            myBlob = value
        End Set
    End Property

    Public Property BlobString() As String
        Get
            Return Medscreen.ByteArraytoString(Blob)
        End Get
        Set(ByVal value As String)
            myBlob = Medscreen.StringToByteArray(value)
        End Set
    End Property


    Public Property RowID() As String
        Get
            Return myRowId
        End Get
        Set(ByVal value As String)
            myRowId = value
        End Set
    End Property

#End Region

#Region "Procedures"
    Public Sub New(ByVal TableName As String, ByVal FieldName As String)
        MyBase.New()
        myFieldName = FieldName
        myTableName = TableName
    End Sub

    Public Sub New(ByVal blobId As String)
        MyBase.new()
        myBlobId = blobId
    End Sub

    Private Sub GetRowId()
        If BLobId.Trim.Length = 0 Then Exit Sub 'No blob id so can't find saved value
        If Not RowID Is Nothing AndAlso RowID.Trim.Length > 0 Then Exit Sub 'RowId found so don't run
        Dim ocmd As New OleDb.OleDbCommand("Select rowid from blob_values where blob_id = ?", CConnection.DbConnection)
        ocmd.Parameters.Add(CConnection.StringParameter("BlobId", BLobId, 10))
        Dim objRet As Object
        Try
            If CConnection.ConnOpen Then
                objRet = ocmd.ExecuteScalar
                If Not objRet Is System.DBNull.Value Then                           'Possibly none to find will return null
                    RowID = CStr(objRet)
                End If
            End If
        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()
        End Try
    End Sub
#End Region

#Region "Functions"
    ''' <developer></developer>
    ''' <summary>
    ''' Write data to database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Function Save() As Boolean
        If BLobId.Trim.Length = 0 Then      ' Will need to do an insert
            BLobId = CConnection.NextSequence(cstSequence)
            Return doInsert()
        Else
            GetRowId()
            If RowID Is Nothing OrElse RowID.Trim.Length = 0 Then
                Return doInsert()
            Else
                Return doUpdate()
            End If
        End If
    End Function

    Private Function doInsert() As Boolean
        Dim cmdString As String = "Insert into " & cstTable & " values(?,?,?,?,?)"
        Dim oCMd As New OleDb.OleDbCommand(cmdString, CConnection.DbConnection)
        Dim intRet As Integer
        Try
            oCMd.Parameters.Add(CConnection.StringParameter("Blobid", BLobId, 10))
            If IsBinary Then
                oCMd.Parameters.Add(CConnection.StringParameter("IsBinary", "T", 10))
            Else
                oCMd.Parameters.Add(CConnection.StringParameter("IsBinary", "F", 10))
            End If
            oCMd.Parameters.Add(CConnection.StringParameter("TableName", TableName, 40))
            oCMd.Parameters.Add(CConnection.StringParameter("FieldName", FieldName, 40))
            oCMd.Parameters.Add(CConnection.BlobParameter("blob", BlobString))
            If CConnection.ConnOpen Then
                intRet = oCMd.ExecuteNonQuery
            End If

        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()
        End Try
        Return intRet = 1
    End Function

    Private Function doUpdate() As Boolean
        Dim blnReturn As Boolean = False
        Dim cmdstring As String = "Update Blob_values set IS_BINARY =?, TABLE_NAME = ?, FIELD_NAME= ?, BLOB_FIELD = ? where rowid = ?"
        Dim oCmd As New OleDb.OleDbCommand(cmdstring, CConnection.DbConnection)
        Dim intReturn As Integer
        Try
            If IsBinary Then
                oCmd.Parameters.Add(CConnection.StringParameter("IsBinary", "T", 10))
            Else
                oCmd.Parameters.Add(CConnection.StringParameter("IsBinary", "F", 10))
            End If
            oCmd.Parameters.Add(CConnection.StringParameter("TableName", TableName, 40))
            oCmd.Parameters.Add(CConnection.StringParameter("FieldName", FieldName, 40))
            oCmd.Parameters.Add(CConnection.BlobParameter("blob", BlobString))
            oCmd.Parameters.Add(CConnection.StringParameter("Rowid", RowID, 60))
            If CConnection.ConnOpen Then
                intReturn = oCmd.ExecuteNonQuery
            End If
            blnReturn = (intReturn = 1)
        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()
        End Try
        Return blnReturn
    End Function

    Public Function Load() As Boolean
        Dim blnReturn As Boolean = False
        If BLobId.Trim.Length > 0 Then
            Dim strCommand As String = "Select rowID,BLOB_ID,IS_BINARY,TABLE_NAME,FIELD_NAME,BLOB_FIELD from " & cstTable & " where blob_id = ?"
            Dim oCmd As New OleDb.OleDbCommand(strCommand, CConnection.DbConnection)
            oCmd.Parameters.Add(CConnection.StringParameter("BlobID", BLobId, 10))
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Try
                If CConnection.ConnOpen Then
                    oRead = oCmd.ExecuteReader
                    If oRead.Read Then 'get data
                        RowID = oRead.GetValue(0)
                        BLobId = oRead.GetValue(1)
                        IsBinary = (oRead.GetValue(2) = "T")
                        TableName = oRead.GetValue(3)
                        FieldName = oRead.GetValue(4)
                        If Not oRead.IsDBNull(5) Then Blob = oRead.GetValue(5)
                    End If
                End If
                blnReturn = True
            Catch ex As Exception
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
            End Try

        End If
        Return blnReturn
    End Function
#End Region
#End Region

End Class
