'Namespace Comments
<CLSCompliant(True)> _
Public Class Comment

#Region "Shared"

  <CLSCompliant(True)> _
Public Class CommentType
    Public Const Accessioning As String = "ACCN"
    Public Const DonorEntry As String = "EDD"
    Public Const Information As String = "INFO"
    Public Const Collection As String = "COLL"
    Public Const OnHold As String = "HOLD"

    Private Sub New()
    End Sub

  End Class

    ''' <summary>
    ''' Types of comment that can exist 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum CommentOn
        ''' <summary>
        ''' Comment on a sample
        ''' </summary>
        ''' <remarks></remarks>
        Sample
        ''' <summary>
        ''' Comment on a customer sample case
        ''' </summary>
        ''' <remarks></remarks>
        [Case]
        ''' <summary>
        ''' Comment on a customer
        ''' </summary>
        ''' <remarks></remarks>
        Customer
        ''' <summary>
        ''' Comment on a collection 
        ''' </summary>
        ''' <remarks></remarks>
        Collection
        ''' <summary>
        ''' Comment on a SMID
        ''' </summary>
        ''' <remarks></remarks>
        SMID
        Officer
        instrument
        Vessel

    End Enum

    ''' <summary>
    ''' Field in the parent table to which this comment is linked (i.e it resolves the foriegn key
    ''' </summary>
    ''' <param name="linkTo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function LinkField(ByVal linkTo As CommentOn) As String
        Select Case linkTo
            Case CommentOn.Sample
                Return "SAMPLE.ID_NUMERIC"
            Case CommentOn.Case
                Return "SAMPLE.BAR_CODE"
            Case CommentOn.Customer
                Return "CUSTOMER.IDENTITY"
            Case CommentOn.SMID
                Return "CUSTOMER.SMID"
            Case CommentOn.Collection
                Return "COLLECTION.IDENTITY"
            Case CommentOn.instrument
                Return "INSTRUMENT.IDENTITY"
            Case CommentOn.Officer
                Return "COLL_OFFICER.IDENTITY"
            Case CommentOn.Vessel
                Return "VESSEL.VESSEL_ID"

            Case Else
                Return linkTo
        End Select
    End Function

    ''' <summary>
    ''' Name by which the link is known
    ''' </summary>
    ''' <param name="linkField"></param>
    ''' <param name="keyID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function LinkName(ByVal linkField As String, ByRef keyID As String) As String
        Select Case linkField
            Case "SAMPLE.ID_NUMERIC"
                Return "Sample"
            Case "SAMPLE.BAR_CODE"
                Return "Case"
            Case "CUSTOMER.IDENTITY"
                keyID = Glossary.Customer.GetSMIDProfile(keyID)
                Return "Customer"
            Case "CUSTOMER.SMID"
                Return "SMID"
            Case "COLLECTION.IDENTITY"
                Return "Collection"
            Case "INSTRUMENT.IDENTITY"
                Return "Instrument"
            Case "COLL_OFFICER.IDENTITY"
                Return "Officer"
            Case "VESSEL.VESSEL_ID"
                Return "Vessel"

            Case Else
                Return linkField
        End Select
    End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Create a comment in the Med_Comments table
  ''' </summary>
  ''' <param name="linkTo"></param>
  ''' <param name="linkKey"></param>
  ''' <param name="comment"></param>
  ''' <param name="userID"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <history>
  ''' 	[Boughton]	20/09/2007	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Shared Sub CreateComment(ByVal linkTo As CommentOn, ByVal linkKey As String, ByVal type As String, ByVal comment As String, ByVal userID As String)
    Dim cmtTable As MedStructure.StructTable = MedStructure.GetTable("Med_Comments")
    Dim fields As MedStructure.StructField.FieldCollection = cmtTable.GetFields("Link_Field,Comment_ID,Comment_Type,CommentText,Creator,Modified_By")
    Dim insertCmd As OleDb.OleDbCommand = cmtTable.GetInsertCommand(fields, False)
    Dim smUser As User = User.GetUser(userID)
    With insertCmd.Parameters
      .Item("Link_Field").Value = LinkField(linkTo).ToUpper
      .Item("Comment_ID").Value = linkKey.ToUpper
      .Item("Comment_Type").Value = type
            .Item("CommentText").Value = comment
            If smUser Is Nothing Then
                .Item("Creator").Value = userID
                .Item("Modified_By").Value = DBNull.Value
            Else
                .Item("Creator").Value = smUser.WindowsIdentity
                .Item("Modified_By").Value = smUser.UserID
            End If
            ' NOTE: Comment_No, Created_On and Deleted are automatically handled by insert trigger
        End With
    Try
      MedConnection.Open(insertCmd.Connection)
      insertCmd.ExecuteNonQuery()
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close(insertCmd.Connection)
    End Try
  End Sub

#End Region

    Protected m_commentData As DataRow

#Region "Data properties"

  Protected Sub New(ByVal commentData As DataRow)
        m_commentData = commentData
  End Sub

    ''' <summary>
    ''' Text of the comment 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public Property Text() As String
    Get
            Return m_commentData!CommentText.ToString
    End Get
    Set(ByVal Value As String)
      If Value <> Me.Text Then
                m_commentData.BeginEdit()
                m_commentData!CommentText = Value
                m_commentData.EndEdit()
      End If
    End Set
  End Property

    ''' <summary>
    ''' Field value which the comment is linked to the parent table by 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public Overridable ReadOnly Property LinksTo() As String
    Get
            Return m_commentData!Link_Field.ToString
    End Get
  End Property
    ''' <summary>
    ''' Field name to which this row is linked 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Link is constructed in the form table.fieldname</remarks>
  Public Overridable ReadOnly Property LinkKey() As String
    Get
            Return m_commentData!Comment_ID.ToString
    End Get
  End Property

    ''' <summary>
    ''' Indexor primary key of the data table
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public ReadOnly Property Index() As Integer
    Get
            Return CInt(m_commentData!Comment_No)
    End Get
  End Property

    ''' <summary>
    ''' Comment type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Within a comment class comments can be of different types, 
    ''' for example within customer comments their might be comments relating to reports sent or to creation of the account </remarks>
  Public ReadOnly Property Type() As String
    Get
            Return m_commentData!Comment_Type.ToString
    End Get
  End Property

    ''' <summary>
    ''' Descriptive explanation of the type of the comment <see cref="Comment.CommentType" />
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public ReadOnly Property TypeText() As String
    Get
      Return Glossary.Phrase.GetPhraseText("MEDCMTTYPE", Me.Type)
    End Get
  End Property

    ''' <summary>
    ''' Date created
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public ReadOnly Property Created() As DateTime
    Get
            Return DirectCast(m_commentData!Created_On, DateTime)
    End Get
  End Property

    ''' <summary>
    ''' The user creating the comment 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public ReadOnly Property CreatedBy() As String
    Get
            Return m_commentData!Creator.ToString
    End Get
  End Property

    ''' <summary>
    ''' User to whom the comment has been passed for action 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This field is really only used for 'action' comments such as the problem resolution comments withing collections</remarks>
  Public ReadOnly Property UserID() As String
    Get
            Return m_commentData!Action_By.ToString
    End Get
  End Property

    ''' <summary>
    ''' This comment has been deleted
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public Property Deleted() As Boolean
    Get
            Return m_commentData!Deleted = C_dbTRUE
    End Get
    Set(ByVal Value As Boolean)
            m_commentData.BeginEdit()
            m_commentData!Deleted = YesNo(Value, C_dbTRUE, C_dbFALSE)
            m_commentData.EndEdit()
    End Set
  End Property

    ''' <summary>
    ''' Person who has last changed the data
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ModifiedBy() As String
        Get
            Return m_commentData!Modified_By.ToString
        End Get
        Set(ByVal Value As String)
            m_commentData!Modified_By = Value
        End Set
    End Property

    ''' <summary>
    ''' Default representation of the comment 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
  Public Overrides Function ToString() As String
    Dim keyID As String = Me.LinkKey
    Dim linkTo As String = LinkName(Me.LinksTo, keyID)
    Return String.Format("{0} for {1} {2}:{3}{4}", Me.TypeText, linkTo, keyID, ControlChars.NewLine, Me.Text)
  End Function

    ''' <summary>
    ''' Method to write data back to the table
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <remarks>As the data may have come from a non updatable view a command is constructed and executed any errors are rethrown
    ''' </remarks>
  Public Sub Update(ByVal userID As String)
        If m_commentData.RowState = DataRowState.Modified Then
            Try
                Me.ModifiedBy = userID
                ' NOTE: The data is retrieved from a view, we need to contruct the update command based on the source
                ' table which can be identified from the LinksTo (Collection comments have their own table, all others
                ' are in the Med_Comments table).  The structure of these tables is not identical!
                Dim tableDef As MedStructure.StructTable
                If Me.LinksTo.ToUpper.StartsWith("COLLECTION") Then
                    tableDef = MedStructure.GetTable("Collection_Comments")
                Else
                    tableDef = MedStructure.GetTable("Med_Comments")
                End If
                ' Create a limited update command: Comment Text always exists...
                Dim updateFields As MedStructure.StructField.FieldCollection = tableDef.GetFields("CommentText")
                ' ...Add any available maintanence fields
                updateFields.AddRange(tableDef.GetFieldsWithUse(MedStructure.FieldUsage.RemoveFlag Or _
                                                                MedStructure.FieldUsage.ModifiedBy Or _
                                                                MedStructure.FieldUsage.ModifiedOn))
                Dim updateCmd As OleDb.OleDbCommand = tableDef.GetUpdateCommand(updateFields)
                ' Populate parameters (list changes depending on table being updated)
                Dim param As OleDb.OleDbParameter
                For Each param In updateCmd.Parameters
                    param.Value = m_commentData.Item(param.SourceColumn)
                Next
                MedConnection.Open()
                Try
                    updateCmd.ExecuteNonQuery()
                    ' Accept changes if update was successful
                    m_commentData.AcceptChanges()
                Catch ex As Exception
                    MedConnection.Close()
                    Throw ex
                End Try
                MedConnection.Close()
                updateCmd.Dispose()
            Catch ex As Exception
                Medscreen.LogError(ex)
                Throw ex
            Finally
            End Try
        End If
  End Sub

#End Region

    ''' <summary>
    ''' Collection of comments 
    ''' </summary>
    ''' <remarks></remarks>
    <CLSCompliant(True)> _
    Public Class Collection
        Inherits CollectionBase
        Implements IDisposable

        Protected m_adapter As OleDb.OleDbDataAdapter
        Protected m_filterCmd As OleDb.OleDbCommand
        Protected m_table As DataTable
        Protected m_subType As Type
        Protected m_disposed As Boolean

        ''' <summary>
        ''' Handles the idispose method
        ''' </summary>
        ''' <remarks>The use of iDispose enables better clean up of unmanaged objects such as data adaptors</remarks>
        Public Sub Dispose() Implements IDisposable.Dispose
            If Not m_disposed Then      'Can't dispose of a disposed object, can't be sure their won't still by references 
                Me.Clear()
                m_table.Dispose()
                m_adapter.Dispose()
                m_filterCmd.Dispose()
                m_disposed = True
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Load all comments for a specified business object
        ''' </summary>
        ''' <param name="linkTo"></param>
        ''' <param name="linkKey"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	20/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal linkTo As CommentOn, ByVal linkKey As String)
            Me.New("Med_Comments")
            Me.SetParams(LinkField(linkTo), linkKey)
            'Me.Load()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Load all comments of a specific type for a business object
        ''' </summary>
        ''' <param name="linkTo"></param>
        ''' <param name="linkKey"></param>
        ''' <param name="commentType"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	20/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal linkTo As CommentOn, ByVal linkKey As String, ByVal commentType As String)
            Me.New(linkTo, linkKey)
            Me.Load(commentType)
        End Sub

        Protected Sub New(ByVal sourceTable As String)
            Me.New(sourceTable, GetType(Comment))
        End Sub

        Protected Sub New(ByVal sourceTable As String, ByVal createType As Type)
            Dim tableDef As MedStructure.StructTable = MedStructure.GetTable(sourceTable)
            m_adapter = New OleDb.OleDbDataAdapter()
            Dim whereFields As New MedStructure.StructField.FieldCollection()
            Dim field As MedStructure.StructField
            ' Add primary keys to where clause, except for numeric index field
            For Each field In tableDef.PrimaryKeys
                If field.SMDataType <> MedStructure.SMFieldDataType.Integer Then
                    whereFields.Add(field)
                End If
            Next
            ' Switch to comments view for creating commands
            tableDef = MedStructure.GetTable("Med_All_Comments")
            m_adapter.SelectCommand = tableDef.GetSelectCommand(tableDef.Fields, whereFields)
            whereFields.Add(tableDef.Fields("COMMENT_TYPE"))
            m_filterCmd = tableDef.GetSelectCommand(tableDef.Fields, whereFields)
            m_table = tableDef.CreateDataTable
            ' Add an alias for Comment ID as CM_Job_Number for updating collection comments
            m_table.Columns.Add("CM_Job_Number", GetType(String), "Comment_ID")
            ' Default to non-removed comments
            m_subType = createType
        End Sub

        ''' <summary>
        ''' Set any parameters of the command object 
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <remarks></remarks>
        Protected Sub SetParams(ByVal ParamArray parameters() As Object)
            'If _adapter.SelectCommand.Parameters.Count <> parameters.Length Then
            '  Throw New ArgumentOutOfRangeException("parameters", "The number of parameters passed must match the numer required for the select command")
            'End If
            Dim index As Integer
            For index = 0 To parameters.Length - 1
                m_adapter.SelectCommand.Parameters(index).Value = parameters(index)
                m_filterCmd.Parameters(index).Value = parameters(index)
            Next
        End Sub

        Protected ReadOnly Property SubType() As Type
            Get
                Return m_subType
            End Get
        End Property

        ''' <summary>
        ''' Load the collection from the database according to the parameters this collection has been initialised to deal with
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Load()
            ' Clear objects, without calling OnClear, thus leaving existing rows in table
            MyBase.InnerList.Clear()
            m_adapter.Fill(m_table)
            ' Create flags for passing to class activator - used to create sample object
            Dim flags As Reflection.BindingFlags = Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Public
            Dim commentRow As DataRow
            For Each commentRow In m_table.Rows
                ' Build an array of arguments to pass on to constructor (containing the data row)
                Dim ctorArgs As Object() = New Object() {commentRow}
                ' Create the sample object, either of type MedSample, or something inheriting from it.
                MyBase.List.Add(Activator.CreateInstance(Me.SubType, flags, Nothing, ctorArgs, Nothing))
            Next
        End Sub

        ''' <summary>
        ''' Load a collection for a particular comment type
        ''' </summary>
        ''' <param name="commentType"></param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal commentType As String)
            Dim oldCmd As OleDb.OleDbCommand = m_adapter.SelectCommand
            Try
                m_adapter.SelectCommand = m_filterCmd
                m_filterCmd.Parameters("COMMENT_TYPE").Value = commentType
                Me.Load()
            Finally
                m_adapter.SelectCommand = oldCmd
            End Try
        End Sub

        ''' <summary>
        ''' Load all values for a particular LinkField value
        ''' </summary>
        ''' <param name="linkTo"></param>
        ''' <param name="linkKey"></param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal linkTo As String, ByVal linkKey As String)
            Me.SetParams(linkTo, linkKey)
            Me.Load()
        End Sub

        ''' <summary>
        ''' Load all values for a particular linkkey value and comment type
        ''' </summary>
        ''' <param name="linkTo"></param>
        ''' <param name="linkKey"></param>
        ''' <param name="commentType"></param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal linkTo As String, ByVal linkKey As String, ByVal commentType As String)
            Me.SetParams(linkTo, linkKey)
            Me.Load(commentType)
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="linkTo"></param>
        ''' <param name="linkKey"></param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal linkTo As CommentOn, ByVal linkKey As String)
            Me.Load(LinkField(linkTo), linkKey)
        End Sub

        Public Sub Load(ByVal linkTo As CommentOn, ByVal linkKey As String, ByVal commentType As String)
            Me.Load(LinkField(linkTo), linkKey, commentType)
        End Sub

        Protected Overrides Sub OnClear()
            MyBase.OnClear()
            m_table.Rows.Clear()
        End Sub

        ''' <summary>
        ''' Returns current active link field 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable ReadOnly Property CurrentLink() As String
            Get
                Return m_adapter.SelectCommand.Parameters("LINK_FIELD").Value.ToString
            End Get
        End Property

        ''' <summary>
        ''' Returns current active link ID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable ReadOnly Property CurrentKey() As String
            Get
                Return m_adapter.SelectCommand.Parameters("COMMENT_ID").Value.ToString
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Gets the item at the specified position,
        '''   NOT the comment with the specified Comment No
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	27/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public ReadOnly Property Item(ByVal index As Integer) As Comment
            Get
                Return DirectCast(MyBase.List.Item(index), Comment)
            End Get
        End Property


        ''' <summary>
        ''' Location of a comment in the collection
        ''' </summary>
        ''' <param name="cmt"></param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IndexOf(ByVal cmt As Comment) As Integer
            Get
                Return MyBase.List.IndexOf(cmt)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Returns an array of comments of the specified type
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Boughton]	07/06/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Items(ByVal commentType As String) As Comment()
            Get
                Dim comments As New ArrayList()
                Dim cmt As Comment
                For Each cmt In Me
                    If cmt.Type = commentType Then
                        comments.Add(cmt)
                    End If
                Next
                Return DirectCast(comments.ToArray(GetType(Comment)), Comment())
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Function to return the Contents of the collection as XML
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function ToXML() As String
            Dim strRet As String = ""
            Return strRet
        End Function

        Public Sub CreateComment(ByVal ActionString As String, ByVal CommentType As String, ByVal Key As String, ByVal _on As CommentOn)
            Comment.CreateComment(_on, Key, CommentType, ActionString, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)


        End Sub

    End Class

End Class
