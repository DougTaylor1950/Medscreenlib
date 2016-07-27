Option Strict On
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing

'--------------------------------------------------------------------------------
#Region "Template classes"
'--------------------------------------------------------------------------------
Public Class Template

#Region "Shared"

  Public Enum TemplateType
    Samp
    Job
    Batch
    Med
  End Enum

  Public Enum FieldEntryType
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Field value is copied from link table after editing
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    AfterEdit = Asc("A")
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Field value is copied from link table before editing
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    BeforeEdit = Asc("B")
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Field value is entered directly
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Value = Asc("V")
  End Enum

  Protected Enum TableType
    Header
    Fields
    Destination
  End Enum

  Protected Shared m_daTemplateHeader As OleDbDataAdapter
  Protected Shared m_Templates As Hashtable


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the medscreen template with the specified ID
  ''' </summary>
  ''' <param name="templateID"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetTemplate(ByVal templateID As String) As Template
    Return GetTemplate(TemplateType.Med, templateID)
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the template of the specified type and ID
  ''' </summary>
  ''' <param name="type"></param>
  ''' <param name="templateID"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetTemplate(ByVal type As TemplateType, ByVal templateID As String) As Template
    If IsBlank(templateID) Then templateID = "<DEFAULT>"
    If m_Templates.Item(FormatID(type, templateID)) Is Nothing Then
      Dim templaterow As DataRow = GetTemplateHeaderData(type, templateID)
      m_Templates.Add(FormatID(type, templateID), New Template(type, templaterow))
    End If
    Return CType(m_Templates.Item(FormatID(type, templateID)), Template)
  End Function


  Protected Shared Function FormatID(ByVal type As TemplateType, ByVal templateID As String) As String
    Return type.ToString() & "." & templateID
  End Function


  Protected Shared Function GetTemplateHeaderData(ByVal type As TemplateType, ByVal templateID As String) As DataRow
    Dim templateRow As DataRow
    InitDataAdapter(type)
    m_daTemplateHeader.SelectCommand.Parameters(0).Value = templateID
    Try
      MedConnection.Open()
      Dim tmplTable As New DataTable(TableName(TableType.Header, type))
      m_daTemplateHeader.Fill(tmplTable)
      If tmplTable.Rows.Count > 0 Then
        templateRow = tmplTable.Rows(0)
      End If
    Catch ex As System.Exception
      Debug.WriteLine(ex.Message)
    Finally
      MedConnection.Close()
    End Try
    Return templateRow
  End Function


  Shared Sub New()
    ' Initialise all shared objects
    m_Templates = New Hashtable()
    m_daTemplateHeader = New OleDbDataAdapter()
    m_daTemplateHeader.SelectCommand = New OleDbCommand()
    m_daTemplateHeader.SelectCommand.Connection = MedConnection.Connection
  End Sub


  Protected Shared Sub InitDataAdapter(ByVal template As TemplateType)
    With m_daTemplateHeader.SelectCommand
      .Parameters.Clear()
      Select Case template
        ' 1st column MUST be identity / template ID
      Case TemplateType.Samp
          .CommandText = "SELECT Identity, Login_Title, Sample_Status, Edit_Tests, Action_Type, ScreenLayout" & _
                         " FROM " & TableName(TableType.Header, template) & _
                         " WHERE Identity = ?"
          .Parameters.Add("Identity", OleDbType.VarChar, 10)
        Case TemplateType.Med
          .CommandText = "SELECT Template_ID, Template_Type, Description, Title" & _
                         " FROM " & TableName(TableType.Header, template) & _
                         " WHERE Template_ID = ?"
          .Parameters.Add("Template_ID", OleDbType.VarChar, 10)
        Case TemplateType.Batch
          .CommandText = "SELECT Identity, Browse_Description, Prompt_For_Name, Sample_Syntax_Id, " & _
                                "Sample_Template_Id, Sample_Repeat_Count, Job_Syntax_Id, Creation_Title" & _
                         " FROM " & TableName(TableType.Header, template) & _
                         " WHERE Identity = ?"
          .Parameters.Add("Identity", OleDbType.VarChar, 10)
        Case TemplateType.Job
          .CommandText = "SELECT Identity, Description, Analysis, Limits, Batch_Class, Syntax, Edit_Criteria, " & _
                                "Minimum_Samples, Maximum_Samples, Automatic_Select, Edit_List, Random_Control, " & _
                                "Random_Number, Print_On_Creation, Printer, Compression" & _
                         " FROM " & TableName(TableType.Header, template) & _
                         " WHERE Identity = ?"
      End Select
    End With
  End Sub


  Protected Shared Function TableName(ByVal entry As TableType, ByVal template As TemplateType) As String
    Dim table As String
    ' Start off with some rough guesses
    Select Case entry
      Case TableType.Fields
        table = "Template_Fields"
      Case TableType.Header
        table = "Tmpl_Header"
    End Select
    ' Destination is the table that will be written to.
    ' Does not apply to Medscreen templates which can write to multiple tables
    If entry = TableType.Destination Then
      Select Case template
        Case TemplateType.Batch, TemplateType.Job
          table = template.ToString.ToUpper & "_HEADERR"
        Case TemplateType.Samp
          table = "SAMPLE"
        Case TemplateType.Med
          table = ""
        Case Else
          table = template.ToString.ToUpper
      End Select
      ' For header tables, or medscreen entries, further work is required
    ElseIf entry = TableType.Header Or template = TemplateType.Med Then
      Select Case template
        Case TemplateType.Batch
          table = "Batch_Template"
        Case Else
          table = template.ToString & "_" & table
      End Select
    End If
    Return table
  End Function

#End Region ' Shared

#Region "Instance"
  '--------------------------- Instance Items -----------------------------------
  Public Event FieldEntered(ByVal sender As Object, ByVal e As EventArgs)
  Protected m_templateData As DataRow, m_type As TemplateType
  Protected m_adapters As Collections.Hashtable
  Protected m_dataSet As DataSet
  Protected WithEvents m_Entries As TemplateFields


  Protected Sub New(ByVal type As TemplateType, ByVal templateData As DataRow)
    m_adapters = New Hashtable()
    m_dataSet = New DataSet()
    m_type = type
    m_templateData = templateData
    m_Entries = New TemplateFields(m_dataSet, Me.Type, Me.TemplateID)
    Me.InitialiseAdapters()
  End Sub


  Protected Sub InitialiseAdapters()
    Dim tableName As String
    For Each tableName In Me.GetTableList
      If tableName.Trim.Length > 0 Then 'Added as on db01-dev blank tables were being returned.
        m_adapters.Add(tableName, New OleDb.OleDbDataAdapter())
        m_dataSet.Tables.Add(New DataTable(tableName))
        With Me.DataAdapter(tableName)
          Try
            Dim tableDef As MedStructure.StructTable = MedStructure.GetTable(tableName)
            Dim fieldList As IList = Me.GetFieldList(tableName)
            .SelectCommand = tableDef.GetSelectCommand(tableDef.GetFields(fieldList))
            .InsertCommand = tableDef.GetInsertCommand(tableDef.GetFields(fieldList))
            .UpdateCommand = tableDef.GetUpdateCommand(tableDef.GetFields(fieldList))
          Catch
          End Try
        End With
      End If
    Next
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the oledbDataAdapter for the specified table within the template
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property DataAdapter(ByVal tableName As String) As OleDb.OleDbDataAdapter
    Get
      Try
        Return CType(m_adapters(tableName.ToUpper), OleDb.OleDbDataAdapter)
      Catch
        Return Nothing
      End Try
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Writes template changes bck to the database
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub UpdateDataSet()
    Dim tableName As String
    ' Begin edit for first row in each table
    For Each tableName In Me.GetTableList
      Dim table As DataTable = m_dataSet.Tables(tableName)
      If table.Rows.Count = 0 Then
        table.NewRow.BeginEdit()
      Else
        table.Rows(0).BeginEdit()
      End If
    Next
    ' Update the fields
    Dim field As TemplateField
    For Each field In Me.Fields
      m_dataSet.Tables(field.TableName).Rows(0).Item(field.FieldName) = field.GetValue
    Next
    ' End edit for first row in each table
    For Each tableName In Me.GetTableList
      Dim table As DataTable = m_dataSet.Tables(tableName)
      table.Rows(0).EndEdit()
    Next
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Retrieve data for the specified row for a multi table template
  ''' </summary>
  ''' <param name="parentTable">The table for which RowID is given</param>
  ''' <param name="rowID"></param>
  ''' <remarks>
  '''   Where a template covers multiple tables, the parent table to which the other
  '''   tables link should be the one passed.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub GetData(ByVal parentTable As String, ByVal rowID As Object)
    Try
      ' Make sure the rowID is correctly formatted for the key field
      rowID = MedStructure.GetTable(parentTable).PrimaryKeys(0).FormatValue(rowID)
      ' Retrieve data for the parent table
      Me.DataAdapter(parentTable).SelectCommand.Parameters(0).Value = rowID
      Me.DataAdapter(parentTable).Fill(m_dataSet, parentTable)
      ' Now attempt to find data for all linked tables
      Dim childTable As String
      For Each childTable In Me.GetTableList
        If String.Compare(childTable, parentTable, True) <> 0 Then
          ' Attempt to get a link from the parent to the child table
          Dim linkField As MedStructure.StructField = MedStructure.GetTable(parentTable).GetLinkToTable(childTable)
          ' Failing that check for a link from child to parent table
          If linkField Is Nothing Then linkField = MedStructure.GetTable(childTable).GetLinkToTable(parentTable)
          ' If we find a link, fill in the key and fill the dataset
          If Not linkField Is Nothing Then
            Try
              DataAdapter(childTable).SelectCommand.Parameters(0).Value = m_dataSet.Tables(parentTable).Rows(0).Item(linkField.Name)
              DataAdapter(childTable).Fill(m_dataSet, childTable)
            Catch ex As Exception
              MessageBox.Show(ex.Message)
            End Try
          End If
        End If
      Next
    Catch ex As Exception
      MessageBox.Show(ex.Message)
    Finally
      MedConnection.Close()
    End Try
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Retrieves template field data for the specified row
  ''' </summary>
  ''' <param name="rowID"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub GetData(ByVal rowID As Object)
    Me.GetData(Me.GetTableList.Item(1), rowID)
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Returns the type of template in use
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Type() As TemplateType
    Get
      Return m_type
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the template ID in use
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property TemplateID() As String
    Get
      ' 1st column must be identity / template ID
      Return m_templateData.Item(0).ToString
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets all the fields associated with this template
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Fields() As TemplateFields
    Get
      Return m_Entries
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Get the template description
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Description() As String
    Get
      Return m_templateData.Item("Description").ToString
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Get templates title
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[taylor]</Author><date> [12/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public ReadOnly Property Title() As String
    Get
      Return m_templateData.Item("Title").ToString
    End Get
  End Property

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Copies data from sticky fields in the passed template
  ''' </summary>
  ''' <param name="oldTemplate"></param>
  ''' <remarks>
  '''   Primarily for use in unpacking, if sample template changes, sticky fields
  '''   are copied to the next template (if they exist) so that data does not need
  '''   to be re-entered.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Overridable Sub UpdateFromTemplate(ByVal oldTemplate As Template)
    If oldTemplate Is Me Then Exit Sub
    Me.Fields.ResetValues(True)
    Dim field As TemplateField
    For Each field In oldTemplate.Fields
      If field.IsSticky Then
        If Not Me.Fields(field.FieldName) Is Nothing Then
          Me.Fields(field.FieldName).CurrentValue = field.CurrentValue
        End If
      End If
    Next
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a collection of all mandatory fields not entered
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Overridable Function RequiredUnenteredFields() As TemplateFields
    Return Me.Fields.RequiredUnenteredFields
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Raises event when a field value is entered
  ''' </summary>
  ''' <param name="field"></param>
  ''' <param name="e"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Private Sub ValueEntered(ByVal field As TemplateField, ByVal e As EventArgs) Handles m_Entries.FieldEntered
    RaiseEvent FieldEntered(field, e)
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a list of all tables associated with the template
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function GetTableList() As Collections.Specialized.StringCollection
    Dim myTables As New Collections.Specialized.StringCollection()
    Dim field As TemplateField
    For Each field In Me.Fields
      If myTables.IndexOf(field.TableName.ToUpper) < 0 Then
        myTables.Add(field.TableName.ToUpper)
      End If
    Next
    Return myTables
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a list of all fields in the template for the specified table
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function GetFieldList(ByVal tableName As String) As Collections.Specialized.StringCollection
    Dim myFields As New Collections.Specialized.StringCollection()
    Dim field As TemplateField
    For Each field In Me.Fields
      If field.TableName.ToUpper = tableName.ToUpper Then
        myFields.Add(field.FieldName.ToUpper)
      End If
    Next
    Return myFields
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets a list of all fields in the template for the default table
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   For multi table templates, tableName shold be specified explicitly
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function GetFieldList() As Collections.Specialized.StringCollection
    Dim tableName As String = Me.GetTableList.Item(1).ToUpper
    Return GetFieldList(tableName)
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a label and control pair on parentControl for each field
  ''' </summary>
  ''' <param name="parentControl"></param>
  ''' <param name="layout"></param>
  ''' <returns>Rectangle giving size required for parentControl</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function CreatePromptControls(ByVal parentControl As Control) As Size
    Return Me.CreatePromptControls(parentControl, Nothing)
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a label and control pair on parentControl for each field
  ''' </summary>
  ''' <param name="parentControl"></param>
  ''' <param name="layout"></param>
  ''' <returns>Rectangle giving size required for parentControl</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function CreatePromptControls(ByVal parentControl As Control, ByVal layout As LayoutAttribs) As Size
    If parentControl Is Nothing Then
      Throw New ArgumentNullException("parentControl")
      Exit Function
    End If
    ' Make sure we have some layout attributes to work with
    If layout Is Nothing Then
      layout = New LayoutAttribs(parentControl, Me.Fields)
    End If
    ' Use parent controls font if none is specified in the layout
    If layout.Font Is Nothing Then layout.Font = parentControl.Font
    Me.ClearPromptControls(parentControl)
    ' Define the sizes to use for labels and prompt controls
    Dim field As Template.TemplateField
    For Each field In Me.Fields
      ' Only support prompting for bool, date, numbers, phrases (lists) and text
      If field.IsVisible Then
        ' Create a lebel and add to the parentcontrol control
        Dim labelControl As Label = field.CreateLabelControl(layout)
        parentControl.Controls.Add(labelControl)
        Dim promptControl As Control = field.CreatePromptControl(layout)
        parentControl.Controls.Add(promptControl)
        ' NOTE: Can't add data bindings until control is on the form for some reason.
        ' Bind the control to the field so value is automatically displayed / entered
        field.CreateDataBindings(promptControl)
        ' Controls arent always layout.promptsize.height tall, so take height from actual control
        layout.CurrentY += promptControl.Height + layout.ControlSpacing.Y
      End If
    Next
    Dim promptRect As Size
    If parentControl.Controls.Count = 0 Then
      promptRect = New Size(layout.TotalWidth, layout.Offset.Y)
    Else
      Dim lastControl As Control = parentControl.Controls(parentControl.Controls.Count - 1)
      promptRect = New Size(layout.TotalWidth, lastControl.Top + lastControl.Height + layout.ControlSpacing.Y)
    End If
    parentControl.Height = promptRect.Height
    Return promptRect
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Clears any prompts from the parentControl and erases references
  ''' </summary>
  ''' <param name="parentControl"></param>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Sub ClearPromptControls(ByVal parentControl As Control)
    parentControl.Controls.Clear()
    parentControl.Height = 0
    Dim field As Template.TemplateField
    For Each field In Me.Fields
      field.PromptControl = Nothing
    Next
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets an error on provider for each incomplete field
  ''' </summary>
  ''' <param name="provider"></param>
  ''' <returns>Number of errors set</returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Function FlagPromptErrors(ByVal provider As ErrorProvider) As Integer
    Dim field As TemplateField, errorCount As Integer
    For Each field In Me.Fields
      If field.FlagPromptError(provider) Then errorCount += 1
    Next
    Return errorCount
  End Function

#End Region ' Instance

#Region "Template Fields collection"
  Public Class TemplateFields
    Inherits Collections.ReadOnlyCollectionBase

    Protected Friend Event FieldEntered(ByVal sender As TemplateField, ByVal e As EventArgs)

    Protected m_adapter As OleDbDataAdapter, m_type As TemplateType

    Protected Sub InitDataAdapter()
      If m_adapter Is Nothing Then
        m_adapter = New OleDbDataAdapter()
        m_adapter.SelectCommand = New OleDbCommand()
        With m_adapter.SelectCommand
          .Connection = MedConnection.Connection
          .CommandText = "SELECT Table_Name, Field_Name, Order_Number, Default_Type, Default_Value, Text_Prompt, " & _
                                "Routine_Name, Library_Name, Prompt_Flag, Copy_Flag, Display_Flag, Mandatory_Flag"
          If m_type = TemplateType.Med Then
            .CommandText &= ", Phrase_Type"
          End If
          .CommandText &= " FROM " & TableName(TableType.Fields, m_type) & " " & _
                          "WHERE Template_ID = ?"
          If TableName(TableType.Destination, m_type) > "" Then
            .CommandText = .CommandText & " AND Table_Name = '" & TableName(TableType.Destination, m_type) & "'"
          End If
          .Parameters.Add("Template_ID", OleDbType.VarChar, 10)
        End With
      End If
    End Sub


    Protected Friend Sub New(ByVal originalData As DataSet, ByVal type As TemplateType, ByVal templateID As String)
      MyBase.new()
      m_type = type
      Me.InitDataAdapter()
      m_adapter.SelectCommand.Parameters.Item(0).Value = templateID
      Dim fieldEntries As New DataTable(TableName(TableType.Fields, m_type))
      Try
        m_adapter.Fill(fieldEntries)
        If fieldEntries.Rows.Count > 0 Then
          Dim fieldEntry As DataRow, field As TemplateField
          For Each fieldEntry In fieldEntries.Rows
            field = New TemplateField(originalData, fieldEntry)
            AddHandler field.FieldEntered, AddressOf Me.ValueEntered
            Me.Add(field)
          Next
        End If
      Catch ex As System.Exception
        Debug.WriteLine(ex.Message)
      Finally
        MedConnection.Close()
      End Try
    End Sub


    Protected Sub New()
    End Sub


    Protected Sub Add(ByVal Entry As TemplateField)
      If Me.Item(Entry.FieldName) Is Nothing Then
        MyBase.InnerList.Add(Entry)
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of all mandatory fields not entered
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function RequiredUnenteredFields() As TemplateFields
      Dim fields As New TemplateFields(), field As TemplateField
      For Each field In Me
        If field.IsMandatory And Not field.IsEntered Then fields.Add(field)
      Next
      Return fields
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the template field at the specified position
    ''' </summary>
    ''' <param name="Index"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal Index As Integer) As TemplateField
      Get
        Return CType(MyBase.InnerList.Item(Index), TemplateField)
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Get the template field of the specified name
    ''' </summary>
    ''' <param name="FieldName"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal FieldName As String) As TemplateField
      Get
        Dim field As TemplateField
        For Each field In Me
          If field.FieldName = FieldName Then
            Return field
          End If
        Next
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether all required fields are entered
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property AreSatisfactory() As Boolean
      Get
        Dim field As TemplateField
        For Each field In Me
          If Not field.IsSatisfactory Then
            Return False
          End If
        Next
        Return True
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Resets the values of all fields
    ''' </summary>
    ''' <param name="IncludeSticky">True to reset sticky fields also</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub ResetValues(ByVal IncludeSticky As Boolean)
      Dim field As TemplateField
      For Each field In Me
        field.ResetValue(IncludeSticky)
      Next
    End Sub


    Private Sub ValueEntered(ByVal sender As TemplateField, ByVal e As EventArgs)
      RaiseEvent FieldEntered(sender, e)
    End Sub


  End Class ' TemplateFields
#End Region ' Fields collection

#Region "Template Field"
  '--------------------------------------------------------------------------------
  ' Single field in a template
  Public Class TemplateField

    Protected Friend Event FieldEntered(ByVal sender As TemplateField, ByVal e As EventArgs)
    Private m_fieldData As DataRow, m_Value As Object, m_ValueEntered As Boolean
    Private m_Options As ArrayList, m_fieldInfo As MedStructure.StructField
    Private m_dataSet As DataSet, m_control As Control


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Converts the value to a common type for comparison - see cref="Accessioning.Template.TemplateField.ValueEquals">
    ''' </summary>
    ''' <param name="newValue"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Shared Function ConvertValue(ByVal newValue As Object) As String
      If newValue.GetType.IsEnum Then
        Return CInt(newValue).ToString
      Else
        Return newValue.ToString
      End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new template field based on a row from one of the template field tables
    ''' </summary>
    ''' <param name="originalData"></param>
    ''' <param name="fieldData"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Sub New(ByVal originalData As DataSet, ByVal fieldData As DataRow)
      m_dataSet = originalData
      m_fieldData = fieldData
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicate whether the passed (value type) object has the same value as this field
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ValueEquals(ByVal value As Object) As Boolean
      Return Me.GetValue = ConvertValue(value)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the index of this field in the collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Index() As Integer
      Get
        Return CType(m_fieldData.Item("Order_Number"), Integer)
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of the field the temlate field value will be written to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FieldName() As String
      Get
        Return m_fieldData.Item("Field_Name").ToString
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a MedStructure.StructField object for the destination field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FieldInfo() As MedStructure.StructField
      Get
        If m_fieldInfo Is Nothing Then
          m_fieldInfo = MedStructure.GetField(Me.TableName, Me.FieldName)
        End If
        Return m_fieldInfo
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a phrase collection to use with this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   For Medscreen templates, the Phrase_Type field can be filled in to override
    '''   any default phrase list (though it should be a subset of the one defined for
    '''   the field).  It can also be used to limit text entry to phrases for text fields
    '''   in which case the phrase text (not ID) is stored.<p/>
    '''   For other templates, the phrase in the one linked to the field being prompted for.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseList() As Glossary.PhraseCollection
      Get
        Dim myList As Glossary.PhraseCollection
        If (Not m_fieldData.Table.Columns("Phrase_Type") Is Nothing) AndAlso _
           m_fieldData.Item("Phrase_Type").ToString.Trim > "" Then
          myList = Glossary.Glossary.PhraseList(m_fieldData.Item("Phrase_Type").ToString.Trim)
        ElseIf Me.FieldInfo.IsPhraseLinked Then
          myList = Me.FieldInfo.PhraseList
        End If
        Return myList
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of the table the template field will be written to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property TableName() As String
      Get
        Return m_fieldData.Item("Table_Name").ToString
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Library in use i.e. VGL Library or .Net class
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LibraryName() As String
      Get
        Return m_fieldData.Item("Library_Name").ToString

      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of the property this field is linked to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   This is taken from the Routine_Name field
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PropertyName() As String
      Get
        ' Sample property should match field name, without underscores
        If Not Me.IsLinkedToProperty Then
          Return Me.FieldName.Replace("_", "")
        Else
          Return m_fieldData.Item("Routine_Name").ToString
        End If
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates how values are entered into the field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   See Sample Manager documentation for how templates work
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property EntryType() As FieldEntryType
      Get
        Try
          Return CType(System.Enum.ToObject(GetType(FieldEntryType), Asc(m_fieldData.Item("Default_Type").ToString)), FieldEntryType)
        Catch ex As Exception
          Console.WriteLine(ex.Message)
          Return FieldEntryType.Value
        End Try
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a value indicating whether this field links to a .NET property or VGL routine
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsLinkedToProperty() As Boolean
      Get
        Return m_fieldData.Item("Routine_Name").ToString.Trim <> ""
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the default value for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DefaultValue() As String
      Get
        Select Case Me.EntryType
          Case FieldEntryType.Value
            Return m_fieldData.Item("Default_Value").ToString.Trim
          Case Else
            ' TODO: Other EntryTypes need handling!
            Return ""
        End Select
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the link information where cref="Accession.template.TemplateField.EntryType"
    '''   is Before or After edit.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property LinkValue() As String
      Get
        Dim value As String = ""
        If Me.EntryType <> FieldEntryType.Value Then
          value = m_fieldData.Item("Default_Value").ToString.Trim
        End If
        Return value
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the link table name if cref="Accession.template.TemplateField.EntryType"
    '''   is Before or After edit.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property LinkTable() As String
      Get
        Dim tableName As String = ""
        If Me.LinkValue.IndexOf(".") > 0 Then
          Try
            tableName = MedStructure.GetField(Me.TableName, Me.LinkField).LinkTable()
          Catch
          End Try
        End If
        Return tableName
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the field name in cref="Accession.Template.TemplateField.LinkTable" if
    '''   cref="Accession.template.TemplateField.EntryType" is Before or After edit.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property LinkField() As String
      Get
        Dim fieldName As String
        If Me.LinkValue.IndexOf(".") > 0 Then
          fieldName = Me.LinkField.Substring(Me.LinkField.IndexOf("."))
        End If
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the StructField object for field cref="Accession.Template.TemplateField.LinkField"
    '''   in table cref="Accession.Template.TemplateField.LinkTable"
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [17/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LinkFieldInfo() As MedStructure.StructField
      Get
        Dim myField As MedStructure.StructField
        If Me.EntryType <> FieldEntryType.Value Then
          myField = MedStructure.GetField(Me.LinkTable, Me.LinkField)
        End If
        Return myField
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the display name (description) for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DisplayName() As String
      Get
        Return m_fieldData.Item("Text_Prompt").ToString
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether field is for display only
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsReadOnly() As Boolean
      Get
        Return m_fieldData.Item("Prompt_Flag").ToString <> C_dbTRUE
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether field value is remembered for subsequent entries
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsSticky() As Boolean
      Get
        Return m_fieldData.Item("Copy_Flag").ToString = C_dbTRUE
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether field is required
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsMandatory() As Boolean
      Get
        Return m_fieldData.Item("Mandatory_Flag").ToString = C_dbTRUE
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether field should be made visible to user
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsVisible() As Boolean
      Get
        Return Me.IsDisplay Or _
               Not Me.IsReadOnly
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether original / default value is displayed to user
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   For fields that are prompted for (ReadOnly = False) IsDisplay indicates whether
    '''   OriginalValue (or DefaultValue) is displayed in edit control, or is hidden from user.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [21/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsDisplay() As Boolean
      Get
        Return m_fieldData.Item("Display_Flag").ToString = C_dbTRUE
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a value indicating whether field value has been entered
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsEntered() As Boolean
      Get
        Return m_ValueEntered
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether field has sufficient information to be accepted
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Non-Mandatory fields are always accepted.<p/>
    '''   Otherwise only fields that have been entered (IsEntered), or where the original
    '''   value is visible (IsVisible) and not blank.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [21/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsSatisfactory() As Boolean
      Get
        ' Always return True for non-mandatory fields
        ' Otherwise only return True if field has been entered, or original value is visible and not blank
        Return (Not Me.IsMandatory) OrElse (Me.CurrentValue > "")
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a value indicating whether value should be selected from a list
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   If property returns True, PromptOptions should return two or more values
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HasSelectList() As Boolean
      Get
        Return Me.FieldInfo.SMDataType = MedStructure.SMFieldDataType.Boolean OrElse _
               Not Me.PhraseList Is Nothing
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a list of suitable values
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   If this is a Boolean or Phrase linked field, value can be selected from a list
    '''   rather than being prompted for directly.  See also HasSelectList property.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PromptOptions() As ArrayList
      Get
        If m_Options Is Nothing Then
          m_Options = New ArrayList()
          If Me.FieldInfo.SMDataType = MedStructure.SMFieldDataType.Boolean Then
            m_Options.Add("True")
            m_Options.Add("False")
          ElseIf Not Me.PhraseList Is Nothing Then
            m_Options.AddRange(Me.PhraseList)
            ' Phrase list always contains a blank entry at top, remove for mandatory fields
            If Me.IsMandatory And IsBlank(m_Options.Item(0).ToString) Then
              m_Options.RemoveAt(0)
            End If
          End If
        End If
        Return m_Options
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a string representation of the value for this field
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   If a value has been entered it is returned, if no value has been entered the
    '''   original value is returned unless blank, in which case the default value is returned.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [21/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetValue() As String
      If Me.CurrentValue > "" Then
        Return Me.CurrentValue
      ElseIf Me.OriginalValue > "" Then
        Return Me.OriginalValue
      Else
        Return Me.DefaultValue
      End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets the field value without raising an event or flagging field as entered
    ''' </summary>
    ''' <param name="newValue"></param>
    ''' <remarks>
    '''   Used for setting initial values in code
    ''' </remarks>
    ''' <history>
    ''' 	[BOUGHTON]	27/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SetValue(ByVal newValue As Object)
      m_Value = TemplateField.ConvertValue(newValue)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the string representation of the field value
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   For phrase linked fields, the value returned will be either the Phrase Text
    '''   or Phrase ID, depending on the field's PhraseLinkType property.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CurrentValue() As String
      Get
        If TypeOf m_Value Is Glossary.Phrase AndAlso _
           Me.FieldInfo.PhraseLink <> MedStructure.PhraseLinkType.Text Then
          ' Return the phrase ID
          Return CType(m_Value, Glossary.Phrase).PhraseID
        ElseIf Not m_Value Is Nothing Then
          ' Return string representation of object (phrase text for phrase objects)
          Return m_Value.ToString()
        ElseIf Me.OriginalValue > "" And Me.IsDisplay Then
          ' Display original value
          If Me.FieldInfo.IsPhraseLinked And Me.FieldInfo.PhraseLink = MedStructure.PhraseLinkType.Text Then
            ' For text linked phrases, return phrase text
            Return Me.FieldInfo.PhraseList.Item(Me.OriginalValue).ToString
          Else
            ' Return original value (may be phrase ID)
            Return Me.OriginalValue
          End If
        Else
          ' For nothing, return an empty string to prevent data binding errors
          Return ""
        End If
      End Get
      Set(ByVal Value As String)
        m_Value = Value
        If Not Value Is Nothing AndAlso Value.Trim > "" Then
          m_ValueEntered = True
          RaiseEvent FieldEntered(Me, New EventArgs())
        End If
      End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or Sets the object used to define the fields current value
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   This will normally be a string or Phrase (see MedscreenLib.Glossary).<p/>
    '''   See also CurrentValue property.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Value() As Object
      Get
        ' If value not entered, for non readonly, visible fields, return original value
        If (m_Value Is Nothing OrElse m_Value.ToString = "") AndAlso _
           Me.OriginalValue > "" And Me.IsDisplay Then
          If Me.FieldInfo.IsPhraseLinked Then
            ' Return phrase entry for original value
            Try
              Return Me.FieldInfo.PhraseList.Item(Me.OriginalValue).ToString
            Catch
              ' Original value was not in phrase list!
              Return ""
            End Try
          Else
            Return Me.OriginalValue
          End If
        Else
          Return m_Value
        End If
      End Get
      Set(ByVal Value As Object)
        m_Value = Value
        If Not Value Is Nothing AndAlso Value.ToString.Trim > "" Then
          m_ValueEntered = True
          RaiseEvent FieldEntered(Me, New EventArgs())
        End If
      End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the value currently held in the database for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property OriginalValue() As String
      Get
        Try
          Return m_dataSet.Tables(Me.TableName).Rows(0).Item(Me.FieldName).ToString.Trim
        Catch
          Return ""
        End Try
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether new entry matches original
    ''' </summary>
    ''' <value>Returns True if OriginalValue is blank or CurrentValue = OriginalValue</value>
    ''' <remarks>
    '''   When doing double entry (IsDisplay = False) it is useful to check that the
    '''   value entered by the user matches that already in the field.<p/>
    '''   Note that DefaultValue is never checked against.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [21/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ValueMatchesOriginal() As Boolean
      Get
        Return Me.OriginalValue = "" OrElse Me.CurrentValue = Me.OriginalValue
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a control for displaying field information
    ''' </summary>
    ''' <param name="displayFont"></param>
    ''' <param name="displaySize"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   For ReadOnly fields a label control is used.<p/>
    '''   For fields browsable from a list (phrases etc) a listbox is used if there
    '''   are eight or less options, otherwise a combo box is used.<p/>
    '''   Date fields are prompted for using a DateTimePicker control.<p/>
    '''   Everything else uses a TextBox.
    ''' </remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CreatePromptControl(ByVal layout As LayoutAttribs) As Control
      ' Make sure we have some layout attributes to work with
      If layout Is Nothing Then
        layout = New LayoutAttribs()
      End If
      ' Create a control to receive data
      If Me.IsReadOnly Then
        m_control = New Label()
      ElseIf Me.HasSelectList Then
        ' Value is selectable from a list, use a combo box
        If Me.PromptOptions.Count > 8 Then
          m_control = New ComboBox()
          CType(m_control, ComboBox).Items.AddRange(Me.PromptOptions.ToArray)
        Else
          m_control = New ListBox()
          CType(m_control, ListBox).Items.AddRange(Me.PromptOptions.ToArray)
        End If
      Else
        ' DBDatatype is more reliable thatn SMDataType
        Select Case Me.FieldInfo.DBDataType
          Case Data.OleDb.OleDbType.Date
            ' For dates, use a date time picker control
            m_control = New DateTimePicker()
          Case Else
            ' For everything else, use a textbox
            m_control = New TextBox()
        End Select
      End If
      ' Set the controls' font if specified
      If Not layout.Font Is Nothing Then
        m_control.Font = layout.Font
      End If
      ' Set the control's size if specified
      If Not layout.PromptSize.IsEmpty Then
        If TypeOf m_control Is ListBox Then
          ' For listbox style controls, size to fit the entire list on screen
          Dim myList As ListBox = CType(m_control, ListBox)
          myList.ItemHeight = CInt(Math.Round(m_control.Font.GetHeight)) + 1
          m_control.Size = New Size(layout.PromptSize.Width, CInt(myList.ItemHeight * myList.Items.Count) + 4)
        Else
          ' Set the size of the control as per layout
          m_control.Size = layout.PromptSize
        End If
      End If
      ' Set the control's location
      m_control.Left = layout.Offset.X + layout.LabelSize.Width + layout.ControlSpacing.X
      m_control.Top = layout.CurrentY
      ' Return the newly create control object
      Return m_control
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the contol created by CreatePromptControl
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property PromptControl() As Control
      Get
        Return m_control
      End Get
      Set(ByVal value As Control)
        m_control = value
      End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a label control to identify the field's prompt control
    ''' </summary>
    ''' <param name="layout"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CreateLabelControl(ByVal layout As LayoutAttribs) As Label
      Dim labelcontrol As New Label()
      labelcontrol.Size = layout.LabelSize
      labelcontrol.Left = layout.Offset.X
      labelcontrol.Top = layout.CurrentY
      labelcontrol.TextAlign = ContentAlignment.MiddleRight
      labelcontrol.Text = Me.DisplayName
      Return labelcontrol
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets an appropriate error on provider for this fields prompt control
    ''' </summary>
    ''' <param name="provider"></param>
    ''' <returns>True if an error was set</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function FlagPromptError(ByVal provider As ErrorProvider) As Boolean
      Dim flag As Boolean = False
      If Not Me.PromptControl Is Nothing Then
        ' TODO: Check IsSatisfactory works for Mandatory fields with and without IsDisplay flag
        If Not Me.IsSatisfactory Then
          provider.SetError(Me.PromptControl, Me.DisplayName & " must be completed")
          flag = True
        ElseIf Not Me.ValueMatchesOriginal Then
          provider.SetError(Me.PromptControl, "Value does not match original entry.")
          flag = True
        Else
          provider.SetError(Me.PromptControl, "")
        End If
      End If
      Return flag
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Binds the field to the passed control
    ''' </summary>
    ''' <param name="promptControl"></param>
    ''' <remarks>
    '''   The control should be created using <CRef/> <p/>
    '''   The control must be added to the form before data bindings are created.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub CreateDataBindings(ByVal promptControl As Control)
      If Me.HasSelectList Then
        ' Bind by object (not text) so that it works with phrases or text selections
        promptControl.DataBindings.Add("SelectedItem", Me, "Value")
      Else
        ' Bind directly to the text property of the field and control
        promptControl.DataBindings.Add("Text", Me, "CurrentValue")
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets the field to nothing and marks as unentered
    ''' </summary>
    ''' <param name="IncludeSticky">If False, Sticky fields are not reset</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Sub ResetValue(ByVal IncludeSticky As Boolean)
      If IncludeSticky Or Not Me.IsSticky Then
        m_Value = Nothing
        m_ValueEntered = False
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a string represenation of the field in the form "Field Name = 'Value'"
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
      Return Me.FieldName & " = '" & Me.GetValue & "'"
    End Function


  End Class 'TemplateField
#End Region ' Template Field

#Region "LayoutAttribs"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : Accessioning
  ''' Class	 : Template.LayoutAttribs
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Holds attributes that determine how template prmopt controls are drawn
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Class LayoutAttribs

    Private m_LabelSize As Size, m_PromptSize As Size, m_Font As Font
    Private m_Offset As Point, m_Gap As Point
    Private m_CurrentY As Integer

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Default constructor
    ''' </summary>
    ''' <remarks>
    '''   Initialises values to reasonable default values.  Font is left empty, but will
    '''   be copied from the parent control if using Template.CreatePromptControls <p/>
    '''   Default Label and Prompt size is 100 x 20 <p/>
    '''   Default offset and spacing is 8, 8
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
      Me.LabelSize = New Size(100, 20)
      Me.PromptSize = New Size(100, 20)
      Me.Offset = New Point(8, 8)
      Me.ControlSpacing = New Point(8, 8)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Constructor - sets font from parentControl and labelSize from maximum field width
    ''' </summary>
    ''' <param name="parentControl"></param>
    ''' <param name="fields"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Sub New(ByVal parentControl As Control, ByVal fields As TemplateFields)
      Me.New()
      Me.SetLabelSize(parentControl, fields)
      ' Groupboxes have an internal label at the top, so move the y offset down by label height
      If TypeOf parentControl Is GroupBox Then
        Me.Offset = New Point(Me.Offset.X, Me.LabelSize.Height)
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets Font and Size properties from parentControl and fields collection
    ''' </summary>
    ''' <param name="parentControl"></param>
    ''' <param name="fields"></param>
    ''' <remarks>
    '''   Label and Promt sizes will not be reduced beyond current sized.
    '''   If this is required, set the Label and Prompt to smaller sizes (e.g. 1, 1)
    '''   before calling this method.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetLabelSize(ByVal parentControl As Control, ByVal fields As TemplateFields)
      If parentControl Is Nothing Then Exit Sub
      Me.Font = parentControl.Font
      Dim width As Single = CSng(Me.LabelSize.Width)
      Dim g As Graphics = parentControl.CreateGraphics
      Dim field As TemplateField
      For Each field In fields
        width = Math.Max(width, g.MeasureString(field.DisplayName, Me.Font).Width)
      Next
      ' Find the maximum height
      Dim height As Integer = Math.Max(Me.Font.Height + 2, Me.LabelSize.Height)
      ' Create new size for the label control
      Me.LabelSize = New Size(CInt(Math.Round(width)) + 2, height)
      ' Ensure prompt height matches label height
      Me.PromptSize = New Size(Me.PromptSize.Width, height)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the size to use for prompt labels
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LabelSize() As Size
      Get
        Return m_LabelSize
      End Get
      Set(ByVal Value As Size)
        If Not Value.Equals(Size.Empty) Then
          m_LabelSize = Value
        End If
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the size to use for prompt controls
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property PromptSize() As Size
      Get
        Return m_PromptSize
      End Get
      Set(ByVal Value As Size)
        If Not Value.Equals(Size.Empty) Then
          m_PromptSize = Value
        End If
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the font to use for prompts and labels
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Font() As Font
      Get
        Return m_Font
      End Get
      Set(ByVal Value As Font)
        If Not Value Is Nothing Then m_Font = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the offset used to position controls
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Setting this property will reset CurrentY to the new Y offset value
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Offset() As Point
      Get
        Return m_Offset
      End Get
      Set(ByVal Value As Point)
        m_Offset = Value
        m_CurrentY = m_Offset.Y
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the spacing used between controls
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ControlSpacing() As Point
      Get
        Return m_Gap
      End Get
      Set(ByVal Value As Point)
        m_Gap = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets or sets the Y position at which the next control will be created
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CurrentY() As Integer
      Get
        Return m_CurrentY
      End Get
      Set(ByVal Value As Integer)
        m_CurrentY = Value
      End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the total width required for creating prompts
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [11/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property TotalWidth() As Integer
      Get
        Return Me.LabelSize.Width + Me.PromptSize.Width + Me.Offset.X + (Me.ControlSpacing.X * 2)
      End Get
    End Property

  End Class
#End Region

End Class
'--------------------------------------------------------------------------------
#End Region ' Template classes

