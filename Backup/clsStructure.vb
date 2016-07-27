Option Strict On
Imports MedscreenLib.Glossary

''' -----------------------------------------------------------------------------
''' <summary>
'''   Contains shared methods for retrieving table and field information
''' </summary>
''' <remarks>
'''   Gets table, field and link information from comments as created by the Parser
'''   for the current database instance as specified by CConnection.DatabaseInUse.
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
<CLSCompliant(True)> _
Public Class MedStructure

#Region "Shared"

#Region "Enumerations"
  <Flags()> _
  Public Enum FieldUsage
    None = 0
    UniqueKey = 1
    RemoveFlag = 2
    Browse = 4
    ModifiedOn = 8
    ModifiedBy = 16
    Modifiable = 32
    SecurityId = 64
    Ordering = 128
    VersionNumber = 256
    ApprovalStatus = 512
    Inspection = 1024
  End Enum

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Enumaration of phrase link types
  ''' </summary>
  ''' <remarks>
  '''   Fields can either contain the Phrase Text or Phrase Identity.
  '''    <UL>
  '''      <LI>Normal - Browse on Text, field contains Identity</LI>
  '''      <LI>Text - Browse on Text, field contains Text</LI>
  '''      <LI>Identity - Browse on Identity, field contains Identity</LI>
  '''   </UL>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Enum PhraseLinkType
    None
    Normal
    Identity
    Text
  End Enum

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Enumeration of incompatible field data types 
  ''' </summary>
  ''' <remarks>
  '''   Some datatypes used in structure.txt do not map directly to the data type used
  '''   in the database.  This enumeration contains an entry for each field datatype
  '''   used by Sample Manager.<p/>
  '''   Those with values >= 10 have individual comments and are ones that don't map
  '''   directly to the database data type.<p/>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Enum SMFieldDataType
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Datatype can not be determined
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Unknown
    Text
    [Integer]
    [Date]
    Real
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Intervals are stored as VarChar(16) in form "dddd hh:mm:ss.ff"
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Interval = 10
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Packed Decimal is stored as VarChar(10), but are numeric
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    PackedDec
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Boolean values are stored as VarChar(1) with values 'T' or 'F' -
    '''   see MedscreenLib contstants C_dbTrue and C_dbFalse.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    [Boolean]
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Identity values are stored as VarChar(n) but are upper case
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Identity
  End Enum

#End Region

  Friend Shared _Parser As Parser

  Shared Sub New()
    _Parser = New Parser(MedConnection.Instance)
  End Sub


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the list of tables currently loaded
  ''' </summary>
  ''' <value></value>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared ReadOnly Property Tables() As StructTable.TableCollection
    Get
      Return _Parser.Tables
    End Get
  End Property


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets table details for the specified table
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <returns>Nothing if table unknown</returns>
  ''' <remarks>
  '''   If the table is not already loaded, it's details are retrieved and it is added
  '''   to the Tables collection.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetTable(ByVal tableName As String) As StructTable
    Dim myTable As StructTable = Tables.Item(tableName)
    If myTable Is Nothing Then
      myTable = New StructTable(tableName)
      If myTable.Name = "" Then
        myTable = Nothing
      Else
        Tables.Add(myTable)
      End If
    End If
    Return myTable
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets details for the specified field
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <param name="fieldName"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   If the table is not already loaded, it's details are retrieved before
  '''   returning the field.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetField(ByVal tableName As String, ByVal fieldName As String) As StructField
    Dim myTable As StructTable = GetTable(tableName)
        Dim myField As StructField = Nothing
    If Not myTable Is Nothing Then
      myField = myTable.Fields.Item(fieldName)
    End If
    Return myField
  End Function


  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Sets the field in the passed datarow to the specified value
  ''' </summary>
  ''' <param name="row">Row containing field to set</param>
  ''' <param name="field">Name of field to set</param>
  ''' <param name="value">Value for field</param>
  ''' <remarks>
  '''   Value is automatically formatted for the field, so for example a boolean value
  '''   will be converted to 'T' or 'F' and Packed Decimal fields will be padded correctly.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [24/01/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Sub SetField(ByVal row As DataRow, ByVal field As String, ByVal value As Object)
    GetField(row.Table.TableName, field).SetField(row, value)
  End Sub

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a view on active and recent committed tables
  ''' </summary>
  ''' <param name="activeTable"></param>
  ''' <returns>Name of the view created in the form Recent(TableName)_View </returns>
  ''' <remarks>
  '''   The active and current committed tables are included in the view.  If the 
  '''   current committed table has been in use less than a month, the previous 
  '''   committed table is also included.<p/>
  '''   Where tables do not have committed versions, the view will contain only the
  '''   active table.<p/>
  '''   A Table_Name field is included in the view indicating the table each row was
  '''   sourced from.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function CreateRecentView(ByVal activeTable As String, ByVal startDate As Date) As String
    Return CreateRecentView(activeTable, startDate, DateTime.Now, "*", "")
  End Function

  Public Shared Function CreateRecentView(ByVal activeTable As String, ByVal startDate As Date, ByVal endDate As Date) As String
    Return CreateRecentView(activeTable, startDate, endDate, "*", "")
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Creates a view on active and recent committed table
  ''' </summary>
  ''' <param name="activeTable"></param>
  ''' <returns>Name of the view created in the form Recent(TableName)_View </returns>
  ''' <remarks>
  '''   The active and current committed tables are included in the view.  If the 
  '''   current committed table has been in use less than a month, the previous 
  '''   committed table is also included.<p/>
  '''   Where tables do not have committed versions, the view will contain only the
  '''   active table.<p/>
  '''   A Table_Name field is included in the view indicating the table each row was
  '''   sourced from.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function CreateRecentView(ByVal activeTable As String, ByVal startDate As Date, ByVal endDate As Date, ByVal fields As String, ByVal viewName As String) As String
    Return CreateRecentView(activeTable, startDate, endDate, fields, viewName, "")
  End Function

  Public Shared Function CreateRecentView(ByVal activeTable As String, ByVal startDate As Date, ByVal endDate As Date, ByVal fields As String, ByVal viewName As String, ByVal whereClause As String) As String
    Dim createView As New OleDb.OleDbCommand()
    createView.Connection = MedConnection.Connection
    If viewName = "" Then
      viewName = "Recent" & activeTable & "_View"
    End If
    If whereClause Is Nothing Then
      whereClause = ""
    ElseIf whereClause.Trim > "" AndAlso Not whereClause.ToUpper.Trim.StartsWith("WHERE ") Then
      whereClause = "WHERE " & whereClause
    End If
    Try
      If fields = "*" Then
        fields = "t.*"
      End If
      ' Create a view on the active, latest committed and last months committed tables
      createView.CommandText = "Create Or Replace View " & viewName & " As " & _
                               "Select " & fields & ", '" & activeTable & "' Table_Name From " & activeTable & _
                               " " & whereClause
      '' Find the current committed table name and add to view
      'Dim committedTable As String = GetTableName(activeTable, endDate)
      '' Only add committed tables if they exist!
      'If committedTable <> activeTable Then
      '  createView.CommandText &= " Union Select " & fields & ", '" & committedTable & "' Table_Name From " & committedTable & _
      '                            " " & whereClause
      ' Now add all other committed tables, back to requested start date
      ' Set end date to first of the month four months ago (table sets span at least four months)
      endDate = endDate.AddDays((endDate.Day * -1) + 1) '.AddMonths(-4)
      ' Set the start date to the first of it's month (table sets always start on 1st of month)
      startDate = startDate.Date.AddDays((startDate.Day * -1) + 1)
      Dim committedtable As String = activeTable
      While startDate <= endDate
        Dim previousTable As String = committedtable
        ' Find the committed table for endDate and add to view if different from previously added table
        committedtable = GetTableName(activeTable, startDate)
        If previousTable <> committedtable Then
          createView.CommandText &= " Union Select " & fields & ", '" & committedtable & "' Table_Name From " & committedtable & _
                                    " " & whereClause
        End If
        ' Jump back another four months (table sets span at least four months)
        startDate = startDate.AddMonths(4)
      End While
      'End If
      ' Open connection and create view
      MedConnection.Open()
      createView.ExecuteNonQuery()
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return viewName
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets the committed table name for a given date
  ''' </summary>
  ''' <param name="activeTable"></param>
  ''' <param name="tableDate"></param>
  ''' <returns></returns>
  ''' <remarks>
  '''   The routine automatically works out whether 1 or 3 committed table sets were
  '''   in use for the year and returns the appropriate table for the date diven.<p/>
  '''   If the passed table name does not a have committ table for the data given,
  '''   the table name is returned as passed.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [10/11/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetTableName(ByVal activeTable As String, ByVal tableDate As Date) As String
    Dim createView As New OleDb.OleDbCommand()
    createView.Connection = MedConnection.Connection
    Try
      ' Find the number of committed tables for this year
      Dim yearString As String = tableDate.Year.ToString
      createView.CommandText = "SELECT COUNT(1) FROM User_Tables " & _
                               "WHERE Table_Name LIKE ?"
            createView.Parameters.Add("Table_Name", "C_" & activeTable.ToUpper & "_" & yearString.Substring(1))
      If MedConnection.Open() Then
        Dim tableCount As Integer = Integer.Parse(createView.ExecuteScalar.ToString)
        If tableCount = 1 Then
          ' One committed table set, append year directly to table name
          activeTable = "C_" & activeTable & yearString
        ElseIf tableCount > 1 Then
          ' Three committed table sets per year, set name is 1|2|3 xxx where xxx is last 3 digits of year.
          yearString = (((tableDate.Month - 1) \ 4) + 1).ToString & yearString.Substring(1)
          activeTable = "C_" & activeTable & yearString
        End If
      End If
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return activeTable
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Indicates whether the specified table exists in the current schema
  ''' </summary>
  ''' <param name="tableName"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function TableExists(ByVal tableName As String) As Boolean
    Dim exists As Boolean
    Dim findTable As New OleDb.OleDbCommand("SELECT Table_Name FROM User_Tables " & _
                                            " WHERE Table_Name = ?", MedConnection.Connection)
        findTable.Parameters.Add("Table_Name", tableName.ToUpper)
    Try
      MedConnection.Open()
      exists = Not findTable.ExecuteScalar Is Nothing
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return exists
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Indicates whether the specified table exists in the current schema
  ''' </summary>
    ''' <param name="viewName"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function ViewExists(ByVal viewName As String) As Boolean
    Dim exists As Boolean
    Dim findTable As New OleDb.OleDbCommand("SELECT View_Name FROM User_Views " & _
                                            " WHERE View_Name = ?", MedConnection.Connection)
        findTable.Parameters.Add("View_Name", viewName.ToUpper)
    Try
      MedConnection.Open()
      exists = Not findTable.ExecuteScalar Is Nothing
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return exists
  End Function

  Public Shared Function GetCommittedTables(ByVal activeTable As String, ByVal includeActive As Boolean) As String()
    Return GetCommittedTables(activeTable, includeActive, 0)
  End Function

  Public Shared Function GetActiveTable(ByVal tableName As String) As String
    Dim activeTable As String = tableName
    If tableName.ToUpper.StartsWith("C_") Then
      activeTable = tableName.Substring(2)
      If IsAllDigits(activeTable.Substring(activeTable.Length - 4, 4)) Then
        activeTable = activeTable.Substring(0, activeTable.Length - 4)
      End If
    End If
    If activeTable Is Nothing OrElse Not TableExists(activeTable) Then
      activeTable = tableName
    End If
    Return activeTable
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Lists the committed versions of the specified active table
  ''' </summary>
  ''' <param name="activeTable"></param>
  ''' <param name="includeActive">True to list the active table as well</param>
  ''' <returns></returns>
  ''' <remarks>
  '''   Only tables that exist within the current schema are included.  If the active
  '''   table does not exist, it will not be returned, even if includeActive is True.<p/>
  '''   Committed tables are assumed to be in the form C_TABLExxxx, this may need revising
  '''   for the new LIMS, depending on our new naming convention.
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [13/12/2006]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  Public Shared Function GetCommittedTables(ByVal activeTable As String, ByVal includeActive As Boolean, ByVal minYear As Integer) As String()
    Dim listTables As New OleDb.OleDbCommand()
    Dim tableList As New ArrayList()
    If includeActive AndAlso TableExists(activeTable) Then
      tableList.Add(activeTable.ToUpper)
    End If
    listTables.Connection = MedConnection.Connection
    Try
      listTables.CommandText = "SELECT Table_Name, GetTableYear(Table_Name) Year FROM User_Tables" & _
                               " WHERE Table_Name LIKE ?" & _
                               " AND GetTableYear(Table_Name) >= ?" & _
                               " ORDER BY Year DESC, Table_Name DESC"
            listTables.Parameters.Add("Table_Name", "C_" & activeTable.ToUpper & "____")
            listTables.Parameters.Add("Year", minYear)
      MedConnection.Open()
      Dim tables As OleDb.OleDbDataReader = listTables.ExecuteReader
      While tables.Read
        Dim tableName As String = tables.GetString(0)
        If tableName.StartsWith("C_") AndAlso _
           IsAllDigits(tableName.Substring(tableName.Length - 4, 4)) Then
          tableList.Add(tableName)
        End If
      End While
      tables.Close()
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return DirectCast(tableList.ToArray(GetType(String)), String())
  End Function


  Public Shared Function GetTableName(ByVal activeTable As String, ByVal tableSet As String) As String
        Dim setIndex As String = ""
    If Not tableSet Is Nothing AndAlso tableSet.Trim.Length >= 4 Then
      setIndex = tableSet.Trim.Substring(tableSet.Trim.Length - 4)
    End If
    If tableSet = "-1" Then
      activeTable = "C_" & activeTable
    ElseIf setIndex > "" AndAlso IsNumeric(setIndex) Then
      activeTable = "C_" & activeTable & setIndex
    End If
    Return activeTable
  End Function

  Public Shared Function FindTablesWhere(ByVal activeTable As String, ByVal whereClause As String, ByVal paramValues() As Object) As String()
    Return FindTablesWhere(activeTable, whereClause, paramValues, 0, 0)
  End Function

  Public Shared Function FindTablesWhere(ByVal activeTable As String, ByVal whereClause As String, ByVal paramValues() As Object, ByVal maxCount As Integer) As String()
    Return FindTablesWhere(activeTable, whereClause, paramValues, maxCount, 0)
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="activeTable" type="String"></param>
  ''' <param name="whereClause" type="String"></param>
  ''' <param name="paramValues" type="Object"></param>
  ''' <param name="maxCount" type="Int32"></param>
  ''' <param name="minYear" type="Int32"></param>
  ''' <returns>
  ''' </returns>
  ''' <remarks>
  ''' </remarks>
  ''' -----------------------------------------------------------------------------
  Public Shared Function FindTablesWhere(ByVal activeTable As String, ByVal whereClause As String, ByVal paramValues() As Object, ByVal maxCount As Integer, ByVal minYear As Integer) As String()
    Dim tableName As String
    Dim tableList As New ArrayList()
    ' Create a command to search the table(s)
    Dim search As New OleDb.OleDbCommand()
    search.Connection = MedConnection.Connection
    Try
      ' Fill in command parameters for where clause
      Dim params As OleDb.OleDbParameter() = GetTable(activeTable).GetParameters(whereClause, "")
      Dim paramIndex As Integer
      For paramIndex = 0 To params.Length - 1
        search.Parameters.Add(params(paramIndex)).Value = paramValues(paramIndex)
      Next
      ' Get a list of all tables and search each one
      For Each tableName In GetCommittedTables(activeTable, True, minYear)
        MedConnection.Open()
        'Console.WriteLine("Search {0} at " & Now.ToString("mm:ss.ff"), tableName)
        search.CommandText = "SELECT Count(1) FROM " & tableName & " WHERE " & whereClause
        Dim result As Object = search.ExecuteScalar
        If TypeOf result Is Decimal AndAlso DirectCast(result, Decimal) > 0 Then
          tableList.Add(tableName)
          If maxCount > 0 AndAlso tableList.Count >= maxCount Then Exit For
        End If
      Next
    Catch ex As System.IndexOutOfRangeException
      Throw New IndexOutOfRangeException("Insufficient parameters passed for WHERE clause", ex)
    Catch ex As Exception
      Medscreen.LogError(ex)
    Finally
      MedConnection.Close()
    End Try
    Return DirectCast(tableList.ToArray(GetType(String)), String())
  End Function

#End Region

#Region "Parser"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : Structure Parser
  ''' Class	 : Parser
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Gets table, field and link information from table structure definition
  '''   for storing in User_Tab_Comments and User_Col_Comments.
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  <CLSCompliant(True)> _
Public Class Parser

#Region "Shared"

    Private Shared binChars As Char() = New Char() {Chr(10), Chr(13), Chr(9)}

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Create table and column comments for specified database instance
    ''' </summary>
    ''' <param name="instance"></param>
    ''' <returns>The parser object created</returns>
    ''' <remarks>
    '''   Although the object is returned, the processing has already been done.<p/>
    '''   To create a parser based on an instance without doing any work use
    '''   <code>dim myParse as New Parser(instance)</code> 
    '''   where instance is a member of the CConnection.UseDatabase enumeration
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function ParseStructure(ByVal instance As MedConnection) As Parser
      Dim myParser As New Parser(instance)
      myParser.ParseStructure()
            myParser.CreateComments(MedConnection.Connection)
            Return myParser
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the next section of the file from startPos
    ''' </summary>
    ''' <param name="text">The text to search</param>
    ''' <param name="startPos">Position to start searching</param>
    ''' <returns>Section between startPos and next ; character</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Shared Function ReadSection(ByRef text As String, ByRef startPos As Integer) As String
            Dim section As String = ""
      Dim sectEnd As Integer = text.IndexOf(";", startPos)
      If sectEnd >= 0 Then
        section = text.Substring(startPos, sectEnd - startPos).Trim
        Dim ch As Char
        For Each ch In binChars
          section = section.Replace(ch, " ")
        Next
        startPos = sectEnd + 1
        section = section.Trim
      Else
        startPos = text.Length
      End If
      Return section
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the name of the table or field defined in definition
    ''' </summary>
    ''' <param name="definition">A section (from ReadSection)</param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Table and field names are always returned trimmed an in upper case
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Shared Function GetItemName(ByVal definition As String) As String
      Dim namestart As Integer = definition.IndexOf(" ")
            Dim name As String = ""
      If namestart >= 0 Then
        name = definition.Substring(namestart).Trim
        Dim nameEnd As Integer = name.IndexOf(" ")
        If nameEnd < 0 Then nameEnd = name.IndexOf(";")
        If nameEnd < 0 Then nameEnd = name.Length
        name = name.Substring(0, nameEnd).Trim.ToUpper
      End If
      Return name
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Removes all comments enclosed in curly braces from internal structure data
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripComments(ByVal vglCode As String) As String
            If vglCode Is Nothing Then
                Return ""
                Exit Function
            End If
      ' Using a string builder is at least 10 times quicker than working on the string itself
      ' because a string is immutable - a new object is created for each change!
      Dim definition As New Text.StringBuilder(vglCode)
      Dim cmtStart As Integer = vglCode.IndexOf("{")
      Dim cmtEnd As Integer, removed As Integer, totalRemoved As Integer
      While cmtStart >= 0 And cmtStart <= vglCode.Length
        Console.Write(".")
        cmtEnd = vglCode.IndexOf("}", cmtStart)
        If cmtEnd >= 0 Then
          removed = cmtEnd + 1 - cmtStart
          definition.Remove(cmtStart - totalRemoved, removed)
          totalRemoved += removed
        End If
        cmtStart = vglCode.IndexOf("{", cmtStart + 1)
      End While
      Console.WriteLine()
      Return definition.ToString
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the nearest matching oleDBType to the specified type name
    ''' </summary>
    ''' <param name="dataType"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Shared Function GetOleDbType(ByVal dataType As String) As OleDb.OleDbType
      Select Case dataType.ToUpper
        Case "VARCHAR2"
          Return Data.OleDb.OleDbType.VarChar
        Case "FLOAT"
          Return Data.OleDb.OleDbType.Double
        Case "NUMBER"
          Return Data.OleDb.OleDbType.Decimal
        Case "LONG"
          Return Data.OleDb.OleDbType.LongVarChar
        Case Else
          Try
            Return CType([Enum].Parse(GetType(OleDb.OleDbType), dataType, True), OleDb.OleDbType)
          Catch
            Return Data.OleDb.OleDbType.IUnknown
          End Try
      End Select
    End Function

#End Region

#Region "Instance"

    Private _file As String, _contents As String, _tables As New StructTable.TableCollection()
    Private _conn As OleDb.OleDbConnection


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new parser for the specified database instance
    ''' </summary>
    ''' <param name="Instance"></param>
    ''' <remarks>
    '''   The created object can be used to parse the structure.txt file, or to retrieve
    '''   table and column information from database comments.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal instance As MedConnection)
      If Not instance Is Nothing Then
        Dim path As String = instance.ServerPath & "\Data"
        If IO.Directory.Exists(path & "\Custom") Then path &= "\Custom"
        _file = path & "\structure.txt"
                _conn = MedConnection.Connection
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new parser for the specified structure file
    ''' </summary>
    ''' <param name="StructureFile"></param>
    ''' <remarks>
    '''   The created object can only be used for parsing the structure file, it cannot
    '''   retrieve data from database comments.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal structureFile As String)
      _file = structureFile
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Parses the structure file
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub ParseStructure()
      Dim structureDef As String
      If System.IO.File.Exists(_file) Then
        ' Read the entire contents of the file
        Console.Write("Reading file and stripping comments")
        With System.IO.File.OpenText(_file)
          structureDef = .ReadToEnd
          .Close()
        End With
        ' Remove any comments sections
        structureDef = Parser.StripComments(structureDef)
        ' Loop through file pickup up sections and processing
        Dim currentChar As Integer, section As String
        While currentChar >= 0 And currentChar < structureDef.Length
          ' Get the next section in the file
          section = ReadSection(structureDef, currentChar)
          If Not section Is Nothing Then
            ' If it's a table definition, add to our tables collection
            If section.ToUpper.StartsWith("TABLE ") Or _
               section.ToUpper.StartsWith("VIEW ") Then
              section = section.Trim
              Dim myTable As New StructTable(section, structureDef, currentChar)
              Me.Tables.Add(myTable)
            End If
          End If
        End While
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns a collection of all tables from structure
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Tables() As StructTable.TableCollection
      Get
        Return _tables
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates comments on the current instance
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   Parser must have been created using the constructor taking a MedConnection
    '''   object.  If the structure File constructor was used, call CreateComments
    '''   passing in an oleDbConnection.
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	24/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CreateComments() As Boolean
      If Not _conn Is Nothing Then
        Me.CreateComments(_conn)
      End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Attempts to create table and column comments in the database
    ''' </summary>
    ''' <param name="Connection"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CreateComments(ByVal connection As OleDb.OleDbConnection) As Boolean
      Dim failed As Boolean
      Try
        If MedConnection.Open(connection) Then
          Console.WriteLine("Updating comments")
          Dim table As StructTable
          For Each table In Me.Tables
            table.CreateComments(connection)
          Next
          Console.WriteLine("Update complete")
        End If
        MedConnection.Close(connection)
      Catch
        failed = True
      End Try
      Return Not failed
    End Function

#End Region

  End Class
#End Region

#Region "Table"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : Structure Parser
  ''' Class	 : Parser.StructTable
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Contains name and fields information for a table<p/>
  '''   Contains routines to parse table definition to extract the data.
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  <CLSCompliant(True)> _
Public Class StructTable

    Private _fields As StructField.FieldCollection
    Private _name As String
    Private _tables As TableCollection
    Private _parent As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new table based on structure file
    ''' </summary>
        ''' <param name="definition"></param>
    ''' <param name="text"></param>
    ''' <param name="startPos"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub New(ByVal definition As String, ByVal text As String, ByRef startPos As Integer)
            Dim section As String = "", newTable As Boolean
      _fields = New StructField.FieldCollection()
      ' Find the name of this table
      _name = Parser.GetItemName(definition)
      ' Check whether it is a child of another table
      Dim parentStart As Integer = definition.ToUpper.IndexOf("PARENT_TABLE")
      If parentStart > 0 Then
        _parent = Parser.GetItemName(definition.Substring(parentStart))
      End If
      Console.WriteLine("Parsing Table " & _name)
      ' Now look for field definitions
      While startPos < text.Length And Not newTable
        section = Parser.ReadSection(text, startPos)
        If Not section Is Nothing Then
          If section.ToUpper.StartsWith("FIELD ") Then
            section = section.Trim
            Dim myField As New StructField(Me, section)
            _fields.Add(myField)
          Else
            newTable = True
          End If
        End If
      End While
      If newTable Then
        startPos -= startPos - InStrRev(text, ";", startPos - (section.Length + 1))
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new table based on database comments
    ''' </summary>
        ''' <param name="TableName"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub New(ByVal tableName As String)
      Try
        tableName = tableName.ToUpper
        ' Find table comment
        Dim getComment As New OleDb.OleDbCommand("SELECT Comments FROM User_Tab_COmments WHERE Table_Name = ?", _
                                                  MedConnection.Connection)
                getComment.Parameters.Add("Table", tableName)
        Try
          MedConnection.Open()
          Me.ParseComments(getComment.ExecuteScalar.ToString)
        Catch ex As Exception
          Medscreen.LogError(ex)
        Finally
          MedConnection.Close()
        End Try
        ' For committed tables, use active tables field collection
        If Me.IsChild Then
          _name = tableName
          _fields = Me.ActiveTable.Fields
        Else
          ' Get field details and comments
          Dim daFields As New OleDb.OleDbDataAdapter("SELECT d.Column_Name, d.Data_Type, d.Data_Length, c.Comments" & _
                                                     "  FROM User_Col_Comments c, User_Tab_Columns d" & _
                                                     " WHERE d.Table_Name = ? " & _
                                                     "   AND c.Table_Name = d.Table_Name and c.Column_Name = d.Column_Name", _
                                                     MedConnection.Connection)
          daFields.SelectCommand.Parameters.Add("Table_Name", OleDb.OleDbType.VarChar).Value = tableName
          Dim fieldsTable As New DataTable()
          daFields.Fill(fieldsTable)
          If fieldsTable.Rows.Count > 0 Then
            _name = tableName
            _fields = New StructField.FieldCollection()
          End If
          Dim fieldRow As DataRow
          For Each fieldRow In fieldsTable.Rows
            Me.Fields.Add(New StructField(Me, fieldRow))
          Next
        End If
      Catch ex As Exception
        Console.WriteLine(ex.Message)
      Finally
        MedConnection.Close()
      End Try
    End Sub

    Private Sub ParseComments(ByVal comments As String)
      Dim comment As String, key As String, value As String
      If Not IsBlank(comments) Then
        For Each comment In comments.Split(";"c)
          If comment.IndexOf("=") > 0 Then
            key = comment.Substring(0, comment.IndexOf("="))
            value = comment.Substring(comment.IndexOf("=") + 1)
            Select Case key
              Case "PARENT"
                _parent = value
            End Select
          End If
        Next
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of the table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
      Get
        Return _name
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the name of the parent table for committed tables
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ParentName() As String
      Get
        Return _parent
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the parent table for committed tables
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ParentTable() As StructTable
            Get
                Dim myTable As StructTable = Nothing
                If Me.IsChild Then
                    myTable = GetTable(Me.ParentName)
                End If
                Return myTable
            End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether this table is a child (committed) table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsChild() As Boolean
      Get
        Return Not _parent Is Nothing
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the Active table for this table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   For tables without children (committed tables) this instance is returned.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [20/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ActiveTable() As StructTable
      Get
        If Me.IsChild Then
          Return Me.ParentTable
        Else
          Return Me
        End If
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the collection of fields for the table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Fields() As StructField.FieldCollection
      Get
        Return _fields
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Populates the data row with default values for this table
    ''' </summary>
        ''' <param name="newrow"></param>
    ''' <remarks>
    '''   Field values are only changed if they are Null and field has a default
    '''   value defined in structure.txt
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub PopulateDataRow(ByVal newRow As DataRow)
      If newRow Is Nothing Then
        Throw New ArgumentNullException("row")
      ElseIf newRow.Table.TableName.ToUpper.Trim <> Me.Name Then
        Throw New ArgumentOutOfRangeException("row", "Row does not belong to this table.")
      End If
      Dim column As DataColumn
      newRow.BeginEdit()
      For Each column In newRow.Table.Columns
        Try
          Dim field As StructField = Me.Fields(column.ColumnName)
          If newRow.IsNull(column) AndAlso _
             Not field Is Nothing AndAlso _
             field.HasDefaultValue AndAlso _
             field.IsValidValue(field.DefaultValue) Then
            ' Only populate field if value is null and default is specified and is valid
            newRow.Item(column) = field.DefaultValue
          End If
        Catch ex As Exception
          Console.WriteLine(column.ColumnName)
        End Try
      Next
      newRow.EndEdit()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Builds a select query based on primary key and browse information
    ''' </summary>
    ''' <param name="linkField"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function BuildSelect(ByVal linkField As String, ByVal whereClause As String) As String
      Dim clause As New System.Text.StringBuilder(linkField)
      Dim browse As New Collections.SortedList()
      Dim field As StructField
            Dim orderfield As String = ""

            Dim strReturn As String = ""
            For Each field In Me.Fields
                If field.Name <> linkField Then
                    If field.IsPrimaryKey Then
                        clause.Append(", " & field.Name)
                    ElseIf field.DisplayIndex >= 0 Then
                        browse.Add(field.DisplayIndex, field)
                    End If
                    If field.IsOrder Then orderfield = field.Name
                End If
            Next
            For Each field In browse.Values
                If Not field Is Nothing Then
                    clause.Append(", " & field.Name)
                End If
            Next
            If clause.Length > 0 Then
                If clause.Chars(0) = ","c Then clause.Remove(0, 2)
                clause.Append(" FROM " & Me.Name)
                If Not IsBlank(whereClause) Then clause.Append(whereClause)
                If orderfield > "" Then clause.Append(" ORDER BY " & orderfield)
                strReturn = "SELECT " & clause.ToString
            End If
            Return strReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a field collection from an enumeratable list of field names
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [16/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFields(ByVal fields As IEnumerable) As StructField.FieldCollection
      Dim field As String, myFields As New StructField.FieldCollection()
      For Each field In fields
        If Not Me.Fields(field) Is Nothing Then
          myFields.Add(Me.Fields(field))
        End If
      Next
      Return myFields
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a field collection from a data column collection
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [16/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFields(ByVal fields As DataColumnCollection) As StructField.FieldCollection
      Dim field As DataColumn, myFields As New StructField.FieldCollection()
      For Each field In fields
        If Not Me.Fields(field.ColumnName) Is Nothing Then
          myFields.Add(Me.Fields(field.ColumnName))
        End If
      Next
      Return myFields
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a field collection from a comma delimited list of field names
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [16/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFields(ByVal fields As String) As StructField.FieldCollection
      Return GetFields(fields.Replace(" ", "").Split(","c))
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a FieldCollection from an array of fields
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [16/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function Getfields(ByVal fields() As StructField) As StructField.FieldCollection
      Dim field As StructField, myFields As New StructField.FieldCollection()
      For Each field In fields
        myFields.Add(field)
      Next
      Return myFields
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a select command based on the fields required with specified where clause
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Note field names are formatted lower case in the SQL command for ease of
    '''   readability when debugging.  Primary Keys are always included at the start of
    '''   the select statement.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectCommand(ByVal fields As StructField.FieldCollection, ByVal whereFields As StructField.FieldCollection) As OleDb.OleDbCommand
      Dim whereClause As New System.Text.StringBuilder(" WHERE ")
      Dim field As StructField
      ' Create where clause
      For Each field In whereFields
        whereClause.Append(field.Name.ToLower & " = ? AND ")
      Next
            ' Remove the final AND
            If whereClause.Length = 7 Then
                whereClause.Remove(0, 7)
            Else
                whereClause.Remove(whereClause.Length - 4, 4)
            End If

            Return GetSelectCommand(fields, whereClause.ToString)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a select command based on the fields required with specified where clause
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Note field names are formatted lower case in the SQL command for ease of
    '''   readability when debugging.  Primary Keys are always included at the start of
    '''   the select statement.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectCommand(ByVal fields As StructField.FieldCollection, ByVal whereClause As String) As OleDb.OleDbCommand
      Dim selectCommand As New OleDb.OleDbCommand("SELECT")
      Dim mySQL As New System.Text.StringBuilder()
      Dim field As StructField
      ' Always put primary keys first in select list, even if not asked for
      For Each field In Me.PrimaryKeys
        mySQL.Append(", " & field.Name)
      Next
      ' List the other fields to be selectd
      For Each field In fields
        If Not field.IsPrimaryKey Then mySQL.Append(", " & field.Name.ToLower)
      Next
      ' Create where clause
      If whereClause Is Nothing OrElse _
         Not (whereClause.Trim.ToUpper.StartsWith("WHERE") Or _
              whereclause.trim.length = 0) Then
        whereClause = "WHERE " & whereClause
      End If
      mySQL.Append(" FROM " & Me.Name.ToLower & " " & whereClause)
      ' Add parameters
      Dim param As OleDb.OleDbParameter
      For Each param In GetParameters(whereClause, "")
        selectCommand.Parameters.Add(param)
      Next
      '  Add order field
      If whereClause.ToUpper.IndexOf(" ORDER ") < 0 Then
        Dim order As StructField = Me.GetFieldFor(FieldUsage.Ordering)
        If Not order Is Nothing Then
          mySQL.Append(" ORDER BY " & order.Name)
        End If
      End If
      ' Remove first comma
      mySQL.Remove(0, 1)
      selectCommand.CommandText &= mySQL.ToString
      selectCommand.Connection = MedConnection.Connection
      Return selectCommand
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets parameters required for a where clause
    ''' </summary>
    ''' <param name="whereClause"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Rather simplistic approach, but works for majority of queries.  Feel free
    '''   to enhance!
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [24/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Function GetParameters(ByVal whereClause As String, ByVal prefix As String) As OleDb.OleDbParameter()
      Dim params As New ArrayList()
      ' Look for ?
      Dim findPar As New System.Text.RegularExpressions.Regex("\?")
      ' Loop backwards for field name
      Dim findName As New System.Text.RegularExpressions.Regex("[A-Za-z0-9_]+", _
                                                               Text.RegularExpressions.RegexOptions.RightToLeft)
      Dim parEnd As Text.RegularExpressions.Match
      For Each parEnd In findPar.Matches(whereClause)
        Dim parName As Text.RegularExpressions.Match
        Dim startIndex As Integer = parEnd.Index
        Do
          parName = findName.Match(whereClause, startIndex)
          If Not parName Is Nothing Then
            startIndex = parName.Index
            If " AND OR LIKE BETWEEN ".IndexOf(" " & parName.Value.ToUpper & " ") < 0 Then
              Dim fieldDef As MedStructure.StructField = Me.Fields(parName.Value)
              If Not fieldDef Is Nothing Then
                Dim param As New OleDb.OleDbParameter(prefix & fieldDef.Name, fieldDef.DBDataType, fieldDef.DBDataLength)
                param.SourceColumn = fieldDef.Name
                startIndex = 0
                ' Check if parameter name is already used and append an index if it is
                Dim oldParam As OleDb.OleDbParameter
                For Each oldParam In params
                  If oldParam.ParameterName = param.ParameterName Then
                    startIndex += 1
                    param.ParameterName = prefix & fieldDef.Name & "_" & startIndex
                  End If
                Next
                params.Add(param)
                ' Parameter found, Exit loop
                Exit Do
              End If
            End If
          End If
        Loop Until parName Is Nothing Or startIndex = 0
      Next
      Return CType(params.ToArray(GetType(OleDb.OleDbParameter)), OleDb.OleDbParameter())
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a select command based on the fields required
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [16/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectCommand(ByVal fields As StructField.FieldCollection) As OleDb.OleDbCommand
      Return Me.GetSelectCommand(fields, GetFields(Me.PrimaryKeys))
    End Function

    Public Function GetSelectCommand(ByVal whereClause As String) As OleDb.OleDbCommand
      Return Me.GetSelectCommand(Me.Fields, whereClause)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a select command for all fields in the table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectCommand() As OleDb.OleDbCommand
      Return Me.GetSelectCommand(Me.Fields)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an update command based on the fields required
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Note field names are formatted lower case in the SQL command for ease of
    '''   readability when debugging.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetUpdateCommand(ByVal fields As StructField.FieldCollection, ByVal whereClause As String) As OleDb.OleDbCommand
      Dim updateCommand As New OleDb.OleDbCommand("UPDATE " & Me.Name & " SET")
      Dim mySQL As New System.Text.StringBuilder()
      Dim field As StructField
      ' List the fields to be updated
      For Each field In fields
        'If Not field.IsPrimaryKey Then
        If field.IsUsedFor(FieldUsage.ModifiedOn) Then
          mySQL.Append(",  " & field.Name & " = SYSDATE")
        Else
          mySQL.Append(", " & field.Name.ToLower & " = ?")
          updateCommand.Parameters.Add(field.Name.ToLower, field.DBDataType, field.DBDataLength).SourceColumn = field.Name
        End If
        'End If
      Next
      ' Remove first comma
      If mySQL.Length > 0 Then mySQL.Remove(0, 1)
      ' Create where clause
      If whereClause Is Nothing OrElse _
         Not whereClause.Trim.ToUpper.StartsWith("WHERE") Then
        whereClause = "WHERE " & whereClause
      End If
      mySQL.Append(" " & whereClause)
      ' Add parameters
      Dim param As OleDb.OleDbParameter
      For Each param In GetParameters(whereClause, "W_")
        updateCommand.Parameters.Add(param)
      Next
      updateCommand.CommandText &= mySQL.ToString
      updateCommand.Connection = MedConnection.Connection
      Return updateCommand
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an update command based on the fields required
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   Note field names are formatted lower case in the SQL command for ease of
    '''   readability when debugging.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetUpdateCommand(ByVal fields As StructField.FieldCollection) As OleDb.OleDbCommand
      Dim updateCommand As New OleDb.OleDbCommand("UPDATE " & Me.Name & " SET")
      Dim mySQL As New System.Text.StringBuilder()
      Dim field As StructField
      ' List the fields to be updated
      For Each field In fields
        'If Not field.IsPrimaryKey Then
        mySQL.Append(", " & field.Name.ToLower & " = ?")
        updateCommand.Parameters.Add(field.Name.ToLower, field.DBDataType, field.DBDataLength).SourceColumn = field.Name
        'End If
      Next
      ' Remove first comma
      If mySQL.Length > 0 Then mySQL.Remove(0, 1)
      ' Create where clause
      mySQL.Append(" WHERE ")
      For Each field In Me.PrimaryKeys
        mySQL.Append(field.Name.ToLower & " = ? AND ")
        With updateCommand.Parameters.Add("K_" & field.Name.ToLower, field.DBDataType, field.DBDataLength)
          .SourceColumn = field.Name
          .SourceVersion = DataRowVersion.Original
        End With
      Next
      ' Remove the final AND
      mySQL.Remove(mySQL.Length - 4, 4)
      updateCommand.CommandText &= mySQL.ToString
      updateCommand.Connection = MedConnection.Connection
      Return updateCommand
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an update command for all fields in the table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetUpdateCommand() As OleDb.OleDbCommand
      Return Me.GetUpdateCommand(Me.Fields)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an insert command based on the fields required
    ''' </summary>
    ''' <param name="fields"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   The primary key field(s) are always listed whether or not they are included
    '''   in the passed collection.<p/>
    '''   Note: field names are formatted lower case in the SQL command for ease of
    '''   readability when debugging.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetInsertCommand(ByVal fields As StructField.FieldCollection) As OleDb.OleDbCommand
      Return Me.GetInsertCommand(fields, True)
    End Function

    Public Function GetInsertCommand(ByVal fields As StructField.FieldCollection, ByVal includeKeys As Boolean) As OleDb.OleDbCommand
      Return GetInsertCommand(fields, includeKeys, "")
    End Function

    Public Function GetInsertCommand(ByVal fields As StructField.FieldCollection, ByVal sequenceID As String) As OleDb.OleDbCommand
      Return GetInsertCommand(fields, True, sequenceID)
    End Function

    Private Function GetInsertCommand(ByVal fields As StructField.FieldCollection, ByVal includeKeys As Boolean, ByVal sequenceID As String) As OleDb.OleDbCommand
      Dim insertCommand As New OleDb.OleDbCommand("INSERT INTO " & Me.Name & " (")
      Dim mySQL As New System.Text.StringBuilder()
      Dim field As StructField
      ' List the primary keys first
      If includeKeys Then
        For Each field In Me.PrimaryKeys
          mySQL.Append(", " & field.Name.ToLower)
          If sequenceID = "" Then
            insertCommand.Parameters.Add(field.Name.ToLower, field.DBDataType, field.DBDataLength).SourceColumn = field.Name
          End If
        Next
      End If
      ' Remove first comma DT Added check for length being greater than 0 
      If mySQL.Length > 0 Then mySQL.Remove(0, 2)
      ' List the fields to be insertd
      For Each field In fields
        ' Omit Key fields from this list, unless they were not included earlier
        If (Not field.IsUsedFor(FieldUsage.ModifiedOn)) AndAlso _
           (includeKeys = False OrElse Not field.IsUsedFor(FieldUsage.UniqueKey)) Then
          mySQL.Append(", " & field.Name.ToLower)
          insertCommand.Parameters.Add(field.Name.ToLower, field.DBDataType, field.DBDataLength).SourceColumn = field.Name
        End If
      Next
      ' Remove first comma if necessary
      If mySQL.Length > 0 And Not includeKeys Then
        mySQL.Remove(0, 2)
      End If
      ' Add modified on field
      field = Me.GetFieldFor(FieldUsage.ModifiedOn)
      If Not field Is Nothing Then
        mySQL.Append(", " & field.Name.ToLower)
      End If
      ' Create the VALUES section
      mySQL.Append(") VALUES (")
      ' Include the sequence name for tables keyed by a sequence
      If sequenceID > "" Then
        mySQL.Append(sequenceID & ".NextVal,")
      End If
      ' Insert one value per parameter
      mySQL.Insert(mySQL.Length, "?,", insertCommand.Parameters.Count)
      ' Add SYSDATE for Modified on field and close VALUES section bracket
      If Not field Is Nothing Then
        mySQL.Append("SYSDATE)")
      Else
        mySQL.Chars(mySQL.Length - 1) = ")"c
      End If
      insertCommand.CommandText &= mySQL.ToString
      insertCommand.Connection = MedConnection.Connection
      Return insertCommand
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an insert command for all fields in the table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [10/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetInsertCommand() As OleDb.OleDbCommand
      Return Me.GetInsertCommand(Me.Fields)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a delete command based on primary key(s) in table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [23/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetDeleteCommand() As OleDb.OleDbCommand
      Dim deleteCommand As New OleDb.OleDbCommand("DELETE FROM " & Me.Name & " WHERE ")
      Dim mySQL As New System.Text.StringBuilder()
      Dim field As StructField
      For Each field In Me.PrimaryKeys
        mySQL.Append(field.Name.ToLower & " = ? AND ")
        With deleteCommand.Parameters.Add(field.Name.ToLower, field.DBDataType, field.DBDataLength)
          .SourceColumn = field.Name
          .SourceVersion = DataRowVersion.Original
        End With
      Next
      If mySQL.Length > 0 Then
        ' Remove the final AND
        mySQL.Remove(mySQL.Length - 4, 4)
        deleteCommand.CommandText &= mySQL.ToString
        deleteCommand.Connection = MedConnection.Connection
      Else
        deleteCommand = Nothing
      End If
      Return deleteCommand
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the number of primary key fields in the table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function KeyCount() As Integer
      Dim field As StructField, keys As Integer = 0
      For Each field In Me.Fields
        If field.IsPrimaryKey Then
          keys = keys + 1
        End If
      Next
      Return keys
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns an array of the table's primary key fields in order
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PrimaryKeys() As StructField()
      Get
        Dim field As StructField, keyFields As New SortedList()
        For Each field In Me.GetFieldsWithUse(FieldUsage.UniqueKey)
          keyFields.Add(field.KeyIndex, field)
        Next
        ' Copy the (now sorted) list of fields into an array
        Dim orderedFields(keyFields.Count - 1) As StructField
        keyFields.Values.CopyTo(orderedFields, 0)
        Return orderedFields
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the first field with the specified use
    ''' </summary>
    ''' <param name="use"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [02/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFieldFor(ByVal use As FieldUsage) As StructField
            Dim field As StructField = Nothing
            For Each field In Me.Fields
                If field.IsUsedFor(use) Then
                    Exit For
                End If
                field = Nothing                     'Null out variable so if we spin through all the fields and find nothing we will have a null return
            Next
            Return field
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether table has a field for the specified use
    ''' </summary>
    ''' <param name="use"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [02/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HasFieldFor(ByVal use As FieldUsage) As Boolean
      Get
        Return Not GetFieldFor(use) Is Nothing
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets an array of all fields with the specified use
    ''' </summary>
    ''' <param name="use"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [02/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFieldsWithUse(ByVal use As FieldUsage) As StructField()
      Dim field As StructField, useFields As New ArrayList()
      For Each field In Me.Fields
        If field.IsUsedFor(use) Then
          useFields.Add(field)
        End If
      Next
      Return DirectCast(useFields.ToArray(GetType(StructField)), StructField())
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns an array of the table's browse fields in order
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property BrowseFields() As StructField()
      Get
        Dim field As StructField, keyFields As New SortedList()
        For Each field In Me.GetFieldsWithUse(FieldUsage.Browse)
          keyFields.Add(field.DisplayIndex, field)
        Next
        ' Copy the (now sorted) list of fields into an array
        Dim orderedFields(keyFields.Count - 1) As StructField
        keyFields.Values.CopyTo(orderedFields, 0)
        Return orderedFields
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns an array of data relations for the table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   Note that the relations must be added to a dataset before they can be used.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetRelations() As DataRelation()
      Dim relations As New ArrayList()
      Dim field As StructField
      For Each field In Me.Fields
        Dim relation As DataRelation = field.GetRelation
        If Not relation Is Nothing Then
          relations.Add(relation)
        End If
      Next
      Return CType(relations.ToArray(GetType(DataRelation)), DataRelation())
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the StructField that links to the specified table
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns>Nothing if no link found</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetLinkToTable(ByVal tableName As String) As StructField
            Dim field As StructField = Nothing, retField As StructField = Nothing
      For Each field In Me.Fields
        If String.Compare(field.LinkTable, tableName, True) = 0 Then
          retField = field
          Exit For
        End If
      Next
      Return retField
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the string that is written to the User_Tab_Comments table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Comments() As String
      If Me.IsChild Then
        Return "PARENT=" & Me.ParentName
      Else
        Return Me.Name
      End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a string detailing the table name and comments
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
      Return Me.Name & ":" & Me.Comments
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates entries in User_Tab_Comments and User_Col_Comments
    ''' </summary>
    ''' <param name="OpenConnection"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub CreateComments(ByVal OpenConnection As OleDb.OleDbConnection)
      If Me.CreateComment(OpenConnection) Then
        Console.WriteLine("Writing comments for " & Me.Name)
        Dim field As StructField
        For Each field In Me.Fields
          field.CreateComment(OpenConnection)
        Next
      End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an entry in User_Tab_Comments
    ''' </summary>
    ''' <param name="OpenConnection"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function CreateComment(ByVal OpenConnection As OleDb.OleDbConnection) As Boolean
      Dim failed As Boolean
      failed = Not TableExists(Me.Name)
      If (Not failed) AndAlso Me.Comments > "" Then
        Dim createCmd As New OleDb.OleDbCommand("COMMENT ON TABLE " & Me.Name & " IS '" & Me.Comments & "'", OpenConnection)
        Try
          MedConnection.Open(OpenConnection)
          createCmd.ExecuteNonQuery()
        Catch ex As Exception
          Console.WriteLine(ex.Message & " on table " & Me.Name)
          failed = True
        Finally
          MedConnection.Close(OpenConnection)
        End Try
      End If
      Return Not failed
    End Function


    Public Function CreateDataTable() As DataTable
      Dim table As New DataTable(Me.Name)
      ' Add all the columns
      Dim field As StructField
      For Each field In Me.Fields
        table.Columns.Add(New DataColumn(field.Name, field.DataType))
      Next
      ' Set up primary keys
      Dim keys As New ArrayList()
      For Each field In Me.PrimaryKeys
        keys.Add(table.Columns(field.Name))
      Next
      If keys.Count > 0 Then
        table.PrimaryKey() = DirectCast(keys.ToArray(GetType(DataColumn)), DataColumn())
      End If
      ' Return the table
      Return table
    End Function

#Region "Collection"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : Structure Parser
    ''' Class	 : Parser.StructTable.TableCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Collection of tables
    ''' </summary>
    ''' <remarks>
    '''   This class is a strongly typed wrapper around HashTable
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(True)> _
Public Class TableCollection
      Inherits Hashtable

      Friend Sub New()

      End Sub

      Friend Shadows Sub Add(ByVal Table As StructTable)
        If Table.Name > "" Then MyBase.Add(Table.Name.ToUpper, Table)
      End Sub

      Default Public Shadows ReadOnly Property Item(ByVal Tablename As String) As StructTable
                Get
                    Dim objReturn As StructTable = Nothing
                    Try
                        objReturn = CType(MyBase.Item(Tablename.ToUpper), StructTable)
                    Catch
                    End Try
                    Return objReturn
                End Get
      End Property

      Public Shadows Function GetEnumerator() As IEnumerator
        Return MyBase.Values.GetEnumerator()
      End Function

    End Class
#End Region

  End Class

#End Region

#Region "Field"
  ''' -----------------------------------------------------------------------------
  ''' Project	 : Structure Parser
  ''' Class	 : Parser.StructField
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Holds name and link information on a field<p/>
  '''   Contains code to parse field definitions to extract this data.
  ''' </summary>
  ''' <remarks>
  ''' </remarks>
  ''' <revisionHistory>
  ''' <revision><Author>[Boughton]</Author><date> [25/11/2005]</date><Action></Action></revision>
  ''' </revisionHistory>
  ''' -----------------------------------------------------------------------------
  <CLSCompliant(True)> _
Public Class StructField

    Private Const C_KeyChar As String = "CHARS"
    Private Const C_KeyPrimary As String = "KEY"
    Private Const C_KeyBrowse As String = "BROWSE"
    Private Const C_KeyPhrase As String = "PHRASE"
    Private Const C_KeyLink As String = "LINK"
    Private Const C_KeyType As String = "TYPE"
    Private Const C_KeyUsage As String = "USAGE"
    Private Const C_Default As String = "DEFAULT"

    Private Shared breakChars As Char() = New Char() {" "c, "("c, ";"c}

    Private _table As StructTable, _name As String, _linkTable As String = "", _linkField As String = ""
    Private _phrase As String = "", _keyIndex As Integer, _browseIndex As Integer = -1, _chars As String
    Private _isOrder As Boolean, _phraseType As PhraseLinkType, _smType As SMFieldDataType
    Private _dataType As OleDb.OleDbType, _dataLen As Integer, _default As String
    Private _trueWord As String, _falseWord As String, _usage As FieldUsage
    Private _charValidator As CharRangeValidator

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a new field with information based on structure definition file
    ''' </summary>
    ''' <param name="parentTable"></param>
    ''' <param name="definition"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub New(ByRef parentTable As StructTable, ByVal definition As String)
            Dim uCaseDef As String
      _table = parentTable
      _name = Parser.GetItemName(definition).ToUpper
      uCaseDef = definition.ToUpper

      If uCaseDef.IndexOf("LINKS_TO") >= 0 AndAlso _
         Me.Name <> "LINKS_TO" Then

        ' Format is links_to table.field
        Dim linkStart As Integer = uCaseDef.IndexOf("LINKS_TO") + 8
        Dim linkEnd As Integer = uCaseDef.IndexOf(".", linkStart)
        _linkTable = uCaseDef.Substring(linkStart, linkEnd - linkStart).Trim
        _linkField = uCaseDef.Substring(linkEnd + 1).Trim
        linkEnd = _linkField.IndexOfAny(breakChars)
        If linkEnd >= 0 Then _linkField = _linkField.Substring(0, linkEnd).Trim

      ElseIf uCaseDef.IndexOf("PHRASE_TYPE") >= 0 AndAlso _
             Me.Name <> "PHRASE_TYPE" Then

        Dim linkStart As Integer = uCaseDef.IndexOf("PHRASE_TYPE") + 11
        Dim linktype As String = uCaseDef.Substring(linkStart + 1).Trim.ToUpper
        Dim linkEnd As Integer = linkType.IndexOfAny(breakChars)
        If linkEnd >= 0 Then linkType = linkType.Substring(0, linkEnd).Trim
        ' Browse on Phrase (store and browse on same field)
        ' Use Phrase_ID if "Identities" set, else Phrase_text
        _linkTable = "PHRASE"
        _phrase = linktype
        If uCaseDef.IndexOf("IDENTITIES") >= 0 Then
          _linkField = "PHRASE_ID"
          _phraseType = PhraseLinkType.Identity
        Else
          _linkField = "PHRASE_TEXT"
          _phraseType = PhraseLinkType.Text
        End If

      ElseIf uCaseDef.IndexOf("CHOOSE_TYPE") >= 0 AndAlso _
             Me.Name <> "CHOOSE_TYPE" Then

        ' Browse on Phrase.Phrase_Text, store Phrase.Phrase_ID
        Dim linkStart As Integer = uCaseDef.IndexOf("CHOOSE_TYPE") + 11
        Dim linktype As String = uCaseDef.Substring(linkStart + 1).Trim
        Dim linkEnd As Integer = linkType.IndexOfAny(breakChars)
        If linkEnd >= 0 Then linkType = linkType.Substring(0, linkEnd).Trim
        _linkTable = "PHRASE"
        _linkField = "PHRASE_ID"
        _phrase = linktype
        _phraseType = PhraseLinkType.Normal

      ElseIf uCaseDef.IndexOf("DATATYPE") >= 0 AndAlso _
             Me.Name <> "DATATYPE" Then
        Dim linkStart As Integer = uCaseDef.IndexOf("DATATYPE") + 8
        Dim linktype As String = uCaseDef.Substring(linkStart + 1).Trim
        Dim linkEnd As Integer = linkType.IndexOfAny(breakChars)
        'If linkEnd < 0 Then linkEnd = linkType.IndexOf(";")
        If linkEnd >= 0 Then linkType = linkType.Substring(0, linkEnd).Trim
        ' Other SM types can be determined from the database type, only check the odd ones
        Select Case linktype
          Case "PACKED_DECIMAL"
            _smType = SMFieldDataType.PackedDec
          Case "INTERVAL"
            _smType = SMFieldDataType.Interval
          Case "BOOLEAN"
            _smType = SMFieldDataType.Boolean
          Case "IDENTITY"
            _smType = SMFieldDataType.Identity
        End Select
      End If

      If uCaseDef.IndexOf("USED_FOR") >= 0 AndAlso _
         Me.Name <> "USED_FOR" Then
        Dim linkStart As Integer = uCaseDef.IndexOf("USED_FOR") + 8
        Dim linktype As String = uCaseDef.Substring(linkStart + 1).Trim
        Dim uses() As String = linktype.Replace("_", "").Split(","c)
        Dim use As String, exitUses As Boolean
        For Each use In uses
          ' Check whether we've reached end of Used_For section
          use = use.Trim
          Dim wordEnd As Integer = use.IndexOf(" ")
          If wordEnd > 0 Then
            Dim bracketStart As Integer = use.IndexOf("(")
            If bracketStart > wordEnd Then
              ' Check if there is any other word between our keyword and the brackets
              Dim midSection As String = use.Substring(wordEnd, bracketStart - wordEnd).Trim
              If midSection = "" Then
                wordEnd = use.IndexOf(")") + 1
                use = use.Substring(0, wordEnd).Replace(" ", "")
              Else
                ' This must be the last identifier
                use = use.Substring(0, wordEnd)
                exitUses = True
              End If
            Else
              ' This must be the last identifier
              use = use.Substring(0, wordEnd)
              exitUses = True
            End If
          End If
          ' Parse use
          If use.IndexOf("BROWSE") > 0 Then
            _usage = _usage Or FieldUsage.Browse
            If use.IndexOf("PRIMARY") >= 0 Then
              _browseIndex = 0
            ElseIf use.IndexOf("(") > 0 Then
              _browseIndex = CInt(ExtractStringSection(use, 13, "(", ")").Trim)
            Else
              _browseIndex = 1
            End If
          Else
            Try
              _usage = _usage Or DirectCast([Enum].Parse(GetType(FieldUsage), use, True), FieldUsage)
            Catch
              Medscreen.LogError("Could not identify Used_For keyword: " & use, False)
            End Try
          End If
          If exitUses Then
            Exit For
          End If
        Next
        ' Check if key field
        If Me.IsPrimaryKey Then
          _keyIndex = Me.Table.KeyCount + 1
        End If
        ' Check if primary display field
        'If linktype.IndexOf("PRIMARY_BROWSE") >= 0 Then
        '  _browseIndex = 0
        'End If
        '' Check if order by field
        'If linktype.IndexOf("ORDERING") >= 0 Then
        '  _isOrder = True
        'End If
        '' Check if display field
        'linkstart = linktype.IndexOf("DISPLAY_BROWSE")
        'If linkstart >= 0 Then
        '  Dim browseIndex As String = linktype.Substring(linkstart + 14).Trim
        '  If browseIndex.StartsWith("(") Then
        '    _browseIndex = CInt(Val(browseIndex.Substring(1).Trim))
        '  Else
        '    _browseIndex = 1
        '  End If
        'End If
      End If

      If uCaseDef.IndexOf("ALLOWED_CHARACTERS") >= 0 AndAlso _
         Me.Name <> "ALLOWED_CHARACTERS" Then
        Dim linkStart As Integer = uCaseDef.IndexOf("'", uCaseDef.IndexOf("ALLOWED_CHARACTERS"))
        ' NOTE: Allowed_Characters may contain lower case chars, so don't use uCaseDef
        If linkstart >= 0 Then
          Dim linktype As String = definition.Substring(linkStart + 1)
          _chars = linktype.Substring(0, linktype.IndexOf("'"))
        End If
      End If

      If uCaseDef.IndexOf(" DEFAULT ") >= 0 AndAlso _
         Me.Name <> " DEFAULT " Then
        Dim linkStart As Integer = uCaseDef.IndexOf(" ", uCaseDef.IndexOf(" DEFAULT ")) + 9
        ' NOTE: Default may contain lower case chars, so don't use uCaseDef
        If linkstart >= 0 Then
          Dim linktype As String = definition.Substring(linkStart)
          Dim linkend As Integer
          If linktype.StartsWith("'") Then
            ' Add 1 because single quote need to be included
            linkend = linktype.IndexOf("'", 1) + 1
          Else
            linkend = linktype.IndexOfany(New Char() {" "c, ";"c})
          End If
          If linkend < 0 Then linkend = linktype.Length
          _default = linktype.Substring(0, linkend).Trim
        End If
      End If

      ' Fill in True and False words for boolean types
      If Me.SMDataType = SMFieldDataType.Boolean Then
        If uCaseDef.IndexOf(" TRUE_WORD") >= 0 Then
          Dim linkStart As Integer = uCaseDef.IndexOf("'", uCaseDef.IndexOf(" TRUE_WORD"))
          ' NOTE: Allowed_Characters may contain lower case chars, so don't use uCaseDef
          If linkstart >= 0 Then
            Dim linktype As String = definition.Substring(linkStart + 1)
            _trueWord = linktype.Substring(0, linktype.IndexOf("'")).Trim
          End If
        End If
        If uCaseDef.IndexOf(" FALSE_WORD") >= 0 Then
          Dim linkStart As Integer = uCaseDef.IndexOf("'", uCaseDef.IndexOf(" FALSE_WORD"))
          ' NOTE: words may contain lower case chars, so don't use uCaseDef
          If linkstart >= 0 Then
            Dim linktype As String = definition.Substring(linkStart + 1)
            _falseWord = linktype.Substring(0, linktype.IndexOf("'")).Trim
          End If
        End If
        ' Store words in allowed characters field, separated by a |
        If _trueWord > "" And _falseWord > "" Then _chars = _trueWord & "|" & _falseWord
      End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates a field based on information in User_Col_Comments data row
    ''' </summary>
    ''' <param name="parentTable"></param>
        ''' <param name="fieldData"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub New(ByVal parentTable As StructTable, ByVal fieldData As DataRow)
      _table = parentTable
      _name = fieldData!Column_Name.ToString
			_dataType = Parser.GetOleDbType(fieldData!Data_type.ToString)
      _dataLen = CInt(fieldData!Data_Length)
      If Not IsBlank(fieldData!Comments) Then
        Dim comments() As String = fieldData!Comments.ToString.Split(";"c)
        Dim comment As String, key As String, value As String
        For Each comment In comments
          If comment.IndexOf("=") > 0 Then
            key = comment.Substring(0, comment.IndexOf("="))
            value = comment.Substring(comment.IndexOf("=") + 1).Trim
            Try
              Select Case key
                Case C_KeyChar
                  _chars = value.Substring(1, value.Length - 2)
                Case C_KeyPrimary
                  _keyIndex = CInt(Val(value))
                Case C_KeyBrowse
                  _browseIndex = CInt(Val(value))
                Case C_KeyPhrase
                  _phrase = value.Substring(0, value.IndexOf("."))
                  _phraseType = CType([Enum].Parse(GetType(PhraseLinkType), _phrase, True), PhraseLinkType)
                  _phrase = value.Substring(value.IndexOf(".") + 1)
                Case C_KeyLink
                  _linkTable = value.Substring(0, value.IndexOf(".")).ToUpper
                  _linkField = value.Substring(value.IndexOf(".") + 1).ToUpper
                Case C_KeyType
                  Try
                    _smType = CType([Enum].Parse(GetType(SMFieldDataType), value, True), SMFieldDataType)
                  Catch
                  End Try
                Case C_KeyUsage
                  _usage = DirectCast([Enum].ToObject(GetType(FieldUsage), CInt(value)), FieldUsage)
                Case C_Default
                  _default = value
              End Select
            Catch
            End Try
          End If
        Next
      End If
      ' Attempt to determine SM DataType if not specified explicitly
      If Me.SMDataType = SMFieldDataType.Unknown Then
        Select Case Me.DBDataType
          Case Data.OleDb.OleDbType.VarChar
            _smType = SMFieldDataType.Text
          Case Data.OleDb.OleDbType.Date
            _smType = SMFieldDataType.Date
          Case Data.OleDb.OleDbType.Integer, Data.OleDb.OleDbType.Decimal
            _smType = SMFieldDataType.Integer
          Case Data.OleDb.OleDbType.Double
            _smType = SMFieldDataType.Real
        End Select
      End If
      ' Fill in True and False words for boolean types
      If Me.SMDataType = SMFieldDataType.Boolean Then
        If (Not Me.AllowedChars Is Nothing) AndAlso _
           Me.AllowedChars.IndexOf("|") > 0 Then
          Dim words As String() = Me.AllowedChars.Split("|"c)
          _trueWord = words(0)
          _falseWord = words(1)
        Else
          _trueWord = C_dbTRUE
          _falseWord = C_dbFALSE
        End If
      End If
      ' Set defaults for certain types where no default already specified
      If Not Me.HasDefaultValue Then
        Select Case Me.SMDataType
          Case SMFieldDataType.Integer, SMFieldDataType.Real, SMFieldDataType.Interval, SMFieldDataType.PackedDec
            _default = "0"
          Case SMFieldDataType.Boolean
            _default = Me.FalseWord
          Case SMFieldDataType.Identity, SMFieldDataType.Text
            ' Keep Sample Manager happpy - can't cope with Nulls in Foreign Keys
            If Me.PhraseLink <> PhraseLinkType.Text And Me.LinkField > "" Then
              _default = " "
            End If
        End Select
      End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   A bitwise value indicating the special uses of this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [02/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property UsedFor() As FieldUsage
      Get
        Return _usage
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the field is used for the specified purpose
    ''' </summary>
    ''' <param name="use"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [02/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IsUsedFor(ByVal use As FieldUsage) As Boolean
      Get
        Return (Me.UsedFor And use) > 0
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the table to which this field belongs
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Table() As StructTable
      Get
        Return _table
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the field name
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
      Get
        Return _name
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a string in the form Table.Field if field is linked to another table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Link() As String
      Get
        If Me.LinkTable > "" Then
          Return Me.LinkTable & "." & Me.LinkField
        Else
          Return ""
        End If
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Field in parent table for which this field is a foreign key
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LinkField() As String
      Get
        Return _linkField
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the table in which LinkField exists
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LinkTable() As String
      Get
        Return _linkTable
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets details of this field and it's comments
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
      Return Me.Identifier & ":" & Me.Comments
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns an identifier in the form TableName.ColumnName
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Identifier() As String
      Get
        Return Me.Table.Name & "." & Me.Name
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the string that is written to User_Col_Comments
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend ReadOnly Property Comments() As String
      Get
        Dim comment As String = ""
        If Me.LinkTable > "" Then
          comment = C_KeyLink & "=" & Me.Link & ";"
        End If
        ' Values of 10 and over cannot be mapped directly to database types, so specify in comments
        ' See comments on the SMFieldDataType enumeration for more details
        If Me.SMDataType >= 10 Then
          comment &= C_KeyType & "=" & Me.SMDataType.ToString & ";"
        End If
        If Me.UsedFor > FieldUsage.None Then
          comment &= C_KeyUsage & "=" & Me.UsedFor.ToString("D") & ";"
        End If
        If Me.IsPrimaryKey Then
          comment &= C_KeyPrimary & "=" & Me.KeyIndex & ";"
        End If
        If Me.IsDisplay Then
          comment &= C_KeyBrowse & "=" & Me.DisplayIndex & ";"
        End If
        If Me.IsPhraseLinked Then
          comment &= C_KeyPhrase & "=" & Me.PhraseLink.ToString & "." & Me.PhraseType & ";"
        End If
        If Me.AllowedChars > "" Then
                    comment &= StructField.C_KeyChar & "=""" & Me.AllowedChars & """;"
        End If
        If Me.HasDefaultValue Then
          comment &= C_Default & "=" & _default & ";"
        End If
        Return comment
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the data type specified in structure.txt
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property SMDataType() As SMFieldDataType
      Get
        Return _smType
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the oracle data type used for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DBDataType() As OleDb.OleDbType
      Get
        Return _dataType
      End Get
		End Property


		Public ReadOnly Property DataType() As Type
			Get
				Select Case Me.DBDataType
					Case Data.OleDb.OleDbType.VarChar
						Return GetType(String)
					Case Data.OleDb.OleDbType.Double
						Return GetType(Double)
					Case Data.OleDb.OleDbType.Decimal
						Return GetType(Decimal)
					Case Data.OleDb.OleDbType.Date
						Return GetType(DateTime)
					Case Data.OleDb.OleDbType.LongVarChar
                        Return GetType(String)
                    Case Else
                        Return GetType(String)                  'This is the safest return 
                End Select
			End Get
		End Property


		Public ReadOnly Property DBDataLength() As Integer
			Get
				Return _dataLen
			End Get
		End Property

		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets a valie indicating whether this field is a primary key
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property IsPrimaryKey() As Boolean
			Get
				Return (Me.UsedFor And FieldUsage.UniqueKey) > 0
			End Get
		End Property


		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets the position of this field within the Primary Key fields for this table
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property KeyIndex() As Integer
			Get
				Return _keyIndex
			End Get
		End Property


		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets a value indicating whether this field is the default for ordering
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property IsOrder() As Boolean
			Get
				Return _isOrder
			End Get
		End Property


		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets a value indicating whether this field is normally included on browses
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property IsDisplay() As Boolean
			Get
				Return (Me.UsedFor And FieldUsage.Browse) > 0
			End Get
		End Property


		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets the (column) position for this field when browsing
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property DisplayIndex() As Integer
			Get
				Return _browseIndex
			End Get
		End Property


		''' -----------------------------------------------------------------------------
		''' <summary>
		'''   Gets a value indicating whether this field links to a phrase
		''' </summary>
		''' <value></value>
		''' <remarks>
		''' </remarks>
		''' <revisionHistory>
		''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
		''' </revisionHistory>
		''' -----------------------------------------------------------------------------
		Public ReadOnly Property IsPhraseLinked() As Boolean
			Get
				Return Me.PhraseLink <> PhraseLinkType.None
			End Get
    End Property

    Public ReadOnly Property IsPhraseIDLinked() As Boolean
      Get
        Return Me.PhraseLink = PhraseLinkType.Normal Or Me.PhraseLink = PhraseLinkType.Identity
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a value identifying how this field links to a phrase
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Fields can either contain the Phrase Text or Phrase Identity.
    '''    <UL>
    '''      <LI>Normal - Browse on Text, field contains Identity</LI>
    '''      <LI>Text - Browse on Text, field contains Text</LI>
    '''      <LI>Identity - Browse on Identity, field contains Identity</LI>
    '''   </UL>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseLink() As PhraseLinkType
      Get
        Return _phraseType
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the Phrase this field links to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseType() As String
      Get
        Return _phrase
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether a default value is specified for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HasDefaultValue() As Boolean
      Get
        Return Not _default Is Nothing
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the default value for this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Some data fields have a default of Now + Interval.  The default value will
    '''   be calculated at the time this property is called.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DefaultValue() As Object
      Get
        If _default Is Nothing OrElse _default.ToUpper = "NULL" Then
          Return DBNull.Value
        ElseIf Me.DBDataType = Data.OleDb.OleDbType.Date AndAlso _
         (_default.StartsWith("'") And _default.EndsWith("'")) Then
          ' Default for dates is normally in interval format to be added to current date
          Dim addTime As TimeSpan = Interval.Parse(_default.Trim(New Char() {"'"c})).ToTimeSpan
          Return DateTime.Now.Add(addTime)
        Else
          ' Return default in correct type for this field - remove single quotes
          Return Me.FormatValue(_default.Trim(New Char() {"'"c}))
        End If
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns an outline select clause for browsing on linked table
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property SelectClause() As String
      Get
                Dim sql As String = ""
        If Me.LinkTable <> "" Then
          Try
            'sql = CType(Table.Tables.Item(Me.LinkTable), StructTable).BuildSelect(Me.LinkField, "")
            sql = GetTable(Me.LinkTable).BuildSelect(Me.LinkField, "")
          Catch
          End Try
        End If
        Return sql
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a string representing characters allowed in this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property AllowedChars() As String
      Get
        Return _chars
      End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''    Gets the character validator for Allowed Chars
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CharValidator() As CharRangeValidator
      Get
        If _charValidator Is Nothing Then
          _charValidator = New CharRangeValidator(Me.AllowedChars)
        End If
        Return _charValidator
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the True word for boolean types
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   The True and False words are only used for user data entry.  The values stored
    '''   in the database field will be C_dbTrue or C_dbFalse.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property TrueWord() As String
      Get
        Return _trueWord
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the False word for boolean types
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   The True and False words are only used for user data entry.  The values stored
    '''   in the database field will be C_dbTrue or C_dbFalse.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FalseWord() As String
      Get
        Return _falseWord
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Returns the PhraseCollection for the phrase used by this field
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   The PhraseCollection holds a list of Phrase objects (phrase entries), one
    '''   for each entry in the phrase table under the relevant phrase_type.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseList() As PhraseCollection
      Get
        Dim phrase As PhraseCollection = MedscreenLib.Glossary.Glossary.PhraseList.Item(Me.PhraseType)
        If phrase Is Nothing Then
          phrase = New MedscreenLib.Glossary.PhraseCollection(Me.PhraseType)
          phrase.Load()
          MedscreenLib.Glossary.Glossary.PhraseList.Add(phrase)
        End If
        Return phrase
      End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Creates an entry in User_Col_Comments
    ''' </summary>
    ''' <param name="OpenConnection"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub CreateComment(ByVal OpenConnection As OleDb.OleDbConnection)
      Dim createCmd As New OleDb.OleDbCommand("COMMENT ON COLUMN " & Me.Table.Name & "." & Me.Name & " IS '" & _
           Replace(Me.Comments, "'", "''") & "'", OpenConnection)
      'createCmd.Parameters.Add("Cmt", OleDb.OleDbType.VarChar).Value = Me.Comments
      Try
        MedConnection.Open(OpenConnection)
        createCmd.ExecuteNonQuery()
      Catch ex As Exception
        Console.WriteLine(ex.Message & " on column " & Me.Table.Name & "." & Me.Name)
      Finally
        MedConnection.Close(OpenConnection)
      End Try
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets a data relation object linking this field to it's parent table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetRelation() As DataRelation
      ' Do not return a relation for phrase text linked fields - it's not strong enough
      If Me.LinkTable > "" And (Me.PhraseLink <> PhraseLinkType.Text) Then
        Return New DataRelation(Me.Identifier, Me.LinkTable, Me.Table.Name, New String() {Me.LinkField}, New String() {Me.Name}, False)
      Else
        Return Nothing
      End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed value is valid for this field
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidValue(ByVal value As String) As Boolean
      Dim isValid As Boolean = (value = "" And Not Me.IsPrimaryKey)
      If Me.PhraseLink = PhraseLinkType.Identity Or Me.PhraseLink = PhraseLinkType.Normal Then
        isValid = Not Me.PhraseList.Item(value) Is Nothing
      ElseIf Me.LinkField > "" Then
        isValid = GetTable(Me.LinkTable).Fields(Me.LinkField).IsValidValue(value)
      ElseIf Not isValid Then
        Select Case Me.SMDataType
          Case SMFieldDataType.Identity
            isValid = (value Is Nothing OrElse _
              (value.ToUpper = value And Me.DBDataLength >= value.Length))
          Case SMFieldDataType.Boolean
            isValid = (value = C_dbTRUE Or value = C_dbFALSE) OrElse _
              value.ToUpper = Me.TrueWord.ToCharArray Or value.ToUpper = Me.FalseWord.ToUpper
          Case SMFieldDataType.PackedDec
            isValid = (IsNumeric(value) AndAlso Me.DBDataLength >= value.Length)
          Case SMFieldDataType.Interval
            isValid = Glossary.Interval.IsInterval(value)
          Case SMFieldDataType.Integer, SMFieldDataType.Real
            isValid = IsNumeric(value)
          Case Else
            Select Case Me.DBDataType
              Case Data.OleDb.OleDbType.VarChar
                isValid = Me.DBDataLength >= value.Length AndAlso _
                          Me.IsValidString(value)
              Case Data.OleDb.OleDbType.Date
                isValid = IsDate(value)
            End Select
        End Select
      End If
      Return isValid
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed value is valid for this field
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidValue(ByVal value As Date) As Boolean
      Dim isValid As Boolean
      If Me.LinkField > "" Then
        isValid = GetTable(Me.LinkTable).Fields(Me.LinkField).IsValidValue(value)
      ElseIf Me.SMDataType = SMFieldDataType.Text Then
        isValid = IsValidValue(value.ToString)
      Else
        isValid = (Me.DBDataType = Data.OleDb.OleDbType.Date)
      End If
      Return isValid OrElse (IsDBNull(value) And Not Me.IsPrimaryKey)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed value is valid for this field
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidValue(ByVal value As Integer) As Boolean
      Dim isValid As Boolean
      If Me.LinkField > "" Then
        isValid = GetTable(Me.LinkTable).Fields(Me.LinkField).IsValidValue(value)
      ElseIf Me.SMDataType = SMFieldDataType.Text Then
        isValid = IsValidValue(value.ToString)
      Else
        isValid = (Me.SMDataType = SMFieldDataType.Integer) OrElse _
          (Me.SMDataType = SMFieldDataType.Real) OrElse _
          (Me.DBDataType = Data.OleDb.OleDbType.VarChar And Me.DBDataLength >= value.ToString.Length)
      End If
      Return isValid OrElse (IsDBNull(value) And Not Me.IsPrimaryKey)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed value is valid for this field
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidValue(ByVal value As Char) As Boolean
      Dim isValid As Boolean
      If Me.PhraseLink = PhraseLinkType.Identity Or Me.PhraseLink = PhraseLinkType.Normal Then
        isValid = Not Me.PhraseList.Item(value) Is Nothing
      ElseIf Me.LinkField > "" Then
        isValid = GetTable(Me.LinkTable).Fields(Me.LinkField).IsValidValue(value)
      ElseIf Me.SMDataType = SMFieldDataType.Boolean Then
        isValid = value = C_dbTRUE Or value = C_dbFALSE
      Else
        isValid = Me.DBDataType = Data.OleDb.OleDbType.VarChar AndAlso _
          Me.IsValidString(value.ToString)
      End If
      Return isValid
    End Function


        ' TODO: IsValidValue - use Parse to simplify logic and cover more circumstances
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Indicates whether the passed object is valid for this field
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks>
        '''   This routine attempts to convert the value into one of the types supported
        '''   by this method's overloads and return the value from that method.<p/>
        '''   Failing that it returns True for NULL values in Non Primary Key fields.
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Boughton]</Author><date> [01/12/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function IsValidValue(ByVal value As Object) As Boolean
            Dim isValid As Boolean
            If TypeOf value Is Glossary.Phrase Then
                isValid = Me.PhraseType > "" AndAlso _
                          Me.PhraseList.PhraseType = DirectCast(value, Glossary.Phrase).PhraseType
            ElseIf Me.LinkField > "" And Not Me.IsPhraseIDLinked Then
                isValid = GetTable(Me.LinkTable).Fields(Me.LinkField).IsValidValue(value)
            ElseIf TypeOf value Is String Then
                isValid = Me.IsValidValue(CType(value, String))
            ElseIf TypeOf value Is Integer Then
                isValid = Me.IsValidValue(CType(value, Integer))
            ElseIf TypeOf value Is Date Then
                isValid = Me.IsValidValue(CType(value, Date))
            ElseIf TypeOf value Is Char Then
                isValid = Me.IsValidValue(CType(value, Char))
            ElseIf TypeOf value Is Single Or TypeOf value Is Double Then
                isValid = Me.SMDataType = SMFieldDataType.Real Or Me.SMDataType = SMFieldDataType.Integer
            Else
                isValid = (Not value Is Nothing) AndAlso Me.IsValidValue(value.ToString)
            End If
            Return isValid
        End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Indicates whether the passed string only contains allowed characters
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [08/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValidString(ByVal value As String) As Boolean
      Dim isValid As Boolean = Me.AllowedChars = ""
      If Not isValid Then
        isValid = Me.CharValidator.IsValidValue(value)
      End If
      Return isValid
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Formats the passed value into an object for assigning to field value
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[BOUGHTON]</Author><date> [20/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
        Public Function FormatValue(ByVal value As Object) As Object
            Dim objReturn As Object = Nothing                               'Declare return variable

            If Me.LinkField > "" And Not Me.IsPhraseIDLinked Then
                objReturn = GetTable(Me.LinkTable).Fields(Me.LinkField).FormatValue(value)
            ElseIf Me.IsValidValue(value) Then
                If value Is Nothing Then
                    objReturn = DBNull.Value
                ElseIf TypeOf value Is Phrase Then
                    Dim phraseEntry As Glossary.Phrase = DirectCast(value, Glossary.Phrase)
                    Select Case Me.PhraseLink
                        Case PhraseLinkType.Text
                            value = phraseEntry.PhraseText
                        Case Else
                            value = phraseEntry.PhraseID
                    End Select
                Else
                    Select Case Me.SMDataType
                        Case SMFieldDataType.Boolean
                            If (value.ToString.Length = 1 AndAlso value.ToString.Substring(0, 1) = C_dbTRUE) OrElse _
                               value.ToString.ToUpper = Me.TrueWord.ToUpper Then
                                objReturn = C_dbTRUE
                            Else
                                objReturn = C_dbFALSE
                            End If
                        Case SMFieldDataType.PackedDec
                            objReturn = value.ToString.PadLeft(Me.DBDataLength)
                        Case Else
                            objReturn = value
                    End Select
                End If
            End If
            Return objReturn
        End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets the value for this field in the passed data row
    ''' </summary>
    ''' <param name="row">Data row to populate</param>
    ''' <param name="value">Value for field</param>
    ''' <remarks>
    '''   Value is automatically formatted for the field, so for example a boolean value
    '''   will be converted to 'T' or 'F' and Packed Decimal fields will be padded correctly.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [24/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetField(ByVal row As DataRow, ByVal value As Object)
      If row Is Nothing Then
        Throw New ArgumentNullException("row")
      End If
      row.Item(Me.Name) = Me.FormatValue(value)
    End Sub

#Region "Collection"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : Structure Parser
    ''' Class	 : Parser.StructTable.FieldCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Collection of fields
    ''' </summary>
    ''' <remarks>
    '''   This class is a strongly typed wrapper around IndexedCollection
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(True)> _
Public Class FieldCollection
      Inherits IndexedCollection

      Friend Sub New()
        MyBase.New("Name")
      End Sub

      Friend Shadows Sub Add(ByVal field As StructField)
        If field.Name > "" Then MyBase.Add(field)
      End Sub


      Friend Sub AddRange(ByVal fields As IEnumerable)
        Dim field As StructField
        For Each field In fields
          Me.Add(field)
        Next
      End Sub


      Default Public Shadows ReadOnly Property Item(ByVal Fieldname As String) As StructField
                Get
                    Dim objReturn As StructField = Nothing
                    Try
                        objReturn = CType(MyBase.Item(Fieldname.ToUpper), StructField)
                    Catch
                    End Try
                    Return objReturn
                End Get
      End Property

    End Class
#End Region

  End Class

#End Region

#Region "Sequences"
  Friend Class Sequences

    Private Shared _sequences As Collections.Specialized.StringDictionary

    Shared Sub New()
      _sequences = New Collections.Specialized.StringDictionary()
      Dim daFields As New OleDb.OleDbCommand("SELECT Sequence_name FROM User_Sequences WHERE Sequence_Name LIKE 'SEQ%'", _
                                             MedConnection.Connection)
      MedConnection.Open()
      Dim getNames As OleDb.OleDbDataReader = daFields.ExecuteReader
      Dim sequence As String, tablename As String
      While getNames.Read
        sequence = getNames.GetString(0)
        tablename = sequence.Substring(4)
        If tablename.EndsWith("_ID") Then tablename = tablename.Substring(0, tablename.Length - 3)
        _sequences.Add(tablename.Replace("_", ""), sequence)
      End While
      getNames.Close()
      MedConnection.Close()
    End Sub

    Shared Function GetTableSequence(ByVal tableName As String) As String
      Return _sequences.Item(tableName.Replace("_", ""))
    End Function

  End Class
#End Region

End Class
