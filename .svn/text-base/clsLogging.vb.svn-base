'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2006-01-04 17:45:07+00 $
'$Log: clsLogging.vb,v $
'Revision 1.0  2006-01-04 17:45:07+00  taylor
'check in after createing
'
Imports MedscreenLib.CConnection
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : SendLogInfo
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class representing an element in a send log file 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class SendLogInfo
    Inherits SMTable
#Region "Declarations"
    Dim objID As IntegerField = New IntegerField("ID", 0, True)
    Dim objFilename As StringField = New StringField("FILENAME", "", 30)
    Dim objStatus As StringField = New StringField("STATUS", "", 1)
    Dim objSentBY As StringField = New StringField("SENT_BY", "", 7)
    Dim objSentON As DateField = New DateField("SENT_ON")
    Dim objSentFrom As StringField = New StringField("SENT_FROM", "", 20)
    Dim objSentTo As StringField = New StringField("SENT_TO", "", 120)
    Dim objReferenceTO As StringField = New StringField("REFERENCE_TO", "", 20)
    Dim objReferenceID As StringField = New StringField("REFERENCE_ID", "", 20)

#End Region

 


#Region "Public Instance"

#Region "Functions"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert data to XML 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML() As String
        Dim strRet As String = "<SendLog>"
        Dim objTF As TableField
        For Each objTF In MyBase.Fields
            strRet += objTF.ToXMLSchema
        Next
        strRet += "</SendLog>" & vbCrLf
        Return strRet

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update collection 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Update() As Boolean
        Return MyBase.DoUpdate
    End Function
#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new send log item
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("SEND_LOG")
        MyBase.Fields.Add(Me.objID)
        MyBase.Fields.Add(Me.objFilename)
        MyBase.Fields.Add(Me.objStatus)
        MyBase.Fields.Add(Me.objSentBY)
        MyBase.Fields.Add(Me.objSentON)
        MyBase.Fields.Add(Me.objSentFrom)
        MyBase.Fields.Add(Me.objReferenceTO)
        MyBase.Fields.Add(Me.objReferenceID)
        MyBase.Fields.Add(Me.objSentTo)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a send log by ID
    ''' </summary>
    ''' <param name="newId"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	21/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal newId As Integer)
        MyClass.New()
        Me.ID = newId
        MyBase.Load(CConnection.DbConnection)
    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Id of send log item 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ID() As Integer
        Get
            Return Me.objID.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objID.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Filename associated with item
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Filename() As String
        Get
            Return Me.objFilename.Value
        End Get
        Set(ByVal Value As String)
            Me.objFilename.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Status of item
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Status() As String
        Get
            Return Me.objStatus.Value
        End Get
        Set(ByVal Value As String)
            Me.objStatus.Value = Value
        End Set
    End Property

    Friend Property DataFields() As TableFields
        Get
            Return MyBase.Fields
        End Get
        Set(ByVal Value As TableFields)
            MyBase.Fields = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The way in which the item has been sent out from Medscreen
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SendMethod() As String
        Get
            Return Me.objSentBY.Value
        End Get
        Set(ByVal Value As String)
            Me.objSentBY.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Date report was sent 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SentOn() As Date
        Get
            Return Me.objSentON.Value
        End Get
        Set(ByVal Value As Date)
            objSentON.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Program sending the report out 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SendingProgram() As String
        Get
            Return Me.objSentFrom.Value
        End Get
        Set(ByVal Value As String)
            objSentFrom.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address sent to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReceivingAddress() As String
        Get
            Return Me.objSentTo.Value
        End Get
        Set(ByVal Value As String)
            Me.objSentTo.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' What object id has caused this to be sent
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReferenceID() As String
        Get
            Return Me.objReferenceID.Value
        End Get
        Set(ByVal Value As String)
            objReferenceID.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Item type that has been sent, for use when there are multiple types for a refrence id
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReferenceTo() As String
        Get
            Return Me.objReferenceTO.Value
        End Get
        Set(ByVal Value As String)

            objReferenceTO.Value = Value
        End Set
    End Property
#End Region
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : SendLogCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Collection of send log items 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class SendLogCollection
    Inherits SMTableCollection
#Region "Declarations"
    Dim objID As IntegerField = New IntegerField("ID", 0, True)
    Dim objFilename As StringField = New StringField("FILENAME", "", 30)
    Dim objStatus As StringField = New StringField("STATUS", "", 1)
    Dim objSentBY As StringField = New StringField("SENT_BY", "", 7)
    Dim objSentON As DateField = New DateField("SENT_ON")
    Dim objSentFrom As StringField = New StringField("SENT_FROM", "", 20)
    Dim objReferenceTO As StringField = New StringField("REFERENCE_TO", "", 20)
    Dim objReferenceID As StringField = New StringField("REFERENCE_ID", "", 20)
    Dim objSentTo As StringField = New StringField("SENT_TO", "", 120)
    Dim strWhereClause As String = ""
    Dim strReference As String
#End Region




#Region "Public Instance"

#Region "Functions"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a send log item to collection
    ''' </summary>
    ''' <param name="item">item to add</param>
    ''' <returns>Position of added item in the collection</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal item As SendLogInfo) As Integer
        Return MyBase.List.Add(item)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' load collection or part of the data from the database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Load() As Boolean
        Dim blnRet As Boolean = False

        Dim ocmd As New OleDb.OleDbCommand()
        Dim oRead As OleDb.OleDbDataReader = Nothing

        Try
            Dim strQuery As String = MyBase.Fields.SelectString
            If Me.strWhereClause.Trim.Length > 0 Then
                strQuery += " Where " & Me.strWhereClause
            End If
            ocmd.CommandText = strQuery
            ocmd.Connection = CConnection.DbConnection

            If CConnection.ConnOpen Then
                oRead = ocmd.ExecuteReader
                While oRead.Read
                    Dim objSendLogItem As New SendLogInfo()
                    objSendLogItem.DataFields.ReadFields(oRead)
                    Me.Add(objSendLogItem)
                End While
            End If
            blnRet = True
        Catch ex As Exception
            Medscreen.LogError(ex, , "getting send log info")
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            CConnection.SetConnClosed()
        End Try
        Return blnRet
    End Function

    Public Shared Sub TidyForHTML(ByRef strBody As String)
        strBody = Medscreen.ReplaceString(strBody, vbCrLf, "<BR/>")
        strBody = Medscreen.ReplaceString(strBody, Chr(13), "<BR/>")
        strBody = Medscreen.ReplaceString(strBody, Chr(10), "<BR/>")
        strBody = Medscreen.ReplaceString(strBody, "&lt;", "<")
        strBody = Medscreen.ReplaceString(strBody, "&amp;lt;", "<")
        strBody = Medscreen.ReplaceString(strBody, "&gt;", ">")
        strBody = Medscreen.ReplaceString(strBody, "&amp;gt;", ">")
        strBody = Medscreen.ReplaceString(strBody, "&nbsp;", " ")
        strBody = Medscreen.ReplaceString(strBody, "<0", "&gt;0")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert to XML
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML() As String
        Dim strRet As String = "<SendLogs>"
        'Dim objSL As SendLogInfo
        'For Each objSL In Me.List
        '    strRet += objSL.ToXML & vbCrLf
        'Next
        Dim oCmd As New OleDb.OleDbCommand()
        Dim oread As OleDb.OleDbDataReader = Nothing
        Try
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "select lib_collection.GetSendLogRowXML(id)from send_log  " & _
                "where REFERENCE_ID = ? order by id "
            oCmd.Parameters.Add(CConnection.StringParameter("Custid", strReference, 20))
            If CConnection.ConnOpen Then
                oread = oCmd.ExecuteReader
                While oread.Read
                    strRet += oread.GetValue(0) & vbCrLf

                End While
            End If
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            CConnection.SetConnClosed()             'Close connection
            If Not oread Is Nothing Then            'Try and close reader
                If Not oread.IsClosed Then oread.Close()
            End If


        End Try
        'TidyForHTML(strRet)

        strRet += "</SendLogs>"
        Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new send log item 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CreateSendLog() As SendLogInfo
        Dim intID As Integer = CConnection.NextSequence("SEQ_SEND_LOG")
        Dim objSendLog As New SendLogInfo()
        objSendLog.ID = intID
        objSendLog.DataFields.Insert(DbConnection)
        Return objSendLog
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Log an Email to an Officer
    ''' </summary>
    ''' <param name="CMID">Collection ID</param>
    ''' <param name="Recipients">Sent To </param>
    ''' <param name="SendingProgram">Program sending it defaults to "Collection Manager"</param>
    ''' <param name="ReferenceTo">Type of item logged defaults to "Officer Comment"</param>
    ''' <returns>True if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function LogEmailToOfficer(ByVal CMID As String, _
        ByVal Recipients As String, _
    Optional ByVal SendingProgram As String = "Collection Manager", _
    Optional ByVal ReferenceTo As String = "Officer Comment") As Boolean
        Dim objSendLog As SendLogInfo = CreateSendLog()
        With objSendLog
            .ReferenceID = CMID
            .ReferenceTo = ReferenceTo
            .SendingProgram = SendingProgram
            .Status = "C"
            .SentOn = Now
            .SendMethod = "EMAIL"

            .ReceivingAddress = Recipients
            .DoUpdate()
        End With

    End Function

    Public Shared Sub LogEmailToOfficerNC(ByVal OfficerID As String, ByVal Recipients As String)
        Dim objSendInfo As MedscreenLib.SendLogInfo = MedscreenLib.SendLogCollection.CreateSendLog
        With objSendInfo
            .SendingProgram = "CCTOOL"
            .ReferenceID = OfficerID
            .SentOn = Now
            .ReferenceTo = "Coll Officer"
            .ReceivingAddress = Recipients
            .SendMethod = "EMAIL"
            .Status = "C"
            .Update()
        End With

    End Sub

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create a new send log collection (empty)
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("SEND_LOG")
        MyBase.Fields.Add(Me.objID)
        MyBase.Fields.Add(Me.objFilename)
        MyBase.Fields.Add(Me.objStatus)
        MyBase.Fields.Add(Me.objSentBY)
        MyBase.Fields.Add(Me.objSentON)
        MyBase.Fields.Add(Me.objSentFrom)
        MyBase.Fields.Add(Me.objReferenceTO)
        MyBase.Fields.Add(Me.objReferenceID)
        MyBase.Fields.Add(Me.objSentTo)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new send log collection populating it from a reference
    ''' </summary>
    ''' <param name="Reference">Refernce to use</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Reference As String)
        MyClass.New()
        Me.strReference = Reference
        Me.strWhereClause = " REFERENCE_ID = '" & Me.strReference.Trim & "'"
        'Me.Load()
    End Sub
#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a Send Log item from the collection
    ''' </summary>
    ''' <param name="Index">Position of item required</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Default Public Overloads Property Item(ByVal Index As Integer) As SendLogInfo
        Get
            Return CType(MyBase.List.Item(Index), SendLogInfo)
        End Get
        Set(ByVal Value As SendLogInfo)
            MyBase.List.Item(Index) = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' reference item used 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>Will fill the collection using reference
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Reference() As String
        Get
            Return Me.strReference
        End Get
        Set(ByVal Value As String)
            strReference = Value
            If strReference.Trim.Length > 0 Then
                Me.List.Clear()
                Me.strWhereClause = " REFERENCE_ID = '" & Me.strReference.Trim & "'"
                Me.Load()

            End If

        End Set
    End Property
#End Region

#End Region

End Class