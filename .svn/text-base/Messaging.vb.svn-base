'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-08-25 07:58:40+01 $
'$Log: Messaging.vb,v $
'Revision 1.0  2005-08-25 07:58:40+01  taylor
'Base check in
'
'<>
Namespace Messaging

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Messaging.Messaging
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deals with the messaging table
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(True)> _
Public Class Messaging
        Inherits MedscreenLib.SMTable

#Region "Declarations"
        Private objMessageNumber As StringField = New StringField("MESSAGE_NUMBER", "", 10, True) '    
        Private objOperatorID As UserField = New UserField("FROM_OPERATOR_ID")
        Private objRead As BooleanField = New BooleanField("READ", "F")
        Private objMessageText As StringField = New StringField("MESSAGE_TEXT", "", 2000)
        Private objSubject As StringField = New StringField("SUBJECT", "", 100)
        Private objExpires As DateField = New DateField("EXPIRES")
        Private objRecipient As UserField = New UserField("OPERATOR_ID")
        Private objSent As DateField = New DateField("DATE_SENT")
        Private objDateRead As DateField = New DateField("DATE_READ")
        Private objType As StringField = New StringField("CATEGORY", "", 10)


#End Region

#Region "Public Constants"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Message Type Text
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const GCST_MessageTypeText As String = "TEXT"
        '''<summary>HTML message</summary>
        Public Const GCST_MessageTypeHTML As String = "HTML"
        '''<summary>System Message</summary>
        Public Const GCST_MessageTypeSystem As String = "SYS"
        '''<summary>Attachment</summary>
        Public Const GCST_MessageTypeAttachment As String = "ATCH"
#End Region

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' update or insert message
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Update() As Boolean
            If Not MyBase.Fields.Loaded Then   'inserting
                Me.Sent = Now
                If Me.MessageType Is Nothing Then
                    Me.MessageType = "TEXT"
                End If
            End If
            Return MyBase.DoUpdate
        End Function


#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor Creates a new messaging entity
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("MESSAGING")
            MyBase.Fields.Add(Me.objMessageNumber)
            MyBase.Fields.Add(Me.objOperatorID)
            MyBase.Fields.Add(Me.objRead)
            MyBase.Fields.Add(Me.objMessageText)
            MyBase.Fields.Add(Me.objExpires)
            MyBase.Fields.Add(Me.objRecipient)
            MyBase.Fields.Add(Me.objSent)
            MyBase.Fields.Add(Me.objType)
        End Sub

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The Id of the Message
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property MessageId() As Integer
            Get
                Return Me.objMessageNumber.Value
            End Get
            Set(ByVal Value As Integer)
                objMessageNumber.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Text of the message
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property MessageText() As String
            Get
                Return Me.objMessageText.Value
            End Get
            Set(ByVal Value As String)
                objMessageText.Value = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' user creating the message 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property [Operator]() As String
            Get
                Return Me.objOperatorID.Value
            End Get
            Set(ByVal Value As String)
                objOperatorID.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Person receiving the message
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Recipient() As String
            Get
                Return Me.objRecipient.Value
            End Get
            Set(ByVal Value As String)
                objRecipient.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Indicates the message has been read
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Read() As Boolean
            Get
                Return Me.objRead.Value
            End Get
            Set(ByVal Value As Boolean)
                Me.objRead.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Date message expires
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Expires() As Date
            Get
                Return Me.objExpires.Value
            End Get
            Set(ByVal Value As Date)
                Me.objExpires.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Date Message sent
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Sent() As Date
            Get
                Return Me.objSent.Value
            End Get
            Set(ByVal Value As Date)
                Me.objSent.Value = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Type of Message
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property MessageType() As String
            Get
                Return Me.objType.Value
            End Get
            Set(ByVal Value As String)
                Me.objType.Value = Value
            End Set
        End Property


#End Region
#End Region

        Friend Property DataFields() As TableFields
            Get
                Return MyBase.Fields
            End Get
            Set(ByVal Value As TableFields)
                MyBase.Fields = Value
            End Set
        End Property


    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Messaging.MessageList
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A list of messages
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(True)> _
Public Class MessageList
        Inherits MedscreenLib.SMTableCollection

#Region "Declarations"
        Private objMessageNumber As StringField = New StringField("MESSAGE_NUMBER", "", 10, True) '    
        Private objOperatorID As UserField = New UserField("FROM_OPERATOR_ID")
        Private objRead As BooleanField = New BooleanField("READ", "F")
        Private objMessageText As StringField = New StringField("MESSAGE_TEXT", "", 2000)
        Private objSubject As StringField = New StringField("SUBJECT", "", 100)
        Private objExpires As DateField = New DateField("EXPIRES")
        Private objRecipient As UserField = New UserField("OPERATOR_ID")
        Private objSent As DateField = New DateField("DATE_SENT")
        Private objDateRead As DateField = New DateField("DATE_READ")
        Private objType As StringField = New StringField("CATEGORY", "", 10)

#End Region

#Region "Public Instance"

#Region "Functions"


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a Message
        ''' </summary>
        ''' <param name="Message">Message to send</param>
        ''' <param name="recipient">Recipient</param>
        ''' <param name="Originator">Originator</param>
        ''' <param name="DaysBeforeExpires">Number of days before expiry, default = 2</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function CreateMessage(ByVal Message As String, _
      ByVal recipient As String, ByVal Originator As String, Optional ByVal DaysBeforeExpires As Integer = 2) As Messaging
            Dim objMessage As Messaging = New Messaging()
            objMessage.MessageId = NewMessageId()
            objMessage.MessageText = Message
            objMessage.Expires = Today.AddDays(DaysBeforeExpires)
            objMessage.[Operator] = Originator
            objMessage.Recipient = recipient
            objMessage.Update()
            Return objMessage
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send a message to a Role
        ''' </summary>
        ''' <param name="Role">Role to send message to</param>
        ''' <param name="Message">Message to send</param>
        ''' <param name="Originator">Originator of the message</param>
        ''' <param name="MaxRecipients">Maximum no of recipients</param>
        ''' <param name="DaysBeforeExpires">Days before message expires</param>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function SendRoleMessage(ByVal Role As String, ByVal Message As String, _
        ByVal Originator As String, Optional ByVal MaxRecipients As Integer = 1, _
        Optional ByVal DaysBeforeExpires As Integer = 2) As Boolean

            Dim strQuery As String = "Select Operator_id,nvl(rank,1) from role_assignment where role_id  = '" & _
                Role.Trim.ToUpper & "' order by 2,1"

            Dim oCmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(strQuery, MedscreenLib.CConnection.DbConnection)

            Dim OutString As String = ""

            Dim ORead As OleDb.OleDbDataReader = Nothing
            Try
                MedscreenLib.CConnection.SetConnOpen()
                ORead = oCmd.ExecuteReader

                Dim IntCount As Integer = MaxRecipients
                While ORead.Read And IntCount > 0
                    Dim strPerson As String = ORead.GetValue(0)
                    OutString += strPerson & ","
                    IntCount -= 1
                End While
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                CConnection.SetConnClosed()
                If Not ORead Is Nothing Then
                    If Not ORead.IsClosed Then ORead.Close()
                End If
            End Try
            'We should no have an array of people in the role
            Return Me.SendUserMessages(OutString, Message, Originator, DaysBeforeExpires)

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send a message to a user
        ''' </summary>
        ''' <param name="Users">User to send message to</param>
        ''' <param name="Message">Message to send</param>
        ''' <param name="Originator">Originator of message</param>
        ''' <param name="DaysBeforeExpires">Days before message expires, default = 2</param>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function SendUserMessages(ByVal Users As String, ByVal Message As String, _
        ByVal Originator As String, _
        Optional ByVal DaysBeforeExpires As Integer = 2) As Boolean
            Dim Recipients As String()
            Recipients = Users.Split(New Char() {",", ";"}, 100)
            Return MyClass.SendUserMessages(Recipients, Message, Originator, DaysBeforeExpires)

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send A message to an array of users
        ''' </summary>
        ''' <param name="Users">An array of Recipients</param>
        ''' <param name="Message">Message to send</param>
        ''' <param name="Originator">Person sending Message</param>
        ''' <param name="DaysBeforeExpires">Number of days before message expires, default = 2</param>
        ''' <returns>TRUE if Succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function SendUserMessages(ByVal Users As String(), ByVal Message As String, _
            ByVal Originator As String, _
            Optional ByVal DaysBeforeExpires As Integer = 2) As Boolean

            Dim blnRet As Boolean = True
            Try
                Dim intLoop As Integer
                For intLoop = 0 To Users.Length - 1
                    Dim strUser As String = Users.GetValue(intLoop)
                    If strUser.Trim.Length Then
                        Me.CreateMessage(Message, strUser, Originator, DaysBeforeExpires)
                    End If
                Next
            Catch
                blnRet = False
            End Try
            Return blnRet
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get all messages for users
        ''' </summary>
        ''' <param name="User">Recipient of messages</param>
        ''' <param name="UnreadOnly">All messages of just unread, default = TRUE</param>
        ''' <returns>TRUE if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function GetUserMessages(ByVal User As String, _
            Optional ByVal UnreadOnly As Boolean = True) As Boolean
            Dim blnReturn As Boolean = True
            Dim strQuery As String = Me.Fields.SelectString & " where recipient = '" & _
                User.Trim.ToString & "' "
            If UnreadOnly Then
                strQuery += " and read = 'F'"
            End If

            Dim oCmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(strQuery, CConnection.DbConnection)

            Dim oRead As OleDb.OleDbDataReader = Nothing

            Try
                CConnection.SetConnOpen()
                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Dim oMessage As New Messaging()
                    oMessage.DataFields.ReadFields(oRead)
                    Me.Add(oMessage)
                End While
                blnReturn = (Me.Count > 0)
            Catch ex As Exception
                blnReturn = False
            Finally
                CConnection.SetConnClosed()
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
            End Try
            Return blnReturn
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a message to the colelction
        ''' </summary>
        ''' <param name="Item">Message to add</param>
        ''' <returns>Position of Added Message in list</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal Item As Messaging) As Integer
            Return MyBase.List.Add(Item)
        End Function

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor for List
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("MESSAGING")
            MyBase.Fields.Add(Me.objMessageNumber)
            MyBase.Fields.Add(Me.objOperatorID)
            MyBase.Fields.Add(Me.objRead)
            MyBase.Fields.Add(Me.objMessageText)
            MyBase.Fields.Add(Me.objExpires)
            MyBase.Fields.Add(Me.objRecipient)
            MyBase.Fields.Add(Me.objSent)
            MyBase.Fields.Add(Me.objType)
        End Sub

#End Region

#Region "Properties"

#End Region
#End Region

        Private Function NewMessageId() As Integer
            Return CInt(MedscreenLib.CConnection.NextSequence("SEQ_MESSAGE"))
        End Function

    End Class
End Namespace
