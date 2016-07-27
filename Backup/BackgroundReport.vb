''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : BackgroundReport
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Deal with background reports
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Taylor]	18/02/2009	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class BackgroundReport
    Inherits SMTable
#Region "Declarations"
    Private objERID As IntegerField = New IntegerField("ERID", 0, True)
    Private objCommandString As StringField = New StringField("COMMANDSTRING", "", 1500)
    Private objStatus As StringField = New StringField("STATUS", "", 1)
    Private objPrority As StringField = New StringField("PRORITY", "", 10)
    Private objRunafter As DateField = New DateField("RUNAFTER")
    Private objCreatedOn As DateField = New DateField("CREATED_ON")
    Private objCreatedBy As UserField = New UserField("CREATED_BY")
    Private objModifiedOn As TimeStampField = New TimeStampField("MODIFIED_ON")
    Private objModifiedBy As UserField = New UserField("MODIFIED_BY")
    'Added 20-feb-2009
    Private objMessage As StringField = New StringField("MESSAGE", "", 100)
    Private objLog As BooleanField = New BooleanField("LOG", "F")

#End Region

#Region "Public Instance"

#Region "Functions"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new background report
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CreateBackgroundReport() As BackgroundReport
        Dim myBRep As BackgroundReport = New BackgroundReport()
        Dim strID As String = CConnection.NextSequence("SEQ_BACKGROUNDREPORT")
        myBRep.ERID = strID
        myBRep.objCreatedOn.Value = Now
        myBRep.objCreatedBy.Value = Glossary.Glossary.CurrentSMUser.Identity
        myBRep.Status = "W"
        myBRep.Priority = 10
        myBRep.DoUpdate()
        myBRep.Load(CConnection.DbConnection)
        Return myBRep
    End Function

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new Background Report Instance
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("BACKGROUNDREPORT")
        MyBase.Fields.Add(objERID)
        MyBase.Fields.Add(objStatus)
        MyBase.Fields.Add(objPrority)
        MyBase.Fields.Add(objRunafter)
        MyBase.Fields.Add(objCreatedOn)
        MyBase.Fields.Add(objCreatedBy)
        MyBase.Fields.Add(objModifiedOn)
        MyBase.Fields.Add(objModifiedBy)
        MyBase.Fields.Add(Me.objCommandString)
        MyBase.Fields.Add(objMessage)
        MyBase.Fields.Add(objLog)
        MyBase.Fields.Add(Me.objOutputMethod)

    End Sub
#End Region

#Region "Properties"
    Private objOutputMethod As StringField = New StringField("OUTPUTMETHOD", "", 10)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Method by which report should be output
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property OutPutMethod() As String
        Get
            Return objOutputMethod.Value
        End Get
        Set(ByVal value As String)
            objOutputMethod.Value = value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' error or return message
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Message() As String
        Get
            Return objMessage.Value
        End Get
        Set(ByVal value As String)
            objMessage.Value = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the message should be logged
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Log() As Boolean
        Get
            Return objLog.Value
        End Get
        Set(ByVal value As Boolean)
            objLog.Value = value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Report ID
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ERID() As Integer
        Get
            Return Me.objERID.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objERID.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Status of report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Status() As String
        Get
            Return objStatus.Value
        End Get
        Set(ByVal value As String)
            objStatus.Value = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Command to run report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property CommandString() As String
        Get
            Return objCommandString.Value
        End Get
        Set(ByVal value As String)
            objCommandString.Value = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Priority of report
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' default proirity is 10, lower numbers indicate higher priority
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Priority() As Integer
        Get
            Return objPrority.Value
        End Get
        Set(ByVal value As Integer)
            objPrority.Value = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Don't run before this date and time
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Default value of null means can run immediate
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	18/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property RunAfter() As Date
        Get
            Return objRunafter.Value
        End Get
        Set(ByVal value As Date)
            objRunafter.Value = value
        End Set
    End Property
#End Region
#End Region

End Class
