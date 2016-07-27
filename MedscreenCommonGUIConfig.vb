Public Class MedscreenCommonGUIConfig
    Private Shared myInstance As MedscreenCommonGUIConfig
    Private Shared blnCreated As Boolean
    Private Shared CollOut As MedscreenLib.AssemblySettings
    Private Shared myNodePLSQL As MedscreenLib.AssemblySettings
    Private Shared myCollectionQuery As MedscreenLib.AssemblySettings
    Private Shared myNodeQuery As MedscreenLib.AssemblySettings
    Private Shared myLineItems As MedscreenLib.AssemblySettings
    Private Shared myTemplates As MedscreenLib.AssemblySettings
    Private Shared mySampleLineItems As MedscreenLib.AssemblySettings
    Private Shared myIcons As MedscreenLib.AssemblySettings
    Private Shared myDiary As MedscreenLib.AssemblySettings
    Private Shared myMessages As MedscreenLib.AssemblySettings
    Private Shared myMisc As MedscreenLib.AssemblySettings


    Private Sub New()
        'Override default constructor
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Singleton constructor 
    ''' </summary>
    ''' <returns>Class instance</returns>
    ''' <remarks>
    ''' Only one instance of a Singleton can be returned
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/03/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetConfiguration() As MedscreenCommonGUIConfig
        If Not blnCreated Then
            myInstance = New MedscreenCommonGUIConfig()
            blnCreated = True
            Return myInstance
        Else
            Return myInstance
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Access to configuration items relating to output to collectors
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/03/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property NodeStyleSheets() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If CollOut Is Nothing Then
                CollOut = New MedscreenLib.AssemblySettings("NodeStyleSheets")
            End If
            Return CollOut
        End Get

    End Property

    ''' <summary>
    ''' Miscellaneous items
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>16-Apr-2012 06:27</created><Author>CONCATENO\taylor</Author></revision>
    ''' </revisionHistory>
    Public Shared ReadOnly Property Misc() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myMisc Is Nothing Then
                myMisc = New MedscreenLib.AssemblySettings("Misc")
            End If
            Return myMisc
        End Get
    End Property


    Public Shared ReadOnly Property Messages() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myMessages Is Nothing Then
                myMessages = New MedscreenLib.AssemblySettings("Messages")
            End If
            Return myMessages
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Info for collection diary
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/05/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Diary() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myDiary Is Nothing Then
                myDiary = New MedscreenLib.AssemblySettings("Diary")
            End If
            Return myDiary
        End Get

    End Property


    Public Shared ReadOnly Property Icons() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myIcons Is Nothing Then
                myIcons = New MedscreenLib.AssemblySettings("Icons")
            End If
        End Get
    End Property

    Public Shared ReadOnly Property LineItems() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myLineItems Is Nothing Then
                myLineItems = New MedscreenLib.AssemblySettings("LineItems")
            End If
            Return myLineItems
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Types of 'routine Sample line items
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TAYLOR]	02/10/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property SampleLineItems() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If mySampleLineItems Is Nothing Then
                mySampleLineItems = New MedscreenLib.AssemblySettings("SampleLineItems")
            End If
            Return mySampleLineItems
        End Get

    End Property

    Public Shared ReadOnly Property NodePLSQL() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myNodePLSQL Is Nothing Then
                myNodePLSQL = New MedscreenLib.AssemblySettings("NodePLSQL")
            End If
            Return myNodePLSQL
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Queries used with collections
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	31/03/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property CollectionQuery() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myCollectionQuery Is Nothing Then
                myCollectionQuery = New MedscreenLib.AssemblySettings("CollectionQuery")
            End If
            Return myCollectionQuery
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Queries for Nodes 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/04/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property NodeQuery() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myNodeQuery Is Nothing Then
                myNodeQuery = New MedscreenLib.AssemblySettings("NodeQuery")
            End If
            Return myNodeQuery
        End Get

    End Property



    Public Shared ReadOnly Property Templates() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myTemplates Is Nothing Then
                myTemplates = New MedscreenLib.AssemblySettings("Templates")
            End If
            Return myTemplates
        End Get

    End Property

End Class
