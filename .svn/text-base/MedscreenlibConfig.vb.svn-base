''' <summary>
''' Medscreenlib local configuration items
''' </summary>
''' <remarks></remarks>
Public Class MedscreenlibConfig
    Private Shared myInstance As MedscreenlibConfig
    Private Shared blnCreated As Boolean = False
    Private Shared myHyperlinks As MedscreenLib.AssemblySettings
    Private Shared myBlat As MedscreenLib.AssemblySettings
    Private Shared myPaths As MedscreenLib.AssemblySettings
    Private Shared myValidation As MedscreenLib.AssemblySettings
    Private Shared myServers As MedscreenLib.AssemblySettings
    Private Shared myConnections As MedscreenLib.AssemblySettings
    Private Shared myCollectorOut As MedscreenLib.AssemblySettings
    Private Shared myMisc As MedscreenLib.AssemblySettings


    ''' <summary>
    ''' Accessor function to ensure its is a singleton
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetConfiguration() As MedscreenlibConfig
        If Not blnCreated Then
            myInstance = New MedscreenlibConfig
        End If
        Return myInstance
    End Function

    ''' <summary>
    ''' List of hyperlinks
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property HyperLinks() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myHyperlinks Is Nothing Then
                myHyperlinks = New MedscreenLib.AssemblySettings("HyperLinks")
            End If
            Return myHyperlinks
        End Get

    End Property

    Public Shared ReadOnly Property BLAT() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myBlat Is Nothing Then
                myBlat = New MedscreenLib.AssemblySettings("BLAT")
            End If
            Return myBlat
        End Get

    End Property


    Public Shared ReadOnly Property Miscellaneous() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myMisc Is Nothing Then
                myMisc = New MedscreenLib.AssemblySettings("Misc")
            End If
            Return myMisc
        End Get

    End Property

    Public Shared ReadOnly Property Paths() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myPaths Is Nothing Then
                myPaths = New MedscreenLib.AssemblySettings("Paths")
            End If
            Return myPaths
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used in validating info
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	06/01/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Validation() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myValidation Is Nothing Then
                myValidation = New MedscreenLib.AssemblySettings("Validation")
            End If
            Return myValidation
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Server info
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/03/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Servers() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myServers Is Nothing Then
                myServers = New MedscreenLib.AssemblySettings("Servers")
            End If
            Return myServers
        End Get
    End Property

    Public Shared ReadOnly Property Connections() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myConnections Is Nothing Then
                myConnections = New MedscreenLib.AssemblySettings("Connectors")
            End If
            Return myConnections
        End Get
    End Property

    Public Shared ReadOnly Property CollectorOutput() As MedscreenLib.AssemblySettings
        Get
            GetConfiguration()
            If myCollectorOut Is Nothing Then
                myCollectorOut = New MedscreenLib.AssemblySettings("CollectorOut")
            End If
            Return myCollectorOut
        End Get
    End Property


End Class
