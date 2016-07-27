Namespace Support
    ''' <summary>
    ''' Class to expose settings values to other applications that use these values
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    ''' <author></author>
    Public Class Support

#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

#End Region

#Region "Properties"
        Private Shared myPasswordSoapServiceURL As String = My.Settings.PasswordSoapService
        Public Shared Property PasswordSOAPServiceURL() As String
            Get
                Return myPasswordSoapServiceURL
            End Get
            Set(ByVal value As String)
                myPasswordSoapServiceURL = value
            End Set
        End Property


        Private Shared myCardiffSoapService As String = My.Settings.CardiffSoapService
        Public Shared Property CardiffSoapServiceURL() As String
            Get
                Return myCardiffSoapService
            End Get
            Set(ByVal value As String)
                myCardiffSoapService = value
            End Set
        End Property

#End Region
#End Region

    End Class
End Namespace
