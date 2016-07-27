Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace Network

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Network.NetworkDrive
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Network interface code
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class NetworkDrive

#Region "API"
        Private Structure structNetResource
            Public iScope As Integer
            Public iType As Integer
            Public iDisplayType As Integer
            Public iUsage As Integer
            Public sLocalName As String
            Public sRemoteName As String
            Public sComment As String
            Public sProvider As String
        End Structure

        <DllImport("mpr.dll")> _
          Private Shared Function WNetAddConnection2A(ByRef pstNetRes As structNetResource, ByVal psPassword As String, ByVal psUsername As String, ByVal piFlags As Integer) As Integer
        End Function
        <DllImport("mpr.dll")> _
        Private Shared Function WNetCancelConnection2A(ByVal psName As String, ByVal piFlags As Integer, ByVal pfForce As Integer) As Integer
        End Function
        <DllImport("mpr.dll")> _
        Private Shared Function WNetConnectionDialog(ByVal phWnd As Integer, ByVal piType As Integer) As Integer
        End Function
        <DllImport("mpr.dll")> _
        Private Shared Function WNetDisconnectDialog(ByVal phWnd As Integer, ByVal piType As Integer) As Integer
        End Function
        <DllImport("mpr.dll")> _
        Private Shared Function WNetRestoreConnectionW(ByVal phWnd As Integer, ByVal psLocalDrive As String) As Integer
        End Function

        Private Const RESOURCETYPE_DISK As Integer = &H1

        'Standard
        Private Const CONNECT_INTERACTIVE As Integer = &H8
        Private Const CONNECT_PROMPT As Integer = &H10
        Private Const CONNECT_UPDATE_PROFILE As Integer = &H1 'IE4+ Private Const CONNECT_REDIRECT As Integer = &H80
        'NT5 only
        Private Const CONNECT_COMMANDLINE As Integer = &H800
        Private Const CONNECT_CMD_SAVECRED As Integer = &H1000

#End Region

#Region "Properties and options"
        Private lf_SaveCredentials As Boolean = False
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option to save credentials are reconnection...
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property SaveCredentials() As Boolean
            Get
                Return (lf_SaveCredentials)
            End Get
            Set(ByVal value As Boolean)
                lf_SaveCredentials = value
            End Set
        End Property

        Private lf_Persistent As Boolean = False
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option to reconnect drive after log off / reboot ...
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Persistent() As Boolean
            Get
                Return (lf_Persistent)
            End Get
            Set(ByVal value As Boolean)
                lf_Persistent = value
            End Set
        End Property
        Private lf_Force As Boolean = False
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option to force connection if drive is already mapped...
        ''' or force disconnection if network path is not responding...
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Force() As Boolean
            Get
                Return (lf_Force)
            End Get
            Set(ByVal value As Boolean)
                lf_Force = value
            End Set
        End Property

        Private ls_PromptForCredentials As Boolean = False
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option to prompt for user credintals when mapping a drive 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property PromptForCredentials() As Boolean
            Get
                Return (ls_PromptForCredentials)
            End Get
            Set(ByVal value As Boolean)
                ls_PromptForCredentials = value
            End Set
        End Property

        Private ls_Drive As String = "s:"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Drive to be used in mapping / unmapping...
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property LocalDrive() As String
            Get
                Return (ls_Drive)
            End Get
            Set(ByVal value As String)
                If value.Length >= 1 Then
                    ls_Drive = value.Substring(0, 1) & ":"
                Else
                    ls_Drive = ""
                End If
            End Set
        End Property

        Private ls_ShareName As String = "\\Computer\C$"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Share address to map drive to.
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ShareName() As String
            Get
                Return (ls_ShareName)
            End Get
            Set(ByVal value As String)
                ls_ShareName = value
            End Set
        End Property
#End Region

#Region "Function mapping"
        '''
        ''' Map network drive
        '''
        Public Sub MapDrive()
            zMapDrive(Nothing, Nothing)
        End Sub
        '''
        ''' Map network drive (using supplied Password) '''
        Public Sub MapDrive(ByVal Password As String)
            zMapDrive(Nothing, Password)
        End Sub
        ''' Map network drive (using supplied Username and Password) '''
        Public Sub MapDrive(ByVal Username As String, ByVal Password As String)
            zMapDrive(Username, Password)
        End Sub
        '''
        ''' Unmap network drive
        '''
        Public Sub UnMapDrive()
            zUnMapDrive(Me.lf_Force)
        End Sub
        '''
        ''' Check / restore persistent network drive 
        '''         
        Public Sub RestoreDrives()
            zRestoreDrive()
        End Sub
        '''
        ''' Display windows dialog for mapping a network drive 
        ''' '''
        Public Sub ShowConnectDialog(ByVal ParentForm As Form)
            zDisplayDialog(ParentForm, 1)
        End Sub
        '''
        ''' Display windows dialog for disconnecting a network drive 
        ''' '''
        Public Sub ShowDisconnectDialog(ByVal ParentForm As Form)
            zDisplayDialog(ParentForm, 2)
        End Sub


#End Region

#Region "Core functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ' Map network drive
        ''' </summary>
        ''' <param name="psUsername"></param>
        ''' <param name="psPassword"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub zMapDrive(ByVal psUsername As String, ByVal psPassword As String)
            'create struct data
            Dim stNetRes As New structNetResource()
            stNetRes.iScope = 2
            stNetRes.iType = RESOURCETYPE_DISK
            stNetRes.iDisplayType = 3
            stNetRes.iUsage = 1
            stNetRes.sRemoteName = ls_ShareName
            stNetRes.sLocalName = ls_Drive
            'prepare params
            Dim iFlags As Integer = 0
            If lf_SaveCredentials Then
                iFlags += CONNECT_CMD_SAVECRED
            End If
            If lf_Persistent Then
                iFlags += CONNECT_UPDATE_PROFILE
            End If
            If ls_PromptForCredentials Then
                iFlags += CONNECT_INTERACTIVE + CONNECT_PROMPT
            End If
            If psUsername = "" Then
                psUsername = Nothing
            End If
            If psPassword = "" Then
                psPassword = Nothing
            End If 'if force, unmap ready for new connection 
            If lf_Force Then
                Try
                    zUnMapDrive(True)
                Catch
                End Try
            End If
            'call and return
            Dim i As Integer = WNetAddConnection2A(stNetRes, psPassword, psUsername, iFlags)
            If i > 0 Then
                'Throw New System.ComponentModel.Win32Exception(i)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''  ' Unmap network drive
        ''' </summary>
        ''' <param name="pfForce"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub zUnMapDrive(ByVal pfForce As Boolean) 'call unmap and return 
            Dim iFlags As Integer = 0
            If lf_Persistent Then
                iFlags += CONNECT_UPDATE_PROFILE
            End If
            Dim i As Integer = WNetCancelConnection2A(ls_Drive, iFlags, Convert.ToInt32(pfForce))
            If i > 0 Then
                Throw New System.ComponentModel.Win32Exception(i)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ' Check / Restore a network drive
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub zRestoreDrive()
            'call restore and return
            Dim i As Integer = WNetRestoreConnectionW(0, Nothing)
            If i > 0 Then
                Throw New System.ComponentModel.Win32Exception(i)
            End If
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' ' Display windows dialog
        ''' </summary>
        ''' <param name="poParentForm"></param>
        ''' <param name="piDialog"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub zDisplayDialog(ByVal poParentForm As Form, ByVal piDialog As Integer)
            Dim i As Integer = -1
            Dim iHandle As Integer = 0
            'get parent handle
            If Not poParentForm Is Nothing Then
                iHandle = poParentForm.Handle.ToInt32()
            End If
            'show dialog
            If piDialog = 1 Then
                i = WNetConnectionDialog(iHandle, RESOURCETYPE_DISK)
            ElseIf piDialog = 2 Then
                i = WNetDisconnectDialog(iHandle, RESOURCETYPE_DISK)
            End If
            If i > 0 Then
                Throw New System.ComponentModel.Win32Exception(i)
            End If
            'set focus on parent form
            poParentForm.BringToFront()
        End Sub
#End Region

    End Class



End Namespace

