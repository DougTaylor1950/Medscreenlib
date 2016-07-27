'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-09-02 07:41:45+01 $
'$Log: Class1.vb,v $
'<>


Imports System.DirectoryServices
Imports System
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Security.Permissions
Imports Microsoft.VisualBasic
'Imports activeds2
'''<summary>
''' Form allowing users to change their passwords
''' User is initialised to current user on form load
''' </summary>
<CLSCompliant(True)> _
Public Class frmPassword
    Inherits System.Windows.Forms.Form

    Public Enum FormModes
        ChangePassword
        CheckPassword
        CheckPasswordWithSOAP
    End Enum


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtConfirm As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPassword))
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.txtNewPassword = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtConfirm = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "User Name"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(176, 16)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(100, 20)
        Me.txtUser.TabIndex = 1
        '
        'lblPassword
        '
        Me.lblPassword.Location = New System.Drawing.Point(40, 48)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(100, 23)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(176, 48)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(100, 20)
        Me.txtPassword.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 141)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(288, 48)
        Me.Panel1.TabIndex = 4
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(200, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(104, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Ok"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 69)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(288, 72)
        Me.Panel2.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.Location = New System.Drawing.Point(88, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(184, 40)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Enter your Network login  password, remember that these passwords are case sensit" & _
            "ive"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(64, 56)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'txtNewPassword
        '
        Me.txtNewPassword.Location = New System.Drawing.Point(176, 72)
        Me.txtNewPassword.Name = "txtNewPassword"
        Me.txtNewPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNewPassword.Size = New System.Drawing.Size(100, 20)
        Me.txtNewPassword.TabIndex = 2
        Me.txtNewPassword.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(40, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "New Password"
        Me.Label3.Visible = False
        '
        'txtConfirm
        '
        Me.txtConfirm.Location = New System.Drawing.Point(176, 98)
        Me.txtConfirm.Name = "txtConfirm"
        Me.txtConfirm.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirm.Size = New System.Drawing.Size(100, 20)
        Me.txtConfirm.TabIndex = 3
        Me.txtConfirm.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 98)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Confirm Password"
        Me.Label4.Visible = False
        '
        'frmPassword
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(288, 189)
        Me.Controls.Add(Me.txtConfirm)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtNewPassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmPassword"
        Me.Text = "Password Checking"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region



    Private blnChangePassword As Boolean
    Private myFormMode As FormModes = FormModes.ChangePassword

    Property FormMode() As FormModes
        Get
            Return myFormMode
        End Get
        Set(ByVal value As FormModes)
            myFormMode = value
            If value = FormModes.CheckPassword OrElse value = FormModes.CheckPasswordWithSOAP Then
                Me.txtUser.TabIndex = 1
                Me.txtPassword.TabIndex = 0
            End If
        End Set
    End Property

    '''<summary>
    ''' Password (Read Only)
    ''' </summary>
    Public ReadOnly Property Password() As String
        Get
            Return Me.txtPassword.Text
        End Get
    End Property

    Public ReadOnly Property NewPassword() As String
        Get
            Return Me.txtNewPassword.Text
        End Get
    End Property

    '''<summary>
    ''' User Name (Read Only)
    ''' </summary>
    Public ReadOnly Property UserName() As String
        Get
            Return Me.txtUser.Text
        End Get
    End Property

    '''<summary>
    ''' Allow the user to be changed
    ''' </summary>
    Public WriteOnly Property LockUser() As Boolean
        Set(ByVal Value As Boolean)

            Me.txtUser.Enabled = Not Value

        End Set
    End Property

    '''<summary>
    ''' Allow the user to be changed
    ''' </summary>
    Public Property ChangePassword() As Boolean
        Get
            Return Me.blnChangePassword
        End Get
        Set(ByVal Value As Boolean)
            blnChangePassword = Value
            If Value Then
                Me.Height = 280
            Else
                Me.Height = 216
            End If

            Me.Label3.Visible = Value
            Me.Label4.Visible = Value
            Me.txtConfirm.Visible = Value
            Me.txtNewPassword.Visible = Value
        End Set
    End Property

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Me.txtUser.Text = Medscreen.WindowsUser
        Me.txtPassword.Text = ""

        Me.txtPassword.Focus()

    End Sub

    Private Count As Integer = 0
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not Me.blnChangePassword AndAlso FormMode = FormModes.ChangePassword Then Exit Sub

        Dim ErrR As String = ""
        Dim blnRet As Boolean
        If FormMode = FormModes.CheckPassword Then
            blnRet = Medscreen.CheckPassword(Environment.UserName, Me.txtPassword.Text, ErrR)
            ErrR = "Password doesn't match system password"
        ElseIf FormMode = FormModes.CheckPasswordWithSOAP Then
            Try
                blnRet = Medscreen.CheckPasswordBySoap(Environment.UserName, Me.txtPassword.Text, ErrR)
            Catch ex As MedscreenExceptions.MedscreenException
                Throw New MedscreenExceptions.MedscreenLoginCheckTimeoutException
            End Try
        End If
        If Not blnRet Then

            MsgBox(ErrR)
            Me.DialogResult = Windows.Forms.DialogResult.None
            Count += 1
            If Count > 3 Then Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Exit Sub
        End If
        If Me.txtConfirm.Text <> Me.txtNewPassword.Text Then
            MsgBox("Passwords do not match")
            Me.DialogResult = Windows.Forms.DialogResult.None
            Exit Sub
        End If

        If Me.txtPassword.Text = Me.txtNewPassword.Text Then
            MsgBox("You must change the password")
            Me.DialogResult = Windows.Forms.DialogResult.None
            Exit Sub
        End If
        Dim userLogin As String = System.Security.Principal.WindowsIdentity.GetCurrent.Name
        Dim intPos As Integer = InStr(userLogin, "\")
        If intPos > 0 Then
            userLogin = Mid(userLogin, intPos + 1)
        End If

        Try
            If myFormMode = FormModes.ChangePassword Then
                ActiveDirectoryWrapper.PC.ADWrapper.ResetPassword(userLogin, Me.txtNewPassword.Text)
            End If
         Catch ex As Exception
            Medscreen.LogAction(ex.ToString)
            MsgBox("The password has not been changed, due to an issue with the Active Directory permissions, the reason given is :" & ex.ToString)
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Exit Sub
        Finally
        End Try
    End Sub
End Class

Public Class RunAs_Impersonator
#Region "Private Variables and Enum Constants"
    Private tokenHandle As New IntPtr(0)
    Private dupeTokenHandle As New IntPtr(0)
    Private myimpersonatedUser As WindowsImpersonationContext
#End Region
#Region "Properties"

#End Region
#Region "Public Methods"
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean

    Public Declare Auto Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As IntPtr, _
    ByVal SECURITY_IMPERSONATION_LEVEL As Integer, _
    ByRef DuplicateTokenHandle As IntPtr) As Boolean

    Const SecurityImpersonation As Integer = 2

    ' Test harness.
    ' If you incorporate this code into a DLL, be sure to demand FullTrust.
    <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
    Public Sub ImpersonateStart(ByVal Domain As String, ByVal userName As String, ByVal Password As String)
        Try
            tokenHandle = IntPtr.Zero
            ' Call LogonUser to obtain a handle to an access token.
            Dim returnValue As Boolean = LogonUser(userName, Domain, Password, 2, 0, tokenHandle)

            'check if logon successful
            If returnValue = False Then
                Dim ret As Integer = Marshal.GetLastWin32Error()
                Console.WriteLine("LogonUser failed with error code : {0}", ret)
                Throw New System.ComponentModel.Win32Exception(ret)
                Exit Sub
            End If

            'Logon succeeded
            ' Use the token handle returned by LogonUser.
            Dim dupeTokenHandle As IntPtr
            Dim retVal As Boolean = DuplicateToken(tokenHandle, SecurityImpersonation, dupeTokenHandle)
            If (Not retVal) Then
                CloseHandle(tokenHandle)
                Console.WriteLine("Exception in token duplication.")
                Return
            End If


            Dim newId As New WindowsIdentity(dupeTokenHandle)
            myimpersonatedUser = newId.Impersonate()
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try
        'MsgBox(”running as ” & impersonatedUser.ToString & ” — ” & WindowsIdentity.GetCurrent.Name)
    End Sub

    Public Property ImpersonationContext() As WindowsImpersonationContext
        Get
            Return myimpersonatedUser
        End Get
        Set(ByVal value As WindowsImpersonationContext)
            myimpersonatedUser = value
        End Set
    End Property
    <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
    Public Sub ImpersonateStop()
        ' Stop impersonating the user.
        myimpersonatedUser.Undo()

        ' Free the tokens.
        If Not System.IntPtr.op_Equality(tokenHandle, IntPtr.Zero) Then
            CloseHandle(tokenHandle)
        End If
        'MsgBox(”running as ” & Environment.UserName)
    End Sub
#End Region
#Region "Private Methods"
    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
    ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
    ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
    ByRef phToken As IntPtr) As Boolean

    <DllImport("kernel32.dll")> _
    Public Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, _
    ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], _
    ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function
#End Region
End Class