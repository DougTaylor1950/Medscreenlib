'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-09-09 13:42:56+01 $
'$Log: frmLogin.vb,v $
'Revision 1.0  2005-09-09 13:42:56+01  taylor
'Comments added
'
'<>Option Strict On

Imports System.Windows.Forms

'''<summary>Form to provide login to programs requiring login</summary>
Public Class frmLogin
  Inherits System.Windows.Forms.Form

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
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents btnLogin As System.Windows.Forms.Button
  Friend WithEvents btnCancel As System.Windows.Forms.Button
  Friend WithEvents iml16 As System.Windows.Forms.ImageList
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents btnMoreLess As System.Windows.Forms.Button
  Friend WithEvents btnSynchronise As System.Windows.Forms.Button
  Friend WithEvents txtPassword As System.Windows.Forms.TextBox
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnLogin = New System.Windows.Forms.Button
        Me.iml16 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnMoreLess = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnSynchronise = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(20, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Username"
        '
        'txtUser
        '
        Me.txtUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUser.Location = New System.Drawing.Point(88, 12)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(128, 21)
        Me.txtUser.TabIndex = 2
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(88, 40)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(108)
        Me.txtPassword.Size = New System.Drawing.Size(128, 20)
        Me.txtPassword.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(20, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "&Password"
        '
        'btnLogin
        '
        Me.btnLogin.BackColor = System.Drawing.SystemColors.Window
        Me.btnLogin.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnLogin.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnLogin.ImageIndex = 0
        Me.btnLogin.ImageList = Me.iml16
        Me.btnLogin.Location = New System.Drawing.Point(12, 80)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(60, 28)
        Me.btnLogin.TabIndex = 4
        Me.btnLogin.Text = "&Login"
        Me.btnLogin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnLogin.UseVisualStyleBackColor = False
        '
        'iml16
        '
        Me.iml16.ImageStream = CType(resources.GetObject("iml16.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.iml16.TransparentColor = System.Drawing.Color.Transparent
        Me.iml16.Images.SetKeyName(0, "")
        Me.iml16.Images.SetKeyName(1, "")
        Me.iml16.Images.SetKeyName(2, "")
        Me.iml16.Images.SetKeyName(3, "")
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Window
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.ImageIndex = 1
        Me.btnCancel.ImageList = Me.iml16
        Me.btnCancel.Location = New System.Drawing.Point(88, 80)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(68, 28)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnMoreLess
        '
        Me.btnMoreLess.BackColor = System.Drawing.SystemColors.Window
        Me.btnMoreLess.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnMoreLess.ImageIndex = 2
        Me.btnMoreLess.ImageList = Me.iml16
        Me.btnMoreLess.Location = New System.Drawing.Point(172, 80)
        Me.btnMoreLess.Name = "btnMoreLess"
        Me.btnMoreLess.Size = New System.Drawing.Size(56, 28)
        Me.btnMoreLess.TabIndex = 6
        Me.btnMoreLess.Text = "&More"
        Me.btnMoreLess.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnMoreLess.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(212, 28)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Use either your Network username, Sample Manager ID, or User Number."
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(214, 28)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Password is your Network password that you use to log on to Terminal Server"
        '
        'btnSynchronise
        '
        Me.btnSynchronise.BackColor = System.Drawing.SystemColors.Window
        Me.btnSynchronise.Location = New System.Drawing.Point(36, 188)
        Me.btnSynchronise.Name = "btnSynchronise"
        Me.btnSynchronise.Size = New System.Drawing.Size(160, 28)
        Me.btnSynchronise.TabIndex = 9
        Me.btnSynchronise.Text = "Synchronise my Passwords"
        Me.btnSynchronise.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSynchronise.UseVisualStyleBackColor = False
        '
        'frmLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(242, 134)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnSynchronise)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnMoreLess)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximumSize = New System.Drawing.Size(248, 248)
        Me.MinimumSize = New System.Drawing.Size(248, 140)
        Me.Name = "frmLogin"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Medscreen User Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

  Private m_User As User

    '''<summary>
    ''' User who is logged in
    ''' </summary>
    Public Property SelectedUser() As User
        Get
            Return m_User
        End Get
        Set(ByVal value As User)
            m_User = value
        End Set
    End Property


  Private Sub btnMoreLess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMoreLess.Click
    If btnMoreLess.ImageIndex = 2 Then
      btnMoreLess.ImageIndex = 3
      btnMoreLess.Text = "&Less"
      Me.Size = Me.MaximumSize
    Else
      btnMoreLess.ImageIndex = 2
      btnMoreLess.Text = "&More"
      Me.Size = Me.MinimumSize
    End If
  End Sub


  Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click, btnCancel.Click
    Me.DialogResult = CType(sender, Button).DialogResult
    If sender Is btnLogin Then
      If Not m_User Is Nothing Then
        btnLogin.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Dim validated As Boolean = m_User.Validate(txtPassword.Text)
        Me.btnLogin.Enabled = True
        If Validated Then
          Me.Hide()
        Else
          MessageBox.Show("Validation Failed" & ControlChars.NewLine & _
                          m_User.LastMessage, "Login User", _
                          MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
          Me.Cursor = Cursors.Default
        End If
      Else
        MessageBox.Show("User Unknown", "User Login", MessageBoxButtons.OK, MessageBoxIcon.Error)
      End If
    Else
      Me.Hide()
    End If
  End Sub


  Private Sub btnSynchronise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSynchronise.Click
    MessageBox.Show("Sorry, this is not implemented yet", "Synchronise Passwords", MessageBoxButtons.OK, MessageBoxIcon.Information)
  End Sub


  Private Sub txtUser_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtUser.KeyPress
    If e.KeyChar = Chr(Keys.Enter) Then
      SendKeys.Send("{TAB}")
    End If
  End Sub


  Private Sub txtPassword_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles txtPassword.KeyPress
    If e.KeyChar = Chr(Keys.Enter) Then
      btnLogin.PerformClick()
    End If
  End Sub


  Private Sub txtUser_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtUser.Validating
    If txtUser.Text.Length = 4 And IsNumeric(txtUser.Text) Then
      ' 4 digit number is likely to be user numericID (CostCode field in Personnel table))
      m_User = User.GetUser(txtUser.Text, User.SearchType.NumericID)
    Else
      ' Try network ID first, then Sample Manager ID
      Dim networkID As String = txtUser.Text
      If networkID.IndexOf("\") < 0 Then
        networkID = System.Security.Principal.WindowsIdentity.GetCurrent.Name
        networkID = networkID.Substring(0, networkID.IndexOf("\") + 1) & txtUser.Text
      End If
      m_User = User.GetUser(networkID, User.SearchType.NetworkID)
      If m_User Is Nothing Then m_User = User.GetUser(txtUser.Text, User.SearchType.Identity)
    End If
  End Sub


  Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
    Me.Size = Me.MinimumSize
  End Sub


  Private Sub Form_VisibleChanged(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.VisibleChanged
    If Me.Visible Then
      txtPassword.Text = ""
      If btnMoreLess.ImageIndex = 3 Then btnMoreLess.PerformClick()
      If txtUser.Text > "" Then txtPassword.Focus() Else txtUser.Focus()
      Me.Cursor = Cursors.Default
    End If
  End Sub


End Class
