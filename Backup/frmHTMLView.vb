'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-09-02 07:41:45+01 $
'$Log: Class1.vb,v $
'<>

'''<summary>
''' Form to display HTML where an HTML viewer is not available, either using raw HTML 
''' </summary>
<CLSCompliant(True)> _
Public Class frmHTMLView
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents HtmlEditor1 As onlyconnect.HtmlEditor
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHTMLView))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.HtmlEditor1 = New onlyconnect.HtmlEditor()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblMessage, Me.cmdOk, Me.Button2, Me.Button1})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 345)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(552, 64)
        Me.Panel1.TabIndex = 0
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(376, 7)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 7
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Bitmap)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(264, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Print"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(464, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "&Cancel"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 0
        Me.mnuExit.Text = "E&xit"
        '
        'HtmlEditor1
        '
        'Me.HtmlEditor1.DefaultComposeSettings.BackColor = System.Drawing.Color.White
        'Me.HtmlEditor1.DefaultComposeSettings.DefaultFont = New System.Drawing.Font("Arial", 10.0!)
        'Me.HtmlEditor1.DefaultComposeSettings.Enabled = False
        'Me.HtmlEditor1.DefaultComposeSettings.ForeColor = System.Drawing.Color.Black
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        'Me.HtmlEditor1.DocumentEncoding = onlyconnect.EncodingType.WindowsCurrent
        Me.HtmlEditor1.Name = "HtmlEditor1"
        'Me.HtmlEditor1.OpenLinksInNewWindow = True
        'Me.HtmlEditor1.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left
        'Me.HtmlEditor1.SelectionBackColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionBullets = False
        'Me.HtmlEditor1.SelectionFont = Nothing
        'Me.HtmlEditor1.SelectionForeColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionNumbering = False
        Me.HtmlEditor1.Size = New System.Drawing.Size(552, 345)
        Me.HtmlEditor1.TabIndex = 1
        'Me.HtmlEditor1.Text = "HtmlEditor1"
        '
        'lblMessage
        '
        Me.lblMessage.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lblMessage.Location = New System.Drawing.Point(16, 40)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(528, 23)
        Me.lblMessage.TabIndex = 8
        '
        'frmHTMLView
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 409)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.HtmlEditor1, Me.Panel1})
        Me.Menu = Me.MainMenu1
        Me.Name = "frmHTMLView"
        Me.Text = "HTML Viewer"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '''<summary>
    ''' Raw HTML to display
    ''' </summary>
    Public WriteOnly Property Document() As String

        Set(ByVal Value As String)
            Me.HtmlEditor1.LoadDocument(Value)

            'Me.HtmlEditor1.StyleSheet("\\mark\home\medscreen.css")
            Me.HtmlEditor1.Invalidate()

        End Set
    End Property

    Private myDialogMessage As String = ""
    Public Property Dialogmessage() As String
        Get
            Return myDialogMessage
        End Get
        Set(ByVal Value As String)
            myDialogMessage = Value
            Me.lblMessage.Text = Value
        End Set
    End Property

    Private myDialogstyle As MsgBoxStyle = MsgBoxStyle.OKOnly Or MsgBoxStyle.Information
    Public Property DialogStyle() As MsgBoxStyle
        Get
            Return myDialogstyle
        End Get
        Set(ByVal Value As MsgBoxStyle)
            myDialogstyle = Value
        End Set
    End Property

    Private myDialogTitle As String = ""
    Public Property DialogTitle() As String
        Get
            Return myDialogTitle
        End Get
        Set(ByVal Value As String)
            myDialogTitle = Value
        End Set
    End Property

    '''<summary>
    ''' URL to display
    ''' </summary>
    Public WriteOnly Property URL() As String
        Set(ByVal Value As String)
            Me.HtmlEditor1.LoadUrl(Value)
            Me.HtmlEditor1.Invalidate()
        End Set
    End Property

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.HtmlEditor1.Print(True)
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.myDialogMessage.Trim.Length > 0 Then
            Me.DialogResult = MsgBox(Me.myDialogMessage, Me.myDialogstyle, myDialogTitle)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub
End Class
