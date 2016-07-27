Imports System.Drawing
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
        Me.Icon = ImageToIcon(ImageList1.Images(3))

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

    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents HTMLEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHTMLView))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblMessage = New System.Windows.Forms.Label
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.HTMLEditor1 = New System.Windows.Forms.WebBrowser
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lblMessage)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 345)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(552, 64)
        Me.Panel1.TabIndex = 0
        '
        'lblMessage
        '
        Me.lblMessage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMessage.Location = New System.Drawing.Point(16, 40)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(528, 23)
        Me.lblMessage.TabIndex = 8
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(376, 7)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(75, 23)
        Me.cmdOk.TabIndex = 7
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(264, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Print"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(464, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
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
        'HTMLEditor1
        '
        Me.HTMLEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HTMLEditor1.Location = New System.Drawing.Point(0, 0)
        Me.HTMLEditor1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.HTMLEditor1.Name = "HTMLEditor1"
        Me.HTMLEditor1.Size = New System.Drawing.Size(552, 345)
        Me.HTMLEditor1.TabIndex = 1
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "AcroRd320.bmp")
        Me.ImageList1.Images.SetKeyName(1, "_AcroRd320.bmp")
        Me.ImageList1.Images.SetKeyName(2, "firefox0.bmp")
        Me.ImageList1.Images.SetKeyName(3, "iexplore1.bmp")
        '
        'frmHTMLView
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(552, 409)
        Me.Controls.Add(Me.HTMLEditor1)
        Me.Controls.Add(Me.Panel1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmHTMLView"
        Me.Text = "HTML Viewer"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Enum ViewerFormModes
        HTML
        PDF
    End Enum

    Private myFormMode As ViewerFormModes = ViewerFormModes.HTML
    Public Property FormMode() As ViewerFormModes
        Get
            Return myFormMode
        End Get
        Set(ByVal value As ViewerFormModes)
            myFormMode = value
            If myFormMode = ViewerFormModes.HTML Then
                Me.Icon = ImageToIcon(ImageList1.Images(3))
            Else
                Me.Icon = ImageToIcon(ImageList1.Images(0))
            End If
        End Set
    End Property

    Private Function ImageToIcon(ByVal SourceImage As Bitmap) As Icon
        Dim img As Bitmap = SourceImage
        Dim hIcon As IntPtr = img.GetHicon()
        Dim icn As Icon = Icon.FromHandle(hIcon)
        Return icn
    End Function


    Public Sub NavigateTo(ByVal url As String)
        Me.HTMLEditor1.Navigate(url)
    End Sub


    '''<summary>
    ''' Raw HTML to display
    ''' </summary>
    ''' 


    Public WriteOnly Property Document() As String

        Set(ByVal Value As String)
            Me.HTMLEditor1.DocumentText = Value

            'Me.HtmlEditor1.StyleSheet("\\mark\home\medscreen.css")
            Me.HTMLEditor1.Invalidate()

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
    Public WriteOnly Property URL() As System.Uri
        Set(ByVal Value As System.Uri)
            Me.HTMLEditor1.Url = Value
            Me.HTMLEditor1.Invalidate()
            Me.Dialogmessage = "PDF"
        End Set
    End Property

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Me.HtmlEditor1.Print(True)
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
