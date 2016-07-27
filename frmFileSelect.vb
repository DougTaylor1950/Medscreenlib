Public Class frmFileSelect
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
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents cmdOpen As System.Windows.Forms.Button
    Friend WithEvents lblMove As System.Windows.Forms.Label
    Friend WithEvents lblSubject As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents ckSalutation As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFileSelect))
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.lblMove = New System.Windows.Forms.Label()
        Me.lblSubject = New System.Windows.Forms.Label()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.ckSalutation = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 241)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(496, 32)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(413, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(327, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "File"
        '
        'txtFile
        '
        Me.txtFile.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtFile.Location = New System.Drawing.Point(128, 32)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(232, 20)
        Me.txtFile.TabIndex = 3
        Me.txtFile.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtFile, "Word Document that will be sent.")
        '
        'cmdOpen
        '
        Me.cmdOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOpen.Location = New System.Drawing.Point(368, 32)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.Size = New System.Drawing.Size(75, 32)
        Me.cmdOpen.TabIndex = 4
        Me.cmdOpen.Text = "Open"
        Me.cmdOpen.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMove
        '
        Me.lblMove.Location = New System.Drawing.Point(24, 72)
        Me.lblMove.Name = "lblMove"
        Me.lblMove.Size = New System.Drawing.Size(376, 23)
        Me.lblMove.TabIndex = 5
        Me.lblMove.Text = "The file will be copied to the template directory. "
        '
        'lblSubject
        '
        Me.lblSubject.Location = New System.Drawing.Point(16, 104)
        Me.lblSubject.Name = "lblSubject"
        Me.lblSubject.TabIndex = 6
        Me.lblSubject.Text = "Subject"
        '
        'txtSubject
        '
        Me.txtSubject.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtSubject.Location = New System.Drawing.Point(128, 96)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(232, 20)
        Me.txtSubject.TabIndex = 7
        Me.txtSubject.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSubject, "Subject line for Email ")
        '
        'ckSalutation
        '
        Me.ckSalutation.Location = New System.Drawing.Point(24, 136)
        Me.ckSalutation.Name = "ckSalutation"
        Me.ckSalutation.Size = New System.Drawing.Size(224, 24)
        Me.ckSalutation.TabIndex = 8
        Me.ckSalutation.Text = "Include salutation (Dear ...)"
        Me.ToolTip1.SetToolTip(Me.ckSalutation, "A boookmark with name Officer needs to be put in the document where the salutatio" & _
        "n will go.")
        '
        'frmFileSelect
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(496, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ckSalutation, Me.txtSubject, Me.lblSubject, Me.lblMove, Me.cmdOpen, Me.txtFile, Me.Label1, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFileSelect"
        Me.Text = "Select File"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private StrFilename As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Full filename of document to be sent out 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property FileName() As String
        Get
            Return Me.StrFilename
        End Get
        Set(ByVal Value As String)
            Me.OpenFileDialog1.FileName = Value
            StrFilename = Value
            Me.txtFile.Text = StrFilename
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Email subject line for message.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Subject() As String
        Get
            Return Me.txtSubject.Text
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Include Salutation 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IncludeSalutation() As Boolean
        Get
            Return Me.ckSalutation.Checked
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Open file dialog action 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpen.Click
        If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            StrFilename = Me.OpenFileDialog1.FileName
            Me.txtFile.Text = StrFilename

        End If

    End Sub
End Class
