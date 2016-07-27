Public Class frmTeamRegion
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
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbRegions As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents udPriority As System.Windows.Forms.NumericUpDown
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTeamRegion))
        Me.pnlControls = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbRegions = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.udPriority = New System.Windows.Forms.NumericUpDown()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.pnlControls.SuspendLayout()
        CType(Me.udPriority, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlControls
        '
        Me.pnlControls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlControls.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 301)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(376, 48)
        Me.pnlControls.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(288, 11)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(176, 11)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Region"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbRegions
        '
        Me.cbRegions.Location = New System.Drawing.Point(152, 32)
        Me.cbRegions.Name = "cbRegions"
        Me.cbRegions.Size = New System.Drawing.Size(200, 21)
        Me.cbRegions.TabIndex = 2
        Me.cbRegions.Text = "ComboBox1"
        Me.ToolTip1.SetToolTip(Me.cbRegions, "List of Regions")
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Priority"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udPriority
        '
        Me.udPriority.Location = New System.Drawing.Point(152, 72)
        Me.udPriority.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.udPriority.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.udPriority.Name = "udPriority"
        Me.udPriority.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.udPriority, "A region of priority 1 means this team will ")
        Me.udPriority.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 50
        Me.ToolTip1.ShowAlways = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(56, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(272, 128)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "A region of priority 1 means this team will have preference for bookings in this " & _
        "region, as the number rises the preference goes down"
        '
        'frmTeamRegion
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(376, 349)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label3, Me.udPriority, Me.Label2, Me.cbRegions, Me.Label1, Me.pnlControls})
        Me.Name = "frmTeamRegion"
        Me.Text = "Region for Team"
        Me.pnlControls.ResumeLayout(False)
        CType(Me.udPriority, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private regionList As MedscreenLib.Glossary.PhraseCollection
    Private myCountries As String

    Public Property Countries() As String
        Get
            Return Me.myCountries
        End Get
        Set(ByVal Value As String)
            myCountries = Value
            Dim CountryList As String() = myCountries.Split(New Char() {","})
            Dim strQuery As String = "select identity,description from port p, country c where c.country_id = p.country_id and p.type in ('C','R') and p.removeflag = 'F' and c.country_code in ("
            Dim strCountry As String
            For Each strCountry In CountryList
                strQuery += "'" & strCountry.Trim & "',"
            Next
            strQuery = Mid(strQuery, 1, strQuery.Length - 1) & ") order by description"
            regionList = New MedscreenLib.Glossary.PhraseCollection(strQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
            cbRegions.DataSource = regionList
            cbRegions.DisplayMember = "PhraseText"
            cbRegions.ValueMember = "PhraseID"
        End Set
    End Property

    Public Property UsedRegion() As String
        Get
            Return cbRegions.SelectedValue
        End Get
        Set(ByVal Value As String)
            cbRegions.SelectedValue = Value
        End Set
    End Property

    Public Property Priority() As Integer
        Get
            Return Me.udPriority.Value
        End Get
        Set(ByVal Value As Integer)
            Me.udPriority.Value = Value
        End Set
    End Property
End Class
