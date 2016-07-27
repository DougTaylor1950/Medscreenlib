Imports MedscreenLib
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : frmSmidProfile
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to get SMID And Profile
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmSmidProfile
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.cbIndustry.DataSource = MedscreenLib.Glossary.Glossary.Industry
        Me.cbIndustry.DisplayMember = "PhraseText"
        Me.cbIndustry.ValueMember = "PhraseID"

        'Me.CBPanel.DataSource = Intranet.intranet.Support.cSupport.TestPanels
        Me.CBPanel.DisplayMember = "TestPanelDescription"
        Me.CBPanel.ValueMember = "TestPanelID"

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
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSMID As System.Windows.Forms.TextBox
    Friend WithEvents NotificationWindow1 As VbPowerPack.NotificationWindow
    Friend WithEvents PanelSMID As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PanelProfile As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtProfile As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents PanelIndustryCode As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents NotificationWindow2 As VbPowerPack.NotificationWindow
    Friend WithEvents PanelFinish As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdPrev As System.Windows.Forms.Button
    Friend WithEvents pnlCompanyName As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents pnlTestPanel As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CBPanel As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblTemplate As System.Windows.Forms.Label
    Friend WithEvents cbTemplates As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSmidProfile))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdPrev = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSMID = New System.Windows.Forms.TextBox()
        Me.NotificationWindow1 = New VbPowerPack.NotificationWindow(Me.components)
        Me.PanelSMID = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PanelProfile = New System.Windows.Forms.Panel()
        Me.txtProfile = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.PanelIndustryCode = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cbIndustry = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.NotificationWindow2 = New VbPowerPack.NotificationWindow(Me.components)
        Me.PanelFinish = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.pnlCompanyName = New System.Windows.Forms.Panel()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.pnlTestPanel = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CBPanel = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblTemplate = New System.Windows.Forms.Label()
        Me.cbTemplates = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.PanelSMID.SuspendLayout()
        Me.PanelProfile.SuspendLayout()
        Me.PanelIndustryCode.SuspendLayout()
        Me.PanelFinish.SuspendLayout()
        Me.pnlCompanyName.SuspendLayout()
        Me.pnlTestPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.cmdPrev, Me.Button2, Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 197)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(480, 40)
        Me.Panel1.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(120, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 24)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "Next"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Button1, "Next command")
        '
        'cmdPrev
        '
        Me.cmdPrev.Image = CType(resources.GetObject("cmdPrev.Image"), System.Drawing.Bitmap)
        Me.cmdPrev.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrev.Location = New System.Drawing.Point(200, 5)
        Me.cmdPrev.Name = "cmdPrev"
        Me.cmdPrev.Size = New System.Drawing.Size(75, 24)
        Me.cmdPrev.TabIndex = 15
        Me.cmdPrev.Text = "Back"
        Me.cmdPrev.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdPrev, "Previous command")
        Me.cmdPrev.Visible = False
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Bitmap)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(16, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "Search"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(388, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(300, 5)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 12
        Me.cmdOk.Text = "&Create"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdOk.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "SMID"
        '
        'txtSMID
        '
        Me.txtSMID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSMID.Location = New System.Drawing.Point(24, 32)
        Me.txtSMID.Name = "txtSMID"
        Me.txtSMID.TabIndex = 0
        Me.txtSMID.Text = ""
        '
        'NotificationWindow1
        '
        Me.NotificationWindow1.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.SystemColors.Desktop, System.Drawing.Color.Aqua)
        Me.NotificationWindow1.DefaultText = Nothing
        Me.NotificationWindow1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NotificationWindow1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NotificationWindow1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NotificationWindow1.ShowStyle = VbPowerPack.NotificationShowStyle.Slide
        '
        'PanelSMID
        '
        Me.PanelSMID.BackColor = System.Drawing.SystemColors.Info
        Me.PanelSMID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PanelSMID.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2})
        Me.PanelSMID.Location = New System.Drawing.Point(8, 56)
        Me.PanelSMID.Name = "PanelSMID"
        Me.PanelSMID.Size = New System.Drawing.Size(456, 128)
        Me.PanelSMID.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 72)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Enter SMID and press next button, you can use the search button to look for an ex" & _
        "isting SMID to use if you suspect that the customer may be part of a group!"
        '
        'PanelProfile
        '
        Me.PanelProfile.BackColor = System.Drawing.SystemColors.Control
        Me.PanelProfile.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtProfile, Me.Label4, Me.Label3})
        Me.PanelProfile.Location = New System.Drawing.Point(8, 56)
        Me.PanelProfile.Name = "PanelProfile"
        Me.PanelProfile.Size = New System.Drawing.Size(456, 128)
        Me.PanelProfile.TabIndex = 8
        '
        'txtProfile
        '
        Me.txtProfile.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProfile.Location = New System.Drawing.Point(144, 16)
        Me.txtProfile.Name = "txtProfile"
        Me.txtProfile.TabIndex = 2
        Me.txtProfile.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Profile"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Info
        Me.Label3.Location = New System.Drawing.Point(32, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(288, 48)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Enter Profile and press next button, we will then check to ensure that this profi" & _
        "le hasn't been used and move you to the next stage."
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'PanelIndustryCode
        '
        Me.PanelIndustryCode.BackColor = System.Drawing.SystemColors.Control
        Me.PanelIndustryCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PanelIndustryCode.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label6, Me.cbIndustry, Me.Label5})
        Me.PanelIndustryCode.Location = New System.Drawing.Point(8, 56)
        Me.PanelIndustryCode.Name = "PanelIndustryCode"
        Me.PanelIndustryCode.Size = New System.Drawing.Size(456, 128)
        Me.PanelIndustryCode.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Info
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Location = New System.Drawing.Point(16, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(336, 80)
        Me.Label6.TabIndex = 40
        Me.Label6.Text = "Select the industry code for the customer, this will be used to provide defaults " & _
        "for the customer."
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbIndustry
        '
        Me.cbIndustry.Location = New System.Drawing.Point(72, 16)
        Me.cbIndustry.Name = "cbIndustry"
        Me.cbIndustry.Size = New System.Drawing.Size(280, 21)
        Me.cbIndustry.TabIndex = 39
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Industry"
        '
        'NotificationWindow2
        '
        Me.NotificationWindow2.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.Color.Firebrick, System.Drawing.Color.LightPink)
        Me.NotificationWindow2.DefaultText = Nothing
        Me.NotificationWindow2.DefaultTimeout = 50000
        Me.NotificationWindow2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NotificationWindow2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NotificationWindow2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NotificationWindow2.ShowStyle = VbPowerPack.NotificationShowStyle.Immediately
        '
        'PanelFinish
        '
        Me.PanelFinish.BackColor = System.Drawing.SystemColors.Info
        Me.PanelFinish.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PanelFinish.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label7})
        Me.PanelFinish.Location = New System.Drawing.Point(8, 56)
        Me.PanelFinish.Name = "PanelFinish"
        Me.PanelFinish.Size = New System.Drawing.Size(456, 128)
        Me.PanelFinish.TabIndex = 10
        Me.PanelFinish.Visible = False
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(384, 40)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "We have all the information we need, press the create button to create this custo" & _
        "mer profile. Please ensure that the details such as the address are added to the" & _
        " customer"
        '
        'pnlCompanyName
        '
        Me.pnlCompanyName.BackColor = System.Drawing.SystemColors.Info
        Me.pnlCompanyName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCompanyName.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtCompanyName, Me.Label8})
        Me.pnlCompanyName.Location = New System.Drawing.Point(8, 56)
        Me.pnlCompanyName.Name = "pnlCompanyName"
        Me.pnlCompanyName.Size = New System.Drawing.Size(456, 128)
        Me.pnlCompanyName.TabIndex = 11
        Me.pnlCompanyName.Visible = False
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(24, 48)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(416, 20)
        Me.txtCompanyName.TabIndex = 1
        Me.txtCompanyName.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Company Name"
        '
        'pnlTestPanel
        '
        Me.pnlTestPanel.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTestPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTestPanel.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label9, Me.CBPanel, Me.Label10})
        Me.pnlTestPanel.Location = New System.Drawing.Point(8, 56)
        Me.pnlTestPanel.Name = "pnlTestPanel"
        Me.pnlTestPanel.Size = New System.Drawing.Size(456, 128)
        Me.pnlTestPanel.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Info
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label9.Location = New System.Drawing.Point(16, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(336, 80)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Select the test panel for the customer, this can be altered after creation."
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CBPanel
        '
        Me.CBPanel.Location = New System.Drawing.Point(72, 16)
        Me.CBPanel.Name = "CBPanel"
        Me.CBPanel.Size = New System.Drawing.Size(280, 21)
        Me.CBPanel.TabIndex = 39
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Test Panel"
        '
        'lblTemplate
        '
        Me.lblTemplate.Location = New System.Drawing.Point(168, 0)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.TabIndex = 13
        Me.lblTemplate.Text = "Template"
        '
        'cbTemplates
        '
        Me.cbTemplates.Location = New System.Drawing.Point(176, 32)
        Me.cbTemplates.Name = "cbTemplates"
        Me.cbTemplates.Size = New System.Drawing.Size(288, 21)
        Me.cbTemplates.TabIndex = 14
        '
        'frmSmidProfile
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 237)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbTemplates, Me.lblTemplate, Me.txtSMID, Me.Label1, Me.Panel1, Me.pnlTestPanel, Me.PanelIndustryCode, Me.pnlCompanyName, Me.PanelFinish, Me.PanelSMID, Me.PanelProfile})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSmidProfile"
        Me.Text = "Customer Creation"
        Me.Panel1.ResumeLayout(False)
        Me.PanelSMID.ResumeLayout(False)
        Me.PanelProfile.ResumeLayout(False)
        Me.PanelIndustryCode.ResumeLayout(False)
        Me.PanelFinish.ResumeLayout(False)
        Me.pnlCompanyName.ResumeLayout(False)
        Me.pnlTestPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Enum Stages
        InitialStage = 0
        GetProfile = 1
        GetCompanyName = 2
        GetIndustry = 3
        getPanel = 4
        FinalStage = 5
    End Enum

    Public Enum FormModes
        Create = 0
        Clone = 1
        Template = 2
    End Enum

    Private intStage As Stages = Stages.InitialStage
    Private tmpClient As Intranet.intranet.customerns.Client
    Private blnExistingSMID As Boolean = False
    Private blnExistingProfile As Boolean = False
    Private myMode As FormModes = FormModes.Create
    Private mySmidProfiles As Intranet.intranet.customerns.SMIDProfileCollection

    Private Const StdProfilePanelMessage As String = "Enter Profile and press next button. This SMID is not in use."

    Private Const StdSMIDPresntProfilePanelMessage As String = "Enter Profile and press next button, we will then check to ensure that this profi" & _
        "le hasn't been used (SMID in use) and move you to the next stage."


    Public Event StageChanged()

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose client 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return Me.tmpClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            Me.tmpClient = Value
            If Not Value Is Nothing Then
                Me.txtSMID.Text = Value.SMID
                Me.txtProfile.Text = Value.Profile
            End If
        End Set

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Mode form is in 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormMode() As FormModes
        Get
            Return Me.myMode

        End Get
        Set(ByVal Value As FormModes)
            myMode = Value
            Me.lblTemplate.Hide()
            Me.cbTemplates.Hide()
            If myMode = FormModes.Template Then
                Me.lblTemplate.Show()
                Me.cbTemplates.Show()
                Me.intStage = Stages.GetProfile
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Smid in use 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SMID() As String
        Get
            Return Me.txtSMID.Text
        End Get
        Set(ByVal Value As String)
            Me.txtSMID.Enabled = False
            Me.txtSMID.Text = Value
            If Value.Trim.Length = 0 Then
                Me.txtSMID.Enabled = True
            End If
            If Me.myMode = FormModes.Create Then
                Me.intStage = Stages.InitialStage
                Me.StageHasChanged()
            Else
                Me.intStage = Stages.GetProfile
                Me.StageHasChanged()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in the smid profiles to the form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property SmidProfiles() As Intranet.intranet.customerns.SMIDProfileCollection
        Set(ByVal Value As Intranet.intranet.customerns.SMIDProfileCollection)
            mySmidProfiles = value
            If Not mySmidProfiles Is Nothing Then
                Me.cbTemplates.DataSource = mySmidProfiles
                Me.cbTemplates.DisplayMember = "CompanyName"
                Me.cbTemplates.ValueMember = "CustomerID"
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the selected template from the form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property TemplateID() As String
        Get
            If Me.cbTemplates.SelectedIndex > -1 Then
                Return Me.cbTemplates.SelectedValue
            Else
                Return ""
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Company name 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CompanyProfileName() As String
        Get
            Return Me.txtCompanyName.Text
        End Get
        Set(ByVal Value As String)
            Me.txtCompanyName.Text = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose selected industryCode 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IndustryCode() As String
        Get
            Dim strTemp As String = ""
            If Me.cbIndustry.SelectedIndex > 0 Then
                strTemp = MedscreenLib.Glossary.Glossary.Industry.Item(Me.cbIndustry.SelectedIndex).PhraseID
            End If
            Return strTemp
        End Get
    End Property

    Public ReadOnly Property TestPanel() As String
        Get
            Dim strTemp As String = ""
            'If Me.CBPanel.SelectedIndex >= 0 Then
            '    strTemp = Intranet.intranet.Support.cSupport.TestPanels.Item(Me.CBPanel.SelectedIndex).TestPanelID
            'End If
            Return strTemp
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Profile for customer
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Profile() As String
        Get
            Return Me.txtProfile.Text
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Look for a SMID
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>We can have cancel hit nothing found <para/>
    ''' We can have a deleted smid found, the user then needs to be offered whether to undelete
    ''' If they undelete then we have the SMID Profile can get on to editing return YES
    ''' Or we have a SMID we know need to establish the profile.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim frmCustSelect As New frmCustomerSelection2()
        frmCustSelect.IncludeDeleted = True

        Dim intRet As Windows.Forms.DialogResult = frmCustSelect.ShowDialog
        If intRet = DialogResult.OK Then
            If Not frmCustSelect.Client Is Nothing Then

                tmpClient = frmCustSelect.Client
                Me.txtSMID.Text = tmpClient.SMID            'Set the SMID we have found
                If tmpClient.Removed Then 'Check to see if we want to reinstate 
                    If MsgBox("Do you want to reinstate this client?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        'tmpClient.Removed = False
                        'tmpClient.Update()
                        Me.txtProfile.Text = tmpClient.Profile
                        intStage = Stages.GetProfile
                    Else
                        tmpClient = Nothing
                        intStage = Stages.GetProfile
                    End If
                Else
                    tmpClient = Nothing
                    intStage = Stages.GetProfile
                End If
            Else : Exit Sub
            End If
        Else : Exit Sub
        End If

        StageHasChanged()
    End Sub

    Private Sub hidepanels()
        Me.PanelSMID.Hide()
        Me.PanelProfile.Hide()
        Me.PanelIndustryCode.Hide()
        Me.PanelFinish.Hide()
        Me.pnlCompanyName.Hide()
        Me.pnlTestPanel.Hide()
    End Sub

    Private Function ValidateEntry() As Boolean
        Dim blnRet As Boolean = True

        Me.ErrorProvider1.SetError(Me.txtSMID, "")
        'differing code for each stage
        Select Case intStage
            Case Stages.InitialStage
                If Me.txtSMID.Text.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtSMID, "No SMID entered")
                    blnRet = False
                End If
                If Me.txtSMID.Text.Trim.Length > 10 Then
                    Me.ErrorProvider1.SetError(Me.txtSMID, "SMID is too long")
                    blnRet = False
                End If
                'Code added to validate the SMID as being a valid identity 
                Dim strCheck As String = CConnection.PackageStringList("Lib_Utils.IsValidGenericIdentityChar", txtSMID.Text)
                If strCheck = "F" Then
                    Me.ErrorProvider1.SetError(Me.txtSMID, "SMID has invalid characters can be only [0-9][A-Z]+-_, with no spaces")
                    blnRet = False
                End If
                Me.intStage += 1
            Case Stages.GetProfile
                If Me.txtProfile.Text.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtProfile, "No SMID entered")
                    blnRet = False
                End If
                If Me.txtProfile.Text.Trim.Length > 10 Then
                    Me.ErrorProvider1.SetError(Me.txtProfile, "SMID is too long")
                    blnRet = False
                End If
                'Check the identity is valid
                If txtProfile.Text.Trim.Length > 0 Then
                    Dim strCheck As String = CConnection.PackageStringList("Lib_Utils.IsValidGenericIdentityChar", txtProfile.Text)
                    If strCheck = "F" Then
                        Me.ErrorProvider1.SetError(Me.txtProfile, "the profile has invalid characters can be only [0-9][A-Z]+-_, with no spaces")
                        blnRet = False
                    End If
                End If
                'Check that it hasn't been used
                Dim ocmd As New OleDb.OleDbCommand("Select count(*) from customer where SMID ='" & _
                    Me.txtSMID.Text.Trim & "'", CConnection.DbConnection)
                Try
                    If CConnection.ConnOpen Then
                        Dim intProfiles = CInt(ocmd.ExecuteScalar)
                        Me.blnExistingSMID = (intProfiles > 0)
                        If Me.blnExistingSMID Then
                            Me.NotificationWindow1.Notify("SMID - " & Me.txtSMID.Text & _
                            " has " & intProfiles & " profiles at the moment.")
                        End If
                    End If
                Catch
                Finally
                    CConnection.SetConnClosed()
                End Try
            Case Stages.GetIndustry
                    'Need to check that the profile length is in range 
                    If Me.txtProfile.Text.Trim.Length = 0 Then
                        Me.ErrorProvider1.SetError(Me.txtProfile, "No Profile entered")
                        blnRet = False
                    End If
                    If Me.txtProfile.Text.Trim.Length > 10 Then
                        Me.ErrorProvider1.SetError(Me.txtProfile, "Profile is too long")
                        blnRet = False
                    End If
                    Dim ocmd As New OleDb.OleDbCommand("Select count(*) from customer where SMID ='" & _
                        Me.txtSMID.Text.Trim & "' and profile ='" & _
                        Me.txtProfile.Text.Trim & "'", CConnection.DbConnection)
                    Try
                        If CConnection.ConnOpen Then
                            Dim intProfiles = CInt(ocmd.ExecuteScalar)
                            Me.blnExistingProfile = (intProfiles > 0)
                            If Me.blnExistingProfile Then

                                Me.ErrorProvider1.SetError(Me.txtProfile, "Profile name already in use")

                                Me.NotificationWindow2.Notify("SMID - " & Me.txtSMID.Text & _
                                "-" & Me.txtProfile.Text & " already exists.")
                                blnRet = False
                            End If
                        End If
                    Catch
                    Finally
                        CConnection.SetConnClosed()
                    End Try

        End Select

        If blnRet Then
            If intStage = Stages.GetCompanyName And Me.FormMode = FormModes.Template Then
                intStage = Stages.getPanel
            End If
            Me.intStage += 1
        End If


        Return blnRet
    End Function

    Private Sub StageHasChanged()
        hidepanels()
        ValidateEntry() 'Then Exit Sub
        Me.cmdOk.Hide()
        Select Case intStage
            Case Stages.InitialStage
                Me.PanelSMID.Show()
                Me.cmdPrev.Hide()
                Me.txtSMID.Enabled = True
                Me.txtSMID.Focus()
                Me.Text = "Customer Creation"

            Case Stages.GetProfile
                If Me.txtSMID.Text.Trim.Length > 0 Then Me.txtSMID.Enabled = False
                Me.cmdPrev.Show()
                Me.PanelProfile.Show()
                Me.txtProfile.Focus()
                Me.Text = "Customer Creation - " & Me.txtSMID.Text
                If Me.blnExistingSMID Then
                    Me.Label3.Text = Me.StdSMIDPresntProfilePanelMessage
                Else
                    Me.Label3.Text = Me.StdProfilePanelMessage
                End If
            Case Stages.GetCompanyName
                Me.pnlCompanyName.Show()
                Me.txtCompanyName.Focus()
                Me.Text = "Customer Creation - " & Me.txtSMID.Text & "-" & Me.txtProfile.Text
            Case Stages.GetIndustry
                Me.PanelIndustryCode.Show()
                Me.cbIndustry.Focus()
                Me.Text = "Customer Creation - " & Me.txtSMID.Text & "-" & Me.txtProfile.Text
            Case Stages.getPanel
                Dim oIndDef As MedscreenLib.Defaults.IndustryDefault = MedscreenLib.Glossary.Glossary.CustomerDefaults.IndustryDefaults.Item(Me.IndustryCode)
                If Not oIndDef Is Nothing Then
                    Dim strPanel As String = oIndDef.TestPanel
                    'Dim objTP As Intranet.intranet.TestPanels.TestPanel
                    'objTP = Intranet.intranet.Support.cSupport.TestPanels.Item(strPanel)
                    'Dim intPos As Integer
                    'If Not objTP Is Nothing Then
                    '    intPos = Intranet.intranet.Support.cSupport.TestPanels.IndexOf(objTP)
                    '    Me.CBPanel.SelectedIndex = intPos
                    'End If
                End If

                Me.pnlTestPanel.Show()
                Me.CBPanel.Focus()
            Case Stages.FinalStage
                Me.PanelFinish.Show()
                Me.cmdOk.Show()
                Me.cmdOk.Focus()
        End Select

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.myMode = FormModes.Create Then
            intStage += 1
        Else
            StageHasChanged()
            'If intStage = Stages.GetCompanyName Then intStage = Stages.FinalStage
            'If intStage = Stages.GetProfile Then
            '    intStage = Stages.GetCompanyName
            'End If
            End If
    End Sub

    Private Sub cmdPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrev.Click
        intStage -= 2
        StageHasChanged()
    End Sub

    Private Sub txtProfile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProfile.TextChanged

    End Sub

    Private Sub txtProfile_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProfile.Leave, txtSMID.Leave, cbIndustry.Leave, txtCompanyName.Leave
        Me.Button1.Focus()
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    End Sub
End Class
