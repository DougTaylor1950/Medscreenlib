'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-22 06:30:28+00 $
'$Log: frmAbout.vb,v $
Imports MedscreenLib
Imports Intranet.intranet.Support
Imports System.Windows.Forms
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCallOut
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to do callouts
''' </summary>
''' <remarks>A call out is an booking for a collection arranged at less than two hours notice
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCallOut
    Inherits System.Windows.Forms.Form

    Private objSampleTemplates As Glossary.PhraseCollection
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call



        'Set up test type 
        Me.cbCategory.DataSource = MedscreenLib.Glossary.Glossary.CalloutCategory
        Me.cbCategory.DisplayMember = "PhraseText"
        Me.cbCategory.ValueMember = "PhraseID"

        'Me.CBCollectingOfficer.DataSource = cSupport.CollectingOfficers
        'Me.CBCollectingOfficer.DisplayMember = "ReverseName"
        'Me.CBCollectingOfficer.ValueMember = "OfficerID"

        Me.cBLocality.SelectedIndex = 0

        'Add any initialization after the InitializeComponent() call

        'Set up reason for test 
        Dim strReasonForTestCallOut As String = MedscreenCommonGUIConfig.CollectionQuery.Item("ReasonForTestCallOut")
        Dim objReason = New Glossary.PhraseCollection(strReasonForTestCallOut, Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , "MEDSCREEN,WRK_STD")
        Me.cbRandom.DataSource = objReason
        Me.cbRandom.DisplayMember = "PhraseText"
        Me.cbRandom.ValueMember = "PhraseID"

        Me.cbSampleTemplate.DataSource = objSampleTemplates
        Me.cbSampleTemplate.DisplayMember = "PhraseText"
        Me.cbSampleTemplate.ValueMember = "PhraseID"


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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSubmit As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPinNumber As System.Windows.Forms.TextBox
    Friend WithEvents cmdEditSite As System.Windows.Forms.Button
    Friend WithEvents cmdNextDueVessel As System.Windows.Forms.Button
    Friend WithEvents cmdNewSite As System.Windows.Forms.Button
    Friend WithEvents cbCustSites As System.Windows.Forms.ComboBox
    Friend WithEvents lblVessel As System.Windows.Forms.Label
    Friend WithEvents cbSiteVessel As System.Windows.Forms.ComboBox
    Friend WithEvents lblSite As System.Windows.Forms.Label
    Friend WithEvents btnImage As System.Windows.Forms.ImageList
    Friend WithEvents cbVessel As System.Windows.Forms.ComboBox
    Friend WithEvents cmdAddVessel As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCallerName As System.Windows.Forms.TextBox
    Friend WithEvents txtContactNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSiteAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtAccessProblems As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerContact As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DTCollDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblCollDate As System.Windows.Forms.Label
    Friend WithEvents DtCallTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents updnNoDonors As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblNDonors As System.Windows.Forms.Label
    Friend WithEvents txtPO As System.Windows.Forms.TextBox
    Friend WithEvents lblPONumber As System.Windows.Forms.Label
    Friend WithEvents txtCollComments As System.Windows.Forms.TextBox
    Friend WithEvents txtCmNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents lblCollOfficer As System.Windows.Forms.Label
    Friend WithEvents DTSentToOfficer As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblSentToOff As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cBLocality As System.Windows.Forms.ComboBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuPorts As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddPort As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditPort As System.Windows.Forms.MenuItem
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents pnlWarning As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents txtAddress As System.Windows.Forms.RichTextBox
    Friend WithEvents cmdFindPort As System.Windows.Forms.Button
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCoOfficer As System.Windows.Forms.TextBox
    Friend WithEvents cmdFindOfficer As System.Windows.Forms.Button
    Friend WithEvents txtMatrix As System.Windows.Forms.TextBox
    Friend WithEvents cmdViewPanel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbSampleTemplate As System.Windows.Forms.ComboBox
    Friend WithEvents cbRandom As System.Windows.Forms.ComboBox
    Friend WithEvents lblTestPanel As System.Windows.Forms.Label
    Friend WithEvents lblCollType As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCallOut))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSubmit = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtMatrix = New System.Windows.Forms.TextBox()
        Me.cmdViewPanel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbSampleTemplate = New System.Windows.Forms.ComboBox()
        Me.cbRandom = New System.Windows.Forms.ComboBox()
        Me.lblTestPanel = New System.Windows.Forms.Label()
        Me.lblCollType = New System.Windows.Forms.Label()
        Me.cmdFindOfficer = New System.Windows.Forms.Button()
        Me.txtCoOfficer = New System.Windows.Forms.TextBox()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.cmdFindPort = New System.Windows.Forms.Button()
        Me.txtLocation = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtAddress = New System.Windows.Forms.RichTextBox()
        Me.pnlWarning = New System.Windows.Forms.Panel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cBLocality = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DTSentToOfficer = New System.Windows.Forms.DateTimePicker()
        Me.lblSentToOff = New System.Windows.Forms.Label()
        Me.lblCollOfficer = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.txtCmNumber = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtPO = New System.Windows.Forms.TextBox()
        Me.lblPONumber = New System.Windows.Forms.Label()
        Me.updnNoDonors = New System.Windows.Forms.NumericUpDown()
        Me.lblNDonors = New System.Windows.Forms.Label()
        Me.DTCollDate = New System.Windows.Forms.DateTimePicker()
        Me.lblCollDate = New System.Windows.Forms.Label()
        Me.txtCollComments = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtCustomerContact = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtAccessProblems = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cbCategory = New System.Windows.Forms.ComboBox()
        Me.txtSiteAddress = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtContactNumber = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtCallerName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdAddVessel = New System.Windows.Forms.Button()
        Me.cbVessel = New System.Windows.Forms.ComboBox()
        Me.cmdEditSite = New System.Windows.Forms.Button()
        Me.cmdNextDueVessel = New System.Windows.Forms.Button()
        Me.btnImage = New System.Windows.Forms.ImageList(Me.components)
        Me.cmdNewSite = New System.Windows.Forms.Button()
        Me.cbCustSites = New System.Windows.Forms.ComboBox()
        Me.lblVessel = New System.Windows.Forms.Label()
        Me.cbSiteVessel = New System.Windows.Forms.ComboBox()
        Me.lblSite = New System.Windows.Forms.Label()
        Me.txtPinNumber = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DtCallTime = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.mnuPorts = New System.Windows.Forms.MenuItem()
        Me.mnuAddPort = New System.Windows.Forms.MenuItem()
        Me.mnuEditPort = New System.Windows.Forms.MenuItem()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.pnlWarning.SuspendLayout()
        CType(Me.updnNoDonors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdSubmit})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 537)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(920, 56)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(760, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(118, 23)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdSubmit
        '
        Me.cmdSubmit.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSubmit.Image = CType(resources.GetObject("cmdSubmit.Image"), System.Drawing.Bitmap)
        Me.cmdSubmit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSubmit.Location = New System.Drawing.Point(64, 16)
        Me.cmdSubmit.Name = "cmdSubmit"
        Me.cmdSubmit.Size = New System.Drawing.Size(118, 23)
        Me.cmdSubmit.TabIndex = 1
        Me.cmdSubmit.Text = "&Submit Collection"
        Me.cmdSubmit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(920, 537)
        Me.Panel2.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtMatrix, Me.cmdViewPanel, Me.Label1, Me.cbSampleTemplate, Me.cbRandom, Me.lblTestPanel, Me.lblCollType, Me.cmdFindOfficer, Me.txtCoOfficer, Me.txtCustomer, Me.cmdFindPort, Me.txtLocation, Me.Panel3, Me.pnlWarning, Me.cBLocality, Me.Label12, Me.DTSentToOfficer, Me.lblSentToOff, Me.lblCollOfficer, Me.lblLocation, Me.lblInfo, Me.txtCmNumber, Me.Label10, Me.txtPO, Me.lblPONumber, Me.updnNoDonors, Me.lblNDonors, Me.DTCollDate, Me.lblCollDate, Me.txtCollComments, Me.Label9, Me.txtCustomerContact, Me.Label8, Me.txtAccessProblems, Me.Label7, Me.Label6, Me.cbCategory, Me.txtSiteAddress, Me.Label5, Me.txtContactNumber, Me.Label4, Me.txtCallerName, Me.Label3, Me.cmdAddVessel, Me.cbVessel, Me.cmdEditSite, Me.cmdNextDueVessel, Me.cmdNewSite, Me.cbCustSites, Me.lblVessel, Me.cbSiteVessel, Me.lblSite, Me.txtPinNumber, Me.Label2, Me.DtCallTime, Me.Label11})
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(920, 537)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Call Out Collection"
        '
        'txtMatrix
        '
        Me.txtMatrix.Location = New System.Drawing.Point(136, 400)
        Me.txtMatrix.Multiline = True
        Me.txtMatrix.Name = "txtMatrix"
        Me.txtMatrix.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMatrix.Size = New System.Drawing.Size(472, 56)
        Me.txtMatrix.TabIndex = 125
        Me.txtMatrix.Text = ""
        '
        'cmdViewPanel
        '
        Me.cmdViewPanel.Image = CType(resources.GetObject("cmdViewPanel.Image"), System.Drawing.Bitmap)
        Me.cmdViewPanel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdViewPanel.Location = New System.Drawing.Point(512, 464)
        Me.cmdViewPanel.Name = "cmdViewPanel"
        Me.cmdViewPanel.Size = New System.Drawing.Size(96, 23)
        Me.cmdViewPanel.TabIndex = 124
        Me.cmdViewPanel.Text = "View Panel"
        Me.cmdViewPanel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdViewPanel, "See details of the tests to be done")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(32, 400)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 123
        Me.Label1.Text = "Test Type"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbSampleTemplate
        '
        Me.cbSampleTemplate.ItemHeight = 13
        Me.cbSampleTemplate.Location = New System.Drawing.Point(136, 464)
        Me.cbSampleTemplate.Name = "cbSampleTemplate"
        Me.cbSampleTemplate.Size = New System.Drawing.Size(368, 21)
        Me.cbSampleTemplate.TabIndex = 120
        '
        'cbRandom
        '
        Me.cbRandom.ItemHeight = 13
        Me.cbRandom.Location = New System.Drawing.Point(136, 368)
        Me.cbRandom.Name = "cbRandom"
        Me.cbRandom.Size = New System.Drawing.Size(472, 21)
        Me.cbRandom.TabIndex = 119
        '
        'lblTestPanel
        '
        Me.lblTestPanel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestPanel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTestPanel.Location = New System.Drawing.Point(32, 464)
        Me.lblTestPanel.Name = "lblTestPanel"
        Me.lblTestPanel.TabIndex = 122
        Me.lblTestPanel.Text = "Sample Template"
        Me.lblTestPanel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCollType
        '
        Me.lblCollType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCollType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCollType.Location = New System.Drawing.Point(32, 368)
        Me.lblCollType.Name = "lblCollType"
        Me.lblCollType.TabIndex = 121
        Me.lblCollType.Text = "Reason for test"
        Me.lblCollType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdFindOfficer
        '
        Me.cmdFindOfficer.Image = CType(resources.GetObject("cmdFindOfficer.Image"), System.Drawing.Bitmap)
        Me.cmdFindOfficer.Location = New System.Drawing.Point(352, 328)
        Me.cmdFindOfficer.Name = "cmdFindOfficer"
        Me.cmdFindOfficer.Size = New System.Drawing.Size(48, 23)
        Me.cmdFindOfficer.TabIndex = 118
        Me.ToolTip1.SetToolTip(Me.cmdFindOfficer, "Find Port")
        '
        'txtCoOfficer
        '
        Me.txtCoOfficer.Location = New System.Drawing.Point(136, 328)
        Me.txtCoOfficer.Name = "txtCoOfficer"
        Me.txtCoOfficer.Size = New System.Drawing.Size(200, 20)
        Me.txtCoOfficer.TabIndex = 117
        Me.txtCoOfficer.Text = ""
        '
        'txtCustomer
        '
        Me.txtCustomer.Enabled = False
        Me.txtCustomer.Location = New System.Drawing.Point(136, 24)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(760, 20)
        Me.txtCustomer.TabIndex = 116
        Me.txtCustomer.Text = ""
        '
        'cmdFindPort
        '
        Me.cmdFindPort.Image = CType(resources.GetObject("cmdFindPort.Image"), System.Drawing.Bitmap)
        Me.cmdFindPort.Location = New System.Drawing.Point(400, 200)
        Me.cmdFindPort.Name = "cmdFindPort"
        Me.cmdFindPort.Size = New System.Drawing.Size(48, 23)
        Me.cmdFindPort.TabIndex = 115
        Me.ToolTip1.SetToolTip(Me.cmdFindPort, "Find Port")
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(136, 200)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(256, 20)
        Me.txtLocation.TabIndex = 6
        Me.txtLocation.Text = ""
        '
        'Panel3
        '
        Me.Panel3.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtAddress})
        Me.Panel3.Location = New System.Drawing.Point(624, 208)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(272, 306)
        Me.Panel3.TabIndex = 113
        '
        'txtAddress
        '
        Me.txtAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(268, 302)
        Me.txtAddress.TabIndex = 1
        Me.txtAddress.Text = ""
        '
        'pnlWarning
        '
        Me.pnlWarning.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlWarning.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label13})
        Me.pnlWarning.Location = New System.Drawing.Point(40, 160)
        Me.pnlWarning.Name = "pnlWarning"
        Me.pnlWarning.Size = New System.Drawing.Size(320, 32)
        Me.pnlWarning.TabIndex = 112
        Me.pnlWarning.Visible = False
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Info
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(316, 28)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Ask if call out is at Donor's Home - If yes remind customer we won't go to Donors" & _
        "' Homes!"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cBLocality
        '
        Me.cBLocality.ItemHeight = 13
        Me.cBLocality.Items.AddRange(New Object() {"UK", "Overseas"})
        Me.cBLocality.Location = New System.Drawing.Point(248, 48)
        Me.cBLocality.Name = "cBLocality"
        Me.cBLocality.Size = New System.Drawing.Size(104, 21)
        Me.cBLocality.TabIndex = 111
        Me.cBLocality.TabStop = False
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(8, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(112, 23)
        Me.Label12.TabIndex = 110
        Me.Label12.Text = "Customer SMID"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTSentToOfficer
        '
        Me.DTSentToOfficer.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.DTSentToOfficer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTSentToOfficer.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTSentToOfficer.Location = New System.Drawing.Point(480, 328)
        Me.DTSentToOfficer.Name = "DTSentToOfficer"
        Me.DTSentToOfficer.Size = New System.Drawing.Size(119, 20)
        Me.DTSentToOfficer.TabIndex = 19
        Me.DTSentToOfficer.Visible = False
        '
        'lblSentToOff
        '
        Me.lblSentToOff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSentToOff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSentToOff.Location = New System.Drawing.Point(376, 328)
        Me.lblSentToOff.Name = "lblSentToOff"
        Me.lblSentToOff.TabIndex = 109
        Me.lblSentToOff.Text = "Sent To Officer"
        Me.lblSentToOff.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblSentToOff.Visible = False
        '
        'lblCollOfficer
        '
        Me.lblCollOfficer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCollOfficer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCollOfficer.Location = New System.Drawing.Point(32, 328)
        Me.lblCollOfficer.Name = "lblCollOfficer"
        Me.lblCollOfficer.TabIndex = 107
        Me.lblCollOfficer.Text = "Collecting Officer"
        Me.lblCollOfficer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCollOfficer.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocation.Location = New System.Drawing.Point(32, 200)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.TabIndex = 105
        Me.lblLocation.Text = "Location Port/Site"
        Me.lblLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInfo
        '
        Me.lblInfo.ForeColor = System.Drawing.Color.Red
        Me.lblInfo.Location = New System.Drawing.Point(32, 504)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(640, 16)
        Me.lblInfo.TabIndex = 104
        Me.lblInfo.Text = "Label13"
        '
        'txtCmNumber
        '
        Me.txtCmNumber.Enabled = False
        Me.txtCmNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCmNumber.Location = New System.Drawing.Point(480, 304)
        Me.txtCmNumber.Name = "txtCmNumber"
        Me.txtCmNumber.Size = New System.Drawing.Size(88, 20)
        Me.txtCmNumber.TabIndex = 18
        Me.txtCmNumber.Text = ""
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label10.Location = New System.Drawing.Point(368, 304)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 100
        Me.Label10.Text = "CM Number"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPO
        '
        Me.txtPO.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPO.Location = New System.Drawing.Point(480, 272)
        Me.txtPO.Name = "txtPO"
        Me.txtPO.Size = New System.Drawing.Size(136, 20)
        Me.txtPO.TabIndex = 17
        Me.txtPO.Text = ""
        '
        'lblPONumber
        '
        Me.lblPONumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPONumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblPONumber.Location = New System.Drawing.Point(384, 272)
        Me.lblPONumber.Name = "lblPONumber"
        Me.lblPONumber.Size = New System.Drawing.Size(88, 23)
        Me.lblPONumber.TabIndex = 98
        Me.lblPONumber.Text = "Purchase Order"
        Me.lblPONumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'updnNoDonors
        '
        Me.updnNoDonors.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.updnNoDonors.Location = New System.Drawing.Point(528, 200)
        Me.updnNoDonors.Name = "updnNoDonors"
        Me.updnNoDonors.Size = New System.Drawing.Size(36, 20)
        Me.updnNoDonors.TabIndex = 14
        '
        'lblNDonors
        '
        Me.lblNDonors.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNDonors.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNDonors.Location = New System.Drawing.Point(464, 200)
        Me.lblNDonors.Name = "lblNDonors"
        Me.lblNDonors.Size = New System.Drawing.Size(42, 23)
        Me.lblNDonors.TabIndex = 13
        Me.lblNDonors.Text = "&Donors"
        Me.lblNDonors.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTCollDate
        '
        Me.DTCollDate.CustomFormat = "dd-MMM-yyyy"
        Me.DTCollDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTCollDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTCollDate.Location = New System.Drawing.Point(480, 224)
        Me.DTCollDate.Name = "DTCollDate"
        Me.DTCollDate.Size = New System.Drawing.Size(119, 20)
        Me.DTCollDate.TabIndex = 15
        '
        'lblCollDate
        '
        Me.lblCollDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCollDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCollDate.Location = New System.Drawing.Point(384, 224)
        Me.lblCollDate.Name = "lblCollDate"
        Me.lblCollDate.Size = New System.Drawing.Size(88, 23)
        Me.lblCollDate.TabIndex = 92
        Me.lblCollDate.Text = "Collection &Date"
        Me.lblCollDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCollComments
        '
        Me.txtCollComments.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtCollComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCollComments.Location = New System.Drawing.Point(480, 136)
        Me.txtCollComments.Multiline = True
        Me.txtCollComments.Name = "txtCollComments"
        Me.txtCollComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCollComments.Size = New System.Drawing.Size(416, 48)
        Me.txtCollComments.TabIndex = 12
        Me.txtCollComments.Text = ""
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label9.Location = New System.Drawing.Point(376, 144)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(100, 32)
        Me.Label9.TabIndex = 90
        Me.Label9.Text = "Collection Co&mments"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCustomerContact
        '
        Me.txtCustomerContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomerContact.Location = New System.Drawing.Point(136, 296)
        Me.txtCustomerContact.Name = "txtCustomerContact"
        Me.txtCustomerContact.Size = New System.Drawing.Size(200, 20)
        Me.txtCustomerContact.TabIndex = 9
        Me.txtCustomerContact.Text = ""
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(32, 304)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 88
        Me.Label8.Text = "Contact &On Site"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtAccessProblems
        '
        Me.txtAccessProblems.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtAccessProblems.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccessProblems.Location = New System.Drawing.Point(480, 56)
        Me.txtAccessProblems.Multiline = True
        Me.txtAccessProblems.Name = "txtAccessProblems"
        Me.txtAccessProblems.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtAccessProblems.Size = New System.Drawing.Size(416, 48)
        Me.txtAccessProblems.TabIndex = 10
        Me.txtAccessProblems.Text = ""
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(376, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 86
        Me.Label7.Text = "Access P&roblems"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(376, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "Category"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbCategory
        '
        Me.cbCategory.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbCategory.Location = New System.Drawing.Point(480, 112)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(416, 21)
        Me.cbCategory.TabIndex = 11
        '
        'txtSiteAddress
        '
        Me.txtSiteAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSiteAddress.Location = New System.Drawing.Point(136, 248)
        Me.txtSiteAddress.Multiline = True
        Me.txtSiteAddress.Name = "txtSiteAddress"
        Me.txtSiteAddress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSiteAddress.Size = New System.Drawing.Size(187, 48)
        Me.txtSiteAddress.TabIndex = 8
        Me.txtSiteAddress.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSiteAddress, "Please create or select the correct customer site to provide the address")
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(32, 248)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 82
        Me.Label5.Text = "Site Address"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtContactNumber
        '
        Me.txtContactNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContactNumber.Location = New System.Drawing.Point(136, 104)
        Me.txtContactNumber.Name = "txtContactNumber"
        Me.txtContactNumber.Size = New System.Drawing.Size(187, 20)
        Me.txtContactNumber.TabIndex = 4
        Me.txtContactNumber.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(32, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 80
        Me.Label4.Text = "Contact Number "
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCallerName
        '
        Me.txtCallerName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCallerName.Location = New System.Drawing.Point(136, 80)
        Me.txtCallerName.Name = "txtCallerName"
        Me.txtCallerName.Size = New System.Drawing.Size(187, 20)
        Me.txtCallerName.TabIndex = 3
        Me.txtCallerName.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(32, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 78
        Me.Label3.Text = "Caller Name"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdAddVessel
        '
        Me.cmdAddVessel.Image = CType(resources.GetObject("cmdAddVessel.Image"), System.Drawing.Bitmap)
        Me.cmdAddVessel.Location = New System.Drawing.Point(320, 224)
        Me.cmdAddVessel.Name = "cmdAddVessel"
        Me.cmdAddVessel.Size = New System.Drawing.Size(23, 23)
        Me.cmdAddVessel.TabIndex = 77
        '
        'cbVessel
        '
        Me.cbVessel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbVessel.Location = New System.Drawing.Point(136, 224)
        Me.cbVessel.Name = "cbVessel"
        Me.cbVessel.Size = New System.Drawing.Size(187, 21)
        Me.cbVessel.TabIndex = 7
        '
        'cmdEditSite
        '
        Me.cmdEditSite.Image = CType(resources.GetObject("cmdEditSite.Image"), System.Drawing.Bitmap)
        Me.cmdEditSite.Location = New System.Drawing.Point(112, 224)
        Me.cmdEditSite.Name = "cmdEditSite"
        Me.cmdEditSite.Size = New System.Drawing.Size(23, 23)
        Me.cmdEditSite.TabIndex = 74
        '
        'cmdNextDueVessel
        '
        Me.cmdNextDueVessel.Image = CType(resources.GetObject("cmdNextDueVessel.Image"), System.Drawing.Bitmap)
        Me.cmdNextDueVessel.ImageIndex = 0
        Me.cmdNextDueVessel.ImageList = Me.btnImage
        Me.cmdNextDueVessel.Location = New System.Drawing.Point(344, 224)
        Me.cmdNextDueVessel.Name = "cmdNextDueVessel"
        Me.cmdNextDueVessel.Size = New System.Drawing.Size(23, 23)
        Me.cmdNextDueVessel.TabIndex = 73
        '
        'btnImage
        '
        Me.btnImage.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.btnImage.ImageSize = New System.Drawing.Size(16, 16)
        Me.btnImage.ImageStream = CType(resources.GetObject("btnImage.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.btnImage.TransparentColor = System.Drawing.Color.Transparent
        '
        'cmdNewSite
        '
        Me.cmdNewSite.Image = CType(resources.GetObject("cmdNewSite.Image"), System.Drawing.Bitmap)
        Me.cmdNewSite.Location = New System.Drawing.Point(320, 224)
        Me.cmdNewSite.Name = "cmdNewSite"
        Me.cmdNewSite.Size = New System.Drawing.Size(23, 23)
        Me.cmdNewSite.TabIndex = 72
        '
        'cbCustSites
        '
        Me.cbCustSites.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbCustSites.Location = New System.Drawing.Point(136, 224)
        Me.cbCustSites.Name = "cbCustSites"
        Me.cbCustSites.Size = New System.Drawing.Size(187, 21)
        Me.cbCustSites.TabIndex = 71
        Me.cbCustSites.Visible = False
        '
        'lblVessel
        '
        Me.lblVessel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVessel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVessel.Location = New System.Drawing.Point(8, 224)
        Me.lblVessel.Name = "lblVessel"
        Me.lblVessel.TabIndex = 70
        Me.lblVessel.Text = "Vessel Name"
        Me.lblVessel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbSiteVessel
        '
        Me.cbSiteVessel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSiteVessel.Items.AddRange(New Object() {"Site", "Vessel"})
        Me.cbSiteVessel.Location = New System.Drawing.Point(136, 136)
        Me.cbSiteVessel.Name = "cbSiteVessel"
        Me.cbSiteVessel.Size = New System.Drawing.Size(75, 21)
        Me.cbSiteVessel.TabIndex = 5
        '
        'lblSite
        '
        Me.lblSite.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSite.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSite.Location = New System.Drawing.Point(32, 128)
        Me.lblSite.Name = "lblSite"
        Me.lblSite.TabIndex = 68
        Me.lblSite.Text = "Site or Vessel"
        Me.lblSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPinNumber
        '
        Me.txtPinNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPinNumber.Location = New System.Drawing.Point(136, 48)
        Me.txtPinNumber.Name = "txtPinNumber"
        Me.txtPinNumber.TabIndex = 2
        Me.txtPinNumber.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(32, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Pin Number"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'DtCallTime
        '
        Me.DtCallTime.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.DtCallTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DtCallTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtCallTime.Location = New System.Drawing.Point(480, 248)
        Me.DtCallTime.Name = "DtCallTime"
        Me.DtCallTime.Size = New System.Drawing.Size(119, 20)
        Me.DtCallTime.TabIndex = 16
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(376, 248)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 94
        Me.Label11.Text = "Call &Taken At"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPorts})
        '
        'mnuPorts
        '
        Me.mnuPorts.Index = 0
        Me.mnuPorts.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddPort, Me.mnuEditPort})
        Me.mnuPorts.Text = "&Ports"
        '
        'mnuAddPort
        '
        Me.mnuAddPort.Index = 0
        Me.mnuAddPort.Text = "&Add Port"
        '
        'mnuEditPort
        '
        Me.mnuEditPort.Index = 1
        Me.mnuEditPort.Text = "&Edit Port"
        '
        'frmCallOut
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(920, 593)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1})
        Me.Menu = Me.MainMenu1
        Me.Name = "frmCallOut"
        Me.Text = "Callout Collections "
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.pnlWarning.ResumeLayout(False)
        CType(Me.updnNoDonors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declarations"



    Private myCollection As Intranet.intranet.jobs.CMJob
    Private myClient As Intranet.intranet.customerns.Client
    Private myOfficer As Intranet.intranet.jobs.CollectingOfficer

    Private iniFile As New MedscreenLib.IniFile.IniFiles()
    Private strFileDir As String
    Private objVessels As Intranet.intranet.customerns.CVesselCollection
    Private CustSites As Intranet.intranet.Address.CustomerSites
    Private blnFormLoaded As Boolean = False
    Private VesselFirstDate As Date = Now
    Private dblCost As Double
    Private blnFirstPass As Boolean = True
    Private blnSecondPass As Boolean = False
    Private blnDrawing As Boolean = True
#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection that is or will be associated with this callout
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Collection() As Intranet.intranet.jobs.CMJob
        Get
            Return myCollection
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CMJob)
            myCollection = Value                                'Set collection
            'Clear all error providers
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
            Me.ErrorProvider1.SetError(Me.txtPO, "")
            Me.ErrorProvider1.SetError(Me.txtContactNumber, "")
            Me.ErrorProvider1.SetError(Me.txtPinNumber, "")
            Me.ErrorProvider1.SetError(Me.txtAccessProblems, "")
            Me.ErrorProvider1.SetError(Me.txtCustomerContact, "")
            If myCollection Is Nothing Then                     'If no collection clear form
                clearForm()
                myClient = Nothing
                Me.blnFirstPass = True
            Else                                                'Else fill form with data
                myClient = cSupport.CustList.Item(myCollection.ClientId)
                SetUpClient()
                FillForm()
                Me.lblCollOfficer.Show()
                Me.lblSentToOff.Show()
                'Me.CBCollectingOfficer.Show()
                Me.DTSentToOfficer.Show()
                blnFormLoaded = True
                Me.blnFirstPass = False     'Let then no it is a old collection
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clear all controls
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub clearForm()
        Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
        Me.ErrorProvider1.SetError(Me.txtPO, "")
        Me.ErrorProvider1.SetError(Me.txtContactNumber, "")
        Me.ErrorProvider1.SetError(Me.txtPinNumber, "")
        Me.ErrorProvider1.SetError(Me.txtAccessProblems, "")
        Me.ErrorProvider1.SetError(Me.txtCustomerContact, "")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disable the form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DisableForm()
        Me.cbCategory.Enabled = False
        'Me.SmidCombo1.Enabled = False
        'Me.cbCustomer.Enabled = False

        Me.cbCustSites.Enabled = False
        Me.txtLocation.Enabled = False
        Me.cbSiteVessel.Enabled = False
        Me.cbVessel.Enabled = False
        Me.cmdAddVessel.Enabled = False
        Me.cmdEditSite.Enabled = False
        Me.cmdNewSite.Enabled = False
        Me.cmdNextDueVessel.Enabled = False
        Me.DtCallTime.Enabled = False
        Me.DTCollDate.Enabled = False
        Me.txtAccessProblems.Enabled = False
        Me.txtCallerName.Enabled = False
        Me.txtCollComments.Enabled = False
        Me.txtContactNumber.Enabled = False
        Me.txtCustomerContact.Enabled = False
        Me.txtPinNumber.Enabled = False
        Me.txtPO.Enabled = False
        Me.txtSiteAddress.Enabled = False
        Me.updnNoDonors.Enabled = False
        Me.lblCollOfficer.Show()
        Me.lblSentToOff.Show()
        'Me.CBCollectingOfficer.Show()
        Me.DTSentToOfficer.Show()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populate the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillForm()
        Dim objPort As Intranet.intranet.jobs.Port
        Dim myClient As Intranet.intranet.customerns.Client
        Dim strCurr As String
        Dim cr As System.Globalization.CultureInfo
        Dim strCust As String
        Dim oVessel As Intranet.intranet.customerns.Vessel

        Me.Text = "Call out Collections - " & Application.ProductVersion
        With myCollection
            'strCust = .ClientId
            Me.txtCustomer.Text = .SmClient.Info
            myClient = .SmClient
            'Me.cbCustomer.Siblings = Me.SmidCombo1.Siblings

            'Me.cbCustomer.SelectedProfile = myClient
            Me.txtCmNumber.Text = .ID                     'Set the ID of the collection


            'If we don't have client info (should never occur)
            If myClient Is Nothing Then
                'Me.txtCustomer.Text = .ClientId
                'TODO Align combo to correct customer
            Else
                'Check customer requirements Highlight if customer demands a PO Number
                If myClient.InvoiceInfo.PoRequired Then
                    Me.lblPONumber.Text = "Purchase Order*"
                    Me.lblPONumber.ForeColor = System.Drawing.SystemColors.ActiveCaption
                Else
                    Me.lblPONumber.Text = "Purchase Order"
                    Me.lblPONumber.ForeColor = System.Drawing.SystemColors.ControlText
                End If
            End If

            'If it is a vessel dependent collection then say so 
            Me.lblSite.Show()
            If .IsVessel Then
                Me.cbSiteVessel.SelectedIndex = 1
                Me.cbVessel.Show()
                Me.lblVessel.Show()

                Me.cbCustSites.Hide()
                Me.cmdEditSite.Hide()
                Me.cmdNewSite.Hide()
                Me.cmdNextDueVessel.Show()
                Me.cmdAddVessel.Show()

                Me.lblVessel.Text = "Vessel Name"

                Dim i As Integer

                If Not objVessels Is Nothing Then
                    If objVessels.Count > 0 And objVessels.Count < 300 Then
                        For i = 0 To Me.BindingContext(objVessels).Count
                            Me.BindingContext(objVessels).Position = i
                            oVessel = Me.BindingContext(objVessels).Current

                            If oVessel.VesselID = .vessel Or _
                                (oVessel.VesselName = .vessel And _
                                     oVessel.CustomerID = .ClientId) Then
                                Exit For
                            End If
                        Next
                    Else
                        .IsVessel = False
                        .Update()
                    End If

                End If
                'Me.BindingContext(objVessels).Current()
                'Me.txtVessel.Text = .vessel                 'Set the vessel 
            ElseIf myClient.Sites.Count > 0 Then
                Me.cbSiteVessel.SelectedIndex = 0
                Me.cbCustSites.SelectedIndex = -1
                Me.cbVessel.Hide()
                Me.cmdEditSite.Show()
                Me.cbCustSites.Show()
                Me.lblVessel.Text = "Site Name"
                Dim i As Integer
                If Not myClient Is Nothing Then
                    If Not .SiteObject Is Nothing Then
                        For i = 0 To myClient.Sites.Count - 1
                            If myClient.Sites.Item(i).SiteID = .vessel Then
                                Me.cbCustSites.SelectedIndex = i
                                Exit For
                            End If
                        Next

                        Me.cbSiteVessel.Text = Me.cbSiteVessel.Items(Me.cbSiteVessel.SelectedIndex)
                    End If
                End If
            Else
                Me.cbSiteVessel.Hide()
                Me.lblSite.Hide()
                Me.cbVessel.Hide()
                Me.lblVessel.Hide()
                Me.cmdAddVessel.Hide()
                Me.cmdEditSite.Hide()

                Me.cmdNewSite.Show()
                Me.cmdNextDueVessel.Hide()
                'Me.lblVessel.Hide()
            End If



            If .Port.Length > 0 Then
                objPort = .CollectionPort       'Set the site 
            End If   '
            '
            If .HasID < 1 Then                             'New collection
                Me.DTCollDate.Value = Now.AddHours(-2)
            Else
                Me.DTCollDate.Value = .CollectionDate   'Set the collection date 
            End If



            Me.updnNoDonors.Value = .NoOfDonors         'Set the expected number of Donors

            'If .UK Then                                 'If it is a UK based or Overseas
            '    Me.cBLocality.SelectedIndex = 0
            'Else
            '    Me.cBLocality.SelectedIndex = 1
            'End If

            'Me.txtWhototest.Text = .WhoToTest           'Indcate Who to Test 
            'Me.ToolTip1.SetToolTip(Me.txtWhototest, .WhoToTest)

            'Set which of the items is selected 
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            'Dim i As Integer
            Dim I1 As Integer
            If .ConfirmationType.Trim.Length > 0 Then
                For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Count - 1
                    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Current()
                    If objPhrase.PhraseID = .ConfirmationType Then
                        Exit For
                    End If
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Position += 1
                Next
            End If

            If .Tests.Trim.Length > 0 Then
                Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Position = 0
                For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Count - 1
                    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Current()
                    If objPhrase.PhraseID = .Tests Then
                        Exit For
                    End If
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Position += 1

                Next
            End If


            If .CollType.Length > 0 Then
                Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position = 0
                For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Count - 1
                    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Current()
                    If objPhrase.PhraseID = .CollType Then
                        Exit For
                    End If
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position += 1
                Next
            End If

            If .ConfirmationMethod.Length > 0 Then
                Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position = 0
                For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Count - 1
                    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Current()
                    If objPhrase.PhraseID = .ConfirmationMethod Then
                        Exit For
                    End If
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position += 1
                Next
            End If

            'Setup Reason for test
            Me.cbRandom.SelectedValue = .CollType
            'Me.cbConfMethod.SelectionLength = 0
            'Me.cbConfType.SelectionLength = 0
            'Me.cbCustomer.SelectionLength = 0
            'Me.cbTestType.SelectionLength = 0


            'SetListIndex(Me.cbRandom, .CollType)
            'SetListIndex(Me.cbConfMethod, .ConfirmationMethod)
            'SetListIndex(Me.cbConfType, .ConfirmationType)

            'Me.cbRandom.SelectedItem = .CollectionType
            'Fix up purchase order number 
            If .PurchaseOrder.Trim.Length = 0 AndAlso Me.myClient.ContractNumber.Trim.Length > 0 Then
                .PurchaseOrder = Me.myClient.ContractNumber
            End If
            Me.txtPO.Text = .PurchaseOrder                      'Set purchase Order 
            Me.txtCollComments.Text = .CollectionComments           'Set Comments about collection 
            Me.ToolTip1.SetToolTip(Me.txtCollComments, .CollectionComments)


            Me.txtPinNumber.Text = .PinNumber.Trim       'Should be initialised from the client object
            Me.txtCallerName.Text = .CallerName.Trim
            Me.txtContactNumber.Text = .ContactTel.Trim

            Me.txtAccessProblems.Text = .AdditionalComments.Trim
            Me.txtCustomerContact.Text = .ContactOnSite.Trim

            If .RequestReceived.Year < 2003 Then
                Me.DtCallTime.Value = Now
            Else
                Me.DtCallTime.Value = .RequestReceived
            End If

            'Is it a programme managed customer 


            'See if we can't display the site address.
            'Me.txtOther.Text = ""
            If .Addressid > 0 Then
                Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(.Addressid)
                'cmdSiteAddress.Enabled = (cadr Is Nothing)
                If Not cadr Is Nothing Then
                    Me.txtSiteAddress.Text = cadr.BlockAddress(vbCrLf)
                End If
            Else
                If Not objPort Is Nothing Then
                    Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(objPort.AddressId)
                    'cmdSiteAddress.Enabled = (cadr Is Nothing)
                    If Not cadr Is Nothing Then
                        'cmdSiteAddress.Enabled = (cadr Is Nothing)
                        Me.txtSiteAddress.Text = cadr.BlockAddress(vbCrLf)
                    End If

                End If

            End If
            If .HasID > 0 Then
                'Me.cbCustomer.Enabled = True    'Only allow the customer to be set for new customers
                'Me.SmidCombo1.Enabled = True
            Else
                'Me.cbCustomer.Enabled = False
                'Me.SmidCombo1.Enabled = False
            End If
            'Time and date sent to officer
            If .DateCalloutToOfficer > DateField.ZeroDate Then
                Me.DTSentToOfficer.Value = .DateCalloutToOfficer
            End If
            'Collecting Officer 
            'Me.CBCollectingOfficer.SelectedIndex = -1
            Me.myOfficer = Nothing
            If .CollOfficer.Trim.Length > 0 Then 'We have a collecting officer
                Me.myOfficer = cSupport.CollectingOfficers.Item(.CollOfficer)
                If myOfficer Is Nothing Then
                    Me.txtCoOfficer.Text = Me.myOfficer.Info2
                End If
            End If

        End With
        If Me.myCollection.OriginalTestScheduleId <> "?" Then
            Me.cbSampleTemplate.SelectedValue = myCollection.TestScheduleID

        End If
        GetForm()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up controls related to the client
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetUpClient()
        If Not myClient Is Nothing Then
            Collection.PinNumber = myClient.PinNumber
            Collection.ClientId = myClient.Identity
            objVessels = myClient.Vessels
            Me.lblInfo.Hide()
            Me.lblInfo.Text = ""
            'login templates
            Try
                Dim strSQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("SampleTemplate")
                'Dim strSQuery As String = "select identity,description from Samp_Tmpl_Header s " & _
                '"where s.removeflag='F' and action_type in ('.NET','REFER')"
                Me.objSampleTemplates = New Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)

                'Check to see we have a login template.
                'If objSampleTemplates.Count = 0 Then
                '    MsgBox("The login templates have not been set up for this customer, this MUST be done before collections can be created!", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                '    Me.DialogResult = DialogResult.Cancel
                '    Me.cmdCancel_Click(Nothing, Nothing)
                'End If
                Me.cbSampleTemplate.DataSource = objSampleTemplates
                Me.cbSampleTemplate.DisplayMember = "PhraseText"
                Me.cbSampleTemplate.ValueMember = "PhraseID"

                'Check to see if we have a value for the log in template in the defaults
                Dim strSampleTemplate As String = myClient.CollectionInfo.COSampleTemplate
                If strSampleTemplate.Trim.Length = 0 Then
                    strSampleTemplate = myClient.ReportingInfo.SampleTemplate
                End If
                myCollection.TestScheduleID = strSampleTemplate
                Me.cbSampleTemplate.SelectedValue = strSampleTemplate
            Catch ex As Exception
            End Try

            If myClient.InvoiceInfo.CreditCard_pay Then
                Me.lblInfo.Text = "This Client must pay by credit card"
                Me.lblInfo.Show()
            End If
            If myClient.Vessels.Count = 0 Then
                Me.cbVessel.Hide()
                Me.cmdNextDueVessel.Hide()
                myCollection.IsSiteVisit = True
                Me.cbSiteVessel.Enabled = False
                Me.cmdAddVessel.Hide()
                Me.cmdNewSite.Show()
                CustSites = Me.myClient.Sites
                Me.cbCustSites.Show()
                Me.cbCustSites.DataSource = CustSites
                Me.cbCustSites.DisplayMember = "Info"
                Me.cbCustSites.ValueMember = "SiteID"
                Me.lblVessel.Text = "Customer Sites"
            Else
                Me.cmdNextDueVessel.Show()
                Me.cmdNextDueVessel.ImageIndex = 0
                VesselFirstDate = Now
                'myCollection.IsVessel = True
                Me.cbVessel.Show()
                Me.cbSiteVessel.Enabled = True
                Me.cmdAddVessel.Show()
                Me.cmdNewSite.Hide()
                Me.cbCustSites.Hide()
                Me.lblVessel.Text = "Vessel Name"
                Me.cmdAddVessel.Show()
                CustSites = Me.myClient.Sites
                Me.cbCustSites.DataSource = CustSites
                Me.cbCustSites.DisplayMember = "Info"
                Me.cbCustSites.ValueMember = "SiteID"
            End If


        Else
                'objVessels = cSupport.Vessels
        End If
        cbVessel.DataSource = objVessels
        cbVessel.DisplayMember = "Info"
        cbVessel.ValueMember = "VesselID"



        cbCustSites.Text = ""
        'FillForm()

        If Not myClient Is Nothing Then         ' Check we have a client
            Dim blnCustSite As Boolean = False
            Dim objSite As Intranet.intranet.Address.CustomerSite = Nothing

            If Me.myClient.Sites.Count > 0 Then      'Has customer sites
                objSite = Me.myCollection.CustomerSite
                blnCustSite = Not (objSite Is Nothing)
                'Me.cbCustSites.SelectedItem = objSite
            End If

            Dim objPort As Intranet.intranet.jobs.Port = cSupport.UKPorts.Item(myCollection.Port)       'Set the site 
            Me.lblLocation.Text = "Location Port/Site"


        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get contenst of controls
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetForm() As Boolean
        Dim blnRet As Boolean = True
        With myCollection

            Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
            Me.ErrorProvider1.SetError(Me.txtPO, "")
            Me.ErrorProvider1.SetError(Me.txtContactNumber, "")
            Me.ErrorProvider1.SetError(Me.txtPinNumber, "")
            Me.ErrorProvider1.SetError(Me.txtAccessProblems, "")
            Me.ErrorProvider1.SetError(Me.txtCustomerContact, "")
            Me.ErrorProvider1.SetError(Me.txtLocation, "")
            Me.ErrorProvider1.SetError(Me.txtCoOfficer, "")
            Me.ErrorProvider1.SetError(Me.DTSentToOfficer, "")
            If Me.cbSiteVessel.SelectedIndex = 0 Then
                .IsSiteVisit = True
                '.vessel = Nothing
                'TODO see if we have a customer site
                'See if there is a customer site selected index > -1 
                If Me.blnCustSiteSelected Then
                    If Me.cbCustSites.SelectedIndex <= Me.cbCustSites.DataSource.count - 1 Then

                        Dim custSite As Intranet.intranet.Address.CustomerSite = _
                            Me.cbCustSites.DataSource.item(Me.cbCustSites.SelectedIndex)
                        .vessel = custSite.VesselID
                        '.Addressid = custSite.AddressId
                    End If
                End If
                Else
                    .IsVessel = True
                    Dim objVessel As Intranet.intranet.customerns.Vessel
                    objVessel = Me.BindingContext(cbVessel.DataSource).Current
                    If Not objVessel Is Nothing Then
                        .vessel = objVessel.VesselID                 'Set the vessel 
                    Else
                        If Not myClient Is Nothing Then
                            .vessel = Me.myClient.CompanyName
                        End If
                    End If
                End If

                Dim objPort As Intranet.intranet.jobs.Port
                'If Me.cbLocation.SelectedIndex > -1 Then
                '<Added code modified 11-Apr-2007 15:01 by LOGOS\taylor> 
                If Not objPort Is Nothing Then
                    .Port = objPort.Identity
                    If .Port.Trim.Length = 0 Then
                        Me.ErrorProvider1.SetError(Me.txtLocation, "Please select a valid port or create a customer site")
                        blnRet = False
                        Exit Function
                    End If
                    '</Added code modified 11-Apr-2007 15:01 by LOGOS\taylor> 

                    If objPort.IsUK Then
                        .CollectionType = MedscreenLib.Constants.GCST_JobTypeCallout
                        .UK = True
                    Else
                        .CollectionType = MedscreenLib.Constants.GCST_JobTypeCalloutOverseas
                        .UK = False
                    End If
                ElseIf Me.Collection.Port.Trim.Length > 0 Then
                    objPort = Intranet.intranet.Support.cSupport.Ports.Item(Me.Collection.Port)
                    If objPort.IsUK Then
                        .CollectionType = MedscreenLib.Constants.GCST_JobTypeCallout
                        .UK = True
                    Else
                        .CollectionType = MedscreenLib.Constants.GCST_JobTypeCalloutOverseas
                        .UK = False
                    End If
                Else
                    Me.ErrorProvider1.SetError(Me.txtLocation, "Please select a valid port or create a customer site")
                    blnRet = False
                End If

                'objPort = ModTpanel.Ports.Item(.Port)       'Set the site 
                'Me.cbLocation.SelectedItem = objPort
                '
                '
                .CollectionDate = Me.DTCollDate.Value       'Set the collection date 

                .NoOfDonors = Me.updnNoDonors.Value         'Set the expected number of Donors
                If myCollection.NoOfDonors = 0 Then
                    Me.ErrorProvider1.SetError(Me.updnNoDonors, "No of Donors must be greater than 1 ")
                    blnRet = False
                End If


                '.UK = True


                ' .WhoToTest = Me.txtWhototest.Text            'Indcate Who to Test 

                'Set which of the items is selected 
                Dim objPhrase As MedscreenLib.Glossary.Phrase
                'Dim i As Integer
                objPhrase = Me.BindingContext(Me.cbCategory.DataSource).Current()
                If Not objPhrase Is Nothing Then
                    .CallOutCategory = objPhrase.PhraseText
                    .CollType = "CALLOUT"
                End If


                .PurchaseOrder = Me.txtPO.Text                       'Set purchase Order 
                If myClient.InvoiceInfo.PoRequired Then
                    If .PurchaseOrder.Trim.Length = 0 Then
                        Me.ErrorProvider1.SetError(Me.txtPO, "Purchase Order Required for this customer")
                        blnRet = False
                    End If
                End If
                .CollectionComments = Me.txtCollComments.Text          'Set Comments about collection 


                'Customers comments 
                '.CustomerComment = Me.txtPrvComment.Text
                .ContactTel = Me.txtContactNumber.Text
                If .ContactTel.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtContactNumber, "Contact number must be provided")
                    blnRet = False
                End If
                .CallerName = Me.txtCallerName.Text
                If .CallerName.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtCallerName, "Caller's name must be provided")
                    blnRet = False
                End If

                .PinNumber = Me.txtPinNumber.Text       'Should be initialised from the client object
                If .PinNumber.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtPinNumber, "Pin number must be provided")
                    blnRet = False
                End If

                .AdditionalComments = Me.txtAccessProblems.Text
                If .AdditionalComments.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtAccessProblems, "Details must be provided")
                    blnRet = False
                End If

                .ContactOnSite = Me.txtCustomerContact.Text
                If .ContactOnSite.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.txtCustomerContact, "Details must be provided")
                    blnRet = False
                End If
                .RequestReceived = Me.DtCallTime.Value

                Dim i As Integer
                For i = 0 To Me.txtSiteAddress.Lines.Length - 1
                    If i = 0 Then .AddressLine1 = Me.txtSiteAddress.Lines.GetValue(i)
                    If i = 1 Then .AddressLine2 = Me.txtSiteAddress.Lines.GetValue(i)
                    If i = 2 Then .AddressLine3 = Me.txtSiteAddress.Lines.GetValue(i)
                    If i = 3 Then .AddressLine4 = Me.txtSiteAddress.Lines.GetValue(i)
                    If i = 4 Then .AddressLine5 = Me.txtSiteAddress.Lines.GetValue(i)
                Next

                'See if we can't display the site address.
                'ModTpanel.TestPanels.Item(Me.cbTestPanel.SelectedIndex)
                myCollection.DateCalloutToOfficer = Me.DTSentToOfficer.Value
                If myCollection.DateCalloutToOfficer < myCollection.DateReceived Then
                    'Can't be sent to the officer before the collection is done
                    Me.ErrorProvider1.SetError(Me.DTSentToOfficer, "Can't be sent to Officer berore call is received")
                End If
                If Me.myOfficer Is Nothing Then

                    Me.ErrorProvider1.SetError(Me.txtCoOfficer, "A collecting officer needs to be defined")
                    blnRet = False
                Else
                    Me.myCollection.CollOfficer = Me.myOfficer.OfficerID
                    Me.myCollection.CollectingOfficer = Me.myOfficer
                End If
                'If Me.CBCollectingOfficer.SelectedIndex > -1 Then
                '    Dim officer As Intranet.intranet.jobs.CollectingOfficer = _
                '        CType(Me.CBCollectingOfficer.DataSource, Intranet.intranet.jobs.CollectingOfficerCollection).Item( _
                '        Me.CBCollectingOfficer.SelectedIndex)
                '    If Not officer Is Nothing Then
                '        myCollection.CollOfficer = officer.OfficerID
                '    End If
                'End If
        End With

        'Type of collection 
        Me.ErrorProvider1.SetError(Me.cbRandom, "")
        If cbRandom.SelectedIndex > -1 Then
            myCollection.CollType = cbRandom.SelectedValue
            If myCollection.CollType.Trim.Length = 0 Then
                Me.ErrorProvider1.SetError(Me.cbRandom, "A reason for test must be given")
                blnRet = False
            End If
        Else
            Me.ErrorProvider1.SetError(Me.cbRandom, "A reason for test must be given")
            blnRet = False
        End If

        Me.ErrorProvider1.SetError(Me.cbSampleTemplate, "")
        myCollection.TestScheduleID = Me.cbSampleTemplate.SelectedValue
        'Check that we have a sample template 
        If myCollection.TestScheduleID.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.cbSampleTemplate, "A sample template must be provided")
            blnRet = False
        End If

        Return blnRet

    End Function

    'Update collection information and return 
    ''' <summary>
    ''' Created by taylor on ANDREW at 20/01/2007 11:27:54
    '''     Get call out form
    ''' </summary>
    ''' <param name="sender" type="Object">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <param name="e" type="EventArgs">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 20/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub cmdSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        '<EhHeader>
        Try
            '</EhHeader>
            If Me.Collection.Status = "T" Then blnFirstPass = True
            If blnFirstPass Then 'This should only get called if a new collection 
                blnSecondPass = False
                If Not GetForm() Then 'Check data is valid 
                    Me.DialogResult = DialogResult.None
                    Exit Sub
                End If
                Dim objNC As New frmCallOutCMNO()
                Dim dr As DialogResult
                objNC.DefineCMNumber = True
                objNC.DateTimePicker1.Value = Me.DtCallTime.Value
                'Me.DTSentToOfficer.Value = Me.DtCallTime.Value
                dr = objNC.ShowDialog
                Dim oldCMNumber As String = Me.Collection.ID
                If dr = DialogResult.Yes Then
                    Try
                        'Dim OISID As Integer
                        'OISID = CConnection.NextSequence("SEQ_CM_JOB_NUMBER")
                        If Me.Collection.ID.Trim.Length = 0 Then
                            Me.Collection.SetId = CConnection.NextSequence("SEQ_CM_JOB_NUMBER")
                        End If
                        oldCMNumber = Me.Collection.ID
                        Me.Collection.FieldValues.Insert(CConnection.DbConnection)
                        'Dim objC As Intranet.intranet.jobs.CMJob = cSupport.ICollection(False).CreateCollection(Me.Collection, OISID)
                        'If Not objC Is Nothing Then 'Check we managed to create one
                        '    Me.CopyCollection(objc)
                        'End If
                        'Me.txtCmNumber.Text = objc.ID

                    Catch ex As Exception
                        MedscreenLib.Medscreen.LogError(ex, , "Getting ID")

                    End Try
                ElseIf dr = DialogResult.No Then
                    '    'user provided cm number 
                    Try

                        If Not cSupport.ICollection(False).CollectionExists(objNC.UserCMNo) Then
                            Dim strTempId As String
                            If Me.Collection.ID.Trim.Length > 0 Then
                                strTempId = Me.Collection.ID

                            End If
                            'Dim objC As Intranet.intranet.jobs.CMJob = cSupport.ICollection(False).CreateCollection(myCollection.SmClient.Identity, objNC.UserCMNo)
                            'If Not objC Is Nothing Then 'Check we managed to create one
                            '    Me.CopyCollection(objC)
                            'End If
                            Me.txtCmNumber.Text = objNC.UserCMNo
                            Me.Collection.SetId = objNC.UserCMNo
                            myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
                            Me.Collection.Update()
                        Else
                            MsgBox("That collection " & objNC.UserCMNo & " already exists!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
                            Exit Sub
                        End If
                    Catch ex As Exception
                        MedscreenLib.Medscreen.LogError(ex, , "Getting user generated ID")

                    End Try
                Else
                    Me.DialogResult = DialogResult.None
                    Exit Sub        ' User wants to bog off.

                End If
                If Me.Collection Is Nothing Then Exit Sub
                Me.Collection.Update()
                Try

                    With myCollection
                        If Not myPort Is Nothing Then
                            If Me.myPort.IsUK Then
                                .CollectionType = Chr(.CollectionTypes.GCST_JobTypeCallout)
                            Else
                                .CollectionType = Chr(.CollectionTypes.GCST_JobTypeCalloutOverseas)

                            End If
                        End If
                        '
                        'myCollection.SetId = objC.ID
                        'myCollection.CollectionType = "C"
                        myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
                        If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
                            myCollection.Booker = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                        End If
                        'objC = myCollection
                        MsgBox("Collection ID: " & Collection.ID & " Created.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                        myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
                        '.DateReceived = Now


                        .Comments.CreateComment("Callout Collection created for customer " & .ClientId, MedscreenLib.Constants.GCST_COMM_NEWColl)
                        Me.Collection.Update()
                        If .IsVessel Then                           ' a vessel collection needs the vessel info sorted
                            If Not .VesselObject Is Nothing Then
                                .VesselObject.ChangeCollection(.ID)
                            End If
                        Else
                            If Not .SiteObject Is Nothing Then
                                .SiteObject.ChangeCollection(.ID)
                            End If
                        End If

                    End With
                    myCollection.Update()
                    Me.DisableForm()
                    Me.blnFirstPass = False
                    Me.blnSecondPass = True
                    myCollection.DateCalloutToOfficer = Me.DTSentToOfficer.Value
                    If Not myCollection.CollectingOfficer Is Nothing Then
                        'Collection.Status = Constants.GCST_JobStatusConfirmed    'Update the data 
                        myCollection.Update()
                        myCollection.Status = Constants.GCST_JobStatusSent
                        myCollection.Update()
                        myCollection.ConfirmCallOutCollection()       'Pass to SM 
                        myCollection.Comments.CreateComment("Confirmation received from " & myCollection.CollOfficer, MedscreenLib.Constants.GCST_COMM_CollConfirm)

                        myCollection.Update()

                        myCollection.Refresh()                   'Ensure any changes have been updated
                        myCollection.CallOutConfirmation()

                        Me.DialogResult = DialogResult.OK

                    End If
                    Me.Collection.WriteXML(Medscreen.GCSTCollectionXML & myCollection.ID & ".XML", 0)
                    Me.Collection.Update()

                    Exit Sub
                Catch ex As System.Exception
                    ' Handle exception here
                    MedscreenLib.Medscreen.LogError(ex, True, "updating collection")

                End Try
            Else    'Old Collection 
                If Me.Collection.ID <> 0 And Not blnSecondPass Then        'Indicates an existing collection collection
                    If Not GetForm() Then 'Check data is valid 
                        Me.DialogResult = DialogResult.None
                        Exit Sub
                    End If
                    myCollection.Update()
                    Exit Sub
                End If

            End If
            '<EhFooter>
        Catch ex As System.Exception
            ' Handle exception here
            MedscreenLib.Medscreen.LogError(ex, , "Exception log message")
        End Try '</EhFooter>
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 20/01/2007 11:51:09
    '''     
    '''     Copy the collection info
    ''' </summary>
    ''' <param name="objC" type="Intranet.intranet.jobs.CMJob">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 20/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub CopyCollection(ByVal objC As Intranet.intranet.jobs.CMJob)
        With myCollection
            objC.Load()
            myCollection.SetId = objC.ID
            myCollection.FieldValues.RowID = objC.FieldValues.RowID
            myCollection.FieldValues.InitialiseOldValues()
            myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
            If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
                myCollection.Booker = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            End If
            'objC = myCollection
            System.Threading.Thread.CurrentThread.Sleep(200)
            MsgBox("Collection ID: " & objC.ID & " Created.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
            myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
            .DateReceived = Now
            '.CollectionType = Constants.GCST_JobTypeCallout In get form 
            .Comments.CreateComment("Collection created for customer " & .ClientId, MedscreenLib.Constants.GCST_COMM_NEWColl)
            myCollection.Update()       'Ensure it gets written
        End With

    End Sub


    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        'Need to deal with cancelling collection
        If Not Me.Collection Is Nothing Then
            If Me.Collection.Status = Constants.GCST_JobStatusTemp Then
                'Check to see if user wants to keep the collection or delete it 
                'Dim iRet As DialogResult = MsgBox("Do you want to keep this collection to finish later (yes), or delete it (no)?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                'If iRet = DialogResult.No Then
                'Do deletetion code 
                Try
                    cSupport.ICollection(False).DeleteCollection(Me.Collection.ID)
                Catch
                End Try
                'Else'
                '    Me.Collection.CollectionType = Constants.GCST_JobTypeCallout
                ' End If
        End If
        End If
    End Sub

    Private Sub txtPinNumber_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtPinNumber.Validating
        If txtPinNumber.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtPinNumber, "Pin number must be provided")
        Else
            Me.ErrorProvider1.SetError(Me.txtPinNumber, "")
            If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
                myCollection.PinNumber = txtPinNumber.Text
                myCollection.Update()
            End If
        End If


    End Sub

    Private Sub txtCustomerContact_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerContact.Validating
        If txtCustomerContact.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtCustomerContact, "Contact must be provided")
        Else
            Me.ErrorProvider1.SetError(Me.txtCustomerContact, "")
        End If

    End Sub


    Private Sub txtAccessProblems_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAccessProblems.Validating
        If txtAccessProblems.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtAccessProblems, "Details must be provided")
        Else
            Me.ErrorProvider1.SetError(Me.txtAccessProblems, "")
            If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
                myCollection.AdditionalComments = txtAccessProblems.Text
                myCollection.Update()
            End If

        End If
    End Sub

    Private Sub txtSiteAddress_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSiteAddress.Validating
        If txtSiteAddress.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtSiteAddress, "Details must be provided")
            If Not Me.Collection Is Nothing Then
                Dim intLines As Integer = txtSiteAddress.Lines.Length
                Dim i As Integer
                For i = 0 To intLines - 1
                    If i = 0 Then Me.Collection.AddressLine1 = txtSiteAddress.Lines.GetValue(i)
                    If i = 1 Then Me.Collection.AddressLine2 = txtSiteAddress.Lines.GetValue(i)
                    If i = 2 Then Me.Collection.AddressLine3 = txtSiteAddress.Lines.GetValue(i)
                    If i = 3 Then Me.Collection.AddressLine4 = txtSiteAddress.Lines.GetValue(i)
                    If i = 4 Then Me.Collection.AddressLine5 = txtSiteAddress.Lines.GetValue(i)
                Next
                Me.Collection.Update()
            End If
        Else
            Me.ErrorProvider1.SetError(Me.txtSiteAddress, "")
        End If
    End Sub

    Private Sub txtContactNumber_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtContactNumber.Validating
        If txtContactNumber.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtContactNumber, "Contact number must be provided")
        Else
            Me.ErrorProvider1.SetError(Me.txtContactNumber, "")
            If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
                myCollection.ContactTel = txtContactNumber.Text
                myCollection.Update()
            End If
        End If
    End Sub

    Private Sub txtCallerName_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCallerName.Validating
        If txtCallerName.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtCallerName, "Caller's name must be provided")
        Else
            Me.ErrorProvider1.SetError(Me.txtCallerName, "")
            If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
                myCollection.CallerName = txtCallerName.Text
                myCollection.Update()
            End If

        End If
    End Sub

    Private Sub updnNoDonors_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles updnNoDonors.Validating
        If Me.updnNoDonors.Value = 0 Then
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "No of Donors must be greater than 1 ")
        Else
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
            If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
                myCollection.NoOfDonors = updnNoDonors.Value
                myCollection.Update()
            End If
        End If
    End Sub

    Private Sub SetWaitCursor()
        Me.Cursor = Cursors.WaitCursor
        Me.Invalidate()
        Application.DoEvents()
    End Sub

    Private Sub setDefaultCursor()
        Me.Cursor = Cursors.Default
        Me.Invalidate()
        Application.DoEvents()
    End Sub

    Private Sub FlipSites()
        If Me.cbSiteVessel.SelectedIndex = 0 Then
            Me.cbSiteVessel.SelectedIndex = 1
            Me.cbSiteVessel_SelectedIndexChanged(Nothing, Nothing)
            Me.cbSiteVessel.SelectedIndex = 0
            Me.cbSiteVessel_SelectedIndexChanged(Nothing, Nothing)
        Else
            Me.cbSiteVessel.SelectedIndex = 0
            Me.cbSiteVessel_SelectedIndexChanged(Nothing, Nothing)
            Me.cbSiteVessel.SelectedIndex = 1
            Me.cbSiteVessel_SelectedIndexChanged(Nothing, Nothing)
        End If
    End Sub


    Private Sub cmdEditSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditSite.Click

        If Me.myClient.Sites Is Nothing Then Exit Sub

        Dim custSite As Intranet.intranet.Address.CustomerSite

        Me.BindingContext(Me.cbCustSites.DataSource).Position = cbCustSites.SelectedIndex
        If Me.BindingContext(Me.cbCustSites.DataSource).Position = -1 Then Exit Sub
        custSite = Me.BindingContext(Me.cbCustSites.DataSource).Current
        If Not custSite Is Nothing Then
            Dim edFrm As New frmVessel()
            edFrm.Vessel = custSite

            If edFrm.ShowDialog = DialogResult.OK Then
                custSite = edFrm.Vessel
                custSite.Update()
                Me.cbCustSites.DataSource = myClient.Sites

                FlipSites()
                Me.cbCustSites.SelectedIndex = myClient.Sites.Count - 1
                'Dim cadr As MedscreenLib.Address.Caddress = _
                    'cSupport.SMAddresses.item(custSite.AddressId)
                'cmdSiteAddress.Enabled = (cadr Is Nothing)
                'If Not cadr Is Nothing Then
                '    ' Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
                'End If
                Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(custSite.VesselSubType)
                If Not oPort Is Nothing Then
                    Me.txtLocation.Text = oPort.Info2
                    Me.Collection.Port = oPort.Identity
                End If
            End If

        End If
    End Sub



    Private Sub cmdAddVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddVessel.Click
        Dim ovessel As Intranet.intranet.customerns.Vessel

        ovessel = cSupport.AddNewVessel(myClient.Identity)
        If Not ovessel Is Nothing Then
            SetWaitCursor()
            myClient.Vessels.Add(ovessel)
            Me.cbVessel.DataSource = myClient.Vessels
            Me.FlipSites()
        End If
        setDefaultCursor()
    End Sub

    Private Sub cbSiteVessel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSiteVessel.SelectedIndexChanged
        If blnDrawing Then Exit Sub
        If myClient Is Nothing Then Exit Sub
        If myCollection Is Nothing Then Exit Sub
        SetWaitCursor()
        If cbSiteVessel.SelectedIndex = 0 Then
            myCollection.IsVessel = False
            If Not Me.myClient.Sites Is Nothing Then
                Me.myClient.Sites.RefreshChanged()
#If r2004 Then
            myClient.Sites.Sort(Intranet.intranet.customerns.SMClientCollection.SortBy.CompanyName)
#Else
                myClient.Sites.Sort()
#End If
                Me.cbCustSites.DataSource = Me.myClient.Sites
            End If
            Me.cbCustSites.DisplayMember = "Info"
            Me.cbCustSites.ValueMember = "SiteID"
            Me.lblVessel.Text = "Customer Sites"

            Me.cbCustSites.SelectedIndex = -1
            Me.cbCustSites.Show()
            Me.cbCustSites.BringToFront()

            Me.cbVessel.SelectedIndex = -1
            Me.cbVessel.Hide()
            Me.cmdAddVessel.Hide()
            Me.cmdNewSite.Show()
            Me.cmdNewSite.BringToFront()

            Me.cmdEditSite.Show()
            Me.cmdEditSite.BringToFront()
        Else
            myCollection.IsVessel = True
            If Not Me.myClient.Vessels Is Nothing Then
                Me.myClient.Vessels.RefreshChanged()
#If r2004 Then
                Me.myClient.Vessels.Sort(Intranet.intranet.customerns.SMClientCollection.SortBy.CompanyName)
#Else
                Me.myClient.Vessels.Sort()
#End If
                Me.cbVessel.DataSource = Me.myClient.Vessels
                cbVessel.DisplayMember = "Info"
                cbVessel.ValueMember = "VesselID"
            End If
            Me.lblVessel.Text = "Vessel Name"
            Me.cbVessel.SelectedIndex = -1
            Me.cbCustSites.SelectedIndex = -1
            Me.cbCustSites.Hide()
            Me.cbVessel.Show()
            Me.cbVessel.BringToFront()
            Me.cmdAddVessel.Show()
            Me.cmdAddVessel.BringToFront()
            Me.cmdNewSite.Hide()
            Me.cmdEditSite.Hide()
            'Me.txtOther.Text = ""

        End If

        setDefaultCursor()
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = Constants.GCST_JobStatusTemp Then
            myCollection.Update()
        End If
    End Sub

    Private Sub cmdNewSite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNewSite.Click
        'TODO : Add new collection site code needed
        If myClient Is Nothing Then Exit Sub

        Dim oSite As Intranet.intranet.Address.CustomerSite
        SetWaitCursor()
        oSite = myClient.Sites.Create(myClient.Identity)
        oSite.VesselStatus = "ACTIVE"
        oSite.VeselNextTest = Today.AddYears(1)

        Dim edFrm As New frmVessel()
        edFrm.Vessel = oSite

        If edFrm.ShowDialog = DialogResult.OK Then
            oSite = edFrm.Vessel
            oSite.Update()
            Me.cbCustSites.DataSource = myClient.Sites

            FlipSites()
            Me.cbCustSites.SelectedIndex = myClient.Sites.Count - 1
            'Dim cadr As MedscreenLib.Address.Caddress = _
                'cSupport.SMAddresses.item(oSite.AddressId)
            'cmdSiteAddress.Enabled = (cadr Is Nothing)
            'If Not cadr Is Nothing Then
            '    'Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
            'End If
            Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(oSite.VesselSubType)
            If Not oPort Is Nothing Then
                Me.txtLocation.Text = oPort.Info2
                Me.Collection.Port = oPort.Identity
            End If
        End If
        Me.OfficerPortInfo()
        setDefaultCursor()
    End Sub




    Private Sub txtCallerName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCallerName.Leave
    End Sub

    Private Sub DTCollDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTCollDate.ValueChanged
        Me.DtCallTime.Value = DTCollDate.Value
        Me.DTSentToOfficer.Value = DTCollDate.Value.AddHours(3)
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
            myCollection.CollectionDate = DTCollDate.Value
            myCollection.Update()
        End If
        OfficerPortInfo()
    End Sub

    Private Sub OfficerPortInfo()
        If Me.Collection Is Nothing Then Exit Sub
        If Me.Collection.Port.Trim.Length = 0 Then Exit Sub
        Me.Panel2.Show()
        Me.txtAddress.Show()
        Dim Coll As New Collection()
        Coll.Add(Me.Collection.Port)
        'Coll.Add(Me.DTCollDate.Value.ToString("dd-MMM-yy"))

        Dim infoString As String = CConnection.PackageStringList("Lib_collection.GetPortOfficerCO", Coll)
        Me.txtAddress.Rtf = Medscreen.ReplaceString(infoString, Chr(10), vbCrLf)
    End Sub


    Private Sub DtCallTime_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtCallTime.ValueChanged
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
            myCollection.DateReceived = DtCallTime.Value
            myCollection.Update()
        End If

    End Sub

    Private Sub txtPO_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPO.Leave
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
            myCollection.PurchaseOrder = txtPO.Text
            myCollection.Update()
        End If
    End Sub

    Private Sub txtCollComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCollComments.TextChanged
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
            myCollection.CollectionComments = txtCollComments.Text
            myCollection.Update()
        End If
    End Sub

    Private Sub cbCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCategory.SelectedIndexChanged
        If blnDrawing Then Exit Sub
        If myCollection Is Nothing Then Exit Sub
        If myCollection.Status = MedscreenLib.Constants.GCST_JobStatusTemp Then
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            'Dim i As Integer
            objPhrase = Me.BindingContext(Me.cbCategory.DataSource).Current()
            If Not objPhrase Is Nothing Then
                myCollection.CallOutCategory = objPhrase.PhraseText
            End If
            myCollection.Update()
        End If
    End Sub

    Private blnCustSiteSelected As Boolean = False

    Private Sub cbCustSites_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCustSites.SelectedIndexChanged
        If blnDrawing Then Exit Sub
        If myCollection Is Nothing Then Exit Sub
        If myCollection.IsVessel Then Exit Sub
        If Me.myClient.Sites Is Nothing Then Exit Sub

        SetWaitCursor()
        Dim custSite As Intranet.intranet.Address.CustomerSite

        If cbCustSites.SelectedIndex = -1 Then Exit Sub
        custSite = Me.cbCustSites.DataSource.item(cbCustSites.SelectedIndex)
        If Not custSite Is Nothing Then
            If Not myCollection Is Nothing Then
                blnCustSiteSelected = True
                'myCollection.Addressid = custSite.AddressID
                'Dim cadr As MedscreenLib.Address.Caddress = _
                '    cSupport.SMAddresses.item(custSite.AddressID)
                'cmdSiteAddress.Enabled = (cadr Is Nothing)
                'If Not cadr Is Nothing Then
                '    Me.txtSiteAddress.Text = cadr.BlockAddress(vbCrLf)
                'End If
                Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(custSite.VesselSubType)
                If Not oPort Is Nothing Then 'And myCollection.Port.Trim.Length = 0 Then
                    Me.txtLocation.Text = oPort.Info2
                End If
                If Not myCollection Is Nothing Then
                    If myCollection.Status = Constants.GCST_JobStatusTemp Then
                        myCollection.vessel = custSite.VesselID
                        If Not oPort Is Nothing Then
                            myCollection.Port = oPort.Identity
                        End If
                        myCollection.Update()
                    End If
                End If
            End If
        End If
        Me.OfficerPortInfo()
        setDefaultCursor()
    End Sub


    Private Sub txtCallerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCallerName.TextChanged

    End Sub

    Private Sub txtCustomerContact_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerContact.Enter
        If myCollection Is Nothing Then Exit Sub
        If Me.txtCallerName.Text.Trim.Length > 0 And txtCustomerContact.Text.Length = 0 Then
            If Me.txtCustomerContact.Text.Trim.Length = 0 Then
                If MsgBox("Copy Caller to contact?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Me.txtCustomerContact.Text = txtCallerName.Text
                    myCollection.ContactOnSite = txtCallerName.Text
                    Me.ErrorProvider1.SetError(txtCustomerContact, "")
                End If
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change between UK and Overseas
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DomainUpDown1_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cBLocality.SelectedIndexChanged
        cBLocality.Enabled = False
        cBLocality.Enabled = True
    End Sub

    Private Sub mnuAddPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddPort.Click
        Dim aPort As Intranet.intranet.jobs.Port
        'aPort = ModTpanel.Ports.Item(PortId)
        'If Not aPort Is Nothing Then
        '    MsgBox("That Id already exists")
        '    Exit Sub
        'End If
        'aPort = New Intranet.intranet.jobs.Port(PortId, ModTpanel.Ports)
        Dim aFrm As New frmSelectPort()
        With aFrm
            .CountryID = 44
            If .ShowDialog = DialogResult.OK Then
                aPort = .Port
            End If
        End With
        If aPort Is Nothing Then Exit Sub

        Dim frmPort As MedscreenCommonGui.frmPort = New MedscreenCommonGui.frmPort()
        With frmPort
            .Port = aPort
            If .ShowDialog = DialogResult.OK Then

            End If
        End With



    End Sub


    Private Sub mnuEditPort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditPort.Click
        Dim aPort As Intranet.intranet.jobs.Port
        Dim aFrm As New frmSelectPort()
        With aFrm
            .CountryID = 44
            If .ShowDialog = DialogResult.OK Then
                aPort = .Port
            End If
        End With
        If aPort Is Nothing Then Exit Sub
        Dim frmPort As MedscreenCommonGui.frmPort = New MedscreenCommonGui.frmPort()
        With frmPort
            .Port = aPort
            If .ShowDialog = DialogResult.OK Then

            End If
        End With


  

    End Sub

    Private Sub cbLocation_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSiteAddress.Enter
        Me.pnlWarning.Show()
        Application.DoEvents()
    End Sub

    Private Sub cbLocation_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSiteAddress.Leave
        Me.pnlWarning.Hide()
        Application.DoEvents()
    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawing = False
    End Sub

    Private Sub CBCollectingOfficer_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OfficerPortInfo()
    End Sub

    Private Sub cmdFindPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindPort.Click
        Dim afrm As frmSelectPort = New frmSelectPort()
        'Set form location to anything typed in already 
        afrm.txtPort.Text = Me.txtLocation.Text
        afrm.FormMode = frmSelectPort.PortSelectionType.NoFixedSite
        afrm.DoFindPort()
        If afrm.ShowDialog = DialogResult.OK Then
            myPort = afrm.Port
            If Not myPort Is Nothing Then
                Me.txtLocation.Text = myPort.Info2
                Me.myCollection.Port = afrm.Port.Identity
                If afrm.Port.IsUK Then
                    Me.cBLocality.SelectedIndex = 0
                Else
                    Me.cBLocality.SelectedIndex = 1
                End If
                '<Added code modified 07-Aug-2007 06:20 by LOGOS\taylor>   Dealing with Tony's request for officer feedback              
                OfficerPortInfo()
                '</Added code modified 07-Aug-2007 06:20 by LOGOS\taylor> 

            End If
        End If
    End Sub

    Private myPort As Intranet.intranet.jobs.Port                   'Keep a copy of the port for later

    Private Sub txtLocation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLocation.TextChanged
        Dim strLoc As String = Me.txtLocation.Text
        If strLoc.Trim.Length > 0 Then
            Dim ocoll As New Collection()
            ocoll.Add(strLoc)
            ocoll.Add("T")
            Dim strReturn As String = CConnection.PackageStringList("LIB_CCTOOL.PORTSBYID", ocoll)
            If strReturn.Trim.Length > 0 Then 'Found at least one port 
                Dim strPorts As String() = strReturn.Split(New Char() {","})
                If strPorts.Length = 2 Then 'Found just one port 
                    Dim strPortId As String = strPorts.GetValue(0)
                    myPort = New Intranet.intranet.jobs.Port(strPortId)         'Save the port for later
                    If Not myPort Is Nothing Then                               'Display port
                        Me.txtLocation.Text = myPort.Info2
                        Me.myCollection.Port = myPort.Identity
                        If myPort.IsUK Then
                            Me.cBLocality.SelectedIndex = 0
                        Else
                            Me.cBLocality.SelectedIndex = 1
                        End If
                        'Display info about ports and officers
                        Me.OfficerPortInfo()
                    End If

                End If
            End If
        End If
    End Sub

    Private Sub txtCoOfficer_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCoOfficer.TextChanged
        Dim strFind As String = CConnection.PackageStringList("LIB_CCTOOL.FINDOFFICER", Me.txtCoOfficer.Text)
        If strFind.Trim.Length > 0 Then 'See how many found
            Dim strOfficers As String() = strFind.Split(New Char() {","})
            If strOfficers.Length = 1 Then
                Dim strOfficer As String = strOfficers.GetValue(0)
                Me.myOfficer = cSupport.CollectingOfficers.Item(strOfficer)
                If Not myOfficer Is Nothing Then
                    Me.txtCoOfficer.Text = Me.myOfficer.Info2
                End If
            End If
        End If
    End Sub

    Private Sub cmdFindOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindOfficer.Click
        Dim objCS As New frmSelectCollector()

        objCS.OfficerFindText = Me.txtCoOfficer.Text
        objCS.DoOfficerFind()
        If objCS.ShowDialog = DialogResult.OK Then
            Me.myOfficer = objCS.Officer
            If Not myOfficer Is Nothing Then
                Me.txtCoOfficer.Text = Me.myOfficer.Info2
            End If

        End If
    End Sub

    '    Private Sub cbSampleTemplate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSampleTemplate.SelectedIndexChanged
    '        'Need to get matrices and sample login template
    '        If Me.myClient Is Nothing Then Exit Sub
    '        Try  ' Protecting
    '            Dim strTempId As String = Me.cbSampleTemplate.SelectedValue

    '            Dim objReason = New Glossary.PhraseCollection("lib_cctool.REASONFORTESTLISTROUTINE", Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , Me.myClient.Identity & "," & strTempId)
    '            Me.cbRandom.DataSource = objReason
    '            Me.cbRandom.DisplayMember = "PhraseText"
    '            Me.cbRandom.ValueMember = "PhraseID"
    '            'need to make sure we update the combo 
    '            Try  ' Protecting

    '                If Not Me.myCollection Is Nothing Then
    '                    Me.cbRandom.SelectedValue = Me.myCollection.CollType
    '                End If
    '            Catch ex As Exception
    '                MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5214")
    '            Finally
    '            End Try
    '#If r2004 Then
    '            Me.txtMatrix.Text = myClient.CustomerMatrices.MatrixText(strTempId)
    '#End If
    '        Catch ex As Exception
    '            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5202")
    '        Finally
    '        End Try
    '    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' View test panel for customer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdViewPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViewPanel.Click
        If Not myClient.DefaultPanel Is Nothing Then
            Medscreen.TextEditor(myClient.DefaultPanel.ToText(True), True)
        Else
            Dim tp As Intranet.intranet.TestPanels.TestPanel = cSupport.TestPanels.Item(Me.cbSampleTemplate.SelectedIndex)
            If Not tp Is Nothing Then
                Medscreen.TextEditor(tp.ToText(True), True)
            End If
        End If
#If r2004 Then
        Dim strXML As String = myClient.CustomerMatrices.ToXML
        Dim strText As String = Medscreen.ResolveStyleSheet(strXML, "custmatrixtext.xsl", 0)
        Medscreen.ShowHtml(strText)
#End If
    End Sub

    Private Sub cbSampleTemplate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSampleTemplate.SelectedIndexChanged
        If Not myClient Is Nothing Then
            Try  ' Protecting
                Dim strTempId As String = Me.cbSampleTemplate.SelectedValue

                Dim strReasonForTestCallOut As String = MedscreenCommonGUIConfig.CollectionQuery.Item("ReasonForTestCallOut")
                Dim objReason = New Glossary.PhraseCollection(strReasonForTestCallOut, Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , Me.myClient.Identity & "," & strTempId)
                Me.cbRandom.DataSource = objReason
                Me.cbRandom.DisplayMember = "PhraseText"
                Me.cbRandom.ValueMember = "PhraseID"
                'need to make sure we update the combo 
                Try  ' Protecting

                    If Not Me.myClient.CollectionInfo Is Nothing Then
                        Me.cbRandom.SelectedValue = Me.myClient.CollectionInfo.ReasonForTest
                    End If
                Catch ex As Exception
                    MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5214")
                Finally

                End Try

                Me.txtMatrix.Text = myClient.MatrixText(strTempId)

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5202")
            Finally
            End Try

        End If
    End Sub

    Private Sub cbRandom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRandom.SelectedIndexChanged

    End Sub
End Class

