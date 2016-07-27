'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 17:43:55+00 $
'$Log: frmPort.vb,v $
'Revision 1.0  2005-11-29 17:43:55+00  taylor
'Commented
'
'Revision 1.2  2005-06-17 07:05:14+01  taylor
'Milestone check in
'
'Revision 1.1  2004-05-06 14:42:45+01  taylor
'<>
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
Imports MedscreenLib.Constants
Imports Intranet.intranet.support
Imports Intranet.intranet.glossary
Imports Intranet.intranet
Imports System.Windows.Forms
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmPort
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for managing Ports
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmPort
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.cbCountry.DataSource = MedscreenLib.Glossary.Glossary.Countries
        Me.cbCountry.DisplayMember = "CountryName"
        Me.cbCountry.ValueMember = "CountryId"

        Me.CBRegion.DataSource = RegionList
        Me.CBRegion.DisplayMember = "Info"
        Me.CBRegion.ValueMember = "PortId"

        Me.TypeList = New MedscreenLib.Glossary.PhraseCollection("OWNERTYPE")
        TypeList.Load()
        Me.cbOwnerType.DataSource = TypeList
        Me.cbOwnerType.DisplayMember = "PhraseText"
        Me.cbOwnerType.ValueMember = "PhraseID"
        SmidList = New MedscreenLib.Glossary.PhraseCollection("Select distinct SMID,SMID from customer", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query, "Removeflag = 'F'")
        Dim aPhrase As MedscreenLib.Glossary.Phrase = New MedscreenLib.Glossary.Phrase()
        aPhrase.PhraseID = "NONE"
        aPhrase.PhraseText = ""
        SmidList.Add(aPhrase)
        Me.cbOwner.DataSource = SmidList
        Me.cbOwner.DisplayMember = "PhraseText"
        Me.cbOwner.ValueMember = "PhraseID"

        blnDrawn = False
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabDetails As System.Windows.Forms.TabPage
    Friend WithEvents TabAddress As System.Windows.Forms.TabPage
    Friend WithEvents adrPanel As System.Windows.Forms.Panel
    Friend WithEvents AddressPanel1 As MedscreenCommonGui.AddressPanel
    Friend WithEvents lblPortId As System.Windows.Forms.Label
    Friend WithEvents txtPortID As System.Windows.Forms.TextBox
    Friend WithEvents ckFixedSite As System.Windows.Forms.CheckBox
    Friend WithEvents imgListCk As System.Windows.Forms.ImageList
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtMinCharge As System.Windows.Forms.TextBox
    Friend WithEvents lblMinCharge As System.Windows.Forms.Label
    Friend WithEvents txtSiteCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabLocation As System.Windows.Forms.TabPage
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtMapPath As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents tabOpening As System.Windows.Forms.TabPage
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ctxmnuTimes As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddTime As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveTime As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdTime As System.Windows.Forms.MenuItem
    Friend WithEvents PnlTimeEdit As System.Windows.Forms.Panel
    Friend WithEvents lblday As System.Windows.Forms.Label
    Friend WithEvents CbDays As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdAddTime As System.Windows.Forms.Button
    Friend WithEvents ckSpecialNeeds As System.Windows.Forms.CheckBox
    Friend WithEvents txtSpecialNeeds As System.Windows.Forms.TextBox
    Friend WithEvents ckExtraCharges As System.Windows.Forms.CheckBox
    Friend WithEvents TabCO As System.Windows.Forms.TabPage
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgYesNo As System.Windows.Forms.ImageList
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents ctxOfficers As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents CBRegion As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents txtISOPortCode As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CBType As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ckRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIsoRegion As System.Windows.Forms.TextBox
    Friend WithEvents txtPortFunction As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtCoordinates As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents mnuEditPortOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents upDnInterval As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents pnlOwnedSite As System.Windows.Forms.Panel
    Friend WithEvents lbOwner As System.Windows.Forms.Label
    Friend WithEvents cbOwnerType As System.Windows.Forms.ComboBox
    Friend WithEvents cbOwner As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lkDirections As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdSearchAddress As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPort))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabDetails = New System.Windows.Forms.TabPage()
        Me.pnlOwnedSite = New System.Windows.Forms.Panel()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lbOwner = New System.Windows.Forms.Label()
        Me.cbOwnerType = New System.Windows.Forms.ComboBox()
        Me.cbOwner = New System.Windows.Forms.ComboBox()
        Me.txtPortFunction = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtIsoRegion = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ckRemoved = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CBType = New System.Windows.Forms.ComboBox()
        Me.txtISOPortCode = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.CBRegion = New System.Windows.Forms.ComboBox()
        Me.cbCountry = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ckExtraCharges = New System.Windows.Forms.CheckBox()
        Me.txtSpecialNeeds = New System.Windows.Forms.TextBox()
        Me.ckSpecialNeeds = New System.Windows.Forms.CheckBox()
        Me.txtSiteCharge = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMinCharge = New System.Windows.Forms.TextBox()
        Me.lblMinCharge = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.ckFixedSite = New System.Windows.Forms.CheckBox()
        Me.imgListCk = New System.Windows.Forms.ImageList(Me.components)
        Me.txtPortID = New System.Windows.Forms.TextBox()
        Me.lblPortId = New System.Windows.Forms.Label()
        Me.TabLocation = New System.Windows.Forms.TabPage()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.txtCoordinates = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lkDirections = New System.Windows.Forms.LinkLabel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtMapPath = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.TabAddress = New System.Windows.Forms.TabPage()
        Me.adrPanel = New System.Windows.Forms.Panel()
        Me.AddressPanel1 = New MedscreenCommonGui.AddressPanel()
        Me.cmdSearchAddress = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.tabOpening = New System.Windows.Forms.TabPage()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.upDnInterval = New System.Windows.Forms.NumericUpDown()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.PnlTimeEdit = New System.Windows.Forms.Panel()
        Me.cmdAddTime = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.DTEnd = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.DtStart = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CbDays = New System.Windows.Forms.ComboBox()
        Me.lblday = New System.Windows.Forms.Label()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ctxmnuTimes = New System.Windows.Forms.ContextMenu()
        Me.mnuAddTime = New System.Windows.Forms.MenuItem()
        Me.mnuRemoveTime = New System.Windows.Forms.MenuItem()
        Me.mnuEdTime = New System.Windows.Forms.MenuItem()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.TabCO = New System.Windows.Forms.TabPage()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader()
        Me.ctxOfficers = New System.Windows.Forms.ContextMenu()
        Me.mnuAddOfficer = New System.Windows.Forms.MenuItem()
        Me.mnuRemoveOfficer = New System.Windows.Forms.MenuItem()
        Me.mnuEditPortOfficer = New System.Windows.Forms.MenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.imgYesNo = New System.Windows.Forms.ImageList(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabDetails.SuspendLayout()
        Me.pnlOwnedSite.SuspendLayout()
        Me.TabLocation.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabAddress.SuspendLayout()
        Me.adrPanel.SuspendLayout()
        Me.AddressPanel1.SuspendLayout()
        Me.tabOpening.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.upDnInterval, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlTimeEdit.SuspendLayout()
        Me.TabCO.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 469)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(736, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(648, 8)
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
        Me.cmdOk.Location = New System.Drawing.Point(568, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabDetails, Me.TabLocation, Me.TabAddress, Me.tabOpening, Me.TabCO})
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(736, 469)
        Me.TabControl1.TabIndex = 1
        '
        'TabDetails
        '
        Me.TabDetails.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlOwnedSite, Me.txtPortFunction, Me.Label11, Me.txtIsoRegion, Me.Label10, Me.ckRemoved, Me.Label9, Me.CBType, Me.txtISOPortCode, Me.Label8, Me.Label7, Me.CBRegion, Me.cbCountry, Me.Label6, Me.ckExtraCharges, Me.txtSpecialNeeds, Me.ckSpecialNeeds, Me.txtSiteCharge, Me.Label1, Me.txtMinCharge, Me.lblMinCharge, Me.txtName, Me.lblName, Me.txtComments, Me.lblComments, Me.txtDescription, Me.lblDescription, Me.ckFixedSite, Me.txtPortID, Me.lblPortId})
        Me.TabDetails.Location = New System.Drawing.Point(4, 22)
        Me.TabDetails.Name = "TabDetails"
        Me.TabDetails.Size = New System.Drawing.Size(728, 443)
        Me.TabDetails.TabIndex = 0
        Me.TabDetails.Text = "Details"
        '
        'pnlOwnedSite
        '
        Me.pnlOwnedSite.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnedSite.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label15, Me.lbOwner, Me.cbOwnerType, Me.cbOwner})
        Me.pnlOwnedSite.Location = New System.Drawing.Point(576, 8)
        Me.pnlOwnedSite.Name = "pnlOwnedSite"
        Me.pnlOwnedSite.Size = New System.Drawing.Size(144, 136)
        Me.pnlOwnedSite.TabIndex = 249
        Me.pnlOwnedSite.Visible = False
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(16, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(104, 23)
        Me.Label15.TabIndex = 252
        Me.Label15.Text = "Fixed Site Type"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbOwner
        '
        Me.lbOwner.Location = New System.Drawing.Point(12, 14)
        Me.lbOwner.Name = "lbOwner"
        Me.lbOwner.Size = New System.Drawing.Size(64, 23)
        Me.lbOwner.TabIndex = 251
        Me.lbOwner.Text = "Owner"
        Me.lbOwner.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbOwnerType
        '
        Me.cbOwnerType.Items.AddRange(New Object() {"D - Continetal Region", "C - Country", "R - Region of Country", "P - Port", "S - Site "})
        Me.cbOwnerType.Location = New System.Drawing.Point(12, 94)
        Me.cbOwnerType.Name = "cbOwnerType"
        Me.cbOwnerType.Size = New System.Drawing.Size(121, 21)
        Me.cbOwnerType.TabIndex = 250
        '
        'cbOwner
        '
        Me.cbOwner.Items.AddRange(New Object() {"D - Continetal Region", "C - Country", "R - Region of Country", "P - Port", "S - Site "})
        Me.cbOwner.Location = New System.Drawing.Point(12, 46)
        Me.cbOwner.Name = "cbOwner"
        Me.cbOwner.Size = New System.Drawing.Size(121, 21)
        Me.cbOwner.TabIndex = 249
        '
        'txtPortFunction
        '
        Me.txtPortFunction.Enabled = False
        Me.txtPortFunction.Location = New System.Drawing.Point(384, 280)
        Me.txtPortFunction.Name = "txtPortFunction"
        Me.txtPortFunction.Size = New System.Drawing.Size(328, 20)
        Me.txtPortFunction.TabIndex = 245
        Me.txtPortFunction.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(280, 280)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 244
        Me.Label11.Text = "Port Function"
        '
        'txtIsoRegion
        '
        Me.txtIsoRegion.Enabled = False
        Me.txtIsoRegion.Location = New System.Drawing.Point(376, 104)
        Me.txtIsoRegion.Name = "txtIsoRegion"
        Me.txtIsoRegion.Size = New System.Drawing.Size(184, 20)
        Me.txtIsoRegion.TabIndex = 243
        Me.txtIsoRegion.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(272, 104)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 242
        Me.Label10.Text = "Iso Region"
        '
        'ckRemoved
        '
        Me.ckRemoved.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckRemoved.Location = New System.Drawing.Point(42, 104)
        Me.ckRemoved.Name = "ckRemoved"
        Me.ckRemoved.TabIndex = 241
        Me.ckRemoved.Text = "Removed"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(256, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(64, 23)
        Me.Label9.TabIndex = 240
        Me.Label9.Text = "type"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CBType
        '
        Me.CBType.Items.AddRange(New Object() {"D - Continetal Region", "C - Country", "R - Region of Country", "P - Port", "S - Site "})
        Me.CBType.Location = New System.Drawing.Point(328, 16)
        Me.CBType.Name = "CBType"
        Me.CBType.Size = New System.Drawing.Size(121, 21)
        Me.CBType.TabIndex = 239
        '
        'txtISOPortCode
        '
        Me.txtISOPortCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtISOPortCode.Location = New System.Drawing.Point(80, 48)
        Me.txtISOPortCode.Name = "txtISOPortCode"
        Me.txtISOPortCode.Size = New System.Drawing.Size(159, 20)
        Me.txtISOPortCode.TabIndex = 238
        Me.txtISOPortCode.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 23)
        Me.Label8.TabIndex = 237
        Me.Label8.Text = "ISO Port Code:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(272, 74)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 236
        Me.Label7.Text = "Region"
        '
        'CBRegion
        '
        Me.CBRegion.Location = New System.Drawing.Point(376, 74)
        Me.CBRegion.Name = "CBRegion"
        Me.CBRegion.Size = New System.Drawing.Size(192, 21)
        Me.CBRegion.TabIndex = 235
        '
        'cbCountry
        '
        Me.cbCountry.Location = New System.Drawing.Point(376, 48)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(192, 21)
        Me.cbCountry.TabIndex = 234
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(272, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 233
        Me.Label6.Text = "Country"
        '
        'ckExtraCharges
        '
        Me.ckExtraCharges.Location = New System.Drawing.Point(192, 240)
        Me.ckExtraCharges.Name = "ckExtraCharges"
        Me.ckExtraCharges.Size = New System.Drawing.Size(200, 32)
        Me.ckExtraCharges.TabIndex = 18
        Me.ckExtraCharges.Text = "On ports list"
        '
        'txtSpecialNeeds
        '
        Me.txtSpecialNeeds.Location = New System.Drawing.Point(80, 312)
        Me.txtSpecialNeeds.Multiline = True
        Me.txtSpecialNeeds.Name = "txtSpecialNeeds"
        Me.txtSpecialNeeds.Size = New System.Drawing.Size(632, 120)
        Me.txtSpecialNeeds.TabIndex = 17
        Me.txtSpecialNeeds.Text = ""
        '
        'ckSpecialNeeds
        '
        Me.ckSpecialNeeds.Location = New System.Drawing.Point(16, 280)
        Me.ckSpecialNeeds.Name = "ckSpecialNeeds"
        Me.ckSpecialNeeds.Size = New System.Drawing.Size(199, 24)
        Me.ckSpecialNeeds.TabIndex = 16
        Me.ckSpecialNeeds.Text = "Special requirements for port"
        '
        'txtSiteCharge
        '
        Me.txtSiteCharge.Location = New System.Drawing.Point(512, 248)
        Me.txtSiteCharge.Name = "txtSiteCharge"
        Me.txtSiteCharge.Size = New System.Drawing.Size(80, 20)
        Me.txtSiteCharge.TabIndex = 15
        Me.txtSiteCharge.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(416, 248)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 23)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Site Charge:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMinCharge
        '
        Me.txtMinCharge.Location = New System.Drawing.Point(104, 248)
        Me.txtMinCharge.Name = "txtMinCharge"
        Me.txtMinCharge.Size = New System.Drawing.Size(80, 20)
        Me.txtMinCharge.TabIndex = 13
        Me.txtMinCharge.Text = ""
        '
        'lblMinCharge
        '
        Me.lblMinCharge.Location = New System.Drawing.Point(8, 248)
        Me.lblMinCharge.Name = "lblMinCharge"
        Me.lblMinCharge.Size = New System.Drawing.Size(96, 23)
        Me.lblMinCharge.TabIndex = 12
        Me.lblMinCharge.Text = "Minimum Charge:"
        Me.lblMinCharge.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(88, 74)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(136, 20)
        Me.txtName.TabIndex = 11
        Me.txtName.Text = ""
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(16, 74)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(64, 23)
        Me.lblName.TabIndex = 10
        Me.lblName.Text = "Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(80, 152)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(480, 80)
        Me.txtComments.TabIndex = 9
        Me.txtComments.Text = ""
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(8, 152)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(64, 23)
        Me.lblComments.TabIndex = 8
        Me.lblComments.Text = "Comments:"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(80, 128)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(480, 20)
        Me.txtDescription.TabIndex = 7
        Me.txtDescription.Text = ""
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(8, 128)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(64, 23)
        Me.lblDescription.TabIndex = 6
        Me.lblDescription.Text = "Description:"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckFixedSite
        '
        Me.ckFixedSite.Appearance = System.Windows.Forms.Appearance.Button
        Me.ckFixedSite.Image = CType(resources.GetObject("ckFixedSite.Image"), System.Drawing.Bitmap)
        Me.ckFixedSite.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ckFixedSite.ImageIndex = 0
        Me.ckFixedSite.ImageList = Me.imgListCk
        Me.ckFixedSite.Location = New System.Drawing.Point(464, 15)
        Me.ckFixedSite.Name = "ckFixedSite"
        Me.ckFixedSite.TabIndex = 4
        Me.ckFixedSite.Text = "Fixed Site"
        Me.ckFixedSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'imgListCk
        '
        Me.imgListCk.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgListCk.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgListCk.ImageStream = CType(resources.GetObject("imgListCk.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgListCk.TransparentColor = System.Drawing.Color.Transparent
        '
        'txtPortID
        '
        Me.txtPortID.Location = New System.Drawing.Point(80, 17)
        Me.txtPortID.Name = "txtPortID"
        Me.txtPortID.Size = New System.Drawing.Size(159, 20)
        Me.txtPortID.TabIndex = 1
        Me.txtPortID.Text = ""
        '
        'lblPortId
        '
        Me.lblPortId.Location = New System.Drawing.Point(8, 16)
        Me.lblPortId.Name = "lblPortId"
        Me.lblPortId.Size = New System.Drawing.Size(64, 23)
        Me.lblPortId.TabIndex = 0
        Me.lblPortId.Text = "Port ID:"
        Me.lblPortId.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabLocation
        '
        Me.TabLocation.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.HtmlEditor1})
        Me.TabLocation.Location = New System.Drawing.Point(4, 22)
        Me.TabLocation.Name = "TabLocation"
        Me.TabLocation.Size = New System.Drawing.Size(728, 443)
        Me.TabLocation.TabIndex = 2
        Me.TabLocation.Text = "Location"
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.LinkLabel1, Me.txtCoordinates, Me.Label12, Me.lkDirections, Me.Label3, Me.txtMapPath, Me.Label2})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(728, 72)
        Me.Panel2.TabIndex = 9
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(344, 48)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(272, 23)
        Me.LinkLabel1.TabIndex = 246
        '
        'txtCoordinates
        '
        Me.txtCoordinates.Location = New System.Drawing.Point(128, 40)
        Me.txtCoordinates.Name = "txtCoordinates"
        Me.txtCoordinates.Size = New System.Drawing.Size(184, 20)
        Me.txtCoordinates.TabIndex = 245
        Me.txtCoordinates.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtCoordinates, "Enter longitude and latidude or first part of postcode, latitude etc can be enter" & _
        "ed as degrees minutes and seconds or signed decimal degrees")
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(24, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 244
        Me.Label12.Text = "Coordinates"
        '
        'lkDirections
        '
        Me.lkDirections.Location = New System.Drawing.Point(448, 12)
        Me.lkDirections.Name = "lkDirections"
        Me.lkDirections.Size = New System.Drawing.Size(192, 23)
        Me.lkDirections.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(360, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Coordinates:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMapPath
        '
        Me.txtMapPath.Location = New System.Drawing.Point(112, 8)
        Me.txtMapPath.Name = "txtMapPath"
        Me.txtMapPath.Size = New System.Drawing.Size(208, 20)
        Me.txtMapPath.TabIndex = 9
        Me.txtMapPath.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "County/State/etc"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        Me.HtmlEditor1.Size = New System.Drawing.Size(728, 443)
        Me.HtmlEditor1.TabIndex = 8
        'Me.HtmlEditor1.Text = "HtmlEditor1"
        '
        'TabAddress
        '
        Me.TabAddress.Controls.AddRange(New System.Windows.Forms.Control() {Me.adrPanel})
        Me.TabAddress.Location = New System.Drawing.Point(4, 22)
        Me.TabAddress.Name = "TabAddress"
        Me.TabAddress.Size = New System.Drawing.Size(728, 443)
        Me.TabAddress.TabIndex = 1
        Me.TabAddress.Text = "Address"
        '
        'adrPanel
        '
        Me.adrPanel.Controls.AddRange(New System.Windows.Forms.Control() {Me.AddressPanel1})
        Me.adrPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.adrPanel.Name = "adrPanel"
        Me.adrPanel.Size = New System.Drawing.Size(728, 443)
        Me.adrPanel.TabIndex = 0
        '
        'AddressPanel1
        '
        Me.AddressPanel1.Address = Nothing
        Me.AddressPanel1.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AddressPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AddressPanel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdSearchAddress, Me.cmdSave, Me.Button2})
        Me.AddressPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AddressPanel1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AddressPanel1.InfoVisible = False
        Me.AddressPanel1.isReadOnly = False
        Me.AddressPanel1.Name = "AddressPanel1"
        Me.AddressPanel1.ShowCancelCmd = False
        Me.AddressPanel1.ShowCloneCmd = False
        Me.AddressPanel1.ShowContact = True
        Me.AddressPanel1.ShowFindCmd = False
        Me.AddressPanel1.ShowNewCmd = False
        Me.AddressPanel1.ShowPanel = False
        Me.AddressPanel1.ShowSaveCmd = False
        Me.AddressPanel1.ShowStatus = False
        Me.AddressPanel1.Size = New System.Drawing.Size(728, 443)
        Me.AddressPanel1.TabIndex = 33
        Me.AddressPanel1.Tooltip = Nothing
        '
        'cmdSearchAddress
        '
        Me.cmdSearchAddress.Image = CType(resources.GetObject("cmdSearchAddress.Image"), System.Drawing.Bitmap)
        Me.cmdSearchAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSearchAddress.Location = New System.Drawing.Point(320, 256)
        Me.cmdSearchAddress.Name = "cmdSearchAddress"
        Me.cmdSearchAddress.Size = New System.Drawing.Size(104, 23)
        Me.cmdSearchAddress.TabIndex = 232
        Me.cmdSearchAddress.Text = "Find Address"
        Me.cmdSearchAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdSave
        '
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Bitmap)
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(192, 256)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(112, 23)
        Me.cmdSave.TabIndex = 31
        Me.cmdSave.Text = "Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Bitmap)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(56, 256)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 30
        Me.Button2.Text = "New Address"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabOpening
        '
        Me.tabOpening.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel3})
        Me.tabOpening.Location = New System.Drawing.Point(4, 22)
        Me.tabOpening.Name = "tabOpening"
        Me.tabOpening.Size = New System.Drawing.Size(728, 443)
        Me.tabOpening.TabIndex = 3
        Me.tabOpening.Text = "Opening Times"
        '
        'Panel3
        '
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label14, Me.upDnInterval, Me.Label13, Me.PnlTimeEdit, Me.TreeView1, Me.Panel4})
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(728, 443)
        Me.Panel3.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(472, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "mins"
        '
        'upDnInterval
        '
        Me.upDnInterval.Location = New System.Drawing.Point(408, 8)
        Me.upDnInterval.Name = "upDnInterval"
        Me.upDnInterval.Size = New System.Drawing.Size(64, 20)
        Me.upDnInterval.TabIndex = 12
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(304, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.TabIndex = 11
        Me.Label13.Text = "Collection Interval"
        '
        'PnlTimeEdit
        '
        Me.PnlTimeEdit.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAddTime, Me.cmdUpdate, Me.DTEnd, Me.Label5, Me.DtStart, Me.Label4, Me.CbDays, Me.lblday})
        Me.PnlTimeEdit.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlTimeEdit.Location = New System.Drawing.Point(288, 43)
        Me.PnlTimeEdit.Name = "PnlTimeEdit"
        Me.PnlTimeEdit.Size = New System.Drawing.Size(440, 400)
        Me.PnlTimeEdit.TabIndex = 2
        Me.PnlTimeEdit.Visible = False
        '
        'cmdAddTime
        '
        Me.cmdAddTime.Image = CType(resources.GetObject("cmdAddTime.Image"), System.Drawing.Bitmap)
        Me.cmdAddTime.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddTime.Location = New System.Drawing.Point(208, 120)
        Me.cmdAddTime.Name = "cmdAddTime"
        Me.cmdAddTime.TabIndex = 7
        Me.cmdAddTime.Text = "Add"
        Me.cmdAddTime.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAddTime.Visible = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Image = CType(resources.GetObject("cmdUpdate.Image"), System.Drawing.Bitmap)
        Me.cmdUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdate.Location = New System.Drawing.Point(208, 80)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.TabIndex = 6
        Me.cmdUpdate.Text = "Update"
        Me.cmdUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdUpdate.Visible = False
        '
        'DTEnd
        '
        Me.DTEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.DTEnd.Location = New System.Drawing.Point(96, 120)
        Me.DTEnd.Name = "DTEnd"
        Me.DTEnd.ShowUpDown = True
        Me.DTEnd.Size = New System.Drawing.Size(80, 20)
        Me.DTEnd.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 23)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "End time:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DtStart
        '
        Me.DtStart.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.DtStart.Location = New System.Drawing.Point(96, 80)
        Me.DtStart.Name = "DtStart"
        Me.DtStart.ShowUpDown = True
        Me.DtStart.Size = New System.Drawing.Size(80, 20)
        Me.DtStart.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Start time:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CbDays
        '
        Me.CbDays.Enabled = False
        Me.CbDays.Items.AddRange(New Object() {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"})
        Me.CbDays.Location = New System.Drawing.Point(96, 32)
        Me.CbDays.Name = "CbDays"
        Me.CbDays.Size = New System.Drawing.Size(121, 21)
        Me.CbDays.TabIndex = 1
        '
        'lblday
        '
        Me.lblday.Location = New System.Drawing.Point(48, 32)
        Me.lblday.Name = "lblday"
        Me.lblday.Size = New System.Drawing.Size(32, 23)
        Me.lblday.TabIndex = 0
        Me.lblday.Text = "Day:"
        Me.lblday.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TreeView1
        '
        Me.TreeView1.ContextMenu = Me.ctxmnuTimes
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Left
        Me.TreeView1.ImageList = Me.imgListCk
        Me.TreeView1.Location = New System.Drawing.Point(64, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(224, 443)
        Me.TreeView1.TabIndex = 1
        '
        'ctxmnuTimes
        '
        Me.ctxmnuTimes.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddTime, Me.mnuRemoveTime, Me.mnuEdTime})
        '
        'mnuAddTime
        '
        Me.mnuAddTime.Index = 0
        Me.mnuAddTime.Text = "&Add Time"
        '
        'mnuRemoveTime
        '
        Me.mnuRemoveTime.Index = 1
        Me.mnuRemoveTime.Text = "&Remove Time"
        '
        'mnuEdTime
        '
        Me.mnuEdTime.Index = 2
        Me.mnuEdTime.Text = "&Edit Time"
        '
        'Panel4
        '
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(64, 443)
        Me.Panel4.TabIndex = 0
        '
        'TabCO
        '
        Me.TabCO.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListView1})
        Me.TabCO.Location = New System.Drawing.Point(4, 22)
        Me.TabCO.Name = "TabCO"
        Me.TabCO.Size = New System.Drawing.Size(728, 443)
        Me.TabCO.TabIndex = 4
        Me.TabCO.Text = "Collecting officers"
        '
        'ListView1
        '
        Me.ListView1.CheckBoxes = True
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListView1.ContextMenu = Me.ctxOfficers
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(728, 272)
        Me.ListView1.TabIndex = 0
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Officer Id"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Name"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Agreed"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Costs"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Currency"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Comments"
        Me.ColumnHeader6.Width = 400
        '
        'ctxOfficers
        '
        Me.ctxOfficers.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddOfficer, Me.mnuRemoveOfficer, Me.mnuEditPortOfficer})
        '
        'mnuAddOfficer
        '
        Me.mnuAddOfficer.Index = 0
        Me.mnuAddOfficer.Text = "&Add Officer"
        '
        'mnuRemoveOfficer
        '
        Me.mnuRemoveOfficer.Index = 1
        Me.mnuRemoveOfficer.Text = "&Remove Officer"
        '
        'mnuEditPortOfficer
        '
        Me.mnuEditPortOfficer.Index = 2
        Me.mnuEditPortOfficer.Text = "&Edit Port - Officer"
        '
        'imgYesNo
        '
        Me.imgYesNo.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgYesNo.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgYesNo.ImageStream = CType(resources.GetObject("imgYesNo.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgYesNo.TransparentColor = System.Drawing.Color.Transparent
        '
        'frmPort
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 509)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.Panel1})
        Me.Name = "frmPort"
        Me.Text = "Port Maintenance Form"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabDetails.ResumeLayout(False)
        Me.pnlOwnedSite.ResumeLayout(False)
        Me.TabLocation.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.TabAddress.ResumeLayout(False)
        Me.adrPanel.ResumeLayout(False)
        Me.AddressPanel1.ResumeLayout(False)
        Me.tabOpening.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.upDnInterval, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlTimeEdit.ResumeLayout(False)
        Me.TabCO.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private MyPort As Intranet.intranet.jobs.Port
    Private myAddress As MedscreenLib.Address.Caddress
    Private tnode As MedscreenCommonGui.Treenodes.TimeNodes.TimeNode
    Private dNode As MedscreenCommonGui.Treenodes.TimeNodes.DayNode
    Private RegionList As Intranet.intranet.jobs.CentreCollection
    Private myPortType As String = ""
    Private oCountry As MedscreenLib.Address.Country
    Private blnFillingForm As Boolean = True
    Private Times As Intranet.FixedSiteBookings.CentreAvailabilityCollection
    Private myCentre As Intranet.FixedSiteBookings.cCentre
    Private blnDrawn As Boolean = False
    Private SmidList As MedscreenLib.Glossary.PhraseCollection
    Private TypeList As MedscreenLib.Glossary.PhraseCollection

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Port object passed into the form 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Port() As Intranet.intranet.jobs.Port
        Get
            Return MyPort
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.Port)
            MyPort = Value
            myCentre = Nothing
            Me.Times = Nothing
            If Not Value Is Nothing Then
                'MyPort = ModTpanel.Ports.Item(MyPort.PortId)
                If Value.IsFixedSite Then
                    Me.Times = New Intranet.FixedSiteBookings.CentreAvailabilityCollection(Value.PortId)
                End If
                FillForm()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populate the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillForm()

        Me.ClearForm()
        blnFillingForm = True
        Dim objAdr As MedscreenLib.Address.Caddress
        Dim objAdr2 As MedscreenLib.Address.Caddress
        If MyPort Is Nothing Then Exit Sub
        With MyPort                                         'Use a with clause on port object to save typing
            If MyPort.AddressId = -1 Then                   'No address use a null address object
                Me.AddressPanel1.Address = objAdr
            Else
                objAdr2 = MedscreenLib.Glossary.Glossary.SMAddresses.item(MyPort.AddressId)
                myAddress = objAdr2                         'Set  global address object 
                'Me.AddressPanel1.Glossaries = ModTpanel.Glossaries
                Me.AddressPanel1.Address = myAddress        'Put that address into control
                Me.AddressPanel1.Invalidate()               'Redraw
            End If
            Me.txtPortID.Text = .PortId                     'Deal with port ID
            Me.txtPortID.Enabled = (.PortId.Trim.Length = 0)        'ID exists so no changes
            Me.ckFixedSite.Checked = (MyPort.IsFixedSite)           'Fixed Site
            Me.ckExtraCharges.Checked = (Not .IncursExtracharge)        'Extra charges involved 
            Me.txtIsoRegion.Text = .SubDivision
            Me.txtPortFunction.Text = .PortFunction
            Me.txtCoordinates.Text = .Coordinates
            'Port or site
            Dim i1 As Integer
            For i1 = 0 To Me.CBType.Items.Count - 1
                Dim strTemp As String = Me.CBType.Items.Item(i1)
                If Mid(strTemp, 1, 1) = .PortType Then
                    Me.CBType.SelectedIndex = i1
                    Exit For
                End If
            Next

            Me.txtDescription.Text = .PortDescription       'Port Description
            Me.txtComments.Text = .PortComments             'Port Comments
            Me.txtName.Text = .SiteName
            Me.txtMinCharge.Text = .SiteMinCharge
            Me.txtSiteCharge.Text = .SiteCharge
            Me.txtMapPath.Text = .SubDivision
            Me.txtSpecialNeeds.Text = .SpecialNeeds
            Me.txtISOPortCode.Text = .UNLocode.Trim
            Me.ckSpecialNeeds.Checked = .HasSpecialNeeds
            Me.lkDirections.Text = .isoCoordinates.Trim     'Clear directions link
            Me.lkDirections.Links.Clear()
            Me.lkDirections.Links.Add(0, .isoCoordinates.Trim.Length)       'Set up link
            If .Coordinates.Trim.Length > 0 Then
                Try
                    Me.LinkLabel1.Links.Clear()

                    Dim strLink As String = "http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & .Latitude & "&longitude=" & .Longitude & "&dtype=s"
                    LinkLabel1.Text = "View Map"
                    Me.LinkLabel1.Links.Add(0, 8)

                    'http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=52%2e329999999999998&longitude=0%2e53000000000000003&dtype=s
                    'Me.HtmlEditor1.LoadUrl("http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & .Latitude & "&longitude=" & .Longitude & "&dtype=s")
                    'Me.HtmlEditor1.Invalidate()
                Catch
                End Try
            Else
                Try
                    Me.HtmlEditor1.DocumentText = (.MapHTML(True, Intranet.intranet.jobs.Port.PicturePos.LeftPos) & "</body></html>")
                    Me.HtmlEditor1.Invalidate()
                Catch
                End Try
            End If
            'If Not MyPort Is Nothing Then
            If .IsFixedSite Then                'Deal with Fixed Sites
                Dim i As Integer
                Me.ckSpecialNeeds.Hide()        'Hide special needs
                Me.txtSpecialNeeds.Hide()
                myCentre = New Intranet.FixedSiteBookings.cCentre(Me.MyPort.Identity)
                Me.upDnInterval.Value = myCentre.TimeInterval
                For i = System.DayOfWeek.Sunday To System.DayOfWeek.Saturday

                Next
            Else
                Me.ckSpecialNeeds.Show()        'Show special needs
                Me.txtSpecialNeeds.Show()

            End If
            If .CountryId > 0 Then              'Is a country defined
                'Dim oCountry As MedscreenLib.Address.Country            'Get country
                oCountry = MedscreenLib.Glossary.Glossary.Countries.Item(MyPort.CountryId, 0)
                If Not oCountry Is Nothing Then 'We have a country
                    Me.cbCountry.SelectedIndex = MedscreenLib.Glossary.Glossary.Countries.IndexOf(oCountry)
                    Me.RegionList = cSupport.Ports.RegionList(oCountry.CountryId, Me.MyPort.PortType)
                    setRegion()
                    '                    Me.CBRegion.DataSource = Me.RegionList

                End If
            End If
            Me.ckRemoved.Checked = .Removed
            If MyPort.PortType = "P" Or MyPort.PortType = "S" Then
                FillOfficers()                  'Add Officers
            End If
            Me.TreeView1.BeginUpdate()
            Dim rootNode As New TreeNode("Availability")
            Me.TreeView1.Nodes.Add(rootNode)
            If Not Me.Times Is Nothing Then
                Dim objDay As Intranet.FixedSiteBookings.CentreAvailabilityDay
                For Each objDay In Me.Times
                    Dim aNode As New MedscreenCommonGui.Treenodes.TimeNodes.DayNode(objDay)
                    rootNode.Nodes.Add(aNode)
                    Dim objTime As Intranet.FixedSiteBookings.CentreAvailabilityTimes
                    For Each objTime In objDay
                        Dim tNode As New MedscreenCommonGui.Treenodes.TimeNodes.TimeNode(objTime)
                        aNode.Nodes.Add(tNode)
                    Next
                Next
            End If
            Me.TreeView1.EndUpdate()
            rootNode.ExpandAll()
            'Me.cbOwner.SelectedValue = .Owner
            If .Owner.Trim.Length = 0 Then
                Me.cbOwner.SelectedValue = "NONE"
            Else
                Me.cbOwner.SelectedValue = .Owner
            End If
            Me.cbOwnerType.SelectedValue = .OwnerType

            blnFillingForm = False

        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the Officers in 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillOfficers()
        Me.ListView1.Items.Clear()
        Dim objCoff As Intranet.intranet.jobs.CollPort
        For Each objCoff In MyPort.CollectingOfficers       'For each Officer add it to list
            AddListItem(objCoff)
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rewmove the contents of the form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ClearForm()
        Dim objAdr As MedscreenLib.Address.Caddress
        Me.AddressPanel1.Address = Nothing
        Me.txtPortID.Text = ""
        Me.txtPortID.Enabled = True
        Me.ckFixedSite.Checked = False
        Me.CBType.SelectedIndex = -1
        Me.ckSpecialNeeds.Checked = False
        Me.ckExtraCharges.Checked = False
        Me.ckRemoved.Checked = False
        Me.txtDescription.Text = ""
        Me.txtComments.Text = ""
        Me.txtName.Text = ""
        Me.txtMinCharge.Text = ""
        Me.txtSiteCharge.Text = ""
        Me.txtIsoRegion.Text = ""
        Me.txtMapPath.Text = ""
        Me.txtSpecialNeeds.Text = ""
        Me.lkDirections.Text = ""
        Me.txtISOPortCode.Text = ""

        Try
            Me.HtmlEditor1.DocumentText = ("<html><body></body></html>")
            Me.HtmlEditor1.Invalidate()
            'Me.HtmlEditor1.Url = System.Uri.
        Catch
        End Try
        Me.TreeView1.BeginUpdate()
        Me.TreeView1.Nodes.Clear()
        Me.TreeView1.EndUpdate()
        Me.cbCountry.SelectedIndex = -1
        Me.CBRegion.SelectedIndex = -1
        Me.ErrorProvider1.SetError(Me.CBRegion, "")
        Me.ErrorProvider1.SetError(Me.cbCountry, "")

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get details of the port
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetForm() As Boolean
        Dim objAdr As MedscreenLib.Address.Caddress

        Dim blnRet As Boolean = True
        Me.ErrorProvider1.SetError(Me.txtName, "")
        Me.ErrorProvider1.SetError(Me.txtDescription, "")

        If MyPort Is Nothing Then Exit Function
        Try
            With MyPort
                objAdr = Me.AddressPanel1.Address
                If Not objAdr Is Nothing Then
                    .AddressId = objAdr.AddressId
                End If
                MyPort.IsFixedSite = Me.ckFixedSite.Checked
                If Not myCentre Is Nothing Then
                    Me.myCentre.TimeInterval = Me.upDnInterval.Value
                    Me.MyPort.SpecialNeeds = Me.myCentre.myPort.SpecialNeeds
                    '    myCentre.isFixedSite = MyPort.IsFixedSite
                End If
                If Me.CBType.SelectedIndex > -1 Then
                    .PortType = Mid(CBType.Text, 1, 1)
                Else
                    .PortType = "P"
                End If
                .HasSpecialNeeds = Me.ckSpecialNeeds.Checked
                .IncursExtracharge = Not Me.ckExtraCharges.Checked
                If Not MyPort.IsFixedSite Then .SpecialNeeds = Me.txtSpecialNeeds.Text
                If Mid(Me.txtDescription.Text, 1, 5) = "00000" Then
                    Me.ErrorProvider1.SetError(Me.txtDescription, "Site description can not be a number")
                    blnRet = False
                End If
                .PortDescription = Me.txtDescription.Text
                .PortComments = Me.txtComments.Text
                'check site name is not numbers
                If Mid(Me.txtName.Text, 1, 5) = "00000" Then
                    Me.ErrorProvider1.SetError(Me.txtName, "Site name can not be a number")
                    blnRet = False
                End If
                .SiteName = Me.txtName.Text
                .UNLocode = Me.txtISOPortCode.Text
                Try
                    .SiteMinCharge = Me.txtMinCharge.Text
                    .SiteCharge = CDbl(Me.txtSiteCharge.Text)
                Catch
                End Try
                .SubDivision = Me.txtMapPath.Text
                '.isoCoordinates = Me.lkDirections.Text
                .isoCoordinates = jobs.CollectingOfficer.ConvertCoordinates(Me.txtCoordinates.Text)
                Dim oCountry As MedscreenLib.Address.Country
                Me.ErrorProvider1.SetError(Me.CBRegion, "")
                Me.ErrorProvider1.SetError(Me.cbCountry, "")
                If Me.cbCountry.SelectedIndex > 0 Then
                    oCountry = MedscreenLib.Glossary.Glossary.Countries.Item(Me.cbCountry.SelectedIndex)
                    If Not oCountry Is Nothing Then
                        .CountryId = oCountry.CountryId
                    End If
                    Me.ErrorProvider1.SetError(Me.cbCountry, "")
                Else
                    Me.ErrorProvider1.SetError(Me.cbCountry, "There must be a country associated with the Port")
                    blnRet = False
                End If
                If Me.CBRegion.SelectedIndex > -1 _
                  AndAlso Me.CBRegion.DataSource.count >= Me.CBRegion.SelectedIndex + 1 Then
                    Dim oRegion As Intranet.intranet.jobs.Port
                    oRegion = CBRegion.SelectedItem
                    If Not oRegion Is Nothing Then
                        .PortRegion = oRegion.Identity
                    Else
                        blnRet = False
                        Me.ErrorProvider1.SetError(Me.CBRegion, "There must be a region associated with the Port")

                    End If
                Else
                    Me.ErrorProvider1.SetError(Me.CBRegion, "There must be a region associated with the Port")
                    blnRet = False
                End If
                .Removed = Me.ckRemoved.Checked
                .OwnerType = cbOwnerType.SelectedValue
                .Owner = cbOwner.SelectedValue
                If cbOwner.SelectedValue = "NONE" Then
                    .Owner = ""
                Else
                    .Owner = Me.cbOwner.SelectedValue
                End If
            End With
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , MyPort.Identity)
        End Try
        Return blnRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change from fixed site to non fixed site or vice versa
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ckFixedSite_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckFixedSite.CheckedChanged
        If Me.ckFixedSite.Checked Then
            Me.ckFixedSite.ImageIndex = 1
            Me.pnlOwnedSite.Show()
        Else
            Me.ckFixedSite.ImageIndex = 0
            Me.pnlOwnedSite.Hide()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change from site to port or vice versa
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If CheckBox1.Checked Then
    '        Me.CheckBox1.ImageIndex = 2
    '        Me.CheckBox1.Text = "Site"
    '    Else
    '        Me.CheckBox1.ImageIndex = 3
    '        Me.CheckBox1.Text = "Port"
    '    End If
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Look up map
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub btnLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.OpenFileDialog1.InitialDirectory = "x:\map"
        'Me.OpenFileDialog1.FileName = Me.txtMapPath.Text
        'Me.OpenFileDialog1.Filter = "Image Files (*.gif; *.jpg)|*.gif;*.jpg|All Files (*.*)|*.*"

        'If Me.OpenFileDialog1.ShowDialog = DialogResult.OK Then
        '    Me.txtMapPath.Text = Me.OpenFileDialog1.FileName
        '    'MyPort.MapPath = Me.txtMapPath.Text
        '    Me.HtmlEditor1.DocumentText =(MyPort.MapHTML(True, Intranet.intranet.jobs.Port.PicturePos.LeftPos))
        '    Me.HtmlEditor1.Invalidate()
        'End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Look up map
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.OpenFileDialog1.InitialDirectory = "x:\map\Directions"
        'Me.OpenFileDialog1.FileName = Me.lkDirections.Text
        'Me.OpenFileDialog1.Filter = "Document Files (*.Doc)|*.doc;*.rtf|All files (*.*)|*.*"
        'If Me.OpenFileDialog1.ShowDialog = DialogResult.OK Then
        '    Me.lkDirections.Text = Me.OpenFileDialog1.FileName
        '    Me.lkDirections.Links.Clear()
        '    Me.lkDirections.Links.Add(0, Me.lkDirections.Text.Length)
        'End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accept form and exit 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.GetForm() Then
            MyPort.Update()
            If MyPort.Removed Then      'Sort out port and collecting officers at ports
                MyPort.RemovePort()
            Else
                MyPort.RestorePort()
            End If
            Me.DialogResult = DialogResult.OK
        Else
            Me.DialogResult = DialogResult.None
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Directions wanted
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lkDirections_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkDirections.LinkClicked
        Me.lkDirections.Links(lkDirections.Links.IndexOf(e.Link)).Visited = True
        If MyPort Is Nothing Then Exit Sub

        Dim strLink As System.Uri = New System.Uri("http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & MyPort.Latitude & "&longitude=" & MyPort.Longitude & "&dtype=s")
        Try
            Me.HtmlEditor1.Url = strLink
        Catch
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            myAddress = cSupport.SMAddresses.CreateAddress        'Get new address
            If Not myAddress Is Nothing Then
                Me.AddressPanel1.Address = myAddress                'Update GUI
                MyPort.AddressId = myAddress.AddressId              'Update Port
                MyPort.Update()                                     'Save updates
            Else
                MsgBox("error in creating address")
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Save address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        cmdSave.Enabled = False
        myAddress = Me.AddressPanel1.Address
        If Not myAddress Is Nothing Then
            myAddress.Update()
        End If
        cmdSave.Enabled = True
        Me.AddressPanel1.Focus()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Respond to a change in the tree
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Me.cmdAddTime.Visible = False
        Me.cmdUpdate.Visible = False

        If TypeOf TreeView1.SelectedNode Is MedscreenCommonGui.Treenodes.TimeNodes.TimeNode Then
            Me.mnuAddTime.Enabled = False
            Me.mnuRemoveTime.Enabled = True
            tnode = TreeView1.SelectedNode
            dNode = Nothing
            Me.CbDays.SelectedIndex = tnode.TimeValue.Day
            Me.DTEnd.Value = tnode.TimeValue.eTime
            Me.DtStart.Value = tnode.TimeValue.STime
        ElseIf TypeOf TreeView1.SelectedNode Is MedscreenCommonGui.Treenodes.TimeNodes.DayNode Then
            Me.mnuAddTime.Enabled = True
            Me.mnuRemoveTime.Enabled = False
            tnode = Nothing
            dNode = TreeView1.SelectedNode
            Me.CbDays.SelectedIndex = dNode.Day.Day
            Me.DTEnd.Value = Today.AddHours(16)
            Me.DtStart.Value = Today.AddHours(8)
        Else
            Me.mnuAddTime.Enabled = True
            Me.mnuRemoveTime.Enabled = False
            Me.CbDays.SelectedIndex = Me.TreeView1.SelectedNode.Tag
            Me.DTEnd.Value = Today.AddHours(16)
            Me.DtStart.Value = Today.AddHours(8)

        End If
        Me.PnlTimeEdit.Visible = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Remove an available time
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    '''     ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuRemoveTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemoveTime.Click
        'If Me.myCentre Is Nothing Then Exit Sub
        Dim strCommand As String = "delete from tos_centre_availability where port_id = '" & Me.MyPort.Identity & _
         "' and day = " & Me.tnode.TimeValue.Day & " and start_time = to_date('" & Me.tnode.TimeValue.STime.ToString("yyyyMMddHHmm") & "','yyyyMMddHH24mi')"


        'Dim oRead as oledb.OledbDataReader
        Dim oCmd As New OleDb.OleDbCommand()
        Try ' Protecting 
            oCmd.Connection = MedscreenLib.CConnection.DbConnection
            If MedscreenLib.CConnection.ConnOpen Then            'Attempt to open reader
                oCmd.CommandText = strCommand
                Dim intResult As Integer = oCmd.ExecuteNonQuery
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmPort-mnuRemoveTime_Click-1513")
        Finally
            MedscreenLib.CConnection.SetConnClosed()             'Close connection
            ' oRead.IsClosed Then oRead.Close()

        End Try
        'Dim intPos As Integer = Me.Times.IndexOf(Me.tnode.TimeValue)
        'Me.Times.RemoveAt(intPos)
        'me.Times.Item(
        'Me.myCentre.CentreTimes.Delete(Me.tnode.TimeValue)
        FillForm()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a time
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuAddTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddTime.Click
        Me.cmdAddTime.Visible = True
        Me.cmdUpdate.Visible = False
        If Not dNode Is Nothing Then
            Me.DtStart.Value = Me.dNode.Day.LatestEndTime
        Else
            Me.DtStart.Value = Today.AddHours(8)
        End If
        Me.CbDays.Enabled = True
        Me.DTEnd.Value = Me.DtStart.Value.AddHours(4)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit a time
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuEdTime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEdTime.Click
        Me.cmdAddTime.Visible = False
        Me.cmdUpdate.Visible = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update time node
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        If Me.MyPort Is Nothing Then Exit Sub
        If Me.tnode Is Nothing Then Exit Sub
        Me.tnode.TimeValue.STime = Me.DtStart.Value
        Me.tnode.TimeValue.eTime = Me.DTEnd.Value
        Me.Times.Update(Me.tnode.TimeValue)
        'Me.myCentre.CentreTimes.Update(Me.tnode.TimeValue)
        FillForm()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a time
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddTime.Click
        Dim intDay As Integer

        If Me.MyPort Is Nothing Then Exit Sub
        Me.Times.Create(Me.CbDays.SelectedIndex, Me.DtStart.Value, Me.DTEnd.Value)
        FillForm()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit acollecting Officer port information set
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick, mnuEditPortOfficer.Click
        Dim strId As Integer = Me.ListView1.SelectedItems.Item(0).Index
        Dim objCollPort As Intranet.intranet.jobs.CollPort

        objCollPort = Me.MyPort.CollectingOfficers.Item(strId)
        If Not objCollPort Is Nothing Then
            Dim aFrm As frmCollPort = New frmCollPort()
            aFrm.AgreedCosts = objCollPort.Agreed
            aFrm.Cost = objCollPort.Costs
            Dim myCollOfficer As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(objCollPort.OfficerId)
            aFrm.Currency = myCollOfficer.PayCurrency
            aFrm.Removed = objCollPort.Removed
            aFrm.Comments = objCollPort.Comments
            If aFrm.ShowDialog Then
                objCollPort.Agreed = aFrm.AgreedCosts
                objCollPort.Costs = aFrm.Cost
                objCollPort.Removed = aFrm.Removed
                objCollPort.Comments = aFrm.Comments
                objCollPort.Update()
                Me.ListView1.Items(strId).SubItems.Item(1).Text = objCollPort.Agreed
                Me.ListView1.Items(strId).SubItems.Item(2).Text = objCollPort.Agreed
                Me.FillOfficers()
            End If
        End If
        Me.ListView1.Items(strId).Checked = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an officer to port
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuAddOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddOfficer.Click
        Dim objCollPort As Intranet.intranet.jobs.CollPort = New Intranet.intranet.jobs.CollPort()
        objCollPort.PortId = Me.MyPort.PortId

        Dim frmAO As New frmAssignOfficer()

        frmAO.CollPort = objCollPort
        If frmAO.ShowDialog = DialogResult.OK Then
            objCollPort = frmAO.CollPort
            objCollPort.Update()
            Me.MyPort.CollectingOfficers.Add(objCollPort)
            AddListItem(objCollPort)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an officer to list view
    ''' </summary>
    ''' <param name="objCoff">Officer to add</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub AddListItem(ByVal objCoff As Intranet.intranet.jobs.CollPort)
        Dim lvItem As MedscreenCommonGui.ListViewItems.lvOfficerPortItem
        lvItem = New MedscreenCommonGui.ListViewItems.lvOfficerPortItem(objCoff)
        lvItem.DisplayMode = ListViewItems.lvOfficerPortItem.OfficerPortDisplayMode.OfficerMode
        'Dim objOff As Intranet.intranet.jobs.CollectingOfficer
        'objOff = cSupport.CollectingOfficers.Item(objCoff.OfficerId)

        'If Not objOff Is Nothing Then
        '    lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, objOff.Name))
        'Else
        '    lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, ""))
        'End If
        'lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, objCoff.Agreed))
        'lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, objCoff.Costs))
        'If Not objOff Is Nothing Then
        '    lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, objOff.PayCurrency))
        'Else
        '    lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, ""))
        'End If
        'lvItem.SubItems.Add(New ListViewItem.ListViewSubItem(lvItem, objCoff.Comments))
        'lvItem.Checked = True
        Me.ListView1.Items.Add(lvItem)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change country
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCountry.SelectedIndexChanged
        If Not blnDrawn Then Exit Sub
        If blnFillingForm Then Exit Sub
        If MyPort Is Nothing Then Exit Sub

        If Me.cbCountry.SelectedIndex < 0 Then Exit Sub
        oCountry = MedscreenLib.Glossary.Glossary.Countries.Item(Me.cbCountry.SelectedIndex)

        Me.ErrorProvider1.SetError(Me.cbCountry, "")

        setRegion()
    End Sub


    Private Sub setRegion()
        If oCountry Is Nothing Then Exit Sub
        If MyPort Is Nothing Then Exit Sub
        Try
            Me.RegionList = cSupport.Ports.RegionList(oCountry.CountryId, Me.MyPort.PortType)
            'check to see if any regions have been found
            If Me.RegionList.Count = 0 And Not (MyPort.PortType = "R" Or MyPort.PortType = "C") Then
                Me.ErrorProvider1.SetError(Me.cbCountry, "There are no regions are countries defined for this Country, please contact IT.")
            End If
            Me.CBRegion.DataSource = Me.RegionList
            If Not Me.RegionList Is Nothing Then        'Get a list of regions
                Dim oRegion As Intranet.intranet.jobs.Port
                oRegion = Me.RegionList.Item(MyPort.PortRegion)
                If Not oRegion Is Nothing Then
                    Me.CBRegion.SelectedIndex = RegionList.IndexOf(oRegion)
                End If
            End If
        Catch
        End Try

    End Sub
    Private Sub txtISOPortCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtISOPortCode.Leave
    End Sub

    Private Sub CBType_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CBType.Click

    End Sub

    Private Sub CBType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CBType.SelectedIndexChanged
        If blnFillingForm Then Exit Sub
        Me.myPortType = Mid(Me.CBType.Text, 1, 1)
        If oCountry Is Nothing Then Exit Sub
        If Me.myPortType.Trim.Length = 0 Then Exit Sub
        Me.RegionList = cSupport.Ports.RegionList(oCountry.CountryId, Me.myPortType)

    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        Me.blnDrawn = True
    End Sub

    Private Sub LinkLabel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.Click
        If MyPort Is Nothing Then Exit Sub

        Dim strLink As String = "http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & MyPort.Latitude & "&longitude=" & MyPort.Longitude & "&dtype=s"
        Try
            System.Diagnostics.Process.Start(strLink)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtISOPortCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtISOPortCode.Leave
        If Me.Port Is Nothing Then Exit Sub
        Me.ErrorProvider1.SetError(Me.txtISOPortCode, "")
        If txtISOPortCode.Text.Trim.Length < 2 Then Exit Sub

        Dim strCntryCode As String = Mid(txtISOPortCode.Text, 1, 2)
        If Me.Port.CountryId > 0 Then 'check match of country 
            Dim objCountry As MedscreenLib.Address.Country = MedscreenLib.Glossary.Glossary.Countries.Item(Port.CountryId, 0)
            If Not objCountry Is Nothing Then
                If objCountry.CountryCode <> strCntryCode Then Me.ErrorProvider1.SetError(Me.txtISOPortCode, "Mismatch of country")
            End If
        End If

        If Me.txtISOPortCode.Text.Trim.Length = 2 Or Me.txtISOPortCode.Text.Trim.Length = 5 Then
            Me.Port.UNLocode = Me.txtISOPortCode.Text
            Me.Port.GetISOInfo()
            Me.txtIsoRegion.Text = Port.SubDivision
            Me.txtPortFunction.Text = Port.PortFunction
            Me.txtCoordinates.Text = Port.Coordinates
            Me.txtMapPath.Text = Port.SubDivision
            Me.lkDirections.Text = Port.isoCoordinates.Trim     'Clear directions link
            Dim strTemp As String = Port.ISOPortName
            If strTemp.Trim.Length > 0 Then Me.txtDescription.Text = strTemp
            Me.lkDirections.Links.Clear()
            Me.lkDirections.Links.Add(0, Port.isoCoordinates.Trim.Length)       'Set up link
            If Port.Coordinates.Trim.Length > 0 Then
                Try
                    Me.LinkLabel1.Links.Clear()

                    Dim strLink As String = "http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & Port.Latitude & "&longitude=" & Port.Longitude & "&dtype=s"
                    LinkLabel1.Text = "View Map"
                    Me.LinkLabel1.Links.Add(0, 8)
                Catch
                End Try
            Else
                Try
                    Me.HtmlEditor1.DocumentText = (Port.MapHTML(True, Intranet.intranet.jobs.Port.PicturePos.LeftPos) & "</body></html>")
                    Me.HtmlEditor1.Invalidate()
                Catch
                End Try
            End If

        End If

    End Sub


    Private Sub ListView1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.Click

    End Sub

    Private Sub txtCoordinates_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCoordinates.Leave
        Dim strCoords As String = Intranet.intranet.jobs.CollectingOfficer.ConvertCoordinates(txtCoordinates.Text)
        Me.Port.isoCoordinates = strCoords
        Me.lkDirections.Text = Port.isoCoordinates.Trim     'Clear directions link
        Me.lkDirections.Links.Clear()
        Me.lkDirections.Links.Add(0, Port.isoCoordinates.Trim.Length)       'Set up link

        If Port.Coordinates.Trim.Length > 0 Then
            Try
                Me.LinkLabel1.Links.Clear()

                Dim strLink As String = "http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & Port.Latitude & "&longitude=" & Port.Longitude & "&dtype=s"
                LinkLabel1.Text = "View Map"
                Me.LinkLabel1.Links.Add(0, 8)

                'http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=52%2e329999999999998&longitude=0%2e53000000000000003&dtype=s
                'Me.HtmlEditor1.LoadUrl("http://www.mapquest.com/maps/map.adp?formtype=latlong&latlongtype=decimal&latitude=" & .Latitude & "&longitude=" & .Longitude & "&dtype=s")
                'Me.HtmlEditor1.Invalidate()
            Catch
            End Try
        Else
            Try
                Me.HtmlEditor1.DocumentText = (Port.MapHTML(True, Intranet.intranet.jobs.Port.PicturePos.LeftPos) & "</body></html>")
                Me.HtmlEditor1.Invalidate()
            Catch
            End Try
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Call find address GUI
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSearchAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearchAddress.Click
        Dim Faddr As New FindAddress()

        If Faddr.ShowDialog = DialogResult.OK Then
            Dim intId As Integer = Faddr.AddressID          'Get found address Id
            Dim oAddr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(intId)
            Me.myAddress = oAddr                            'Show in GUI
            Me.AddressPanel1.Address = oAddr
            If myAddress Is Nothing Then Exit Sub '          If nowt skip out
            MyPort.AddressId = myAddress.AddressId              'Update Port
            MyPort.Update()                                     'Save updates

            Me.AddressPanel1.isReadOnly = False
            
        End If
    End Sub

End Class
