Imports Intranet
Imports MedscreenLib.Medscreen
Imports MedscreenLib
Imports System.Windows.Forms
Imports System.Drawing
Public Class AddressControl
    Inherits System.Windows.Forms.UserControl

    Private myId As Long
   
    Private myAddress As MedscreenLib.Address.Caddress

    Private blnBuilt As Boolean = False


    Public Event NewAddress(ByVal lngId As Long)
    Public Event SaveAddress(ByVal lngid As Long)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'myAddressList = New oneStopClasses.Address.AddressList()
    End Sub

    'UserControl1 overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Friend WithEvents mfp_editADDRLINE2 As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE1 As System.Windows.Forms.Label
    Friend WithEvents mfp_editADDRLINE1 As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE2 As System.Windows.Forms.Label
    Friend WithEvents mfp_editADDRLINE3 As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE3 As System.Windows.Forms.Label
    Friend WithEvents mfp_editCITY As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblDISTRICT As System.Windows.Forms.Label
    Friend WithEvents mfp_editDISTRICT As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblCITY As System.Windows.Forms.Label
    Friend WithEvents mfp_lblADDRESS_ID As System.Windows.Forms.Label
    Friend WithEvents mfp_lblPOSTCODE As System.Windows.Forms.Label
    Friend WithEvents mfp_lblCOUNTRY As System.Windows.Forms.Label
    Friend WithEvents mfp_editPOSTCODE As System.Windows.Forms.TextBox
    Friend WithEvents mfp_editFAX As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblEMAIL As System.Windows.Forms.Label
    Friend WithEvents mfp_editEMAIL As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblPHONE As System.Windows.Forms.Label
    Friend WithEvents mfp_lblFAX As System.Windows.Forms.Label
    Friend WithEvents mfp_editPHONE As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents mfp_btnCancel As System.Windows.Forms.Button
    Friend WithEvents mfp_btnAdd As System.Windows.Forms.Button
    Friend WithEvents mfp_btnDelete As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Private components As System.ComponentModel.IContainer

    'Required by the Windows Form Designer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    'Friend DsAddress2 As oneStopClasses.dsAddress.TOS_ADDRESSRow
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    'Friend WithEvents DsCountry1 As oneStopClasses.DSCountry
    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    'Friend WithEvents DsAddrType1 As DSAddrType
    Friend WithEvents cbAddrType As System.Windows.Forms.ComboBox
    Friend WithEvents lblAddrType As System.Windows.Forms.Label
    'Friend WithEvents DataView1 As System.Data.DataView
    'Friend WithEvents DsStatusType1 As oneStopClasses.DSStatusType
    Friend WithEvents cbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuProperCase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUpperCase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLowerCase As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CXMenuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuBlockAddress As System.Windows.Forms.MenuItem
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents mfp_editADDRLINE4 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdDownone As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddressControl))
        Me.mfp_editADDRLINE3 = New System.Windows.Forms.TextBox()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.mnuProperCase = New System.Windows.Forms.MenuItem()
        Me.mnuUpperCase = New System.Windows.Forms.MenuItem()
        Me.mnuLowerCase = New System.Windows.Forms.MenuItem()
        Me.mfp_editADDRLINE2 = New System.Windows.Forms.TextBox()
        Me.mfp_editADDRLINE1 = New System.Windows.Forms.TextBox()
        Me.mfp_editPHONE = New System.Windows.Forms.TextBox()
        Me.mfp_lblCITY = New System.Windows.Forms.Label()
        Me.mfp_lblADDRLINE1 = New System.Windows.Forms.Label()
        Me.mfp_lblADDRLINE3 = New System.Windows.Forms.Label()
        Me.mfp_lblPOSTCODE = New System.Windows.Forms.Label()
        Me.mfp_editFAX = New System.Windows.Forms.TextBox()
        Me.mfp_lblEMAIL = New System.Windows.Forms.Label()
        Me.mfp_editPOSTCODE = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.mfp_btnCancel = New System.Windows.Forms.Button()
        Me.mfp_btnAdd = New System.Windows.Forms.Button()
        Me.mfp_btnDelete = New System.Windows.Forms.Button()
        Me.mfp_lblDISTRICT = New System.Windows.Forms.Label()
        Me.mfp_lblPHONE = New System.Windows.Forms.Label()
        Me.mfp_editEMAIL = New System.Windows.Forms.TextBox()
        Me.mfp_lblCOUNTRY = New System.Windows.Forms.Label()
        Me.mfp_lblADDRLINE2 = New System.Windows.Forms.Label()
        Me.mfp_editDISTRICT = New System.Windows.Forms.TextBox()
        Me.mfp_lblFAX = New System.Windows.Forms.Label()
        Me.mfp_lblADDRESS_ID = New System.Windows.Forms.Label()
        Me.mfp_editCITY = New System.Windows.Forms.TextBox()
        Me.cbCountry = New System.Windows.Forms.ComboBox()
        'Me.DsAddrType1 = New DSAddrType()
        Me.cbAddrType = New System.Windows.Forms.ComboBox()
        Me.lblAddrType = New System.Windows.Forms.Label()
        Me.cbStatus = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CXMenuMain = New System.Windows.Forms.ContextMenu()
        Me.mnuBlockAddress = New System.Windows.Forms.MenuItem()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.mfp_editADDRLINE4 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdDownone = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        'CType(Me.DsAddrType1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'mfp_editADDRLINE3
        '
        Me.mfp_editADDRLINE3.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE3.Location = New System.Drawing.Point(56, 81)
        Me.mfp_editADDRLINE3.Name = "mfp_editADDRLINE3"
        Me.mfp_editADDRLINE3.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE3.TabIndex = 10
        Me.mfp_editADDRLINE3.Text = ""
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuProperCase, Me.mnuUpperCase})
        '
        'mnuProperCase
        '
        Me.mnuProperCase.Index = 0
        Me.mnuProperCase.Text = "Proper Case"
        '
        'mnuUpperCase
        '
        Me.mnuUpperCase.Index = 1
        Me.mnuUpperCase.Text = "Upper Case"
        '
        'mfp_editADDRLINE2
        '
        Me.mfp_editADDRLINE2.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE2.Location = New System.Drawing.Point(56, 54)
        Me.mfp_editADDRLINE2.Name = "mfp_editADDRLINE2"
        Me.mfp_editADDRLINE2.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE2.TabIndex = 9
        Me.mfp_editADDRLINE2.Text = ""
        '
        'mfp_editADDRLINE1
        '
        Me.mfp_editADDRLINE1.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE1.Location = New System.Drawing.Point(56, 28)
        Me.mfp_editADDRLINE1.Name = "mfp_editADDRLINE1"
        Me.mfp_editADDRLINE1.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE1.TabIndex = 8
        Me.mfp_editADDRLINE1.Text = ""
        '
        'mfp_editPHONE
        '
        Me.mfp_editPHONE.Location = New System.Drawing.Point(416, 40)
        Me.mfp_editPHONE.Name = "mfp_editPHONE"
        Me.mfp_editPHONE.TabIndex = 21
        Me.mfp_editPHONE.Text = ""
        Me.ToolTip1.SetToolTip(Me.mfp_editPHONE, "Phone number")
        '
        'mfp_lblCITY
        '
        Me.mfp_lblCITY.Location = New System.Drawing.Point(8, 127)
        Me.mfp_lblCITY.Name = "mfp_lblCITY"
        Me.mfp_lblCITY.Size = New System.Drawing.Size(40, 23)
        Me.mfp_lblCITY.TabIndex = 5
        Me.mfp_lblCITY.Text = "City"
        '
        'mfp_lblADDRLINE1
        '
        Me.mfp_lblADDRLINE1.Location = New System.Drawing.Point(10, 28)
        Me.mfp_lblADDRLINE1.Name = "mfp_lblADDRLINE1"
        Me.mfp_lblADDRLINE1.Size = New System.Drawing.Size(38, 23)
        Me.mfp_lblADDRLINE1.TabIndex = 2
        Me.mfp_lblADDRLINE1.Text = "Line1"
        '
        'mfp_lblADDRLINE3
        '
        Me.mfp_lblADDRLINE3.Location = New System.Drawing.Point(8, 82)
        Me.mfp_lblADDRLINE3.Name = "mfp_lblADDRLINE3"
        Me.mfp_lblADDRLINE3.Size = New System.Drawing.Size(38, 23)
        Me.mfp_lblADDRLINE3.TabIndex = 4
        Me.mfp_lblADDRLINE3.Text = "Line 3"
        '
        'mfp_lblPOSTCODE
        '
        Me.mfp_lblPOSTCODE.Location = New System.Drawing.Point(8, 154)
        Me.mfp_lblPOSTCODE.Name = "mfp_lblPOSTCODE"
        Me.mfp_lblPOSTCODE.Size = New System.Drawing.Size(56, 23)
        Me.mfp_lblPOSTCODE.TabIndex = 13
        Me.mfp_lblPOSTCODE.Text = "Post code"
        '
        'mfp_editFAX
        '
        Me.mfp_editFAX.Location = New System.Drawing.Point(416, 72)
        Me.mfp_editFAX.Name = "mfp_editFAX"
        Me.mfp_editFAX.TabIndex = 22
        Me.mfp_editFAX.Text = ""
        Me.ToolTip1.SetToolTip(Me.mfp_editFAX, "Fax number")
        '
        'mfp_lblEMAIL
        '
        Me.mfp_lblEMAIL.Location = New System.Drawing.Point(8, 182)
        Me.mfp_lblEMAIL.Name = "mfp_lblEMAIL"
        Me.mfp_lblEMAIL.Size = New System.Drawing.Size(56, 23)
        Me.mfp_lblEMAIL.TabIndex = 17
        Me.mfp_lblEMAIL.Text = "EMAIL"
        '
        'mfp_editPOSTCODE
        '
        Me.mfp_editPOSTCODE.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.mfp_editPOSTCODE.Location = New System.Drawing.Point(64, 154)
        Me.mfp_editPOSTCODE.Name = "mfp_editPOSTCODE"
        Me.mfp_editPOSTCODE.Size = New System.Drawing.Size(88, 20)
        Me.mfp_editPOSTCODE.TabIndex = 19
        Me.mfp_editPOSTCODE.Text = ""
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdSave, Me.cmdCancel, Me.Button2, Me.cmdDelete, Me.mfp_btnCancel, Me.mfp_btnAdd, Me.mfp_btnDelete})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.DockPadding.All = 30
        Me.Panel1.Location = New System.Drawing.Point(0, 216)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(552, 24)
        Me.Panel1.TabIndex = 24
        '
        'cmdSave
        '
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Bitmap)
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(184, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.TabIndex = 28
        Me.cmdSave.Text = "&Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(456, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 27
        Me.cmdCancel.Text = "&Cancel"
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Bitmap)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(8, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 25
        Me.Button2.Text = "&Add"
        '
        'cmdDelete
        '
        Me.cmdDelete.Image = CType(resources.GetObject("cmdDelete.Image"), System.Drawing.Bitmap)
        Me.cmdDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDelete.Location = New System.Drawing.Point(96, 4)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.TabIndex = 26
        Me.cmdDelete.Text = "&Delete"
        '
        'mfp_btnCancel
        '
        Me.mfp_btnCancel.Location = New System.Drawing.Point(365, 241)
        Me.mfp_btnCancel.Name = "mfp_btnCancel"
        Me.mfp_btnCancel.TabIndex = 27
        Me.mfp_btnCancel.Text = "&Cancel"
        '
        'mfp_btnAdd
        '
        Me.mfp_btnAdd.Location = New System.Drawing.Point(195, 241)
        Me.mfp_btnAdd.Name = "mfp_btnAdd"
        Me.mfp_btnAdd.TabIndex = 25
        Me.mfp_btnAdd.Text = "&Add"
        '
        'mfp_btnDelete
        '
        Me.mfp_btnDelete.Location = New System.Drawing.Point(280, 241)
        Me.mfp_btnDelete.Name = "mfp_btnDelete"
        Me.mfp_btnDelete.TabIndex = 26
        Me.mfp_btnDelete.Text = "&Delete"
        '
        'mfp_lblDISTRICT
        '
        Me.mfp_lblDISTRICT.Location = New System.Drawing.Point(168, 126)
        Me.mfp_lblDISTRICT.Name = "mfp_lblDISTRICT"
        Me.mfp_lblDISTRICT.Size = New System.Drawing.Size(40, 23)
        Me.mfp_lblDISTRICT.TabIndex = 6
        Me.mfp_lblDISTRICT.Text = "District"
        '
        'mfp_lblPHONE
        '
        Me.mfp_lblPHONE.Location = New System.Drawing.Point(344, 40)
        Me.mfp_lblPHONE.Name = "mfp_lblPHONE"
        Me.mfp_lblPHONE.Size = New System.Drawing.Size(64, 23)
        Me.mfp_lblPHONE.TabIndex = 15
        Me.mfp_lblPHONE.Text = "Phone"
        '
        'mfp_editEMAIL
        '
        Me.mfp_editEMAIL.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.mfp_editEMAIL.Location = New System.Drawing.Point(64, 182)
        Me.mfp_editEMAIL.Name = "mfp_editEMAIL"
        Me.mfp_editEMAIL.Size = New System.Drawing.Size(456, 20)
        Me.mfp_editEMAIL.TabIndex = 23
        Me.mfp_editEMAIL.Text = ""
        '
        'mfp_lblCOUNTRY
        '
        Me.mfp_lblCOUNTRY.Location = New System.Drawing.Point(168, 158)
        Me.mfp_lblCOUNTRY.Name = "mfp_lblCOUNTRY"
        Me.mfp_lblCOUNTRY.Size = New System.Drawing.Size(48, 23)
        Me.mfp_lblCOUNTRY.TabIndex = 14
        Me.mfp_lblCOUNTRY.Text = "Country"
        '
        'mfp_lblADDRLINE2
        '
        Me.mfp_lblADDRLINE2.Location = New System.Drawing.Point(10, 55)
        Me.mfp_lblADDRLINE2.Name = "mfp_lblADDRLINE2"
        Me.mfp_lblADDRLINE2.Size = New System.Drawing.Size(46, 23)
        Me.mfp_lblADDRLINE2.TabIndex = 3
        Me.mfp_lblADDRLINE2.Text = "Line2"
        '
        'mfp_editDISTRICT
        '
        Me.mfp_editDISTRICT.ContextMenu = Me.ContextMenu1
        Me.mfp_editDISTRICT.Location = New System.Drawing.Point(216, 126)
        Me.mfp_editDISTRICT.Name = "mfp_editDISTRICT"
        Me.mfp_editDISTRICT.Size = New System.Drawing.Size(120, 20)
        Me.mfp_editDISTRICT.TabIndex = 12
        Me.mfp_editDISTRICT.Text = ""
        '
        'mfp_lblFAX
        '
        Me.mfp_lblFAX.Location = New System.Drawing.Point(344, 72)
        Me.mfp_lblFAX.Name = "mfp_lblFAX"
        Me.mfp_lblFAX.Size = New System.Drawing.Size(48, 23)
        Me.mfp_lblFAX.TabIndex = 16
        Me.mfp_lblFAX.Text = "FAX"
        '
        'mfp_lblADDRESS_ID
        '
        Me.mfp_lblADDRESS_ID.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mfp_lblADDRESS_ID.ForeColor = System.Drawing.SystemColors.Desktop
        Me.mfp_lblADDRESS_ID.Location = New System.Drawing.Point(16, 8)
        Me.mfp_lblADDRESS_ID.Name = "mfp_lblADDRESS_ID"
        Me.mfp_lblADDRESS_ID.Size = New System.Drawing.Size(200, 16)
        Me.mfp_lblADDRESS_ID.TabIndex = 1
        Me.mfp_lblADDRESS_ID.Text = "Address Id:"


        '
        'mfp_editCITY
        '
        Me.mfp_editCITY.ContextMenu = Me.ContextMenu1
        Me.mfp_editCITY.Location = New System.Drawing.Point(56, 127)
        Me.mfp_editCITY.Name = "mfp_editCITY"
        Me.mfp_editCITY.TabIndex = 11
        Me.mfp_editCITY.Text = ""
        '
        'cbCountry
        '
        Me.cbCountry.DisplayMember = "CountryName"
        Me.cbCountry.Location = New System.Drawing.Point(216, 150)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(120, 21)
        Me.cbCountry.TabIndex = 25
        Me.cbCountry.ValueMember = "COUNTRY_ID"
        '
        'DsAddrType1
        '
        'Me.DsAddrType1.DataSetName = "DSAddrType"
        'Me.DsAddrType1.Locale = New System.Globalization.CultureInfo("en-GB")
        'Me.DsAddrType1.Namespace = "http://www.tempuri.org/DSAddrType.xsd"
        '
        'cbAddrType
        '
        'Me.cbAddrType.DataSource = Me.DsAddrType1.TOS_ADDR_TYPE
        Me.cbAddrType.DisplayMember = "PhraseText"
        Me.cbAddrType.Location = New System.Drawing.Point(416, 104)
        Me.cbAddrType.Name = "cbAddrType"
        Me.cbAddrType.Size = New System.Drawing.Size(104, 21)
        Me.cbAddrType.TabIndex = 26
        Me.cbAddrType.Text = "ComboBox2"
        Me.cbAddrType.ValueMember = "PhraseID"
        '
        'lblAddrType
        '
        Me.lblAddrType.Location = New System.Drawing.Point(344, 96)
        Me.lblAddrType.Name = "lblAddrType"
        Me.lblAddrType.Size = New System.Drawing.Size(64, 32)
        Me.lblAddrType.TabIndex = 27
        Me.lblAddrType.Text = "Type"
        '
        'cbStatus
        '
        Me.cbStatus.DisplayMember = "STATUS_TYPE_DESCRIPTION"
        Me.cbStatus.Enabled = False
        Me.cbStatus.Location = New System.Drawing.Point(416, 136)
        Me.cbStatus.Name = "cbStatus"
        Me.cbStatus.Size = New System.Drawing.Size(104, 21)
        Me.cbStatus.TabIndex = 28
        Me.cbStatus.Text = "Status"
        Me.cbStatus.ValueMember = "STATUS_TYPE"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(344, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 32)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Status"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.ContextMenu = Me.CXMenuMain
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdDownone, Me.lblAddrType})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(552, 240)
        Me.Panel2.TabIndex = 30
        '
        'CXMenuMain
        '
        Me.CXMenuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuBlockAddress})
        '
        'mnuBlockAddress
        '
        Me.mnuBlockAddress.Index = 0
        Me.mnuBlockAddress.Text = "&Block Address"
        '
        'mfp_editADDRLINE4
        '
        Me.mfp_editADDRLINE4.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE4.Location = New System.Drawing.Point(56, 104)
        Me.mfp_editADDRLINE4.Name = "mfp_editADDRLINE4"
        Me.mfp_editADDRLINE4.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE4.TabIndex = 32
        Me.mfp_editADDRLINE4.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 23)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Line 4"
        '
        'cmdDownone
        '
        Me.cmdDownone.Image = CType(resources.GetObject("cmdDownone.Image"), System.Drawing.Bitmap)
        Me.cmdDownone.Location = New System.Drawing.Point(336, 0)
        Me.cmdDownone.Name = "cmdDownone"
        Me.cmdDownone.Size = New System.Drawing.Size(32, 23)
        Me.cmdDownone.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.cmdDownone, "Move address lines 1-3 down 1")
        '
        'AddressControl
        '
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.mfp_editADDRLINE4, Me.Label2, Me.Label1, Me.cbStatus, Me.cbAddrType, Me.cbCountry, Me.Panel1, Me.mfp_editFAX, Me.mfp_lblEMAIL, Me.mfp_editEMAIL, Me.mfp_lblPHONE, Me.mfp_lblFAX, Me.mfp_editPHONE, Me.mfp_lblPOSTCODE, Me.mfp_lblCOUNTRY, Me.mfp_editPOSTCODE, Me.mfp_lblADDRESS_ID, Me.mfp_editADDRLINE3, Me.mfp_lblDISTRICT, Me.mfp_editDISTRICT, Me.mfp_lblADDRLINE3, Me.mfp_editCITY, Me.mfp_lblCITY, Me.mfp_editADDRLINE2, Me.mfp_lblADDRLINE1, Me.mfp_editADDRLINE1, Me.mfp_lblADDRLINE2, Me.Panel2})
        Me.Name = "AddressControl"
        Me.Size = New System.Drawing.Size(552, 240)
        Me.Panel1.ResumeLayout(False)
        '  CType(Me.DsAddrType1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Save the address back to the database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Save() As Boolean
        Me.cmdSave_Click(Nothing, Nothing)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the address Id, if avalid id is supllied fill control
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AddressId() As Long
        Get
            AddressId = myId
        End Get
        Set(ByVal Value As Long)
            myAddress = Nothing
            If myId > 0 Then
                GetControl()                        'Get current contents of control

                Dim myAddress2 As MedscreenLib.Address.Caddress     'Set up new address
                myAddress2 = MedscreenLib.Glossary.Glossary.SMAddresses.item(myId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                'If tRow.RowState = DataRowState.Modified Then
                myAddress.Update()                                  'Ensure address is updated
                'End 'If
            End If
            myId = Value                                            'Set new id
            EmptyControl()                                          'Get rid of current contents
            If myId > 0 Then                                        'if new address is valid get it and fill control
                myAddress = MedscreenLib.Glossary.Glossary.SMAddresses.item(myId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                If Not myAddress Is Nothing Then                    'Have we got an address
                    FillControl()                                   'Fill control 
                    BlockAddress()
                    Invalidate()
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass address object to control
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Address() As MedscreenLib.Address.Caddress
        Get
            Return CType(myAddress, MedscreenLib.Address.Caddress)
        End Get
        Set(ByVal Value As MedscreenLib.Address.Caddress)
            myAddress = Value
            If Not myAddress Is Nothing Then
                FillControl()
                BlockAddress()
                Invalidate()
            Else
                EmptyControl()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get rid of contents of control
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub EmptyControl()
        Me.mfp_editADDRLINE1.Text = ""
        Me.mfp_editADDRLINE2.Text = ""
        Me.mfp_editADDRLINE3.Text = ""
        Me.mfp_editCITY.Text = ""
        Me.mfp_editDISTRICT.Text = ""
        Me.mfp_editEMAIL.Text = ""
        Me.mfp_editPOSTCODE.Text = ""
        Me.mfp_editFAX.Text = ""
        Me.mfp_editPHONE.Text = ""
        Me.mfp_lblADDRESS_ID.Text = "Address ID: "


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populate individual elements of the control
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillControl()
        EmptyControl()

        'If myGlossaries Is Nothing Then
        '    myGlossaries = New oneStopClasses.clsGlossary()
        'End If
        'Me.DsCountry1 = myGlossaries.Countrylist
        ' Me.cbCountry.DataSource = Me.DsCountry1.TOS_COUNTRY
        Me.cbCountry.DataSource = MedscreenLib.Glossary.Glossary.Countries
        Me.cbCountry.ValueMember = "CountryID"
        Me.cbCountry.DisplayMember = "CountryName"
        'Me.DsAddrType1 = myGlossaries.AddrTypelist
        Dim addrTypeList As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("ADDR_TYPE")
        Me.cbAddrType.DataSource = AddrTypeList

        'Me.cbAddrType.DataSource = Me.DsAddrType1.TOS_ADDR_TYPE
        Me.cbAddrType.ValueMember = "PhraseID" '"ADDRESS_TYPE"
        Me.cbAddrType.DisplayMember = "PhraseText" ' "ADDR_TYPE_NAME"
        'Me.DataView1 = myGlossaries.StatusTypeList("ADDR")
        'Me.cbStatus.DataSource = Me.DataView1


        Me.mfp_editADDRLINE1.Text = myAddress.AdrLine1
        Me.mfp_editADDRLINE2.Text = myAddress.AdrLine2
        Me.mfp_editADDRLINE3.Text = myAddress.AdrLine3
        Me.mfp_editADDRLINE4.Text = myAddress.AdrLine4

        Me.mfp_editCITY.Text = myAddress.City

        If Not myAddress.Country = 0 Then
            Me.cbCountry.SelectedValue = myAddress.Country
            SetCountrySpecific(myAddress.Country)
        End If
        Me.mfp_editDISTRICT.Text = myAddress.District

        Me.mfp_editEMAIL.Text = myAddress.Email

        Me.mfp_editFAX.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Fax, myAddress.Country)
        Me.mfp_editPHONE.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Phone, myAddress.Country)
        setPhones()

        Me.mfp_editPOSTCODE.Text = myAddress.PostCode

        Me.mfp_lblADDRESS_ID.Text = "Address ID: " & myAddress.AddressId
        Try
            Me.cbAddrType.SelectedValue = Me.myAddress.AddressType()
        Catch
        End Try
        SetStatus(Me.myAddress.Status)
        If myAddress.Status = Intranet.intranet.Support.cSupport.AddrDeleted Then
            Me.cmdDelete.Text = "Un&Delete"
        Else
            Me.cmdDelete.Text = "&Delete"
        End If
        blnBuilt = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sort out phone numbers
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub setPhones()
        Me.ToolTip1.RemoveAll()
        If Not myAddress.Phone.Length > 0 Then
            Me.ToolTip1.SetToolTip(Me.mfp_editPHONE, mfp_editPHONE.Text)
        End If
        If Not myAddress.Fax.Length > 0 Then
            ToolTip1.SetToolTip(Me.mfp_editFAX, "00" & myAddress.Country & " " & mfp_editFAX.Text)
        End If



    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Alter the control to deal with specific countries
    ''' </summary>
    ''' <param name="intCountry">Country ID to use</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetCountrySpecific(ByVal intCountry As Integer)
        Select Case intCountry
            Case Intranet.intranet.Support.cSupport.CountryCodes.CountryUK
                Me.mfp_lblDISTRICT.Text = "County"
                Me.mfp_lblPOSTCODE.Text = "Postcode"
            Case Intranet.intranet.Support.cSupport.CountryCodes.CountryUSA
                Me.mfp_lblDISTRICT.Text = "State"
                Me.mfp_lblPOSTCODE.Text = "Zip code"

        End Select
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up status value
    ''' </summary>
    ''' <param name="strStatusValue">Status value to use</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetStatus(ByVal strStatusValue As String)
        'Dim dr As oneStopClasses.DSStatusType.TOS_STATUS_TYPERow
        'Dim intRow As Integer = 0

        'For Each dr In Me.DataView1.Table.Rows
        '    If dr.STATUS_TYPE = strStatusValue Then

        '        Me.cbStatus.Text = dr.STATUS_TYPE_DESCRIPTION
        '        Me.cbStatus.SelectedIndex = intRow
        '        Exit For
        '    End If
        '    intRow = intRow + 1
        'Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get contents of control
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub GetControl()
        'If myGlossaries Is Nothing Then
        '    myGlossaries = New oneStopClasses.clsGlossary()
        'End If
        With myAddress
            .AdrLine1 = Me.mfp_editADDRLINE1.Text
            .AdrLine2 = Me.mfp_editADDRLINE2.Text
            .AdrLine3 = Me.mfp_editADDRLINE3.Text
            .AdrLine4 = Me.mfp_editADDRLINE4.Text
            .City = Me.mfp_editCITY.Text
            .Country = Me.cbCountry.SelectedValue
            .District = Me.mfp_editDISTRICT.Text
            .Email = Me.mfp_editEMAIL.Text
            .Fax = MedscreenLib.Medscreen.StripCountry(Me.mfp_editFAX.Text)
            .Phone = MedscreenLib.Medscreen.StripCountry(Me.mfp_editPHONE.Text)
            .PostCode = Me.mfp_editPOSTCODE.Text
            .AddressType = Me.cbAddrType.SelectedValue()
        End With
        ' DsAddress2.ADDRESS_STATUS = Me.cbStatus.SelectedValue
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        myAddress = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress()
        'Me.DsAddress2 = tRow
        FillControl()
        RaiseEvent NewAddress(myAddress.AddressId)

    End Sub


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        GetControl()
        myAddress.Update()
        RaiseEvent SaveAddress(myAddress.AddressId)
    End Sub


    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If myAddress.Status = Intranet.intranet.Support.cSupport.AddrActive Then
            myAddress.Status = Intranet.intranet.Support.cSupport.AddrDeleted
        Else
            myAddress.Status = Intranet.intranet.Support.cSupport.AddrActive
        End If
        'tRow = DsAddress2
        'myAddressList.updateAddress()
        myAddress.Update()
        FillControl()
        Me.Invalidate()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        myAddress.Fields.Rollback()
        FillControl()
        Me.Invalidate()


    End Sub

    Private Sub cbCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCountry.SelectedIndexChanged
        If blnBuilt Then
            myAddress.Country = Me.cbCountry.SelectedValue
            SetCountrySpecific(myAddress.Country)
        End If
    End Sub

    Private Sub mnuProperCase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuProperCase.Click
        Dim tmpString As String
        Dim tmpSt2 As String
        Dim c As String
        Dim i As Integer
        Dim blnUpShift As Boolean
        Dim optMen As ContextMenu
        Dim txtBox As TextBox

        If TypeOf sender Is MenuItem Then
            optMen = sender.parent
            If TypeOf optMen.SourceControl Is TextBox Then
                txtBox = optMen.SourceControl
                tmpSt2 = ""
                tmpString = LCase(txtBox.Text)
                blnUpShift = True
                For i = 1 To Len(tmpString)
                    c = Mid(tmpString, i, 1)
                    If blnUpShift Then
                        tmpSt2 = tmpSt2 & UCase(c)
                    Else
                        tmpSt2 = tmpSt2 & c
                    End If
                    blnUpShift = (c = " ")
                Next
                txtBox.Text = tmpSt2
                Invalidate()
            End If
        End If
    End Sub

    Private Sub mnuUpperCase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuUpperCase.Click
        If TypeOf sender.parent.sourcecontrol Is TextBox Then
            sender.parent.sourcecontrol.text = UCase(sender.parent.sourcecontrol.text)
            Invalidate()
        End If
    End Sub


    Private Sub mfp_editPHONE_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editPHONE.Leave
        setPhones()
    End Sub

    Private Sub mfp_editFAX_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editFAX.Leave
        setPhones()
    End Sub


    Private Sub mnuBlockAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuBlockAddress.Click
        Dim strAddress As String

        strAddress = BlockAddress()
        Clipboard.SetDataObject(strAddress)
    End Sub

    Private Function BlockAddress() As String
        Dim strAddress As String
        Dim dr As DataRow()
        'Dim cnt As oneStopClasses.DSCountry.TOS_COUNTRYRow

        With myAddress
            If .AdrLine1.Length > 0 Then strAddress = strAddress & .AdrLine1 & "," & vbCrLf
            If .AdrLine2.Length > 0 Then strAddress = strAddress & .AdrLine2 & "," & vbCrLf
            If .AdrLine3.Length > 0 Then strAddress = strAddress & .AdrLine3 & vbCrLf
            If .AdrLine4.Length > 0 Then strAddress = strAddress & .AdrLine4 & vbCrLf
            Select Case .Country
                Case 45, 47, 49, 33, 65, 46, 30
                    If .Country = 65 Then
                        strAddress = strAddress & "SINGAPORE "
                    End If
                    If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " "
                    If .City.Length > 0 Then
                        strAddress = strAddress & .City & vbCrLf
                    Else
                        strAddress = strAddress & vbCrLf
                    End If
                    If .Country = 45 Then
                        strAddress = strAddress & "DENMARK" & vbCrLf
                    ElseIf .Country = 47 Then
                        strAddress = strAddress & "NORWAY" & vbCrLf
                    ElseIf .Country = 49 Then
                        strAddress = strAddress & "GERMANY" & vbCrLf
                    ElseIf .Country = 30 Then
                        strAddress = strAddress & "GREECE" & vbCrLf
                    ElseIf .Country = 33 Then
                        strAddress = strAddress & "FRANCE" & vbCrLf
                    ElseIf .Country = 46 Then
                        strAddress = strAddress & "SWEDEN" & vbCrLf
                    ElseIf .Country = 65 Then
                        strAddress = strAddress & "REPUBLIC OF SINGAPORE" & vbCrLf
                    End If
                Case 27
                    If .City.Length > 0 Then strAddress = strAddress & .City & vbCrLf
                    If .District.Length > 0 Then strAddress = strAddress & .District & vbCrLf
                    If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " SOUTH AFRICA" & vbCrLf

                Case 44
                    If .City.Length > 0 Then strAddress = strAddress & .City & vbCrLf
                    If .District.Length > 0 Then strAddress = strAddress & .District & vbCrLf
                    If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & vbCrLf
                Case Else
                    If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " "
                    If .City.Length > 0 Then
                        strAddress = strAddress & .City & vbCrLf
                    Else
                        strAddress = strAddress & vbCrLf
                    End If
                    Dim objCnt As MedscreenLib.Address.Country
                    objCnt = MedscreenLib.Glossary.Glossary.Countries.Item(.Country, 0)
                    If Not objCnt Is Nothing Then
                        strAddress = strAddress & UCase(objCnt.CountryName) & vbCrLf
                    End If

                    'dr = myGlossaries.Countrylist.TOS_COUNTRY.Select("Country_ID = " & .Country)
                    'If dr.Length = 1 Then
                    '    cnt = dr(0)
                    '    strAddress = strAddress & UCase(cnt.COUNTRY_NAME) & vbCrLf
                    'End If
            End Select
        End With
        Me.ToolTip1.SetToolTip(Me.Panel2, strAddress)
        Return strAddress
    End Function

    Private Sub Panel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel1.Click
        BlockAddress()
    End Sub

    Private Sub Panel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel1.Click
        BlockAddress()
    End Sub

    Private Sub mfp_editPOSTCODE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mfp_editPOSTCODE.TextChanged
        Me.ErrorProvider1.SetError(mfp_editPOSTCODE, "")
        If Me.cbCountry.SelectedValue = 44 Then
            If Not myAddress Is Nothing Then
                If Not myAddress.IsValidUKPostcode(mfp_editPOSTCODE.Text) Then
                    Me.ErrorProvider1.SetError(mfp_editPOSTCODE, myAddress.PostCodeError)
                End If
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Make a copy of this address id and use that 
    ''' </summary>
    ''' <param name="lngAddrId">Address Id to use</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CopyFrom(ByVal lngAddrId As Integer)
        Dim trow1 As MedscreenLib.Address.Caddress
        Dim lngAddressId As Integer

        AddressId = lngAddrId
        trow1 = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress()
        lngAddressId = trow1.AddressId
        With trow1
            .Status = myAddress.Status
            .AddressType = myAddress.AddressType
            .AdrLine1 = myAddress.AdrLine1
            .AdrLine2 = myAddress.AdrLine2
            .AdrLine3 = myAddress.AdrLine3
            .City = myAddress.City
            .District = myAddress.District
            .Email = myAddress.Email
            .Fax = myAddress.Fax
            .Phone = myAddress.Phone
            .PostCode = myAddress.PostCode

        End With
        myAddress = trow1

        Save()
        RaiseEvent NewAddress(lngAddressId)
    End Function

    Private Sub cmdDownone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDownone.Click
        Me.mfp_editADDRLINE4.Text = Me.mfp_editADDRLINE3.Text
        Me.mfp_editADDRLINE3.Text = Me.mfp_editADDRLINE2.Text
        Me.mfp_editADDRLINE2.Text = Me.mfp_editADDRLINE1.Text
        Me.mfp_editADDRLINE1.Text = ""
    End Sub
End Class

''' -----------------------------------------------------------------------------
''' Project	 : AddressControl
''' Class	 : AddressPanel
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Address display panel
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action>Show status property added</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class AddressPanel
    Inherits Windows.Forms.Panel
#Region "declarations"


    Public Enum HighlightedControl
        none
        Fax
        Email
        Line1
    End Enum

    Private ObjDataSource As MedscreenLib.Glossary.PhraseCollection             'DataSource for Address type

    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents cbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdNewAddress As System.Windows.Forms.Button
    Friend WithEvents cmdFindAddress As System.Windows.Forms.Button
    Friend WithEvents cmdEditAddress As System.Windows.Forms.Button
    Friend WithEvents cmdClone As System.Windows.Forms.Button
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuProperCase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUpperCase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLowerCase As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTipx As System.Windows.Forms.ToolTip
    'Friend WithEvents CXMenuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuBlockAddress As System.Windows.Forms.MenuItem
    Friend WithEvents mfp_editADDRLINE1 As System.Windows.Forms.TextBox
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE1 As System.Windows.Forms.Label
    Friend WithEvents mfp_editADDRLINE2 As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE2 As System.Windows.Forms.Label
    Friend WithEvents mfp_editADDRLINE3 As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblADDRLINE3 As System.Windows.Forms.Label
    Friend WithEvents mfp_lblDISTRICT As System.Windows.Forms.Label
    Friend WithEvents mfp_editDISTRICT As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblCITY As System.Windows.Forms.Label
    Friend WithEvents mfp_lblADDRESS_ID As System.Windows.Forms.Label
    Friend WithEvents mfp_lblPOSTCODE As System.Windows.Forms.Label
    Friend WithEvents mfp_lblCOUNTRY As System.Windows.Forms.Label
    Friend WithEvents mfp_editPOSTCODE As System.Windows.Forms.TextBox
    Friend WithEvents mfp_editFAX As System.Windows.Forms.TextBox
    Friend WithEvents mfp_editCity As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblEMAIL As System.Windows.Forms.Label
    Friend WithEvents mfp_editEMAIL As System.Windows.Forms.TextBox
    Friend WithEvents mfp_lblPHONE As System.Windows.Forms.Label
    Friend WithEvents mfp_lblFAX As System.Windows.Forms.Label
    Friend WithEvents mfp_editPHONE As System.Windows.Forms.TextBox
    Friend WithEvents cbAddrType As System.Windows.Forms.ComboBox
    Friend WithEvents lblAddrType As System.Windows.Forms.Label
    Friend WithEvents mfp_lblContact As System.Windows.Forms.Label
    Friend WithEvents mfp_editContact As System.Windows.Forms.TextBox
    Friend WithEvents mfp_editADDRLINE4 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdDownone As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents mnuPasteNew As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteReplace As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopy As System.Windows.Forms.MenuItem
    Friend WithEvents lblHighLight As System.Windows.Forms.Label
    Friend WithEvents lblAddressType As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel


    Private myAddress As MedscreenLib.Address.Caddress
    Private myCountries As MedscreenLib.Address.CountryCollection
    'Private myGlossaries As oneStopClasses.clsGlossary
    Private blnReadonly As Boolean = False

    Private blnShowContact As Boolean = False
    Private blnShowStatus As Boolean = True
    Private blnShowPanel As Boolean = False
    Private blnShowSave As Boolean = False
    Private blnShowNew As Boolean = False
    Private blnShowFind As Boolean = False
    Private blnShowClone As Boolean = False
    Private blnShowCancel As Boolean = False
    Private blnShowEdit As Boolean = False
    Private myAddressType As MedscreenLib.Constants.AddressType = MedscreenLib.Constants.AddressType.AddressMain

    Private myHighlightedControl As HighlightedControl = HighlightedControl.none
    Private myInfoVisible As Boolean = False

    Private intCmdSaveX As Integer = 180
    Private intcmdNewX As Integer = 8
    Private intCmdFindX As Integer = 8
    Private intcmdCloneX As Integer = 8
    Private intCmdCancelX As Integer = 8
    Private intCmdEditX As Integer = 8
#Region "Events"


    Public Event NewAddress(ByVal AddressId As Integer)
    Public Event Save(ByVal AddressID As Integer)
    Public Event Delete(ByVal AddressID As Integer)
    Public Event FindAddress(ByVal AddressID As Integer)
#End Region
#End Region

#Region "Public Instance"

#Region "Functions"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks that the contents form a valid address
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsValid() As Boolean
        Dim blnRet As Boolean = True
        Me.ErrorProvider1.SetError(Me.cbCountry, "")
        If Not myAddress Is Nothing Then
            If myAddress.Country = 44 Then
                If Not myAddress.IsValidUKPostcode(myAddress.PostCode) Then
                    blnRet = False
                    Exit Function
                End If
            ElseIf myAddress.Country <= 0 Then
                Me.ErrorProvider1.SetError(Me.cbCountry, "Country needs to be set")
                blnRet = False
                Exit Function
            End If

            'Check Phone number 
            Dim EmailError As MedscreenLib.Constants.EmailErrors = MedscreenLib.Medscreen.ValidateEmail(myAddress.Email)
            If Not (EmailError = MedscreenLib.Constants.EmailErrors.None Or EmailError = MedscreenLib.Constants.EmailErrors.NoAddress) Then
                blnRet = False

                If EmailError = MedscreenLib.Constants.EmailErrors.InvalidSet Then
                    Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "At least one address in the set is invalid")
                End If
                Exit Function
            End If
            If Not myAddress.ValidateFax = MedscreenLib.Constants.PhoneNoErrors.NoError Then
                blnRet = False
                Exit Function

            End If
            If Not myAddress.ValidatePhone = MedscreenLib.Constants.PhoneNoErrors.NoError Then
                blnRet = False
                Exit Function

            End If
        End If
        Return blnRet
    End Function
#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new address panel
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'Me.ToolTip1 = New System.Windows.Forms.ToolTip()

        Me.mfp_lblADDRLINE1 = New System.Windows.Forms.Label()
        Me.mfp_lblADDRLINE2 = New System.Windows.Forms.Label()
        Me.mfp_lblADDRLINE3 = New System.Windows.Forms.Label()
        Me.mfp_lblPOSTCODE = New System.Windows.Forms.Label()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        'Me.CXMenuMain = New System.Windows.Forms.ContextMenu()
        Me.mnuProperCase = New System.Windows.Forms.MenuItem()
        Me.mnuUpperCase = New System.Windows.Forms.MenuItem()
        Me.mnuLowerCase = New System.Windows.Forms.MenuItem()
        Me.mnuBlockAddress = New System.Windows.Forms.MenuItem()
        Me.mnuPasteNew = New System.Windows.Forms.MenuItem()
        Me.mnuPasteReplace = New System.Windows.Forms.MenuItem()
        Me.mnuCopy = New System.Windows.Forms.MenuItem()

        Me.mfp_lblPHONE = New System.Windows.Forms.Label()
        Me.mfp_lblCITY = New System.Windows.Forms.Label()
        Me.lblAddrType = New System.Windows.Forms.Label()
        Me.mfp_lblPOSTCODE = New System.Windows.Forms.Label()
        Me.mfp_lblCOUNTRY = New System.Windows.Forms.Label()
        Me.mfp_lblDISTRICT = New System.Windows.Forms.Label()
        Me.mfp_lblFAX = New System.Windows.Forms.Label()
        Me.mfp_lblEMAIL = New System.Windows.Forms.Label()
        Me.mfp_lblContact = New System.Windows.Forms.Label()
        Me.mfp_lblADDRESS_ID = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.mfp_editADDRLINE4 = New System.Windows.Forms.TextBox()


        Me.mfp_editADDRLINE1 = New System.Windows.Forms.TextBox()
        Me.mfp_editADDRLINE3 = New System.Windows.Forms.TextBox()
        Me.mfp_editADDRLINE2 = New System.Windows.Forms.TextBox()
        Me.mfp_editDISTRICT = New System.Windows.Forms.TextBox()
        Me.mfp_editEMAIL = New System.Windows.Forms.TextBox()
        Me.mfp_editContact = New System.Windows.Forms.TextBox()
        Me.mfp_editFAX = New System.Windows.Forms.TextBox()
        Me.mfp_editPHONE = New System.Windows.Forms.TextBox()
        Me.mfp_editPOSTCODE = New System.Windows.Forms.TextBox()
        Me.mfp_editCity = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblAddressType = New System.Windows.Forms.Label()
        Me.cmdDownone = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdNewAddress = New System.Windows.Forms.Button()
        Me.cmdFindAddress = New System.Windows.Forms.Button()
        Me.cmdEditAddress = New System.Windows.Forms.Button
        Me.cmdClone = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()

        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()


        Me.cbAddrType = New System.Windows.Forms.ComboBox()
        Me.cbStatus = New System.Windows.Forms.ComboBox()
        Me.cbCountry = New System.Windows.Forms.ComboBox()
        Me.txtInfo = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()

        Me.Controls.Add(Me.cbCountry)
        Me.Controls.Add(Me.Panel1)

        Me.Panel1.SuspendLayout()

        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdSave, Me.cmdNewAddress, cmdFindAddress, Me.cmdClone, Me.cmdCancel, Me.cmdDelete, Me.cmdEditAddress})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.DockPadding.All = 30
        Me.Panel1.Location = New System.Drawing.Point(0, 200)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(552, 30)
        Me.Panel1.TabIndex = 24

        Me.Panel1.Visible = blnShowPanel
        Me.Panel1.ResumeLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(8, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 27
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.Visible = False


        '
        'cmdNewAddress
        '
        Me.cmdNewAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewAddress.TextAlign = ContentAlignment.MiddleRight
        Me.cmdNewAddress.Location = New System.Drawing.Point(8, 4)
        Me.cmdNewAddress.Name = "cmdNewAddress"
        Me.cmdNewAddress.TabIndex = 25
        Me.cmdNewAddress.Text = "&New"
        Me.cmdNewAddress.Visible = Me.blnShowNew

        '
        'cmdFindAddress
        '
        Me.cmdFindAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindAddress.TextAlign = ContentAlignment.MiddleRight
        Me.cmdFindAddress.Location = New System.Drawing.Point(8, 4)
        Me.cmdFindAddress.Name = "cmdFindAddress"
        Me.cmdFindAddress.TabIndex = 100
        Me.cmdFindAddress.Text = "&Find"
        Me.cmdFindAddress.Visible = Me.blnShowFind

        '
        'cmdClone
        '
        Me.cmdClone.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClone.TextAlign = ContentAlignment.MiddleRight
        Me.cmdClone.Location = New System.Drawing.Point(8, 4)
        Me.cmdClone.Name = "cmdClone"
        Me.cmdClone.TabIndex = 100
        Me.cmdClone.Text = "&Clone"
        Me.cmdClone.Visible = Me.blnShowClone

        '
        'cmdEditAddress
        '
        Me.cmdEditAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEditAddress.TextAlign = ContentAlignment.MiddleRight
        Me.cmdEditAddress.Location = New System.Drawing.Point(8, 4)
        Me.cmdEditAddress.Name = "cmdEditAddress"
        Me.cmdEditAddress.TabIndex = 101
        Me.cmdEditAddress.Text = "&Edit"
        Me.cmdEditAddress.Visible = Me.blnShowEdit


        '
        'Panel2
        '
        Me.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.ContextMenu = Me.ContextMenu1

        Me.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Name = "Panel2"
        Me.Size = New System.Drawing.Size(552, 232)
        Me.TabIndex = 30

        'mfp_editADDRLINE1
        '
        Me.mfp_editADDRLINE1.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE1.Location = New System.Drawing.Point(56, 28)
        Me.mfp_editADDRLINE1.Name = "mfp_editADDRLINE1"
        Me.mfp_editADDRLINE1.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE1.TabIndex = 8
        Me.mfp_editADDRLINE1.Text = ""

        'mfp_lblADDRLINE1
        '
        Me.mfp_lblADDRLINE1.Location = New System.Drawing.Point(8, 28)
        Me.mfp_lblADDRLINE1.Name = "mfp_lblADDRLINE1"
        Me.mfp_lblADDRLINE1.Size = New System.Drawing.Size(38, 23)
        Me.mfp_lblADDRLINE1.TabIndex = 2
        Me.mfp_lblADDRLINE1.Text = "Line1"
        '
        'mfp_editADDRLINE2
        '
        Me.mfp_editADDRLINE2.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE2.Location = New System.Drawing.Point(56, 54)
        Me.mfp_editADDRLINE2.Name = "mfp_editADDRLINE2"
        Me.mfp_editADDRLINE2.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE2.TabIndex = 9
        Me.mfp_editADDRLINE2.Text = ""
        '
        'mfp_lblADDRLINE2
        '
        Me.mfp_lblADDRLINE2.Location = New System.Drawing.Point(8, 54)
        Me.mfp_lblADDRLINE2.Name = "mfp_lblADDRLINE2"
        Me.mfp_lblADDRLINE2.Size = New System.Drawing.Size(46, 23)
        Me.mfp_lblADDRLINE2.TabIndex = 3
        Me.mfp_lblADDRLINE2.Text = "Line2"

        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuProperCase, Me.mnuUpperCase, Me.mnuLowerCase, Me.mnuBlockAddress, Me.mnuCopy, Me.mnuPasteNew, Me.mnuPasteReplace})
        '
        'Me.CXMenuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuBlockAddress})

        '
        'mfp_editADDRLINE3
        '
        Me.mfp_editADDRLINE3.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE3.Location = New System.Drawing.Point(56, 81)
        Me.mfp_editADDRLINE3.Name = "mfp_editADDRLINE3"
        Me.mfp_editADDRLINE3.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE3.TabIndex = 10
        Me.mfp_editADDRLINE3.Text = ""
        '
        'mnuProperCase
        '
        Me.mnuProperCase.Index = 3
        Me.mnuProperCase.Text = "Proper Case"
        '
        'mnuUpperCase
        '
        Me.mnuUpperCase.Index = 1
        Me.mnuUpperCase.Text = "Upper Case"
        '
        'mnuLowerCase
        '
        Me.mnuLowerCase.Index = 2
        Me.mnuLowerCase.Text = "Lower Case"
        '
        '
        'mnuBlockAddress
        '
        Me.mnuBlockAddress.Index = 0
        Me.mnuBlockAddress.Text = "&Block Address"

        '
        '
        '
        Me.mnuPasteNew.Index = 5
        Me.mnuPasteNew.Text = "&Paste as new Address"
        Me.mnuPasteNew.Shortcut = Shortcut.CtrlShiftV
        Me.mnuPasteNew.ShowShortcut = True

        Me.mnuPasteReplace.Index = 6
        Me.mnuPasteReplace.Text = "&Paste replace Address"

        
        '
        'cmdSave
        '
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.TextAlign = ContentAlignment.MiddleRight
        Me.cmdSave.Location = New System.Drawing.Point(184, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.TabIndex = 28
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.Visible = Me.blnShowSave


        '
        'mnuCopy
        '
        Me.mnuCopy.Index = 4
        Me.mnuCopy.Text = "&Copy Address"
        Me.mnuCopy.Shortcut = Shortcut.CtrlShiftC
        Me.mnuCopy.ShowShortcut = True

        lblHighLight = New System.Windows.Forms.Label()
        Me.lblHighLight.Location = New System.Drawing.Point(8, 81)
        Me.lblHighLight.Name = "lblHightLight"
        Me.lblHighLight.Text = "*"
        Me.lblHighLight.Visible = False
        Me.lblHighLight.BackColor = System.Drawing.Color.Green
        Me.lblHighLight.ForeColor = System.Drawing.Color.Yellow
        Me.lblHighLight.Size = New System.Drawing.Size(15, 20)

        '
        'mfp_lblADDRLINE3
        '
        Me.mfp_lblADDRLINE3.Location = New System.Drawing.Point(8, 81)
        Me.mfp_lblADDRLINE3.Name = "mfp_lblADDRLINE3"
        Me.mfp_lblADDRLINE3.Size = New System.Drawing.Size(38, 23)
        Me.mfp_lblADDRLINE3.TabIndex = 4
        Me.mfp_lblADDRLINE3.Text = "Line 3"

        'mfp_editPHONE
        '
        Me.mfp_editPHONE.ContextMenu = Me.ContextMenu1
        Me.mfp_editPHONE.Location = New System.Drawing.Point(416, 40)
        Me.mfp_editPHONE.Name = "mfp_editPHONE"
        Me.mfp_editPHONE.TabIndex = 21
        Me.mfp_editPHONE.Text = ""
        Me.mfp_editPHONE.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left))                 '
        Me.mfp_editPHONE.Size = New System.Drawing.Size(130, 20)

        '
        'mfp_lblPHONE
        '
        Me.mfp_lblPHONE.Location = New System.Drawing.Point(344, 40)
        Me.mfp_lblPHONE.Name = "mfp_lblPHONE"
        Me.mfp_lblPHONE.Size = New System.Drawing.Size(64, 23)
        Me.mfp_lblPHONE.TabIndex = 15
        Me.mfp_lblPHONE.Text = "Phone"
        '
        'mfp_editCITY
        '
        Me.mfp_editCity.ContextMenu = Me.ContextMenu1
        Me.mfp_editCity.Location = New System.Drawing.Point(56, 127)
        Me.mfp_editCity.Name = "mfp_editCITY"
        Me.mfp_editCity.TabIndex = 11
        Me.mfp_editCity.Text = ""
        '
        'mfp_lblCITY
        '
        Me.mfp_lblCITY.Location = New System.Drawing.Point(8, 127)
        Me.mfp_lblCITY.Name = "mfp_lblCITY"
        Me.mfp_lblCITY.Size = New System.Drawing.Size(40, 23)
        Me.mfp_lblCITY.TabIndex = 5
        Me.mfp_lblCITY.Text = "City"
        '
        'mfp_lblPOSTCODE'
        Me.mfp_lblPOSTCODE.Location = New System.Drawing.Point(8, 154)
        Me.mfp_lblPOSTCODE.Name = "mfp_lblPOSTCODE"
        Me.mfp_lblPOSTCODE.Size = New System.Drawing.Size(56, 23)
        Me.mfp_lblPOSTCODE.TabIndex = 13
        Me.mfp_lblPOSTCODE.Text = "Post code"
        '
        'mfp_editFAX
        '
        Me.mfp_editFAX.Location = New System.Drawing.Point(416, 72)
        Me.mfp_editFAX.Name = "mfp_editFAX"
        Me.mfp_editFAX.TabIndex = 22
        Me.mfp_editFAX.Text = ""
        Me.mfp_editFAX.Size = New System.Drawing.Size(130, 20)
        Me.mfp_editFAX.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left))

        Me.txtInfo.Location = New System.Drawing.Point(550, 40)
        Me.txtInfo.Size = New System.Drawing.Size(150, 100)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Visible = False
        Me.txtInfo.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right))

        '
        '
        'mfp_editPOSTCODE
        '
        Me.mfp_editPOSTCODE.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.mfp_editPOSTCODE.Location = New System.Drawing.Point(64, 154)
        Me.mfp_editPOSTCODE.Name = "mfp_editPOSTCODE"
        Me.mfp_editPOSTCODE.Size = New System.Drawing.Size(88, 20)
        Me.mfp_editPOSTCODE.TabIndex = 19
        Me.mfp_editPOSTCODE.Text = ""
        '
        'mfp_editEMAIL
        '
        Me.mfp_editEMAIL.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                   Or System.Windows.Forms.AnchorStyles.Right)
        Me.mfp_editEMAIL.Location = New System.Drawing.Point(64, 182)
        Me.mfp_editEMAIL.Name = "mfp_editEMAIL"
        Me.mfp_editEMAIL.Size = New System.Drawing.Size(456, 20)
        Me.mfp_editEMAIL.TabIndex = 23
        Me.mfp_editEMAIL.Text = ""
        '
        'mfp_lblEMAIL
        '
        Me.mfp_lblEMAIL.Location = New System.Drawing.Point(8, 182)
        Me.mfp_lblEMAIL.Name = "mfp_lblEMAIL"
        Me.mfp_lblEMAIL.Size = New System.Drawing.Size(56, 23)
        Me.mfp_lblEMAIL.TabIndex = 17
        Me.mfp_lblEMAIL.Text = "EMAIL"
        '
        'mfp_editContact
        '
        Me.mfp_editContact.ContextMenu = Me.ContextMenu1

        Me.mfp_editContact.Location = New System.Drawing.Point(64, 218)
        Me.mfp_editContact.Name = "mfp_editEMAIL"
        Me.mfp_editContact.Size = New System.Drawing.Size(456, 20)
        Me.mfp_editContact.TabIndex = 231
        Me.mfp_editContact.Text = ""
        Me.mfp_editContact.Visible = False
        Me.mfp_editContact.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                   Or System.Windows.Forms.AnchorStyles.Right)                '
        'mfp_lblEMAIL
        '
        Me.mfp_lblContact.Location = New System.Drawing.Point(8, 218)
        Me.mfp_lblContact.Name = "mfp_lblContact"
        Me.mfp_lblContact.Size = New System.Drawing.Size(56, 23)
        Me.mfp_lblContact.TabIndex = 17
        Me.mfp_lblContact.Text = "Contact"
        Me.mfp_lblContact.Visible = False
        '
        'mfp_lblCOUNTRY
        '
        Me.mfp_lblCOUNTRY.Location = New System.Drawing.Point(168, 158)
        Me.mfp_lblCOUNTRY.Name = "mfp_lblCOUNTRY"
        Me.mfp_lblCOUNTRY.Size = New System.Drawing.Size(48, 23)
        Me.mfp_lblCOUNTRY.TabIndex = 14
        Me.mfp_lblCOUNTRY.Text = "Country"

        '
        'mfp_editDISTRICT
        '
        Me.mfp_editDISTRICT.ContextMenu = Me.ContextMenu1
        Me.mfp_editDISTRICT.Location = New System.Drawing.Point(216, 126)
        Me.mfp_editDISTRICT.Name = "mfp_editDISTRICT"
        Me.mfp_editDISTRICT.Size = New System.Drawing.Size(120, 20)
        Me.mfp_editDISTRICT.TabIndex = 12
        Me.mfp_editDISTRICT.Text = ""
        '
        'mfp_lblDISTRICT
        '
        Me.mfp_lblDISTRICT.Location = New System.Drawing.Point(168, 126)
        Me.mfp_lblDISTRICT.Name = "mfp_lblDISTRICT"
        Me.mfp_lblDISTRICT.Size = New System.Drawing.Size(40, 23)
        Me.mfp_lblDISTRICT.TabIndex = 6
        Me.mfp_lblDISTRICT.Text = "District"
        '
        'mfp_lblFAX
        '
        Me.mfp_lblFAX.Location = New System.Drawing.Point(344, 72)
        Me.mfp_lblFAX.Name = "mfp_lblFAX"
        Me.mfp_lblFAX.Size = New System.Drawing.Size(48, 23)
        Me.mfp_lblFAX.TabIndex = 16
        Me.mfp_lblFAX.Text = "FAX"
        '
        'mfp_lblADDRESS_ID
        '
        Me.mfp_lblADDRESS_ID.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mfp_lblADDRESS_ID.ForeColor = System.Drawing.SystemColors.Desktop
        Me.mfp_lblADDRESS_ID.Location = New System.Drawing.Point(16, 8)
        Me.mfp_lblADDRESS_ID.Name = "mfp_lblADDRESS_ID"
        Me.mfp_lblADDRESS_ID.Size = New System.Drawing.Size(200, 16)
        Me.mfp_lblADDRESS_ID.TabIndex = 1
        Me.mfp_lblADDRESS_ID.Text = "Address Id:"

        Me.lblAddressType.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddressType.Location = New System.Drawing.Point(400, 8)
        Me.lblAddressType.Text = "Address Type:"
        Me.lblAddressType.Size = New System.Drawing.Size(200, 16)
        Me.lblAddressType.AutoSize = True

        '
        'lblAddrType
        '
        Me.lblAddrType.Location = New System.Drawing.Point(344, 154)
        Me.lblAddrType.Name = "lblAddrType"
        Me.lblAddrType.Size = New System.Drawing.Size(64, 32)
        Me.lblAddrType.TabIndex = 27
        Me.lblAddrType.Text = "Type"

        '
        'cbCountry
        '
        Me.cbCountry.DisplayMember = "CountryName"
        Me.cbCountry.Location = New System.Drawing.Point(216, 154)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(120, 21)
        Me.cbCountry.TabIndex = 25
        Me.cbCountry.Text = ""
        Me.cbCountry.ValueMember = "CountryId"

        Me.cbAddrType.DisplayMember = "PhraseText"
        Me.cbAddrType.ValueMember = "PhraseID"
        Me.cbStatus.ValueMember = "STATUS_TYPE"
        Me.cbStatus.DisplayMember = "STATUS_TYPE_DESCRIPTION"


        '
        'cbAddrType
        '
        Me.cbAddrType.Location = New System.Drawing.Point(416, 154)
        Me.cbAddrType.Name = "cbAddrType"
        'Me.cbAddrType.Size = New System.Drawing.Size(104, 21)
        Me.cbAddrType.TabIndex = 26
        Me.cbAddrType.Text = ""
        Me.cbAddrType.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right))


        'Me.Controls.Add(Me.cbAddrType) '

        '
        'cbStatus
        '
        Me.cbStatus.Enabled = False
        Me.cbStatus.Location = New System.Drawing.Point(416, 136)
        Me.cbStatus.Name = "cbStatus"
        Me.cbStatus.Size = New System.Drawing.Size(104, 21)
        Me.cbStatus.TabIndex = 28
        Me.cbStatus.Text = ""
        Me.cbStatus.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                   Or System.Windows.Forms.AnchorStyles.Right)                '
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(344, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 32)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Status"


        '
        'mfp_editADDRLINE4
        '
        Me.mfp_editADDRLINE4.ContextMenu = Me.ContextMenu1
        Me.mfp_editADDRLINE4.Location = New System.Drawing.Point(56, 104)
        Me.mfp_editADDRLINE4.Name = "mfp_editADDRLINE4"
        Me.mfp_editADDRLINE4.Size = New System.Drawing.Size(280, 20)
        Me.mfp_editADDRLINE4.TabIndex = 32
        Me.mfp_editADDRLINE4.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 23)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Line 4"
        '
        'cmdDownone

        Dim strPath As String = Application.StartupPath & "\adc.resources"
        If IO.File.Exists(strPath) Then
            Dim aRS As Resources.ResourceReader = New Resources.ResourceReader(strPath)
            '
            Dim en As IDictionaryEnumerator = aRS.GetEnumerator()
            While en.MoveNext
                If en.Key = "cmdDownone.Image" Then
                    Me.cmdDownone.Image = CType(en.Value, System.Drawing.Bitmap)
                End If
            End While
        End If
        Me.cmdDownone.Location = New System.Drawing.Point(350, 0)
        Me.cmdDownone.Name = "cmdDownone"
        Me.cmdDownone.Size = New System.Drawing.Size(32, 23)
        Me.cmdDownone.TabIndex = 28
        Me.cmdDownone.Text = "V"

        '''' Add icons for controls
        Dim strIconPath As String = "\\corp.concateno.com\medscreen\common\Lab Programs\Icon Editors\ICONS\medscreen\Address\"

        If IO.Directory.Exists(strIconPath) Then
            Try
                Dim objIco As System.Drawing.Icon
                objIco = New Icon(strIconPath & "cmdSave.ico")
                Me.cmdSave.Image = objIco.ToBitmap()
                objIco = New Icon(strIconPath & "cmdNewAddress.ico")
                Me.cmdNewAddress.Image = objIco.ToBitmap
                objIco = New Icon(strIconPath & "search.ico")
                Me.cmdFindAddress.Image = objIco.ToBitmap
                objIco = New Icon(strIconPath & "AdrClone.ico")
                Me.cmdClone.Image = objIco.ToBitmap
                objIco = New Icon(strIconPath & "Cancel.ico")
                Me.cmdCancel.Image = objIco.ToBitmap
                objIco = New Icon(strIconPath & "edit.ico")
                Me.cmdEditAddress.Image = objIco.ToBitmap

            Catch
            End Try
        End If
        ObjDataSource = New MedscreenLib.Glossary.PhraseCollection("ADDRTYPE")
        ObjDataSource.Load()
        Me.Controls.AddRange(New System.Windows.Forms.Control() _
            {txtInfo, Me.lblHighLight, Me.lblAddressType, Me.mfp_editADDRLINE1, mfp_lblADDRLINE1, _
            Me.mfp_editADDRLINE2, mfp_lblADDRLINE2, _
            Me.mfp_editADDRLINE3, mfp_lblADDRLINE3, _
            Me.mfp_editADDRLINE4, Label2, Me.cmdDownone, _
            mfp_editPOSTCODE, mfp_lblPOSTCODE, _
            mfp_editEMAIL, mfp_lblEMAIL, mfp_editFAX, mfp_lblFAX, _
            mfp_lblCOUNTRY, mfp_editPHONE, mfp_lblPHONE, mfp_lblDISTRICT, mfp_editDISTRICT, _
            mfp_lblCITY, mfp_editCity, mfp_lblADDRESS_ID, lblAddrType, _
            cbCountry, cbAddrType, cbStatus, Label1, mfp_editContact, mfp_lblContact})

        Me.myCountries = MedscreenLib.Glossary.Glossary.Countries
        Me.cbAddrType.DataSource = ObjDataSource
        'Me.cbStatus.DataSource = Me.myGlossaries.StatusTypeList("A")
        Me.cbCountry.DataSource = myCountries
        SetPositions()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set errors from outside
    ''' </summary>
    ''' <param name="AControl"></param>
    ''' <param name="ErrorText"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetErrorText(ByVal Control As HighlightedControl, ByVal ErrorText As String)
        If Control = HighlightedControl.Email Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, ErrorText)
        ElseIf Control = HighlightedControl.Fax Then
            Me.ErrorProvider1.SetError(Me.mfp_editFAX, ErrorText)
        ElseIf Control = HighlightedControl.Line1 Then
            Me.ErrorProvider1.SetError(Me.mfp_editADDRLINE1, ErrorText)

        End If
    End Sub



#End Region

#Region "Properties"

    Public Property InfoVisible() As Boolean
        Get
            Return myInfoVisible
        End Get
        Set(ByVal Value As Boolean)
            myInfoVisible = Value
            Me.txtInfo.Visible = myInfoVisible
        End Set
    End Property

    Public Property AddressType() As MedscreenLib.Constants.AddressType
        Get
            Return myAddressType
        End Get
        Set(ByVal Value As MedscreenLib.Constants.AddressType)
            myAddressType = Value
            Dim strObj As String = [Enum].GetName(GetType(MedscreenLib.Constants.AddressType), Value)
            Me.lblAddressType.Text = "Address Type : " & strObj
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Tooltip to display
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Tooltip() As Tooltip
        Get
            Return ToolTipx
        End Get
        Set(ByVal Value As Tooltip)
            ToolTipx = Value
            If ToolTipx Is Nothing Then Exit Property
            Me.ToolTipx.SetToolTip(Me.mfp_editPHONE, "Phone number")
            Me.ToolTipx.SetToolTip(Me.mfp_editFAX, "Fax number")
            Me.ToolTipx.SetToolTip(Me.cmdDownone, "Move address lines 1-3 down 1")

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' HighLight A control 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property HiLiteControl() As HighlightedControl
        Get
            Return Me.myHighlightedControl
        End Get
        Set(ByVal Value As HighlightedControl)
            Me.myHighlightedControl = Value
            Select Case Me.myHighlightedControl
                Case HighlightedControl.none
                    Me.lblHighLight.Visible = False
                Case HighlightedControl.Email
                    Me.lblHighLight.Visible = True
                    Me.lblHighLight.Top = Me.mfp_editEMAIL.Top
                    Me.lblHighLight.Left = Me.mfp_editEMAIL.Left - Me.lblHighLight.Width
                Case HighlightedControl.Line1
                    Me.lblHighLight.Visible = True
                    Me.lblHighLight.Top = Me.mfp_editADDRLINE1.Top
                    Me.lblHighLight.Left = Me.mfp_editADDRLINE1.Left - Me.lblHighLight.Width
                Case HighlightedControl.Fax
                    Me.lblHighLight.Visible = True
                    Me.lblHighLight.Top = Me.mfp_editFAX.Top
                    Me.lblHighLight.Left = Me.mfp_editFAX.Left - Me.lblHighLight.Width

            End Select
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allow Contact to be shown
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowContact() As Boolean
        Get
            Return Me.blnShowContact
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowContact = Value
            Me.mfp_editContact.Visible = Value
            Me.mfp_lblContact.Visible = Value
            Me.mfp_editContact.Enabled = Not Me.blnReadonly
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allow status combo to be shown 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowStatus() As Boolean
        Get
            Return Me.blnShowStatus
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowStatus = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show command panel
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/03/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowPanel() As Boolean
        Get
            Return blnShowPanel
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowPanel = Value
            Me.Panel1.Visible = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Display save command
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/03/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowSaveCmd() As Boolean
        Get
            Return Me.blnShowSave
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowSave = Value
            Me.cmdSave.Visible = Value
            If Value Then Me.ShowPanel = True
            SetPositions()
        End Set
    End Property

    Public Property ShowCancelCmd() As Boolean
        Get
            Return Me.blnShowCancel
        End Get
        Set(ByVal Value As Boolean)
            blnShowCancel = value
            Me.cmdCancel.Visible = Value
            If Value Then ShowPanel = Visible
            SetPositions()
        End Set
    End Property

    Public Property ShowEditCmd() As Boolean
        Get
            Return Me.blnShowEdit
        End Get
        Set(ByVal value As Boolean)
            Me.blnShowEdit = value
            Me.cmdEditAddress.Visible = value
            If value Then ShowPanel = Visible
            SetPositions()
        End Set
    End Property

    Public Property ShowCloneCmd() As Boolean
        Get
            Return Me.blnShowClone
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowClone = Value
            Me.cmdClone.Visible = Value
            If Value Then Me.ShowPanel = True
            SetPositions()
        End Set
    End Property

    Public Property ShowFindCmd() As Boolean
        Get
            Return Me.blnShowFind
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowFind = Value
            Me.cmdFindAddress.Visible = Value
            If Value Then Me.ShowPanel = True
            SetPositions()
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' display new cmd
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/03/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowNewCmd() As Boolean
        Get
            Return Me.blnShowNew
        End Get
        Set(ByVal Value As Boolean)
            blnShowNew = Value
            If Value Then
                ShowPanel = True
            Else
            End If
            Me.cmdNewAddress.Visible = Value
            SetPositions()
        End Set
    End Property

    Private Sub SetPositions()
        If Me.blnShowNew Then
            Me.intCmdSaveX = 94

        Else
            Me.intCmdSaveX = 8
        End If
        If Me.blnShowSave Then
            Me.intCmdFindX = intCmdSaveX + 86
        Else
            Me.intCmdFindX = intCmdSaveX
        End If
        If Me.blnShowFind Then
            Me.intcmdCloneX = intCmdFindX + 86
        Else
            Me.intcmdCloneX = intCmdFindX
        End If
        If Me.blnShowCancel Then
            Me.intCmdCancelX = intcmdCloneX + 86
        Else
            Me.intCmdCancelX = intcmdCloneX
        End If

        If Me.blnShowEdit Then
            Me.intCmdEditX = intCmdCancelX + 86
        Else
            intCmdEditX = intCmdCancelX
        End If

        Me.cmdSave.Location = New Point(intCmdSaveX, Me.cmdSave.Location.Y)
        Me.cmdFindAddress.Location = New Point(intCmdFindX, Me.cmdFindAddress.Location.Y)
        Me.cmdClone.Location = New Point(intcmdCloneX, Me.cmdClone.Location.Y)
        Me.cmdCancel.Location = New Point(intCmdCancelX, Me.cmdCancel.Location.Y)
        Me.cmdEditAddress.Location = New Point(intCmdEditX, Me.cmdEditAddress.Location.Y)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address associated with this control
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Address() As MedscreenLib.Address.Caddress
        Get
            Me.GetAddressControl()
            Return myAddress
        End Get
        Set(ByVal Value As MedscreenLib.Address.Caddress)
            Try
                myAddress = Value
                If Value Is Nothing Then
                    EmptyAddressControl()
                Else
                    FillAddressControl()
                End If
            Catch ex As Exception
                LogError(ex, True, )
            End Try



        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the control read only 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property isReadOnly() As Boolean
        Get
            Return blnReadonly
        End Get
        Set(ByVal Value As Boolean)
            If Not Me.myAddress Is Nothing AndAlso Me.myAddress.Deleted AndAlso blnReadonly Then
                Exit Property               'Skip out if true 
            End If
            blnReadonly = Value
            Me.mfp_editADDRLINE1.Enabled = Not Value
            Me.mfp_editADDRLINE2.Enabled = Not Value
            Me.mfp_editADDRLINE3.Enabled = Not Value
            Me.mfp_editADDRLINE4.Enabled = Not Value
            Me.mfp_editCity.Enabled = Not Value
            Me.mfp_editDISTRICT.Enabled = Not Value
            Me.mfp_editEMAIL.Enabled = Not Value
            Me.mfp_editFAX.Enabled = Not Value
            Me.mfp_editPHONE.Enabled = Not Value
            Me.mfp_editPOSTCODE.Enabled = Not Value
            Me.mfp_editContact.Enabled = Not Value
            Me.cbAddrType.Enabled = Not Value
            Me.cbStatus.Enabled = Not Value
            Me.cbCountry.Enabled = Not Value
        End Set
    End Property

#End Region
#End Region

#Region "Private Instance"

#Region "Private Procedures"

    'Clear out address info 

    Private Sub EmptyAddressControl()
        Me.mfp_editADDRLINE1.Text = ""
        Me.mfp_editADDRLINE2.Text = ""
        Me.mfp_editADDRLINE3.Text = ""
        Me.mfp_editCity.Text = ""
        Me.mfp_editDISTRICT.Text = ""
        Me.mfp_editEMAIL.Text = ""
        Me.mfp_editPOSTCODE.Text = ""
        Me.mfp_editFAX.Text = ""
        Me.mfp_editPHONE.Text = ""
        Me.mfp_lblADDRESS_ID.Text = "Address Id: "
        Me.mfp_editContact.Text = ""
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the address control with the relevant information 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillAddressControl()
        'Dim intIndex As Integer
        'Dim gItem As MedscreenLib.Glossary.Phrase
        Try
            EmptyAddressControl()
            If myAddress.Deleted Then
                Me.mfp_lblADDRESS_ID.Text = "Address Id: " & myAddress.AddressId & " (deleted)"
                Me.isReadOnly = True
                Me.mfp_lblADDRESS_ID.ForeColor = Color.Red
            Else
                Me.isReadOnly = False
                Me.mfp_lblADDRESS_ID.Text = "Address Id: " & myAddress.AddressId
                Me.mfp_lblADDRESS_ID.ForeColor = System.Drawing.SystemColors.Desktop
            End If
            Me.mfp_editADDRLINE1.Text = myAddress.AdrLine1


            Me.mfp_editADDRLINE2.Text = myAddress.AdrLine2
            Me.mfp_editADDRLINE3.Text = myAddress.AdrLine3
            Me.mfp_editADDRLINE4.Text = myAddress.AdrLine4
            Me.mfp_editContact.Text = myAddress.Contact
            Me.mfp_editContact.MaxLength = Me.myAddress.Fields("CONTACT").FieldLength

            'If Not myAddress.Datatable Is Nothing Then
            '    Me.mfp_editADDRLINE1.MaxLength = myAddress.Fields.Item("ADDRLINE1").FieldLength
            '    Me.mfp_editADDRLINE2.MaxLength = myAddress.Fields.Item("ADDRLINE2").FieldLength
            '    Me.mfp_editADDRLINE3.MaxLength = myAddress.Fields.Item("ADDRLINE3").FieldLength

            'End If

            Me.mfp_editCity.Text = myAddress.City
            Me.mfp_editCity.MaxLength = Me.myAddress.Fields("CITY").FieldLength
            'If myGlossaries Is Nothing Then
            '    myGlossaries = New oneStopClasses.clsGlossary()
            'End If

            'If Not myGlossaries Is Nothing Then


            'End If
            Me.mfp_editDISTRICT.Text = myAddress.District
            Me.mfp_editDISTRICT.MaxLength = Me.myAddress.Fields("DISTRICT").FieldLength


            Me.mfp_editEMAIL.Text = myAddress.Email
            Me.mfp_editEMAIL.MaxLength = Me.myAddress.Fields("EMAIL").FieldLength

            Me.mfp_editFAX.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Fax, myAddress.Country)
            Me.mfp_editFAX.MaxLength = Me.myAddress.Fields("FAX").FieldLength
            Me.mfp_editPHONE.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Phone, myAddress.Country)
            Me.mfp_editPHONE.MaxLength = Me.myAddress.Fields("PHONE").FieldLength
            setPhones()

            Me.mfp_editPOSTCODE.Text = myAddress.PostCode



            'f MedscreenLib.Glossary.Glossary.AddressTypes.Count > 0 Then

            Try  ' Protecting
                Dim StrResp As String = CConnection.PackageStringList("LIB_ADDRESS.GETADDRESSLINKTYPES", myAddress.AddressId)
                If StrResp IsNot Nothing AndAlso StrResp.Trim.Length > 0 Then
                    Dim bits As String() = StrResp.Split(New Char() {"|"})
                    If bits.Length > 1 Then
                        Dim adrPhrase As MedscreenLib.Glossary.Phrase = objdatasource.Item(bits(0), Glossary.PhraseCollection.IndexBy.OrderBY)
                        If adrPhrase IsNot Nothing Then cbAddrType.SelectedValue = adrPhrase.PhraseID
                    End If
                Else
                    If myAddress.AddressType.Trim.Length > 0 Then
                        Dim intType As Integer? = Nothing
                        Try
                            intType = CInt(myAddress.AddressType)

                        Catch ex As Exception

                        End Try
                        If intType IsNot Nothing Then
                            Dim oPhrase As MedscreenLib.Glossary.Phrase = ObjDataSource.Item(CStr(intType), Glossary.PhraseCollection.IndexBy.OrderBY)
                            If oPhrase IsNot Nothing Then
                                cbAddrType.SelectedValue = oPhrase.PhraseID
                            End If
                        End If
                    End If
                End If
                'Me.cbAddrType.SelectedValue = Me.myAddress.AddressType
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "AddressPanel-FillAddressControl-1958")
            Finally
            End Try            'intIndex = MedscreenLib.Glossary.Glossary.AddressTypes.IndexOf(Me.myAddress.AddressType)
            'Me.cbAddrType.SelectedIndex = intIndex

            'If intIndex > -1 Then
            '    Dim objAddrType As MedscreenLib.Glossary.AddressType
            '    objAddrType = MedscreenLib.Glossary.Glossary.AddressTypes.Item(intIndex)
            '    Me.cbAddrType.SelectedIndex = intIndex
            '    'If Me.cbAddrType.SelectedIndex <> -1 Then
            '    '    Me.cbAddrType.SelectedIndex = 0
            '    '    intIndex = Intranet.intranet.support.cSupport.AddressTypes.IndexOf(Me.myAddress.AddressType)
            '    '    If intIndex > -1 Then
            '    '        Me.cbAddrType.SelectedIndex = intIndex
            '    '        Me.cbAddrType.SelectedText = objAddrType.Description
            '    '    End If
            '    'End If
            'End If
            'End If
            Try
                If myAddress.Country <= 0 Then
                    myAddress.Country = 44                  'Set undefined countries to UK
                End If
                If myAddress.Country > 0 Then
                    Dim objCont As MedscreenLib.Address.Country
                    Dim ik As Integer = 0
                    For Each objCont In Me.cbCountry.DataSource
                        If objCont.CountryId = myAddress.Country Then
                            Me.cbCountry.SelectedIndex = ik
                            Exit For
                        End If
                        ik += 1
                    Next
                    SetCountrySpecific(myAddress.Country)
                End If
            Catch
            End Try
            SetStatus(Me.myAddress.Status)
            IsValid()
            Dim StrOutput As String = ""
            Me.txtInfo.Text = ""

            If Me.Address.AddressId > 0 Then
                Dim strInfo As String = MedscreenLib.CConnection.PackageStringList("lib_address.GetAddressLinkTypes", CStr(Me.Address.AddressId))
                If Not strInfo Is Nothing Then
                    Dim strAddress As String() = strInfo.Split(New Char() {","})
                    Dim i As Integer
                    StrOutput = "Address used for " & vbCrLf
                    Dim strBits As String()
                    For i = 0 To strAddress.Length - 1
                        Dim strThis As String = strAddress.GetValue(i)
                        If strThis.Trim.Length > 0 Then
                            strBits = strThis.Split(New Char() {"|"})
                            If strBits.Length > 1 Then StrOutput += strBits.GetValue(1) & vbCrLf

                        End If
                    Next
                    'see if we only have one type
                    If strAddress.Length = 1 Then
                        If Not strBits Is Nothing Then
                            Try
                                Dim adrPhrase As MedscreenLib.Glossary.Phrase = ObjDataSource.Item(strBits(0), Glossary.PhraseCollection.IndexBy.OrderBY)
                                If adrPhrase IsNot Nothing Then cbAddrType.SelectedValue = adrPhrase.PhraseID
                            Catch
                            End Try
                        End If
                    End If
                    Me.txtInfo.Text = StrOutput
                    If Me.myInfoVisible Then 'Infocontrol visible so fill it 
                        Me.txtInfo.Show()
                    End If
                End If
            End If
        Catch ex As Exception
            LogError(ex, True, )
        End Try



        'Me.mfp_editADDRLINE1.Focus()
    End Sub

    Private Sub GetAddressControl()
        If myAddress Is Nothing Then Exit Sub
        myAddress.AdrLine1 = Me.mfp_editADDRLINE1.Text
        myAddress.AdrLine2 = Me.mfp_editADDRLINE2.Text
        myAddress.AdrLine3 = Me.mfp_editADDRLINE3.Text
        myAddress.AdrLine4 = Me.mfp_editADDRLINE4.Text

        myAddress.City = Me.mfp_editCity.Text

        myAddress.Country = Me.cbCountry.SelectedValue
        myAddress.District = Me.mfp_editDISTRICT.Text


        myAddress.Email = Me.mfp_editEMAIL.Text
        If myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.None Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.NoAt Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "Email address invalid no @ symbol")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.NoDomain Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "Email address invalid missing a domain")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.InvalidSet Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "At least one Email address is invalid")
        End If

        ' myAddress.Fax = Me.mfp_editFAX.Text
        myAddress.Fax = MedscreenLib.Medscreen.StripCountry(Me.mfp_editFAX.Text)
        myAddress.Phone = MedscreenLib.Medscreen.StripCountry(Me.mfp_editPHONE.Text)
        'myAddress.Phone = Me.mfp_editPHONE.Text
        myAddress.PostCode = Me.mfp_editPOSTCODE.Text
        If Not Me.cbAddrType.SelectedItem Is Nothing AndAlso TypeOf cbAddrType.SelectedItem Is MedscreenLib.Glossary.Phrase Then

            Me.myAddress.AddressType = TryCast(Me.cbAddrType.SelectedItem, MedscreenLib.Glossary.Phrase).OrderNum
        End If
        If Me.ShowContact Then
            Me.myAddress.Contact = Me.mfp_editContact.Text
        End If
    End Sub

    Private Sub setPhones()
        If ToolTipx Is Nothing Then Exit Sub
        Me.ToolTipx.SetToolTip(Me.mfp_editPHONE, "")
        Me.ToolTipx.SetToolTip(Me.mfp_editFAX, "")

        If Not myAddress.Phone.Length > 0 Then
            Me.ToolTipx.SetToolTip(Me.mfp_editPHONE, mfp_editPHONE.Text)
        End If
        If Not myAddress.Fax.Length > 0 Then
            ToolTipx.SetToolTip(Me.mfp_editFAX, "00" & myAddress.Country & " " & mfp_editFAX.Text)
        End If

    End Sub

    Private Sub SetCountrySpecific(ByVal intCountry As Integer)
        Select Case intCountry
            Case Intranet.intranet.Support.cSupport.CountryCodes.CountryUK
                Me.mfp_lblDISTRICT.Text = "County"
                Me.mfp_lblPOSTCODE.Text = "Postcode"
            Case Intranet.intranet.Support.cSupport.CountryCodes.CountryUSA
                Me.mfp_lblDISTRICT.Text = "State"
                Me.mfp_lblPOSTCODE.Text = "Zip code"
            Case Else
                Me.mfp_lblDISTRICT.Text = "District"
                Me.mfp_lblPOSTCODE.Text = "Postal code"

        End Select
        Me.mfp_editPhone_Leave(Nothing, Nothing)
        Me.mfp_editFAX_Leave(Nothing, Nothing)
        Me.mfp_editPOSTCODE_Leave(Nothing, Nothing)
    End Sub

    Private Sub SetStatus(ByVal strStatusValue As String)
        'Dim dr As oneStopClasses.DSStatusType.TOS_STATUS_TYPERow
        'Dim intRow As Integer = 0

    End Sub

    Private Sub AddressPanel_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.VisibleChanged
        Me.mfp_editContact.Visible = Me.blnShowContact
        Me.mfp_lblContact.Visible = Me.blnShowContact
        Me.cbStatus.Visible = Me.blnShowStatus
        Me.Label1.Visible = Me.blnShowStatus
    End Sub

    Private Sub cmdDownone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDownone.Click
        Me.mfp_editADDRLINE4.Text = Me.mfp_editADDRLINE3.Text
        Me.mfp_editADDRLINE3.Text = Me.mfp_editADDRLINE2.Text
        Me.mfp_editADDRLINE2.Text = Me.mfp_editADDRLINE1.Text
        Me.mfp_editADDRLINE1.Text = ""
    End Sub

    Private Sub mnuBlockAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuBlockAddress.Click
        Dim strAddress As String

        strAddress = BlockAddress()
        Clipboard.SetDataObject(strAddress)
    End Sub
    Private Sub mnuProperCase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuProperCase.Click
        Dim tmpString As String
        Dim tmpSt2 As String
        Dim i As Integer
        Dim blnUpShift As Boolean
        Dim optMen As ContextMenu
        Dim txtBox As TextBox

        If TypeOf sender Is MenuItem Then
            optMen = sender.parent
            If TypeOf optMen.SourceControl Is Panel Or TypeOf optMen.SourceControl Is TextBox Then
                If TypeOf optMen.SourceControl Is TextBox Then   'Check to see if it is a textbox
                    Dim myText As TextBox = optMen.SourceControl
                    If Not (myText.Name Like "*ADDRLINE*") Then  'Is the text box one of the top lines
                        ShiftToProperCase(myText)                'No just alter this line
                    Else
                        Shift1to4()
                    End If
                Else
                    Shift1to4()
                End If
                'Need to add checks for the text box calling.
            End If
            Invalidate()

        End If
    End Sub

    Private Sub Shift1to4()
        ShiftToProperCase(Me.mfp_editADDRLINE1)
        ShiftToProperCase(Me.mfp_editADDRLINE2)
        ShiftToProperCase(Me.mfp_editADDRLINE3)
        ShiftToProperCase(Me.mfp_editADDRLINE4)

    End Sub

    Private Sub mnuPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPasteNew.Click, mnuPasteReplace.Click
        Dim optMen As ContextMenu
        Dim optMenu As MenuItem
        If TypeOf sender Is MenuItem Then           'Get sender
            optMenu = sender                        'Save for later use
            optMen = sender.parent                  'Find source control 
            If TypeOf optMen.SourceControl Is Panel Or TypeOf optMen.SourceControl Is TextBox Then
                ' Want to ensure am not in a text box
                ' Declares an IDataObject to hold the data returned from the clipboard.
                ' Retrieves the data from the clipboard.
                Dim iData As IDataObject = Clipboard.GetDataObject()
                Dim MedAddressFormat As DataFormats.Format = DataFormats.GetFormat("MedAddressFormat")
                If iData.GetDataPresent("MedAddressFormat", True) Then
                    'translate the XML object to a string 
                    Dim tmpAddress As String = CType(iData.GetData(MedAddressFormat.Name), String)
                    'Create an XML DOM object 
                    Dim mydoc As Xml.XmlDocument = New Xml.XmlDocument()
                    mydoc.LoadXml(tmpAddress)           'Populate it 
                    Dim iNode As Xml.XmlNode = mydoc.ChildNodes.Item(0) 'Get the inner node data
                    Dim myTmpAddress As New MedscreenLib.Address.Caddress()
                    myTmpAddress.Fields.readfields(iNode.ChildNodes)    'Translate XML back to data 
                    'See which menu item has been called
                    If optMenu.Index = Me.mnuPasteReplace.Index Then 'Doing a replace copy address id into tmpaddress
                        myTmpAddress.AddressId = Me.Address.AddressId
                        myTmpAddress.Fields.RowID = Me.Address.Fields.RowID
                        myTmpAddress.Fields.InitialiseOldValues()
                        myTmpAddress.Update()
                        Me.Address = myTmpAddress
                    Else    'Get new address ID 
                        myTmpAddress.AddressId = MedscreenLib.Address.AddressCollection.NextAddressId
                        myTmpAddress.Fields.Insert(MedscreenLib.CConnection.DbConnection)
                        Me.Address = myTmpAddress
                        Me.Address.Update()
                        RaiseEvent NewAddress(myTmpAddress.AddressId)       'Propogate change upwards
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub mnuCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        Dim optMen As ContextMenu
        If TypeOf sender Is MenuItem Then       'Sender should be a menu item
            optMen = sender.parent              'Get context menu, so we can see who called it 
            If TypeOf optMen.SourceControl Is Panel Or TypeOf optMen.SourceControl Is TextBox Then
                ' Declares an IDataObject to hold the data returned from the clipboard.
                ' Retrieves the data from the clipboard.
                ' Creates a new data format.
                Dim MedAddressFormat As DataFormats.Format = DataFormats.GetFormat("MedAddressFormat")

                Dim myDataObject As New DataObject(MedAddressFormat.Name, Me.Address.Fields.ToXMLSchema)
                ' Copies XML representation of address into the clipboard.
                Clipboard.SetDataObject(myDataObject)
            End If
        End If
    End Sub

    Private Sub ShiftToProperCase(ByRef txtbox As TextBox)
        Dim tmpSt2 As String = ""
        Dim tmpString As String = LCase(txtbox.Text)
        Dim blnUpShift As Boolean = True
        Dim I As Integer
        For I = 1 To Len(tmpString)
            Dim c As String
            c = Mid(tmpString, I, 1)
            If blnUpShift Then
                tmpSt2 = tmpSt2 & UCase(c)
            Else
                tmpSt2 = tmpSt2 & c
            End If
            blnUpShift = (c = " " Or c = "/" Or c = "(" Or c = ".")
        Next
        txtbox.Text = tmpSt2

    End Sub

    Private Sub mnuUpperCase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuUpperCase.Click
        Dim optMen As ContextMenu
        Dim txtBox As TextBox
        If TypeOf sender Is MenuItem Then
            optMen = sender.parent
            If TypeOf optMen.SourceControl Is Panel Then
                Me.mfp_editADDRLINE1.Text = Me.mfp_editADDRLINE1.Text.ToUpper
                Me.mfp_editADDRLINE2.Text = Me.mfp_editADDRLINE2.Text.ToUpper
                Me.mfp_editADDRLINE3.Text = Me.mfp_editADDRLINE3.Text.ToUpper
                Me.mfp_editADDRLINE4.Text = Me.mfp_editADDRLINE4.Text.ToUpper
            ElseIf TypeOf optMen.SourceControl Is TextBox Then
                Dim aTxtBox As TextBox = optMen.SourceControl
                aTxtBox.SelectedText = aTxtBox.SelectedText.ToUpper
            End If
            'sender.parent.sourcecontrol.text = UCase(sender.text)
            Invalidate()
        End If
    End Sub

    Private Sub mnuLowerCase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuLowerCase.Click
        Dim optMen As ContextMenu
        Dim txtBox As TextBox
        If TypeOf sender Is MenuItem Then
            optMen = sender.parent
            If TypeOf optMen.SourceControl Is Panel Then
                Me.mfp_editADDRLINE1.Text = Me.mfp_editADDRLINE1.Text.ToLower
                Me.mfp_editADDRLINE2.Text = Me.mfp_editADDRLINE2.Text.ToLower
                Me.mfp_editADDRLINE3.Text = Me.mfp_editADDRLINE3.Text.ToLower
                Me.mfp_editADDRLINE4.Text = Me.mfp_editADDRLINE4.Text.ToLower
            ElseIf TypeOf optMen.SourceControl Is TextBox Then
                Dim aTxtBox As TextBox = optMen.SourceControl
                aTxtBox.SelectedText = aTxtBox.SelectedText.ToLower
            End If
            'sender.parent.sourcecontrol.text = UCase(sender.text)
            Invalidate()
        End If
    End Sub

    Private Function BlockAddress() As String
        Dim strAddress As String
        'Dim dr As DataRow()
        ''Dim cnt As oneStopClasses.DSCountry.TOS_COUNTRYRow

        'If myAddress Is Nothing Then Exit Function
        'With myAddress
        '    If .AdrLine1.Length > 0 Then strAddress = strAddress & .AdrLine1 & "," & vbCrLf
        '    If .AdrLine2.Length > 0 Then strAddress = strAddress & .AdrLine2 & "," & vbCrLf
        '    If .AdrLine3.Length > 0 Then strAddress = strAddress & .AdrLine3 & vbCrLf
        '    If .AdrLine4.Length > 0 Then strAddress = strAddress & .AdrLine4 & vbCrLf
        '    Select Case .Country
        '        Case 45, 47, 49, 33, 65, 46, 30
        '            If .Country = 65 Then
        '                strAddress = strAddress & "SINGAPORE "
        '            End If
        '            If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " "
        '            If .City.Length > 0 Then
        '                strAddress = strAddress & .City & vbCrLf
        '            Else
        '                strAddress = strAddress & vbCrLf
        '            End If
        '            If .Country = 45 Then
        '                strAddress = strAddress & "DENMARK" & vbCrLf
        '            ElseIf .Country = 47 Then
        '                strAddress = strAddress & "NORWAY" & vbCrLf
        '            ElseIf .Country = 49 Then
        '                strAddress = strAddress & "GERMANY" & vbCrLf
        '            ElseIf .Country = 30 Then
        '                strAddress = strAddress & "GREECE" & vbCrLf
        '            ElseIf .Country = 33 Then
        '                strAddress = strAddress & "FRANCE" & vbCrLf
        '            ElseIf .Country = 46 Then
        '                strAddress = strAddress & "SWEDEN" & vbCrLf
        '            ElseIf .Country = 65 Then
        '                strAddress = strAddress & "REPUBLIC OF SINGAPORE" & vbCrLf
        '            End If
        '        Case 27
        '            If .City.Length > 0 Then strAddress = strAddress & .City & vbCrLf
        '            If .District.Length > 0 Then strAddress = strAddress & .District & vbCrLf
        '            If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " SOUTH AFRICA" & vbCrLf

        '        Case 44
        '            If .City.Length > 0 Then strAddress = strAddress & .City & vbCrLf
        '            If .District.Length > 0 Then strAddress = strAddress & .District & vbCrLf
        '            If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & vbCrLf
        '        Case Else
        '            If .PostCode.Length > 0 Then strAddress = strAddress & .PostCode & " "
        '            If .City.Length > 0 Then
        '                strAddress = strAddress & .City & vbCrLf
        '            Else
        '                strAddress = strAddress & vbCrLf
        '            End If
        '            strAddress += .CountryName & vbCrLf
        '            Dim objCnt As MedscreenLib.Address.Country
        '            objCnt = MedscreenLib.Glossary.Glossary.Countries.Item(.Country, 0)
        '            If Not objCnt Is Nothing Then
        '                strAddress = strAddress & UCase(objCnt.CountryName) & vbCrLf
        '            End If

        '            'dr = myGlossaries.Countrylist.TOS_COUNTRY.Select("Country_ID = " & .Country)
        '            'If dr.Length = 1 Then
        '            '    cnt = dr(0)
        '            '    strAddress = strAddress & UCase(cnt.COUNTRY_NAME) & vbCrLf
        '            'End If
        '    End Select

        'End With
        Dim oParams As New Collection()
        oParams.Add(CStr(Me.Address.AddressId))
        oParams.Add(vbLf)
        strAddress = MedscreenLib.CConnection.PackageStringList("lib_address.GetAddressLabel", oParams)
        If Not ToolTipx Is Nothing Then
            Me.ToolTipx.SetToolTip(Me, strAddress)
        End If

        Return strAddress
    End Function

    Private Sub mfp_editEMAIL_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editEMAIL.Leave

        Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "")
        If Me.myAddress Is Nothing Then Exit Sub
        Me.myAddress.Email = Me.mfp_editEMAIL.Text
        If myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.None Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.NoAt Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "Email address invalid no @ symbol")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.NoDomain Then
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "Email address invalid missing a domain")
        ElseIf myAddress.ValidateEmail = MedscreenLib.Constants.EmailErrors.InvalidSet Then
            Dim intPos As Integer = MedscreenLib.Medscreen.isValidEmailSet(mfp_editEMAIL.Text)
            Me.ErrorProvider1.SetError(Me.mfp_editEMAIL, "Email address no " & CStr(intPos + 1) & " invalid in set")
        End If

    End Sub

    Private Sub mfp_editFAX_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editFAX.Leave
        If Me.myAddress Is Nothing Then Exit Sub
        Me.myAddress.Fax = MedscreenLib.Medscreen.StripCountry(Me.mfp_editFAX.Text, CStr(myAddress.Country))
        Me.mfp_editFAX.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Fax, myAddress.Country)

        If myAddress.ValidateFax = MedscreenLib.Constants.PhoneNoErrors.NoError Then
            Me.ErrorProvider1.SetError(Me.mfp_editFAX, "")
        ElseIf myAddress.ValidateFax = MedscreenLib.Constants.PhoneNoErrors.IllegalCharacterPresent Then
            Me.ErrorProvider1.SetError(Me.mfp_editFAX, "An illegal character present in the phone number, only digits or spaces allowed")
        End If

    End Sub

    Private Sub mfp_editPhone_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editPHONE.Leave
        If Me.myAddress Is Nothing Then Exit Sub
        Me.myAddress.Phone = MedscreenLib.Medscreen.StripCountry(Me.mfp_editPHONE.Text, CStr(myAddress.Country))
        Me.mfp_editPHONE.Text = MedscreenLib.Medscreen.AddCountry(myAddress.Phone, myAddress.Country)

        If myAddress.ValidatePhone = MedscreenLib.Constants.PhoneNoErrors.NoError Then
            Me.ErrorProvider1.SetError(Me.mfp_editPHONE, "")
        ElseIf myAddress.ValidatePhone = MedscreenLib.Constants.PhoneNoErrors.IllegalCharacterPresent Then
            Me.ErrorProvider1.SetError(Me.mfp_editPHONE, "An illegal character present in the phone number, only digits or spaces allowed")
        End If

    End Sub

    Private Sub ContextMenu1_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles ContextMenu1.Popup
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim hasAddress As Boolean = iData.GetDataPresent("MedAddressFormat", True)
        Me.mnuPasteNew.Enabled = False
        Me.mnuPasteReplace.Enabled = False
        If Me.Address Is Nothing OrElse Me.Address.AddressId = 0 Then
            Me.mnuCopy.Enabled = False
            If hasAddress Then
                Me.mnuPasteNew.Enabled = True
            End If
        Else
            Me.mnuCopy.Enabled = True
            If hasAddress Then
                Me.mnuPasteReplace.Enabled = True
            End If
        End If

    End Sub
#End Region
#End Region

    Private Sub cbCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCountry.SelectedIndexChanged
        If Not Me.Address Is Nothing Then
            myAddress.Country = Me.cbCountry.SelectedValue
            SetCountrySpecific(myAddress.Country)
            IsValid()
        End If

    End Sub

    Private Sub mfp_editPOSTCODE_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles mfp_editPOSTCODE.Leave
        Me.ErrorProvider1.SetError(Me.mfp_editPOSTCODE, "")
        If Not myAddress Is Nothing AndAlso myAddress.Country = 44 Then
            If Not myAddress.IsValidUKPostcode(Me.mfp_editPOSTCODE.Text) Then
                Me.ErrorProvider1.SetError(Me.mfp_editPOSTCODE, myAddress.PostCodeError)
            End If
        End If
    End Sub

    
    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
        Me.mfp_editADDRLINE1.Focus()
        Me.mfp_editADDRLINE1.SelectionStart = 1
        Me.mfp_editADDRLINE1.SelectionLength = 0
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        SaveAddress()
    End Sub

    Private Sub cmdNewAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNewAddress.Click
        myAddress = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress()
        'Me.DsAddress2 = tRow
        Me.FillAddressControl()
        RaiseEvent NewAddress(myAddress.AddressId)
    End Sub

    Private Sub cmdFindAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindAddress.Click
        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                myAddress = New MedscreenLib.Address.Caddress(intAddrID)
                RaiseEvent FindAddress(intAddrID)
                Me.FillAddressControl()
            End If
        End With
    End Sub

    Private Sub cmdClone_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClone.Click
        Dim oCmd As New OleDb.OleDbCommand("Lib_address.CloneAddress", MedscreenLib.CConnection.DbConnection)
        Dim intRet As Integer
        Try
            If CConnection.ConnOpen Then
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.Parameters.Add(CConnection.IntegerParameter("AddressId", Me.Address.AddressId))
                oCmd.Parameters.Add(CConnection.StringParameter("", Glossary.Glossary.CurrentSMUser.Identity, 10))
                intRet = oCmd.ExecuteNonQuery()
                If intRet = 1 Then
                    intRet = CConnection.PackageStringList("Lib_address.GetClonedAddress", "")
                    Me.Address = New Address.Caddress(intRet)
                    RaiseEvent NewAddress(intRet)
                End If
            End If
        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Address.Fields.Rollback()
        Me.FillAddressControl()

    End Sub

    Private Sub cmdEditAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEditAddress.Click
        Me.isReadOnly = False

        ShowSaveCmd = True

    End Sub

    Public Sub SaveAddress()
        Me.GetAddressControl()
        If myAddress IsNot Nothing Then myAddress.Update()

        RaiseEvent Save(myAddress.AddressId)
    End Sub
End Class