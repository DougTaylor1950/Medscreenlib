'$Revision: 1.2 $
'$Author: taylor $
'$Date: 2005-06-17 06:59:41+01 $
'$Log: frmAddress.vb,v $
'Revision 1.2  2005-06-17 06:59:41+01  taylor
'Milestone check in

'Revision 1.1  2004-05-04 06:30:38+01  taylor
'Fixed address set property to display Address
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports MedscreenCommonGui
Imports Intranet.intranet.Support
Imports MedscreenLib
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmAddressx
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for the maintenance of addresses
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmAddress
    Inherits System.Windows.Forms.Form

    Private MySMID As String = ""
    Private myProfile As String = ""

#Region " Windows Form Designer generated code "
    Private blnDrawing As Boolean = True
    Friend WithEvents txtInvShipAddr As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtInvMailAddr As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Private objSiblings As Intranet.intranet.customerns.SMIDProfileCollection
    Private AddrTypes As New MedscreenLib.Glossary.PhraseCollection("ADDRTYPE")

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.AddressPanel1.Tooltip = Me.ToolTip1

        'Me.cbSMID.DataSource = Intranet.intranet.Support.cSupport.SMIDS
        'Me.cbSMID.DisplayMember = "SMID"
        'Me.cbSMID.ValueMember = "SMID"

        Me.cbProfile.DataSource = Nothing
        'Me.cbProfile.ValueMember = "Identity"
        'Me.cbProfile.DisplayMember = "Info"
        AddrTypes.Load()
        Me.cbADRLink.DataSource = AddrTypes
        Me.cbADRLink.BindingContext(Me.cbADRLink.DataSource).SuspendBinding()
        Me.cbADRLink.BindingContext(Me.cbADRLink.DataSource).ResumeBinding()
        Me.cbADRLink.ValueMember = "PhraseID"
        Me.cbADRLink.DisplayMember = "PhraseText"

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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents AddressPanel1 As MedscreenCommonGui.AddressPanel
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents txtSiteLoc As System.Windows.Forms.TextBox
    Friend WithEvents lblSageId As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSiteName As System.Windows.Forms.TextBox
    Friend WithEvents lblAddressID As System.Windows.Forms.Label
    Friend WithEvents txtAddressID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblAdrLink As System.Windows.Forms.Label
    Friend WithEvents cbADRLink As System.Windows.Forms.ComboBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblSMID As System.Windows.Forms.Label
    Friend WithEvents cbSMID As MedscreenCommonGui.Combos.SMIDCombo
    Friend WithEvents cbProfile As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtMailID As System.Windows.Forms.TextBox
    Friend WithEvents txtInvAddrID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtShipAddrID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents txtSameAddress As System.Windows.Forms.RichTextBox
    Friend WithEvents lblSiteLoc As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddress))
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtInvShipAddr = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtInvMailAddr = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtShipAddrID = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtInvAddrID = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtMailID = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cbProfile = New System.Windows.Forms.ComboBox
        Me.cbSMID = New MedscreenCommonGui.Combos.SMIDCombo
        Me.lblSMID = New System.Windows.Forms.Label
        Me.lblAdrLink = New System.Windows.Forms.Label
        Me.lblSiteLoc = New System.Windows.Forms.Label
        Me.txtAddressID = New System.Windows.Forms.TextBox
        Me.txtSiteName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblSageId = New System.Windows.Forms.Label
        Me.txtSiteLoc = New System.Windows.Forms.TextBox
        Me.lblCustomer = New System.Windows.Forms.Label
        Me.cbADRLink = New System.Windows.Forms.ComboBox
        Me.lblAddressID = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.AddressPanel1 = New MedscreenCommonGui.AddressPanel
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.txtSameAddress = New System.Windows.Forms.RichTextBox
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtInvShipAddr)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.txtInvMailAddr)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.txtShipAddrID)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txtInvAddrID)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtMailID)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.cbProfile)
        Me.Panel2.Controls.Add(Me.cbSMID)
        Me.Panel2.Controls.Add(Me.lblSMID)
        Me.Panel2.Controls.Add(Me.lblAdrLink)
        Me.Panel2.Controls.Add(Me.lblSiteLoc)
        Me.Panel2.Controls.Add(Me.txtAddressID)
        Me.Panel2.Controls.Add(Me.txtSiteName)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.lblSageId)
        Me.Panel2.Controls.Add(Me.txtSiteLoc)
        Me.Panel2.Controls.Add(Me.lblCustomer)
        Me.Panel2.Controls.Add(Me.cbADRLink)
        Me.Panel2.Controls.Add(Me.lblAddressID)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(936, 196)
        Me.Panel2.TabIndex = 35
        '
        'txtInvShipAddr
        '
        Me.txtInvShipAddr.Location = New System.Drawing.Point(344, 98)
        Me.txtInvShipAddr.Name = "txtInvShipAddr"
        Me.txtInvShipAddr.Size = New System.Drawing.Size(72, 20)
        Me.txtInvShipAddr.TabIndex = 25
        Me.txtInvShipAddr.Tag = "14"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(240, 98)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 30)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "Invoice Ship Address"
        '
        'txtInvMailAddr
        '
        Me.txtInvMailAddr.Location = New System.Drawing.Point(128, 98)
        Me.txtInvMailAddr.Name = "txtInvMailAddr"
        Me.txtInvMailAddr.Size = New System.Drawing.Size(72, 20)
        Me.txtInvMailAddr.TabIndex = 23
        Me.txtInvMailAddr.Tag = "13"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 98)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 30)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Invoice Mail Address"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.BackColor = System.Drawing.SystemColors.Info
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 145)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(904, 43)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "You can select a customer address by Selecting the SMID, then profile then addres" & _
            "s type.  You can also double click on an address id"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtShipAddrID
        '
        Me.txtShipAddrID.Location = New System.Drawing.Point(544, 72)
        Me.txtShipAddrID.Name = "txtShipAddrID"
        Me.txtShipAddrID.Size = New System.Drawing.Size(72, 20)
        Me.txtShipAddrID.TabIndex = 20
        Me.txtShipAddrID.Tag = "2"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(440, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 24)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Secondary Address"
        '
        'txtInvAddrID
        '
        Me.txtInvAddrID.Location = New System.Drawing.Point(344, 72)
        Me.txtInvAddrID.Name = "txtInvAddrID"
        Me.txtInvAddrID.Size = New System.Drawing.Size(72, 20)
        Me.txtInvAddrID.TabIndex = 18
        Me.txtInvAddrID.Tag = "3"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(240, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Invoice Address"
        '
        'txtMailID
        '
        Me.txtMailID.Location = New System.Drawing.Point(128, 72)
        Me.txtMailID.Name = "txtMailID"
        Me.txtMailID.Size = New System.Drawing.Size(72, 20)
        Me.txtMailID.TabIndex = 16
        Me.txtMailID.Tag = "1"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Results Address"
        '
        'cbProfile
        '
        Me.cbProfile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbProfile.DisplayMember = "Info"
        Me.cbProfile.Location = New System.Drawing.Point(112, 40)
        Me.cbProfile.MaxDropDownItems = 10
        Me.cbProfile.Name = "cbProfile"
        Me.cbProfile.Size = New System.Drawing.Size(672, 21)
        Me.cbProfile.TabIndex = 14
        Me.cbProfile.ValueMember = "CustomerID"
        '
        'cbSMID
        '
        Me.cbSMID.DisplayMember = "SMID"
        Me.cbSMID.Location = New System.Drawing.Point(512, 8)
        Me.cbSMID.Name = "cbSMID"
        Me.cbSMID.Siblings = Nothing
        Me.cbSMID.Size = New System.Drawing.Size(121, 21)
        Me.cbSMID.TabIndex = 13
        Me.cbSMID.ValueMember = "SMID"
        '
        'lblSMID
        '
        Me.lblSMID.Location = New System.Drawing.Point(456, 8)
        Me.lblSMID.Name = "lblSMID"
        Me.lblSMID.Size = New System.Drawing.Size(48, 23)
        Me.lblSMID.TabIndex = 12
        Me.lblSMID.Text = "SMID"
        '
        'lblAdrLink
        '
        Me.lblAdrLink.Location = New System.Drawing.Point(208, 8)
        Me.lblAdrLink.Name = "lblAdrLink"
        Me.lblAdrLink.Size = New System.Drawing.Size(80, 32)
        Me.lblAdrLink.TabIndex = 10
        Me.lblAdrLink.Text = "Address Link Type"
        '
        'lblSiteLoc
        '
        Me.lblSiteLoc.Location = New System.Drawing.Point(632, 72)
        Me.lblSiteLoc.Name = "lblSiteLoc"
        Me.lblSiteLoc.Size = New System.Drawing.Size(100, 23)
        Me.lblSiteLoc.TabIndex = 1
        Me.lblSiteLoc.Text = "Site Location"
        Me.lblSiteLoc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddressID
        '
        Me.txtAddressID.Location = New System.Drawing.Point(104, 8)
        Me.txtAddressID.Name = "txtAddressID"
        Me.txtAddressID.Size = New System.Drawing.Size(100, 20)
        Me.txtAddressID.TabIndex = 7
        '
        'txtSiteName
        '
        Me.txtSiteName.Location = New System.Drawing.Point(369, 40)
        Me.txtSiteName.Name = "txtSiteName"
        Me.txtSiteName.Size = New System.Drawing.Size(135, 20)
        Me.txtSiteName.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(248, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Site Name"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSageId
        '
        Me.lblSageId.Location = New System.Drawing.Point(792, 40)
        Me.lblSageId.Name = "lblSageId"
        Me.lblSageId.Size = New System.Drawing.Size(124, 23)
        Me.lblSageId.TabIndex = 3
        Me.lblSageId.Text = "Sage ID :"
        '
        'txtSiteLoc
        '
        Me.txtSiteLoc.Location = New System.Drawing.Point(752, 72)
        Me.txtSiteLoc.Name = "txtSiteLoc"
        Me.txtSiteLoc.Size = New System.Drawing.Size(136, 20)
        Me.txtSiteLoc.TabIndex = 2
        '
        'lblCustomer
        '
        Me.lblCustomer.Location = New System.Drawing.Point(664, 8)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.Size = New System.Drawing.Size(256, 23)
        Me.lblCustomer.TabIndex = 0
        Me.lblCustomer.Text = "Customer :"
        '
        'cbADRLink
        '
        Me.cbADRLink.Location = New System.Drawing.Point(288, 8)
        Me.cbADRLink.Name = "cbADRLink"
        Me.cbADRLink.Size = New System.Drawing.Size(160, 21)
        Me.cbADRLink.TabIndex = 11
        '
        'lblAddressID
        '
        Me.lblAddressID.Location = New System.Drawing.Point(24, 8)
        Me.lblAddressID.Name = "lblAddressID"
        Me.lblAddressID.Size = New System.Drawing.Size(100, 23)
        Me.lblAddressID.TabIndex = 6
        Me.lblAddressID.Text = "Address Id"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Customer"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AddressPanel1
        '
        Me.AddressPanel1.Address = Nothing
        Me.AddressPanel1.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AddressPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AddressPanel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.AddressPanel1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AddressPanel1.InfoVisible = False
        Me.AddressPanel1.isReadOnly = False
        Me.AddressPanel1.Location = New System.Drawing.Point(0, 196)
        Me.AddressPanel1.Name = "AddressPanel1"
        Me.AddressPanel1.ShowCancelCmd = True
        Me.AddressPanel1.ShowCloneCmd = True
        Me.AddressPanel1.ShowContact = True
        Me.AddressPanel1.ShowEditCmd = True
        Me.AddressPanel1.ShowFindCmd = True
        Me.AddressPanel1.ShowNewCmd = True
        Me.AddressPanel1.ShowPanel = True
        Me.AddressPanel1.ShowSaveCmd = True
        Me.AddressPanel1.ShowStatus = True
        Me.AddressPanel1.Size = New System.Drawing.Size(680, 322)
        Me.AddressPanel1.TabIndex = 36
        Me.AddressPanel1.Tooltip = Nothing
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 518)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(936, 56)
        Me.Panel1.TabIndex = 37
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(844, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(764, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 8
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(680, 196)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 322)
        Me.Splitter1.TabIndex = 38
        Me.Splitter1.TabStop = False
        '
        'txtSameAddress
        '
        Me.txtSameAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSameAddress.Location = New System.Drawing.Point(683, 196)
        Me.txtSameAddress.Name = "txtSameAddress"
        Me.txtSameAddress.Size = New System.Drawing.Size(253, 322)
        Me.txtSameAddress.TabIndex = 39
        Me.txtSameAddress.Text = ""
        '
        'frmAddress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(936, 574)
        Me.Controls.Add(Me.txtSameAddress)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.AddressPanel1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAddress"
        Me.Text = "Address Maintenance"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible address dialog types 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum AddressDialogTypes
        ''' <summary>Show Site address</summary>
        site
        ''' <summary>Edit mode</summary>
        edit
    End Enum

    Private myAddress As MedscreenLib.Address.Caddress
    Private myClient As Intranet.intranet.customerns.Client
    Private strSiteName As String
    Private strSiteLocale As String
    Private FrmType As AddressDialogTypes = AddressDialogTypes.site

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Customer site locale
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Locale() As String
        Get
            strSiteLocale = Me.txtSiteLoc.Text
            Return strSiteLocale
        End Get
        Set(ByVal Value As String)
            strSiteLocale = Value
            Me.txtSiteLoc.Text = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the form type and change display to suit
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AddressFormType() As AddressDialogTypes
        Get
            Return FrmType
        End Get
        Set(ByVal Value As AddressDialogTypes)
            FrmType = Value
            If FrmType = AddressDialogTypes.site Then       'If site type show various items
                Me.lblCustomer.Show()
                Me.lblSageId.Show()
                Me.lblSiteLoc.Show()
                Me.Label2.Show()
                Me.txtSiteLoc.Show()
                Me.txtSiteName.Show()
                Me.lblAddressID.Hide()
                Me.txtAddressID.Hide()
                Me.Label3.Hide()
                Me.cbProfile.Hide()
                Me.cbSMID.Hide()
                Me.lblSMID.Hide()
                Me.lblAdrLink.Hide()
                Me.cbADRLink.Hide()
                'Me.Button2.Hide()
            Else
                Me.lblCustomer.Hide()
                Me.lblSageId.Hide()
                Me.lblSiteLoc.Hide()
                Me.Label2.Hide()
                Me.txtSiteLoc.Hide()
                Me.txtSiteName.Hide()
                Me.lblAddressID.Show()
                Me.txtAddressID.Show()
                Me.Label3.Show()
                Me.cbProfile.Show()
                Me.lblAdrLink.Show()
                Me.cbADRLink.Show()
                'Me.Button2.Show()
                Me.cbSMID.Show()
                Me.lblSMID.Show()
                Me.cbProfile.SelectedIndex = -1
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Client who owns this address if relevant
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            myClient = Value
            If Not Value Is Nothing Then
                Me.blnDrawing = True
                Me.cbSMID.SelectedValue = myClient.SMID
                Me.cbSMID_SelectedIndexChanged(Nothing, Nothing)
                Me.setupProfileCombo("")
                Me.cbProfile.SelectedValue = myClient.Identity
                'e.blnDrawing = False
                Me.cbCustomer_SelectedIndexChanged(Nothing, Nothing)
                Me.lblCustomer.Text = "Company ID : "
                Me.lblSageId.Text = "Sage Id : "
                If myClient Is Nothing Then Exit Property

                Me.lblCustomer.Text = "Company ID : " & myClient.CompanyName
                Me.lblSageId.Text = "Sage Id : " & myClient.InvoiceInfo.SageID
                Me.lblSiteLoc.Hide()
                Me.txtSiteLoc.Hide()
                SetCustomerAddresses()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Name of site
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SiteName() As String
        Get
            strSiteName = Me.txtSiteName.Text
            Return Me.strSiteName
        End Get
        Set(ByVal Value As String)
            strSiteName = Value
            Me.txtSiteName.Text = Value
            Me.lblSiteLoc.Show()
            Me.txtSiteLoc.Show()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address id of the address to be displayed 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AddressID() As Integer
        Get
            If myAddress Is Nothing Then
                Return -1
            Else
                Return myAddress.AddressId
            End If
        End Get
        Set(ByVal Value As Integer)
            If Value > 0 Then           'If Id is valid get address and display it 
                myAddress = cSupport.SMAddresses.item(Value)
                Me.AddressPanel1.Address = myAddress
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create an address 
    ''' </summary>
    ''' <param name="Line1">First line of address</param>
    ''' <param name="AdrType">Type of address (Alpha)</param>
    ''' <param name="ADRLink">Address type Numeric</param>
    ''' <returns>Address ID</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CreateAddress(Optional ByVal Line1 As String = "<NewAddress>", _
    Optional ByVal AdrType As String = "CUST", Optional ByVal ADRLink As Integer = 1) As Integer
        Try
            myAddress = cSupport.SMAddresses.CreateAddress(Line1, AdrType, , ADRLink) 'Get new address
            If Not myAddress Is Nothing Then
                Me.AddressPanel1.Address = myAddress                            'Update GUI
                Return myAddress.AddressId
            Else
                MsgBox("error in creating address")
                Return -1
            End If

        Catch ex As Exception
            Return -1
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Do new address action
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub CmdNewAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.FrmType = AddressDialogTypes.edit Then
            If Me.cbSMID.SelectedIndex < 1 Then
                MsgBox("You need to select a customer first")
                Exit Sub
            End If
            Dim strCustomerId As String = ""
            strCustomerId = Me.cbProfile.SelectedValue
            If strCustomerId.Trim.Length = 0 Then
                MsgBox("You need to select a customer first")
                Exit Sub
            End If
            If Me.CurrentTag = -1 Then Exit Sub
            myAddress = cSupport.SMAddresses.CreateAddress(, , strCustomerId, Me.CurrentTag)       'Get new address


            If Not myAddress Is Nothing Then
                Me.AddressPanel1.Address = myAddress                'Update GUI
                Me.txtAddressID.Text = myAddress.AddressId
                If Not Me.myClient Is Nothing Then
                    If CurrentTag = 1 Then
                        myClient.MailAddressId = myAddress.AddressId
                    ElseIf CurrentTag = 2 Then
                        myClient.ShippingAddressId = myAddress.AddressId
                    Else
                        myClient.InvoiceAddressId = myAddress.AddressId
                    End If
                    myClient.Update()
                End If
            Else
                MsgBox("error in creating address")
            End If
            'End If
        Catch ex As Exception
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Save address to database 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    cmdSave.Enabled = False
    '    myAddress = Me.AddressPanel1.Address
    '    If Not myAddress Is Nothing Then                'Check for valid address
    '        myAddress.Update()                          'Do update 
    '    End If
    '    cmdSave.Enabled = True                          'Reenable button 
    '    Me.AddressPanel1.Focus()                        'Reset focus to address

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exit form 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        AddressPanel1.Address.Update()
        AddressPanel1.SaveAddress()
        'Me.cmdSave_Click(Nothing, Nothing)
        If Not myClient Is Nothing Then
            myClient.Update()
        End If
    End Sub
    Private TargetAddressId As Integer? = Nothing
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address id has changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtAddressID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddressID.Leave
        Try
            TargetAddressId = Nothing
            Me.cbProfile.Enabled = True
            If Me.txtAddressID.Text.Trim.Length > 0 Then        'If we have an id, get the address that matches
                myAddress = cSupport.SMAddresses.item(Me.txtAddressID.Text, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                If Not myAddress Is Nothing Then                'If address is valid
                    TargetAddressId = CInt(txtAddressID.Text)
                    Me.AddressPanel1.Address = myAddress        'Populate control
                    If myAddress.Usage = 0 Then                 'If not used elsewhere allow editing 
                        'AddressPanel1.cmdEditAddres.Enabled = True
                        Me.AddressPanel1.isReadOnly = True      'Set  to read only edit will enable writing  
                    Else
                        Me.AddressPanel1.isReadOnly = True
                    End If
                    If myAddress.Customer.Trim.Length > 0 Then  'If there is a customer
                        AlignCustomer(myAddress.Customer)
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' To aid in customer editing 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/06/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property CustomerID() As String
        Set(ByVal Value As String)
            AlignCustomer(Value)
        End Set
    End Property

    Private Sub AlignCustomer(ByVal CustID As String)
        Dim objClient As Intranet.intranet.customerns.Client = Intranet.intranet.Support.cSupport.CustList.Item(CustID)
        'now get the smid aligned 
        Me.cbProfile.Enabled = True
        If Not objClient Is Nothing Then
            Dim objSmid As Intranet.intranet.customerns.SMID = _
                Intranet.intranet.Support.cSupport.SMIDS.Item(objClient.SMID)
            If Not objSmid Is Nothing Then
                Dim intRet As Integer = Intranet.intranet.Support.cSupport.SMIDS.IndexOf(objSmid)
                If intRet > -1 Then
                    Me.cbSMID.SelectedIndex = intRet
                    Dim objClient2 As Intranet.intranet.customerns.SMIDProfile
                    intRet = 0
                    For Each objClient2 In objSiblings
                        If objClient2.CustomerID = objClient.Identity Then
                            Me.cbProfile.SelectedIndex = intRet
                            Exit For
                        End If
                        intRet += 1
                    Next
                    Me.cbProfile.Enabled = False
                    Me.cbProfile.SelectionStart = 0
                    Me.cbProfile.SelectionLength = 0
                End If
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Copy the address 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdCopyAddress2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim Faddr As New MedscreenCommonGui.FindAddress()

    '    'use the findaddress method to find an address to copy 
    '    If Faddr.ShowDialog = DialogResult.OK Then
    '        Dim intId As Integer = Faddr.AddressID          'Get the address id to copy 
    '        Dim oAddr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.Copy(intId)  'Create a copy (new ID)
    '        Me.myAddress = oAddr                            'Set this address to be current
    '        Me.AddressPanel1.Address = oAddr                'Show in GUI
    '        'If myAddress.Usage = 0 Then                     'This must be 0 
    '        Me.cmdEditAddress.Enabled = True                'allow editing
    '        Me.AddressPanel1.isReadOnly = False             'make read write
    '        'Else
    '        '    Me.AddressPanel1.isReadOnly = True
    '        'End If
    '    End If
    'End Sub

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
    Private Sub cmdSearchAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Faddr As New FindAddress()

        If Faddr.ShowDialog = DialogResult.OK Then
            Dim intId As Integer = Faddr.AddressID          'Get found address Id
            Dim oAddr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(intId)
            Me.myAddress = oAddr                            'Show in GUI
            Me.AddressPanel1.Address = oAddr
            If myAddress Is Nothing Then Exit Sub '          If nowt skip out

            Me.AddressPanel1.isReadOnly = True
            If myAddress.Usage <= 1 Then                    'If not used 
                'Me.cmdEditAddress.Enabled = True            'Allow editing 
            End If
            If myAddress.Customer.Trim.Length > 0 Then      'Set up customer control 
                AlignCustomer(myAddress.Customer)
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit the address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdEditAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    'If myAddress.Usage <= 1 Then
    '    Me.AddressPanel1.isReadOnly = False         'Set control to read write
    '    'Me.cmdSave.Enabled = True                   'Allow saving 
    '    'End If

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create an address from the sample manager addresses for this customer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdCopySMAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim blnMatch As Boolean = False
    '    Dim ADRID As Integer
    '    Dim adrLink As Integer
    '    Try
    '        If Me.FrmType = AddressDialogTypes.edit Then            'If editing allowed
    '            If Me.cbProfile.SelectedIndex < 1 Then             'Check customer selection 
    '                MsgBox("You need to select a customer first")
    '                Exit Sub
    '            End If
    '            Dim smClient As Intranet.intranet.customerns.Client = _
    '                Me.cbProfile.SelectedProfile
    '            If smClient Is Nothing Then
    '                MsgBox("You need to select a customer first")
    '                Exit Sub
    '            End If
    '            If Not smClient Is Nothing Then
    '                Select Case Me.DomainUpDown1.SelectedIndex
    '                    Case 0  'Results Address (mail)
    '                        'blnMatch = (smClient.MailAddressId <> -1)
    '                        'ADRID = smClient.MailAddressId
    '                        adrLink = 1
    '                    Case 1  'Invoice Address (mail)
    '                        blnMatch = (smClient.InvoiceAddressId <> -1)
    '                        ADRID = smClient.InvoiceAddressId
    '                        adrLink = 3
    '                    Case 2  'Shipping Address (mail)
    '                        'blnMatch = (smClient.ShippingAddressId <> -1)
    '                        'ADRID = smClient.ShippingAddressId
    '                        'adrLink = 2
    '                End Select
    '                If blnMatch Then
    '                    MsgBox("This address has already been copied (id : " & ADRID & ") This will be shown in the address panel", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)
    '                    myAddress = cSupport.SMAddresses.item(ADRID)
    '                    Me.AddressPanel1.Address = myAddress
    '                    Exit Sub
    '                End If
    '            End If
    '            myAddress = cSupport.SMAddresses.CreateAddress(, , smClient.Identity, adrLink)     'Get new address
    '            If Not myAddress Is Nothing Then
    '                Dim intTempID As Integer = myAddress.AddressId
    '                Select Case adrLink
    '                    Case 1
    '                        'myAddress = smClient.MainAddress
    '                        'smClient.MailAddressId = intTempID
    '                        'Me.txtMailID.Text = intTempID
    '                        'Me.rbMailAddress.Checked = True
    '                    Case 2
    '                        'myAddress = smClient.SecondaryAddress
    '                        'smClient.ShippingAddressId = intTempID
    '                        'Me.txtShipAddrID.Text = intTempID
    '                        'Me.RBShipAddress.Checked = True
    '                    Case 3
    '                        myAddress = smClient.InvoiceInfo.InvoiceAddress
    '                        smClient.InvoiceAddressId = intTempID
    '                        Me.txtInvAddrID.Text = intTempID
    '                        Me.RBInvAdd.Checked = True
    '                End Select
    '                myAddress.AddressId = intTempID
    '                myAddress.Update()
    '                Me.AddressPanel1.Address = myAddress                'Update GUI
    '                smClient.Update()
    '                Me.txtAddressID.Text = myAddress.AddressId
    '            Else
    '                MsgBox("error in creating address")
    '            End If
    '        End If
    '    Catch ex As Exception
    '    End Try


    'End Sub

    Private Sub cbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbProfile.SelectedIndexChanged
        If blnDrawing Then Exit Sub
        'clear out info in case we find nothing 
        Me.AddressPanel1.Address = Nothing
        Me.txtAddressID.Text = ""
        'Me.cbADRLink.SelectedIndex = 0
        If cbProfile.SelectedIndex > -1 Then
            myClient = objSiblings.Item(Me.cbProfile.SelectedIndex).Client
            SetCustomerAddresses()
        End If
    End Sub

    Private Sub SetCustomerAddresses()
        Me.txtInvAddrID.Text = ""
        Me.txtMailID.Text = ""
        Me.txtShipAddrID.Text = ""
        Dim strAddressMismatch As String = "lib_Customer.SharedMailAddress"
        Dim strXML As String = CConnection.PackageStringList(strAddressMismatch, Client.Identity)
        Dim strText As String = Medscreen.ResolveStyleSheet(strXML, "PairedAddresstxt1.xsl", 0)
        Me.txtSameAddress.Rtf = strText

        If Not myClient Is Nothing Then
            'If Not myClient.ExtraInfo Is Nothing Then
            If myClient.InvoiceAddressId > -1 Then Me.txtInvAddrID.Text = myClient.InvoiceAddressId
            'Set up results address
            If myClient.MailAddressId > -1 Then
                Me.txtMailID.Text = myClient.MailAddressId
                'Me.txtAddressID.Text = myClient.MailAddressId
                'Me.txtAddressID_Leave(Nothing, Nothing)
                'Me.cbADRLink.SelectedIndex = 1
            End If
            If myClient.ShippingAddressId > -1 Then Me.txtShipAddrID.Text = myClient.ShippingAddressId
            If myClient.InvoiceMailAddressID > -1 Then Me.txtInvMailAddr.Text = myClient.InvoiceMailAddressID
            If myClient.InvoiceShippingAddressId > -1 Then Me.txtInvShipAddr.Text = myClient.InvoiceShippingAddressId
            'End 'If
        End If
        If TargetAddressId IsNot Nothing Then
            txtAddressID.Text = TargetAddressId
            If myAddress IsNot Nothing Then
                AddressPanel1.Address = myAddress
                Try  ' Protecting
                    Dim StrResp As String = CConnection.PackageStringList("LIB_ADDRESS.GETADDRESSLINKTYPES", myAddress.AddressId)
                    If StrResp IsNot Nothing AndAlso StrResp.Trim.Length > 0 Then
                        Dim bits As String() = StrResp.Split(New Char() {"|"})
                        If bits.Length > 1 Then
                            Dim adrPhrase As MedscreenLib.Glossary.Phrase = AddrTypes.Item(bits(0), Glossary.PhraseCollection.IndexBy.OrderBY)
                            If adrPhrase IsNot Nothing Then cbADRLink.SelectedValue = adrPhrase.PhraseID
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
        End If
    End Sub
    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawing = False
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address id box left 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub txtShipAddrID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    Try
    '        If Me.txtShipAddrID.Text.Trim.Length = 0 Then
    '            myClient.ShippingAddressId = -1
    '        Else
    '            myClient.ShippingAddressId = CInt(Me.txtShipAddrID.Text)
    '        End If

    '    Catch ex As Exception
    '    End Try

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address id box left
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub txtInvAddrID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    Try
    '        If Me.txtInvAddrID.Text.Trim.Length = 0 Then
    '            myClient.InvoiceAddressId = -1
    '        Else
    '            myClient.InvoiceAddressId = CInt(Me.txtInvAddrID.Text)
    '        End If

    '    Catch ex As Exception
    '    End Try

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address id box left
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub txtMailID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    Try
    '        If Me.txtMailID.Text.Trim.Length = 0 Then
    '            myClient.MailAddressId = -1
    '        Else
    '            myClient.MailAddressId = CInt(Me.txtMailID.Text)
    '        End If

    '    Catch ex As Exception
    '    End Try

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checked box changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub rbMailAddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    If rbMailAddress.Checked = True Then
    '        myAddress = cSupport.SMAddresses(myClient.MailAddressId)
    '        Me.AddressPanel1.Address = myAddress
    '    End If
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checked box changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub RBShipAddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    If RBShipAddress.Checked = True Then
    '        myAddress = cSupport.SMAddresses(myClient.ShippingAddressId)
    '        Me.AddressPanel1.Address = myAddress
    '    End If

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checked box changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub RBInvAdd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If myClient Is Nothing Then Exit Sub
    '    'If myClient.ExtraInfo Is Nothing Then Exit Sub
    '    If RBInvAdd.Checked = True Then
    '        myAddress = cSupport.SMAddresses(myClient.InvoiceAddressId)
    '        Me.AddressPanel1.Address = myAddress
    '    End If

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find out new smid and set up profile
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbSMID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSMID.SelectedIndexChanged
        If Me.blnDrawing Then Exit Sub
        'Dim objSMID As Intranet.intranet.customerns.SMID = Intranet.intranet.Support.cSupport.SMIDS.Item(Me.cbSMID.SelectedValue)
        'If objSMID Is Nothing Then Exit Sub
        'clear out info in case we find nothing 
        Me.AddressPanel1.Address = Nothing
        Me.txtAddressID.Text = ""
        Me.setupProfileCombo("")


    End Sub

    Private Sub setupProfileCombo(ByVal SMIDID As String)
        Me.cbProfile.Enabled = True
        'objSiblings = ModTpanel.SMClients.Siblings(SMIDID, "xxx", Intranet.intranet.customerns.SMClientCollection.SiblingTypes.SMIDSibling)
        objSiblings = Me.cbSMID.Siblings
        'Me.cbProfile.DataSource = objSiblings
        Me.cbProfile.DataSource = objSiblings
        Me.cbProfile.BindingContext(Me.cbProfile.DataSource).SuspendBinding()
        Me.cbProfile.BindingContext(Me.cbProfile.DataSource).ResumeBinding()
        Me.cbProfile.ValueMember = "CustomerID"
        Me.cbProfile.DisplayMember = "Info"
        Me.cbProfile.Enabled = (objSiblings.Count > 0)
    End Sub

    Private Sub cbADRLink_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbADRLink.SelectedIndexChanged
        If blnDrawing Then Exit Sub

        'clear out info in case we find nothing 
        Me.AddressPanel1.Address = Nothing
        Me.txtAddressID.Text = ""

        'Find out which address the user is after
        Dim strAddrLinkID As String
        If cbADRLink.SelectedItem IsNot Nothing Then
            strAddrLinkID = TryCast(cbADRLink.SelectedItem, MedscreenLib.Glossary.Phrase).OrderNum
        End If
        Dim oColl As New Collection()
        oColl.Add(Me.myClient.Identity)
        oColl.Add(strAddrLinkID)
        oColl.Add("Customer.Identity")
        Dim strAddressId As String = MedscreenLib.CConnection.PackageStringList("Lib_address.GetCustomerAddressId", oColl)
        If Not strAddressId Is Nothing AndAlso strAddressId.Trim.Length > 0 Then
            Me.txtAddressID.Text = strAddressId
            Me.txtAddressID_Leave(Nothing, Nothing)
        Else
            If CInt(strAddrLinkID) = 1 AndAlso myClient.MailAddressId > 0 Then
                Me.txtAddressID.Text = myClient.MailAddressId
                Me.txtAddressID_Leave(Nothing, Nothing)
                If Not Me.AddressPanel1.Address Is Nothing Then
                    Me.AddressPanel1.Address.SetAddressLink(myClient.Identity, 1)
                End If
            ElseIf CInt(strAddrLinkID) = 3 AndAlso myClient.InvoiceAddressId > 0 Then
                Me.txtAddressID.Text = myClient.InvoiceAddressId
                Me.txtAddressID_Leave(Nothing, Nothing)
                If Not Me.AddressPanel1.Address Is Nothing Then
                    Me.AddressPanel1.Address.SetAddressLink(myClient.Identity, 3)
                End If
            ElseIf CInt(strAddrLinkID) = 2 AndAlso myClient.ShippingAddressId > 0 Then
                Me.txtAddressID.Text = myClient.ShippingAddressId
                Me.txtAddressID_Leave(Nothing, Nothing)
                If Not Me.AddressPanel1.Address Is Nothing Then
                    Me.AddressPanel1.Address.SetAddressLink(myClient.Identity, 2)
                End If
            End If
        End If

    End Sub

    Private CurrentTag As Integer = -1

    Private Sub txtMailId_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMailID.DoubleClick, txtInvAddrID.DoubleClick, txtShipAddrID.DoubleClick, txtInvMailAddr.DoubleClick, txtInvShipAddr.DoubleClick
        'clear out info in case we find nothing 
        Me.AddressPanel1.Address = Nothing
        Me.txtAddressID.Text = ""
        CurrentTag = -1
        If TypeOf sender Is System.Windows.Forms.TextBox Then
            Dim txtaddrBox As Windows.Forms.TextBox = sender
            CurrentTag = txtaddrBox.Tag
            If txtaddrBox.Text.Trim.Length = 0 Then Exit Sub

            If CInt(txtaddrBox.Text) <= 0 Then Exit Sub
            Me.txtAddressID.Text = txtaddrBox.Text
            Me.txtAddressID.Tag = CurrentTag
            Me.txtAddressID_Leave(Nothing, Nothing)
            If CInt(txtaddrBox.Tag) > 0 AndAlso CInt(txtaddrBox.Tag) <= 3 Then

                Me.AddressPanel1.Address.SetAddressLink(Me.myClient.Identity, CInt(txtAddressID.Tag))
                'See if we have to put up a warning 
                If Not Me.myClient Is Nothing AndAlso CStr(txtaddrBox.Tag).Trim.Length > 0 Then
                    Dim strAddressMismatch As String = "lib_Customer.SharedMailAddress"
                    If CInt(txtaddrBox.Tag) = 3 Then strAddressMismatch = "lib_Customer.SharedInvMailAddress"
                    Dim strXML As String = CConnection.PackageStringList(strAddressMismatch, Client.Identity)
                    Dim strXSLSheet As String = "PairedAddresstxt1.xsl"
                    If CInt(txtaddrBox.Tag) = 3 Then strXSLSheet = "PairedInvAddresstxt1.xsl"

                    Dim strText As String = Medscreen.ResolveStyleSheet(strXML, strXSLSheet, CInt(txtaddrBox.Tag) - 1)
                    Me.txtSameAddress.Rtf = strText

                End If

            End If
            Try  ' Protecting
                Dim StrResp As String = CConnection.PackageStringList("LIB_ADDRESS.GETADDRESSLINKTYPES", CInt(txtAddressID.Text))
                If StrResp IsNot Nothing AndAlso StrResp.Trim.Length > 0 Then
                    Dim bits As String() = StrResp.Split(New Char() {"|"})
                    If bits.Length > 1 Then
                        Dim adrPhrase As MedscreenLib.Glossary.Phrase = AddrTypes.Item(bits(0), Glossary.PhraseCollection.IndexBy.OrderBY)
                        If adrPhrase IsNot Nothing Then cbADRLink.SelectedValue = adrPhrase.PhraseID
                    End If
                End If
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub txtMailID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMailID.Leave
        Dim intAddrID As Integer
        Try
            intAddrID = CInt(Me.txtMailID.Text)
        Catch
        End Try

        If intAddrID > 0 AndAlso Not myClient Is Nothing Then
            If myClient.MailAddressId <> intAddrID Then
                Dim dlgRet As MsgBoxResult = MsgBox("Do you want to change the result address Id from " & myClient.MailAddressId & " to " & intAddrID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = MsgBoxResult.Yes Then
                    myClient.MailAddressId = intAddrID
                    myClient.Update()
                End If
            End If

        End If
    End Sub

    Private Sub txtInvMailAddr_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvMailAddr.Leave
        If myClient Is Nothing Then Exit Sub
        Dim intAddrID As Integer
        Try
            intAddrID = CInt(Me.txtInvMailAddr.Text)
        Catch
        End Try

        If intAddrID > 0 AndAlso Not myClient Is Nothing Then
            If myClient.InvoiceMailAddressID <> intAddrID Then
                Dim dlgRet As MsgBoxResult = MsgBox("Do you want to change the Invoice Mail Address Id from " & myClient.MailAddressId & " to " & intAddrID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = MsgBoxResult.Yes Then
                    myClient.InvoiceMailAddressID = intAddrID
                    myClient.Update()
                End If
            End If

        End If
    End Sub

    Private Sub txtInvMailAddr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvMailAddr.TextChanged

    End Sub

    Private Sub AddressPanel1_FindAddress(ByVal AddressID As Integer) Handles AddressPanel1.FindAddress, AddressPanel1.NewAddress
        If myClient Is Nothing Then Exit Sub
        If txtAddressID.Tag = 1 Then
            myClient.MailAddressId = AddressID
            txtMailID.Text = CStr(AddressID)
        ElseIf txtAddressID.Tag = 3 Then
            myClient.InvoiceAddressId = AddressID
            txtInvAddrID.Text = CStr(AddressID)
        ElseIf txtAddressID.Tag = 2 Then
            myClient.ShippingAddressId = AddressID
            txtShipAddrID.Text = CStr(AddressID)
        ElseIf txtAddressID.Tag = 13 Then
            myClient.InvoiceMailAddressID = AddressID
            txtInvMailAddr.Text = CStr(AddressID)
        ElseIf txtAddressID.Tag = 14 Then
            myClient.InvoiceShippingAddressId = AddressID
            txtInvShipAddr.Text = CStr(AddressID)
        End If
        myClient.Update()
    End Sub

    Private Sub txtInvAddrID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvAddrID.Leave
        If myClient Is Nothing Then Exit Sub
        Dim intAddrID As Integer
        Try
            intAddrID = CInt(Me.txtInvAddrID.Text)
        Catch
        End Try

        If intAddrID > 0 AndAlso Not myClient Is Nothing Then
            If myClient.InvoiceAddressId <> intAddrID Then
                Dim dlgRet As MsgBoxResult = MsgBox("Do you want to change the Invoice address Id from " & myClient.MailAddressId & " to " & intAddrID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = MsgBoxResult.Yes Then
                    myClient.InvoiceAddressId = intAddrID
                    myClient.Update()
                End If
            End If

        End If
    End Sub

    Private Sub txtShipAddrID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipAddrID.Leave
        If myClient Is Nothing Then Exit Sub
        Dim intAddrID As Integer
        Try
            intAddrID = CInt(Me.txtShipAddrID.Text)
        Catch
        End Try

        If intAddrID > 0 AndAlso Not myClient Is Nothing Then
            If myClient.ShippingAddressId <> intAddrID Then
                Dim dlgRet As MsgBoxResult = MsgBox("Do you want to change the Secondary Results address Id from " & myClient.MailAddressId & " to " & intAddrID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = MsgBoxResult.Yes Then
                    myClient.ShippingAddressId = intAddrID
                    myClient.Update()
                End If
            End If

        End If
    End Sub

    Private Sub txtInvShipAddr_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvShipAddr.Leave
        If myClient Is Nothing Then Exit Sub
        Dim intAddrID As Integer
        Try
            intAddrID = CInt(Me.txtInvShipAddr.Text)
        Catch
        End Try

        If intAddrID > 0 AndAlso Not myClient Is Nothing Then
            If myClient.InvoiceShippingAddressId <> intAddrID Then
                Dim dlgRet As MsgBoxResult = MsgBox("Do you want to change the Invoice Shipping Address Id from " & myClient.MailAddressId & " to " & intAddrID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = MsgBoxResult.Yes Then
                    myClient.InvoiceShippingAddressId = intAddrID
                    myClient.Update()
                End If
            End If

        End If
    End Sub
End Class
