'''$Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2006-01-18 06:14:42+00 $ <para/>
'''$Log: FrmEdInvoice.vb,v $
'''Revision 1.0  2006-01-18 06:14:42+00  taylor
'''Altered locked display to allow editing of Purchase Orders request from Adebola
''' <para/>

Imports Intranet.intranet
Imports Intranet.intranet.support
Imports MedscreenLib
Imports MedscreenLib.Glossary.Glossary
Imports System.Xml
Imports System.Windows.Forms
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : FrmEdInvoice
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for editing invoices 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action>
''' $Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2006-01-18 06:14:42+00 $ <para/>
'''$Log: FrmEdInvoice.vb,v $
'''Revision 1.0  2006-01-18 06:14:42+00  taylor
'''Altered locked display to allow editing of Purchase Orders request from Adebola
'''
'''Revision 1.0  2005-11-29 18:16:07+00  taylor
'''Commented
''' <para/>
''' </Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class FrmEdInvoice
    Inherits System.Windows.Forms.Form

    Private DsLineItems As New DataSet()
    Private Client As Intranet.intranet.customerns.Client
    Private blnNoInvoice As Boolean = True
    Private strUserDestText As String = "the currently logged in user."
    Private strOutputType As String = "Email"
    Private prColl As IniFile.IniCollection
    Private myInvAddressID As Integer = 0
    Private myMailAddressID As Integer = 0
    Private myShipAddressID As Integer = 0
    Private blnChanging As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Me.AddressPanel1.Glossaries = ModTpanel.Glossaries
        Me.AddressPanel1.Tooltip = Me.ToolTip1
        Dim strXml As String = MedscreenLib.Glossary.Glossary.LineItems.ToXml
        Dim datadoc As New System.Xml.XmlDataDocument()

        Dim st As New IO.MemoryStream(Medscreen.StringToByteArray(strXml))

        Dim xr As XmlTextReader = New XmlTextReader(st)

        datadoc.DataSet.ReadXml(xr, XmlReadMode.InferSchema)

        DsLineItems.Merge(datadoc.DataSet)

        Me.DGItemCode = New cb.DataGridComboBoxColumnStyle(DsLineItems.Tables(0), "description", "itemcode")
        'DGItemCode
        '
        Me.DGItemCode.HeaderText = "Item"
        Me.DGItemCode.MappingName = "itemcode"
        Me.DGItemCode.Width = 200

        'DTSInvoiceItems
        '
        Me.DTSInvoiceItems.DataGrid = Me.DataGrid1
        Me.DTSInvoiceItems.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DGDeleted, Me.DGLineNumber, DataGridTextBoxColumn1, Me.DGQty, Me.DGPrice, Me.DGDiscount, Me.DGReference, DGItemCode, Me.DGComment})
        Me.DTSInvoiceItems.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DTSInvoiceItems.MappingName = "INVOICE_ITEMS"

        'Customer information 
        Me.cbCustomer.DataSource = cSupport.CustList
        Me.cbCustomer.ValueMember = "Identity"
        Me.cbCustomer.DisplayMember = "InfoAcc"

        'Currency info
        Me.cbCurrency.DataSource = MedscreenLib.Glossary.Glossary.Currencies
        Me.cbCurrency.ValueMember = "CurrencyId"
        Me.cbCurrency.DisplayMember = "Description"

        'Vat type info
        Me.cbVATPhrase.DataSource = MedscreenLib.Glossary.Glossary.InvoiceVATType
        Me.cbVATPhrase.ValueMember = "PhraseID"
        Me.cbVATPhrase.DisplayMember = "PhraseText"

        Me.CBInvDest.SelectedIndex = 2

        Me.cbType.SelectedIndex = 0

        Me.txtDestination.Text = CurrentSMUserEmail()
        cmdSave.Enabled = False

        Dim inF As IniFile.IniFiles = New IniFile.IniFiles()

        inF.FileName = "\\john\live\inifiles\reporterprinters.ini"
        Dim PColl As Collection = New Collection()
        PColl = inF.ReadSection("Printers")

        prColl = New IniFile.IniCollection()
        Dim pStr As String
        Dim i As Integer
        Dim iItem As IniFile.IniElement
        Dim inIndex As Integer

        For i = 1 To PColl.Count
            pStr = PColl.Item(i)
            inIndex = InStr(pStr, "=")
            If inIndex > 0 Then
                iItem = New IniFile.IniElement()
                iItem.Header = Mid(pStr, 1, inIndex - 1)
                iItem.Item = Mid(pStr, inIndex + 1)
                prColl.Add(iItem)
                Me.CBPrinters.Items.Add(iItem.Item)
            End If
        Next
        If prColl.Count > 0 Then
            Me.CBPrinters.SelectedIndex = 0
        End If

        'myInvoice = New Intranet.intranet.jobs.Invoice()
    End Sub

    Function CurrentSMUserEmail() As String
        Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item(Security.Principal.WindowsIdentity.GetCurrent.Name, _
        MedscreenLib.personnel.PersonnelList.FindIdBy.WindowsIdentity)
        If Not up Is Nothing Then
            Return up.Email
        Else
            Return ""
        End If
    End Function

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
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents ds As Intranet.Invoice
    Friend WithEvents DGLineNumber As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGItemCode As cb.DataGridComboBoxColumnStyle
    Friend WithEvents DGQty As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGPrice As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGDiscount As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGReference As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGDeleted As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DGComment As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtSentBy As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DTSentOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ckSample As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CKLocked As System.Windows.Forms.CheckBox
    Friend WithEvents txtVAT As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceValue As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtInvoicedBy As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpInvoiceOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DTSInvoiceItems As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RBInvAdd As System.Windows.Forms.RadioButton
    Friend WithEvents RBShipAddress As System.Windows.Forms.RadioButton
    Friend WithEvents rbMailAddress As System.Windows.Forms.RadioButton
    Friend WithEvents txtPurchOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtMailID As System.Windows.Forms.TextBox
    Friend WithEvents txtInvAddrID As System.Windows.Forms.TextBox
    Friend WithEvents txtShipAddrID As System.Windows.Forms.TextBox
    Friend WithEvents rbCustMailAddr As System.Windows.Forms.RadioButton
    Friend WithEvents RBCustInvAddr As System.Windows.Forms.RadioButton
    Friend WithEvents rbExtraAddr As System.Windows.Forms.RadioButton
    Friend WithEvents cmdEmail As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents txtSageID As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtExchRate As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbTransType As System.Windows.Forms.ComboBox
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents cbVATPhrase As System.Windows.Forms.ComboBox
    Friend WithEvents CBInvDest As System.Windows.Forms.ComboBox
    Friend WithEvents lblEmailed As System.Windows.Forms.Label
    Friend WithEvents txtDestination As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents RBEmail As System.Windows.Forms.RadioButton
    Friend WithEvents RBFax As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RBSelf As System.Windows.Forms.RadioButton
    Friend WithEvents RBMail As System.Windows.Forms.RadioButton
    Friend WithEvents RBInvoice As System.Windows.Forms.RadioButton
    Friend WithEvents RBShipping As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents cmdDeleteRow As System.Windows.Forms.Button
    Friend WithEvents cmdInsertLine As System.Windows.Forms.Button
    Friend WithEvents RBPrint As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents AddressPanel1 As MedscreenCommonGui.AddressPanel
    Friend WithEvents cmdEditAddress As System.Windows.Forms.Button
    Friend WithEvents cmdSearchAddress As System.Windows.Forms.Button
    Friend WithEvents cmdCopyAddress2 As System.Windows.Forms.Button
    Friend WithEvents lblAddressType As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents CBPrinters As System.Windows.Forms.ComboBox
    Friend WithEvents cmdPreview As System.Windows.Forms.Button
    Friend WithEvents cmdSend As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents lblCredited As System.Windows.Forms.Label
    Friend WithEvents cbType As System.Windows.Forms.ComboBox
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents txtInvNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DTLastSentOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtLastSentBy As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtVATNo As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmEdInvoice))
        Me.ds = New Intranet.Invoice()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.DGLineNumber = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGQty = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGPrice = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGDiscount = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGReference = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGDeleted = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DGComment = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cbType = New System.Windows.Forms.ComboBox()
        Me.CBPrinters = New System.Windows.Forms.ComboBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RBShipping = New System.Windows.Forms.RadioButton()
        Me.RBInvoice = New System.Windows.Forms.RadioButton()
        Me.RBMail = New System.Windows.Forms.RadioButton()
        Me.RBSelf = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.RBPrint = New System.Windows.Forms.RadioButton()
        Me.RBFax = New System.Windows.Forms.RadioButton()
        Me.RBEmail = New System.Windows.Forms.RadioButton()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtDestination = New System.Windows.Forms.TextBox()
        Me.lblEmailed = New System.Windows.Forms.Label()
        Me.CBInvDest = New System.Windows.Forms.ComboBox()
        Me.cmdEmail = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.DTSInvoiceItems = New System.Windows.Forms.DataGridTableStyle()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.cmdInsertLine = New System.Windows.Forms.Button()
        Me.cmdDeleteRow = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblCredited = New System.Windows.Forms.Label()
        Me.cmdSend = New System.Windows.Forms.Button()
        Me.cmdPreview = New System.Windows.Forms.Button()
        Me.cbVATPhrase = New System.Windows.Forms.ComboBox()
        Me.cbTransType = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cbCurrency = New System.Windows.Forms.ComboBox()
        Me.cbCustomer = New System.Windows.Forms.ComboBox()
        Me.txtSageID = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtExchRate = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPurchOrder = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtSentBy = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.DTSentOn = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ckSample = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CKLocked = New System.Windows.Forms.CheckBox()
        Me.txtVAT = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtInvoiceValue = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtInvoicedBy = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpInvoiceOn = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.AddressPanel1 = New MedscreenCommonGui.AddressPanel()
        Me.cmdEditAddress = New System.Windows.Forms.Button()
        Me.cmdSearchAddress = New System.Windows.Forms.Button()
        Me.cmdCopyAddress2 = New System.Windows.Forms.Button()
        Me.lblAddressType = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbExtraAddr = New System.Windows.Forms.RadioButton()
        Me.RBCustInvAddr = New System.Windows.Forms.RadioButton()
        Me.rbCustMailAddr = New System.Windows.Forms.RadioButton()
        Me.txtShipAddrID = New System.Windows.Forms.TextBox()
        Me.txtInvAddrID = New System.Windows.Forms.TextBox()
        Me.txtMailID = New System.Windows.Forms.TextBox()
        Me.RBInvAdd = New System.Windows.Forms.RadioButton()
        Me.RBShipAddress = New System.Windows.Forms.RadioButton()
        Me.rbMailAddress = New System.Windows.Forms.RadioButton()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtInvNumber = New System.Windows.Forms.TextBox()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DTLastSentOn = New System.Windows.Forms.DateTimePicker()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtLastSentBy = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtVATNo = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        CType(Me.ds, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.AddressPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'ds
        '
        Me.ds.DataSetName = "Invoice"
        Me.ds.Locale = New System.Globalization.CultureInfo("en-GB")
        Me.ds.Namespace = "http://tempuri.org/Invoice.xsd"
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(808, 3)
        Me.Splitter1.TabIndex = 1
        Me.Splitter1.TabStop = False
        '
        'DGLineNumber
        '
        Me.DGLineNumber.Format = ""
        Me.DGLineNumber.FormatInfo = Nothing
        Me.DGLineNumber.HeaderText = "Line Number"
        Me.DGLineNumber.MappingName = "linenumber"
        Me.DGLineNumber.Width = 75
        '
        'DGQty
        '
        Me.DGQty.Format = ""
        Me.DGQty.FormatInfo = Nothing
        Me.DGQty.HeaderText = "Quantity"
        Me.DGQty.MappingName = "quantity"
        Me.DGQty.Width = 75
        '
        'DGPrice
        '
        Me.DGPrice.Format = ""
        Me.DGPrice.FormatInfo = Nothing
        Me.DGPrice.HeaderText = "Price"
        Me.DGPrice.MappingName = "price"
        Me.DGPrice.Width = 75
        '
        'DGDiscount
        '
        Me.DGDiscount.Format = ""
        Me.DGDiscount.FormatInfo = Nothing
        Me.DGDiscount.HeaderText = "Discount"
        Me.DGDiscount.MappingName = "discount"
        Me.DGDiscount.Width = 75
        '
        'DGReference
        '
        Me.DGReference.Format = ""
        Me.DGReference.FormatInfo = Nothing
        Me.DGReference.HeaderText = "Reference"
        Me.DGReference.MappingName = "reference"
        Me.DGReference.Width = 75
        '
        'DGDeleted
        '
        Me.DGDeleted.Format = ""
        Me.DGDeleted.FormatInfo = Nothing
        Me.DGDeleted.HeaderText = "Deleted"
        Me.DGDeleted.MappingName = "Deleted"
        Me.DGDeleted.Width = 30
        '
        'DGComment
        '
        Me.DGComment.Format = ""
        Me.DGComment.FormatInfo = Nothing
        Me.DGComment.HeaderText = "Comments"
        Me.DGComment.MappingName = "comments"
        Me.DGComment.Width = 600
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Item Code"
        Me.DataGridTextBoxColumn1.MappingName = "itemcode1"
        Me.DataGridTextBoxColumn1.ReadOnly = True
        Me.DataGridTextBoxColumn1.Width = 75
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbType, Me.CBPrinters, Me.GroupBox4, Me.GroupBox3, Me.Label16, Me.txtDestination, Me.lblEmailed, Me.CBInvDest, Me.cmdEmail, Me.cmdCancel, Me.cmdOk})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 493)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(808, 80)
        Me.Panel2.TabIndex = 3
        '
        'cbType
        '
        Me.cbType.Items.AddRange(New Object() {"Invoice", "Credit Note"})
        Me.cbType.Location = New System.Drawing.Point(280, 8)
        Me.cbType.Name = "cbType"
        Me.cbType.Size = New System.Drawing.Size(104, 21)
        Me.cbType.TabIndex = 12
        Me.cbType.Visible = False
        '
        'CBPrinters
        '
        Me.CBPrinters.Location = New System.Drawing.Point(88, 48)
        Me.CBPrinters.Name = "CBPrinters"
        Me.CBPrinters.Size = New System.Drawing.Size(200, 21)
        Me.CBPrinters.TabIndex = 11
        Me.CBPrinters.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.RBShipping, Me.RBInvoice, Me.RBMail, Me.RBSelf})
        Me.GroupBox4.Location = New System.Drawing.Point(304, 40)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(384, 40)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "&Destination"
        '
        'RBShipping
        '
        Me.RBShipping.Location = New System.Drawing.Point(292, 16)
        Me.RBShipping.Name = "RBShipping"
        Me.RBShipping.Size = New System.Drawing.Size(72, 16)
        Me.RBShipping.TabIndex = 3
        Me.RBShipping.Text = "&Shipping"
        '
        'RBInvoice
        '
        Me.RBInvoice.Location = New System.Drawing.Point(200, 16)
        Me.RBInvoice.Name = "RBInvoice"
        Me.RBInvoice.Size = New System.Drawing.Size(72, 16)
        Me.RBInvoice.TabIndex = 2
        Me.RBInvoice.Text = "&Invoice"
        '
        'RBMail
        '
        Me.RBMail.Location = New System.Drawing.Point(108, 16)
        Me.RBMail.Name = "RBMail"
        Me.RBMail.Size = New System.Drawing.Size(72, 16)
        Me.RBMail.TabIndex = 1
        Me.RBMail.Text = "&Mail"
        '
        'RBSelf
        '
        Me.RBSelf.Checked = True
        Me.RBSelf.Location = New System.Drawing.Point(16, 16)
        Me.RBSelf.Name = "RBSelf"
        Me.RBSelf.Size = New System.Drawing.Size(72, 16)
        Me.RBSelf.TabIndex = 0
        Me.RBSelf.TabStop = True
        Me.RBSelf.Text = "&User"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.RBPrint, Me.RBFax, Me.RBEmail})
        Me.GroupBox3.Location = New System.Drawing.Point(400, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(216, 40)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "&Output"
        '
        'RBPrint
        '
        Me.RBPrint.Location = New System.Drawing.Point(146, 16)
        Me.RBPrint.Name = "RBPrint"
        Me.RBPrint.Size = New System.Drawing.Size(56, 16)
        Me.RBPrint.TabIndex = 2
        Me.RBPrint.Text = "&Print"
        '
        'RBFax
        '
        Me.RBFax.Location = New System.Drawing.Point(81, 16)
        Me.RBFax.Name = "RBFax"
        Me.RBFax.Size = New System.Drawing.Size(56, 16)
        Me.RBFax.TabIndex = 1
        Me.RBFax.Text = "&Fax"
        '
        'RBEmail
        '
        Me.RBEmail.Checked = True
        Me.RBEmail.Location = New System.Drawing.Point(16, 16)
        Me.RBEmail.Name = "RBEmail"
        Me.RBEmail.Size = New System.Drawing.Size(56, 16)
        Me.RBEmail.TabIndex = 0
        Me.RBEmail.TabStop = True
        Me.RBEmail.Text = "&Email"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(16, 48)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 23)
        Me.Label16.TabIndex = 8
        Me.Label16.Text = "Destination"
        '
        'txtDestination
        '
        Me.txtDestination.Location = New System.Drawing.Point(88, 48)
        Me.txtDestination.Name = "txtDestination"
        Me.txtDestination.Size = New System.Drawing.Size(200, 20)
        Me.txtDestination.TabIndex = 7
        Me.txtDestination.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtDestination, "Address to which the Invoice will be sent, please check for accuracy")
        '
        'lblEmailed
        '
        Me.lblEmailed.BackColor = System.Drawing.Color.Coral
        Me.lblEmailed.ForeColor = System.Drawing.Color.BlanchedAlmond
        Me.lblEmailed.Location = New System.Drawing.Point(288, 8)
        Me.lblEmailed.Name = "lblEmailed"
        Me.lblEmailed.TabIndex = 6
        Me.lblEmailed.Text = "Emailed"
        Me.lblEmailed.Visible = False
        '
        'CBInvDest
        '
        Me.CBInvDest.Items.AddRange(New Object() {"", "Original", "Copy", "Delivery note"})
        Me.CBInvDest.Location = New System.Drawing.Point(152, 8)
        Me.CBInvDest.Name = "CBInvDest"
        Me.CBInvDest.Size = New System.Drawing.Size(121, 21)
        Me.CBInvDest.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.CBInvDest, "Type of invoice to produce")
        '
        'cmdEmail
        '
        Me.cmdEmail.Image = CType(resources.GetObject("cmdEmail.Image"), System.Drawing.Bitmap)
        Me.cmdEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEmail.Location = New System.Drawing.Point(16, 8)
        Me.cmdEmail.Name = "cmdEmail"
        Me.cmdEmail.Size = New System.Drawing.Size(120, 23)
        Me.cmdEmail.TabIndex = 4
        Me.cmdEmail.Text = "&Email Invoice"
        Me.cmdEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEmail, "Email Invoice Copy  of the invoice to the currently logged in user.")
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(723, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(639, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2, Me.TabPage3})
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 32)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(808, 461)
        Me.TabControl1.TabIndex = 1
        Me.TabControl1.TabStop = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.DataGrid1, Me.Panel3, Me.Panel1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(800, 435)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Details"
        '
        'DataGrid1
        '
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.Lavender
        Me.DataGrid1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.LightGray
        Me.DataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.DataSource = Me.ds.INVOICE_ITEMS
        Me.DataGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGrid1.FlatMode = True
        Me.DataGrid1.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.DataGrid1.ForeColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.GridLineColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.DataGrid1.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.LinkColor = System.Drawing.Color.Teal
        Me.DataGrid1.Location = New System.Drawing.Point(88, 232)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsBackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.DataGrid1.SelectionForeColor = System.Drawing.Color.WhiteSmoke
        Me.DataGrid1.Size = New System.Drawing.Size(712, 203)
        Me.DataGrid1.TabIndex = 3
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DTSInvoiceItems})
        '
        'DTSInvoiceItems
        '
        Me.DTSInvoiceItems.DataGrid = Me.DataGrid1
        Me.DTSInvoiceItems.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DTSInvoiceItems.MappingName = ""
        '
        'Panel3
        '
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdInsertLine, Me.cmdDeleteRow})
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel3.Location = New System.Drawing.Point(0, 232)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(88, 203)
        Me.Panel3.TabIndex = 4
        '
        'cmdInsertLine
        '
        Me.cmdInsertLine.Location = New System.Drawing.Point(8, 72)
        Me.cmdInsertLine.Name = "cmdInsertLine"
        Me.cmdInsertLine.TabIndex = 1
        Me.cmdInsertLine.Text = "Insert Line"
        '
        'cmdDeleteRow
        '
        Me.cmdDeleteRow.Location = New System.Drawing.Point(8, 40)
        Me.cmdDeleteRow.Name = "cmdDeleteRow"
        Me.cmdDeleteRow.TabIndex = 0
        Me.cmdDeleteRow.Text = "Delete Line"
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtVATNo, Me.Label18, Me.txtLastSentBy, Me.Label17, Me.DTLastSentOn, Me.Label15, Me.lblCredited, Me.cmdSend, Me.cmdPreview, Me.cbVATPhrase, Me.cbTransType, Me.GroupBox2, Me.txtPurchOrder, Me.Label14, Me.Label13, Me.txtSentBy, Me.Label11, Me.DTSentOn, Me.Label12, Me.ckSample, Me.Label9, Me.CKLocked, Me.txtVAT, Me.Label7, Me.txtInvoiceValue, Me.Label6, Me.txtInvoicedBy, Me.Label4, Me.dtpInvoiceOn, Me.Label3})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(800, 232)
        Me.Panel1.TabIndex = 0
        '
        'lblCredited
        '
        Me.lblCredited.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lblCredited.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCredited.ForeColor = System.Drawing.Color.PaleGoldenrod
        Me.lblCredited.Location = New System.Drawing.Point(39, 64)
        Me.lblCredited.Name = "lblCredited"
        Me.lblCredited.Size = New System.Drawing.Size(153, 16)
        Me.lblCredited.TabIndex = 35
        Me.lblCredited.Text = "Invoice Has been credited "
        Me.lblCredited.Visible = False
        '
        'cmdSend
        '
        Me.cmdSend.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdSend.Enabled = False
        Me.cmdSend.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdSend.Image = CType(resources.GetObject("cmdSend.Image"), System.Drawing.Bitmap)
        Me.cmdSend.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSend.Location = New System.Drawing.Point(720, 146)
        Me.cmdSend.Name = "cmdSend"
        Me.cmdSend.TabIndex = 34
        Me.cmdSend.Text = "Send"
        Me.cmdSend.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSend, "Send and lock the Invoice")
        '
        'cmdPreview
        '
        Me.cmdPreview.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdPreview.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPreview.Image = CType(resources.GetObject("cmdPreview.Image"), System.Drawing.Bitmap)
        Me.cmdPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPreview.Location = New System.Drawing.Point(720, 112)
        Me.cmdPreview.Name = "cmdPreview"
        Me.cmdPreview.TabIndex = 33
        Me.cmdPreview.Text = "XML"
        Me.cmdPreview.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdPreview, "Preview XML to a separate window")
        '
        'cbVATPhrase
        '
        Me.cbVATPhrase.Location = New System.Drawing.Point(592, 112)
        Me.cbVATPhrase.Name = "cbVATPhrase"
        Me.cbVATPhrase.Size = New System.Drawing.Size(121, 21)
        Me.cbVATPhrase.TabIndex = 6
        '
        'cbTransType
        '
        Me.cbTransType.Items.AddRange(New Object() {"4 - Charge", "5 - Credit"})
        Me.cbTransType.Location = New System.Drawing.Point(592, 148)
        Me.cbTransType.Name = "cbTransType"
        Me.cbTransType.Size = New System.Drawing.Size(121, 21)
        Me.cbTransType.TabIndex = 9
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbCurrency, Me.cbCustomer, Me.txtSageID, Me.Label10, Me.txtExchRate, Me.Label8, Me.Label5, Me.Label2})
        Me.GroupBox2.Location = New System.Drawing.Point(216, -2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(568, 69)
        Me.GroupBox2.TabIndex = 30
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Customer specific details"
        Me.ToolTip1.SetToolTip(Me.GroupBox2, "Please ensure that if you change any of the customer specific details all others " & _
        "match.")
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(325, 19)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(96, 21)
        Me.cbCurrency.TabIndex = 39
        '
        'cbCustomer
        '
        Me.cbCustomer.Location = New System.Drawing.Point(83, 17)
        Me.cbCustomer.Name = "cbCustomer"
        Me.cbCustomer.Size = New System.Drawing.Size(152, 21)
        Me.cbCustomer.TabIndex = 38
        Me.ToolTip1.SetToolTip(Me.cbCustomer, "Customers no invoice selected")
        '
        'txtSageID
        '
        Me.txtSageID.Location = New System.Drawing.Point(86, 43)
        Me.txtSageID.Name = "txtSageID"
        Me.txtSageID.Size = New System.Drawing.Size(64, 20)
        Me.txtSageID.TabIndex = 37
        Me.txtSageID.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSageID, "Id of customer within the accounting system")
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(12, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 23)
        Me.Label10.TabIndex = 36
        Me.Label10.Text = "Sage ID:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtExchRate
        '
        Me.txtExchRate.Location = New System.Drawing.Point(491, 17)
        Me.txtExchRate.Name = "txtExchRate"
        Me.txtExchRate.Size = New System.Drawing.Size(64, 20)
        Me.txtExchRate.TabIndex = 35
        Me.txtExchRate.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtExchRate, "Exchange rate in force between Currency and Sterling.")
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(411, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 23)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Exch Rate:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(235, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 23)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "Invoice Currency:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(11, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 23)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Customer ID:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPurchOrder
        '
        Me.txtPurchOrder.Location = New System.Drawing.Point(117, 16)
        Me.txtPurchOrder.Name = "txtPurchOrder"
        Me.txtPurchOrder.Size = New System.Drawing.Size(64, 20)
        Me.txtPurchOrder.TabIndex = 1
        Me.txtPurchOrder.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtPurchOrder, "Purchase order supplied by customer")
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(13, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(96, 23)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Purch Order:"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(514, 146)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(70, 23)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "Trans Type:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSentBy
        '
        Me.txtSentBy.Location = New System.Drawing.Point(376, 144)
        Me.txtSentBy.Name = "txtSentBy"
        Me.txtSentBy.Size = New System.Drawing.Size(112, 20)
        Me.txtSentBy.TabIndex = 8
        Me.txtSentBy.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSentBy, "Person who sent the invoice ")
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(293, 146)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 23)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Sent By:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTSentOn
        '
        Me.DTSentOn.Location = New System.Drawing.Point(115, 146)
        Me.DTSentOn.Name = "DTSentOn"
        Me.DTSentOn.ShowCheckBox = True
        Me.DTSentOn.Size = New System.Drawing.Size(160, 20)
        Me.DTSentOn.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.DTSentOn, "Date on which the invoice was sent to the customer")
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 146)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Sent On :"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckSample
        '
        Me.ckSample.Location = New System.Drawing.Point(688, 83)
        Me.ckSample.Name = "ckSample"
        Me.ckSample.TabIndex = 19
        Me.ckSample.Text = "Sample Invoice"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(504, 112)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 23)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "VAT Type:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CKLocked
        '
        Me.CKLocked.Location = New System.Drawing.Point(552, 83)
        Me.CKLocked.Name = "CKLocked"
        Me.CKLocked.TabIndex = 16
        Me.CKLocked.Text = "Locked"
        Me.ToolTip1.SetToolTip(Me.CKLocked, "If you unlock an invoice and then edit it, you must be aware of the implications " & _
        "of your actions")
        '
        'txtVAT
        '
        Me.txtVAT.Enabled = False
        Me.txtVAT.Location = New System.Drawing.Point(376, 112)
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(112, 20)
        Me.txtVAT.TabIndex = 5
        Me.txtVAT.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtVAT, "VAT added to the invoice (if appropriate), this information will be updated when " & _
        "the invoice is printed")
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(293, 112)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 23)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Invoice VAT:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtInvoiceValue
        '
        Me.txtInvoiceValue.Location = New System.Drawing.Point(115, 112)
        Me.txtInvoiceValue.Name = "txtInvoiceValue"
        Me.txtInvoiceValue.Size = New System.Drawing.Size(112, 20)
        Me.txtInvoiceValue.TabIndex = 4
        Me.txtInvoiceValue.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtInvoiceValue, "Value of the invoice")
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 23)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Invoice Value:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtInvoicedBy
        '
        Me.txtInvoicedBy.Location = New System.Drawing.Point(376, 83)
        Me.txtInvoicedBy.Name = "txtInvoicedBy"
        Me.txtInvoicedBy.Size = New System.Drawing.Size(112, 20)
        Me.txtInvoicedBy.TabIndex = 3
        Me.txtInvoicedBy.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(293, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Invoiced By:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtpInvoiceOn
        '
        Me.dtpInvoiceOn.Location = New System.Drawing.Point(115, 83)
        Me.dtpInvoiceOn.Name = "dtpInvoiceOn"
        Me.dtpInvoiceOn.ShowCheckBox = True
        Me.dtpInvoiceOn.Size = New System.Drawing.Size(160, 20)
        Me.dtpInvoiceOn.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.dtpInvoiceOn, "Date on which the invoice was created")
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(6, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Invoiced On :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox5, Me.GroupBox1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(800, 435)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Addresses"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.AddressPanel1})
        Me.GroupBox5.Location = New System.Drawing.Point(16, 88)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(784, 360)
        Me.GroupBox5.TabIndex = 32
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Address Details"
        '
        'AddressPanel1
        '
        Me.AddressPanel1.Address = Nothing
        Me.AddressPanel1.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AddressPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AddressPanel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdEditAddress, Me.cmdSearchAddress, Me.cmdCopyAddress2, Me.lblAddressType, Me.Button2, Me.cmdSave})
        Me.AddressPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AddressPanel1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AddressPanel1.InfoVisible = False
        Me.AddressPanel1.isReadOnly = True
        Me.AddressPanel1.Location = New System.Drawing.Point(3, 16)
        Me.AddressPanel1.Name = "AddressPanel1"
        Me.AddressPanel1.ShowContact = True
        Me.AddressPanel1.ShowStatus = True
        Me.AddressPanel1.Size = New System.Drawing.Size(778, 341)
        Me.AddressPanel1.TabIndex = 33
        Me.AddressPanel1.Tooltip = Nothing
        '
        'cmdEditAddress
        '
        Me.cmdEditAddress.Image = CType(resources.GetObject("cmdEditAddress.Image"), System.Drawing.Bitmap)
        Me.cmdEditAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEditAddress.Location = New System.Drawing.Point(624, 184)
        Me.cmdEditAddress.Name = "cmdEditAddress"
        Me.cmdEditAddress.Size = New System.Drawing.Size(104, 23)
        Me.cmdEditAddress.TabIndex = 41
        Me.cmdEditAddress.Text = "&Edit"
        Me.cmdEditAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdSearchAddress
        '
        Me.cmdSearchAddress.Image = CType(resources.GetObject("cmdSearchAddress.Image"), System.Drawing.Bitmap)
        Me.cmdSearchAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSearchAddress.Location = New System.Drawing.Point(624, 134)
        Me.cmdSearchAddress.Name = "cmdSearchAddress"
        Me.cmdSearchAddress.Size = New System.Drawing.Size(104, 23)
        Me.cmdSearchAddress.TabIndex = 40
        Me.cmdSearchAddress.Text = "Find Address"
        Me.cmdSearchAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCopyAddress2
        '
        Me.cmdCopyAddress2.Image = CType(resources.GetObject("cmdCopyAddress2.Image"), System.Drawing.Bitmap)
        Me.cmdCopyAddress2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCopyAddress2.Location = New System.Drawing.Point(624, 100)
        Me.cmdCopyAddress2.Name = "cmdCopyAddress2"
        Me.cmdCopyAddress2.Size = New System.Drawing.Size(104, 23)
        Me.cmdCopyAddress2.TabIndex = 39
        Me.cmdCopyAddress2.Text = "Copy Address"
        Me.cmdCopyAddress2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAddressType
        '
        Me.lblAddressType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddressType.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lblAddressType.Location = New System.Drawing.Point(256, 7)
        Me.lblAddressType.Name = "lblAddressType"
        Me.lblAddressType.Size = New System.Drawing.Size(472, 23)
        Me.lblAddressType.TabIndex = 32
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Bitmap)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(624, 66)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(104, 23)
        Me.Button2.TabIndex = 31
        Me.Button2.Text = "New"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdSave
        '
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Bitmap)
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(624, 32)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(104, 23)
        Me.cmdSave.TabIndex = 30
        Me.cmdSave.Text = "Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbExtraAddr, Me.RBCustInvAddr, Me.rbCustMailAddr, Me.txtShipAddrID, Me.txtInvAddrID, Me.txtMailID, Me.RBInvAdd, Me.RBShipAddress, Me.rbMailAddress})
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(800, 72)
        Me.GroupBox1.TabIndex = 31
        Me.GroupBox1.TabStop = False
        '
        'rbExtraAddr
        '
        Me.rbExtraAddr.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.rbExtraAddr.Location = New System.Drawing.Point(504, 44)
        Me.rbExtraAddr.Name = "rbExtraAddr"
        Me.rbExtraAddr.Size = New System.Drawing.Size(128, 24)
        Me.rbExtraAddr.TabIndex = 8
        Me.rbExtraAddr.Text = "Client Extra Addr"
        '
        'RBCustInvAddr
        '
        Me.RBCustInvAddr.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.RBCustInvAddr.Location = New System.Drawing.Point(249, 44)
        Me.RBCustInvAddr.Name = "RBCustInvAddr"
        Me.RBCustInvAddr.TabIndex = 7
        Me.RBCustInvAddr.Text = "Client Inv Addr"
        '
        'rbCustMailAddr
        '
        Me.rbCustMailAddr.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.rbCustMailAddr.Location = New System.Drawing.Point(39, 44)
        Me.rbCustMailAddr.Name = "rbCustMailAddr"
        Me.rbCustMailAddr.TabIndex = 6
        Me.rbCustMailAddr.Text = "Client Mail Addr"
        '
        'txtShipAddrID
        '
        Me.txtShipAddrID.Location = New System.Drawing.Point(616, 13)
        Me.txtShipAddrID.Name = "txtShipAddrID"
        Me.txtShipAddrID.Size = New System.Drawing.Size(40, 20)
        Me.txtShipAddrID.TabIndex = 5
        Me.txtShipAddrID.Text = ""
        '
        'txtInvAddrID
        '
        Me.txtInvAddrID.Location = New System.Drawing.Point(352, 15)
        Me.txtInvAddrID.Name = "txtInvAddrID"
        Me.txtInvAddrID.Size = New System.Drawing.Size(40, 20)
        Me.txtInvAddrID.TabIndex = 4
        Me.txtInvAddrID.Text = ""
        '
        'txtMailID
        '
        Me.txtMailID.Location = New System.Drawing.Point(128, 13)
        Me.txtMailID.Name = "txtMailID"
        Me.txtMailID.Size = New System.Drawing.Size(40, 20)
        Me.txtMailID.TabIndex = 3
        Me.txtMailID.Text = ""
        '
        'RBInvAdd
        '
        Me.RBInvAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.RBInvAdd.Location = New System.Drawing.Point(248, 13)
        Me.RBInvAdd.Name = "RBInvAdd"
        Me.RBInvAdd.TabIndex = 1
        Me.RBInvAdd.Text = "Invoice Address"
        '
        'RBShipAddress
        '
        Me.RBShipAddress.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.RBShipAddress.Location = New System.Drawing.Point(504, 13)
        Me.RBShipAddress.Name = "RBShipAddress"
        Me.RBShipAddress.Size = New System.Drawing.Size(144, 24)
        Me.RBShipAddress.TabIndex = 2
        Me.RBShipAddress.Text = "Shipping Address"
        '
        'rbMailAddress
        '
        Me.rbMailAddress.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.rbMailAddress.Location = New System.Drawing.Point(40, 13)
        Me.rbMailAddress.Name = "rbMailAddress"
        Me.rbMailAddress.TabIndex = 0
        Me.rbMailAddress.Text = "Mail Address"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel4})
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(800, 435)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "History"
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.AddRange(New System.Windows.Forms.Control() {Me.HtmlEditor1})
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.DockPadding.All = 5
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(800, 435)
        Me.Panel4.TabIndex = 0
        '
        'HtmlEditor1
        '
        'Me.HtmlEditor1.DefaultComposeSettings.BackColor = System.Drawing.Color.White
        'Me.HtmlEditor1.DefaultComposeSettings.DefaultFont = New System.Drawing.Font("Arial", 10.0!)
        'Me.HtmlEditor1.DefaultComposeSettings.Enabled = False
        'Me.HtmlEditor1.DefaultComposeSettings.ForeColor = System.Drawing.Color.Black
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        'Me.HtmlEditor1.DocumentEncoding = onlyconnect.EncodingType.WindowsCurrent
        Me.HtmlEditor1.Location = New System.Drawing.Point(5, 5)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        'Me.HtmlEditor1.OpenLinksInNewWindow = True
        'Me.HtmlEditor1.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left
        'Me.HtmlEditor1.SelectionBackColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionBullets = False
        'Me.HtmlEditor1.SelectionFont = Nothing
        'Me.HtmlEditor1.SelectionForeColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionNumbering = False
        Me.HtmlEditor1.Size = New System.Drawing.Size(786, 421)
        Me.HtmlEditor1.TabIndex = 0
        'Me.HtmlEditor1.Text = "HtmlEditor1"
        '
        'txtInvNumber
        '
        Me.txtInvNumber.BackColor = System.Drawing.SystemColors.Info
        Me.txtInvNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtInvNumber.Location = New System.Drawing.Point(109, 3)
        Me.txtInvNumber.Name = "txtInvNumber"
        Me.txtInvNumber.Size = New System.Drawing.Size(75, 20)
        Me.txtInvNumber.TabIndex = 0
        Me.txtInvNumber.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtInvNumber, "Move off control to accept invoice number ")
        '
        'Panel5
        '
        Me.Panel5.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtInvNumber, Me.Label1})
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 3)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(808, 29)
        Me.Panel5.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(7, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Invoice Number :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTLastSentOn
        '
        Me.DTLastSentOn.Enabled = False
        Me.DTLastSentOn.Location = New System.Drawing.Point(115, 176)
        Me.DTLastSentOn.Name = "DTLastSentOn"
        Me.DTLastSentOn.ShowCheckBox = True
        Me.DTLastSentOn.Size = New System.Drawing.Size(160, 20)
        Me.DTLastSentOn.TabIndex = 36
        Me.ToolTip1.SetToolTip(Me.DTLastSentOn, "Date on which the invoice was sent to the customer")
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 176)
        Me.Label15.Name = "Label15"
        Me.Label15.TabIndex = 37
        Me.Label15.Text = "Last Sent On :"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLastSentBy
        '
        Me.txtLastSentBy.Enabled = False
        Me.txtLastSentBy.Location = New System.Drawing.Point(376, 176)
        Me.txtLastSentBy.Name = "txtLastSentBy"
        Me.txtLastSentBy.Size = New System.Drawing.Size(112, 20)
        Me.txtLastSentBy.TabIndex = 38
        Me.txtLastSentBy.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtLastSentBy, "Person who sent the invoice ")
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(293, 176)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(72, 23)
        Me.Label17.TabIndex = 39
        Me.Label17.Text = "Last Sent By:"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVATNo
        '
        Me.txtVATNo.Enabled = True
        Me.txtVATNo.Location = New System.Drawing.Point(592, 176)
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New System.Drawing.Size(112, 20)
        Me.txtVATNo.TabIndex = 40
        Me.txtVATNo.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtVATNo, "Person who sent the invoice ")
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(512, 176)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 23)
        Me.Label18.TabIndex = 41
        Me.Label18.Text = "VAT NO"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FrmEdInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(808, 573)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.Panel5, Me.Panel2, Me.Splitter1})
        Me.Name = "FrmEdInvoice"
        Me.ShowInTaskbar = False
        Me.Text = "Edit Invoices"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ds, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.AddressPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private myInvoice As Intranet.intranet.jobs.Invoice = Nothing

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the invoice number and leave the form filling it
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtInvNumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvNumber.Leave
        Dim strInv As String = txtInvNumber.Text
        Dim blnChanges As Boolean = False

        strInv = strInv.ToUpper                 ' convert to upper case
        blnNoInvoice = False

        If strInv.Trim.Length > 0 Then          'If we have an invoice value 
            If Not myInvoice Is Nothing Then
                If myInvoice.InvoiceID = strInv Then Exit Sub
                blnChanges = GetBindings()
            End If
            CheckForChanges(blnChanges)
            myInvoice = cSupport.Invoices.Item(strInv)
            If Not myInvoice Is Nothing Then
                Try
                    Dim strXml As String = myInvoice.ToXML(Intranet.intranet.jobs.Invoice.InvoiceFormat.Normal)
                    Dim datadoc As New System.Xml.XmlDataDocument()

                    Dim st As New IO.MemoryStream(Medscreen.StringToByteArray(strXml))

                    Dim xr As XmlTextReader = New XmlTextReader(st)
                    DoBindings()
                    ds.EnforceConstraints = False

                    ds.INVOICE_ITEMS.Clear()
                    ds.Invoice.Clear()
                    datadoc.DataSet.ReadXml(xr, XmlReadMode.InferSchema)
                    ds.Merge(datadoc.DataSet, False, MissingSchemaAction.Add)
                    ds.AcceptChanges()
                    Me.DataGrid1.DataSource = ds.INVOICE_ITEMS
                    Me.BindingContext(ds, "Invoice").Position = ds.Invoice.Rows.Count - 1
                    Me.BindingContext(ds.Invoice).Position = ds.Invoice.Rows.Count - 1
                    Me.BindingContext(ds).Position = 0
                    'Me.txtCustID.Refresh()
                    Me.dtpInvoiceOn.Refresh()

                Catch ex As Exception
                    Medscreen.LogError(ex)
                End Try

            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if invoice has been edited
    ''' </summary>
    ''' <param name="blnChanges"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub CheckForChanges(ByVal blnChanges As Boolean)
        If ds.HasChanges Or blnChanges Then
            Dim mRet As MsgBoxResult = MsgBox("The invoice has changed do you want to save the changes?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If mRet = MsgBoxResult.Yes Then
                Dim i As Integer
                Dim iir As Intranet.Invoice.INVOICE_ITEMSRow
                For Each iir In ds.INVOICE_ITEMS.Rows
                    If iir.RowState = DataRowState.Modified Or iir.RowState = DataRowState.Added _
                    Or iir.RowState = DataRowState.Deleted Then
                        If iir.RowState = DataRowState.Added Then
                            Dim objII As Intranet.intranet.jobs.InvoiceItem = myInvoice.InvoiceItems.Create(iir.linenumber)
                            If Not objII Is Nothing Then
                                If Not iir.IscommentsNull Then objII.Comment = iir.comments
                                If Not iir.IsitemcodeNull Then objII.ItemCode = iir.itemcode
                                If Not iir.IspriceNull Then objII.Price = iir.price
                                If Not iir.IsdiscountNull Then objII.Discount = iir.discount
                                If Not iir.IsreferenceNull Then objII.Reference = iir.reference
                                If Not iir.IsquantityNull Then objii.quantity = iir.quantity
                                objII.Update()
                            End If
                        ElseIf iir.RowState = DataRowState.Modified Then
                            Dim objII As Intranet.intranet.jobs.InvoiceItem = myInvoice.InvoiceItems.Item(iir.linenumber, 0)
                            If Not objii Is Nothing Then
                                If Not iir.IscommentsNull Then objII.Comment = iir.comments
                                If Not iir.IsitemcodeNull Then objII.ItemCode = iir.itemcode
                                If Not iir.IspriceNull Then objII.Price = iir.price
                                If Not iir.IsdiscountNull Then objII.Discount = iir.discount
                                If Not iir.IsreferenceNull Then objII.Reference = iir.reference
                                If Not iir.IsquantityNull Then objii.quantity = iir.quantity
                                If iir.Deleted = "*" Then
                                    If objii.Delete() Then
                                        myInvoice.InvoiceItems.RemoveAt(myInvoice.InvoiceItems.IndexOf(objii))
                                    End If
                                Else
                                    objII.Update()
                                End If
                            End If
                        ElseIf iir.RowState = DataRowState.Deleted Then
                        End If
                    End If
                Next
                'If blnChanges Then
                myInvoice.ModifiedOn = Now

                myInvoice.Update(True)
                'End 'If
        Else
            myInvoice.DataFields.Rollback()
            Me.ds.INVOICE_ITEMS.RejectChanges()
        End If
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change the state of teh controls from enabled to disabled and vice versa
    ''' </summary>
    ''' <param name="Value">State  to take</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub EnableControls(ByVal Value As Boolean)
        Me.txtInvoiceValue.Enabled = Value
        Me.txtExchRate.Enabled = Value
        Me.txtInvoicedBy.Enabled = Value
        'Me.txtPurchOrder.Enabled = Value
        Me.txtSageID.Enabled = Value
        Me.txtSentBy.Enabled = Value
        Me.cbCurrency.Enabled = Value
        Me.cbCustomer.Enabled = Value
        Me.cbTransType.Enabled = Value
        Me.cbVATPhrase.Enabled = Value
        Me.dtpInvoiceOn.Enabled = Value
        Me.DTSentOn.Enabled = Value
        Me.DataGrid1.Enabled = Value
        Me.cmdDeleteRow.Enabled = Value
        Me.cmdInsertLine.Enabled = Value
        Me.cmdSend.Enabled = Value
        'Me.txtVATNo.Enabled = Value
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Bind the invoice object to the controls 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DoBindings()
        Me.lblEmailed.Hide()
        With myInvoice
            Try
                If .DateSent <> DateField.ZeroDate Then
                    Me.DTSentOn.Value = .DateSent
                    Me.DTSentOn.Checked = True
                Else
                    Me.DTSentOn.Value = Now
                    Me.DTSentOn.Checked = False
                End If

                If .DateLastSent <> DateField.ZeroDate Then
                    Me.DTLastSentOn.Value = .DateLastSent
                    Me.DTLastSentOn.Checked = True
                Else
                    Me.DTLastSentOn.Value = Now
                    Me.DTLastSentOn.Checked = False
                End If


                Me.txtSentBy.Text = .SentBy.Trim
                Me.txtLastSentBy.Text = .LastSentBy.Trim
                Me.txtSageID.Text = .SageId.Trim
                Me.txtVATNo.Text = .VATNo.Trim
                Me.txtPurchOrder.Text = .PurchaseOrder.Trim
                Me.ckSample.Checked = .SampleInvoice
                'Me.txtVatType.Text = .VatType.Trim
                Me.cbVATPhrase.SelectedIndex = MedscreenLib.Glossary.Glossary.InvoiceVATType.IndexOf(.VatType.Trim)
                Me.CKLocked.Checked = .Locked
                Me.EnableControls(Not .Locked)
                Me.txtExchRate.Text = .ExchangeRate.Trim
                Me.txtVAT.Text = .InvoiceVat
                Me.txtInvoiceValue.Text = .InvoiceAmount
                Me.txtInvoicedBy.Text = .InvoicedBy.Trim
                Me.dtpInvoiceOn.Value = .InvoiceDate
                Client = Nothing
                Client = cSupport.CustList.Item(.ClientID.Trim)
                blnChanging = True
                Me.cbCustomer.DataSource = Nothing

                Me.cbCustomer.DataSource = cSupport.CustList
                Me.cbCustomer.ValueMember = "Identity"
                Me.cbCustomer.DisplayMember = "InfoAcc"
                Threading.Thread.CurrentThread.Sleep(1000)
                blnChanging = False
                'Me.cbCustomer.Refresh()
                Dim intIndex As Integer = cSupport.CustList.IndexOf(Client)


                Me.cbCustomer.SelectedIndex = intIndex
                Me.cbCustomer.SelectedValue = .ClientID.Trim
                Me.cbCustomer_SelectedValueChanged(Nothing, Nothing)

                'Me.txtCustID.Text = .ClientID.Trim
                If .TransactionType = 5 Then
                    Me.cbTransType.SelectedIndex = 1
                Else
                    Me.cbTransType.SelectedIndex = 0
                End If
                'Me.txtInvCurrency.Text = .InvoiceCurrency.Trim
                Dim i As Integer '
                'See if we have 1 or 3 character invoice currency codes
                If .InvoiceCurrency.Trim.Length = 1 Then
                    For i = 0 To MedscreenLib.Glossary.Glossary.Currencies.Count - 1
                        If MedscreenLib.Glossary.Glossary.Currencies.Item(i).SingleCharId = .InvoiceCurrency.Trim Then
                            Me.cbCurrency.SelectedIndex = i
                            Exit For
                        End If
                    Next
                Else
                    Me.cbCurrency.SelectedValue = .InvoiceCurrency.Trim
                End If
                Me.txtMailID.Text = .MailaddressID
                Me.myMailAddressID = .MailaddressID
                Me.txtInvAddrID.Text = .InvoiceAddressID
                Me.myInvAddressID = .InvoiceAddressID
                Me.txtShipAddrID.Text = .ShippingAddressID
                Me.myShipAddressID = .ShippingAddressID

                Dim myaddress As MedscreenLib.Address.Caddress

                myaddress = Intranet.intranet.Support.cSupport.SMAddresses.item(.MailaddressID)

                Me.AddressPanel1.Address = myaddress
                Me.rbMailAddress.Checked = True
                Me.lblAddressType.Text = "Mail Address"

                'check for credited invoices

                Me.lblCredited.Visible = .Credited
                Me.cbType.Visible = .Credited

            Catch ex As Exception
            End Try
        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the contents of the control
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetBindings() As Boolean
        Dim myblnRet As Boolean = True
        With myInvoice
            Try
                .DataFields.Commit()
                .Locked = (Me.CKLocked.Checked)
                .PurchaseOrder = Me.txtPurchOrder.Text
                .VATNo = Me.txtVATNo.Text.Trim
                If Not .Locked Then
                    If Me.DTSentOn.Checked Then .DateSent = Me.DTSentOn.Value
                    .SentBy = Me.txtSentBy.Text
                    .SageId = Me.txtSageID.Text
                    .SampleInvoice = Me.ckSample.Checked
                    If Me.cbVATPhrase.SelectedIndex > -1 Then
                        .VatType = MedscreenLib.Glossary.Glossary.InvoiceVATType.Item(Me.cbVATPhrase.SelectedIndex).PhraseID
                    End If
                    .Locked = (Me.CKLocked.Checked)
                    .ExchangeRate = Me.txtExchRate.Text
                    If Me.txtVAT.Text.Trim.Length > 0 Then .InvoiceVat = Me.txtVAT.Text
                    'Check to see if we have a valid customer 
                    If (Me.txtInvoiceValue.Text.Trim.Length > 0) AndAlso Me.cbCustomer.SelectedIndex > -1 Then
                        .InvoiceAmount = Me.txtInvoiceValue.Text
                    End If
                    .InvoicedBy = Me.txtInvoicedBy.Text
                    .InvoiceDate = Me.dtpInvoiceOn.Value
                    If Not Me.cbCustomer.SelectedValue = Nothing Then .ClientID = Me.cbCustomer.SelectedValue
                    'check Client id if we have no client raise an exception
                    If .ClientID.Trim.Length = 0 Then
                        Throw New Exception("There is no client associated with this Invoice")
                    End If
                    If Me.cbTransType.SelectedIndex = 0 Then
                        .TransactionType = 4
                    Else
                        .TransactionType = 5
                    End If
                    If Me.cbCurrency.SelectedIndex > -1 Then
                        .InvoiceCurrency = cbCurrency.SelectedValue
                    End If
                End If
                If ((.InvoiceAddressID + .ShippingAddressID + .MailaddressID) <> 0) And _
                ((.InvoiceAddressID = 0) Or (.MailaddressID = 0) Or (.ShippingAddressID = 0)) Then
                    'We have an address not set
                    MsgBox("Either all addresses must be set or none")
                    myblnRet = False
                End If
                .MailaddressID = Me.myMailAddressID
                If .MailaddressID = 0 And .DataFields("MAILADDRESSID").OldValue <> 0 Then .DataFields("MAILADDRESSID").SetNull()
                .InvoiceAddressID = Me.myInvAddressID
                If .InvoiceAddressID = 0 And .DataFields("INVADDRESSID").OldValue <> 0 Then .DataFields("INVADDRESSID").SetNull()
                .ShippingAddressID = Me.myShipAddressID
                If .ShippingAddressID = 0 And .DataFields("SHIPADDRESSID").OldValue <> 0 Then .DataFields("SHIPADDRESSID").SetNull()

            Catch ex As Exception
      End Try

      Dim blnRet As Boolean = .DataFields.Changed
      If blnRet Then 'Update mod info 
        .ModifiedOn = Now
        Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item( _
            System.Security.Principal.WindowsIdentity.GetCurrent.Name, _
            MedscreenLib.personnel.PersonnelList.FindIdBy.WindowsIdentity)

        If Not up Is Nothing Then
          .ModifiedBy = up.Identity
        End If
      End If
      Return blnRet
    End With
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accept the form and exit 
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

        Dim blnChanges As Boolean = True

        If Not myInvoice Is Nothing Then
            blnChanges = GetBindings()
        End If
        Me.CheckForChanges(blnChanges)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RBMail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMailAddress.CheckedChanged
        Dim myaddress As MedscreenLib.Address.Caddress

        If Me.rbMailAddress.Checked Then
            'Me.cmdSave.Enabled = True
            Me.Button2.Enabled = True
            myaddress = Intranet.intranet.Support.cSupport.SMAddresses.item(myMailAddressID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            SetAddressStatus(myaddress)
            Me.lblAddressType.Text = "Mail Address"

            Me.AddressPanel1.Address = myaddress
            'Me.txtContact2.Text = myaddress.Contact
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RBInvAdd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBInvAdd.CheckedChanged
        Dim myaddress As MedscreenLib.Address.Caddress
        If RBInvAdd.Checked Then
            'Me.cmdsave.Enabled = True
            Me.Button2.Enabled = True
            myaddress = Intranet.intranet.Support.cSupport.SMAddresses.item(myInvAddressID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            SetAddressStatus(myaddress)

            Me.lblAddressType.Text = "Invoice Address"
            Me.AddressPanel1.Address = myaddress
            'Me.txtContact2.Text = myaddress.Contact
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up address status
    ''' </summary>
    ''' <param name="myAddress"></param>
    ''' <param name="DisableAll"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetAddressStatus(ByVal myAddress As MedscreenLib.Address.Caddress, _
    Optional ByVal DisableAll As Boolean = False)
        If Not myAddress Is Nothing Then
            Me.cmdCopyAddress2.Enabled = True
            Me.cmdSearchAddress.Enabled = True
            'Me.cmdsave.Enabled = True
            If myAddress.Usage = 0 Or myAddress.Usage = 1 Then
                Me.cmdEditAddress.Enabled = True
                Me.AddressPanel1.isReadOnly = True
                'Me.txtContact2.Enabled = False
            Else
                Me.cmdEditAddress.Enabled = False
                Me.AddressPanel1.isReadOnly = True
                'Me.txtContact2.Enabled = False
            End If
        End If
        If DisableAll Then
            'Me.cmdsave.Enabled = False
            Me.Button2.Enabled = False
            Me.cmdCopyAddress2.Enabled = False
            Me.cmdSearchAddress.Enabled = False
            Me.cmdEditAddress.Enabled = False
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RadioButtonEx1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBShipAddress.CheckedChanged
        Dim myaddress As MedscreenLib.Address.Caddress
        If Me.RBShipAddress.Checked Then
            'Me.cmdsave.Enabled = True
            Me.Button2.Enabled = True
            myaddress = Intranet.intranet.Support.cSupport.SMAddresses.item(myShipAddressID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            SetAddressStatus(myaddress)
            Me.lblAddressType.Text = "Shipping Address"
            Me.AddressPanel1.Address = myaddress
            'Me.txtContact2.Text = myaddress.Contact
        End If

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
        Dim oAddress As MedscreenLib.Address.Caddress = Me.AddressPanel1.Address
        If Not oAddress Is Nothing Then
            'oAddress.Contact = Me.txtContact2.Text
            oAddress.Update()

        End If
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an addresss
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
        'MsgBox("This code is not yet functional")

        Dim oAddress As MedscreenLib.Address.Caddress

        oAddress = cSupport.SMAddresses.CreateAddress()
        If Not oAddress Is Nothing Then
            oAddress.Update()                                       'Update the address to ensure that it is saved okay
            If Me.RBInvAdd.Checked Then
                myInvoice.InvoiceAddressID = oAddress.AddressId
            ElseIf Me.RBShipAddress.Checked Then
                myInvoice.ShippingAddressID = oAddress.AddressId
            ElseIf Me.rbMailAddress.Checked Then
                myInvoice.MailaddressID = oAddress.AddressId
            End If
        End If
        Me.AddressPanel1.Address = oAddress
        Me.AddressPanel1.isReadOnly = False
        'Me.txtContact2.Enabled = True
        Me.cmdSave.Enabled = True
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub rbCustMailAddr_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbCustMailAddr.CheckedChanged
        Dim myaddress As MedscreenLib.Address.Caddress

        If Me.rbCustMailAddr.Checked Then
            If Not Client Is Nothing Then

                Me.lblAddressType.Text = "Results Address"
                Dim iAddress As MedscreenLib.Address.Caddress
                If Not myaddress Is Nothing Then
                    iAddress = New MedscreenLib.Address.Caddress()
                    With iAddress
                        .AddressId = 0
                        .AddressType = "CUST"
                        .AdrLine1 = myaddress.AdrLine1
                        .AdrLine2 = myaddress.AdrLine2
                        .AdrLine3 = myaddress.AdrLine3
                        .City = myaddress.City
                        .District = myaddress.District
                        .PostCode = myaddress.PostCode
                        .Country = 0
                    End With
                    Me.AddressPanel1.Address = iAddress
                    'Me.txtContact2.Text = iAddress.Contact
                End If
                Me.SetAddressStatus(iAddress, True)
            End If
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    'Private Sub rbCustInvAddr_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBCustInvAddr.CheckedChanged
    '    Dim mySmClient As Intranet.intranet.customerns.Client

    '    If Me.RBCustInvAddr.Checked Then
    '        If Not Client Is Nothing Then
    '            'Me.cmdsave.Enabled = False
    '            Me.Button2.Enabled = False
    '            mySmClient = Client
    '            Me.lblAddressType.Text = "Customer Invoice Address"
    '            Dim iAddress As MedscreenLib.Address.Caddress
    '            If Not mySmClient Is Nothing Then
    '                'iAddress = mySmClient.InvoiceInfo.InvoiceAddress
    '                'With iAddress
    '                '    .AddressId = 0
    '                '    .AddressType = "CUST"
    '                '    .AdrLine1 = mySmClient.InvoiceInfo.InvoiceAddress
    '                '    .AdrLine2 = mySmClient.InvoiceInfo.InvAddress2
    '                '    .AdrLine3 = mySmClient.InvoiceInfo.InvAddress3
    '                '    .City = mySmClient.InvoiceInfo.InvAddress4
    '                '    .District = mySmClient.InvoiceInfo.InvAddress5
    '                '    .PostCode = mySmClient.InvoiceInfo.InvAddress6
    '                '    .Country = 0
    '                'End With
    '                Me.AddressPanel1.Address = iAddress
    '                'Me.txtContact2.Text = iAddress.Contact
    '            End If
    '            Me.SetAddressStatus(iAddress, True)
    '        End If
    '    End If


    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with address radio button change 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ExtraAddr_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbExtraAddr.CheckedChanged
        Dim mySmClient As Intranet.intranet.customerns.Client

        If Me.rbExtraAddr.Checked Then
            'Me.cmdsave.Enabled = False
            Me.Button2.Enabled = False
            If Not Client Is Nothing Then
                mySmClient = Client
                If Not mySmClient Is Nothing Then
                    Me.lblAddressType.Text = "Customer Secondary Address"
                    If Not mySmClient Is Nothing Then
                        'Me.AddressPanel1.Address = mySmClient.SecondaryAddress
                    End If
                End If
                'Me.SetAddressStatus(mySmClient.SecondaryAddress, True)
            End If
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find an address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSearchAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearchAddress.Click

        Dim Faddr As New MedscreenCommonGui.FindAddress()

        If Faddr.ShowDialog = DialogResult.OK Then
            Dim intId As Integer = Faddr.AddressID
            Dim oAddr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(intId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            If Me.RBInvAdd.Checked Then
                Me.myInvoice.InvoiceAddressID = oAddr.AddressId
                Me.txtInvAddrID.Text = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Vessel_Address_Inv)
            ElseIf Me.rbMailAddress.Checked Then
                Me.myInvoice.MailaddressID = oAddr.AddressId
                Me.txtMailID.Text = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Vessel_Address_Mail)
            ElseIf Me.RBShipAddress.Checked Then
                Me.myInvoice.ShippingAddressID = oAddr.AddressId
                Me.txtShipAddrID.Text = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Delivery_Address)
            End If
            Me.myInvoice.Update()
            Me.AddressPanel1.Address = oAddr
            Me.AddressPanel1.isReadOnly = True
            'Me.txtContact2.Enabled = False

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Copy address to new address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdCopyAddress2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyAddress2.Click
        Dim Faddr As New MedscreenCommonGui.FindAddress()

        If Faddr.ShowDialog = DialogResult.OK Then
            Dim intId As Integer = Faddr.AddressID
            Dim oAddr As MedscreenLib.Address.Caddress = MedscreenLib.Glossary.Glossary.SMAddresses.Copy(intId)
            If Me.RBInvAdd.Checked Then
                Me.myInvoice.InvoiceAddressID = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Vessel_Address_Inv)
            ElseIf Me.rbMailAddress.Checked Then
                Me.myInvoice.MailaddressID = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Vessel_Address_Mail)
            ElseIf Me.RBShipAddress.Checked Then
                Me.myInvoice.ShippingAddressID = oAddr.AddressId
                Me.AddressPanel1.Address = oAddr
                cSupport.SMAddresses.AddLink(oAddr.AddressId, "Customer.Identity", myInvoice.ClientID, MedscreenLib.Constants.GCST_Delivery_Address)
            End If
            Me.myInvoice.Update()
            Me.AddressPanel1.Address = oAddr
            Me.AddressPanel1.isReadOnly = False
            'Me.txtContact2.Enabled = True

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Senbd invoice as Email
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmail.Click
        Me.cmdOk_Click(Nothing, Nothing)
        'Givie it a chance to do something
        System.Threading.Thread.CurrentThread.Sleep(1000)

        Dim ACCInt As AccountsInterface = New AccountsInterface()
        With ACCInt
            .Status = MedscreenLib.Constants.GCST_IFace_StatusRequest
            .Reference = Me.myInvoice.InvoiceID
            .UserId = CurrentSMUser.Identity
            .SendAddress = Me.txtDestination.Text
            .SetSendDateNull()
            If Me.CBInvDest.SelectedIndex > -1 Then
                .CopyType = Me.CBInvDest.SelectedIndex
            Else
                .CopyType = 2
            End If
            .InvoiceType = myInvoice.TransactionType
            If myInvoice.Credited Then
                .InvoiceType = Me.cbType.SelectedIndex + 4
            End If

            .RequestCode = "SEND"

            .Update()
            Dim strError As String = ""
            Dim ChRet As Char
            ChRet = .Wait(strError, 100)
            While ChRet <> Constants.GCST_IFace_StatusCreated And strError.Trim.Length = 0
                System.Threading.Thread.CurrentThread.Sleep(100)
                ChRet = .Wait(strError, 1)
            End While


            'Me.lblEmailed.Show()
            If strError.Trim.Length = 0 Then
                MsgBox("Invoice Emailed/Printed", MsgBoxStyle.OKOnly)
            Else
                MsgBox(strError)
            End If
            ACCInt.Status = Constants.GCST_IFace_StatusDelete
            ACCInt.Delete()
            Me.txtInvAddrID.Focus()
        End With
    End Sub

    Private Sub txtMailID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMailID.Leave
        If txtMailID.Text.Trim.Length <> 0 Then
            Me.myMailAddressID = txtMailID.Text
        Else
            Me.myMailAddressID = 0
        End If
        Me.rbMailAddress.Checked = True
        RBMail_CheckedChanged(Nothing, Nothing)
    End Sub

    Private Sub txtInvAddrID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvAddrID.Leave
        If txtInvAddrID.Text.Trim.Length <> 0 Then
            Me.myInvAddressID = txtInvAddrID.Text
        Else
            Me.myInvAddressID = 0
        End If
        Me.RBInvAdd.Checked = True
        RBInvAdd_CheckedChanged(Nothing, Nothing)
    End Sub

    Private Sub txtShipAddrID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtShipAddrID.Leave
        If txtShipAddrID.Text.Trim.Length <> 0 Then
            Me.myShipAddressID = txtShipAddrID.Text
        Else
            Me.myShipAddressID = 0
        End If
        Me.RBShipAddress.Checked = True
        Me.RadioButtonEx1_CheckedChanged(Nothing, Nothing)
    End Sub

    Private Sub cbCustomer_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If blnNoInvoice Then Exit Sub
        If blnChanging Then Exit Sub

        If Me.cbCustomer.SelectedIndex > -1 Then
            Dim CustId As String = Me.cbCustomer.SelectedValue
            Dim oCust As Intranet.intranet.customerns.Client = cSupport.CustList.Item(CustId)
            If Not oCust Is Nothing Then
                Me.ToolTip1.SetToolTip(Me.cbCustomer, oCust.InfoAcc)
            End If
            If CustId <> myInvoice.ClientID Then
                MsgBox("As you have changed the customer for this invoice please check " & vbCrLf & _
                "The sage id" & vbCrLf & _
                "The Currency" & vbCrLf & _
                "The Exchange Rate ", MsgBoxStyle.OKOnly)
            End If
        End If

    End Sub

    Private Sub cbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCustomer.SelectedIndexChanged
        If blnNoInvoice Then Exit Sub
        If blnChanging Then Exit Sub

        If Me.cbCustomer.SelectedIndex > -1 Then
            Dim oCust As Intranet.intranet.customerns.Client = cSupport.CustList.Item(Me.cbCustomer.SelectedIndex)
            Dim CustId As String
            If Not oCust Is Nothing Then
                CustId = oCust.Identity
                Me.ToolTip1.SetToolTip(Me.cbCustomer, oCust.InfoAcc)
            End If
            If CustId.Trim.Length > 0 AndAlso CustId <> myInvoice.ClientID Then
                MsgBox("As you have changed the customer for this invoice please check " & vbCrLf & _
                "The sage id" & vbCrLf & _
                "The Currency" & vbCrLf & _
                "The Exchange Rate ", MsgBoxStyle.OKOnly)
            End If
        End If

    End Sub

    Private Sub CBInvDest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBInvDest.SelectedIndexChanged
        If Me.CBInvDest.SelectedIndex < 1 Then Exit Sub
        If Me.CBInvDest.SelectedIndex = 1 Then
            Me.ToolTip1.SetToolTip(Me.cmdEmail, strOutputType & " original of the invoice to " & strUserDestText)
        ElseIf Me.CBInvDest.SelectedIndex = 2 Then
            Me.ToolTip1.SetToolTip(Me.cmdEmail, strOutputType & " copy of the invoice to " & strUserDestText)
        ElseIf Me.CBInvDest.SelectedIndex = 3 Then
            Me.ToolTip1.SetToolTip(Me.cmdEmail, strOutputType & " delivery note to " & strUserDestText)
        End If
    End Sub

    Private Sub SetAddressLocal(ByVal AddressID As Integer, ByVal Email As Boolean)
        Dim lAddress As MedscreenLib.Address.Caddress

        Me.txtDestination.Text = ""

        If AddressID < 1 Then Exit Sub

        lAddress = cSupport.SMAddresses.item(AddressID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
        If Not lAddress Is Nothing Then
            If Email Then
                Me.txtDestination.Text = lAddress.Email
            Else
                Me.txtDestination.Text = lAddress.Fax.Trim.Replace(" ", "") & "@starfax.co.uk"
            End If
        End If
    End Sub

    Private Sub SetDestAddress()
        If Me.RBEmail.Checked Then
            Me.cmdEmail.Text = "&Email Invoice"
            Me.txtDestination.Show()
            Me.CBPrinters.Hide()
            Me.strOutputType = "Email"
            If Me.RBSelf.Checked Then
                Me.txtDestination.Text = CurrentSMUserEmail()
                strUserDestText = "the currently logged in user."
            ElseIf Me.RBMail.Checked Then
                SetAddressLocal(myInvoice.MailaddressID, True)
                strUserDestText = "the mail address"
            ElseIf Me.RBInvoice.Checked Then
                SetAddressLocal(myInvoice.InvoiceAddressID, True)
                strUserDestText = "the invoice address"
            ElseIf Me.RBShipping.Checked Then
                SetAddressLocal(myInvoice.ShippingAddressID, True)
                strUserDestText = "the shipping address"
            End If
        ElseIf Me.RBFax.Checked Then
            Me.cmdEmail.Text = "&Fax Invoice"
            Me.txtDestination.Show()
            Me.CBPrinters.Hide()
            Me.strOutputType = "Fax"
            If Me.RBSelf.Checked Then
                Me.txtDestination.Text = ""
                strUserDestText = ""
            ElseIf Me.RBMail.Checked Then
                SetAddressLocal(myInvoice.MailaddressID, False)
                strUserDestText = "the mail address"
            ElseIf Me.RBInvoice.Checked Then
                SetAddressLocal(myInvoice.InvoiceAddressID, False)
                strUserDestText = "the invoice address"
            ElseIf Me.RBShipping.Checked Then
                SetAddressLocal(myInvoice.ShippingAddressID, False)
                strUserDestText = "the shipping address"
            End If
        ElseIf Me.RBPrint.Checked Then
            Me.cmdEmail.Text = "&Print Invoice"
            Me.txtDestination.Hide()
            Me.CBPrinters.Show()
            Me.strOutputType = "Fax"
            If Me.RBSelf.Checked Then
                Me.txtDestination.Text = ""
                strUserDestText = ""
            ElseIf Me.RBMail.Checked Then
                SetAddressLocal(myInvoice.MailaddressID, False)
                strUserDestText = "the mail address"
            ElseIf Me.RBInvoice.Checked Then
                SetAddressLocal(myInvoice.InvoiceAddressID, False)
                strUserDestText = "the invoice address"
            ElseIf Me.RBShipping.Checked Then
                SetAddressLocal(myInvoice.ShippingAddressID, False)
                strUserDestText = "the shipping address"
            End If
            If Me.CBPrinters.SelectedIndex > -1 Then
                Me.txtDestination.Text = prColl.Item(CBPrinters.SelectedIndex).Header
            End If
        End If
        CBInvDest_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub RBEmail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBEmail.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub RBFax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBFax.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub RBSelf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBSelf.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub RBMail_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBMail.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub RBInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBInvoice.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub RBShipping_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBShipping.CheckedChanged
        SetDestAddress()
    End Sub

    Private Sub cmdEditAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditAddress.Click
        Me.AddressPanel1.isReadOnly = False
        'Me.txtContact2.Enabled = True
        Me.cmdSave.Enabled = True
    End Sub


    Private Sub cmdDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteRow.Click
        If Me.DataGrid1.CurrentRowIndex > -1 Then
            Dim iir As Intranet.Invoice.INVOICE_ITEMSRow
            iir = ds.INVOICE_ITEMS.Rows(Me.DataGrid1.CurrentRowIndex)
            iir.Deleted = "*"
        End If
    End Sub

    Private Sub cmdInsertLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertLine.Click
        Dim objII As Intranet.intranet.jobs.InvoiceItem = _
            myInvoice.InvoiceItems.Create(myInvoice.InvoiceItems.MaxLineNumber + 1)
        With objII
            Me.ds.INVOICE_ITEMS.AddINVOICE_ITEMSRow(.LineNumber, .ItemCode, .Quantity, _
                .Discount, .Comment, .Price, .Reference, " ", 0.0, " ", _
                Me.ds.Invoice.Rows(0))
        End With
        Me.ds.INVOICE_ITEMS.AcceptChanges()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If Not myInvoice Is Nothing Then
            Me.myInvoice.DataFields.Rollback()
        End If
    End Sub

    Private Sub RBPrint_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBPrint.CheckedChanged
        SetDestAddress()
    End Sub



    Private Sub cmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreview.Click
        If Me.myInvoice Is Nothing Then Exit Sub

        Dim ACCInt As AccountsInterface = New AccountsInterface()
        With ACCInt
            .Status = MedscreenLib.Constants.GCST_IFace_StatusRequest
            .Reference = Me.myInvoice.InvoiceID
            .UserId = CurrentSMUser.Identity

            .InvoiceType = myInvoice.TransactionType
            .RequestCode = "PRXP"

            .Update()

            Dim strErrorMessage As String = ""
            Dim ChRet As Char

            ChRet = .Wait(strErrorMessage)
            While ChRet <> MedscreenLib.Constants.GCST_IFace_StatusCreated _
                And strErrorMessage.Trim.Length = 0

                System.Threading.Thread.CurrentThread.Sleep(200)
                ChRet = .Wait(strErrorMessage)
            End While

            If strErrorMessage.Trim.Length <> 0 Then
                MsgBox(strErrorMessage, MsgBoxStyle.OKOnly)
            Else
                Dim strXML As String
                Dim sFileName As String = _
                    Medscreen.LiveRoot & "userfiles\sage export\Preview " & Me.myInvoice.InvoiceID & ".xml"
                Try
                    Dim cString As String = "C:\Program Files\Internet Explorer\iexplore.exe " & sFileName
                    Dim Pid As Integer = Shell(cString, AppWinStyle.MaximizedFocus, True, 2000)

                    'Medscreen.LoadHtml(sFileName)

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    IO.File.Delete(sFileName)
                End Try
            End If
            'Me.lblEmailed.Show()
            ACCInt.Status = Constants.GCST_IFace_StatusDelete
            ACCInt.Delete()
        End With

    End Sub



    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged


        If Me.TabControl1.SelectedIndex = 2 Then
            Dim strXML As String = myInvoice.InvoiceHistory.ToXML(myInvoice.ToXML)
            Try
                Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "InvoiceHistory.xsl", 0)

                Me.HtmlEditor1.DocumentText = (strHTML)
            Catch ex As Exception
            End Try

        End If

        If Me.TabControl1.SelectedIndex = 1 AndAlso myInvoice Is Nothing Then
            MsgBox("You can't alter addresses without an invoice loaded")
        End If
    End Sub

    Private Sub cmdSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSend.Click
        'Code to lock an invoice basically it needs to call the accounts interface with various parameters set.


        Dim iRet As MsgBoxResult = MsgBox("If the send date is correct then select Ok, any errors select cancel." & vbCrLf & _
            "Ensure that the correct send address and method has been choosen.", MsgBoxStyle.OKCancel)

        If iRet = MsgBoxResult.Cancel Then Exit Sub

        Dim accInt As AccountsInterface = New AccountsInterface()

        With accInt
            .Status = Constants.GCST_IFace_StatusRequest
            .UserId = CurrentSMUser.Identity
            .SendDate = Me.DTSentOn.Value
            .Reference = Me.myInvoice.InvoiceID
            .RequestCode = "SEND"

            'Set up destination of the invoice 
            .SendAddress = Me.txtDestination.Text
            If Me.CBInvDest.SelectedIndex > -1 Then
                .CopyType = Me.CBInvDest.SelectedIndex
            Else
                .CopyType = 2
            End If
            .InvoiceType = myInvoice.TransactionType

            .Update()

            'Look for return waiting for a message or Completed flag
            Dim strErrorMessage As String = ""
            Dim ChRet As Char

            ChRet = .Wait(strErrorMessage)
            While ChRet <> MedscreenLib.Constants.GCST_IFace_StatusCreated _
                And strErrorMessage.Trim.Length = 0

                System.Threading.Thread.CurrentThread.Sleep(200)
                ChRet = .Wait(strErrorMessage)
            End While

            'If we have a message display it
            If strErrorMessage.Trim.Length <> 0 Then
                MsgBox(strErrorMessage, MsgBoxStyle.OKOnly)
            End If

            'Delete entry to clear table
            accInt.Delete()

        End With
    End Sub



    Private Sub CKLocked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CKLocked.CheckedChanged
        Me.EnableControls(Not CKLocked.Checked)
    End Sub


    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Me.txtInvNumber.Focus()
        Application.DoEvents()
    End Sub
End Class
