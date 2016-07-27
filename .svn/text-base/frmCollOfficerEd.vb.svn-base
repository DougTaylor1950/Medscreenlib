'$Revision: 1.1 $
'$Author: taylor $
'$Date: 2005-12-12 10:51:31+00 $
'$Log: frmCollOfficerEd.vb,v $
'Revision 1.1  2005-12-12 10:51:31+00  taylor
'Added control over pay roll number
'
'Revision 1.0  2005-11-29 08:11:28+00  taylor
'Commented and holiday list item moved to Medscreen common GUI
'
'Revision 1.2  2005-06-16 17:39:06+01  taylor
'Milestone Check in
'
'Revision 1.1  2004-05-04 10:05:31+01  taylor
'Header addition check in
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports Intranet.collectorpay
Imports Intranet.collectorpay.ExpenseTypeCollections
Imports Intranet.intranet.jobs
Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports MedscreenCommonGui.ListViewItems
Imports Intranet.intranet.Support
Imports System.Windows.Forms
Imports System.Drawing.Printing


''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCollOfficerEd
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for editing Collecting Officer Details
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-12-12 10:51:31+00 $</date><Action>$Revision: 1.1 $</Action></revision>
''' <revision><Author>[taylor]</Author><date>  2005-12-12 10:51:31+00 </date><Action>Revision: 1.1 </Action></revision>
'''  </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCollOfficerEd
    Inherits System.Windows.Forms.Form

    Private loading As Boolean = True
    Friend WithEvents TabComms As System.Windows.Forms.TabPage
    Friend WithEvents wbComms As System.Windows.Forms.WebBrowser
    Friend WithEvents txtPayMailAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents cmdAddExpense As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader18 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader19 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader20 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader21 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader29 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddMiscPay As System.Windows.Forms.Button
    Friend WithEvents cbPayType2 As System.Windows.Forms.ComboBox
    Friend WithEvents HtmlPay As System.Windows.Forms.WebBrowser
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents cmdFind As System.Windows.Forms.Button
    Friend WithEvents cbPayType As System.Windows.Forms.ComboBox
    Friend WithEvents cmdAddJob As System.Windows.Forms.Button
    Friend WithEvents cmdAddOCR As System.Windows.Forms.Button
    Private blnCanEnterPay As Boolean = False
    Friend WithEvents txtTTOfficerID As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents chPortCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPortName As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNoColls As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCostPerDonor As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPLocation As System.Windows.Forms.ColumnHeader
    Private blnDisplayPay As Boolean = False
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.cbCountry.DataSource = MedscreenLib.Glossary.Glossary.Countries
        Me.cbCountry.DisplayMember = "CountryName"
        Me.cbCountry.ValueMember = "CountryId"

        Me.cbSendOption.DataSource = MedscreenLib.Glossary.Glossary.SendOption
        Me.cbSendOption.DisplayMember = "PhraseText"
        Me.cbSendOption.ValueMember = "PhraseID"

        Dim objPaymethods As New MedscreenLib.Glossary.PhraseCollection("PAYMETHOD")
        objPaymethods.Load()
        Me.cbPaymethod.DataSource = objPaymethods
        Me.cbPaymethod.DisplayMember = "PhraseText"
        Me.cbPaymethod.ValueMember = "PhraseID"

        Dim objSubDivisions As New MedscreenLib.Glossary.PhraseCollection("OFF_DIVISN")
        objSubDivisions.Load()
        Me.cbSubDivision.DataSource = objSubDivisions
        Me.cbSubDivision.DisplayMember = "PhraseText"
        Me.cbSubDivision.ValueMember = "PhraseID"



        Dim objDivisions As New MedscreenLib.Glossary.PhraseCollection("select identity,company_name from customer", Glossary.PhraseCollection.BuildBy.Query, "removeflag = 'F' and customertype = 'SUPPLIER'")
        'objDivisions.load() 'Auto loaded on create
        Me.cbDivision.DataSource = objDivisions
        Me.cbDivision.DisplayMember = "PhraseText"
        Me.cbDivision.ValueMember = "PhraseID"

        'Get currencies from currency table
        Dim objCurr As New Glossary.PhraseCollection("Select Currencyid,description from currency", Glossary.PhraseCollection.BuildBy.Query)
        objCurr.Load()
        Me.cbCurrency.DataSource = objCurr
        Me.cbCurrency.ValueMember = "PhraseID"
        Me.cbCurrency.DisplayMember = "PhraseText"

        SetOutputMenu()

        Dim objPayTypes As New Glossary.PhraseCollection("select e.exptypeid,l.description from expense_type e, lineitems l  where e.itemcode = l.itemcode and expensegroup = 'PAY' and expensesubgroup = 'LIST'", Glossary.PhraseCollection.BuildBy.Query)
        objPayTypes.Load()
        Me.cbPayType.DataSource = objPayTypes
        Me.cbPayType.ValueMember = "PhraseID"
        Me.cbPayType.DisplayMember = "PhraseText"
        Me.cbPayType2.DataSource = objPayTypes
        Me.cbPayType2.ValueMember = "PhraseID"
        Me.cbPayType2.DisplayMember = "PhraseText"

        loading = False
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub SetOutputMenu()
        Me.mnuOutputDefault.Text = "&Output Default - " & [Enum].GetName(GetType(MedscreenLib.Constants.SendMethod), MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod)
        Me.mnuDefaultPrinter.Text = "&Default Printer - " & MedscreenCommonGui.CommonForms.UserOptions.DefaultPrinter
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents txtSurName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtForeName As System.Windows.Forms.TextBox
    Friend WithEvents lblForeName As System.Windows.Forms.Label
    Friend WithEvents txtCollID As System.Windows.Forms.TextBox
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tabPaydetails As System.Windows.Forms.TabPage
    Friend WithEvents tabAddress As System.Windows.Forms.TabPage
    Friend WithEvents lblPayMethod As System.Windows.Forms.Label
    Friend WithEvents lblPayCurrency As System.Windows.Forms.Label
    Friend WithEvents lblRateSample As System.Windows.Forms.Label
    Friend WithEvents txtSampleRate As System.Windows.Forms.TextBox
    Friend WithEvents txtHourlyRate As System.Windows.Forms.TextBox
    Friend WithEvents lblHourlyRate As System.Windows.Forms.Label
    Friend WithEvents txtPayrollNo As System.Windows.Forms.TextBox
    Friend WithEvents lblPayRollNo As System.Windows.Forms.Label
    Friend WithEvents txtCancSamp As System.Windows.Forms.TextBox
    Friend WithEvents lblCancSamp As System.Windows.Forms.Label
    Friend WithEvents txtMinCharge As System.Windows.Forms.TextBox
    Friend WithEvents lblMinCharge As System.Windows.Forms.Label
    Friend WithEvents tabBankDetails As System.Windows.Forms.TabPage
    Friend WithEvents tabTraining As System.Windows.Forms.TabPage
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents lblBankName As System.Windows.Forms.Label
    Friend WithEvents txtBankName As System.Windows.Forms.TextBox
    Friend WithEvents txtSortCode As System.Windows.Forms.TextBox
    Friend WithEvents lblBankSortCode As System.Windows.Forms.Label
    Friend WithEvents txtKitEdit As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNINum As System.Windows.Forms.TextBox
    Friend WithEvents lblDob As System.Windows.Forms.Label
    Friend WithEvents dtDOB As System.Windows.Forms.DateTimePicker
    Friend WithEvents tabHolidays As System.Windows.Forms.TabPage
    Friend WithEvents lvHolidays As System.Windows.Forms.ListView
    Friend WithEvents ctxHoliday As System.Windows.Forms.ContextMenu
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chListPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLargeIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSmallIcons As System.Windows.Forms.MenuItem
    Friend WithEvents MnuList As System.Windows.Forms.MenuItem
    Friend WithEvents MnuDetails As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBar1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveHol As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddHol As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdHoliday As System.Windows.Forms.MenuItem
    Friend WithEvents imgSmall As System.Windows.Forms.ImageList
    Friend WithEvents imgLarge As System.Windows.Forms.ImageList
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtBadge As System.Windows.Forms.TextBox
    Friend WithEvents DTExpire As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents TabPort As System.Windows.Forms.TabPage
    Friend WithEvents lvOfficerPorts As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BankAddress As MedscreenCommonGui.AddressPanel
    Friend WithEvents adrPanelOfficer As MedscreenCommonGui.AddressPanel
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents TabComments As System.Windows.Forms.TabPage
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblFaxError As System.Windows.Forms.Label
    Friend WithEvents lblEmailError As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lvComments As System.Windows.Forms.ListView

    Friend WithEvents chCommDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommType As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHComment As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommBy As System.Windows.Forms.ColumnHeader
    Friend WithEvents ChCommAction As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddComment As System.Windows.Forms.Button

    Friend WithEvents ctxPorts As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddPort As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemovePort As System.Windows.Forms.MenuItem
    Friend WithEvents CHPortCountry As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuEditPort As System.Windows.Forms.MenuItem
    Friend WithEvents TabPay As System.Windows.Forms.TabPage
    Friend WithEvents TabExpenses As System.Windows.Forms.TabPage
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDiscount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVALUE As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader

    'Friend WithEvents cmdFindBankAddress As System.Windows.Forms.Button
    'Friend WithEvents cmdNewBankAddress As System.Windows.Forms.Button

    Friend WithEvents cmdFindBankAddress As System.Windows.Forms.Button
    Friend WithEvents cmdNewBankAddress As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents cmdFindAddress As System.Windows.Forms.Button
    Friend WithEvents cmdNewAddress As System.Windows.Forms.Button
    Friend WithEvents lblOfficePhone As System.Windows.Forms.Label
    Friend WithEvents lblOfficeFax As System.Windows.Forms.Label
    Friend WithEvents txtOfficeFax As System.Windows.Forms.TextBox
    Friend WithEvents txtPager As System.Windows.Forms.TextBox
    Friend WithEvents lblPager As System.Windows.Forms.Label
    Friend WithEvents txtMobile As System.Windows.Forms.TextBox
    Friend WithEvents lblMobile As System.Windows.Forms.Label
    Friend WithEvents txtOfficePhone As System.Windows.Forms.TextBox
    Friend WithEvents cbSendOption As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvExpenses As System.Windows.Forms.ListView
    Friend WithEvents lvExpenseHeader As System.Windows.Forms.ListView
    Friend WithEvents LvCostHeader As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader22 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuReports As System.Windows.Forms.MenuItem
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents txtOutOfHours As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtCOOutHours As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtCOInHours As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents ctxPayHeader As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuApproveItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeclineItem As System.Windows.Forms.MenuItem
    Friend WithEvents MnuReferItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCancelItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEntered As System.Windows.Forms.MenuItem
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuPaySummary As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPayStatement As System.Windows.Forms.MenuItem
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents dtPayMonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFaxWarning As System.Windows.Forms.Label
    Friend WithEvents mnuLockPay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaySummary2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOutputDefault As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDefaultPrinter As System.Windows.Forms.MenuItem
    Friend WithEvents ErrorProvider2 As System.Windows.Forms.ErrorProvider
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents ckPinPerson As System.Windows.Forms.CheckBox
    Friend WithEvents txtAreaFee As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtTransFee As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents lvPayRates As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader2a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8a As System.Windows.Forms.ColumnHeader
    Friend WithEvents ctxDiscount As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuPayRateStartDate As System.Windows.Forms.MenuItem
    Friend WithEvents lblNopayentryRole As System.Windows.Forms.Label
    Friend WithEvents TabEquipment As System.Windows.Forms.TabPage
    Friend WithEvents lvEquipment As System.Windows.Forms.ListView
    Friend WithEvents CHModel As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHSerialNo As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHLastService As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHLastCalib As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPay As System.Windows.Forms.ColumnHeader
    Friend WithEvents chExp As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTotal As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbPaymethod As System.Windows.Forms.ComboBox
    Friend WithEvents cmdaddInstrument As System.Windows.Forms.Button
    Friend WithEvents chRemoved As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cbDivision As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cbSubDivision As System.Windows.Forms.ComboBox
    Friend WithEvents chBanded As System.Windows.Forms.ColumnHeader
    Friend WithEvents chStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents chComments As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuCollectionByPort As System.Windows.Forms.MenuItem
    Friend WithEvents txtVolumeDiscount As System.Windows.Forms.TextBox
    Friend WithEvents TabOrders As System.Windows.Forms.TabPage
    Friend WithEvents lvOrders As System.Windows.Forms.ListView
    Friend WithEvents chInvNumber As System.Windows.Forms.ColumnHeader
    Friend WithEvents chInvoicedOn As System.Windows.Forms.ColumnHeader
    Friend WithEvents chInvoiceValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDespatch As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvOrderItems As System.Windows.Forms.ListView
    Friend WithEvents chItem As System.Windows.Forms.ColumnHeader
    Friend WithEvents chOrderQty As System.Windows.Forms.ColumnHeader
    Friend WithEvents udMaxOrders As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceDisp As System.Windows.Forms.Label
    Friend WithEvents mnuEmailSummary As System.Windows.Forms.MenuItem
    Friend WithEvents ctxPayEdit As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuEditPay As System.Windows.Forms.MenuItem
    Friend WithEvents chActPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents TabLocation As System.Windows.Forms.TabPage
    Friend WithEvents pnlLocation As System.Windows.Forms.Panel
    Friend WithEvents txtRegion As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLongitude As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtLatitude As System.Windows.Forms.TextBox
    Friend WithEvents lLatitude As System.Windows.Forms.Label
    Friend WithEvents pnlTeam As System.Windows.Forms.Panel
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents treTeams As System.Windows.Forms.TreeView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents pnlTeamHTML As System.Windows.Forms.Panel
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtCoordinates As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader30 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader31 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader17 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter



    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCollOfficerEd))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label17 = New System.Windows.Forms.Label
        Me.dtPayMonth = New System.Windows.Forms.DateTimePicker
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tabPaydetails = New System.Windows.Forms.TabPage
        Me.txtTTOfficerID = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.txtVolumeDiscount = New System.Windows.Forms.TextBox
        Me.cbPaymethod = New System.Windows.Forms.ComboBox
        Me.lblNopayentryRole = New System.Windows.Forms.Label
        Me.ckPinPerson = New System.Windows.Forms.CheckBox
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.txtTransFee = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtAreaFee = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.lblFaxWarning = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtCOOutHours = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtCOInHours = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtOutOfHours = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.cbCurrency = New System.Windows.Forms.ComboBox
        Me.txtCancSamp = New System.Windows.Forms.TextBox
        Me.lblCancSamp = New System.Windows.Forms.Label
        Me.txtMinCharge = New System.Windows.Forms.TextBox
        Me.lblMinCharge = New System.Windows.Forms.Label
        Me.txtPayrollNo = New System.Windows.Forms.TextBox
        Me.lblPayRollNo = New System.Windows.Forms.Label
        Me.txtHourlyRate = New System.Windows.Forms.TextBox
        Me.lblHourlyRate = New System.Windows.Forms.Label
        Me.txtSampleRate = New System.Windows.Forms.TextBox
        Me.lblRateSample = New System.Windows.Forms.Label
        Me.lblPayCurrency = New System.Windows.Forms.Label
        Me.lblPayMethod = New System.Windows.Forms.Label
        Me.txtKitEdit = New System.Windows.Forms.TextBox
        Me.lvPayRates = New ListViewEx.ListViewEx
        Me.ColumnHeader2a = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3a = New System.Windows.Forms.ColumnHeader
        Me.chListPriceKit = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5a = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6a = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7a = New System.Windows.Forms.ColumnHeader
        Me.chPriceKit = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8a = New System.Windows.Forms.ColumnHeader
        Me.chBanded = New System.Windows.Forms.ColumnHeader
        Me.ctxDiscount = New System.Windows.Forms.ContextMenu
        Me.mnuPayRateStartDate = New System.Windows.Forms.MenuItem
        Me.TabPay = New System.Windows.Forms.TabPage
        Me.HtmlPay = New System.Windows.Forms.WebBrowser
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.cmdFind = New System.Windows.Forms.Button
        Me.cbPayType = New System.Windows.Forms.ComboBox
        Me.cmdAddJob = New System.Windows.Forms.Button
        Me.cmdAddOCR = New System.Windows.Forms.Button
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.LvCostHeader = New System.Windows.Forms.ListView
        Me.ColumnHeader22 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader
        Me.chPay = New System.Windows.Forms.ColumnHeader
        Me.chExp = New System.Windows.Forms.ColumnHeader
        Me.chTotal = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader30 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader31 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader17 = New System.Windows.Forms.ColumnHeader
        Me.ctxPayHeader = New System.Windows.Forms.ContextMenu
        Me.mnuApproveItem = New System.Windows.Forms.MenuItem
        Me.mnuDeclineItem = New System.Windows.Forms.MenuItem
        Me.MnuReferItem = New System.Windows.Forms.MenuItem
        Me.mnuCancelItem = New System.Windows.Forms.MenuItem
        Me.mnuEntered = New System.Windows.Forms.MenuItem
        Me.mnuLockPay = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuPayStatement = New System.Windows.Forms.MenuItem
        Me.TabComments = New System.Windows.Forms.TabPage
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.lvComments = New System.Windows.Forms.ListView
        Me.chCommDate = New System.Windows.Forms.ColumnHeader
        Me.chCommType = New System.Windows.Forms.ColumnHeader
        Me.CHComment = New System.Windows.Forms.ColumnHeader
        Me.chCommBy = New System.Windows.Forms.ColumnHeader
        Me.ChCommAction = New System.Windows.Forms.ColumnHeader
        Me.cmdAddComment = New System.Windows.Forms.Button
        Me.TabExpenses = New System.Windows.Forms.TabPage
        Me.cbPayType2 = New System.Windows.Forms.ComboBox
        Me.cmdAddMiscPay = New System.Windows.Forms.Button
        Me.cmdAddExpense = New System.Windows.Forms.Button
        Me.lvExpenses = New System.Windows.Forms.ListView
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.chPrice = New System.Windows.Forms.ColumnHeader
        Me.chDiscount = New System.Windows.Forms.ColumnHeader
        Me.chActPrice = New System.Windows.Forms.ColumnHeader
        Me.chVALUE = New System.Windows.Forms.ColumnHeader
        Me.chStatus = New System.Windows.Forms.ColumnHeader
        Me.chComments = New System.Windows.Forms.ColumnHeader
        Me.lvExpenseHeader = New System.Windows.Forms.ListView
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader18 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader19 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader20 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader21 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader29 = New System.Windows.Forms.ColumnHeader
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.tabAddress = New System.Windows.Forms.TabPage
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.txtPayMailAddress = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.cmdFindAddress = New System.Windows.Forms.Button
        Me.cmdNewAddress = New System.Windows.Forms.Button
        Me.lblOfficePhone = New System.Windows.Forms.Label
        Me.lblOfficeFax = New System.Windows.Forms.Label
        Me.txtOfficeFax = New System.Windows.Forms.TextBox
        Me.txtPager = New System.Windows.Forms.TextBox
        Me.lblPager = New System.Windows.Forms.Label
        Me.txtMobile = New System.Windows.Forms.TextBox
        Me.lblMobile = New System.Windows.Forms.Label
        Me.txtOfficePhone = New System.Windows.Forms.TextBox
        Me.cbSendOption = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TabPort = New System.Windows.Forms.TabPage
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.lvOfficerPorts = New System.Windows.Forms.ListView
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.CHPortCountry = New System.Windows.Forms.ColumnHeader
        Me.chPortCode = New System.Windows.Forms.ColumnHeader
        Me.chPortName = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ctxPorts = New System.Windows.Forms.ContextMenu
        Me.mnuAddPort = New System.Windows.Forms.MenuItem
        Me.mnuRemovePort = New System.Windows.Forms.MenuItem
        Me.mnuEditPort = New System.Windows.Forms.MenuItem
        Me.TabLocation = New System.Windows.Forms.TabPage
        Me.pnlTeamHTML = New System.Windows.Forms.Panel
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.pnlTeam = New System.Windows.Forms.Panel
        Me.treTeams = New System.Windows.Forms.TreeView
        Me.Label23 = New System.Windows.Forms.Label
        Me.pnlLocation = New System.Windows.Forms.Panel
        Me.txtCoordinates = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.txtLongitude = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtLatitude = New System.Windows.Forms.TextBox
        Me.lLatitude = New System.Windows.Forms.Label
        Me.txtRegion = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cbCountry = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tabBankDetails = New System.Windows.Forms.TabPage
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.txtSortCode = New System.Windows.Forms.TextBox
        Me.lblBankSortCode = New System.Windows.Forms.Label
        Me.txtBankName = New System.Windows.Forms.TextBox
        Me.lblBankName = New System.Windows.Forms.Label
        Me.tabTraining = New System.Windows.Forms.TabPage
        Me.tabHolidays = New System.Windows.Forms.TabPage
        Me.lvHolidays = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ctxHoliday = New System.Windows.Forms.ContextMenu
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuLargeIcons = New System.Windows.Forms.MenuItem
        Me.mnuSmallIcons = New System.Windows.Forms.MenuItem
        Me.MnuList = New System.Windows.Forms.MenuItem
        Me.MnuDetails = New System.Windows.Forms.MenuItem
        Me.mnuBar1 = New System.Windows.Forms.MenuItem
        Me.mnuRemoveHol = New System.Windows.Forms.MenuItem
        Me.mnuAddHol = New System.Windows.Forms.MenuItem
        Me.mnuEdHoliday = New System.Windows.Forms.MenuItem
        Me.imgLarge = New System.Windows.Forms.ImageList(Me.components)
        Me.imgSmall = New System.Windows.Forms.ImageList(Me.components)
        Me.TabEquipment = New System.Windows.Forms.TabPage
        Me.cmdaddInstrument = New System.Windows.Forms.Button
        Me.lvEquipment = New System.Windows.Forms.ListView
        Me.CHModel = New System.Windows.Forms.ColumnHeader
        Me.CHSerialNo = New System.Windows.Forms.ColumnHeader
        Me.CHLastService = New System.Windows.Forms.ColumnHeader
        Me.CHLastCalib = New System.Windows.Forms.ColumnHeader
        Me.chRemoved = New System.Windows.Forms.ColumnHeader
        Me.TabOrders = New System.Windows.Forms.TabPage
        Me.lblInvoiceDisp = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.udMaxOrders = New System.Windows.Forms.NumericUpDown
        Me.lvOrderItems = New System.Windows.Forms.ListView
        Me.chOrderQty = New System.Windows.Forms.ColumnHeader
        Me.chItem = New System.Windows.Forms.ColumnHeader
        Me.lvOrders = New System.Windows.Forms.ListView
        Me.chInvNumber = New System.Windows.Forms.ColumnHeader
        Me.chInvoicedOn = New System.Windows.Forms.ColumnHeader
        Me.chInvoiceValue = New System.Windows.Forms.ColumnHeader
        Me.chDespatch = New System.Windows.Forms.ColumnHeader
        Me.TabComms = New System.Windows.Forms.TabPage
        Me.wbComms = New System.Windows.Forms.WebBrowser
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.cbSubDivision = New System.Windows.Forms.ComboBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.cbDivision = New System.Windows.Forms.ComboBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.ckRemoved = New System.Windows.Forms.CheckBox
        Me.DTExpire = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtBadge = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtDOB = New System.Windows.Forms.DateTimePicker
        Me.lblDob = New System.Windows.Forms.Label
        Me.txtNINum = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSurName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtForeName = New System.Windows.Forms.TextBox
        Me.lblForeName = New System.Windows.Forms.Label
        Me.txtCollID = New System.Windows.Forms.TextBox
        Me.lblID = New System.Windows.Forms.Label
        Me.ctxPayEdit = New System.Windows.Forms.ContextMenu
        Me.mnuEditPay = New System.Windows.Forms.MenuItem
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuReports = New System.Windows.Forms.MenuItem
        Me.mnuPaySummary = New System.Windows.Forms.MenuItem
        Me.mnuPaySummary2 = New System.Windows.Forms.MenuItem
        Me.mnuOutputDefault = New System.Windows.Forms.MenuItem
        Me.mnuDefaultPrinter = New System.Windows.Forms.MenuItem
        Me.mnuCollectionByPort = New System.Windows.Forms.MenuItem
        Me.mnuEmailSummary = New System.Windows.Forms.MenuItem
        Me.ErrorProvider2 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.chNoColls = New System.Windows.Forms.ColumnHeader
        Me.chCostPerDonor = New System.Windows.Forms.ColumnHeader
        Me.adrPanelOfficer = New MedscreenCommonGui.AddressPanel
        Me.lblFaxError = New System.Windows.Forms.Label
        Me.lblEmailError = New System.Windows.Forms.Label
        Me.BankAddress = New MedscreenCommonGui.AddressPanel
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.cmdFindBankAddress = New System.Windows.Forms.Button
        Me.cmdNewBankAddress = New System.Windows.Forms.Button
        Me.chPLocation = New System.Windows.Forms.ColumnHeader
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.tabPaydetails.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.TabPay.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.TabComments.SuspendLayout()
        Me.TabExpenses.SuspendLayout()
        Me.tabAddress.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.TabPort.SuspendLayout()
        Me.TabLocation.SuspendLayout()
        Me.pnlTeamHTML.SuspendLayout()
        Me.pnlTeam.SuspendLayout()
        Me.pnlLocation.SuspendLayout()
        Me.tabBankDetails.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.tabHolidays.SuspendLayout()
        Me.TabEquipment.SuspendLayout()
        Me.TabOrders.SuspendLayout()
        CType(Me.udMaxOrders, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabComms.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.adrPanelOfficer.SuspendLayout()
        Me.BankAddress.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.dtPayMonth)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 400)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(938, 32)
        Me.Panel1.TabIndex = 1
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(24, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(100, 23)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "Pay Month"
        '
        'dtPayMonth
        '
        Me.dtPayMonth.CustomFormat = "MMMM yyyy"
        Me.dtPayMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPayMonth.Location = New System.Drawing.Point(152, 8)
        Me.dtPayMonth.Name = "dtPayMonth"
        Me.dtPayMonth.Size = New System.Drawing.Size(120, 20)
        Me.dtPayMonth.TabIndex = 6
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(854, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(774, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TabControl1)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(938, 400)
        Me.Panel2.TabIndex = 2
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabPaydetails)
        Me.TabControl1.Controls.Add(Me.TabPay)
        Me.TabControl1.Controls.Add(Me.TabComments)
        Me.TabControl1.Controls.Add(Me.TabExpenses)
        Me.TabControl1.Controls.Add(Me.tabAddress)
        Me.TabControl1.Controls.Add(Me.TabPort)
        Me.TabControl1.Controls.Add(Me.TabLocation)
        Me.TabControl1.Controls.Add(Me.tabBankDetails)
        Me.TabControl1.Controls.Add(Me.tabTraining)
        Me.TabControl1.Controls.Add(Me.tabHolidays)
        Me.TabControl1.Controls.Add(Me.TabEquipment)
        Me.TabControl1.Controls.Add(Me.TabOrders)
        Me.TabControl1.Controls.Add(Me.TabComms)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.ItemSize = New System.Drawing.Size(62, 36)
        Me.TabControl1.Location = New System.Drawing.Point(0, 120)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(934, 276)
        Me.TabControl1.TabIndex = 1
        '
        'tabPaydetails
        '
        Me.tabPaydetails.Controls.Add(Me.txtTTOfficerID)
        Me.tabPaydetails.Controls.Add(Me.Label27)
        Me.tabPaydetails.Controls.Add(Me.txtVolumeDiscount)
        Me.tabPaydetails.Controls.Add(Me.cbPaymethod)
        Me.tabPaydetails.Controls.Add(Me.lblNopayentryRole)
        Me.tabPaydetails.Controls.Add(Me.ckPinPerson)
        Me.tabPaydetails.Controls.Add(Me.Panel6)
        Me.tabPaydetails.Controls.Add(Me.lblFaxWarning)
        Me.tabPaydetails.Controls.Add(Me.Label16)
        Me.tabPaydetails.Controls.Add(Me.txtCOOutHours)
        Me.tabPaydetails.Controls.Add(Me.Label14)
        Me.tabPaydetails.Controls.Add(Me.txtCOInHours)
        Me.tabPaydetails.Controls.Add(Me.Label15)
        Me.tabPaydetails.Controls.Add(Me.txtOutOfHours)
        Me.tabPaydetails.Controls.Add(Me.Label13)
        Me.tabPaydetails.Controls.Add(Me.cbCurrency)
        Me.tabPaydetails.Controls.Add(Me.txtCancSamp)
        Me.tabPaydetails.Controls.Add(Me.lblCancSamp)
        Me.tabPaydetails.Controls.Add(Me.txtMinCharge)
        Me.tabPaydetails.Controls.Add(Me.lblMinCharge)
        Me.tabPaydetails.Controls.Add(Me.txtPayrollNo)
        Me.tabPaydetails.Controls.Add(Me.lblPayRollNo)
        Me.tabPaydetails.Controls.Add(Me.txtHourlyRate)
        Me.tabPaydetails.Controls.Add(Me.lblHourlyRate)
        Me.tabPaydetails.Controls.Add(Me.txtSampleRate)
        Me.tabPaydetails.Controls.Add(Me.lblRateSample)
        Me.tabPaydetails.Controls.Add(Me.lblPayCurrency)
        Me.tabPaydetails.Controls.Add(Me.lblPayMethod)
        Me.tabPaydetails.Controls.Add(Me.txtKitEdit)
        Me.tabPaydetails.Controls.Add(Me.lvPayRates)
        Me.tabPaydetails.Location = New System.Drawing.Point(4, 76)
        Me.tabPaydetails.Name = "tabPaydetails"
        Me.tabPaydetails.Size = New System.Drawing.Size(926, 196)
        Me.tabPaydetails.TabIndex = 0
        Me.tabPaydetails.Text = "PayDetails"
        Me.tabPaydetails.UseVisualStyleBackColor = True
        '
        'txtTTOfficerID
        '
        Me.txtTTOfficerID.Location = New System.Drawing.Point(750, 16)
        Me.txtTTOfficerID.Name = "txtTTOfficerID"
        Me.txtTTOfficerID.Size = New System.Drawing.Size(100, 20)
        Me.txtTTOfficerID.TabIndex = 36
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(662, 16)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(64, 23)
        Me.Label27.TabIndex = 35
        Me.Label27.Text = "TT OFF ID"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVolumeDiscount
        '
        Me.txtVolumeDiscount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVolumeDiscount.Location = New System.Drawing.Point(8, 169)
        Me.txtVolumeDiscount.Multiline = True
        Me.txtVolumeDiscount.Name = "txtVolumeDiscount"
        Me.txtVolumeDiscount.Size = New System.Drawing.Size(906, 20)
        Me.txtVolumeDiscount.TabIndex = 34
        '
        'cbPaymethod
        '
        Me.cbPaymethod.Location = New System.Drawing.Point(96, 16)
        Me.cbPaymethod.Name = "cbPaymethod"
        Me.cbPaymethod.Size = New System.Drawing.Size(104, 21)
        Me.cbPaymethod.TabIndex = 33
        '
        'lblNopayentryRole
        '
        Me.lblNopayentryRole.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNopayentryRole.ForeColor = System.Drawing.Color.Red
        Me.lblNopayentryRole.Location = New System.Drawing.Point(24, 200)
        Me.lblNopayentryRole.Name = "lblNopayentryRole"
        Me.lblNopayentryRole.Size = New System.Drawing.Size(480, 40)
        Me.lblNopayentryRole.TabIndex = 32
        Me.lblNopayentryRole.Text = "You are not in a role which allows entry of pay statements, if you think you shou" & _
            "ld be, contact your manager!"
        '
        'ckPinPerson
        '
        Me.ckPinPerson.Location = New System.Drawing.Point(432, 112)
        Me.ckPinPerson.Name = "ckPinPerson"
        Me.ckPinPerson.Size = New System.Drawing.Size(144, 24)
        Me.ckPinPerson.TabIndex = 31
        Me.ckPinPerson.Text = "Area Coordinator"
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.txtTransFee)
        Me.Panel6.Controls.Add(Me.Label19)
        Me.Panel6.Controls.Add(Me.txtAreaFee)
        Me.Panel6.Controls.Add(Me.Label18)
        Me.Panel6.Location = New System.Drawing.Point(40, 144)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(560, 56)
        Me.Panel6.TabIndex = 30
        '
        'txtTransFee
        '
        Me.txtTransFee.Location = New System.Drawing.Point(384, 16)
        Me.txtTransFee.Name = "txtTransFee"
        Me.txtTransFee.Size = New System.Drawing.Size(100, 20)
        Me.txtTransFee.TabIndex = 29
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(296, 16)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(72, 23)
        Me.Label19.TabIndex = 28
        Me.Label19.Text = "Transfer Fee "
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAreaFee
        '
        Me.txtAreaFee.Location = New System.Drawing.Point(120, 16)
        Me.txtAreaFee.Name = "txtAreaFee"
        Me.txtAreaFee.Size = New System.Drawing.Size(100, 20)
        Me.txtAreaFee.TabIndex = 27
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(32, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 23)
        Me.Label18.TabIndex = 26
        Me.Label18.Text = "Area Coord Fee"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFaxWarning
        '
        Me.lblFaxWarning.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFaxWarning.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFaxWarning.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblFaxWarning.Font = New System.Drawing.Font("Impact", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFaxWarning.ForeColor = System.Drawing.Color.Firebrick
        Me.lblFaxWarning.Location = New System.Drawing.Point(32, 240)
        Me.lblFaxWarning.Name = "lblFaxWarning"
        Me.lblFaxWarning.Size = New System.Drawing.Size(834, 196)
        Me.lblFaxWarning.TabIndex = 29
        Me.lblFaxWarning.Text = "Fax is expensive - Try and persuade the collecting officer to switch to email"
        Me.lblFaxWarning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblFaxWarning.Visible = False
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(9, 80)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(72, 23)
        Me.Label16.TabIndex = 28
        Me.Label16.Text = "CallOuts"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCOOutHours
        '
        Me.txtCOOutHours.Location = New System.Drawing.Point(504, 80)
        Me.txtCOOutHours.Name = "txtCOOutHours"
        Me.txtCOOutHours.Size = New System.Drawing.Size(100, 20)
        Me.txtCOOutHours.TabIndex = 27
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(416, 80)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 23)
        Me.Label14.TabIndex = 26
        Me.Label14.Text = "out of hours Rate"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCOInHours
        '
        Me.txtCOInHours.Location = New System.Drawing.Point(296, 80)
        Me.txtCOInHours.Name = "txtCOInHours"
        Me.txtCOInHours.Size = New System.Drawing.Size(100, 20)
        Me.txtCOInHours.TabIndex = 25
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(208, 80)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 23)
        Me.Label15.TabIndex = 24
        Me.Label15.Text = "in hours Rate"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOutOfHours
        '
        Me.txtOutOfHours.Location = New System.Drawing.Point(504, 48)
        Me.txtOutOfHours.Name = "txtOutOfHours"
        Me.txtOutOfHours.Size = New System.Drawing.Size(100, 20)
        Me.txtOutOfHours.TabIndex = 23
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(416, 48)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 23)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "out of hours Rate"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(296, 16)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(121, 21)
        Me.cbCurrency.TabIndex = 21
        '
        'txtCancSamp
        '
        Me.txtCancSamp.Location = New System.Drawing.Point(296, 110)
        Me.txtCancSamp.Name = "txtCancSamp"
        Me.txtCancSamp.Size = New System.Drawing.Size(100, 20)
        Me.txtCancSamp.TabIndex = 20
        '
        'lblCancSamp
        '
        Me.lblCancSamp.Location = New System.Drawing.Point(208, 110)
        Me.lblCancSamp.Name = "lblCancSamp"
        Me.lblCancSamp.Size = New System.Drawing.Size(72, 23)
        Me.lblCancSamp.TabIndex = 19
        Me.lblCancSamp.Text = "Canc Samp Deduction"
        Me.lblCancSamp.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMinCharge
        '
        Me.txtMinCharge.Location = New System.Drawing.Point(96, 112)
        Me.txtMinCharge.Name = "txtMinCharge"
        Me.txtMinCharge.Size = New System.Drawing.Size(100, 20)
        Me.txtMinCharge.TabIndex = 18
        '
        'lblMinCharge
        '
        Me.lblMinCharge.Location = New System.Drawing.Point(8, 112)
        Me.lblMinCharge.Name = "lblMinCharge"
        Me.lblMinCharge.Size = New System.Drawing.Size(72, 23)
        Me.lblMinCharge.TabIndex = 17
        Me.lblMinCharge.Text = "Min Charge"
        Me.lblMinCharge.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPayrollNo
        '
        Me.txtPayrollNo.Enabled = False
        Me.txtPayrollNo.Location = New System.Drawing.Point(536, 16)
        Me.txtPayrollNo.Name = "txtPayrollNo"
        Me.txtPayrollNo.Size = New System.Drawing.Size(100, 20)
        Me.txtPayrollNo.TabIndex = 16
        '
        'lblPayRollNo
        '
        Me.lblPayRollNo.Location = New System.Drawing.Point(448, 16)
        Me.lblPayRollNo.Name = "lblPayRollNo"
        Me.lblPayRollNo.Size = New System.Drawing.Size(64, 23)
        Me.lblPayRollNo.TabIndex = 15
        Me.lblPayRollNo.Text = "Pay Roll No"
        Me.lblPayRollNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHourlyRate
        '
        Me.txtHourlyRate.Location = New System.Drawing.Point(296, 48)
        Me.txtHourlyRate.Name = "txtHourlyRate"
        Me.txtHourlyRate.Size = New System.Drawing.Size(100, 20)
        Me.txtHourlyRate.TabIndex = 14
        '
        'lblHourlyRate
        '
        Me.lblHourlyRate.Location = New System.Drawing.Point(206, 48)
        Me.lblHourlyRate.Name = "lblHourlyRate"
        Me.lblHourlyRate.Size = New System.Drawing.Size(72, 23)
        Me.lblHourlyRate.TabIndex = 13
        Me.lblHourlyRate.Text = "in hours Rate"
        Me.lblHourlyRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSampleRate
        '
        Me.txtSampleRate.Location = New System.Drawing.Point(96, 48)
        Me.txtSampleRate.Name = "txtSampleRate"
        Me.txtSampleRate.Size = New System.Drawing.Size(100, 20)
        Me.txtSampleRate.TabIndex = 12
        '
        'lblRateSample
        '
        Me.lblRateSample.Location = New System.Drawing.Point(8, 48)
        Me.lblRateSample.Name = "lblRateSample"
        Me.lblRateSample.Size = New System.Drawing.Size(72, 23)
        Me.lblRateSample.TabIndex = 11
        Me.lblRateSample.Text = "Sample Rate"
        Me.lblRateSample.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPayCurrency
        '
        Me.lblPayCurrency.Location = New System.Drawing.Point(198, 17)
        Me.lblPayCurrency.Name = "lblPayCurrency"
        Me.lblPayCurrency.Size = New System.Drawing.Size(80, 23)
        Me.lblPayCurrency.TabIndex = 9
        Me.lblPayCurrency.Text = "Pay Currency"
        Me.lblPayCurrency.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPayMethod
        '
        Me.lblPayMethod.Location = New System.Drawing.Point(24, 16)
        Me.lblPayMethod.Name = "lblPayMethod"
        Me.lblPayMethod.Size = New System.Drawing.Size(64, 23)
        Me.lblPayMethod.TabIndex = 0
        Me.lblPayMethod.Text = "Pay Method"
        Me.lblPayMethod.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtKitEdit
        '
        Me.txtKitEdit.Location = New System.Drawing.Point(336, 8)
        Me.txtKitEdit.Name = "txtKitEdit"
        Me.txtKitEdit.Size = New System.Drawing.Size(100, 20)
        Me.txtKitEdit.TabIndex = 3
        Me.txtKitEdit.Visible = False
        '
        'lvPayRates
        '
        Me.lvPayRates.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvPayRates.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2a, Me.ColumnHeader3a, Me.chListPriceKit, Me.ColumnHeader5a, Me.ColumnHeader6a, Me.ColumnHeader7a, Me.chPriceKit, Me.ColumnHeader8a, Me.chBanded})
        Me.lvPayRates.ContextMenu = Me.ctxDiscount
        Me.lvPayRates.FullRowSelect = True
        Me.lvPayRates.GridLines = True
        Me.lvPayRates.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvPayRates.Location = New System.Drawing.Point(8, 240)
        Me.lvPayRates.Name = "lvPayRates"
        Me.lvPayRates.Size = New System.Drawing.Size(906, 196)
        Me.lvPayRates.TabIndex = 1
        Me.lvPayRates.UseCompatibleStateImageBehavior = False
        Me.lvPayRates.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2a
        '
        Me.ColumnHeader2a.Text = "Description/Kit"
        Me.ColumnHeader2a.Width = 300
        '
        'ColumnHeader3a
        '
        Me.ColumnHeader3a.Text = ""
        Me.ColumnHeader3a.Width = 40
        '
        'chListPriceKit
        '
        Me.chListPriceKit.Text = "Standard Rate"
        Me.chListPriceKit.Width = 100
        '
        'ColumnHeader5a
        '
        Me.ColumnHeader5a.Text = ""
        '
        'ColumnHeader6a
        '
        Me.ColumnHeader6a.Text = "Discount (%)"
        Me.ColumnHeader6a.Width = 100
        '
        'ColumnHeader7a
        '
        Me.ColumnHeader7a.Text = ""
        '
        'chPriceKit
        '
        Me.chPriceKit.Text = "Pay Rate"
        '
        'ColumnHeader8a
        '
        Me.ColumnHeader8a.Text = "Dates"
        Me.ColumnHeader8a.Width = 400
        '
        'chBanded
        '
        Me.chBanded.Text = "Banded"
        '
        'ctxDiscount
        '
        Me.ctxDiscount.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPayRateStartDate})
        '
        'mnuPayRateStartDate
        '
        Me.mnuPayRateStartDate.Index = 0
        Me.mnuPayRateStartDate.Text = "Set Start Date"
        '
        'TabPay
        '
        Me.TabPay.Controls.Add(Me.HtmlPay)
        Me.TabPay.Controls.Add(Me.Panel7)
        Me.TabPay.Controls.Add(Me.Splitter2)
        Me.TabPay.Controls.Add(Me.LvCostHeader)
        Me.TabPay.Location = New System.Drawing.Point(4, 40)
        Me.TabPay.Name = "TabPay"
        Me.TabPay.Size = New System.Drawing.Size(926, 252)
        Me.TabPay.TabIndex = 7
        Me.TabPay.Text = "Coll Pay Information"
        Me.TabPay.UseVisualStyleBackColor = True
        '
        'HtmlPay
        '
        Me.HtmlPay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlPay.Location = New System.Drawing.Point(0, 163)
        Me.HtmlPay.Name = "HtmlPay"
        Me.HtmlPay.Size = New System.Drawing.Size(926, 63)
        Me.HtmlPay.TabIndex = 13
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.cmdFind)
        Me.Panel7.Controls.Add(Me.cbPayType)
        Me.Panel7.Controls.Add(Me.cmdAddJob)
        Me.Panel7.Controls.Add(Me.cmdAddOCR)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel7.Location = New System.Drawing.Point(0, 226)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(926, 26)
        Me.Panel7.TabIndex = 12
        '
        'cmdFind
        '
        Me.cmdFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdFind.Image = CType(resources.GetObject("cmdFind.Image"), System.Drawing.Image)
        Me.cmdFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFind.Location = New System.Drawing.Point(407, 1)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.Size = New System.Drawing.Size(112, 24)
        Me.cmdFind.TabIndex = 13
        Me.cmdFind.Text = "Find Job"
        Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbPayType
        '
        Me.cbPayType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbPayType.FormattingEnabled = True
        Me.cbPayType.Location = New System.Drawing.Point(252, 2)
        Me.cbPayType.Name = "cbPayType"
        Me.cbPayType.Size = New System.Drawing.Size(151, 21)
        Me.cbPayType.Sorted = True
        Me.cbPayType.TabIndex = 12
        '
        'cmdAddJob
        '
        Me.cmdAddJob.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAddJob.Image = CType(resources.GetObject("cmdAddJob.Image"), System.Drawing.Image)
        Me.cmdAddJob.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddJob.Location = New System.Drawing.Point(134, 2)
        Me.cmdAddJob.Name = "cmdAddJob"
        Me.cmdAddJob.Size = New System.Drawing.Size(112, 24)
        Me.cmdAddJob.TabIndex = 8
        Me.cmdAddJob.Text = "Job"
        Me.cmdAddJob.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAddOCR
        '
        Me.cmdAddOCR.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAddOCR.Image = CType(resources.GetObject("cmdAddOCR.Image"), System.Drawing.Image)
        Me.cmdAddOCR.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddOCR.Location = New System.Drawing.Point(3, 2)
        Me.cmdAddOCR.Name = "cmdAddOCR"
        Me.cmdAddOCR.Size = New System.Drawing.Size(112, 24)
        Me.cmdAddOCR.TabIndex = 7
        Me.cmdAddOCR.Text = "OCR"
        Me.cmdAddOCR.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter2.Location = New System.Drawing.Point(0, 160)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(926, 3)
        Me.Splitter2.TabIndex = 9
        Me.Splitter2.TabStop = False
        '
        'LvCostHeader
        '
        Me.LvCostHeader.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader22, Me.ColumnHeader16, Me.ColumnHeader23, Me.ColumnHeader27, Me.chPay, Me.chExp, Me.chTotal, Me.ColumnHeader30, Me.ColumnHeader31, Me.ColumnHeader17})
        Me.LvCostHeader.ContextMenu = Me.ctxPayHeader
        Me.LvCostHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.LvCostHeader.FullRowSelect = True
        Me.LvCostHeader.GridLines = True
        Me.LvCostHeader.Location = New System.Drawing.Point(0, 0)
        Me.LvCostHeader.Name = "LvCostHeader"
        Me.LvCostHeader.Size = New System.Drawing.Size(926, 160)
        Me.LvCostHeader.TabIndex = 4
        Me.LvCostHeader.UseCompatibleStateImageBehavior = False
        Me.LvCostHeader.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader22
        '
        Me.ColumnHeader22.Text = "Coll Date"
        Me.ColumnHeader22.Width = 150
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "Customer"
        Me.ColumnHeader16.Width = 98
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "Vessel/Site Name"
        Me.ColumnHeader23.Width = 300
        '
        'ColumnHeader27
        '
        Me.ColumnHeader27.Text = "Samples"
        Me.ColumnHeader27.Width = 50
        '
        'chPay
        '
        Me.chPay.Text = "Costs"
        '
        'chExp
        '
        Me.chExp.Text = "Expenses"
        Me.chExp.Width = 120
        '
        'chTotal
        '
        Me.chTotal.Text = "Gross Total"
        '
        'ColumnHeader30
        '
        Me.ColumnHeader30.Text = "CM Number"
        Me.ColumnHeader30.Width = 100
        '
        'ColumnHeader31
        '
        Me.ColumnHeader31.Text = "Status"
        '
        'ColumnHeader17
        '
        Me.ColumnHeader17.Text = "Comment"
        Me.ColumnHeader17.Width = 200
        '
        'ctxPayHeader
        '
        Me.ctxPayHeader.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuApproveItem, Me.mnuDeclineItem, Me.MnuReferItem, Me.mnuCancelItem, Me.mnuEntered, Me.mnuLockPay, Me.MenuItem1, Me.mnuPayStatement})
        '
        'mnuApproveItem
        '
        Me.mnuApproveItem.Index = 0
        Me.mnuApproveItem.Text = "&Approve Item"
        '
        'mnuDeclineItem
        '
        Me.mnuDeclineItem.Index = 1
        Me.mnuDeclineItem.Text = "&Decline Item"
        '
        'MnuReferItem
        '
        Me.MnuReferItem.Index = 2
        Me.MnuReferItem.Text = "&Refer Item"
        '
        'mnuCancelItem
        '
        Me.mnuCancelItem.Index = 3
        Me.mnuCancelItem.Text = "&Cancel Item"
        '
        'mnuEntered
        '
        Me.mnuEntered.Index = 4
        Me.mnuEntered.Text = "Revert to &Entered"
        '
        'mnuLockPay
        '
        Me.mnuLockPay.Index = 5
        Me.mnuLockPay.Text = "&Lock Item"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 6
        Me.MenuItem1.Text = "-"
        '
        'mnuPayStatement
        '
        Me.mnuPayStatement.Index = 7
        Me.mnuPayStatement.Text = "Pay Statement"
        '
        'TabComments
        '
        Me.TabComments.Controls.Add(Me.txtComments)
        Me.TabComments.Controls.Add(Me.Label10)
        Me.TabComments.Controls.Add(Me.lvComments)
        Me.TabComments.Controls.Add(Me.cmdAddComment)
        Me.TabComments.Location = New System.Drawing.Point(4, 40)
        Me.TabComments.Name = "TabComments"
        Me.TabComments.Size = New System.Drawing.Size(926, 252)
        Me.TabComments.TabIndex = 6
        Me.TabComments.Text = "Comments"
        Me.TabComments.UseVisualStyleBackColor = True
        '
        'txtComments
        '
        Me.txtComments.AcceptsReturn = True
        Me.txtComments.AcceptsTab = True
        Me.txtComments.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtComments.Location = New System.Drawing.Point(16, 40)
        Me.txtComments.MaxLength = 400
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(824, 130)
        Me.txtComments.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(100, 23)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Comments"
        '
        'lvComments
        '
        Me.lvComments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvComments.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chCommDate, Me.chCommType, Me.CHComment, Me.chCommBy, Me.ChCommAction})
        Me.lvComments.Location = New System.Drawing.Point(15, 176)
        Me.lvComments.Name = "lvComments"
        Me.lvComments.Size = New System.Drawing.Size(890, 252)
        Me.lvComments.TabIndex = 0
        Me.lvComments.UseCompatibleStateImageBehavior = False
        Me.lvComments.View = System.Windows.Forms.View.Details
        '
        'chCommDate
        '
        Me.chCommDate.Text = "Date"
        Me.chCommDate.Width = 100
        '
        'chCommType
        '
        Me.chCommType.Text = "Type"
        Me.chCommType.Width = 80
        '
        'CHComment
        '
        Me.CHComment.Text = "Comment"
        Me.CHComment.Width = 200
        '
        'chCommBy
        '
        Me.chCommBy.Text = "Created By"
        Me.chCommBy.Width = 100
        '
        'ChCommAction
        '
        Me.ChCommAction.Text = "Action By"
        '
        'cmdAddComment
        '
        Me.cmdAddComment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAddComment.Image = Global.MedscreenCommonGui.AdcResource.Pharmacy
        Me.cmdAddComment.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddComment.Location = New System.Drawing.Point(843, 18)
        Me.cmdAddComment.Name = "cmdAddComment"
        Me.cmdAddComment.Size = New System.Drawing.Size(80, 23)
        Me.cmdAddComment.TabIndex = 1
        Me.cmdAddComment.Text = "Comment"
        Me.cmdAddComment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabExpenses
        '
        Me.TabExpenses.Controls.Add(Me.cbPayType2)
        Me.TabExpenses.Controls.Add(Me.cmdAddMiscPay)
        Me.TabExpenses.Controls.Add(Me.cmdAddExpense)
        Me.TabExpenses.Controls.Add(Me.lvExpenses)
        Me.TabExpenses.Controls.Add(Me.lvExpenseHeader)
        Me.TabExpenses.Controls.Add(Me.cmdAdd)
        Me.TabExpenses.Location = New System.Drawing.Point(4, 40)
        Me.TabExpenses.Name = "TabExpenses"
        Me.TabExpenses.Size = New System.Drawing.Size(926, 252)
        Me.TabExpenses.TabIndex = 8
        Me.TabExpenses.Text = "Misc/ Expenses"
        Me.TabExpenses.UseVisualStyleBackColor = True
        '
        'cbPayType2
        '
        Me.cbPayType2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbPayType2.FormattingEnabled = True
        Me.cbPayType2.Location = New System.Drawing.Point(300, 219)
        Me.cbPayType2.Name = "cbPayType2"
        Me.cbPayType2.Size = New System.Drawing.Size(151, 21)
        Me.cbPayType2.Sorted = True
        Me.cbPayType2.TabIndex = 12
        '
        'cmdAddMiscPay
        '
        Me.cmdAddMiscPay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAddMiscPay.Image = CType(resources.GetObject("cmdAddMiscPay.Image"), System.Drawing.Image)
        Me.cmdAddMiscPay.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddMiscPay.Location = New System.Drawing.Point(172, 219)
        Me.cmdAddMiscPay.Name = "cmdAddMiscPay"
        Me.cmdAddMiscPay.Size = New System.Drawing.Size(112, 24)
        Me.cmdAddMiscPay.TabIndex = 9
        Me.cmdAddMiscPay.Text = "Add Misc Pay"
        Me.cmdAddMiscPay.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAddExpense
        '
        Me.cmdAddExpense.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAddExpense.Image = CType(resources.GetObject("cmdAddExpense.Image"), System.Drawing.Image)
        Me.cmdAddExpense.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddExpense.Location = New System.Drawing.Point(21, 219)
        Me.cmdAddExpense.Name = "cmdAddExpense"
        Me.cmdAddExpense.Size = New System.Drawing.Size(112, 24)
        Me.cmdAddExpense.TabIndex = 8
        Me.cmdAddExpense.Text = "Add Expense"
        Me.cmdAddExpense.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvExpenses
        '
        Me.lvExpenses.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvExpenses.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader8, Me.ColumnHeader13, Me.ColumnHeader14, Me.chPrice, Me.chDiscount, Me.chActPrice, Me.chVALUE, Me.chStatus, Me.chComments})
        Me.lvExpenses.FullRowSelect = True
        Me.lvExpenses.GridLines = True
        Me.lvExpenses.Location = New System.Drawing.Point(0, 45)
        Me.lvExpenses.Name = "lvExpenses"
        Me.lvExpenses.Size = New System.Drawing.Size(926, 168)
        Me.lvExpenses.TabIndex = 2
        Me.lvExpenses.UseCompatibleStateImageBehavior = False
        Me.lvExpenses.View = System.Windows.Forms.View.Details
        Me.lvExpenses.Visible = False
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Expense Type"
        Me.ColumnHeader8.Width = 0
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Quantity"
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Description"
        Me.ColumnHeader14.Width = 200
        '
        'chPrice
        '
        Me.chPrice.Text = "List Price/Rate"
        Me.chPrice.Width = 118
        '
        'chDiscount
        '
        Me.chDiscount.Text = "Discount"
        '
        'chActPrice
        '
        Me.chActPrice.Text = "Act Price/Rate"
        Me.chActPrice.Width = 118
        '
        'chVALUE
        '
        Me.chVALUE.Text = "Value"
        Me.chVALUE.Width = 120
        '
        'chStatus
        '
        Me.chStatus.Text = "Status"
        '
        'chComments
        '
        Me.chComments.Text = "Comment"
        Me.chComments.Width = 200
        '
        'lvExpenseHeader
        '
        Me.lvExpenseHeader.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvExpenseHeader.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader15, Me.ColumnHeader24, Me.ColumnHeader26, Me.ColumnHeader18, Me.ColumnHeader19, Me.ColumnHeader20, Me.ColumnHeader21, Me.ColumnHeader25, Me.ColumnHeader28, Me.ColumnHeader29})
        Me.lvExpenseHeader.ContextMenu = Me.ctxPayHeader
        Me.lvExpenseHeader.FullRowSelect = True
        Me.lvExpenseHeader.GridLines = True
        Me.lvExpenseHeader.Location = New System.Drawing.Point(0, 0)
        Me.lvExpenseHeader.Name = "lvExpenseHeader"
        Me.lvExpenseHeader.Size = New System.Drawing.Size(922, 252)
        Me.lvExpenseHeader.TabIndex = 3
        Me.lvExpenseHeader.UseCompatibleStateImageBehavior = False
        Me.lvExpenseHeader.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "Pay Month"
        Me.ColumnHeader15.Width = 98
        '
        'ColumnHeader24
        '
        Me.ColumnHeader24.Text = "Customer"
        '
        'ColumnHeader26
        '
        Me.ColumnHeader26.Text = "Expense"
        Me.ColumnHeader26.Width = 200
        '
        'ColumnHeader18
        '
        Me.ColumnHeader18.Text = "Quantity"
        '
        'ColumnHeader19
        '
        Me.ColumnHeader19.Text = "Cost"
        '
        'ColumnHeader20
        '
        Me.ColumnHeader20.Text = "Expense"
        '
        'ColumnHeader21
        '
        Me.ColumnHeader21.Text = "Gross Total"
        '
        'ColumnHeader25
        '
        Me.ColumnHeader25.Text = "ID"
        '
        'ColumnHeader28
        '
        Me.ColumnHeader28.Text = "Status"
        '
        'ColumnHeader29
        '
        Me.ColumnHeader29.Text = "Comment"
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.Image = CType(resources.GetObject("cmdAdd.Image"), System.Drawing.Image)
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(16, 5)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(96, 24)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabAddress
        '
        Me.tabAddress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabAddress.Controls.Add(Me.adrPanelOfficer)
        Me.tabAddress.Controls.Add(Me.Panel4)
        Me.tabAddress.Location = New System.Drawing.Point(4, 40)
        Me.tabAddress.Name = "tabAddress"
        Me.tabAddress.Size = New System.Drawing.Size(926, 252)
        Me.tabAddress.TabIndex = 1
        Me.tabAddress.Text = "Contact Details"
        Me.tabAddress.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.txtPayMailAddress)
        Me.Panel4.Controls.Add(Me.Label26)
        Me.Panel4.Controls.Add(Me.cmdFindAddress)
        Me.Panel4.Controls.Add(Me.cmdNewAddress)
        Me.Panel4.Controls.Add(Me.lblOfficePhone)
        Me.Panel4.Controls.Add(Me.lblOfficeFax)
        Me.Panel4.Controls.Add(Me.txtOfficeFax)
        Me.Panel4.Controls.Add(Me.txtPager)
        Me.Panel4.Controls.Add(Me.lblPager)
        Me.Panel4.Controls.Add(Me.txtMobile)
        Me.Panel4.Controls.Add(Me.lblMobile)
        Me.Panel4.Controls.Add(Me.txtOfficePhone)
        Me.Panel4.Controls.Add(Me.cbSendOption)
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(922, 100)
        Me.Panel4.TabIndex = 41
        '
        'txtPayMailAddress
        '
        Me.txtPayMailAddress.Location = New System.Drawing.Point(712, 18)
        Me.txtPayMailAddress.Name = "txtPayMailAddress"
        Me.txtPayMailAddress.Size = New System.Drawing.Size(189, 20)
        Me.txtPayMailAddress.TabIndex = 237
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(603, 19)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(102, 13)
        Me.Label26.TabIndex = 236
        Me.Label26.Text = "Pay Mailing Address"
        '
        'cmdFindAddress
        '
        Me.cmdFindAddress.Image = CType(resources.GetObject("cmdFindAddress.Image"), System.Drawing.Image)
        Me.cmdFindAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindAddress.Location = New System.Drawing.Point(480, 48)
        Me.cmdFindAddress.Name = "cmdFindAddress"
        Me.cmdFindAddress.Size = New System.Drawing.Size(112, 23)
        Me.cmdFindAddress.TabIndex = 235
        Me.cmdFindAddress.Text = "Find Address"
        Me.cmdFindAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdNewAddress
        '
        Me.cmdNewAddress.Image = CType(resources.GetObject("cmdNewAddress.Image"), System.Drawing.Image)
        Me.cmdNewAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewAddress.Location = New System.Drawing.Point(480, 16)
        Me.cmdNewAddress.Name = "cmdNewAddress"
        Me.cmdNewAddress.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewAddress.TabIndex = 234
        Me.cmdNewAddress.Text = "New Address"
        Me.cmdNewAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOfficePhone
        '
        Me.lblOfficePhone.Location = New System.Drawing.Point(16, 16)
        Me.lblOfficePhone.Name = "lblOfficePhone"
        Me.lblOfficePhone.Size = New System.Drawing.Size(80, 23)
        Me.lblOfficePhone.TabIndex = 42
        Me.lblOfficePhone.Text = "Office  Phone"
        Me.lblOfficePhone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblOfficePhone.Visible = False
        '
        'lblOfficeFax
        '
        Me.lblOfficeFax.Location = New System.Drawing.Point(264, 16)
        Me.lblOfficeFax.Name = "lblOfficeFax"
        Me.lblOfficeFax.Size = New System.Drawing.Size(72, 23)
        Me.lblOfficeFax.TabIndex = 41
        Me.lblOfficeFax.Text = "Office Fax"
        Me.lblOfficeFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblOfficeFax.Visible = False
        '
        'txtOfficeFax
        '
        Me.txtOfficeFax.Location = New System.Drawing.Point(344, 16)
        Me.txtOfficeFax.Name = "txtOfficeFax"
        Me.txtOfficeFax.Size = New System.Drawing.Size(100, 20)
        Me.txtOfficeFax.TabIndex = 40
        Me.txtOfficeFax.Visible = False
        '
        'txtPager
        '
        Me.txtPager.Location = New System.Drawing.Point(344, 19)
        Me.txtPager.Name = "txtPager"
        Me.txtPager.Size = New System.Drawing.Size(100, 20)
        Me.txtPager.TabIndex = 39
        '
        'lblPager
        '
        Me.lblPager.Location = New System.Drawing.Point(256, 19)
        Me.lblPager.Name = "lblPager"
        Me.lblPager.Size = New System.Drawing.Size(72, 23)
        Me.lblPager.TabIndex = 38
        Me.lblPager.Text = "Pager"
        Me.lblPager.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMobile
        '
        Me.txtMobile.Location = New System.Drawing.Point(112, 19)
        Me.txtMobile.Name = "txtMobile"
        Me.txtMobile.Size = New System.Drawing.Size(100, 20)
        Me.txtMobile.TabIndex = 37
        '
        'lblMobile
        '
        Me.lblMobile.Location = New System.Drawing.Point(24, 19)
        Me.lblMobile.Name = "lblMobile"
        Me.lblMobile.Size = New System.Drawing.Size(72, 23)
        Me.lblMobile.TabIndex = 36
        Me.lblMobile.Text = "Mobile"
        Me.lblMobile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOfficePhone
        '
        Me.txtOfficePhone.Location = New System.Drawing.Point(112, 16)
        Me.txtOfficePhone.Name = "txtOfficePhone"
        Me.txtOfficePhone.Size = New System.Drawing.Size(100, 20)
        Me.txtOfficePhone.TabIndex = 35
        Me.txtOfficePhone.Visible = False
        '
        'cbSendOption
        '
        Me.ErrorProvider1.SetIconAlignment(Me.cbSendOption, System.Windows.Forms.ErrorIconAlignment.TopRight)
        Me.ErrorProvider1.SetIconPadding(Me.cbSendOption, 2)
        Me.cbSendOption.Location = New System.Drawing.Point(112, 72)
        Me.cbSendOption.Name = "cbSendOption"
        Me.cbSendOption.Size = New System.Drawing.Size(320, 21)
        Me.cbSendOption.TabIndex = 42
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 23)
        Me.Label8.TabIndex = 42
        Me.Label8.Text = "Contact Method"
        '
        'TabPort
        '
        Me.TabPort.Controls.Add(Me.Label7)
        Me.TabPort.Controls.Add(Me.Label6)
        Me.TabPort.Controls.Add(Me.lvOfficerPorts)
        Me.TabPort.Location = New System.Drawing.Point(4, 76)
        Me.TabPort.Name = "TabPort"
        Me.TabPort.Size = New System.Drawing.Size(926, 196)
        Me.TabPort.TabIndex = 5
        Me.TabPort.Text = "Ports"
        Me.TabPort.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label7.Location = New System.Drawing.Point(24, 157)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(229, 23)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Right Click  port to bring up menu"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label6.Location = New System.Drawing.Point(259, 157)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(248, 23)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Double click on port to edit details"
        '
        'lvOfficerPorts
        '
        Me.lvOfficerPorts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvOfficerPorts.CheckBoxes = True
        Me.lvOfficerPorts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader7, Me.CHPortCountry, Me.chPortCode, Me.chPortName, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.chNoColls, Me.chCostPerDonor, Me.chPLocation})
        Me.lvOfficerPorts.ContextMenu = Me.ctxPorts
        Me.lvOfficerPorts.FullRowSelect = True
        Me.lvOfficerPorts.GridLines = True
        Me.lvOfficerPorts.Location = New System.Drawing.Point(0, 0)
        Me.lvOfficerPorts.Name = "lvOfficerPorts"
        Me.lvOfficerPorts.Size = New System.Drawing.Size(926, 139)
        Me.lvOfficerPorts.TabIndex = 1
        Me.lvOfficerPorts.UseCompatibleStateImageBehavior = False
        Me.lvOfficerPorts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Port"
        Me.ColumnHeader7.Width = 130
        '
        'CHPortCountry
        '
        Me.CHPortCountry.DisplayIndex = 2
        Me.CHPortCountry.Text = "Country"
        Me.CHPortCountry.Width = 120
        '
        'chPortCode
        '
        Me.chPortCode.DisplayIndex = 3
        Me.chPortCode.Text = "UN Locode"
        '
        'chPortName
        '
        Me.chPortName.DisplayIndex = 1
        Me.chPortName.Text = "Name"
        Me.chPortName.Width = 100
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Agreed"
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Costs"
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Currency"
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "comments"
        Me.ColumnHeader12.Width = 300
        '
        'ctxPorts
        '
        Me.ctxPorts.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddPort, Me.mnuRemovePort, Me.mnuEditPort})
        '
        'mnuAddPort
        '
        Me.mnuAddPort.Index = 0
        Me.mnuAddPort.Text = "&Add Port"
        '
        'mnuRemovePort
        '
        Me.mnuRemovePort.Index = 1
        Me.mnuRemovePort.Text = "&Remove Port"
        '
        'mnuEditPort
        '
        Me.mnuEditPort.Index = 2
        Me.mnuEditPort.Text = "&Edit Port - Officer"
        '
        'TabLocation
        '
        Me.TabLocation.Controls.Add(Me.pnlTeamHTML)
        Me.TabLocation.Controls.Add(Me.Splitter1)
        Me.TabLocation.Controls.Add(Me.pnlTeam)
        Me.TabLocation.Controls.Add(Me.pnlLocation)
        Me.TabLocation.Location = New System.Drawing.Point(4, 40)
        Me.TabLocation.Name = "TabLocation"
        Me.TabLocation.Size = New System.Drawing.Size(926, 252)
        Me.TabLocation.TabIndex = 11
        Me.TabLocation.Text = "Teams / regions etc"
        Me.TabLocation.UseVisualStyleBackColor = True
        '
        'pnlTeamHTML
        '
        Me.pnlTeamHTML.Controls.Add(Me.HtmlEditor1)
        Me.pnlTeamHTML.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTeamHTML.Location = New System.Drawing.Point(0, 217)
        Me.pnlTeamHTML.Name = "pnlTeamHTML"
        Me.pnlTeamHTML.Size = New System.Drawing.Size(926, 35)
        Me.pnlTeamHTML.TabIndex = 3
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 0)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(926, 35)
        Me.HtmlEditor1.TabIndex = 8
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 214)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(926, 3)
        Me.Splitter1.TabIndex = 2
        Me.Splitter1.TabStop = False
        '
        'pnlTeam
        '
        Me.pnlTeam.Controls.Add(Me.treTeams)
        Me.pnlTeam.Controls.Add(Me.Label23)
        Me.pnlTeam.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTeam.Location = New System.Drawing.Point(0, 64)
        Me.pnlTeam.Name = "pnlTeam"
        Me.pnlTeam.Size = New System.Drawing.Size(926, 150)
        Me.pnlTeam.TabIndex = 1
        '
        'treTeams
        '
        Me.treTeams.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.treTeams.Location = New System.Drawing.Point(24, 48)
        Me.treTeams.Name = "treTeams"
        Me.treTeams.Size = New System.Drawing.Size(890, 96)
        Me.treTeams.TabIndex = 29
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(24, 16)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(128, 23)
        Me.Label23.TabIndex = 28
        Me.Label23.Text = "Teams / Regions / Ports"
        '
        'pnlLocation
        '
        Me.pnlLocation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlLocation.Controls.Add(Me.txtCoordinates)
        Me.pnlLocation.Controls.Add(Me.Label25)
        Me.pnlLocation.Controls.Add(Me.txtLongitude)
        Me.pnlLocation.Controls.Add(Me.Label24)
        Me.pnlLocation.Controls.Add(Me.txtLatitude)
        Me.pnlLocation.Controls.Add(Me.lLatitude)
        Me.pnlLocation.Controls.Add(Me.txtRegion)
        Me.pnlLocation.Controls.Add(Me.Label9)
        Me.pnlLocation.Controls.Add(Me.cbCountry)
        Me.pnlLocation.Controls.Add(Me.Label5)
        Me.pnlLocation.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLocation.Location = New System.Drawing.Point(0, 0)
        Me.pnlLocation.Name = "pnlLocation"
        Me.pnlLocation.Size = New System.Drawing.Size(926, 64)
        Me.pnlLocation.TabIndex = 0
        '
        'txtCoordinates
        '
        Me.txtCoordinates.Location = New System.Drawing.Point(488, 40)
        Me.txtCoordinates.Name = "txtCoordinates"
        Me.txtCoordinates.Size = New System.Drawing.Size(152, 20)
        Me.txtCoordinates.TabIndex = 38
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(392, 40)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(80, 23)
        Me.Label25.TabIndex = 37
        Me.Label25.Text = "Coordinates"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLongitude
        '
        Me.txtLongitude.Location = New System.Drawing.Point(688, 16)
        Me.txtLongitude.Name = "txtLongitude"
        Me.txtLongitude.Size = New System.Drawing.Size(100, 20)
        Me.txtLongitude.TabIndex = 36
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(616, 16)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(56, 23)
        Me.Label24.TabIndex = 35
        Me.Label24.Text = "Longitude"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLatitude
        '
        Me.txtLatitude.Location = New System.Drawing.Point(488, 16)
        Me.txtLatitude.Name = "txtLatitude"
        Me.txtLatitude.Size = New System.Drawing.Size(100, 20)
        Me.txtLatitude.TabIndex = 34
        '
        'lLatitude
        '
        Me.lLatitude.Location = New System.Drawing.Point(416, 16)
        Me.lLatitude.Name = "lLatitude"
        Me.lLatitude.Size = New System.Drawing.Size(56, 23)
        Me.lLatitude.TabIndex = 33
        Me.lLatitude.Text = "Latitude"
        Me.lLatitude.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRegion
        '
        Me.txtRegion.Location = New System.Drawing.Point(80, 16)
        Me.txtRegion.Name = "txtRegion"
        Me.txtRegion.Size = New System.Drawing.Size(100, 20)
        Me.txtRegion.TabIndex = 28
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(24, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 23)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Region"
        '
        'cbCountry
        '
        Me.cbCountry.Location = New System.Drawing.Point(272, 16)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(121, 21)
        Me.cbCountry.TabIndex = 26
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(192, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 23)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Country"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabBankDetails
        '
        Me.tabBankDetails.Controls.Add(Me.Panel5)
        Me.tabBankDetails.Controls.Add(Me.BankAddress)
        Me.tabBankDetails.Location = New System.Drawing.Point(4, 40)
        Me.tabBankDetails.Name = "tabBankDetails"
        Me.tabBankDetails.Size = New System.Drawing.Size(926, 252)
        Me.tabBankDetails.TabIndex = 2
        Me.tabBankDetails.Text = "Bank Details"
        Me.tabBankDetails.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.txtSortCode)
        Me.Panel5.Controls.Add(Me.lblBankSortCode)
        Me.Panel5.Controls.Add(Me.txtBankName)
        Me.Panel5.Controls.Add(Me.lblBankName)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(0, 212)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(926, 40)
        Me.Panel5.TabIndex = 0
        '
        'txtSortCode
        '
        Me.txtSortCode.Location = New System.Drawing.Point(376, 10)
        Me.txtSortCode.Name = "txtSortCode"
        Me.txtSortCode.Size = New System.Drawing.Size(100, 20)
        Me.txtSortCode.TabIndex = 3
        '
        'lblBankSortCode
        '
        Me.lblBankSortCode.Location = New System.Drawing.Point(256, 8)
        Me.lblBankSortCode.Name = "lblBankSortCode"
        Me.lblBankSortCode.Size = New System.Drawing.Size(100, 23)
        Me.lblBankSortCode.TabIndex = 2
        Me.lblBankSortCode.Text = "Sort Code"
        Me.lblBankSortCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBankName
        '
        Me.txtBankName.Location = New System.Drawing.Point(128, 10)
        Me.txtBankName.Name = "txtBankName"
        Me.txtBankName.Size = New System.Drawing.Size(100, 20)
        Me.txtBankName.TabIndex = 1
        '
        'lblBankName
        '
        Me.lblBankName.Location = New System.Drawing.Point(8, 8)
        Me.lblBankName.Name = "lblBankName"
        Me.lblBankName.Size = New System.Drawing.Size(100, 23)
        Me.lblBankName.TabIndex = 0
        Me.lblBankName.Text = "Bank Name"
        Me.lblBankName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabTraining
        '
        Me.tabTraining.Location = New System.Drawing.Point(4, 40)
        Me.tabTraining.Name = "tabTraining"
        Me.tabTraining.Size = New System.Drawing.Size(926, 252)
        Me.tabTraining.TabIndex = 3
        Me.tabTraining.Text = "Training"
        Me.tabTraining.UseVisualStyleBackColor = True
        '
        'tabHolidays
        '
        Me.tabHolidays.Controls.Add(Me.lvHolidays)
        Me.tabHolidays.Location = New System.Drawing.Point(4, 40)
        Me.tabHolidays.Name = "tabHolidays"
        Me.tabHolidays.Size = New System.Drawing.Size(926, 252)
        Me.tabHolidays.TabIndex = 4
        Me.tabHolidays.Text = "Holidays"
        Me.tabHolidays.UseVisualStyleBackColor = True
        '
        'lvHolidays
        '
        Me.lvHolidays.BackColor = System.Drawing.SystemColors.Control
        Me.lvHolidays.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.lvHolidays.ContextMenu = Me.ctxHoliday
        Me.lvHolidays.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvHolidays.LargeImageList = Me.imgLarge
        Me.lvHolidays.Location = New System.Drawing.Point(0, 0)
        Me.lvHolidays.Name = "lvHolidays"
        Me.lvHolidays.Size = New System.Drawing.Size(926, 252)
        Me.lvHolidays.SmallImageList = Me.imgSmall
        Me.lvHolidays.TabIndex = 0
        Me.lvHolidays.UseCompatibleStateImageBehavior = False
        Me.lvHolidays.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Start Date"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "End Date"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Days"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Deputy"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Comments"
        Me.ColumnHeader5.Width = 180
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Type"
        '
        'ctxHoliday
        '
        Me.ctxHoliday.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuView, Me.mnuLargeIcons, Me.mnuSmallIcons, Me.MnuList, Me.MnuDetails, Me.mnuBar1, Me.mnuRemoveHol, Me.mnuAddHol, Me.mnuEdHoliday})
        '
        'mnuView
        '
        Me.mnuView.Index = 0
        Me.mnuView.Text = "View"
        '
        'mnuLargeIcons
        '
        Me.mnuLargeIcons.Index = 1
        Me.mnuLargeIcons.Text = "&Large Icons"
        '
        'mnuSmallIcons
        '
        Me.mnuSmallIcons.Index = 2
        Me.mnuSmallIcons.Text = "&Small Icons"
        '
        'MnuList
        '
        Me.MnuList.Index = 3
        Me.MnuList.Text = "L&ist"
        '
        'MnuDetails
        '
        Me.MnuDetails.Index = 4
        Me.MnuDetails.Text = "&Details"
        '
        'mnuBar1
        '
        Me.mnuBar1.Index = 5
        Me.mnuBar1.Text = "-"
        '
        'mnuRemoveHol
        '
        Me.mnuRemoveHol.Index = 6
        Me.mnuRemoveHol.Text = "&Remove Holiday"
        '
        'mnuAddHol
        '
        Me.mnuAddHol.Index = 7
        Me.mnuAddHol.Text = "&Add Holiday"
        '
        'mnuEdHoliday
        '
        Me.mnuEdHoliday.Index = 8
        Me.mnuEdHoliday.Text = "&Edit Holiday"
        '
        'imgLarge
        '
        Me.imgLarge.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgLarge.ImageSize = New System.Drawing.Size(32, 32)
        Me.imgLarge.TransparentColor = System.Drawing.Color.Transparent
        '
        'imgSmall
        '
        Me.imgSmall.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgSmall.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgSmall.TransparentColor = System.Drawing.Color.Transparent
        '
        'TabEquipment
        '
        Me.TabEquipment.Controls.Add(Me.cmdaddInstrument)
        Me.TabEquipment.Controls.Add(Me.lvEquipment)
        Me.TabEquipment.Location = New System.Drawing.Point(4, 40)
        Me.TabEquipment.Name = "TabEquipment"
        Me.TabEquipment.Size = New System.Drawing.Size(926, 252)
        Me.TabEquipment.TabIndex = 9
        Me.TabEquipment.Text = "Officer Equipment"
        Me.TabEquipment.UseVisualStyleBackColor = True
        '
        'cmdaddInstrument
        '
        Me.ErrorProvider1.SetIconAlignment(Me.cmdaddInstrument, System.Windows.Forms.ErrorIconAlignment.BottomRight)
        Me.cmdaddInstrument.Image = CType(resources.GetObject("cmdaddInstrument.Image"), System.Drawing.Image)
        Me.cmdaddInstrument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdaddInstrument.Location = New System.Drawing.Point(656, 32)
        Me.cmdaddInstrument.Name = "cmdaddInstrument"
        Me.cmdaddInstrument.Size = New System.Drawing.Size(75, 23)
        Me.cmdaddInstrument.TabIndex = 1
        Me.cmdaddInstrument.Text = "Add"
        Me.cmdaddInstrument.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvEquipment
        '
        Me.lvEquipment.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CHModel, Me.CHSerialNo, Me.CHLastService, Me.CHLastCalib, Me.chRemoved})
        Me.lvEquipment.Location = New System.Drawing.Point(8, 16)
        Me.lvEquipment.Name = "lvEquipment"
        Me.lvEquipment.Size = New System.Drawing.Size(632, 360)
        Me.lvEquipment.TabIndex = 0
        Me.lvEquipment.UseCompatibleStateImageBehavior = False
        Me.lvEquipment.View = System.Windows.Forms.View.Details
        '
        'CHModel
        '
        Me.CHModel.Text = "Model"
        '
        'CHSerialNo
        '
        Me.CHSerialNo.Text = "Serial No"
        Me.CHSerialNo.Width = 192
        '
        'CHLastService
        '
        Me.CHLastService.Text = "Last Service"
        Me.CHLastService.Width = 144
        '
        'CHLastCalib
        '
        Me.CHLastCalib.Text = "Last Calibration "
        Me.CHLastCalib.Width = 136
        '
        'chRemoved
        '
        Me.chRemoved.Text = "Removed"
        '
        'TabOrders
        '
        Me.TabOrders.Controls.Add(Me.lblInvoiceDisp)
        Me.TabOrders.Controls.Add(Me.Label22)
        Me.TabOrders.Controls.Add(Me.udMaxOrders)
        Me.TabOrders.Controls.Add(Me.lvOrderItems)
        Me.TabOrders.Controls.Add(Me.lvOrders)
        Me.TabOrders.Location = New System.Drawing.Point(4, 40)
        Me.TabOrders.Name = "TabOrders"
        Me.TabOrders.Size = New System.Drawing.Size(926, 252)
        Me.TabOrders.TabIndex = 10
        Me.TabOrders.Tag = "Orders"
        Me.TabOrders.Text = "Orders"
        Me.TabOrders.UseVisualStyleBackColor = True
        '
        'lblInvoiceDisp
        '
        Me.lblInvoiceDisp.Location = New System.Drawing.Point(24, 208)
        Me.lblInvoiceDisp.Name = "lblInvoiceDisp"
        Me.lblInvoiceDisp.Size = New System.Drawing.Size(416, 23)
        Me.lblInvoiceDisp.TabIndex = 4
        Me.lblInvoiceDisp.Text = "Items for "
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(24, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(144, 23)
        Me.Label22.TabIndex = 3
        Me.Label22.Text = "Max no of orders to display"
        '
        'udMaxOrders
        '
        Me.udMaxOrders.Location = New System.Drawing.Point(168, 0)
        Me.udMaxOrders.Minimum = New Decimal(New Integer() {6, 0, 0, 0})
        Me.udMaxOrders.Name = "udMaxOrders"
        Me.udMaxOrders.Size = New System.Drawing.Size(48, 20)
        Me.udMaxOrders.TabIndex = 2
        Me.udMaxOrders.Value = New Decimal(New Integer() {6, 0, 0, 0})
        '
        'lvOrderItems
        '
        Me.lvOrderItems.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvOrderItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chOrderQty, Me.chItem})
        Me.lvOrderItems.Location = New System.Drawing.Point(24, 232)
        Me.lvOrderItems.Name = "lvOrderItems"
        Me.lvOrderItems.Size = New System.Drawing.Size(842, 252)
        Me.lvOrderItems.TabIndex = 1
        Me.lvOrderItems.UseCompatibleStateImageBehavior = False
        Me.lvOrderItems.View = System.Windows.Forms.View.Details
        '
        'chOrderQty
        '
        Me.chOrderQty.Text = "Quantity"
        '
        'chItem
        '
        Me.chItem.Text = "Item"
        Me.chItem.Width = 400
        '
        'lvOrders
        '
        Me.lvOrders.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chInvNumber, Me.chInvoicedOn, Me.chInvoiceValue, Me.chDespatch})
        Me.lvOrders.Location = New System.Drawing.Point(24, 24)
        Me.lvOrders.Name = "lvOrders"
        Me.lvOrders.Size = New System.Drawing.Size(744, 176)
        Me.lvOrders.TabIndex = 0
        Me.lvOrders.UseCompatibleStateImageBehavior = False
        Me.lvOrders.View = System.Windows.Forms.View.Details
        '
        'chInvNumber
        '
        Me.chInvNumber.Text = "Invoice Number"
        Me.chInvNumber.Width = 100
        '
        'chInvoicedOn
        '
        Me.chInvoicedOn.Text = "Invoiced On"
        Me.chInvoicedOn.Width = 100
        '
        'chInvoiceValue
        '
        Me.chInvoiceValue.Text = "Value"
        Me.chInvoiceValue.Width = 100
        '
        'chDespatch
        '
        Me.chDespatch.Text = "Despatched On"
        Me.chDespatch.Width = 100
        '
        'TabComms
        '
        Me.TabComms.Controls.Add(Me.wbComms)
        Me.TabComms.Location = New System.Drawing.Point(4, 76)
        Me.TabComms.Name = "TabComms"
        Me.TabComms.Size = New System.Drawing.Size(926, 216)
        Me.TabComms.TabIndex = 12
        Me.TabComms.Text = "Communications"
        Me.TabComms.UseVisualStyleBackColor = True
        '
        'wbComms
        '
        Me.wbComms.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbComms.Location = New System.Drawing.Point(0, 0)
        Me.wbComms.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbComms.Name = "wbComms"
        Me.wbComms.Size = New System.Drawing.Size(926, 216)
        Me.wbComms.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.cbSubDivision)
        Me.Panel3.Controls.Add(Me.Label21)
        Me.Panel3.Controls.Add(Me.cbDivision)
        Me.Panel3.Controls.Add(Me.Label20)
        Me.Panel3.Controls.Add(Me.ckRemoved)
        Me.Panel3.Controls.Add(Me.DTExpire)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.txtBadge)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.dtDOB)
        Me.Panel3.Controls.Add(Me.lblDob)
        Me.Panel3.Controls.Add(Me.txtNINum)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.txtSurName)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.txtForeName)
        Me.Panel3.Controls.Add(Me.lblForeName)
        Me.Panel3.Controls.Add(Me.txtCollID)
        Me.Panel3.Controls.Add(Me.lblID)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(934, 120)
        Me.Panel3.TabIndex = 0
        '
        'cbSubDivision
        '
        Me.cbSubDivision.Location = New System.Drawing.Point(720, 40)
        Me.cbSubDivision.Name = "cbSubDivision"
        Me.cbSubDivision.Size = New System.Drawing.Size(121, 21)
        Me.cbSubDivision.TabIndex = 28
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(664, 40)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(48, 23)
        Me.Label21.TabIndex = 27
        Me.Label21.Text = "Sub Division"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbDivision
        '
        Me.cbDivision.Location = New System.Drawing.Point(720, 8)
        Me.cbDivision.Name = "cbDivision"
        Me.cbDivision.Size = New System.Drawing.Size(121, 21)
        Me.cbDivision.TabIndex = 26
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(664, 8)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(48, 23)
        Me.Label20.TabIndex = 25
        Me.Label20.Text = "Division"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckRemoved
        '
        Me.ckRemoved.Location = New System.Drawing.Point(24, 64)
        Me.ckRemoved.Name = "ckRemoved"
        Me.ckRemoved.Size = New System.Drawing.Size(104, 24)
        Me.ckRemoved.TabIndex = 20
        Me.ckRemoved.Text = "Removed"
        '
        'DTExpire
        '
        Me.DTExpire.CustomFormat = "dd-MMM-yyyy"
        Me.DTExpire.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTExpire.Location = New System.Drawing.Point(552, 64)
        Me.DTExpire.Name = "DTExpire"
        Me.DTExpire.Size = New System.Drawing.Size(104, 20)
        Me.DTExpire.TabIndex = 19
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(488, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 23)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Expires"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBadge
        '
        Me.txtBadge.Location = New System.Drawing.Point(552, 36)
        Me.txtBadge.Name = "txtBadge"
        Me.txtBadge.Size = New System.Drawing.Size(100, 20)
        Me.txtBadge.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(504, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 23)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Badge"
        '
        'dtDOB
        '
        Me.dtDOB.CustomFormat = "dd-MMM-yyyy"
        Me.dtDOB.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtDOB.Location = New System.Drawing.Point(552, 8)
        Me.dtDOB.Name = "dtDOB"
        Me.dtDOB.Size = New System.Drawing.Size(104, 20)
        Me.dtDOB.TabIndex = 15
        '
        'lblDob
        '
        Me.lblDob.Location = New System.Drawing.Point(496, 8)
        Me.lblDob.Name = "lblDob"
        Me.lblDob.Size = New System.Drawing.Size(40, 23)
        Me.lblDob.TabIndex = 14
        Me.lblDob.Text = "DOB"
        Me.lblDob.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtNINum
        '
        Me.txtNINum.Location = New System.Drawing.Point(304, 8)
        Me.txtNINum.Name = "txtNINum"
        Me.txtNINum.Size = New System.Drawing.Size(100, 20)
        Me.txtNINum.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(224, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "NI Number"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSurName
        '
        Me.txtSurName.Location = New System.Drawing.Point(304, 36)
        Me.txtSurName.Name = "txtSurName"
        Me.txtSurName.Size = New System.Drawing.Size(176, 20)
        Me.txtSurName.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(224, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Surname"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtForeName
        '
        Me.txtForeName.Location = New System.Drawing.Point(88, 36)
        Me.txtForeName.Name = "txtForeName"
        Me.txtForeName.Size = New System.Drawing.Size(100, 20)
        Me.txtForeName.TabIndex = 9
        '
        'lblForeName
        '
        Me.lblForeName.Location = New System.Drawing.Point(16, 36)
        Me.lblForeName.Name = "lblForeName"
        Me.lblForeName.Size = New System.Drawing.Size(56, 23)
        Me.lblForeName.TabIndex = 8
        Me.lblForeName.Text = "Forename"
        Me.lblForeName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCollID
        '
        Me.txtCollID.Enabled = False
        Me.txtCollID.Location = New System.Drawing.Point(88, 8)
        Me.txtCollID.Name = "txtCollID"
        Me.txtCollID.Size = New System.Drawing.Size(100, 20)
        Me.txtCollID.TabIndex = 7
        '
        'lblID
        '
        Me.lblID.Location = New System.Drawing.Point(16, 8)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(56, 23)
        Me.lblID.TabIndex = 6
        Me.lblID.Text = "Identity"
        Me.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ctxPayEdit
        '
        Me.ctxPayEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditPay})
        '
        'mnuEditPay
        '
        Me.mnuEditPay.Index = 0
        Me.mnuEditPay.Text = "&Edit Item"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.AlwaysBlink
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        Me.ErrorProvider1.Icon = CType(resources.GetObject("ErrorProvider1.Icon"), System.Drawing.Icon)
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuReports})
        '
        'mnuReports
        '
        Me.mnuReports.Index = 0
        Me.mnuReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPaySummary, Me.mnuPaySummary2, Me.mnuOutputDefault, Me.mnuDefaultPrinter, Me.mnuCollectionByPort, Me.mnuEmailSummary})
        Me.mnuReports.Text = "&Reports"
        '
        'mnuPaySummary
        '
        Me.mnuPaySummary.Index = 0
        Me.mnuPaySummary.Text = "&Pay Overview"
        '
        'mnuPaySummary2
        '
        Me.mnuPaySummary2.Index = 1
        Me.mnuPaySummary2.Text = "Pay &Summary"
        '
        'mnuOutputDefault
        '
        Me.mnuOutputDefault.Index = 2
        Me.mnuOutputDefault.Text = "Output Default"
        '
        'mnuDefaultPrinter
        '
        Me.mnuDefaultPrinter.Index = 3
        Me.mnuDefaultPrinter.Text = "&Default Printer"
        '
        'mnuCollectionByPort
        '
        Me.mnuCollectionByPort.Index = 4
        Me.mnuCollectionByPort.Text = "&Collections by Port"
        '
        'mnuEmailSummary
        '
        Me.mnuEmailSummary.Index = 5
        Me.mnuEmailSummary.Text = "&Email Officer Summary"
        '
        'ErrorProvider2
        '
        Me.ErrorProvider2.ContainerControl = Me
        '
        'chNoColls
        '
        Me.chNoColls.Text = "No Collections"
        '
        'chCostPerDonor
        '
        Me.chCostPerDonor.Text = "Cost per Donor"
        '
        'adrPanelOfficer
        '
        Me.adrPanelOfficer.Address = Nothing
        Me.adrPanelOfficer.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.adrPanelOfficer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.adrPanelOfficer.Controls.Add(Me.lblFaxError)
        Me.adrPanelOfficer.Controls.Add(Me.lblEmailError)
        Me.adrPanelOfficer.Dock = System.Windows.Forms.DockStyle.Top
        Me.adrPanelOfficer.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.adrPanelOfficer.InfoVisible = False
        Me.adrPanelOfficer.isReadOnly = False
        Me.adrPanelOfficer.Location = New System.Drawing.Point(0, 100)
        Me.adrPanelOfficer.Name = "adrPanelOfficer"
        Me.adrPanelOfficer.ShowCancelCmd = False
        Me.adrPanelOfficer.ShowCloneCmd = False
        Me.adrPanelOfficer.ShowContact = True
        Me.adrPanelOfficer.ShowEditCmd = False
        Me.adrPanelOfficer.ShowFindCmd = False
        Me.adrPanelOfficer.ShowNewCmd = False
        Me.adrPanelOfficer.ShowPanel = False
        Me.adrPanelOfficer.ShowSaveCmd = False
        Me.adrPanelOfficer.ShowStatus = False
        Me.adrPanelOfficer.Size = New System.Drawing.Size(922, 247)
        Me.adrPanelOfficer.TabIndex = 39
        Me.adrPanelOfficer.Tooltip = Nothing
        '
        'lblFaxError
        '
        Me.lblFaxError.Location = New System.Drawing.Point(392, 72)
        Me.lblFaxError.Name = "lblFaxError"
        Me.lblFaxError.Size = New System.Drawing.Size(16, 16)
        Me.lblFaxError.TabIndex = 235
        '
        'lblEmailError
        '
        Me.lblEmailError.Location = New System.Drawing.Point(48, 184)
        Me.lblEmailError.Name = "lblEmailError"
        Me.lblEmailError.Size = New System.Drawing.Size(16, 16)
        Me.lblEmailError.TabIndex = 234
        '
        'BankAddress
        '
        Me.BankAddress.Address = Nothing
        Me.BankAddress.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.BankAddress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.BankAddress.Controls.Add(Me.Label11)
        Me.BankAddress.Controls.Add(Me.Label12)
        Me.BankAddress.Controls.Add(Me.cmdFindBankAddress)
        Me.BankAddress.Controls.Add(Me.cmdNewBankAddress)
        Me.BankAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BankAddress.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.BankAddress.InfoVisible = False
        Me.BankAddress.isReadOnly = False
        Me.BankAddress.Location = New System.Drawing.Point(0, 0)
        Me.BankAddress.Name = "BankAddress"
        Me.BankAddress.ShowCancelCmd = False
        Me.BankAddress.ShowCloneCmd = False
        Me.BankAddress.ShowContact = True
        Me.BankAddress.ShowEditCmd = False
        Me.BankAddress.ShowFindCmd = False
        Me.BankAddress.ShowNewCmd = False
        Me.BankAddress.ShowPanel = False
        Me.BankAddress.ShowSaveCmd = False
        Me.BankAddress.ShowStatus = False
        Me.BankAddress.Size = New System.Drawing.Size(926, 252)
        Me.BankAddress.TabIndex = 17
        Me.BankAddress.Tooltip = Nothing
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(392, 72)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(16, 16)
        Me.Label11.TabIndex = 235
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(48, 184)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(16, 16)
        Me.Label12.TabIndex = 234
        '
        'cmdFindBankAddress
        '
        Me.cmdFindBankAddress.Image = CType(resources.GetObject("cmdFindBankAddress.Image"), System.Drawing.Image)
        Me.cmdFindBankAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindBankAddress.Location = New System.Drawing.Point(440, 280)
        Me.cmdFindBankAddress.Name = "cmdFindBankAddress"
        Me.cmdFindBankAddress.Size = New System.Drawing.Size(112, 23)
        Me.cmdFindBankAddress.TabIndex = 233
        Me.cmdFindBankAddress.Text = "Find Address"
        Me.cmdFindBankAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdNewBankAddress
        '
        Me.cmdNewBankAddress.Image = CType(resources.GetObject("cmdNewBankAddress.Image"), System.Drawing.Image)
        Me.cmdNewBankAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewBankAddress.Location = New System.Drawing.Point(440, 248)
        Me.cmdNewBankAddress.Name = "cmdNewBankAddress"
        Me.cmdNewBankAddress.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewBankAddress.TabIndex = 232
        Me.cmdNewBankAddress.Text = "New Address"
        Me.cmdNewBankAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chPLocation
        '
        Me.chPLocation.Text = "Location"
        '
        'frmCollOfficerEd
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(938, 432)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmCollOfficerEd"
        Me.Text = "Collecting Officer Details"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.tabPaydetails.ResumeLayout(False)
        Me.tabPaydetails.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.TabPay.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.TabComments.ResumeLayout(False)
        Me.TabComments.PerformLayout()
        Me.TabExpenses.ResumeLayout(False)
        Me.tabAddress.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.TabPort.ResumeLayout(False)
        Me.TabLocation.ResumeLayout(False)
        Me.pnlTeamHTML.ResumeLayout(False)
        Me.pnlTeam.ResumeLayout(False)
        Me.pnlLocation.ResumeLayout(False)
        Me.pnlLocation.PerformLayout()
        Me.tabBankDetails.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.tabHolidays.ResumeLayout(False)
        Me.TabEquipment.ResumeLayout(False)
        Me.TabOrders.ResumeLayout(False)
        CType(Me.udMaxOrders, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabComms.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.adrPanelOfficer.ResumeLayout(False)
        Me.adrPanelOfficer.PerformLayout()
        Me.BankAddress.ResumeLayout(False)
        Me.BankAddress.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myCollOfficer As Intranet.intranet.jobs.CollectingOfficer
    Dim myAddress As MedscreenLib.Address.Caddress
    Dim myBankAddress As MedscreenLib.Address.Caddress
    Private myOffPortItem As ListViewItems.lvOfficerPortItem

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in collecting officer to edit 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CollectingOfficer() As Intranet.intranet.jobs.CollectingOfficer
        Get
            Return myCollOfficer
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CollectingOfficer)
            myCollOfficer = Value
            If Not Value Is Nothing Then
                objCommentList = New MedscreenLib.Comment.Collection(Comment.CommentOn.Officer, Me.myCollOfficer.OfficerID)
                FillForm()
            Else
                EmptyForm()
                Me.Invalidate()
            End If
        End Set
    End Property
    Private blnUserCanLockPay As Boolean
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the entire form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillForm()
        Dim objCountry As MedscreenLib.Address.Country
        Dim lvItem As ListViewItem

        Dim cp As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)

        'Check to see if the user has the rights to edit payroll number 
        If cp.IsInRole("LOGOS\ACCOUNTS") Or cp.IsInRole("LOGOS\IT") Then
            Me.txtPayrollNo.Enabled = True
        End If
        'See if we can lock pay
        blnUserCanLockPay = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_PAY_LOCK")
        'See if we can enter pay 
        Me.blnCanEnterPay = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_PAY_ENTRY")
        If Me.blnUserCanLockPay Then
            Me.mnuLockPay.Enabled = True
        Else
            Me.mnuLockPay.Enabled = False
        End If

        'If the user is not in the pay entry role lock currency and pay rate areas.

        Me.cbCurrency.Enabled = Me.blnCanEnterPay
        Me.lvPayRates.Enabled = Me.blnCanEnterPay
        Me.cbPaymethod.Enabled = Me.blnCanEnterPay
        Me.txtPayrollNo.Enabled = Me.blnCanEnterPay
        Me.lblNopayentryRole.Visible = Not Me.blnCanEnterPay
        'deal with pay entry 
        Me.cmdAdd.Enabled = Me.blnCanEnterPay
        Me.cmdAddOCR.Enabled = Me.blnCanEnterPay

        Me.lblFaxWarning.Hide()                         'Hide the message about faxing
        With myCollOfficer
            Me.txtCollID.Text = .OfficerID
            txtTTOfficerID.Text = .CardiffOfficerId
            Me.txtForeName.Text = .FirstName
            Me.txtSurName.Text = .Surname
            Me.txtComments.Text = .Comment
            txtPayMailAddress.Text = .SendInnvNum
            'Me.txtPayCuurency.Text = .PayCurrency
            Dim tp As String
            Dim int3 As Integer = 0

            Me.cbCurrency.SelectedValue = .PayCurrency
            'For Each tp In Me.UDCurrency.Items
            '    If tp = .PayCurrency Then
            '        Me.UDCurrency.SelectedIndex = int3
            '    End If
            '    int3 += 1
            'Next
            Me.cbPaymethod.SelectedValue = .PayMethod
            Me.cbDivision.SelectedValue = .Division
            Me.cbSubDivision.SelectedValue = .SubDivision
            Dim dblHourlyRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Routine_In_hours), Now)
            If dblHourlyRate = 0 Then dblHourlyRate = .HourlyRate
            Dim dblOutOfHoursRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Routine_Out_hours), Now)
            Dim dblSampleRate As Double
            If Me.CollectingOfficer.UkBased Then
                dblSampleRate = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Per_Sample_Pay), Now)
            Else
                dblSampleRate = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Per_Sample_Pay), Now)
            End If
            If dblSampleRate = 0 Then dblSampleRate = .SampleRate
            Dim dblMinCharge As Double = Me.CollectingOfficer.PayForItem(ExpenseTypeCollections.GCST_Routine_Min_Charge, Now)
            If dblMinCharge = 0 Then dblMinCharge = .MinCharge

            Me.txtHourlyRate.Text = dblHourlyRate
            Me.txtOutOfHours.Text = dblOutOfHoursRate
            Me.txtSampleRate.Text = dblSampleRate
            Me.txtMinCharge.Text = dblMinCharge

            Me.txtCOInHours.Text = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_CallOut_In_hours), Now)
            Me.txtCOOutHours.Text = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_CallOut_Out_hours), Now)


            Me.txtMobile.Text = FixUpPhone(.MobilePhone, .CountryId)
            Me.txtPager.Text = FixUpPhone(.Bleeper, .CountryId)
            Me.txtOfficePhone.Text = FixUpPhone(.OfficePhone, .CountryId)
            Me.txtOfficeFax.Text = FixUpPhone(.WorkFax, .CountryId)
            Me.txtCancSamp.Text = .CancelSampleDeduction
            Me.txtPayrollNo.Text = .PayRollNo.Trim
            Me.txtBadge.Text = .BadgeNo.Trim
            Me.ckRemoved.Checked = .Removed
            'Me.txtCountry.Text = .Country.Trim
            If .AddressId > 0 Then          'See if we have a valid address id 
                myAddress = cSupport.SMAddresses.Item(.AddressId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            Else
                myAddress = cSupport.SMAddresses.Item(0, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            End If
            If myAddress IsNot Nothing Then
                myAddress.SetAddressLink(Me.CollectingOfficer.OfficerID, 11)
            End If
            'myAddress = ModTpanel.SMAddresses.item(0, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            If .CountryId > -1 Then
                Dim objCountryv As MedscreenLib.Address.Country = _
                MedscreenLib.Glossary.Glossary.Countries.Item(.CountryId, 0)
                If Not objCountryv Is Nothing Then
                    Dim intIndex As Integer = MedscreenLib.Glossary.Glossary.Countries.IndexOf(objCountryv)
                    Me.cbCountry.SelectedIndex = intIndex
                    Me.cbCountry.SelectedValue = objCountryv.CountryId
                    If myAddress.Country <= 0 Then myAddress.Country = objCountryv.CountryId
                End If
                'Me.cbCountry.SelectedIndex = .CountryId
            End If

            Me.txtRegion.Text = .Region             'Set the collecting officers region

            'If .AddressId = 0 Then
            '    myAddress.Phone = .HomePhone
            '    myAddress.Fax = .HomeFax
            '    myAddress.AdrLine1 = .Address1
            '    myAddress.AdrLine2 = .Address2
            '    myAddress.AdrLine3 = .Address3
            '    myAddress.City = .City
            '    myAddress.District = .District
            '    myAddress.PostCode = .PostCode
            '    myAddress.AddressType = "COLL"
            '    myAddress.Email = .Email
            '    myAddress.Status = "A"
            'End If
            'If .Country.Length = 0 Then
            '    myAddress.Country = 44
            'Else
            '    'objCountry = Intranet.intranet.support.cSupport.Countries.Item(.Country)
            '    'If Not objCountry Is Nothing Then
            '    '    Me.myCollOfficer.UkBased = (objCountry.CountryId = 44)
            '    '    myAddress.Country = objCountry.CountryId
            '    'End If
            'End If

            With myAddress
                If Not myAddress Is Nothing AndAlso Not adrPanelOfficer Is Nothing Then Me.adrPanelOfficer.Address = myAddress

            End With

            Me.txtBankName.Text = .BankName
            Me.txtSortCode.Text = .BankSortCode

            If .BankAddressId > 0 Then
                myBankAddress = cSupport.SMAddresses.Item(.BankAddressId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            Else
                myBankAddress = cSupport.SMAddresses.Item(0, MedscreenLib.Address.AddressCollection.ItemType.AddressId)


                myBankAddress.AddressType = "BANK"
                'myBankAddress.AdrLine1 = .BankAddress1
                'myBankAddress.AdrLine2 = .BankAddress2
                'myBankAddress.AdrLine3 = .BankAddress3
                'myBankAddress.City = .BankCity
                'myBankAddress.District = .BankDistrict
                'myBankAddress.PostCode = .BankPostCode
                myBankAddress.Country = myAddress.Country
            End If
            If Not myBankAddress Is Nothing Then
                myBankAddress.SetAddressLink(CollectingOfficer.OfficerID, 12)
            End If

            Me.BankAddress.Address = myBankAddress
            Me.txtNINum.Text = .NINumber
            If Date.Compare(.DateOfBirth, DateSerial(1900, 1, 1)) > 0 Then
                Me.dtDOB.Value = .DateOfBirth
            Else
                Me.dtDOB.Checked = False
            End If

            If .BadgeExpire = MedscreenLib.DateField.ZeroDate Then
                Me.DTExpire.Checked = False
                Me.DTExpire.Value = Now
            Else
                Me.DTExpire.Checked = True
                Me.DTExpire.Value = .BadgeExpire
            End If

            SetVisibility(.UkBased)
            Me.cbSendOption.SelectedIndex = 0
            If .SendOption.Length = 1 Then          'Then we have an ID 
                Me.cbSendOption.SelectedValue = .SendOption()
            Else            'No we have a phrase description ugh
                Dim soPhraseInt As Integer = MedscreenLib.Glossary.Glossary.SendOption.IndexOf(.SendOption, Glossary.PhraseCollection.IndexBy.PhraseText)
                Me.cbSendOption.SelectedIndex = soPhraseInt
            End If
            'Do we warn about faxing 
            If .SendOption = "F" Or .SendOption = "P" Then
                Me.lblFaxWarning.Show()
            End If

            If .IsPinPerson Then
                Me.ckPinPerson.Checked = True
                Me.Panel6.Visible = True
                Dim dblCoordRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay), Now)
                Me.txtAreaFee.Text = dblCoordRate
                Dim dblTransRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Transfer_Pay), Now)
                Me.txtTransFee.Text = dblTransRate

            Else
                Me.ckPinPerson.Checked = False
                Me.Panel6.Visible = False
            End If
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            'For Each objPhrase In MedscreenLib.Glossary.Glossary.SendOption
            '    If objPhrase.PhraseText = .SendOption Then
            '        Exit For
            '    End If
            '    objPhrase = Nothing
            'Next
            'If Not objPhrase Is Nothing Then
            '    Dim i As Integer = MedscreenLib.Glossary.Glossary.SendOption.IndexOf(objPhrase.PhraseID)
            '    Me.cbSendOption.SelectedIndex = i
            'End If

            'FillHolidays()
            'FillPorts()
            'FillHeaderExpListView()
            'FillHeaderCostListView()
            'FillTeams()
            'Me.DrawComments()
            Me.DrawPayRates()
            'Me.DrawInstruments()
            Me.txtLatitude.Text = .Latitude
            Me.txtLongitude.Text = .Longitude
            Try
                Dim oColl As New Collection()
                oColl.Add(.Latitude)
                oColl.Add(.Longitude)
                Me.txtCoordinates.Text = CConnection.PackageStringList("Lib_officer.LatLongToCoord", oColl)
            Catch

            End Try
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Auxillary method to fill the ports aspect of the form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillPorts()
        Me.lvOfficerPorts.Items.Clear()

        Dim objCollPort As Intranet.intranet.jobs.CollPort
        With myCollOfficer
            .Ports = Nothing
            Dim lvItem As MedscreenCommonGui.ListViewItems.lvOfficerPortItem
            .Ports.SortByIsoId()
            For Each objCollPort In .Ports
               
                lvItem = New MedscreenCommonGui.ListViewItems.lvOfficerPortItem(objCollPort)
                
                Me.lvOfficerPorts.Items.Add(lvItem)

            Next
        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the tree view with teams
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/12/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillTeams()
        Me.treTeams.Nodes.Clear()
        Dim strTeams As String = CConnection.PackageStringList("Lib_Officer.OfficerTeams", Me.myCollOfficer.OfficerID)
        If strTeams Is Nothing OrElse strTeams.Trim.Length = 0 Then Exit Sub
        Dim TeamArray As String() = strTeams.Split(New Char() {"'"})
        Dim i As Integer
        For i = 0 To TeamArray.Length - 1
            Dim teamEntry As New Intranet.Teams.TeamHeader(TeamArray(i))

            Dim teamNode As New Treenodes.PeopleNodes.TeamHeader(teamEntry)
            Me.treTeams.Nodes.Add(teamNode)
            teamNode.Refresh()
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Auxillary method to fill the holiday information for form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillHolidays()

        Me.lvHolidays.Items.Clear()
        Me.lvHolidays.View = View.Details
        With myCollOfficer
            Dim lvItem As HolidayItem
            If Not .Holidays Is Nothing Then

                If .Holidays.Count > 0 Then
                    Dim objHol As Intranet.intranet.jobs.CollectorHoliday
                    For Each objHol In .Holidays
                        lvItem = New HolidayItem(objHol)
                        Me.lvHolidays.Items.Add(lvItem)
                    Next
                End If
            End If
        End With

    End Sub



    Dim objHeaders As New Intranet.intranet.jobs.CollPayCollection()

    Private Sub FillHeaderCostListView()

        Dim PayMonth As Date = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)

        'objHeaders.Load("Coll_officer = '" & Me.CollectingOfficer.OfficerID & "' and paytype in ('C','R') and paymonth = '" _
        '  & PayMonth.ToString("dd-MMM-yy") & "'")
        objHeaders.OfficerID = Me.CollectingOfficer.OfficerID
        objHeaders.Load(Me.CollectingOfficer.OfficerID, PayMonth)
        LvCostHeader.BeginUpdate()
        lvExpenseHeader.BeginUpdate()
        Dim objPayItem As Intranet.intranet.jobs.CollPay
        LvCostHeader.Items.Clear()
        lvExpenseHeader.Items.Clear()
        'Me.LvCostHeader.SelectedItems.Clear()
        '' Populate costs from collectorpay_header in list view
        Dim ExpTotal As Double = 0
        Dim payTotal As Double = 0
        Dim total As Double = 0
        Dim TotSamples As Integer = 0
        For Each objPayItem In objHeaders
            If objPayItem.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_OTHERCOSTSEXPENSES OrElse objPayItem.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_OTHERCOST _
            OrElse objPayItem.PayType = CollPay.CST_PAYTYPE_EQUIPPCCODE OrElse objPayItem.PayType = CollPay.CST_PAYTYPE_TELEPHONE OrElse objPayItem.PayType = CollPay.CST_PAYTYPE_STATIONERY _
            OrElse objPayItem.PayType = CollPay.CST_PAYTYPE_DEDUCTION Then
                Dim objListItem As New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(objPayItem, Me.CollectingOfficer.Culture)
                Me.lvExpenseHeader.Items.Add(objListItem)
                ExpTotal += objPayItem.Expenses
                payTotal += objPayItem.TotalCost
                total += objPayItem.CollTotal
            Else
                Dim objListItem As New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(objPayItem, Me.CollectingOfficer.Culture)
                Me.LvCostHeader.Items.Add(objListItem)
                ExpTotal += objListItem.Expenses
                payTotal += objListItem.PayTotal
                total += objListItem.Total
                TotSamples += objPayItem.Donors
            End If

        Next
        Dim TotalPay As New Intranet.intranet.jobs.CollPay()
        Dim OtherCosts As New Intranet.intranet.jobs.CollPay()
        Dim OtherExp As New Intranet.intranet.jobs.CollPay()
        TotalPay.ExpensesTotal = objHeaders.ExpensesTotal
        TotalPay.TotalCost = objHeaders.Total
        TotalPay.CollTotal = objHeaders.TotalPay
        TotalPay.Collection = "TOTAL"
        TotalPay.Donors = objHeaders.TotalDonors
        OtherCosts.TotalCost = objHeaders.OtherCosts
        OtherCosts.CollTotal = objHeaders.OtherCosts
        OtherCosts.Collection = "Other Costs"
        OtherExp.ExpensesTotal = objHeaders.OtherExpense
        OtherExp.CollTotal = objHeaders.OtherExpense
        OtherExp.Collection = "Other Expenses"

        Me.LvCostHeader.Items.Add(New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(TotalPay, Me.CollectingOfficer.Culture))
        Me.LvCostHeader.Items.Add(New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(OtherCosts, Me.CollectingOfficer.Culture))

        Me.LvCostHeader.Items.Add(New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(OtherExp, Me.CollectingOfficer.Culture))
        Me.lvExpenseHeader.Items.Add(New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(OtherCosts, Me.CollectingOfficer.Culture))

        Me.lvExpenseHeader.Items.Add(New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(OtherExp, Me.CollectingOfficer.Culture))
        LvCostHeader.EndUpdate()
        lvExpenseHeader.EndUpdate()


        'Me.lvPayItems.Items.Clear()
        'Dim lvItem As ListViewItems.LVCollectorpayHeader
        'If objHeaders.Count > 0 Then
        '    lvItem = Me.LvCostHeader.Items(0)
        '    lvItem.Selected = True
        '    lvItem.EnsureVisible()
        '    Me.lvCostHeader_Click(Nothing, Nothing)
        '    lvItem.Refresh()
        'End If


    End Sub

    'Private Sub FillEntryCostListView()
    '    Me.lvPayItems.Items.Clear()
    '    If Me.lvPayCurrentItem Is Nothing Then Exit Sub 'Check that we have clicked an item 
    '    If Me.lvPayCurrentItem.CollPayHeader Is Nothing Then Exit Sub ' check that entry is valid 

    '    Dim objExpPayItem As Intranet.collectorpay.CollPayItem

    '    'loop through all pay items displaying them 
    '    For Each objExpPayItem In lvPayCurrentItem.CollPayHeader.PayItems
    '        'populate the costs and associated expense items
    '        Dim obExpjListItem As New ExpenseDescription(objExpPayItem, Me.myCollOfficer)
    '        If Not objExpPayItem.Removed Then
    '            Me.lvPayItems.Items.Add(obExpjListItem)
    '        End If
    '    Next

    '    Dim objTotal As New ExpenseTotal(lvPayCurrentItem.CollPayHeader.PayItems.TotalValue, myCollOfficer.Culture)
    '    Me.lvPayItems.Items.Add(objTotal)

    'End Sub

    Private Sub FillHeaderExpListView()
        'Dim objHeaders As New Intranet.collectorpay.CollectorPayCollection()
        ''return all expense items from header table
        'Dim PayMonth As Date = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)

        'objHeaders.Load("Coll_officer = '" & Me.CollectingOfficer.OfficerID & "' and paytype = 'E' and trunc(paymonth) = '" _
        '  & PayMonth.ToString("dd-MMM-yy") & "'")

        'Dim objPayItem As Intranet.collectorpay.CollectorPay
        'Me.lvExpenseHeader.Items.Clear()
        'Me.lvExpenses.Items.Clear()
        'For Each objPayItem In objHeaders
        '    Dim objListItem As New MedscreenCommonGui.ListViewItems.LVCollectorpayHeader(objPayItem)
        '    Me.lvExpenseHeader.Items.Add(objListItem)
        'Next
        'Dim aLitem As ListViewItem = New ListViewItem("Add new Expense")
        'Me.lvExpenseHeader.Items.Add(aLitem)
        'Dim lvItem As ListViewItems.LVCollectorpayHeader
        'If objHeaders.Count > 0 Then
        '    lvItem = Me.lvExpenseHeader.Items(0)
        '    lvItem.Selected = True
        '    lvItem.EnsureVisible()
        '    Me.lvExpenseHeader_Click(Nothing, Nothing)
        '    lvItem.Refresh()
        'End If


    End Sub

    Private Sub FillEntryExpListView()
        Me.lvExpenses.Items.Clear()
        'If Me.lvCurrentItem Is Nothing Then Exit Sub 'Check that we have clicked an item 
        'If Me.lvCurrentItem.CollPayHeader Is Nothing Then Exit Sub ' check that entry is valid 

        'Dim objExpPayItem As Intranet.collectorpay.CollPayItem

        ''loop through all pay items displaying them 
        'For Each objExpPayItem In lvCurrentItem.CollPayHeader.PayItems
        '    Dim obExpjListItem As New ExpenseDescription(objExpPayItem, Me.myCollOfficer)
        '    Me.lvExpenses.Items.Add(obExpjListItem)
        'Next
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Method to set the visibility of the controls 
    ''' </summary>
    ''' <param name="blnUk"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetVisibility(ByVal blnUk As Boolean)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clear form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub EmptyForm()
        Me.txtCollID.Text = ""
        Me.txtForeName.Text = ""
        Me.txtSurName.Text = ""
        Me.cbCurrency.SelectedIndex = -1
        'Me.txtPaymethod.Text = ""
        Me.txtHourlyRate.Text = ""
        Me.txtSampleRate.Text = ""
        Me.txtMobile.Text = ""
        Me.txtPager.Text = ""
        Me.txtOfficePhone.Text = ""
        Me.txtOfficeFax.Text = ""
        Me.txtCancSamp.Text = ""
        Me.txtPayrollNo.Text = ""
        Me.txtBadge.Text = ""
        'Me.txtCountry.Text = ""
        Me.cbCountry.SelectedIndex = -1
        Me.dtDOB.Value = DateField.ZeroDate
        Me.DTExpire.Value = DateField.ZeroDate


        Me.adrPanelOfficer.Address = Nothing
        Me.txtBankName.Text = ""
        Me.txtSortCode.Text = ""

        Me.BankAddress.Address = Nothing
        Me.DTExpire.Checked = True
        'Me.DTExpire.Value = 
    End Sub


    Private Sub mnuLargeIcons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuLargeIcons.Click
        Me.lvHolidays.View = View.LargeIcon
    End Sub

    Private Sub mnuSmallIcons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSmallIcons.Click
        Me.lvHolidays.View = View.SmallIcon
    End Sub

    Private Sub MnuList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuList.Click
        Me.lvHolidays.View = View.List
    End Sub

    Private Sub MnuDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuDetails.Click
        Me.lvHolidays.View = View.Details
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit holiday info 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuEdHoliday_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEdHoliday.Click
        Dim frmhol As New frmHolidays()

        Dim lvItem As HolidayItem = Me.lvHolidays.SelectedItems.Item(0)
        Dim objHol As Intranet.intranet.jobs.CollectorHoliday
        objHol = lvItem.Holiday
        frmhol.Holiday = objHol
        If frmhol.ShowDialog Then
            objHol = frmhol.Holiday
            objHol.Update()
            lvItem.Holiday = objHol
            lvItem.Refresh()
        End If
    End Sub

    Private Sub mnuAddHol_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAddHol.Click
        Dim frmhol As New frmHolidays()

        Dim objHol As New Intranet.intranet.jobs.CollectorHoliday()

        objHol.HolidayEnd = Now.AddDays(14)
        objHol.HolidayStart = Now
        objHol.CollectorId = myCollOfficer.OfficerID
        objHol.DeputyId = ""
        objHol.Comment = ""

        frmhol.Holiday = objHol
        If frmhol.ShowDialog = DialogResult.OK Then
            objHol = frmhol.Holiday
            objHol.Insert()
            myCollOfficer.Holidays.Add(objHol)
            Me.FillHolidays()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get contents of the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetForm() As Boolean
        Dim blnReturn As Boolean = True
        With Me.myCollOfficer
            .OfficerID = Me.txtCollID.Text
            .CardiffOfficerId = txtTTOfficerID.Text
            .FirstName = Me.txtForeName.Text
            .Removed = (Me.ckRemoved.Checked)
            .Surname = Me.txtSurName.Text
            .PayCurrency = Me.cbCurrency.SelectedValue
            .PayMethod = Me.cbPaymethod.SelectedValue
            .Division = Me.cbDivision.SelectedValue
            .SubDivision = Me.cbSubDivision.SelectedValue
            .SendInnvNum = txtPayMailAddress.Text

            '.HourlyRate = Me.txtHourlyRate.Text
            '.SampleRate = Me.txtSampleRate.Text
            .MobilePhone = Me.txtMobile.Text
            .Comment = Me.txtComments.Text
            .Bleeper = Me.txtPager.Text
            .OfficePhone = Me.txtOfficePhone.Text
            .WorkFax = Me.txtOfficeFax.Text
            '.CancelSampleDeduction = Me.txtCancSamp.Text
            .PayRollNo = Me.txtPayrollNo.Text
            .BadgeNo = Me.txtBadge.Text
            .IsPinPerson = ckPinPerson.Checked

            .BankName = Me.txtBankName.Text
            .BankSortCode = Me.txtSortCode.Text
            .Region = Me.txtRegion.Text
            .Latitude = Me.txtLatitude.Text
            .Longitude = Me.txtLongitude.Text

            If Me.cbCountry.SelectedIndex > -1 Then
                Dim objCountry As MedscreenLib.Address.Country = _
                    MedscreenLib.Glossary.Glossary.Countries(Me.cbCountry.SelectedIndex)
                If Not objCountry Is Nothing Then
                    .CountryId = objCountry.CountryId
                    .Country = objCountry.CountryName
                End If
            End If

            Dim tmpDate As Object
            tmpDate = Me.dtDOB.Value

            If Not tmpDate Is Nothing Then .DateOfBirth = tmpDate
            tmpDate = Me.DTExpire.Value
            If Not tmpDate Is Nothing Then .BadgeExpire = tmpDate

            Dim myaddress As MedscreenLib.Address.Caddress

            If Me.adrPanelOfficer.Address.Fields.Changed Then           'See if address has changed if so update it
                Me.adrPanelOfficer.Address.Update()                     'Do update 
            End If

            If Me.BankAddress.Address.Fields.Changed Then           'See if address has changed if so update it
                If .AddressId = .BankAddressId Then
                    Dim tmpAddress As MedscreenLib.Address.Caddress _
                        = Intranet.intranet.Support.cSupport.SMAddresses.CreateAddress(Me.BankAddress.Address.AdrLine1, "COLL", "", 11)
                    With tmpAddress
                        .AdrLine2 = Me.BankAddress.Address.AdrLine2
                        .AdrLine3 = Me.BankAddress.Address.AdrLine3
                        .AdrLine4 = Me.BankAddress.Address.AdrLine4
                        .City = Me.BankAddress.Address.City
                        .Contact = Me.BankAddress.Address.Contact
                        .Country = Me.BankAddress.Address.Country
                        .District = Me.BankAddress.Address.District
                        .Email = Me.BankAddress.Address.Email
                        .Fax = Me.BankAddress.Address.Fax
                        .Phone = Me.BankAddress.Address.Phone
                        .Deleted = Me.BankAddress.Address.Deleted
                        .Status = Me.BankAddress.Address.Status
                        .Update()
                        Me.BankAddress.Address = tmpAddress
                    End With

                End If
                Me.BankAddress.Address.Update()   'Do update 
            End If
            .BankAddressId = Me.BankAddress.Address.AddressId

            myaddress = Me.adrPanelOfficer.Address
            If myaddress.AddressId <= 0 Then 'Create a new address 
                Dim tmpAddress As MedscreenLib.Address.Caddress _
                    = Intranet.intranet.Support.cSupport.SMAddresses.CreateAddress(myaddress.AdrLine1, "COLL", "", 11)
                If Not tmpAddress Is Nothing Then           'Address created okay 
                    myaddress.AddressId = tmpAddress.AddressId
                    myaddress.AddressType = "COLL"          'Set address type 
                    myaddress.Country = .CountryId
                    myaddress.Update()                      'Fill info 
                    .AddressId = myaddress.AddressId
                End If


            End If
            .HomePhone = myaddress.PhoneFormatted
            .HomeFax = myaddress.FaxFormatted
            '.Address1 = myaddress.AdrLine1
            '.Address2 = myaddress.AdrLine2
            '.Address3 = myaddress.AdrLine3
            '.City = myaddress.City
            '.District = myaddress.District
            '.PostCode = myaddress.PostCode
            ''myaddress.AddressType = "OFF"
            .Email = myaddress.Email
            'myaddress.Status = "A"

            'Check we are providing send information
            Me.ErrorProvider2.SetError(Me.cbSendOption, "")

            If Me.cbSendOption.SelectedIndex > 0 Then
                'Whilst still using the phrase text we have to do this 
                Dim objSOPhrase As MedscreenLib.Glossary.Phrase = Me.cbSendOption.SelectedItem
                If Not objSOPhrase Is Nothing Then
                    .SendOption = objSOPhrase.PhraseText
                End If
            Else
                .SendOption = ""
            End If
            If .SendOption.Trim.Length = 0 Then
                blnReturn = False
                Me.ErrorProvider2.SetError(Me.cbSendOption, "A method of sending information must be provided")
            End If

        End With

        Return blnReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Action the form 
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
        If GetForm() Then
            Me.DialogResult = DialogResult.OK
        Else
            Me.DialogResult = DialogResult.None
        End If
    End Sub

    Private Sub lvOfficerPorts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvOfficerPorts.DoubleClick, mnuEditPort.Click
        Dim lvitem As MedscreenCommonGui.ListViewItems.lvOfficerPortItem = Me.lvOfficerPorts.SelectedItems.Item(0)
        Dim objCollPort As Intranet.intranet.jobs.CollPort

        objCollPort = lvitem.OfficerPort
        If Not objCollPort Is Nothing Then
            Dim aFrm As frmCollPort = New frmCollPort()
            aFrm.AgreedCosts = objCollPort.Agreed
            aFrm.Cost = objCollPort.Costs
            aFrm.Currency = myCollOfficer.PayCurrency
            aFrm.Removed = objCollPort.Removed
            aFrm.Comments = objCollPort.Comments
            If aFrm.ShowDialog Then
                objCollPort.Agreed = aFrm.AgreedCosts
                objCollPort.Costs = aFrm.Cost
                objCollPort.Removed = aFrm.Removed
                objCollPort.Comments = aFrm.Comments
                objCollPort.Update()
                lvitem.OfficerPort = objCollPort
                lvitem.Refresh()
                Me.FillPorts()
            End If
        End If
        'Me.lvOfficerPorts.Items(strId).Checked = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get rid of the holiday 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuRemoveHol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemoveHol.Click

        If Me.lvHolidays.SelectedItems.Count = 0 Then Exit Sub

        Try
            Dim lvItem As HolidayItem = Me.lvHolidays.SelectedItems.Item(0)
            Dim objHol As Intranet.intranet.jobs.CollectorHoliday
            objHol = lvItem.Holiday
            If Not objHol Is Nothing Then
                objHol.Delete()
                Dim intInd As Integer = myCollOfficer.Holidays.IndexOf(objHol)
                If intInd > -1 Then myCollOfficer.Holidays.RemoveAt(intInd)
            End If
            FillHolidays()
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        End Try
    End Sub


    Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedTab.Tag = "Orders" Then
            DrawOrders()
        ElseIf TabControl1.SelectedTab.Name = "TabComms" Then
            DrawLogs()
        ElseIf TabControl1.SelectedTab.Name = TabPay.Name Then
            TabPay_Click(Nothing, Nothing)
        ElseIf TabControl1.SelectedTab.Name = TabExpenses.Name Then
            TabExpenses_Click(Nothing, Nothing)
        ElseIf TabControl1.SelectedTab.Name = TabPort.Name Then
            FillPorts()
        ElseIf TabControl1.SelectedTab.Name = TabComments.Name Then
            DrawComments()
        ElseIf TabControl1.SelectedTab.Name = tabHolidays.Name Then
            FillHolidays()
        ElseIf TabControl1.SelectedTab.Name = TabLocation.Name Then
            FillTeams()
        ElseIf TabControl1.SelectedTab.Name = TabEquipment.Name Then
            DrawInstruments()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''Draw a list of orders made by the Officer 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	21/10/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawOrders()
        Me.lvOrders.Items.Clear()
        Dim oColl As New Collection()
        If Me.myCollOfficer Is Nothing Then Exit Sub
        oColl.Add(Me.myCollOfficer.OfficerID)
        Dim intInv As Integer = Me.udMaxOrders.Value
        If intInv = 0 Then intInv = 6
        oColl.Add(intInv)
        Dim StrOrders As String = CConnection.PackageStringList("Lib_collpay.getInvoiceList", oColl)
        If Not StrOrders Is Nothing AndAlso StrOrders.Trim.Length > 0 Then
            Dim strOrderList As String() = StrOrders.Split(New Char() {","})
            Dim i As Integer
            For i = 0 To strOrderList.Length - 1
                Dim strInvoice As String = strOrderList.GetValue(i)
                If strInvoice.Trim.Length > 0 Then
                    Dim aItem As New ListViewItems.lvInvoice(strInvoice)
                    Me.lvOrders.Items.Add(aItem)
                End If
            Next
        End If
    End Sub

    Private Sub DrawLogs()
        Dim strXML As String = CollectingOfficer.OfficerLogsXML(myCollOfficer.OfficerID)
        Dim strStyleSheet As String = "OfficerLogs.xslt"
        Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strStyleSheet, 0)
        wbComms.DocumentText = strHTML
    End Sub

    Private Sub DrawOrderItems()
        Me.lvOrderItems.Items.Clear()
        If Me.currentOrder Is Nothing Then Exit Sub
        Dim myInvItem As Intranet.intranet.jobs.InvoiceItem
        For Each myInvItem In Me.currentOrder.Invoice.InvoiceItems
            Dim anItem As ListViewItems.lvInvoiceItem = New ListViewItems.lvInvoiceItem(myInvItem)
            Me.lvOrderItems.Items.Add(anItem)
        Next
    End Sub

    Private Sub tabAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabAddress.Click
        Me.adrPanelOfficer.Address = Me.myAddress
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validate phone number 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtOfficeFax_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOfficeFax.Leave
        If Me.CollectingOfficer Is Nothing Then Exit Sub
        If Me.CollectingOfficer.CountryId <= 0 Then Exit Sub
        Me.txtOfficeFax.Text = MedscreenLib.Medscreen.FixUpPhone(Me.txtOfficeFax.Text, Me.CollectingOfficer.CountryId)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validate Pager info 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtPager_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPager.Leave
        If Me.CollectingOfficer Is Nothing Then Exit Sub
        If Me.CollectingOfficer.CountryId <= 0 Then Exit Sub
        Me.txtPager.Text = MedscreenLib.Medscreen.FixUpPhone(Me.txtPager.Text, Me.CollectingOfficer.CountryId)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check mobile phone
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtMobile_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMobile.Leave
        If Me.CollectingOfficer Is Nothing Then Exit Sub
        If Me.CollectingOfficer.CountryId <= 0 Then Exit Sub
        Me.txtMobile.Text = MedscreenLib.Medscreen.FixUpPhone(Me.txtMobile.Text, Me.CollectingOfficer.CountryId)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check Office Phone
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtOfficePhone_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOfficePhone.Leave
        If Me.CollectingOfficer Is Nothing Then Exit Sub
        If Me.CollectingOfficer.CountryId <= 0 Then Exit Sub
        Me.txtOfficePhone.Text = MedscreenLib.Medscreen.FixUpPhone(Me.txtOfficePhone.Text, Me.CollectingOfficer.CountryId)

    End Sub


    Private Sub mnuAddPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddPort.Click
        Dim frmPortSel As New frmSelectPort()
        With frmPortSel
            .Listview = False
            If .ShowDialog = DialogResult.OK Then       'Want to go ahead 
                If .Port Is Nothing Then Exit Sub 'Not a valid port bog off 
                Debug.WriteLine(.Port.SiteName)
                If CConnection.ConnOpen Then            'See if we can open connection 
                    'Is this already used ?
                    Dim oCmd As New OleDb.OleDbCommand("Select COLL_ID from coll_ports where coll_id = '" & _
                        Me.myCollOfficer.OfficerID & "' and port_id = '" & .Port.PortId & "'", CConnection.DbConnection)
                    Dim intRet As String = oCmd.ExecuteScalar
                    If intRet Is Nothing Then      'Not in list 
                        Dim objCollPort As Intranet.intranet.jobs.CollPort = New Intranet.intranet.jobs.CollPort()
                        objCollPort.PortId = .Port.PortId
                        objCollPort.OfficerId = Me.myCollOfficer.OfficerID

                        Dim frmAO As New MedscreenCommonGui.frmAssignOfficer()
                        Try
                            frmAO.CollPort = objCollPort
                            If frmAO.ShowDialog = DialogResult.OK Then
                                objCollPort = frmAO.CollPort

                                objCollPort.Update()
                                .Port.CollectingOfficers.Add(objCollPort)

                                Dim lvItem As lvOfficerPortItem = New lvOfficerPortItem(objCollPort)
                                lvItem = Me.lvOfficerPorts.Items.Add(lvItem)
                                lvItem.Refresh()
                            End If
                        Catch ex As Exception
                            Medscreen.LogError(ex, , "AddOfficer - frmcollofficer")
                        End Try
                    Else
                        MsgBox("This port already exists for this officer!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub mnuRemovePort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemovePort.Click
        If myOffPortItem Is Nothing Then Exit Sub
        myOffPortItem.OfficerPort.Removed = True
        myOffPortItem.OfficerPort.Update()
        myOffPortItem.Refresh()
    End Sub

    Private Sub lvOfficerPorts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvOfficerPorts.Click
        myOffPortItem = Nothing
        If Me.lvOfficerPorts.SelectedItems.Count = 1 Then
            myOffPortItem = lvOfficerPorts.SelectedItems.Item(0)
            If Not myOffPortItem Is Nothing Then
                Me.mnuRemovePort.Visible = Not myOffPortItem.OfficerPort.Removed
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	21/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNewAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewAddress.Click
        Dim addr As MedscreenLib.Address.Caddress
        'If Me.AddressUsed = Constants.AddressType.InvoiceShipping Then
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "INSA", "", MedscreenLib.Constants.AddressType.AddressCollOfficer)
        Me.CollectingOfficer.AddressId = addr.AddressId
        MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Coll_Officer.Identity", Me.CollectingOfficer.OfficerID, MedscreenLib.Constants.AddressType.AddressCollOfficer)
        Me.adrPanelOfficer.Address = addr

    End Sub

    Private Sub cmdFindAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindAddress.Click
        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                Me.CollectingOfficer.AddressId = intAddrID
                'get address and fill control
                Dim addr As MedscreenLib.Address.Caddress
                If intAddrID > 0 Then
                    addr = MedscreenLib.Glossary.Glossary.SMAddresses.Item(intAddrID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                    Me.adrPanelOfficer.Address = addr
                    MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Coll_Officer.Identity", Me.CollectingOfficer.OfficerID, MedscreenLib.Constants.AddressType.AddressCollOfficer)
                    'TODO: Need to change address links
                End If

            End If
        End With

    End Sub

    Private Sub cmdNewBankAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewBankAddress.Click
        Dim addr As MedscreenLib.Address.Caddress
        'If Me.AddressUsed = Constants.AddressType.InvoiceShipping Then
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "BANK", "", MedscreenLib.Constants.AddressType.AddressBank)
        Me.CollectingOfficer.BankAddressId = addr.AddressId
        MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Coll_Officer.Identity", Me.CollectingOfficer.OfficerID, MedscreenLib.Constants.AddressType.AddressBank)
        Me.BankAddress.Address = addr
    End Sub

    Private Sub cmdFindBankAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindBankAddress.Click
        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                Me.CollectingOfficer.BankAddressId = intAddrID
                'get address and fill control
                Dim addr As MedscreenLib.Address.Caddress
                If intAddrID > 0 Then
                    addr = MedscreenLib.Glossary.Glossary.SMAddresses.Item(intAddrID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                    Me.BankAddress.Address = addr
                    MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Coll_Officer.Identity", Me.CollectingOfficer.OfficerID, MedscreenLib.Constants.AddressType.AddressBank)
                    'TODO: Need to change address links
                End If

            End If
        End With

    End Sub

    Private lvCurrentItem As LVOldCollectorpayHeader

    Private Sub lvExpenseHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvExpenseHeader.Click
        Me.mnuLockPay.Enabled = False
        Me.lvCurrentItem = Nothing
        If Me.lvExpenseHeader.SelectedItems.Count > 0 Then
            Dim lvItem As ListViewItem = Me.lvExpenseHeader.SelectedItems.Item(0)
            If TypeOf lvItem Is MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader Then
                Me.lvCurrentItem = lvItem
            End If
            'Else
            '    Dim lvItem As ListViewItem
            '    Dim i As Integer
            '    For i = 0 To lvExpenseHeader.Items.Count - 1
            '        lvItem = lvExpenseHeader.Items(i)
            '        If lvItem.Selected Then
            '            Me.lvCurrentItem = lvItem
            '            Exit For
            '        End If
            '    Next

        End If
        If Not Me.lvCurrentItem Is Nothing Then
            Me.SetStatusMenus(Me.lvCurrentItem.CollPayHeader.PayStatus)
            'FillEntryExpListView()
        End If

    End Sub

    Private lvPayCurrentItem As LVOldCollectorpayHeader

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set menu appearance 
    ''' </summary>
    ''' <param name="Status"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	02/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub SetStatusMenus(ByVal Status As String)
        If Me.blnUserCanLockPay Then Me.mnuLockPay.Enabled = True

        If Status = "L" Then
            Me.mnuLockPay.Enabled = False
            Me.mnuApproveItem.Enabled = False
            Me.mnuCancelItem.Enabled = False
            Me.MnuReferItem.Enabled = False
            Me.mnuDeclineItem.Enabled = False
            Me.mnuEntered.Enabled = False
        ElseIf Status = "E" Then
            Me.mnuLockPay.Enabled = False
            Me.mnuApproveItem.Enabled = True
            Me.mnuCancelItem.Enabled = True
            Me.MnuReferItem.Enabled = True
            Me.mnuDeclineItem.Enabled = True
            Me.mnuEntered.Enabled = False
        ElseIf Status = "X" Then
            Me.mnuLockPay.Enabled = True
            Me.mnuApproveItem.Enabled = False
            Me.mnuCancelItem.Enabled = False
            Me.MnuReferItem.Enabled = False
            Me.mnuDeclineItem.Enabled = False
            Me.mnuEntered.Enabled = True
        ElseIf Status = "A" Then
            Me.mnuLockPay.Enabled = True
            Me.mnuApproveItem.Enabled = False
            Me.mnuCancelItem.Enabled = True
            Me.MnuReferItem.Enabled = True
            Me.mnuDeclineItem.Enabled = True
            Me.mnuEntered.Enabled = True
        ElseIf Status = "D" Then
            Me.mnuLockPay.Enabled = True
            Me.mnuApproveItem.Enabled = True
            Me.mnuCancelItem.Enabled = True
            Me.MnuReferItem.Enabled = True
            Me.mnuDeclineItem.Enabled = False
            Me.mnuEntered.Enabled = True
        ElseIf Status = "R" Then
            Me.mnuLockPay.Enabled = False
            Me.mnuApproveItem.Enabled = True
            Me.mnuCancelItem.Enabled = True
            Me.MnuReferItem.Enabled = False
            Me.mnuDeclineItem.Enabled = True
            Me.mnuEntered.Enabled = True

        End If
    End Sub


    Private Sub lvCostHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LvCostHeader.Click
        Me.lvPayCurrentItem = Nothing
        Me.mnuLockPay.Enabled = False
        If Me.LvCostHeader.SelectedItems.Count > 0 Then
            Me.lvPayCurrentItem = Me.LvCostHeader.SelectedItems.Item(0)
            Try
                Me.HtmlPay.DocumentText = (Me.lvPayCurrentItem.CollPayHeader.ToHTML)
            Catch
            End Try
            'Check to see if it is of type OCR.
            If Not Me.lvPayCurrentItem.CollPayHeader Is Nothing AndAlso Me.lvPayCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_OCR Then
                'CheckForOCRItems(Me.lvPayCurrentItem.CollPayHeader)
            End If
        Else
            Dim lvItem As ListViewItems.LVOldCollectorpayHeader
            For Each lvItem In LvCostHeader.Items
                If lvItem.Selected Then
                    Me.lvPayCurrentItem = lvItem
                    Exit For
                End If
            Next
        End If
        If Not Me.lvPayCurrentItem Is Nothing Then
            SetStatusMenus(Me.lvPayCurrentItem.CollPayHeader.PayStatus)

            'FillEntryCostListView()
        End If
    End Sub

    'Private Sub CheckForOCRItems(ByVal PayHead As CollPay)
    '    If Me.myCollOfficer.IsPinPerson Then 'Now we need to check on the presence of the right items
    '        Dim oColl As New Collection()
    '        oColl.Add(Intranet.collectorpay.ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay)
    '        oColl.Add(Me.myCollOfficer.OfficerID)
    '        oColl.Add(PayHead.PaymentMonth.ToString("dd-MMM-yy"))
    '        Dim strRet As String = CConnection.PackageStringList("Lib_collpay.HasPayItem", oColl)
    '        'See if we have item if not create it 
    '        If Not strRet Is Nothing AndAlso strRet = "F" Then
    '            Dim payItem As CollPayItem = PayHead.PayItems.Create(GCST_Uk_Std_Coord_Pay)
    '            payItem.ItemPrice = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay), PayHead.PaymentMonth)
    '            payItem.Quantity = 1
    '            payItem.ItemValue = payItem.ItemPrice * payItem.Quantity
    '            payItem.DoUpdate()
    '        End If
    '        oColl = New Collection()
    '        oColl.Add(Intranet.collectorpay.ExpenseTypeCollections.GCST_Uk_Std_Transfer_Pay)
    '        oColl.Add(Me.myCollOfficer.OfficerID)
    '        oColl.Add(PayHead.PaymentMonth.ToString("dd-MMM-yy"))
    '        strRet = CConnection.PackageStringList("Lib_collpay.HasPayItem", oColl)
    '        If Not strRet Is Nothing AndAlso strRet = "F" Then
    '            Dim payItem As CollPayItem = PayHead.PayItems.Create(GCST_Uk_Std_Transfer_Pay)
    '            payItem.ItemPrice = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Transfer_Pay), PayHead.PaymentMonth)
    '            payItem.Quantity = 1
    '            payItem.ItemValue = payItem.ItemPrice * payItem.Quantity
    '            payItem.DoUpdate()
    '        End If

    '    End If

    'End Sub

    Private Sub lvCostHeader_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LvCostHeader.DoubleClick
        If lvPayCurrentItem Is Nothing Then Exit Sub

        Dim frmEdit As New MedscreenCommonGui.frmCollPay()
        'frmEdit.col = Me.CollectingOfficer
        Dim frmOCREdit As New frmExpEntry()

        'See what type of item we are editing 
        If lvPayCurrentItem.CollPayHeader.PayType.Trim.Length = 0 OrElse lvPayCurrentItem.CollPayHeader.PayType = "PAYHALFDAY" Then
            'frmEdit.FormType = frmCollPayEntry.enFormType.CollRelatedPay
            frmEdit.PayItem = lvPayCurrentItem.CollPayHeader
            If frmEdit.ShowDialog = DialogResult.OK Then
                'FillEntryCostListView()
            End If
        ElseIf lvPayCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_OCR Then
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.OnCallRotaEntry
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay
            If lvPayCurrentItem.CollPayHeader.ExpenseTypeID.Trim.Length = 0 Then
                lvPayCurrentItem.CollPayHeader.ExpenseTypeID = "OCRSTD"
            End If
            If lvPayCurrentItem.CollPayHeader.ItemPrice = 0 Then
                Dim dblSampleRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay), lvPayCurrentItem.CollPayHeader.PayMonth)
                lvPayCurrentItem.CollPayHeader.ItemPrice = dblSampleRate
                lvPayCurrentItem.CollPayHeader.Quantity = lvPayCurrentItem.CollPayHeader.ItemValue / dblSampleRate
            End If
            frmOCREdit.PayItem = lvPayCurrentItem.CollPayHeader
            Dim tmpAreaEntry As Intranet.intranet.jobs.CollPay = objHeaders.CoordinatorFee(Me.CollectingOfficer, True)
            frmOCREdit.ckCoordinator.Enabled = tmpAreaEntry Is Nothing
            Dim tmpTransEntry As Intranet.intranet.jobs.CollPay = objHeaders.TransferFee(Me.CollectingOfficer, True)
            frmOCREdit.ckTransfer.Enabled = (tmpTransEntry Is Nothing)
            'tmpTransEntry = lvPayCurrentItem.CollPayHeader.PayItems.Item("PAYTRANSF001", CollectorPayItemCollection.Findby.LineItem)
            frmOCREdit.AddTransferFee = Not (tmpTransEntry Is Nothing)
            'tmpAreaEntry = lvPayCurrentItem.CollPayHeader.PayItems.Item("PAYAREACO001", CollectorPayItemCollection.Findby.LineItem)
            frmOCREdit.AddCoordinatorFee = Not (tmpAreaEntry Is Nothing)
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.ExpenseTypeID = Intranet.collectorpay.ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay
                ocrEntry.DoUpdate()
                If frmOCREdit.ckCoordinator.Checked = True Then
                    Dim CoordFee As Intranet.intranet.jobs.CollPay = objHeaders.CoordinatorFee(Me.CollectingOfficer)
                    CoordFee.DoUpdate()
                End If
                If frmOCREdit.ckTransfer.Checked = True Then
                    Dim CoordFee As Intranet.intranet.jobs.CollPay = objHeaders.TransferFee(Me.CollectingOfficer)
                    CoordFee.DoUpdate()

                End If
                '        End If
                '        frmOCREdit.ckCoordinator.Enabled = True
                '        frmOCREdit.ckTransfer.Enabled = True

            End If

            'End If
            'frmEdit.FormType = frmCollPayEntry.enFormType.OCR
        ElseIf lvPayCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_AREA Or lvPayCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_TRANSFER Then
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.OnCallRotaEntry
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay

            frmOCREdit.PayItem = lvPayCurrentItem.CollPayHeader
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.DoUpdate()
            End If
        Else
        End If
        'redraw item
        lvPayCurrentItem.Refresh()
        Me.FillHeaderCostListView()
        'FillEntryCostListView()
    End Sub

    Private Sub lvExpenseHeader_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvExpenseHeader.ColumnClick

    End Sub

    Private MyCollPayEntry As Intranet.intranet.jobs.CollPay

    'Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    '    If lvCurrentItem Is Nothing Then
    '        MyCollPayEntry = Intranet.collectorpay.CollectorPayCollection.CreateHeader(CollectorPayCollection.PayTypes.OFFICER)
    '        MyCollPayEntry.PayType = "E"
    '        MyCollPayEntry.Currency = Me.myCollOfficer.PayCurrency
    '        MyCollPayEntry.DoUpdate()
    '    Else
    '        MyCollPayEntry = lvCurrentItem.CollPayHeader
    '    End If
    '    MyCollPayEntry.PaymentMonth = Me.dtPayMonth.Value
    '    Dim frmEdit As New frmExpEntry()
    '    Dim objCollPayItem As Intranet.collectorpay.CollPayItem = Me.MyCollPayEntry.GetExpenseItem

    '    frmEdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
    '    If objCollPayItem Is Nothing Then Exit Sub 'Fatal error

    '    objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '    objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '    'If Me.FormType = enFormType.CollRelatedPay Then
    '    '    frmEdit.CollOfficer = Me.objCollOfficer
    '    'Else
    '    frmEdit.CollOfficer = Me.myCollOfficer
    '    'End If
    '    frmEdit.PayItem = objCollPayItem
    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        objCollPayItem = frmEdit.PayItem
    '        objCollPayItem.SetPriceInfo(frmEdit.Quantity.Value, CDbl(frmEdit.txtPrice.Text))
    '        objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '        objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '        'If Me.FormType = enFormType.CollRelatedPay Then
    '        '    Me.lvOfficerPorts.Items.Add(New ExpenseDescription(objCollPayItem, Me.objCollOfficer))
    '        'Else
    '        '    Me.lvOfficerPorts.Items.Add(New ExpenseDescription(objCollPayItem, Me.cbCollOfficer.SelectedItem))
    '        'End If
    '    Else
    '        objCollPayItem.CancelItem() ' = True

    '    End If

    '    MyCollPayEntry.SetOfficer(Me.myCollOfficer.OfficerID, Me.dtPayMonth.Value, "E")

    '    Me.FillHeaderExpListView()

    'End Sub

    Private Sub lvExpenseHeader_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvExpenseHeader.DoubleClick
        If lvCurrentItem Is Nothing Then Exit Sub

        Dim frmEdit As New MedscreenCommonGui.frmCollPay()
        'frmEdit.col = Me.CollectingOfficer
        Dim frmOCREdit As New frmExpEntry()

        'See what type of item we are editing 
        If lvCurrentItem.CollPayHeader.PayType.Trim.Length = 0 Then
            'frmEdit.FormType = frmCollPayEntry.enFormType.CollRelatedPay
            frmEdit.PayItem = lvCurrentItem.CollPayHeader
            If frmEdit.ShowDialog = DialogResult.OK Then
                'FillEntryCostListView()
            End If
        ElseIf lvCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_OCR Then
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.OnCallRotaEntry
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay

            frmOCREdit.PayItem = lvCurrentItem.CollPayHeader
            Dim tmpAreaEntry As Intranet.intranet.jobs.CollPay = objHeaders.CoordinatorFee(Me.CollectingOfficer, True)
            frmOCREdit.ckCoordinator.Enabled = tmpAreaEntry Is Nothing
            Dim tmpTransEntry As Intranet.intranet.jobs.CollPay = objHeaders.TransferFee(Me.CollectingOfficer, True)
            frmOCREdit.ckTransfer.Enabled = (tmpTransEntry Is Nothing)
            'tmpTransEntry = lvCurrentItem.CollPayHeader.PayItems.Item("PAYTRANSF001", CollectorPayItemCollection.Findby.LineItem)
            frmOCREdit.AddTransferFee = Not (tmpTransEntry Is Nothing)
            'tmpAreaEntry = lvCurrentItem.CollPayHeader.PayItems.Item("PAYAREACO001", CollectorPayItemCollection.Findby.LineItem)
            frmOCREdit.AddCoordinatorFee = Not (tmpAreaEntry Is Nothing)
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.ExpenseTypeID = Intranet.collectorpay.ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay
                ocrEntry.PayType = CollPay.CST_PAYTYPE_OCR
                ocrEntry.DoUpdate()
                If frmOCREdit.ckCoordinator.Checked = True Then
                    Dim CoordFee As Intranet.intranet.jobs.CollPay = objHeaders.CoordinatorFee(Me.CollectingOfficer)
                    CoordFee.DoUpdate()
                End If
                If frmOCREdit.ckTransfer.Checked = True Then
                    Dim CoordFee As Intranet.intranet.jobs.CollPay = objHeaders.TransferFee(Me.CollectingOfficer)
                    CoordFee.DoUpdate()

                End If
                '        End If
                '        frmOCREdit.ckCoordinator.Enabled = True
                '        frmOCREdit.ckTransfer.Enabled = True

            End If

            'End If
            'frmEdit.FormType = frmCollPayEntry.enFormType.OCR
        ElseIf lvCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_AREA Or lvCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_TRANSFER Then
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.OnCallRotaEntry
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay

            frmOCREdit.PayItem = lvCurrentItem.CollPayHeader
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.PayType = CollPay.CST_PAYTYPE_AREA
                ocrEntry.DoUpdate()
            End If
        ElseIf lvCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_OTHERCOST Then
            frmEdit.PayItem = lvCurrentItem.CollPayHeader
            If frmEdit.ShowDialog Then

            End If
        ElseIf lvCurrentItem.CollPayHeader.PayType = CollPay.CST_PAYTYPE_OTHERCOSTSEXPENSES Then
            If lvCurrentItem.CollPayHeader.ExpenseTypeID.Trim.Length = 0 Then lvCurrentItem.CollPayHeader.ExpenseTypeID = "CEXPNOREC"
            With lvCurrentItem.CollPayHeader
                .ItemPrice = .ExpensesTotal
                .ItemValue = .ExpensesTotal
            End With
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay

            frmOCREdit.PayItem = lvCurrentItem.CollPayHeader
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.DoUpdate()
            End If

        Else
            If lvCurrentItem.CollPayHeader.ExpenseTypeID.Trim.Length = 0 Then lvCurrentItem.CollPayHeader.ExpenseTypeID = "CEXPNOREC"
            frmOCREdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
            frmOCREdit.CollOfficer = Me.CollectingOfficer
            Dim ocrEntry As CollPay

            frmOCREdit.PayItem = lvCurrentItem.CollPayHeader
            If frmOCREdit.ShowDialog = DialogResult.OK Then         'See if the user wants to save changes.
                ocrEntry = frmOCREdit.PayItem
                ocrEntry.DoUpdate()
            End If

        End If
        'redraw item
        lvCurrentItem.Refresh()
        Me.FillHeaderExpListView()


    End Sub

    Private Sub FillOCRexpFrm()

    End Sub

    Private MyCollOCREntry As Intranet.collectorpay.CollectorPay
    Private objCollOfficer As Intranet.intranet.jobs.CollectingOfficer

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' add an On Call Rota payment
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddOCR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddOCR.Click

        Dim frmOCREdit As New frmExpEntry()
        Dim strOCRRet As String

        Dim objDate As Object = Medscreen.GetParameter(MyTypes.typDate, "Pay Month", , DateSerial(Now.Year, Now.Month, 1))
        If objDate Is Nothing Then Exit Sub ' No date entered

        Dim strOCRID As String = Me.myCollOfficer.StubID & CDate(objDate).ToString("MMMyy")
        Dim oCmd As New OleDb.OleDbCommand("Select identity from coll_pay ch where  identity = '" & _
                            strOCRID & "'", CConnection.DbConnection)
        Dim objRet As Object


        'Dim intRet As String = oCmd.ExecuteScalar
        Try
            If CConnection.ConnOpen Then

                objRet = oCmd.ExecuteScalar
                If Not objRet Is Nothing Then
                    MsgBox("A On call rota payment has been entered for this month, it can be edited by double clicking on it.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)

                    Exit Sub 'We already have a pay item for this month.
                End If

            End If
        Catch ex As Exception
        Finally
            CConnection.SetConnClosed()

        End Try
        Dim objRotaPayment As CollPay

        'If frmOCREdit.DialogResult = DialogResult.OK Then

        'End If

        'Create a new rota payment header
        objRotaPayment = New CollPay()
        objRotaPayment.CID = strOCRID
        objRotaPayment.Officer = Me.CollectingOfficer.OfficerID
        'objRotaPayment.Currency = Me.CollectingOfficer.PayCurrency
        objRotaPayment.InvoiceMonth = CDate(objDate).ToString("MMMMyyyy")
        objRotaPayment.PayType = CollPay.CST_PAYTYPE_OCR
        objRotaPayment.PayStatus = "E"
        'Dim objCollPayItem As Intranet.collectorpay.CollPayItem = objRotaPayment.PayItems.Create()

        frmOCREdit.FormType = frmExpEntry.enumExpFormType.OnCallRotaEntry
        'If objCollPayItem Is Nothing Then Exit Sub 'Fatal error


        objRotaPayment.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
        objRotaPayment.ExpenseTypeID = ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay
        Dim dblSampleRate As Double = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay), Now)
        objRotaPayment.ItemPrice = dblSampleRate

        frmOCREdit.CollOfficer = Me.CollectingOfficer
        objRotaPayment.DoUpdate()
        'CheckForOCRItems(objRotaPayment)

        frmOCREdit.PayItem = objRotaPayment
        If frmOCREdit.ShowDialog = DialogResult.OK Then
            objRotaPayment = frmOCREdit.PayItem
            objRotaPayment.PayType = CollPay.CST_PAYTYPE_OCR
            'objCollPayItem = frmOCREdit.PayItem
            'objCollPayItem.SetPriceInfo(frmOCREdit.Quantity.Value, CDbl(frmOCREdit.txtPrice.Text))
            'objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
            'objCollPayItem.ExpenseTypeID = ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay
            'Me.lvOfficerPorts.Items.Add(New ExpenseDescription(objCollPayItem, Me.CollectingOfficer))
            If frmOCREdit.AddTransferFee Then
                Dim expItem As Intranet.intranet.jobs.CollPay = objHeaders.CoordinatorFee(Me.CollectingOfficer)  '"PAYAREACO001")
                'If Not expItem Is Nothing Then
                'expItem = objRotaPayment.PayItems.Create()
                expItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
                expItem.ExpenseTypeID = ExpenseTypeCollections.GCST_Uk_Std_Transfer_Pay
                expItem.Quantity = 1

                dblSampleRate = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Transfer_Pay), Now)
                expItem.ItemPrice = dblSampleRate
                expItem.ItemValue = expItem.Quantity * expItem.ItemPrice
                expItem.DoUpdate()
                'End 'If
            ElseIf frmOCREdit.AddCoordinatorFee Then
                Dim expItem As Intranet.collectorpay.CollPayItem '= objRotaPayment.PayItems.Item(ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay) '"PAYAREACO001")
                'If Not expItem Is Nothing Then
                'expItem = objRotaPayment.PayItems.Create()
                expItem.Quantity = 1
                expItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
                expItem.ExpenseTypeID = ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay
                dblSampleRate = Me.CollectingOfficer.PayForItem(LookupItemCode(ExpenseTypeCollections.GCST_Uk_Std_Coord_Pay), Now)
                expItem.ItemPrice = dblSampleRate
                expItem.ItemValue = expItem.Quantity * expItem.ItemPrice
                expItem.DoUpdate()
                ' End If
            End If

        Else
            'objCollPayItem.CancelItem() ' = True

        End If
        'objCollPayItem.ExpenseTypeID = ExpenseTypeCollections.GCST_Uk_Std_OCR_Pay
        'objCollPayItem.DoUpdate()
        'MyCollPayEntry.SetOfficer(Me.CollectingOfficer, Today, objCollPayItem.PayType)
        objRotaPayment.DoUpdate()
        Me.FillHeaderCostListView()
    End Sub


    Private Sub cmdAddJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddJob.Click
        Dim objCm As Object = Medscreen.GetParameter(MyTypes.typString, "CM Number")
        If objCm Is Nothing Then Exit Sub 'User has canceleld

        Dim strCm As String = CStr(objCm).PadLeft(10)
        strCm = strCm.Replace(" ", "0")

        Dim objRet As Object
        Dim oCmd As New OleDb.OleDbCommand()
        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "Select identity from coll_pay  where identity = ? and COLL_OFF = ?"
            oCmd.Parameters.Add(CConnection.StringParameter("CMNum", strCm, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("CollOff", Me.myCollOfficer.OfficerID, 20))
            If CConnection.ConnOpen Then            'Attempt to open reader
                objRet = oCmd.ExecuteScalar           'Get Date                
            End If
            Dim StrOfficers As String = CollPay.OtherOfficers(strCm, Me.myCollOfficer.OfficerID)
            If StrOfficers.Trim.Length > 0 Then
                MsgBox("This collection has pay statements for these Officers :- " & StrOfficers, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollOfficerEd-cmdAddJob_Click-2469")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try

        Dim objCollPayItem As CollPay
        If Not objRet Is Nothing Then
            MsgBox("A pay statement has been entered for this officer already", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            objCollPayItem = New CollPay()
            objCollPayItem.CID = objRet
            objCollPayItem.Load(objCollPayItem.CID, Me.CollectingOfficer.OfficerID)
            If objCollPayItem.PayMonth = DateField.ZeroDate Then objCollPayItem.PayMonth = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)
            objCollPayItem.InvoiceMonth = Me.dtPayMonth.Value.ToString("MMMM yyyy")
        Else
            objCollPayItem = New CollPay()
            objCollPayItem.PayMonth = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)
            objCollPayItem.InvoiceMonth = Me.dtPayMonth.Value.ToString("MMMM yyyy")


        End If
        'Now Create the item and edit it.

        objCollPayItem.PayType = ""
        objCollPayItem.Officer = Me.myCollOfficer.OfficerID
        objCollPayItem.CID = strCm
        objCollPayItem.ExpenseTypeID = cbPayType.SelectedItem.phraseid

        objCollPayItem.DoUpdate()
        Dim CollPayForm As New MedscreenCommonGui.frmCollPay()
        CollPayForm.PayItem = objCollPayItem

        If CollPayForm.ShowDialog() = DialogResult.OK Then
            CollPayForm.PayItem.DoUpdate()
        Else
            CollPayForm.PayItem.Delete()                        'Creating so delete item
        End If
        Dim objListItem As New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(objCollPayItem, Me.CollectingOfficer.Culture)
        LvCostHeader.Items.Add(objListItem)


    End Sub

    Private Sub CalcDeduction(ByVal LineItem As String, ByVal newPrice As Double)
        Dim Discount As Double
        Dim price As Double = Me.CollectingOfficer.PriceForItem(LineItem, Now)
        Dim oldPrice As Double = Me.CollectingOfficer.PayForItem(LineItem, Now)

        If price = 0 Then
            MsgBox("No price defined for " & LineItem & " - Currency - " & Me.CollectingOfficer.PayCurrency, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If price <> oldPrice Then Exit Sub
        If newPrice = oldPrice Then Exit Sub

        Discount = (price - newPrice) * 100 / price
        Dim oColl As New Collection()
        oColl.Add(LineItem)
        oColl.Add(CStr(Discount))
        Dim strDiscounts As String = CConnection.PackageStringList("PckWarehouse.FindDiscountSchemes", oColl)

        oColl = New Collection()
        oColl.Add(LineItem)
        oColl.Add(Me.CollectingOfficer.OfficerID)
        oColl.Add("Coll_Officer.Identity")
        Dim strExistingDiscount As String = CConnection.PackageStringList("PckWarehouse.GetDiscountSchemes", oColl)
        If Discount = 0 And strExistingDiscount <> "-1" Then
            Me.RemoveScheme(strExistingDiscount)
            Exit Sub
        End If


        If strDiscounts.Trim.Length = 0 Then 'no scheme create one

            Dim oCmd As New OleDb.OleDbCommand()
            Try ' Protecting 
                oCmd.Connection = CConnection.DbConnection
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.CommandText = "pckWarehouse.CreateDiscountScheme"
                If CConnection.ConnOpen Then            'Attempt to open reader
                    oCmd.Parameters.Add(CConnection.StringParameter("LineItem", LineItem, 20))
                    oCmd.Parameters.Add(CConnection.StringParameter("Discount", Discount, 20))
                    oCmd.Parameters.Add(CConnection.StringParameter("User", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 20))
                    Dim intRet As Integer = oCmd.ExecuteNonQuery
                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "frmCollOfficerEd-CalcDeduction-2615")
            Finally
                CConnection.SetConnClosed()             'Close connection
            End Try

            'now remove any existing scheme
        End If
        If strExistingDiscount <> "-1" Then
            RemoveScheme(strExistingDiscount)
        End If
        'oColl = New Collection()
        'oColl.Add(LineItem)
        'oColl.Add(CStr(Discount))
        'strDiscounts = CConnection.PackageStringList("PckWarehouse.FindDiscountSchemes", oColl)
        'Dim strDiscArray As String() = strDiscounts.Split(New Char() {","})
        'If strDiscArray.Length = 2 Then
        '    'CreateCustomerScheme(strDiscArray.GetValue(0))
        'End If
    End Sub

    Private Sub RemoveScheme(ByVal strExistingDiscount As String)
        Dim oCmd As New OleDb.OleDbCommand()
        oCmd.CommandText = "pckWarehouse.RemoveDiscountScheme"
        oCmd.Parameters.Clear()
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Connection = CConnection.DbConnection

        Try  ' Protecting
            If CConnection.ConnOpen Then            'Attempt to open reader
                oCmd.Parameters.Add(CConnection.StringParameter("SchemeId", strExistingDiscount, 20))
                oCmd.Parameters.Add(CConnection.StringParameter("CustId", Me.CollectingOfficer.OfficerID, 20))
                oCmd.Parameters.Add(CConnection.StringParameter("ClientType", "Coll_Officer.Identity", 30))
                Dim intRet As Integer = oCmd.ExecuteNonQuery
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollOfficerEd-CalcDeduction-2640")
        Finally
        End Try

    End Sub

    Private Sub CreateCustomerScheme(ByVal strExistingDiscount As String)
        Dim oCmd As New OleDb.OleDbCommand()
        oCmd.CommandText = "pckWarehouse.CreateCustomerDiscountScheme"
        oCmd.Parameters.Clear()
        oCmd.Connection = CConnection.DbConnection
        oCmd.CommandType = CommandType.StoredProcedure

        Try  ' Protecting
            If CConnection.ConnOpen Then            'Attempt to open reader
                oCmd.Parameters.Add(CConnection.StringParameter("SchemeId", strExistingDiscount, 20))
                oCmd.Parameters.Add(CConnection.StringParameter("CustId", Me.CollectingOfficer.OfficerID, 20))
                oCmd.Parameters.Add(CConnection.StringParameter("ClientType", "Coll_Officer.Identity", 30))
                oCmd.Parameters.Add(CConnection.StringParameter("User", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 20))
                Dim intRet As Integer = oCmd.ExecuteNonQuery
            End If
            mnuPayRateStartDate_Click(Nothing, Nothing)
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollOfficerEd-CalcDeduction-2640")
        Finally
        End Try

    End Sub

    Private Sub txtSampleRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSampleRate.Leave
        Dim newPrice As Double
        Try
            Dim strRate As String = Me.txtSampleRate.Text
            If Not Char.IsDigit(strRate.Chars(0)) Then
                strRate = Mid(strRate, 2)
            End If
            newPrice = CDbl(strRate)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(ExpenseTypeCollections.GCST_Sample_Pay, newPrice)
    End Sub

    Private Sub txtMinCharge_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinCharge.Leave
        Dim newPrice As Double
        Try
            Dim strRate As String = txtMinCharge.Text
            If Not Char.IsDigit(strRate.Chars(0)) Then
                strRate = Mid(strRate, 2)
            End If
            newPrice = CDbl(strRate)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(ExpenseTypeCollections.GCST_Routine_Min_Charge, newPrice)

    End Sub

    Private Sub txtHourlyRate_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHourlyRate.Leave
        Dim newPrice As Double
        Try
            Dim strRate As String = Me.txtHourlyRate.Text
            If Not Char.IsDigit(strRate.Chars(0)) Then
                strRate = Mid(strRate, 2)
            End If
            newPrice = CDbl(strRate)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(LookupItemCode(ExpenseTypeCollections.GCST_Routine_In_hours), newPrice)

    End Sub

    Private Sub txtOutOfHours_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOutOfHours.Leave
        Dim newPrice As Double
        Try
            newPrice = CDbl(Me.txtOutOfHours.Text)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(LookupItemCode(ExpenseTypeCollections.GCST_Routine_Out_hours), newPrice)

    End Sub

    Private Sub txtCOInHours_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCOInHours.Leave
        Dim newPrice As Double
        Try
            newPrice = CDbl(Me.txtCOInHours.Text)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(LookupItemCode(ExpenseTypeCollections.GCST_CallOut_In_hours), newPrice)

    End Sub

    Private Sub txtCOOutHours_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCOOutHours.Leave
        Dim newPrice As Double
        Try
            newPrice = CDbl(Me.txtCOOutHours.Text)
        Catch
            newPrice = 0
        End Try
        CalcDeduction(LookupItemCode(ExpenseTypeCollections.GCST_CallOut_Out_hours), newPrice)

    End Sub

    '''Approve pay this also stops it changing 
    'Private Sub mnuApproveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuApproveItem.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub
    '            If lvPayCurrentItem.CollPayHeader.PayStatus = "L" Then
    '                MsgBox("Locked items can't be changed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
    '                Exit Sub
    '            End If

    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Approved)
    '            lvPayCurrentItem.Refresh()
    '            Me.lvCostHeader_Click(Nothing, Nothing)
    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Approved)
    '            lvCurrentItem.Refresh()
    '            Me.lvExpenseHeader_Click(Nothing, Nothing)

    '        End If
    '    End If
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refer item back to officer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	23/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub MnuReferItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuReferItem.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub
    '            If lvPayCurrentItem.CollPayHeader.PayStatus = "L" Then
    '                MsgBox("Locked items can't be changed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
    '                Exit Sub
    '            End If
    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Referred)
    '            lvPayCurrentItem.Refresh()
    '            'Me.FillEntryCostListView()
    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Referred)
    '            lvCurrentItem.Refresh()
    '            Me.FillEntryExpListView()
    '        End If
    '    End If

    'End Sub

    'Private Sub mnuEntered_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEntered.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub
    '            If lvPayCurrentItem.CollPayHeader.PayStatus = "L" Then
    '                MsgBox("Locked items can't be changed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
    '                Exit Sub
    '            End If

    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Entered)
    '            lvPayCurrentItem.Refresh()
    '            Me.lvCostHeader_Click(Nothing, Nothing)

    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Entered)
    '            lvCurrentItem.Refresh()
    '            Me.lvExpenseHeader_Click(Nothing, Nothing)

    '        End If
    '    End If

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up the formula for the report
    ''' </summary>
    ''' <param name="objReport"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	02/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub SetUpParameters(ByVal objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal paymonth As Date)
        Dim FF As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        For Each FF In objReport.DataDefinition.FormulaFields
            If FF.FormulaName.ToUpper = "{@OFFICER}" Then
                FF.Text = """" & Me.CollectingOfficer.OfficerID & """"
            End If
            If FF.FormulaName.ToUpper = "{@PAYMONTH}" Then
                FF.Text = """" & CStr(DateSerial(paymonth.Year, paymonth.Month, 1)) & """"
            End If
        Next

    End Sub

    Private Sub mnuPaySummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaySummary.Click
        mnuPaySummary2_Click(sender, e)
        'Dim objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        'objReport = Me.LogInCrystal("\\corp.concateno.com\medscreen\common\it\crystal\collpayoff41.rpt")
        'Dim objDate As Object = Medscreen.GetParameter(MyTypes.typDate, "Pay Month", DateSerial(Now.Year, Now.Month, 1))
        'If objDate Is Nothing Then Exit Sub ' User cancelled 

        ''Set up report formula
        'SetUpParameters(objReport, CDate(objDate))
        'OutputReport(objReport, "Pay summary " & Me.CollectingOfficer.Name & " Month of " & CDate(objDate).ToString("MMM-yyyy"))
        'objReport.Close()

    End Sub

    Private Function LogInCrystal(ByVal ReportPath As String) As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim objTab As CrystalDecisions.CrystalReports.Engine.Table
        Dim Cr As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim objLogInf As CrystalDecisions.Shared.TableLogOnInfo

        Try
            Cr = New CrystalDecisions.CrystalReports.Engine.ReportDocument()

            Cr.Load(ReportPath)
            Medscreen.LogAction(" Setting log on ")
            For Each objTab In Cr.Database.Tables
                objLogInf = objTab.LogOnInfo
                objLogInf.ConnectionInfo.ServerName = CConnection.DBDatabase
                objLogInf.ConnectionInfo.Password = CConnection.DBUserName
                objLogInf.ConnectionInfo.UserID = CConnection.DBUserName
                objTab.ApplyLogOnInfo(objLogInf)
            Next
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try

        Return Cr
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pay month has changed controls need to be redrawn
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	24/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub dtPayMonth_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPayMonth.ValueChanged
        If blnDisplayPay Then Me.FillHeaderCostListView()
        'Me.FillHeaderExpListView()
    End Sub


    Private Sub LvCostHeader_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles LvCostHeader.ColumnClick
        Me.LvCostHeader.ListViewItemSorter = New Comparers.ListViewItemComparer(e.Column, 1)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check that the send option data matches 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	25/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cbSendOption_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSendOption.SelectedValueChanged
        Me.ErrorProvider1.SetError(Me.cbSendOption, "")
        Me.ErrorProvider2.SetError(Me.cbSendOption, "")
        Me.ErrorProvider1.SetError(Me.txtOfficeFax, "")
        If loading Then Exit Sub
        Dim emailError As MedscreenLib.Constants.EmailErrors
        Dim objPhrase As MedscreenLib.Glossary.Phrase
        Dim PhraseId As String
        If TypeOf Me.cbSendOption.SelectedValue Is MedscreenLib.Glossary.Phrase Then
            objPhrase = Me.cbSendOption.SelectedValue
            If objPhrase Is Nothing Then Exit Sub
            PhraseId = objPhrase.PhraseID
        ElseIf TypeOf Me.cbSendOption.SelectedValue Is String Then
            PhraseId = Me.cbSendOption.SelectedValue
        Else : Exit Sub
        End If

        Select Case PhraseId
            Case "E"
                emailError = Medscreen.ValidateEmail(Me.adrPanelOfficer.Address.Email)
                If emailError <> Constants.EmailErrors.None Then
                    Me.ErrorProvider1.SetError(Me.cbSendOption, "Error in email address")
                End If
            Case "F"
                If Me.txtOfficeFax.Text.Trim.Length > 0 Then
                    Dim strTest As String = Me.txtOfficeFax.Text.Trim
                    If strTest.Length > 1 AndAlso strTest.Chars(0) = "+" Then
                        strTest = Mid(strTest, 2)
                    End If
                    If Not Medscreen.IsNumber(strTest) Then
                        Me.ErrorProvider1.SetError(Me.cbSendOption, "Error in fax address")
                    End If
                Else
                    'If Not Medscreen.IsNumber(strTest) Then
                    Me.ErrorProvider1.SetError(Me.txtOfficeFax, "No Fax number provided")
                End If
        End Select
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find a job and show it 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
        Dim objCMNo As Object = Medscreen.GetParameter(MyTypes.typString, "CM Number", "CM Number of collection")
        If objCMNo Is Nothing Then Exit Sub ' User cancelled
        Dim strCollPayId As String = CStr(objCMNo).PadLeft(10).Replace(" ", "0")

        'Dim oColl As New Collection()
        'oColl.Add(CStr(objCMNo).PadLeft(10).Replace(" ", "0"))
        'oColl.Add(Me.myCollOfficer.OfficerID)
        'Dim strCollPayId As String = CConnection.PackageStringList("Lib_collPay.FindPayHeader", oColl)

        'See if we have collection 
        If strCollPayId Is Nothing OrElse strCollPayId.Trim.Length = 0 Then
            MsgBox("Payment for " & CStr(objCMNo) & " not found ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim objCollPayID As New CollPay()              'Create the coll pay object
        objCollPayID.CID = strCollPayId
        objCollPayID.Load(strCollPayId, Me.CollectingOfficer.OfficerID)          'Load it 

        If objCollPayID.PayMonth <= DateField.ZeroDate Then
            Dim strDate As String = "01-" & Mid(objCollPayID.InvoiceMonth, 1, 3) & "-" & Mid(objCollPayID.InvoiceMonth, objCollPayID.InvoiceMonth.Length - 3)
            Try
                objCollPayID.PayMonth = Date.Parse(strDate)
            Catch ex As Exception
                Exit Sub
            End Try

        End If
        'Need to ceheck the validity then move to the correct pay date
        Me.dtPayMonth.Value = objCollPayID.PayMonth
        'Me.FillHeaderCostListView()
        LvCostHeader.SelectedItems.Clear()
        'now need to search for item in list and select it.
        Dim LvItem As MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader
        For Each LvItem In Me.LvCostHeader.Items
            If LvItem.CollPayHeader.CID = strCollPayId Then
                LvItem.EnsureVisible()
                LvItem.Selected = True

                Me.lvCostHeader_Click(Nothing, Nothing)
                LvItem.Refresh()
                Exit For
            End If
        Next
    End Sub

    'Private Sub mnuLockPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLockPay.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub

    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Locked)
    '            lvPayCurrentItem.CollPayHeader.LockPay()
    '            lvPayCurrentItem.Refresh()
    '            Me.lvCostHeader_Click(Nothing, Nothing)
    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Locked)
    '            lvCurrentItem.CollPayHeader.LockPay()
    '            lvCurrentItem.Refresh()
    '            Me.lvExpenseHeader_Click(Nothing, Nothing)
    '        End If
    '    End If

    'End Sub

    'Private Sub mnuDeclineItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeclineItem.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub
    '            If lvPayCurrentItem.CollPayHeader.PayStatus = "L" Then
    '                MsgBox("Locked items can't be changed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
    '                Exit Sub
    '            End If

    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Declined)
    '            lvPayCurrentItem.Refresh()
    '            Me.lvCostHeader_Click(Nothing, Nothing)
    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Declined)
    '            lvCurrentItem.Refresh()
    '            Me.lvExpenseHeader_Click(Nothing, Nothing)

    '        End If
    '    End If

    'End Sub

    'Private Sub mnuCancelItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancelItem.Click
    '    Dim objMenu As MenuItem = sender
    '    Dim ctxMenu As ContextMenu = objMenu.Parent

    '    If TypeOf ctxMenu.SourceControl Is ListView Then
    '        Dim tmpListView As ListView
    '        tmpListView = ctxMenu.SourceControl
    '        If tmpListView.Name = Me.LvCostHeader.Name Then
    '            If Me.lvPayCurrentItem Is Nothing Then Exit Sub
    '            If lvPayCurrentItem.CollPayHeader.PayStatus = "L" Then
    '                MsgBox("Locked items can't be changed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
    '                Exit Sub
    '            End If

    '            lvPayCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Cancelled)
    '            lvPayCurrentItem.Refresh()
    '            Me.lvCostHeader_Click(Nothing, Nothing)
    '        ElseIf tmpListView.Name = Me.lvExpenseHeader.Name Then
    '            If lvCurrentItem Is Nothing Then Exit Sub
    '            lvCurrentItem.CollPayHeader.SetPayStatus(CollPayItem.GCST_Pay_Cancelled)
    '            lvCurrentItem.Refresh()
    '            Me.lvExpenseHeader_Click(Nothing, Nothing)


    '        End If
    '    End If

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raise a detailed pay statement for this item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	02/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub mnuPayStatement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPayStatement.Click
    '    Dim objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
    '    'This needs moving to config file
    '    objReport = Me.LogInCrystal("\\andrew\solutions\worktemp\PayDetail41.rpt")

    '    'Set up report formula

    '    Dim FF As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

    '    For Each FF In objReport.DataDefinition.FormulaFields
    '        If FF.FormulaName.ToUpper = "{@COLLPAYID}" Then
    '            FF.Text = """" & Me.lvPayCurrentItem.CollPayHeader.CollPayID & """"
    '        End If
    '    Next

    '    OutputReport(objReport, "Detailed pay statement " & Me.CollectingOfficer.Name & " for " & Me.lvPayCurrentItem.CollPayHeader.CMidentity)
    '    objReport.Close()


    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Output a report 
    ''' </summary>
    ''' <param name="objReport"></param>
    ''' <param name="Subject"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub OutputReport(ByVal objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal Subject As String)
        If Not MedscreenCommonGui.CommonForms.UserOptions Is Nothing Then
            MedscreenCommonGui.CommonForms.UserOptions.OutputCrystalReport(objReport, Subject)
        End If
    End Sub

   
   
    
    Private Sub mnuPaySummary2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaySummary2.Click

        CollectingOfficerCollection.TidyPay()
        If myCollOfficer Is Nothing Then Exit Sub
        Dim oCollr As New Collection()
        oCollr.Add(myCollOfficer.OfficerID)
        oCollr.Add("PAYPERMILE")
        oCollr.Add(Me.dtPayMonth.Value)
        Dim MileageRate As Double = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oCollr)
        objHeaders.Clear()
        objHeaders.OfficerID = Me.CollectingOfficer.OfficerID
        objHeaders.Load(Me.CollectingOfficer.OfficerID, dtPayMonth.Value)

        myCollOfficer.PaySummary(myCollOfficer.OfficerID, MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod, Me.dtPayMonth.Value, MedscreenLib.Glossary.Glossary.CurrentSMUserEmail, MileageRate, MedscreenCommonGui.CommonForms.UserOptions.DefaultPrinter, objHeaders)


    End Sub

    Private Sub mnuOutputDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOutputDefault.Click
        Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.Constants.SendMethod))
        Dim obj As Object = Medscreen.GetParameter(GetType(MedscreenLib.Constants.SendMethod), "Test")
        If Not obj Is Nothing Then
            MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod = [Enum].Parse(GetType(MedscreenLib.Constants.SendMethod), obj)
            SetOutputMenu()
        End If
    End Sub

    Private Sub mnuDefaultPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDefaultPrinter.Click
        Dim frm As New MedscreenLib.frmOutputSelect()
        frm.OutputMethod = MedscreenLib.frmOutputSelect.cstPrinterOption
        frm.PrinterOutputTo = frmOutputSelect.PrinterOutput.LogicalName
        If frm.ShowDialog = DialogResult.OK Then
            MedscreenCommonGui.CommonForms.UserOptions.DefaultPrinter = frm.OutputDestination
        End If
        SetOutputMenu()
    End Sub

    Private CurrentExpListItem As ExpenseDescription
    Private Sub lvExpenses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvExpenses.Click
        Me.CurrentExpListItem = Nothing
        If Me.lvExpenses.SelectedItems.Count > 0 Then
            Me.CurrentExpListItem = Me.lvExpenses.SelectedItems.Item(0)
        End If
    End Sub

    'Private Sub lvExpenses_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvExpenses.DoubleClick
    '    If CurrentExpListItem Is Nothing Then Exit Sub
    '    If Me.blnCanEnterPay = False Then Exit Sub
    '    Dim frmEdit As New frmExpEntry()
    '    frmEdit.CollOfficer = Me.CollectingOfficer
    '    frmEdit.PayItem = CurrentExpListItem.Expense
    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        CurrentExpListItem.Expense = frmEdit.PayItem
    '        CurrentExpListItem.Expense.DoUpdate()
    '        CurrentExpListItem.Refresh()
    '    End If
    'End Sub



    Dim objCommentList As MedscreenLib.Comment.Collection

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:57:46
    '''     Display the customer comments
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub DrawComments()
        Dim objComm As MedscreenLib.Comment
        Me.lvComments.Items.Clear()
        objCommentList.Clear()
        'objCommentList.Load()

        objCommentList.Load(Comment.CommentOn.Officer, Me.myCollOfficer.OfficerID)
        For Each objComm In objCommentList
            Me.lvComments.Items.Add(New MedscreenCommonGui.ListViewItems.lvComment(objComm))
        Next
    End Sub

    Private myCustComment As MedscreenCommonGui.ListViewItems.lvComment

    Private Sub lvComments_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvComments.SelectedIndexChanged
        Me.myCustComment = Nothing
        txtComments.Text = ""
        If Me.lvComments.SelectedItems.Count = 1 Then
            Me.myCustComment = Me.lvComments.SelectedItems.Item(0)
            If Not Me.myCustComment Is Nothing AndAlso Not Me.myCustComment.MedComment Is Nothing AndAlso (InStr("INFO,COFF,MAIL,COLL", myCustComment.MedComment.Type) > 0) Then
                Me.txtComments.Text = myCustComment.MedComment.Text
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' add a comment
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddComment.Click
        Dim afrm As New FrmCollComment()
        afrm.FormType = FrmCollComment.FCCFormTypes.CollectorComment
        afrm.CommentType = MedscreenLib.Constants.GCST_COMM_CollCOOfficerInfo
        afrm.CMNumber = Me.myCollOfficer.OfficerID
        If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Dim aComm As Intranet.intranet.jobs.CollectionComment = _
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.Officer, Me.myCollOfficer.OfficerID, afrm.CommentType, afrm.Comment, _
            MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            'If Not aComm Is Nothing Then
            '    aComm.ActionBy = afrm.ActionBy
            '    aComm.Update()
            Me.DrawComments()
            'End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw a list of instruments belonging to officer.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	15/07/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawInstruments()
        lvEquipment.Items.Clear()
        If Me.myCollOfficer Is Nothing Then Exit Sub
        If Me.myCollOfficer.Instruments Is Nothing Then Exit Sub
        Dim objOffInst As Intranet.INSTRUMENT.OfficerInstrument
        For Each objOffInst In Me.myCollOfficer.Instruments
            Dim objInst As Intranet.INSTRUMENT.Instrument = New Intranet.INSTRUMENT.Instrument(objOffInst.InstrumentId)
            Dim lvItem As lvInstrument = New lvInstrument(objInst)
            Me.lvEquipment.Items.Add(lvItem)
        Next
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw the pay rates
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	12/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawPayRates()
        Me.lvPayRates.Items.Clear()

        'clear volume discount info
        Me.txtVolumeDiscount.Text = ""

        'get line items and fill list view
        Dim oColl As New Collection()
        oColl.Add("COLLECPAY")
        oColl.Add("EXPENSE")
        Dim strLineItems As String = CConnection.PackageStringList("LIB_CCTOOL.GETLINEITEMSBYSERVICE", oColl)
        If Not strLineItems Is Nothing AndAlso strLineItems.Trim.Length > 0 Then 'Process Line Items
            Dim LineItemArray As String() = strLineItems.Split(New Char() {","})
            Dim i As Integer
            For i = 0 To LineItemArray.Length - 1
                Dim strItemCode As String = LineItemArray.GetValue(i)
                Dim objLVPayRate As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry
                objLVPayRate = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.myCollOfficer, strItemCode)
                Me.lvPayRates.Items.Add(objLVPayRate)
            Next

        End If

        'Get volume discount items list
        Dim strVDInfo As String = CConnection.PackageStringList("Lib_collpay.VolumeDiscountList", Me.myCollOfficer.OfficerID)
        If strVDInfo Is Nothing OrElse strVDInfo.Trim.Length = 0 Then Exit Sub
        Dim strVDList As String() = strVDInfo.Split(New Char() {","})
        Dim j As Integer
        Dim strOutput As String = ""
        For j = 0 To strVDList.Length - 1           'Process list
            Dim strItem As String = strVDList.GetValue(0)
            If strItem.Trim.Length > 0 Then         'If valid item get text
                oColl = New Collection()
                oColl.Add(Me.myCollOfficer.OfficerID)
                oColl.Add(strItem)

                strOutput += CConnection.PackageStringList("pckWarehouse.GetVolDiscText", oColl)
            End If
        Next

        Me.txtVolumeDiscount.Text = strOutput
    End Sub

    Private lvKitItem As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a sub item being clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	20/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub lvPayRates_SubItemClicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs) Handles lvPayRates.SubItemClicked
        Dim lvItem As ListViewItem = e.Item

        Dim lvSubItem As Integer = e.SubItem
        lvKitItem = Nothing
        'Dim textbox1 As TextBox = New TextBox()
        If TypeOf e.Item Is MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry Then
            lvKitItem = e.Item
        Else
            Exit Sub
        End If
        Dim Editors As Control() = New Control() {Me.txtKitEdit}
        If lvSubItem = 6 Then
            Dim strPrice As String = Me.myCollOfficer.PayForItem(lvKitItem.LineItem.ItemCode, Now).ToString("0.00")

            Me.txtKitEdit.Text = strPrice

            sender.StartEditing(Editors(0), e.Item, e.SubItem)
        ElseIf lvSubItem = 8 Then
            Debug.WriteLine("Banded")
        End If

    End Sub

    Private Sub txtKitEdit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKitEdit.Leave
        Dim strText As String = Me.txtKitEdit.Text
        If Not Char.IsDigit(strText.Chars(0)) Then
            strText = Mid(strText, 2)
        End If

        Dim existingItem As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry = lvKitItem

        Me.Cursor = Cursors.WaitCursor
        If Not lvKitItem Is Nothing Then
            Dim strTest = lvKitItem.LineItem.ItemCode
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim Discount As Double
            If CDbl(strText) = 0 Then
                Discount = 100
            ElseIf CDbl(strText) = Me.myCollOfficer.PriceForItem(strTest, Now) Then
                Me.RemoveScheme(lvKitItem.DiscountScheme)
                lvKitItem.Refresh()
                Me.Cursor = Cursors.Default
                Exit Sub
            Else
                Dim Price As Double = Me.myCollOfficer.PriceForItem(strTest, Now)
                Dim NewPrice As Double = CDbl(strText)
                Discount = (Price - NewPrice) * 100 / Price

            End If
            'do we want a scheme for this sample type or all samples
            DiscType = "CODE"
            Dim Cancel As Boolean = False
            ds = FindDiscountScheme(Discount, strTest, "CODE", Cancel)
            If Not ds Is Nothing Then 'We have a scheme
                If lvKitItem.DiscountScheme = ds.DiscountSchemeId Then 'Matches existing scheme
                    Exit Sub

                End If
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                CreateCustomerScheme(ds.DiscountSchemeId)
                If lvKitItem.DiscountScheme <> "-1" AndAlso ds.DiscountSchemeId <> CInt(lvKitItem.DiscountScheme) Then
                    Me.RemoveScheme(lvKitItem.DiscountScheme)
                End If
                mnuPayRateStartDate_Click(Nothing, Nothing)
                'End If
            ElseIf Cancel Then
                lvKitItem = existingItem
                lvKitItem.Refresh()

                txtKitEdit.Hide()
                Exit Sub
            Else
                CreateScheme(Discount, lvKitItem.LineItem.Description, lvKitItem.DiscountScheme, strTest, DiscType)
            End If
            lvKitItem.Refresh()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Function FindDiscountScheme(ByVal Discount As Double, ByVal ItemCode As String, ByVal SampleType As String, ByRef Cancel As Boolean) As MedscreenLib.Glossary.DiscountScheme
        Dim ds As MedscreenLib.Glossary.DiscountScheme
        Dim From As Integer = 0
        ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
        If ds Is Nothing Then
            Return Nothing
        Else
            Dim strRet As String = ds.FullDescription(Me.myCollOfficer.OfficerID)
            Dim intRet As MsgBoxResult = MsgBox("Do you want to use this scheme?" & vbCrLf & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            While intRet = MsgBoxResult.No And Not ds Is Nothing
                From += 1
                ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
                If Not ds Is Nothing Then
                    strRet = ds.FullDescription
                    intRet = MsgBox("Do you want to use this scheme?" & vbCrLf & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                End If
            End While
            If intRet = MsgBoxResult.No Then
                Cancel = True
                Return Nothing
            Else
                Return ds
            End If
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a scheme and add one entry 
    ''' </summary>
    ''' <param name="Discount"></param>
    ''' <param name="strType"></param>
    ''' <param name="ExistingScheme"></param>
    ''' <param name="ItemCode"></param>
    ''' <param name="DiscountType"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub CreateScheme(ByVal Discount As Double, ByVal strType As String, _
    ByVal ExistingScheme As String, ByVal ItemCode As String, ByVal DiscountType As String)
        'We need to create a scheme
        Dim strSchemeTitle As String = Math.Abs(Discount).ToString("0.00") & "% "
        If Discount < 0 Then
            strSchemeTitle += "Surcharge on "
        Else
            strSchemeTitle += "Discount on "
        End If
        'Check for discount type 
        If DiscountType = "TYPE" Then
            strSchemeTitle += "All "
        End If
        strSchemeTitle += strType
        If MsgBox("create discount scheme for " & strSchemeTitle, MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim objNewScheme As New MedscreenLib.Glossary.DiscountScheme()
            objNewScheme.Description = strSchemeTitle
            objNewScheme.DiscountSchemeId = MedscreenLib.CConnection.NextID("DISCOUNTSCHEME_HEADER", "IDENTITY")
            objNewScheme.Insert()

            objNewScheme.Update()
            MedscreenLib.Glossary.Glossary.DiscountSchemes.Add(objNewScheme)
            Dim objNewEntry As New MedscreenLib.Glossary.DiscountSchemeEntry()
            objNewEntry.ApplyTo = ItemCode
            objNewEntry.Discount = Discount
            objNewEntry.DiscountSchemeId = objNewScheme.DiscountSchemeId
            objNewEntry.DiscountType = DiscountType
            objNewEntry.EntryNumber = 1
            objNewEntry.Insert()
            objNewEntry.Update()
            objNewScheme.Entries.Add(objNewEntry)

            RemoveScheme(ExistingScheme)
            CreateCustomerScheme(objNewScheme.DiscountSchemeId)
        End If



    End Sub

    Private Sub mnuPayRateStartDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPayRateStartDate.Click
        If Me.lvKitItem Is Nothing Then Exit Sub

        Dim strDiscId As String = Me.lvKitItem.DiscountScheme
        If strDiscId = "-1" Then Exit Sub 'No discount in place 
        'Dim cdl As Intranet.intranet.customerns.CustomerDiscounts = New Intranet.intranet.customerns.CustomerDiscounts(Me.myCollOfficer.OfficerID, Intranet.customerns.CustomerDiscounts.CollectionType.Officer)
        Dim cd As Intranet.intranet.customerns.CustomerDiscount
        If strDiscId.Trim.Length <> 0 Then       'Check it is not a dummy string 
            cd = New Intranet.intranet.customerns.CustomerDiscount()
            cd.CustomerID = Me.myCollOfficer.OfficerID
            cd.DiscountSchemeId = strDiscId
            cd.LinkField = Intranet.intranet.customerns.CustomerDiscount.GCST_Officer_link
            cd.Load()
            If Not cd Is Nothing Then       'We've found the item remove it 
                Dim objForm As New MedscreenCommonGui.frmCustomerService()
                objForm.CustomerDiscount = cd
                If objForm.ShowDialog = DialogResult.OK Then
                    cd = objForm.CustomerDiscount
                    cd.Update()
                    Me.lvKitItem.Refresh()
                    'Me.FillPanels()
                End If
            End If
        End If

    End Sub

    Private Sub lblFaxWarning_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblFaxWarning.Click
        Me.lblFaxWarning.Hide()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Display list of instrument classes to user
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	11/07/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdaddInstrument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdaddInstrument.Click
        'Need to display to the users available models.
        Dim objModels As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_MODEL")
        objModels.Load("Phrase_text")
        Dim objModel As Object = Medscreen.GetParameter(objModels, "Model")
        If objModel Is Nothing Then Exit Sub
        'Find serioal numbers for model
        Dim objSerialNo As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("Select identity,serial_no from instrument", Glossary.PhraseCollection.BuildBy.Query, "model = '" & objModel & "' order by serial_no")
        'objSerialNo.Load("Description")
        Dim objInstID As Object = Medscreen.GetParameter(objSerialNo, "Select by serial no, cancel if not in list")
        If objInstID Is Nothing Then
            'need to ask whether they want to create the instrument.
        Else
            'we need to see if the instrument is assigned to anyone.
            Dim objRetOff As Object = CConnection.PackageStringList("Lib_instrument.OfficerHasInstrument", objInstID)
            Dim objCinst As Intranet.INSTRUMENT.OfficerInstrument
            If CStr(objRetOff).Trim.Length > 0 Then 'assigned to officer
                If objRetOff = Me.myCollOfficer.OfficerID Then

                    Dim oCmd As New OleDb.OleDbCommand()
                    oCmd.CommandText = "update officer_instrument set removeflag = 'F' where officer_id = ? and instrument_id = ?"
                    oCmd.Parameters.Add(CConnection.StringParameter("OfficerID", objRetOff, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("InstId", objInstID, 10))
                    Try ' Protecting 
                        oCmd.Connection = CConnection.DbConnection
                        If CConnection.ConnOpen Then            'Attempt to open reader
                            oCmd.ExecuteNonQuery()
                        End If

                    Catch ex As Exception
                        MedscreenLib.Medscreen.LogError(ex, , "frmCollOfficerEd-cmdaddInstrument_Click-4380")
                    Finally
                        CConnection.SetConnClosed()             'Close connection
                    End Try
                    Me.myCollOfficer.Instruments.Load()
                    Me.DrawInstruments()
                    Exit Sub
                End If

            End If

            objCinst = New Intranet.INSTRUMENT.OfficerInstrument()
            objCinst.InstrumentId = objInstID
            objCinst.OfficerId = Me.myCollOfficer.OfficerID
            objCinst.Load()
            If objCinst.CreatedOn = DateField.ZeroDate Then objCinst.CreatedOn = Now
            objCinst.ModifiedOn = Now
            If objCinst.CreatedBy.Trim.Length = 0 Then objCinst.CreatedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            objCinst.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            objCinst.Removed = False
            objCinst.DoUpdate()
            Dim objInst As Intranet.INSTRUMENT.Instrument = New Intranet.INSTRUMENT.Instrument(objCinst.InstrumentId)
            objInst.Load(CConnection.DbConnection)
            objInst.LocationId = Intranet.INSTRUMENT.Instrument.GCST_LocationField
            objInst.DoUpdate()

            Me.myCollOfficer.Instruments.Add(objCinst)
            Me.DrawInstruments()
        End If
    End Sub

    Private Sub mnuCollectionByPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollectionByPort.Click
        Dim objDate As Object
        objDate = Medscreen.GetParameter(MyTypes.typDate, "Satrt Date", "Start date of report", DateSerial(Today.Year, 1, 1))
        If objDate Is Nothing Then Exit Sub 'User cancelled 

        Dim objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        'This needs moving to config file
        objReport = Me.LogInCrystal("\\andrew\solutions\worktemp\OfficerPortColl.rpt")

        'Set up report formula

        Dim FF As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        For Each FF In objReport.DataDefinition.FormulaFields
            If FF.FormulaName.ToUpper = "{@OFFICER}" Then
                FF.Text = """" & Me.CollectingOfficer.OfficerID & """"
            End If
            If FF.FormulaName.ToUpper = "{@COLLDATE}" Then
                FF.Text = """" & CDate(objDate).ToString("dd-MMM-yyyy") & """"
            End If
        Next

        Dim oldDefault As Integer = MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod
        MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod = Constants.SendMethod.PDF
        OutputReport(objReport, "Collections by port " & Me.CollectingOfficer.Name & " from " & CDate(objDate).ToString("dd-MMM-yyyy"))
        objReport.Close()
        MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod = oldDefault
    End Sub

    Private currentOrder As ListViewItems.lvInvoice
    Private Sub lvOrders_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvOrders.SelectedIndexChanged
        currentOrder = Nothing

        If lvOrders.SelectedItems.Count = 1 Then
            currentOrder = Me.lvOrders.SelectedItems(0)
            Me.lblInvoiceDisp.Text = "Items for order - " & currentOrder.Invoice.InvoiceID
            Me.DrawOrderItems()
        End If
    End Sub

    Private Sub udMaxOrders_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udMaxOrders.ValueChanged
        Me.DrawOrders()
    End Sub

    Private Sub mnuEmailSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEmailSummary.Click
        Dim strHTML As String = Me.myCollOfficer.ToHtml
        Medscreen.BlatEmail("Collecting Officer Report for " & Me.myCollOfficer.Name, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit individual pay items
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	27/01/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub mnuEditPay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditPay.Click
    '    If Not CurrentExpPayItem Is Nothing Then
    '        Dim frmEdit As New frmExpEntry()
    '        frmEdit.CollOfficer = Me.myCollOfficer
    '        'End If
    '        frmEdit.FormType = frmExpEntry.enumExpFormType.PayDetail
    '        frmEdit.PayItem = CurrentExpPayItem.Expense
    '        frmEdit.FormType = frmExpEntry.enumExpFormType.PayDetail
    '        If frmEdit.ShowDialog = DialogResult.OK Then
    '            CurrentExpPayItem.Expense = frmEdit.PayItem
    '            CurrentExpPayItem.Expense.SetPriceInfo(frmEdit.Quantity.Value, CDbl(frmEdit.txtPrice.Text))
    '            CurrentExpPayItem.Expense.PayMonth = Me.dtPayMonth.Value
    '            CurrentExpPayItem.Expense.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
    '            CurrentExpPayItem.Refresh()
    '            CurrentExpPayItem.Expense.DoUpdate()
    '        End If


    '    End If
    'End Sub

    Private CurrentExpPayItem As ExpenseDescription

    'Private Sub lvPayItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    CurrentExpPayItem = Nothing
    '    'If Me.lvPayItems.SelectedItems.Count > 0 Then
    '    '    Me.CurrentExpPayItem = Me.lvPayItems.SelectedItems.Item(0)

    '    'End If

    'End Sub

    Private Sub treTeams_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treTeams.AfterSelect
        Dim anode As TreeNode = Me.treTeams.SelectedNode
        Me.treTeams.ContextMenu = Nothing
        If TypeOf anode Is Treenodes.PeopleNodes.TeamHeadersNode Then
            Dim onode As Treenodes.PeopleNodes.TeamHeadersNode = anode
            Me.treTeams.ContextMenu = onode.ContextMenu
        ElseIf TypeOf anode Is Treenodes.PeopleNodes.TeamEntryNode Then
            Dim onode As Treenodes.PeopleNodes.TeamEntryNode = anode
            Me.treTeams.ContextMenu = onode.ContextMenu
            Me.HtmlEditor1.DocumentText = (onode.ToHTML)
        ElseIf TypeOf anode Is Treenodes.PeopleNodes.TeamRegionNode Then
            Dim onode As Treenodes.PeopleNodes.TeamRegionNode = anode
            Me.ContextMenu = onode.ContextMenu
            HtmlEditor1.DocumentText = (onode.ToHTML)
        ElseIf TypeOf anode Is Treenodes.PeopleNodes.TeamHeader Then
            Dim onode As Treenodes.PeopleNodes.TeamHeader = anode
            onode.Refresh()
            HtmlEditor1.DocumentText = (onode.ToHTML)

            Me.ContextMenu = onode.ContextMenu
        End If

    End Sub

    Private Sub txtLatitude_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLatitude.Leave
        Dim strCoords As String = Intranet.intranet.jobs.CollectingOfficer.ConvertCoordinate(txtLatitude.Text)
        Me.txtLatitude.Text = strCoords
    End Sub

    Private Sub txtLongitude_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLongitude.Leave
        Dim strCoords As String = Intranet.intranet.jobs.CollectingOfficer.ConvertCoordinate(txtLongitude.Text)
        Me.txtLongitude.Text = strCoords
    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click

    End Sub

    Private Sub txtCoordinates_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoordinates.Leave
        If Me.txtCoordinates.Text.Trim.Length > 8 Then
            Try
                Dim strCoords As String = Intranet.intranet.jobs.CollectingOfficer.ConvertCoordinates(txtCoordinates.Text)
                Dim strLat As String = CConnection.PackageStringList("Lib_officer.GetCoordinateLatitude", strCoords)
                Dim strLong As String = CConnection.PackageStringList("Lib_officer.GetCoordinateLongitude", strCoords)
                Me.txtLatitude.Text = strLat
                Me.txtLongitude.Text = strLong
            Catch ex As Exception
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' Check that the Officer has no active collections if trying to remove
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/05/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ckRemoved_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckRemoved.CheckedChanged
        If ckRemoved.Checked Then
            Dim strRet As String = CConnection.PackageStringList("lib_officer.GetActiveCollections", Me.myCollOfficer.OfficerID)
            If strRet.Trim.Length > 0 Then
                MsgBox("The following collections are active for this officer, they need to be reassigned, before the officer can be marked as removed." & vbCrLf & _
                strRet, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                ckRemoved.Checked = False
            End If
        End If
    End Sub

    Private Sub LvCostHeader_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LvCostHeader.SelectedIndexChanged

    End Sub

    Private Sub mnuPayStatement_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPayStatement.Click

    End Sub
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Private Sub cmdAddExpense_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddExpense.Click
        Dim objCollPayItem As CollPay
        objCollPayItem = New CollPay()
        objCollPayItem.PayMonth = Me.dtPayMonth.Value
        objCollPayItem.PayType = CollPay.CST_PAYTYPE_OTHERCOSTSEXPENSES
        Dim strID As String = "E" & CConnection.NextSequence("SEQ_COLLECTORPAY_HEADER", 9)
        With objCollPayItem
            .CID = strID
            .Officer = Me.myCollOfficer.OfficerID
            .Collection = "OTHER EXPE"
            .PayMonth = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)
            .InvoiceMonth = Me.dtPayMonth.Value.ToString("MMMM yyyy")
            .ExpenseTypeID = "CEXPNOREC"
        End With
        'MyCollPayEntry.PaymentMonth = 
        Dim frmEdit As New frmExpEntry()
        '    Dim objCollPayItem As Intranet.collectorpay.CollPayItem = Me.MyCollPayEntry.GetExpenseItem

        frmEdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
        '    If objCollPayItem Is Nothing Then Exit Sub 'Fatal error

        '    objCollPayItem.PayMonth = Me.dtPayMonth.Value
        '    objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

        '    'If Me.FormType = enFormType.CollRelatedPay Then
        'frmEdit.CollOfficer = Me.objCollOfficer
        '    'Else
        frmEdit.CollOfficer = Me.myCollOfficer
        '    'End If
        frmEdit.PayItem = objCollPayItem
        If frmEdit.ShowDialog = DialogResult.OK Then
            objCollPayItem = frmEdit.PayItem
            'objCollPayItem.SetPriceInfo(frmEdit.Quantity.Value, CDbl(frmEdit.txtPrice.Text))
            objCollPayItem.PayMonth = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)
            objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
            objCollPayItem.DoUpdate()
            Dim objListItem As New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(objCollPayItem, Me.CollectingOfficer.Culture)
            Me.lvExpenseHeader.Items.Add(objListItem)

        Else
            'objCollPayItem.CancelItem() ' = True
            objCollPayItem.Delete()

        End If

        '    MyCollPayEntry.SetOfficer(Me.myCollOfficer.OfficerID, Me.dtPayMonth.Value, "E")

        '    Me.FillHeaderExpListView()

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

    End Sub

    Private Sub cmdAddMiscPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddMiscPay.Click
        Dim objCollPayItem As CollPay
        objCollPayItem = New CollPay()
        objCollPayItem.PayMonth = Me.dtPayMonth.Value
        objCollPayItem.PayType = CollPay.CST_PAYTYPE_OTHERCOST
        Dim strID As String = "P" & CConnection.NextSequence("SEQ_COLLECTORPAY_HEADER", 9)
        With objCollPayItem
            .CID = strID
            .Officer = Me.myCollOfficer.OfficerID
            .Collection = "OTHER COST"
            .PayMonth = DateSerial(Me.dtPayMonth.Value.Year, Me.dtPayMonth.Value.Month, 1)
            .InvoiceMonth = .PayMonth.ToString("MMMMyyyy")
            .PayType = CollPay.CST_PAYTYPE_OTHERCOST
            .LeaveSite = Today
            .ArriveOnSite = Today
            .TravelStart = Today
            .TravelEnd = Today
            objCollPayItem.ExpenseTypeID = cbPayType2.SelectedItem.phraseid
            .DoUpdate()
        End With
        Dim CollPayForm As New MedscreenCommonGui.frmCollPay()
        CollPayForm.PayItem = objCollPayItem

        If CollPayForm.ShowDialog() = DialogResult.OK Then
            CollPayForm.PayItem.DoUpdate()
            Dim objListItem As New MedscreenCommonGui.ListViewItems.LVOldCollectorpayHeader(objCollPayItem, Me.CollectingOfficer.Culture)
            Me.lvExpenseHeader.Items.Add(objListItem)

        Else
            'delete item
            objCollPayItem.Delete()
        End If
    End Sub



    Private Sub LvCostHeader_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LvCostHeader.KeyDown
        If lvPayCurrentItem Is Nothing Then Exit Sub
        If e.KeyCode = Keys.Delete Then
            Dim Result As MsgBoxResult = MsgBox("Do you want to delete this item?" & lvPayCurrentItem.CollPayHeader.CID, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If Result = MsgBoxResult.Yes Then
                lvPayCurrentItem.CollPayHeader.Delete()
                LvCostHeader.Items.Remove(lvPayCurrentItem)
            End If
        End If
    End Sub

    Private Sub lvExpenseHeader_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvExpenseHeader.KeyDown
        If lvCurrentItem Is Nothing Then Exit Sub
        If e.KeyCode = Keys.Delete Then
            Dim Result As MsgBoxResult = MsgBox("Do you want to delete this item?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If Result = MsgBoxResult.Yes Then
                lvCurrentItem.CollPayHeader.Delete()
                lvExpenseHeader.Items.Remove(lvCurrentItem)
            End If

        End If
    End Sub

    Private Sub TabPay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPay.Click
        blnDisplayPay = True
        'FillHeaderExpListView()
        FillHeaderCostListView()

    End Sub

    Private Sub TabExpenses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabExpenses.Click
        blnDisplayPay = True
        FillHeaderExpListView()
        'FillHeaderCostListView()
    End Sub

    Private Sub tabPaydetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabPaydetails.Click
        'blnDisplayPay = True
        'FillHeaderExpListView()
        'FillHeaderCostListView()
    End Sub

    
    Private Sub lvHolidays_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvHolidays.Click

    End Sub
End Class
