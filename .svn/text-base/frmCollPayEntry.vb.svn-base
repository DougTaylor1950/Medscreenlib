Imports Intranet
Imports Intranet.collectorpay
Imports System.Windows.Forms
Imports MedscreenLib
Imports System.Drawing
Imports MedscreenCommonGui.ListViewItems

Public Class frmCollPayEntry
    Inherits System.Windows.Forms.Form
    Private objCM As Intranet.intranet.jobs.CMJob
    Private objCollOfficer As Intranet.intranet.jobs.CollectingOfficer
    Private lvCurrentItem As ExpenseDescription
    Private lvCurrentPayItem As ExpenseDescription
    Private MyCollPayInHoursCost As Intranet.collectorpay.CollPayItem               'Primary collecting officers pay item, can be 
    Private MyMileageCost As Intranet.collectorpay.CollPayItem               'Primary collecting officers pay item, can be 
    '                                                                               'Either sample or time based (in hours only) or minimum charge.
    Private MyCollPayOutOfHoursCost As Intranet.collectorpay.CollPayItem            'Secondary pay item (only used for Out of hours pay)
    Private NewForm As Boolean = True
    Private UpdateCollPayDetails As Boolean = False
    Private myFormType As enFormType = enFormType.CollRelatedPay

    'Roles
    Private blnCanEnterPay As Boolean = False
    Private blnUserCanLockPay As Boolean = False
    Private blnDrawing As Boolean = False
    Private Const cstPaySampleStub = "PAYSAMPLE00"

    'enumeration to control form behaviour
    Public Enum enFormType
        CollRelatedPay
        OCR
        NonCollRelatedExp
    End Enum

    Private Enum ItemType
        Pay
        Expense
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the form type to the outside world
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	02/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property FormType() As enFormType
        Get
            Return Me.myFormType
        End Get
        Set(ByVal Value As enFormType)
            Me.myFormType = Value
            Me.PanelSamplePlus.Hide()
            Select Case myFormType
                'Deal with collection pay 
            Case enFormType.CollRelatedPay
                    Me.gBoxCollOpts.Show()
                    Me.gBoxPayOpts.Show()
                    Me.gBoxTimeEntry.Show()
                    Me.grpPerSampleRate.Show()

                    Me.TimeSumBox.Show()
                    Me.PaySummBox.Show()
                    Me.TabControl1.TabPages(0).Text = "Collection Costs"
                    Me.lvPayItems.Show()

                    Me.CollDate.Visible = True
                    Me.Donors.Visible = True
                    Me.lblCancSamples.Enabled = True
                    Me.txtCMidentity.Enabled = True
                    Me.cbCollOfficer.Enabled = True
                    Me.cbCollOfficer.Width = 121
                    Me.txtCustomerID.Enabled = True
                    Me.txtVesselName.Enabled = True
                    Me.txtSitePort.Enabled = True

                    Me.cbCollOfficer.Width = 121
                    Me.lblCollDate.Visible = True
                    Me.CollDate.Visible = True
                    Me.lblDonors.Visible = True
                    Me.Donors.Visible = True
                    Me.lblCancSamples.Visible = True
                    Me.txtCMidentity.Visible = True
                    Me.txtCustomerID.Visible = True
                    Me.txtVesselName.Visible = True
                    Me.txtSitePort.Visible = True
                    Me.lblvessel.Visible = True
                    Me.lblCustomer.Visible = True
                    Me.lblSite.Visible = True
                    Me.txtPayAmount.Visible = True
                    Me.lblExpSubTotal.Visible = True
                    Me.lblCostsSubTotal.Visible = True
                    Me.HtmlEditor1.Visible = False
                    Me.Panel2.Show()

                    'If we are dealing with an on call rota payment or expense header
                Case enFormType.NonCollRelatedExp, enFormType.OCR
                    Me.gBoxCollOpts.Hide()
                    Me.gBoxPayOpts.Hide()
                    Me.gBoxTimeEntry.Hide()
                    Me.grpPerSampleRate.Hide()
                    Me.Panel2.Hide()

                    Me.TimeSumBox.Hide()
                    Me.PaySummBox.Hide()

                    Me.cbCollOfficer.Enabled = True
                    Me.cbCollOfficer.Width = 408
                    Me.TabControl1.TabPages(0).Text = "Non-Collection Related Expenses"

                    Me.lblCollDate.Visible = False
                    Me.CollDate.Visible = False
                    Me.lblDonors.Visible = False
                    Me.Donors.Visible = False
                    Me.lblCancSamples.Visible = False
                    Me.txtCMidentity.Visible = False
                    Me.txtCustomerID.Visible = False
                    Me.txtVesselName.Visible = False
                    Me.txtSitePort.Visible = False
                    Me.lblvessel.Visible = False
                    Me.lblCustomer.Visible = False
                    Me.lblSite.Visible = False
                    Me.txtPayAmount.Visible = False
                    Me.lblCostsSubTotal.Visible = False
                    Me.Text = "Non-Collection Associated Expenses."
                    Me.HtmlEditor1.Visible = True

                    Me.lvPayItems.Hide()
                    Me.panelCollPayDefaults.Visible = True

                    If Me.MyCollPayEntry Is Nothing Then
                        MyCollPayEntry = Intranet.collectorpay.CollectorPayCollection.CreateHeader(CollectorPayCollection.PayTypes.OFFICER)
                    Else
                        MyCollPayEntry.PayType = ExpenseTypeCollections.GCST_Coll_Pay_Entry
                    End If
                    'MyCollPayEntry.CollOfficer = Me.cbCollOfficer.SelectedIndex
                    'MyCollPayEntry.PaymentMonth = Me.dtPayMonth.Value
                    'mycollpayentry.PayType = 
                    MyCollPayEntry.DoUpdate()

            End Select
            If Value = enFormType.CollRelatedPay Then
                UpDateCollPayItem()
                Me.UpdateMileage(True)
            Else 'Deal with collecting officer items
                If Not Me.CollPay Is Nothing Then
                    Me.cbCollOfficer.SelectedValue = Me.CollPay.CollOfficer
                    RefreshCollCostListView()
                End If
            End If
        End Set
    End Property

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer. 
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Debug.Write("test")
        Me.cbCollOfficer.DataSource = Intranet.intranet.Support.cSupport.CollectingOfficers
        Me.cbCollOfficer.DisplayMember = "Info2"
        Me.cbCollOfficer.ValueMember = "OfficerID"

        Dim objCurr As New Glossary.PhraseCollection("Select Currencyid,description from currency", Glossary.PhraseCollection.BuildBy.Query)
        objCurr.Load()
        Me.cbCurrency.DataSource = objCurr
        Me.cbCurrency.ValueMember = "PhraseID"
        Me.cbCurrency.DisplayMember = "PhraseText"

        'See if we are allowed to do things
        Me.blnCanEnterPay = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_PAY_ENTRY")
        blnUserCanLockPay = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_PAY_LOCK")


        'Get sample lineitems
        Dim i As Integer
        For i = 1 To MedscreenCommonGUIConfig.SampleLineItems.Count
            Me.updLineitem.Items.Add(MedscreenCommonGUIConfig.SampleLineItems.Item(cstPaySampleStub & CStr(i).Trim))
        Next
        Me.updLineitem.SelectedIndex = 0

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
    '<System.Diagnostics.DebuggerStepThrough()> 
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents panelCollPayDefaults As System.Windows.Forms.Panel
    Friend WithEvents lblPaymentMonth As System.Windows.Forms.Label
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Friend WithEvents txtSitePort As System.Windows.Forms.TextBox
    Friend WithEvents txtCMidentity As System.Windows.Forms.TextBox
    Friend WithEvents txtVesselName As System.Windows.Forms.TextBox
    Friend WithEvents lblvessel As System.Windows.Forms.Label
    Friend WithEvents lblCollDate As System.Windows.Forms.Label
    Friend WithEvents lblCMIdentity As System.Windows.Forms.Label
    Friend WithEvents lblSite As System.Windows.Forms.Label
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents lblCollOff As System.Windows.Forms.Label
    Friend WithEvents lblCancSamples As System.Windows.Forms.Label
    Friend WithEvents Donors As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblDonors As System.Windows.Forms.Label
    Friend WithEvents CollDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVALUE As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtExpSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents dtPayMonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents lvPayItems As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtPayAmount As System.Windows.Forms.TextBox
    Friend WithEvents chDiscount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDisc As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents txtCostTotal As System.Windows.Forms.TextBox
    Friend WithEvents lblSample As System.Windows.Forms.Label
    Friend WithEvents lblCostsTotal As System.Windows.Forms.Label
    Friend WithEvents txtCostsTotal As System.Windows.Forms.TextBox
    Friend WithEvents lblInHours As System.Windows.Forms.Label
    Friend WithEvents gBoxTimeEntry As System.Windows.Forms.GroupBox
    Friend WithEvents lblimeArrHome As System.Windows.Forms.Label
    Friend WithEvents lblTimeLeftSite As System.Windows.Forms.Label
    Friend WithEvents lblArrOnSite As System.Windows.Forms.Label
    Friend WithEvents lblTimeSetOff As System.Windows.Forms.Label
    Friend WithEvents DteTimeArrHome As System.Windows.Forms.DateTimePicker
    Friend WithEvents DteTimeLeftSite As System.Windows.Forms.DateTimePicker
    Friend WithEvents DteTimeArrOnSite As System.Windows.Forms.DateTimePicker
    Friend WithEvents DteTimeSetOff As System.Windows.Forms.DateTimePicker
    Friend WithEvents gBoxCollOpts As System.Windows.Forms.GroupBox
    Friend WithEvents rButCallOut As System.Windows.Forms.RadioButton
    Friend WithEvents rButRoutine As System.Windows.Forms.RadioButton
    Friend WithEvents gBoxPayOpts As System.Windows.Forms.GroupBox
    Friend WithEvents rButHourly As System.Windows.Forms.RadioButton
    Friend WithEvents rButSamples As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOutHours As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalInHours As System.Windows.Forms.TextBox
    Friend WithEvents lblHoursWorked As System.Windows.Forms.Label
    Friend WithEvents txtHoursWorked As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblPerSample As System.Windows.Forms.Label
    Friend WithEvents grpPerSampleRate As System.Windows.Forms.GroupBox
    Friend WithEvents TimeSumBox As System.Windows.Forms.GroupBox
    Friend WithEvents PaySummBox As System.Windows.Forms.GroupBox
    Friend WithEvents txtTotalExp As System.Windows.Forms.TextBox
    Friend WithEvents lblUserInfo As System.Windows.Forms.Label
    Friend WithEvents cbCollOfficer As System.Windows.Forms.ComboBox
    Friend WithEvents lblExpSubTotal As System.Windows.Forms.Label
    Friend WithEvents lblCostsSubTotal As System.Windows.Forms.Label
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents udMileage As System.Windows.Forms.NumericUpDown
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuPay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLockPay As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuApproveItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeclineItem As System.Windows.Forms.MenuItem
    Friend WithEvents MnuReferItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCancelItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEntered As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Comment As System.Windows.Forms.Label
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents rbOther As System.Windows.Forms.RadioButton
    Friend WithEvents cmdAddPay As System.Windows.Forms.Button
    Friend WithEvents udKMMiles As System.Windows.Forms.DomainUpDown
    Friend WithEvents txtMileage As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ckIgnoreMinimum As System.Windows.Forms.CheckBox
    Friend WithEvents rbCollection As System.Windows.Forms.RadioButton
    Friend WithEvents rbWorkingTime As System.Windows.Forms.RadioButton
    Friend WithEvents DTWorkingTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTravelTime As System.Windows.Forms.Label
    Friend WithEvents dtTravelTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents mnuReports As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaystatement As System.Windows.Forms.MenuItem
    Friend WithEvents chComment As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommentP As System.Windows.Forms.ColumnHeader
    Friend WithEvents udHalfdays As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblHalfDays As System.Windows.Forms.Label
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents cmdAddOfficer As System.Windows.Forms.Button
    Friend WithEvents lvOfficers As System.Windows.Forms.ListView
    Friend WithEvents lblTimeWarning As System.Windows.Forms.Label
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents lblTimeOnSite As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents lblLeaveHome As System.Windows.Forms.Label
    Friend WithEvents lblArriveHome As System.Windows.Forms.Label
    Friend WithEvents PanelSamplePlus As System.Windows.Forms.Panel
    Friend WithEvents updLineitem As System.Windows.Forms.DomainUpDown
    Friend WithEvents mnuUnlockPay As System.Windows.Forms.MenuItem
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtGrandTotal As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuApprovePay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditItem As System.Windows.Forms.MenuItem
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCollPayEntry))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtComment = New System.Windows.Forms.TextBox()
        Me.Comment = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblTimeWarning = New System.Windows.Forms.Label()
        Me.ckIgnoreMinimum = New System.Windows.Forms.CheckBox()
        Me.PaySummBox = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtMileage = New System.Windows.Forms.TextBox()
        Me.txtCostTotal = New System.Windows.Forms.TextBox()
        Me.txtTotalExp = New System.Windows.Forms.TextBox()
        Me.lblSample = New System.Windows.Forms.Label()
        Me.lblCostsTotal = New System.Windows.Forms.Label()
        Me.txtCostsTotal = New System.Windows.Forms.TextBox()
        Me.lblInHours = New System.Windows.Forms.Label()
        Me.gBoxTimeEntry = New System.Windows.Forms.GroupBox()
        Me.lblArriveHome = New System.Windows.Forms.Label()
        Me.lblLeaveHome = New System.Windows.Forms.Label()
        Me.lblTimeOnSite = New System.Windows.Forms.Label()
        Me.udKMMiles = New System.Windows.Forms.DomainUpDown()
        Me.udMileage = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblimeArrHome = New System.Windows.Forms.Label()
        Me.lblTimeLeftSite = New System.Windows.Forms.Label()
        Me.lblArrOnSite = New System.Windows.Forms.Label()
        Me.lblTimeSetOff = New System.Windows.Forms.Label()
        Me.DteTimeArrHome = New System.Windows.Forms.DateTimePicker()
        Me.DteTimeLeftSite = New System.Windows.Forms.DateTimePicker()
        Me.DteTimeArrOnSite = New System.Windows.Forms.DateTimePicker()
        Me.DteTimeSetOff = New System.Windows.Forms.DateTimePicker()
        Me.gBoxCollOpts = New System.Windows.Forms.GroupBox()
        Me.rButCallOut = New System.Windows.Forms.RadioButton()
        Me.rButRoutine = New System.Windows.Forms.RadioButton()
        Me.gBoxPayOpts = New System.Windows.Forms.GroupBox()
        Me.rbWorkingTime = New System.Windows.Forms.RadioButton()
        Me.rbCollection = New System.Windows.Forms.RadioButton()
        Me.rbOther = New System.Windows.Forms.RadioButton()
        Me.rButHourly = New System.Windows.Forms.RadioButton()
        Me.rButSamples = New System.Windows.Forms.RadioButton()
        Me.TimeSumBox = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtOutHours = New System.Windows.Forms.TextBox()
        Me.txtTotalInHours = New System.Windows.Forms.TextBox()
        Me.lblHoursWorked = New System.Windows.Forms.Label()
        Me.txtHoursWorked = New System.Windows.Forms.TextBox()
        Me.grpPerSampleRate = New System.Windows.Forms.GroupBox()
        Me.lblPerSample = New System.Windows.Forms.Label()
        Me.panelCollPayDefaults = New System.Windows.Forms.Panel()
        Me.PanelSamplePlus = New System.Windows.Forms.Panel()
        Me.updLineitem = New System.Windows.Forms.DomainUpDown()
        Me.lvOfficers = New System.Windows.Forms.ListView()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.cmdAddOfficer = New System.Windows.Forms.Button()
        Me.lblHalfDays = New System.Windows.Forms.Label()
        Me.udHalfdays = New System.Windows.Forms.NumericUpDown()
        Me.dtTravelTime = New System.Windows.Forms.DateTimePicker()
        Me.lblTravelTime = New System.Windows.Forms.Label()
        Me.DTWorkingTime = New System.Windows.Forms.DateTimePicker()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cbCurrency = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbCollOfficer = New System.Windows.Forms.ComboBox()
        Me.dtPayMonth = New System.Windows.Forms.DateTimePicker()
        Me.CollDate = New System.Windows.Forms.DateTimePicker()
        Me.lblPaymentMonth = New System.Windows.Forms.Label()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.txtCustomerID = New System.Windows.Forms.TextBox()
        Me.txtSitePort = New System.Windows.Forms.TextBox()
        Me.txtVesselName = New System.Windows.Forms.TextBox()
        Me.lblvessel = New System.Windows.Forms.Label()
        Me.lblCollDate = New System.Windows.Forms.Label()
        Me.lblSite = New System.Windows.Forms.Label()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.lblCollOff = New System.Windows.Forms.Label()
        Me.Donors = New System.Windows.Forms.NumericUpDown()
        Me.lblDonors = New System.Windows.Forms.Label()
        Me.lblCancSamples = New System.Windows.Forms.Label()
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtGrandTotal = New System.Windows.Forms.TextBox()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.cmdAddPay = New System.Windows.Forms.Button()
        Me.lblCostsSubTotal = New System.Windows.Forms.Label()
        Me.lblExpSubTotal = New System.Windows.Forms.Label()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.txtPayAmount = New System.Windows.Forms.TextBox()
        Me.lvPayItems = New System.Windows.Forms.ListView()
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader()
        Me.chDisc = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.chCommentP = New System.Windows.Forms.ColumnHeader()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.mnuApproveItem = New System.Windows.Forms.MenuItem()
        Me.mnuDeclineItem = New System.Windows.Forms.MenuItem()
        Me.MnuReferItem = New System.Windows.Forms.MenuItem()
        Me.mnuCancelItem = New System.Windows.Forms.MenuItem()
        Me.mnuEntered = New System.Windows.Forms.MenuItem()
        Me.mnuEditItem = New System.Windows.Forms.MenuItem()
        Me.txtExpSubTotal = New System.Windows.Forms.TextBox()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader()
        Me.chPrice = New System.Windows.Forms.ColumnHeader()
        Me.chDiscount = New System.Windows.Forms.ColumnHeader()
        Me.chVALUE = New System.Windows.Forms.ColumnHeader()
        Me.chStatus = New System.Windows.Forms.ColumnHeader()
        Me.chComment = New System.Windows.Forms.ColumnHeader()
        Me.txtCMidentity = New System.Windows.Forms.TextBox()
        Me.lblCMIdentity = New System.Windows.Forms.Label()
        Me.lblUserInfo = New System.Windows.Forms.Label()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.mnuPay = New System.Windows.Forms.MenuItem()
        Me.mnuLockPay = New System.Windows.Forms.MenuItem()
        Me.mnuUnlockPay = New System.Windows.Forms.MenuItem()
        Me.mnuApprovePay = New System.Windows.Forms.MenuItem()
        Me.mnuReports = New System.Windows.Forms.MenuItem()
        Me.mnuPaystatement = New System.Windows.Forms.MenuItem()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.PaySummBox.SuspendLayout()
        Me.gBoxTimeEntry.SuspendLayout()
        CType(Me.udMileage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gBoxCollOpts.SuspendLayout()
        Me.gBoxPayOpts.SuspendLayout()
        Me.TimeSumBox.SuspendLayout()
        Me.grpPerSampleRate.SuspendLayout()
        Me.panelCollPayDefaults.SuspendLayout()
        Me.PanelSamplePlus.SuspendLayout()
        CType(Me.udHalfdays, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Donors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 585)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(824, 40)
        Me.Panel1.TabIndex = 10
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(744, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(656, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 25)
        Me.cmdOk.TabIndex = 1
        Me.cmdOk.Text = "OK"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2})
        Me.TabControl1.Location = New System.Drawing.Point(0, 32)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(824, 548)
        Me.TabControl1.TabIndex = 11
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtComment, Me.Comment, Me.Panel2, Me.panelCollPayDefaults, Me.HtmlEditor1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(816, 522)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Collection Costs"
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(96, 192)
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(568, 20)
        Me.txtComment.TabIndex = 18
        Me.txtComment.Text = ""
        '
        'Comment
        '
        Me.Comment.Location = New System.Drawing.Point(16, 192)
        Me.Comment.Name = "Comment"
        Me.Comment.TabIndex = 17
        Me.Comment.Text = "Comment"
        '
        'Panel2
        '
        Me.Panel2.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblTimeWarning, Me.ckIgnoreMinimum, Me.PaySummBox, Me.gBoxTimeEntry, Me.gBoxCollOpts, Me.gBoxPayOpts, Me.TimeSumBox, Me.grpPerSampleRate})
        Me.Panel2.Location = New System.Drawing.Point(0, 216)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(808, 280)
        Me.Panel2.TabIndex = 15
        '
        'lblTimeWarning
        '
        Me.lblTimeWarning.BackColor = System.Drawing.SystemColors.Info
        Me.lblTimeWarning.ForeColor = System.Drawing.Color.Red
        Me.lblTimeWarning.Location = New System.Drawing.Point(312, 200)
        Me.lblTimeWarning.Name = "lblTimeWarning"
        Me.lblTimeWarning.Size = New System.Drawing.Size(400, 16)
        Me.lblTimeWarning.TabIndex = 13
        '
        'ckIgnoreMinimum
        '
        Me.ckIgnoreMinimum.Location = New System.Drawing.Point(160, 96)
        Me.ckIgnoreMinimum.Name = "ckIgnoreMinimum"
        Me.ckIgnoreMinimum.TabIndex = 12
        Me.ckIgnoreMinimum.Text = "Ignore Minimum"
        '
        'PaySummBox
        '
        Me.PaySummBox.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label5, Me.txtMileage, Me.txtCostTotal, Me.txtTotalExp, Me.lblSample, Me.lblCostsTotal, Me.txtCostsTotal, Me.lblInHours})
        Me.PaySummBox.Location = New System.Drawing.Point(160, 136)
        Me.PaySummBox.Name = "PaySummBox"
        Me.PaySummBox.Size = New System.Drawing.Size(128, 136)
        Me.PaySummBox.TabIndex = 10
        Me.PaySummBox.TabStop = False
        Me.PaySummBox.Text = "Payment Summary"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Travel:"
        '
        'txtMileage
        '
        Me.txtMileage.Enabled = False
        Me.txtMileage.Location = New System.Drawing.Point(72, 72)
        Me.txtMileage.Name = "txtMileage"
        Me.txtMileage.Size = New System.Drawing.Size(48, 20)
        Me.txtMileage.TabIndex = 7
        Me.txtMileage.Text = "0"
        '
        'txtCostTotal
        '
        Me.txtCostTotal.Enabled = False
        Me.txtCostTotal.Location = New System.Drawing.Point(72, 16)
        Me.txtCostTotal.Name = "txtCostTotal"
        Me.txtCostTotal.Size = New System.Drawing.Size(48, 20)
        Me.txtCostTotal.TabIndex = 6
        Me.txtCostTotal.Text = "0"
        '
        'txtTotalExp
        '
        Me.txtTotalExp.Enabled = False
        Me.txtTotalExp.Location = New System.Drawing.Point(72, 40)
        Me.txtTotalExp.Name = "txtTotalExp"
        Me.txtTotalExp.Size = New System.Drawing.Size(48, 20)
        Me.txtTotalExp.TabIndex = 1
        Me.txtTotalExp.Text = "0"
        '
        'lblSample
        '
        Me.lblSample.Location = New System.Drawing.Point(8, 40)
        Me.lblSample.Name = "lblSample"
        Me.lblSample.Size = New System.Drawing.Size(56, 16)
        Me.lblSample.TabIndex = 0
        Me.lblSample.Text = "Expenses:"
        '
        'lblCostsTotal
        '
        Me.lblCostsTotal.Location = New System.Drawing.Point(8, 104)
        Me.lblCostsTotal.Name = "lblCostsTotal"
        Me.lblCostsTotal.Size = New System.Drawing.Size(48, 16)
        Me.lblCostsTotal.TabIndex = 0
        Me.lblCostsTotal.Text = "Total:"
        '
        'txtCostsTotal
        '
        Me.txtCostsTotal.Enabled = False
        Me.txtCostsTotal.Location = New System.Drawing.Point(72, 104)
        Me.txtCostsTotal.Name = "txtCostsTotal"
        Me.txtCostsTotal.Size = New System.Drawing.Size(48, 20)
        Me.txtCostsTotal.TabIndex = 1
        Me.txtCostsTotal.Text = "0"
        '
        'lblInHours
        '
        Me.lblInHours.Location = New System.Drawing.Point(8, 16)
        Me.lblInHours.Name = "lblInHours"
        Me.lblInHours.Size = New System.Drawing.Size(40, 16)
        Me.lblInHours.TabIndex = 5
        Me.lblInHours.Text = "Costs:"
        '
        'gBoxTimeEntry
        '
        Me.gBoxTimeEntry.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.gBoxTimeEntry.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblArriveHome, Me.lblLeaveHome, Me.lblTimeOnSite, Me.udKMMiles, Me.udMileage, Me.Label4, Me.lblimeArrHome, Me.lblTimeLeftSite, Me.lblArrOnSite, Me.lblTimeSetOff, Me.DteTimeArrHome, Me.DteTimeLeftSite, Me.DteTimeArrOnSite, Me.DteTimeSetOff})
        Me.gBoxTimeEntry.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gBoxTimeEntry.Location = New System.Drawing.Point(312, 16)
        Me.gBoxTimeEntry.Name = "gBoxTimeEntry"
        Me.gBoxTimeEntry.Size = New System.Drawing.Size(472, 184)
        Me.gBoxTimeEntry.TabIndex = 4
        Me.gBoxTimeEntry.TabStop = False
        Me.gBoxTimeEntry.Text = "Record Of Times"
        '
        'lblArriveHome
        '
        Me.lblArriveHome.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lblArriveHome.Location = New System.Drawing.Point(256, 120)
        Me.lblArriveHome.Name = "lblArriveHome"
        Me.lblArriveHome.Size = New System.Drawing.Size(208, 56)
        Me.lblArriveHome.TabIndex = 13
        '
        'lblLeaveHome
        '
        Me.lblLeaveHome.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lblLeaveHome.Location = New System.Drawing.Point(256, 24)
        Me.lblLeaveHome.Name = "lblLeaveHome"
        Me.lblLeaveHome.Size = New System.Drawing.Size(200, 48)
        Me.lblLeaveHome.TabIndex = 12
        '
        'lblTimeOnSite
        '
        Me.lblTimeOnSite.Location = New System.Drawing.Point(256, 96)
        Me.lblTimeOnSite.Name = "lblTimeOnSite"
        Me.lblTimeOnSite.Size = New System.Drawing.Size(168, 23)
        Me.lblTimeOnSite.TabIndex = 11
        Me.lblTimeOnSite.Text = "Time On Site :"
        '
        'udKMMiles
        '
        Me.udKMMiles.Items.Add("Miles")
        Me.udKMMiles.Items.Add("Km")
        Me.udKMMiles.Location = New System.Drawing.Point(216, 153)
        Me.udKMMiles.Name = "udKMMiles"
        Me.udKMMiles.Size = New System.Drawing.Size(40, 20)
        Me.udKMMiles.TabIndex = 10
        Me.udKMMiles.Text = "Miles"
        '
        'udMileage
        '
        Me.udMileage.Location = New System.Drawing.Point(128, 154)
        Me.udMileage.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.udMileage.Name = "udMileage"
        Me.udMileage.Size = New System.Drawing.Size(88, 20)
        Me.udMileage.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 24)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Mileage:"
        '
        'lblimeArrHome
        '
        Me.lblimeArrHome.Location = New System.Drawing.Point(8, 123)
        Me.lblimeArrHome.Name = "lblimeArrHome"
        Me.lblimeArrHome.Size = New System.Drawing.Size(112, 24)
        Me.lblimeArrHome.TabIndex = 7
        Me.lblimeArrHome.Text = "Time Arrived Home:"
        '
        'lblTimeLeftSite
        '
        Me.lblTimeLeftSite.Location = New System.Drawing.Point(8, 91)
        Me.lblTimeLeftSite.Name = "lblTimeLeftSite"
        Me.lblTimeLeftSite.Size = New System.Drawing.Size(112, 24)
        Me.lblTimeLeftSite.TabIndex = 6
        Me.lblTimeLeftSite.Text = "Time Left Site:"
        '
        'lblArrOnSite
        '
        Me.lblArrOnSite.Location = New System.Drawing.Point(8, 59)
        Me.lblArrOnSite.Name = "lblArrOnSite"
        Me.lblArrOnSite.Size = New System.Drawing.Size(112, 24)
        Me.lblArrOnSite.TabIndex = 5
        Me.lblArrOnSite.Text = "Time Arrived On Site:"
        '
        'lblTimeSetOff
        '
        Me.lblTimeSetOff.Location = New System.Drawing.Point(8, 27)
        Me.lblTimeSetOff.Name = "lblTimeSetOff"
        Me.lblTimeSetOff.Size = New System.Drawing.Size(112, 24)
        Me.lblTimeSetOff.TabIndex = 4
        Me.lblTimeSetOff.Text = "Time Set Off:"
        '
        'DteTimeArrHome
        '
        Me.DteTimeArrHome.CustomFormat = "dd-MMM-yy HH:mm"
        Me.DteTimeArrHome.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DteTimeArrHome.Location = New System.Drawing.Point(128, 123)
        Me.DteTimeArrHome.Name = "DteTimeArrHome"
        Me.DteTimeArrHome.Size = New System.Drawing.Size(112, 20)
        Me.DteTimeArrHome.TabIndex = 5
        Me.DteTimeArrHome.Value = New Date(2007, 1, 30, 13, 37, 15, 442)
        '
        'DteTimeLeftSite
        '
        Me.DteTimeLeftSite.CustomFormat = "dd-MMM-yy HH:mm"
        Me.DteTimeLeftSite.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DteTimeLeftSite.Location = New System.Drawing.Point(128, 91)
        Me.DteTimeLeftSite.Name = "DteTimeLeftSite"
        Me.DteTimeLeftSite.Size = New System.Drawing.Size(112, 20)
        Me.DteTimeLeftSite.TabIndex = 4
        Me.DteTimeLeftSite.Value = New Date(2007, 1, 30, 13, 37, 15, 482)
        '
        'DteTimeArrOnSite
        '
        Me.DteTimeArrOnSite.CustomFormat = "dd-MMM-yy HH:mm"
        Me.DteTimeArrOnSite.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DteTimeArrOnSite.Location = New System.Drawing.Point(128, 59)
        Me.DteTimeArrOnSite.Name = "DteTimeArrOnSite"
        Me.DteTimeArrOnSite.Size = New System.Drawing.Size(112, 20)
        Me.DteTimeArrOnSite.TabIndex = 3
        Me.DteTimeArrOnSite.Value = New Date(2007, 1, 30, 13, 37, 15, 492)
        '
        'DteTimeSetOff
        '
        Me.DteTimeSetOff.CustomFormat = "dd-MMM-yy HH:mm"
        Me.DteTimeSetOff.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DteTimeSetOff.Location = New System.Drawing.Point(128, 27)
        Me.DteTimeSetOff.Name = "DteTimeSetOff"
        Me.DteTimeSetOff.Size = New System.Drawing.Size(112, 20)
        Me.DteTimeSetOff.TabIndex = 2
        Me.DteTimeSetOff.Value = New Date(2007, 1, 30, 13, 37, 15, 502)
        '
        'gBoxCollOpts
        '
        Me.gBoxCollOpts.Controls.AddRange(New System.Windows.Forms.Control() {Me.rButCallOut, Me.rButRoutine})
        Me.gBoxCollOpts.Location = New System.Drawing.Point(160, 16)
        Me.gBoxCollOpts.Name = "gBoxCollOpts"
        Me.gBoxCollOpts.Size = New System.Drawing.Size(128, 72)
        Me.gBoxCollOpts.TabIndex = 3
        Me.gBoxCollOpts.TabStop = False
        Me.gBoxCollOpts.Text = "Collection Type"
        '
        'rButCallOut
        '
        Me.rButCallOut.Enabled = False
        Me.rButCallOut.Location = New System.Drawing.Point(16, 40)
        Me.rButCallOut.Name = "rButCallOut"
        Me.rButCallOut.Size = New System.Drawing.Size(80, 16)
        Me.rButCallOut.TabIndex = 1
        Me.rButCallOut.Text = "CALL-OUT"
        '
        'rButRoutine
        '
        Me.rButRoutine.Enabled = False
        Me.rButRoutine.Location = New System.Drawing.Point(16, 16)
        Me.rButRoutine.Name = "rButRoutine"
        Me.rButRoutine.Size = New System.Drawing.Size(80, 24)
        Me.rButRoutine.TabIndex = 0
        Me.rButRoutine.Text = "ROUTINE"
        '
        'gBoxPayOpts
        '
        Me.gBoxPayOpts.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbWorkingTime, Me.rbCollection, Me.rbOther, Me.rButHourly, Me.rButSamples})
        Me.gBoxPayOpts.Location = New System.Drawing.Point(16, 16)
        Me.gBoxPayOpts.Name = "gBoxPayOpts"
        Me.gBoxPayOpts.Size = New System.Drawing.Size(136, 152)
        Me.gBoxPayOpts.TabIndex = 2
        Me.gBoxPayOpts.TabStop = False
        Me.gBoxPayOpts.Text = "Payment Options"
        '
        'rbWorkingTime
        '
        Me.rbWorkingTime.Location = New System.Drawing.Point(16, 120)
        Me.rbWorkingTime.Name = "rbWorkingTime"
        Me.rbWorkingTime.TabIndex = 4
        Me.rbWorkingTime.Text = "Working Time"
        '
        'rbCollection
        '
        Me.rbCollection.Location = New System.Drawing.Point(16, 88)
        Me.rbCollection.Name = "rbCollection"
        Me.rbCollection.TabIndex = 3
        Me.rbCollection.Text = "Collection Rate"
        '
        'rbOther
        '
        Me.rbOther.Location = New System.Drawing.Point(16, 63)
        Me.rbOther.Name = "rbOther"
        Me.rbOther.Size = New System.Drawing.Size(72, 24)
        Me.rbOther.TabIndex = 2
        Me.rbOther.Text = "Day Rate"
        '
        'rButHourly
        '
        Me.rButHourly.Location = New System.Drawing.Point(16, 40)
        Me.rButHourly.Name = "rButHourly"
        Me.rButHourly.Size = New System.Drawing.Size(72, 24)
        Me.rButHourly.TabIndex = 1
        Me.rButHourly.Text = "Hourly"
        '
        'rButSamples
        '
        Me.rButSamples.Location = New System.Drawing.Point(16, 16)
        Me.rButSamples.Name = "rButSamples"
        Me.rButSamples.Size = New System.Drawing.Size(72, 24)
        Me.rButSamples.TabIndex = 0
        Me.rButSamples.Text = "Sample"
        '
        'TimeSumBox
        '
        Me.TimeSumBox.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Label1, Me.txtOutHours, Me.txtTotalInHours, Me.lblHoursWorked, Me.txtHoursWorked})
        Me.TimeSumBox.Location = New System.Drawing.Point(16, 168)
        Me.TimeSumBox.Name = "TimeSumBox"
        Me.TimeSumBox.Size = New System.Drawing.Size(128, 104)
        Me.TimeSumBox.TabIndex = 8
        Me.TimeSumBox.TabStop = False
        Me.TimeSumBox.Text = "Time Summary"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 24)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Out-Hrs"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "In-Hrs:"
        '
        'txtOutHours
        '
        Me.txtOutHours.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtOutHours.Enabled = False
        Me.txtOutHours.Location = New System.Drawing.Point(80, 40)
        Me.txtOutHours.Name = "txtOutHours"
        Me.txtOutHours.Size = New System.Drawing.Size(40, 20)
        Me.txtOutHours.TabIndex = 2
        Me.txtOutHours.Text = "0"
        '
        'txtTotalInHours
        '
        Me.txtTotalInHours.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtTotalInHours.Enabled = False
        Me.txtTotalInHours.Location = New System.Drawing.Point(80, 16)
        Me.txtTotalInHours.Name = "txtTotalInHours"
        Me.txtTotalInHours.Size = New System.Drawing.Size(40, 20)
        Me.txtTotalInHours.TabIndex = 1
        Me.txtTotalInHours.Text = "0"
        '
        'lblHoursWorked
        '
        Me.lblHoursWorked.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHoursWorked.Location = New System.Drawing.Point(8, 72)
        Me.lblHoursWorked.Name = "lblHoursWorked"
        Me.lblHoursWorked.Size = New System.Drawing.Size(56, 24)
        Me.lblHoursWorked.TabIndex = 0
        Me.lblHoursWorked.Text = "Total:"
        '
        'txtHoursWorked
        '
        Me.txtHoursWorked.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtHoursWorked.Enabled = False
        Me.txtHoursWorked.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHoursWorked.Location = New System.Drawing.Point(80, 72)
        Me.txtHoursWorked.Name = "txtHoursWorked"
        Me.txtHoursWorked.Size = New System.Drawing.Size(40, 20)
        Me.txtHoursWorked.TabIndex = 1
        Me.txtHoursWorked.Text = "0"
        '
        'grpPerSampleRate
        '
        Me.grpPerSampleRate.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblPerSample})
        Me.grpPerSampleRate.Location = New System.Drawing.Point(424, 216)
        Me.grpPerSampleRate.Name = "grpPerSampleRate"
        Me.grpPerSampleRate.Size = New System.Drawing.Size(344, 56)
        Me.grpPerSampleRate.TabIndex = 11
        Me.grpPerSampleRate.TabStop = False
        Me.grpPerSampleRate.Text = "Sample Payment"
        '
        'lblPerSample
        '
        Me.lblPerSample.Location = New System.Drawing.Point(8, 24)
        Me.lblPerSample.Name = "lblPerSample"
        Me.lblPerSample.Size = New System.Drawing.Size(320, 16)
        Me.lblPerSample.TabIndex = 0
        '
        'panelCollPayDefaults
        '
        Me.panelCollPayDefaults.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.panelCollPayDefaults.Controls.AddRange(New System.Windows.Forms.Control() {Me.PanelSamplePlus, Me.lvOfficers, Me.cmdAddOfficer, Me.lblHalfDays, Me.udHalfdays, Me.dtTravelTime, Me.lblTravelTime, Me.DTWorkingTime, Me.lblStatus, Me.cbCurrency, Me.Label3, Me.cbCollOfficer, Me.dtPayMonth, Me.CollDate, Me.lblPaymentMonth, Me.MonthCalendar1, Me.txtCustomerID, Me.txtSitePort, Me.txtVesselName, Me.lblvessel, Me.lblCollDate, Me.lblSite, Me.lblCustomer, Me.lblCollOff, Me.Donors, Me.lblDonors, Me.lblCancSamples})
        Me.panelCollPayDefaults.Location = New System.Drawing.Point(8, 0)
        Me.panelCollPayDefaults.Name = "panelCollPayDefaults"
        Me.panelCollPayDefaults.Size = New System.Drawing.Size(828, 184)
        Me.panelCollPayDefaults.TabIndex = 13
        '
        'PanelSamplePlus
        '
        Me.PanelSamplePlus.Controls.AddRange(New System.Windows.Forms.Control() {Me.updLineitem})
        Me.PanelSamplePlus.Location = New System.Drawing.Point(416, 120)
        Me.PanelSamplePlus.Name = "PanelSamplePlus"
        Me.PanelSamplePlus.Size = New System.Drawing.Size(336, 40)
        Me.PanelSamplePlus.TabIndex = 27
        '
        'updLineitem
        '
        Me.updLineitem.Location = New System.Drawing.Point(16, 0)
        Me.updLineitem.Name = "updLineitem"
        Me.updLineitem.TabIndex = 0
        Me.updLineitem.Text = "DomainUpDown1"
        '
        'lvOfficers
        '
        Me.lvOfficers.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lvOfficers.LargeImageList = Me.ImageList1
        Me.lvOfficers.Location = New System.Drawing.Point(392, 8)
        Me.lvOfficers.Name = "lvOfficers"
        Me.lvOfficers.Size = New System.Drawing.Size(408, 72)
        Me.lvOfficers.SmallImageList = Me.ImageList1
        Me.lvOfficers.TabIndex = 26
        Me.lvOfficers.View = System.Windows.Forms.View.SmallIcon
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'cmdAddOfficer
        '
        Me.cmdAddOfficer.Image = CType(resources.GetObject("cmdAddOfficer.Image"), System.Drawing.Bitmap)
        Me.cmdAddOfficer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddOfficer.Location = New System.Drawing.Point(288, 24)
        Me.cmdAddOfficer.Name = "cmdAddOfficer"
        Me.cmdAddOfficer.Size = New System.Drawing.Size(96, 24)
        Me.cmdAddOfficer.TabIndex = 25
        Me.cmdAddOfficer.Text = "Add Officer"
        Me.cmdAddOfficer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblHalfDays
        '
        Me.lblHalfDays.AutoSize = True
        Me.lblHalfDays.Location = New System.Drawing.Point(416, 121)
        Me.lblHalfDays.Name = "lblHalfDays"
        Me.lblHalfDays.Size = New System.Drawing.Size(56, 13)
        Me.lblHalfDays.TabIndex = 24
        Me.lblHalfDays.Text = "Half Days:"
        Me.lblHalfDays.Visible = False
        '
        'udHalfdays
        '
        Me.udHalfdays.Location = New System.Drawing.Point(480, 120)
        Me.udHalfdays.Maximum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.udHalfdays.Name = "udHalfdays"
        Me.udHalfdays.Size = New System.Drawing.Size(40, 20)
        Me.udHalfdays.TabIndex = 23
        Me.udHalfdays.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.udHalfdays.Visible = False
        '
        'dtTravelTime
        '
        Me.dtTravelTime.CustomFormat = "hh:mm"
        Me.dtTravelTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTravelTime.Location = New System.Drawing.Point(600, 120)
        Me.dtTravelTime.Name = "dtTravelTime"
        Me.dtTravelTime.Size = New System.Drawing.Size(72, 20)
        Me.dtTravelTime.TabIndex = 22
        Me.dtTravelTime.Tag = "PAYTRVTIME"
        '
        'lblTravelTime
        '
        Me.lblTravelTime.AutoSize = True
        Me.lblTravelTime.Location = New System.Drawing.Point(520, 120)
        Me.lblTravelTime.Name = "lblTravelTime"
        Me.lblTravelTime.Size = New System.Drawing.Size(61, 13)
        Me.lblTravelTime.TabIndex = 21
        Me.lblTravelTime.Text = "TravelTime"
        '
        'DTWorkingTime
        '
        Me.DTWorkingTime.CustomFormat = "hh:mm"
        Me.DTWorkingTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTWorkingTime.Location = New System.Drawing.Point(432, 120)
        Me.DTWorkingTime.Name = "DTWorkingTime"
        Me.DTWorkingTime.Size = New System.Drawing.Size(72, 20)
        Me.DTWorkingTime.TabIndex = 20
        Me.DTWorkingTime.Tag = "PAYWRKTIME"
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(296, 152)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(208, 23)
        Me.lblStatus.TabIndex = 19
        Me.lblStatus.Text = "Status:"
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(120, 152)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(121, 21)
        Me.cbCurrency.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 24)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Currency:"
        '
        'cbCollOfficer
        '
        Me.cbCollOfficer.Location = New System.Drawing.Point(104, 24)
        Me.cbCollOfficer.Name = "cbCollOfficer"
        Me.cbCollOfficer.Size = New System.Drawing.Size(168, 21)
        Me.cbCollOfficer.TabIndex = 16
        '
        'dtPayMonth
        '
        Me.dtPayMonth.CustomFormat = "MMMM-yyyy"
        Me.dtPayMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPayMonth.Location = New System.Drawing.Point(144, 120)
        Me.dtPayMonth.Name = "dtPayMonth"
        Me.dtPayMonth.Size = New System.Drawing.Size(112, 20)
        Me.dtPayMonth.TabIndex = 1
        '
        'CollDate
        '
        Me.CollDate.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.CollDate.CalendarMonthBackground = System.Drawing.SystemColors.InactiveBorder
        Me.CollDate.CalendarTitleForeColor = System.Drawing.SystemColors.ActiveBorder
        Me.CollDate.CustomFormat = "dd-MMMMMMMMMM-yy"
        Me.CollDate.Enabled = False
        Me.CollDate.Location = New System.Drawing.Point(616, 88)
        Me.CollDate.Name = "CollDate"
        Me.CollDate.Size = New System.Drawing.Size(160, 20)
        Me.CollDate.TabIndex = 15
        '
        'lblPaymentMonth
        '
        Me.lblPaymentMonth.Location = New System.Drawing.Point(16, 120)
        Me.lblPaymentMonth.Name = "lblPaymentMonth"
        Me.lblPaymentMonth.Size = New System.Drawing.Size(96, 24)
        Me.lblPaymentMonth.TabIndex = 13
        Me.lblPaymentMonth.Text = "Payment Month:"
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.MonthCalendar1.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.MonthCalendar1.Location = New System.Drawing.Point(112, 112)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 12
        Me.MonthCalendar1.TitleBackColor = System.Drawing.SystemColors.InactiveBorder
        Me.MonthCalendar1.TitleForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.MonthCalendar1.Visible = False
        '
        'txtCustomerID
        '
        Me.txtCustomerID.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.txtCustomerID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCustomerID.Location = New System.Drawing.Point(112, 56)
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.Size = New System.Drawing.Size(120, 13)
        Me.txtCustomerID.TabIndex = 10
        Me.txtCustomerID.Text = ""
        '
        'txtSitePort
        '
        Me.txtSitePort.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.txtSitePort.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtSitePort.Location = New System.Drawing.Point(112, 88)
        Me.txtSitePort.Name = "txtSitePort"
        Me.txtSitePort.Size = New System.Drawing.Size(120, 13)
        Me.txtSitePort.TabIndex = 9
        Me.txtSitePort.Text = ""
        '
        'txtVesselName
        '
        Me.txtVesselName.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.txtVesselName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtVesselName.Location = New System.Drawing.Point(296, 88)
        Me.txtVesselName.Name = "txtVesselName"
        Me.txtVesselName.Size = New System.Drawing.Size(224, 13)
        Me.txtVesselName.TabIndex = 7
        Me.txtVesselName.Text = ""
        '
        'lblvessel
        '
        Me.lblvessel.Location = New System.Drawing.Point(240, 80)
        Me.lblvessel.Name = "lblvessel"
        Me.lblvessel.Size = New System.Drawing.Size(56, 23)
        Me.lblvessel.TabIndex = 5
        Me.lblvessel.Text = "Vessel:"
        '
        'lblCollDate
        '
        Me.lblCollDate.Location = New System.Drawing.Point(520, 88)
        Me.lblCollDate.Name = "lblCollDate"
        Me.lblCollDate.Size = New System.Drawing.Size(88, 24)
        Me.lblCollDate.TabIndex = 4
        Me.lblCollDate.Text = "Collection Date:"
        '
        'lblSite
        '
        Me.lblSite.Location = New System.Drawing.Point(16, 88)
        Me.lblSite.Name = "lblSite"
        Me.lblSite.Size = New System.Drawing.Size(96, 24)
        Me.lblSite.TabIndex = 2
        Me.lblSite.Text = "Site/Port Name"
        '
        'lblCustomer
        '
        Me.lblCustomer.Location = New System.Drawing.Point(16, 56)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.Size = New System.Drawing.Size(88, 24)
        Me.lblCustomer.TabIndex = 1
        Me.lblCustomer.Text = "Customer:"
        '
        'lblCollOff
        '
        Me.lblCollOff.Location = New System.Drawing.Point(16, 24)
        Me.lblCollOff.Name = "lblCollOff"
        Me.lblCollOff.Size = New System.Drawing.Size(96, 24)
        Me.lblCollOff.TabIndex = 0
        Me.lblCollOff.Text = "Collecting Officer:"
        '
        'Donors
        '
        Me.Donors.Location = New System.Drawing.Point(368, 120)
        Me.Donors.Name = "Donors"
        Me.Donors.Size = New System.Drawing.Size(40, 20)
        Me.Donors.TabIndex = 1
        Me.Donors.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblDonors
        '
        Me.lblDonors.AutoSize = True
        Me.lblDonors.Location = New System.Drawing.Point(312, 120)
        Me.lblDonors.Name = "lblDonors"
        Me.lblDonors.Size = New System.Drawing.Size(51, 13)
        Me.lblDonors.TabIndex = 0
        Me.lblDonors.Text = "Samples:"
        '
        'lblCancSamples
        '
        Me.lblCancSamples.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCancSamples.ForeColor = System.Drawing.Color.Red
        Me.lblCancSamples.Location = New System.Drawing.Point(424, 120)
        Me.lblCancSamples.Name = "lblCancSamples"
        Me.lblCancSamples.Size = New System.Drawing.Size(144, 16)
        Me.lblCancSamples.TabIndex = 5
        Me.lblCancSamples.Text = "Canc Samples:"
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        'Me.HtmlEditor1.DefaultComposeSettings.BackColor = System.Drawing.Color.White
        'Me.HtmlEditor1.DefaultComposeSettings.DefaultFont = New System.Drawing.Font("Arial", 10.0!)
        'Me.HtmlEditor1.DefaultComposeSettings.Enabled = False
        'Me.HtmlEditor1.DefaultComposeSettings.ForeColor = System.Drawing.Color.Black
        'Me.HtmlEditor1.DocumentEncoding = onlyconnect.EncodingType.WindowsCurrent
        Me.HtmlEditor1.Enabled = False
        Me.HtmlEditor1.Location = New System.Drawing.Point(16, 216)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        'Me.HtmlEditor1.OpenLinksInNewWindow = True
        'Me.HtmlEditor1.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left
        'Me.HtmlEditor1.SelectionBackColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionBullets = False
        'Me.HtmlEditor1.SelectionFont = Nothing
        'Me.HtmlEditor1.SelectionForeColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionNumbering = False
        Me.HtmlEditor1.Size = New System.Drawing.Size(792, 384)
        Me.HtmlEditor1.TabIndex = 16
        'Me.HtmlEditor1.Text = "HtmlEditor1"
        Me.HtmlEditor1.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label6, Me.txtGrandTotal, Me.cmdRefresh, Me.cmdAddPay, Me.lblCostsSubTotal, Me.lblExpSubTotal, Me.cmdRemove, Me.txtPayAmount, Me.lvPayItems, Me.txtExpSubTotal, Me.cmdAdd, Me.ListView1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(816, 522)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Collection Expenses"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(376, 424)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 24)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Grand Total:"
        '
        'txtGrandTotal
        '
        Me.txtGrandTotal.Location = New System.Drawing.Point(480, 424)
        Me.txtGrandTotal.Name = "txtGrandTotal"
        Me.txtGrandTotal.TabIndex = 10
        Me.txtGrandTotal.Text = ""
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Image = CType(resources.GetObject("cmdRefresh.Image"), System.Drawing.Bitmap)
        Me.cmdRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRefresh.Location = New System.Drawing.Point(128, 392)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(112, 24)
        Me.cmdRefresh.TabIndex = 9
        Me.cmdRefresh.Text = "Refresh"
        '
        'cmdAddPay
        '
        Me.cmdAddPay.Image = CType(resources.GetObject("cmdAddPay.Image"), System.Drawing.Bitmap)
        Me.cmdAddPay.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddPay.Location = New System.Drawing.Point(16, 392)
        Me.cmdAddPay.Name = "cmdAddPay"
        Me.cmdAddPay.Size = New System.Drawing.Size(96, 24)
        Me.cmdAddPay.TabIndex = 8
        Me.cmdAddPay.Text = "Add"
        '
        'lblCostsSubTotal
        '
        Me.lblCostsSubTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCostsSubTotal.Location = New System.Drawing.Point(376, 392)
        Me.lblCostsSubTotal.Name = "lblCostsSubTotal"
        Me.lblCostsSubTotal.Size = New System.Drawing.Size(96, 24)
        Me.lblCostsSubTotal.TabIndex = 7
        Me.lblCostsSubTotal.Text = "Costs Sub Total:"
        '
        'lblExpSubTotal
        '
        Me.lblExpSubTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpSubTotal.Location = New System.Drawing.Point(336, 184)
        Me.lblExpSubTotal.Name = "lblExpSubTotal"
        Me.lblExpSubTotal.Size = New System.Drawing.Size(112, 24)
        Me.lblExpSubTotal.TabIndex = 6
        Me.lblExpSubTotal.Text = "Expense Sub Total:"
        '
        'cmdRemove
        '
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.Image = CType(resources.GetObject("cmdRemove.Image"), System.Drawing.Bitmap)
        Me.cmdRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRemove.Location = New System.Drawing.Point(128, 184)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(112, 24)
        Me.cmdRemove.TabIndex = 5
        Me.cmdRemove.Text = "Remove"
        '
        'txtPayAmount
        '
        Me.txtPayAmount.Location = New System.Drawing.Point(480, 392)
        Me.txtPayAmount.Name = "txtPayAmount"
        Me.txtPayAmount.TabIndex = 4
        Me.txtPayAmount.Text = ""
        '
        'lvPayItems
        '
        Me.lvPayItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.chDisc, Me.ColumnHeader10, Me.ColumnHeader9, Me.ColumnHeader1, Me.chCommentP})
        Me.lvPayItems.ContextMenu = Me.ContextMenu1
        Me.lvPayItems.FullRowSelect = True
        Me.lvPayItems.Location = New System.Drawing.Point(0, 232)
        Me.lvPayItems.Name = "lvPayItems"
        Me.lvPayItems.Size = New System.Drawing.Size(672, 152)
        Me.lvPayItems.TabIndex = 3
        Me.lvPayItems.Tag = "PAY"
        Me.lvPayItems.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Cost Type"
        Me.ColumnHeader5.Width = 0
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Quantity"
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Description"
        Me.ColumnHeader7.Width = 200
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "ListPay Rate"
        Me.ColumnHeader8.Width = 118
        '
        'chDisc
        '
        Me.chDisc.Text = "Discount"
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Act Pay Rate"
        Me.ColumnHeader10.Width = 118
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Pay Amount"
        Me.ColumnHeader9.Width = 120
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Status"
        '
        'chCommentP
        '
        Me.chCommentP.Text = "Comment"
        Me.chCommentP.Width = 200
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuApproveItem, Me.mnuDeclineItem, Me.MnuReferItem, Me.mnuCancelItem, Me.mnuEntered, Me.mnuEditItem})
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
        'mnuEditItem
        '
        Me.mnuEditItem.Index = 5
        Me.mnuEditItem.Text = "Ed&it Item"
        '
        'txtExpSubTotal
        '
        Me.txtExpSubTotal.Location = New System.Drawing.Point(456, 184)
        Me.txtExpSubTotal.Name = "txtExpSubTotal"
        Me.txtExpSubTotal.TabIndex = 2
        Me.txtExpSubTotal.Text = ""
        '
        'cmdAdd
        '
        Me.cmdAdd.Image = CType(resources.GetObject("cmdAdd.Image"), System.Drawing.Bitmap)
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(16, 184)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(96, 24)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "Add"
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.chPrice, Me.chDiscount, Me.chVALUE, Me.chStatus, Me.chComment})
        Me.ListView1.ContextMenu = Me.ContextMenu1
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ListView1.FullRowSelect = True
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(816, 168)
        Me.ListView1.TabIndex = 1
        Me.ListView1.Tag = "EXP"
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Expense Type"
        Me.ColumnHeader2.Width = 0
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Quantity"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Description"
        Me.ColumnHeader4.Width = 200
        '
        'chPrice
        '
        Me.chPrice.Text = "Price"
        Me.chPrice.Width = 118
        '
        'chDiscount
        '
        Me.chDiscount.Text = "Discount"
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
        'chComment
        '
        Me.chComment.Text = "Comment"
        Me.chComment.Width = 200
        '
        'txtCMidentity
        '
        Me.txtCMidentity.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.txtCMidentity.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCMidentity.Location = New System.Drawing.Point(560, 8)
        Me.txtCMidentity.Name = "txtCMidentity"
        Me.txtCMidentity.Size = New System.Drawing.Size(120, 13)
        Me.txtCMidentity.TabIndex = 8
        Me.txtCMidentity.Text = ""
        '
        'lblCMIdentity
        '
        Me.lblCMIdentity.Location = New System.Drawing.Point(472, 8)
        Me.lblCMIdentity.Name = "lblCMIdentity"
        Me.lblCMIdentity.Size = New System.Drawing.Size(88, 24)
        Me.lblCMIdentity.TabIndex = 3
        Me.lblCMIdentity.Text = "CM Identity:"
        '
        'lblUserInfo
        '
        Me.lblUserInfo.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserInfo.Location = New System.Drawing.Point(8, 8)
        Me.lblUserInfo.Name = "lblUserInfo"
        Me.lblUserInfo.Size = New System.Drawing.Size(456, 16)
        Me.lblUserInfo.TabIndex = 16
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPay, Me.mnuReports})
        '
        'mnuPay
        '
        Me.mnuPay.Index = 0
        Me.mnuPay.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuLockPay, Me.mnuUnlockPay, Me.mnuApprovePay})
        Me.mnuPay.Text = "Pay Actions"
        '
        'mnuLockPay
        '
        Me.mnuLockPay.Index = 0
        Me.mnuLockPay.Text = "Lock Pay"
        '
        'mnuUnlockPay
        '
        Me.mnuUnlockPay.Index = 1
        Me.mnuUnlockPay.Text = "&Unlock Pay"
        '
        'mnuApprovePay
        '
        Me.mnuApprovePay.Index = 2
        Me.mnuApprovePay.Text = "&Approve Pay"
        '
        'mnuReports
        '
        Me.mnuReports.Index = 1
        Me.mnuReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPaystatement})
        Me.mnuReports.Text = "Reports"
        '
        'mnuPaystatement
        '
        Me.mnuPaystatement.Index = 0
        Me.mnuPaystatement.Text = "Pay Statement"
        '
        'frmCollPayEntry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(824, 625)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.Panel1, Me.lblUserInfo, Me.txtCMidentity, Me.lblCMIdentity})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmCollPayEntry"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.PaySummBox.ResumeLayout(False)
        Me.gBoxTimeEntry.ResumeLayout(False)
        CType(Me.udMileage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gBoxCollOpts.ResumeLayout(False)
        Me.gBoxPayOpts.ResumeLayout(False)
        Me.TimeSumBox.ResumeLayout(False)
        Me.grpPerSampleRate.ResumeLayout(False)
        Me.panelCollPayDefaults.ResumeLayout(False)
        Me.PanelSamplePlus.ResumeLayout(False)
        CType(Me.udHalfdays, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Donors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private MyCollPayEntry As Intranet.collectorpay.CollectorPay

    Public Property CollPay() As Intranet.collectorpay.CollectorPay
        Get
            Return Me.MyCollPayEntry
        End Get
        Set(ByVal Value As Intranet.collectorpay.CollectorPay)
            MyCollPayEntry = Value
            If MyCollPayEntry Is Nothing Then Exit Property
            objCollOfficer = New Intranet.intranet.jobs.CollectingOfficer(MyCollPayEntry.CollOfficer)
            objCollOfficer.Load()
            FillForm()
        End Set

    End Property


    Private Function HoursMinutesWorked() As System.TimeSpan

        Dim dteHoursMinWorked As System.TimeSpan

        dteHoursMinWorked = Me.DteTimeArrHome.Value.Subtract(Me.DteTimeSetOff.Value)

        Return dteHoursMinWorked
    End Function

    Public Function UpdateUserInfo(ByVal UserInfoComm)
        If Me.lblUserInfo.Text = "" Then
            Me.lblUserInfo.Text += UserInfoComm
        Else
            Me.lblUserInfo.Text += "; " & UserInfoComm
        End If
    End Function


    Private Sub FillForm()
        blnDrawing = True
        If MyCollPayEntry Is Nothing Then
            UpdateUserInfo("New Collector Pay Entry.")
            Exit Sub
        End If
        NewForm = True
        With MyCollPayEntry

            LockControls((Me.CollPay.PayStatus = "E"), Me.CollPay.PayStatus)
            If Not objCollOfficer Is Nothing Then
                Me.cbCollOfficer.SelectedValue = objCollOfficer.OfficerID
                'Me.cbCollOfficer.Text = objCollOfficer.Surname & objCollOfficer.FirstName
                '<Added code modified 16-Jul-2007 07:54 by taylor> 
            End If

            If .CMidentity.Trim.Length > 0 Then
                objCM = New Intranet.intranet.jobs.CMJob(.CMidentity)
                If Not objCM Is Nothing Then
                    '<Added code modified 17-Jul-2007 08:16 by taylor> 
                    If Me.objCollOfficer Is Nothing Then objCollOfficer = objCollOfficer
                    '</Added code modified 17-Jul-2007 08:16 by taylor> 


                    If Not objCollOfficer Is Nothing Then
                        'Me.cbCollOfficer.Text = objCollOfficer.Surname & objCollOfficer.FirstName
                        '<Added code modified 16-Jul-2007 07:54 by taylor> 


                        If .PayMethod.Trim.Length = 0 Then .PayMethod = objCollOfficer.PayMethod
                        '</Added code modified 16-Jul-2007 07:54 by taylor> 
                        'Check country information 
                        BuildItems()

                        If objCollOfficer.CountryId = 44 Then
                            Me.udKMMiles.SelectedIndex = 0
                        Else
                            Me.udKMMiles.SelectedIndex = 1
                        End If
                    End If

                    Select Case .PayMethod
                        Case "S"
                            Me.rButSamples.Checked = True
                            Me.rButSamples.Refresh()
                            Me.TimeSumBox.Enabled = False
                            Me.PaySummBox.Enabled = True
                        Case "H"
                            Me.rButHourly.Checked = True
                            Me.rButHourly.Checked = True
                            Me.rButHourly.Refresh()
                            Me.TimeSumBox.Enabled = True
                            Me.PaySummBox.Enabled = True
                        Case "O"
                            Me.rbOther.Checked = True
                            Me.rbOther.Checked = True
                            Me.rbOther.Refresh()
                            Me.TimeSumBox.Enabled = False
                            Me.PaySummBox.Enabled = True
                        Case "C"
                            Me.rbCollection.Checked = True
                        Case "W"
                            Me.rbWorkingTime.Checked = True
                    End Select

                    'Deal with ignoring discounts 
                    Me.ckIgnoreMinimum.Checked = .NoMinimum

                    'Do travel values 
                    If Not CollPay Is Nothing AndAlso Me.CollPay.PayItems.IsMileagePresent Then
                        MyMileageCost = CollPay.PayItems.MileagePay()
                        If Not MyMileageCost Is Nothing Then
                            Me.txtMileage.Text = MyMileageCost.ItemValue.ToString("c", Me.objCollOfficer.Culture)
                            If MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerMile Then
                                Me.udKMMiles.SelectedIndex = 0
                            Else
                                Me.udKMMiles.SelectedIndex = 1
                            End If
                            Me.udMileage.Value = Me.MyMileageCost.Quantity
                        End If
                    Else
                        Me.udMileage.Value = 0 'zero out mileage value
                    End If
                    If MyCollPayEntry.TimeLeftHome = DateField.ZeroDate Then
                        MyCollPayEntry.TimeLeftHome = objCM.CollectionDate
                        MyCollPayEntry.TimeLeftSite = MyCollPayEntry.TimeLeftHome
                        MyCollPayEntry.TimeArrHome = MyCollPayEntry.TimeLeftHome
                        MyCollPayEntry.TimeArrOnSite = MyCollPayEntry.TimeLeftHome

                    End If
                    'see abou the comment 
                    If .Comment.Trim.Length = 0 Then
                        .Comment = CConnection.PackageStringList("Lib_collection.GetCollectionBrowseText", objCM.ID)
                    End If
                    If Not IsUkCollection() Then ' Remove UK Fields leaving O/S format.                     
                        Me.lblPerSample.Text = objCollOfficer.SampleRate
                        'objCollOfficer.Update()
                        UpdateUserInfo("Overseas Collection")
                    Else
                        UpdateUserInfo("UK Collection")
                    End If
                    Me.cbCollOfficer.SelectedValue = objCollOfficer.OfficerID
                    If .Currency.Trim.Length = 0 Then .Currency = objCollOfficer.PayCurrency
                    Me.cbCurrency.SelectedValue = .Currency
                    If Not objCM.SmClient Is Nothing Then

                        Me.txtCustomerID.Text = objCM.SmClient.SMID & "-" _
                                              & objCM.SmClient.Profile
                        UpdateUserInfo(objCM.SmClient.SMID & "-" & objCM.SmClient.Profile & "; " & objCM.VesselName)
                    End If

                    Me.txtVesselName.Text = objCM.VesselName
                    'If MyCollPayEntry.PayItems.IsStandardPayItemPresent Then
                    'myCollPayInHoursCost = MyCollPayEntry.PayItems.StandardPayItem
                    'Else
                    'MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_Std_Sample_Pay)
                    'End If
                    Me.DTWorkingTime.Value = Today
                    Me.dtTravelTime.Value = Today
                    'See if we have the number of samples saved
                    If Me.CollPay.PayMethod = "S" Or Me.CollPay.PayMethod = "H" Then
                        Me.Donors.Value = Me.CollPay.NoSamples
                        Me.Donors.Value = Me.CollPay.NoSamples
                    End If
                    If MyCollPayEntry.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODSample Then
                        If MyCollPayInHoursCost Is Nothing Then
                            MyCollPayInHoursCost = CollPay.PayItems.Item("PAYROUNTINHRS001")
                        End If
                        If MyCollPayInHoursCost Is Nothing Then
                            MyCollPayInHoursCost = CollPay.PayItems.Item("PAYSAMPLE001")
                        End If

                        MyCollPayInHoursCost.ExpenseTypeID = SampleRatePayType()
                        If Not MyCollPayInHoursCost Is Nothing And Me.CollPay.NoSamples = 0 Then
                            Dim strNSamples As String = CConnection.PackageStringList("lib_collpay.NoSamples", MyCollPayEntry.CollPayID)
                            If MyCollPayInHoursCost.Quantity = 0 AndAlso Not objCM Is Nothing Then MyCollPayInHoursCost.Quantity = objCM.NoReceived
                            Me.Donors.Value = strNSamples
                        Else
                            Dim strNSamples As String = CConnection.PackageStringList("lib_collpay.NoSamples", MyCollPayEntry.CollPayID)

                            Me.Donors.Value = CInt(strnsamples)
                        End If
                    ElseIf MyCollPayEntry.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODHourly Then
                        MyCollPayInHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_Routine_In_hours

                    ElseIf MyCollPayEntry.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODDay Then
                        MyCollPayInHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerDAY
                        'Me.Donors.Value = 1
                    ElseIf MyCollPayEntry.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODWorkingTime Then
                        MyCollPayInHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayWorkingTime
                        Me.Donors.Value = 0
                    ElseIf MyCollPayEntry.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODCollection Then
                        MyCollPayInHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_Fixed_Price_Charge
                        Me.Donors.Value = 1

                    End If
                    Me.CollDate.Text = objCM.CollectionDate
                    Me.txtSitePort.Text = objCM.SiteName

                    Me.DteTimeSetOff.Value = MyCollPayEntry.TimeLeftHome
                    Me.DteTimeArrOnSite.Value = MyCollPayEntry.TimeArrOnSite
                    Me.DteTimeLeftSite.Value = MyCollPayEntry.TimeLeftSite
                    Me.DteTimeArrHome.Value = MyCollPayEntry.TimeArrHome
                    Dim strTos As String = CConnection.PackageStringList("Lib_collpay.TimeOnSite", Me.CollPay.CollPayID)
                    Me.lblTimeOnSite.Text = "Time On Site : " & strTos

                End If

                ChkCancelledSamples() 'Checks for and display number of cancelled sample.

                SetCollectorSampleRate() 'Find if the c/o has a per sample collection rate.



                If .PaymentMonth <> DateField.ZeroDate Then
                    Me.dtPayMonth.Value = DateSerial(.PaymentMonth.Year, .PaymentMonth.Month, 1)
                Else
                    Me.dtPayMonth.Value = DateSerial(Now.Year, Now.Month, 1)
                End If

                Select Case objCM.CollectionType ' format collection type section of form.
                    Case Is = "C", "O"
                        Me.rButRoutine.Checked = False
                        Me.rButCallOut.Checked = True
                        Me.rButRoutine.Refresh()
                        Me.rButCallOut.Refresh()
                    Case Else
                        Me.rButRoutine.Checked = True
                        Me.rButCallOut.Checked = False
                        Me.rButRoutine.Refresh()
                        Me.rButCallOut.Refresh()
                End Select


                Me.txtCMidentity.Text = objCM.ID

                SetDateTimeFields() ' set collection time fields


                If MyCollPayEntry Is Nothing Then Exit Sub
                UpDateCollPayItem()
            End If
            Me.txtComment.Text = .Comment
        End With
        'See if we need to do pay items

        UpdateCollPayDetails = False
        'Update status info
        SetStatus()
        DrawOfficerList()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill officer list view
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	15/08/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawOfficerList()
        Me.lvOfficers.Items.Clear()
        If objCM Is Nothing Then Exit Sub
        Dim strOfficers As String = CConnection.PackageStringList("Lib_Collpay.GetOfficers", objCM.ID)
        Dim strOffArray As String() = strOfficers.Split(New Char() {","})
        Dim i As Integer
        For i = 0 To strOffArray.Length - 2
            Dim strOff As String = strOffArray.GetValue(i)
            Dim strOffText As String = CConnection.PackageStringList("Lib_collpay.OfficerName", strOff)
            Dim lvItem As New ListViewItem(strOffText)
            lvItem.Tag = strOff
            Me.lvOfficers.Items.Add(lvItem)
        Next
    End Sub

    Private Sub RefreshCollCostListView()
        Dim objCollExpItem As Intranet.collectorpay.CollPayItem
        ListView1.Items.Clear()
        lvPayItems.Items.Clear()

        'Check we have a collecting Officer
        If Me.objCollOfficer Is Nothing Then Exit Sub

        For Each objCollExpItem In MyCollPayEntry.PayItems ' loop through expense items, populate the correct listview dep on paytype.
            'Ensure the appropriate cost item is displayed to match the select costs type.
            If objCollExpItem.IsExpenseItem(objCollExpItem.ExpenseTypeID) Then
                If Not objCollExpItem.Removed Then
                    Dim lvItem As New ExpenseDescription(objCollExpItem, objCollOfficer)
                    Me.ListView1.Items.Add(lvItem)
                End If
            ElseIf MyCollPayEntry.PaymentType = CollectorPayCollection.PayTypes.COLLECTION Or MyCollPayEntry.PaymentType = CollectorPayCollection.PayTypes.ONCALLROTA Then
                If Not objCollExpItem.Removed Then
                    Dim lvItem As New ExpenseDescription(objCollExpItem, objCollOfficer)
                    Me.lvPayItems.Items.Add(lvItem)
                End If
            ElseIf objCollExpItem.IsSampleTypeExpenseId(objCollExpItem.ExpenseTypeID) Then
                If Not objCollExpItem.Removed And Me.rButSamples.Checked Then
                    Dim lvItem As New ExpenseDescription(objCollExpItem, objCollOfficer)
                    Me.lvPayItems.Items.Add(lvItem)
                End If
            End If

        Next
        NewForm = False

    End Sub
    ''' <summary>
    ''' Created by tonyw on ANDREW at 26/02/2007 16:04:01
    '''     
    ''' </summary> Method for updating the secondary collection cost object. This will include the out
    ''' of hours costs and WILL NOT include any per sample related pay rates or in-hour payments.
    ''' <param name="CollPayOutOfHoursCost" type="Intranet.collectorpay.CollPayItem">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>tonyw</Author><date> 26/02/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 



    Private Sub DrawPayItems(ByVal MyCollPayExp)

        If MyCollPayExp Is Nothing Then Exit Sub
        Dim CostsTotal As Double = (Me.MyCollPayEntry.PayItems.TotalValue.ToString - Me.MyCollPayEntry.PayItems.ExpenseValue.ToString)
        Dim outHours As TimeSpan = Medscreen.HoursOutOfHours(Me.DteTimeSetOff.Value, Me.DteTimeArrHome.Value)

        ' Determine the total In/Out of hours values and display them
        Me.txtTotalInHours.Text = CDbl((HoursMinutesWorked.TotalMinutes - outHours.TotalMinutes) / 60).ToString("00.00")
        Me.txtOutHours.Text = CDbl(outHours.TotalMinutes / 60).ToString("00.00")
        Me.txtHoursWorked.Text = CDbl(HoursMinutesWorked.TotalMinutes / 60).ToString("00.00")

        'Populate the form objects on the expenses tab
        Me.txtExpSubTotal.Text = Me.MyCollPayEntry.PayItems.ExpenseValue.ToString("c", objCollOfficer.Culture)
        Me.txtPayAmount.Text = Me.MyCollPayEntry.PayItems.PayValue.ToString("c", objCollOfficer.Culture)
        Me.txtGrandTotal.Text = CDbl(Me.MyCollPayEntry.PayItems.PayValue + Me.MyCollPayEntry.PayItems.ExpenseValue).ToString("c", objCollOfficer.Culture)
        'Populate the form objects on the costs tab
        'Me.txtCostTotal.Text = CostsTotal.ToString("c", objCollOfficer.Culture)
        Me.txtTotalExp.Text = (Me.MyCollPayEntry.PayItems.ExpenseValue).ToString("c", objCollOfficer.Culture)
        'Subtract out any mileage cost
        'If Not MyMileageCost Is Nothing Then
        'Me.txtCostTotal.Text = (Me.MyCollPayEntry.PayItems.PayValue - MyMileageCost.ItemValue).ToString("c", objCollOfficer.Culture)
        'Else
        Me.txtCostTotal.Text = Me.CollPay.PayValue.ToString("c", objCollOfficer.Culture)

        'End If
        Me.txtCostsTotal.Text = Me.CollPay.TotalValue.ToString("c", objCollOfficer.Culture) 'Me.MyCollPayEntry.PayItems.TotalValue.ToString("c", objCollOfficer.Culture)

        Dim lvItem As ExpenseDescription
        For Each lvItem In Me.lvPayItems.Items
            lvItem.Refresh()
        Next

        If MyCollPayOutOfHoursCost Is Nothing Then Exit Sub

        If outHours.TotalMinutes > 0 Then 'We have an out of hour value 
            Me.MyCollPayOutOfHoursCost.Quantity = outHours.TotalMinutes / 60
            MyCollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
            MyCollPayOutOfHoursCost.ItemValue = objCollOfficer.PayForItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value) * Me.MyCollPayOutOfHoursCost.Quantity
            MyCollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
            MyCollPayOutOfHoursCost.DoUpdate()

        End If
    End Sub
    ''' <summary>
    ''' Created by tonyw on ANDREW at 19/03/2007 15:16:34
    '''     
    ''' </summary>
    ''' <param name="CollCostTotal" type="Object">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     A System.Object value...
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>tonyw</Author><date> 19/03/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Function ChkMinCollCarge(ByVal CollCostTotal)

        Dim MinCharge As Double = objCollOfficer.MinCharge

        If CollCostTotal = 0.0 Then Exit Function ' do not over-write minimum value on zero values.
        If objCollOfficer.UkBased Then
            Return CollCostTotal
        End If
        'see what the override is saying 
        If ckIgnoreMinimum.Checked Then         'Ignore it 
            MinCharge = CollCostTotal
        Else
            If MinCharge <= CollCostTotal Then
                MinCharge = CollCostTotal
            End If
        End If

        Return MinCharge

    End Function



    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Me.lvCurrentItem = Nothing
        If Me.ListView1.SelectedItems.Count = 1 Then
            Me.lvCurrentItem = Me.ListView1.SelectedItems.Item(0)
            If Not Me.lvCurrentItem Is Nothing Then
                If Me.lvCurrentItem.Expense.Removed Then
                    Me.cmdRemove.Text = "&Reinstate"
                Else
                    Me.cmdRemove.Text = "&Remove"
                End If
                SetStatusMenus(Me.lvCurrentItem.Expense.PayStatus)

            End If
        End If

        Me.cmdRemove.Enabled = True

    End Sub


    'Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick

    '    If Me.lvCurrentItem Is Nothing Then Exit Sub

    '    Dim frmEdit As New frmExpEntry()
    '    frmEdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
    '    frmEdit.CollOfficer = Me.lvCurrentItem.CollOfficer
    '    frmEdit.PayItem = Me.lvCurrentItem.Expense

    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        Me.lvCurrentItem.Expense = frmEdit.PayItem
    '        Me.lvCurrentItem.Expense.Fields.Update(MedscreenLib.CConnection.DbConnection)
    '        Me.lvCurrentItem.Refresh()
    '        DrawPayItems(MyCollPayInHoursCost)
    '    End If

    'End Sub



    Private Sub DteTimeArrHome_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeArrHome.Leave

    End Sub

    Private Sub DteTimeSetOff_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeSetOff.Leave
        If Me.blnDrawing Then Exit Sub
        If DteTimeArrOnSite.Value < Me.DteTimeSetOff.Value Then Me.DteTimeArrOnSite.Value = Me.DteTimeSetOff.Value
        If DteTimeLeftSite.Value < Me.DteTimeSetOff.Value Then Me.DteTimeLeftSite.Value = Me.DteTimeSetOff.Value
        If DteTimeArrHome.Value < Me.DteTimeSetOff.Value Then Me.DteTimeArrHome.Value = Me.DteTimeSetOff.Value

        UpdateTimes()
        DrawPayItems(MyCollPayInHoursCost)
        'CalcPay()
    End Sub

    Private Sub UpdateTimes(Optional ByVal force As Boolean = False)
        If Me.CollPay.PayMethod <> "H" Then Exit Sub
        If Me.blnDrawing AndAlso Not force Then Exit Sub
        'see if these times overlap another 
        Dim oColl As New Collection()
        oColl.Add(Me.objCollOfficer.OfficerID)
        oColl.Add(Me.DteTimeSetOff.Value.ToString("dd-MMM-yyyy HH:mm:ss"))
        oColl.Add(Me.DteTimeArrHome.Value.ToString("dd-MMM-yyyy HH:mm:ss"))
        oColl.Add(Me.CollPay.CollPayID)
        Dim strReturn As String = CConnection.PackageStringList("Lib_collpay.TimeOverlap", oColl)
        'Check for an overlap
        Me.lblLeaveHome.Text = ""
        Me.lblArriveHome.Text = ""

        If strReturn.Trim.Length > 0 Then
            Dim strErrors As String() = strReturn.Split(New Char() {","})
            Dim StrError As String = CConnection.PackageStringList("Lib_collpay.CollpayDescription", strErrors.GetValue(0))
            'StrError += " at " & strErrors.GetValue(1)
            If CStr(strErrors.GetValue(1)).ToUpper = "END" Then
                Me.ErrorProvider1.SetError(Me.DteTimeArrHome, StrError)
                Me.lblArriveHome.Text = StrError
            Else
                Me.ErrorProvider1.SetError(Me.DteTimeSetOff, StrError)
                Me.lblLeaveHome.Text = StrError

            End If
        End If
        Dim outHours As TimeSpan = Medscreen.HoursOutOfHours(Me.DteTimeSetOff.Value, Me.DteTimeArrHome.Value)
        Dim inHoursMins As Double = CDbl((HoursMinutesWorked.TotalMinutes - outHours.TotalMinutes))
        Me.txtTotalInHours.Text = CDbl(inHoursMins / 60).ToString("00.00")
        Me.txtOutHours.Text = CDbl(outHours.TotalMinutes / 60).ToString("00.00")
        Me.txtHoursWorked.Text = CDbl(HoursMinutesWorked.TotalMinutes / 60).ToString("0.00")
        Me.txtOutHours.Refresh()
        Me.txtHoursWorked.Refresh()
        Me.txtTotalInHours.Refresh()

        Dim price1 As Double = 0
        Dim price2 As Double = 0
        Me.CollPay.PayItems.Clear()
        Me.CollPay.PayItems.load(Me.CollPay.CollPayID)

        Dim objEntry As CollPayItem
        For Each objEntry In Me.CollPay.PayItems
            If Mid(objEntry.ExpenseTypeID, 1, 8) = "COLLCOST" Then
                If objEntry.ExpenseTypeID = "COLLCOST" Or objEntry.ExpenseTypeID = "COLLCOSTCI" Then
                    price1 = Me.objCollOfficer.PayForItem(objEntry.ItemCode, objEntry.PayMonth)
                Else
                    price2 = Me.objCollOfficer.PayForItem(objEntry.ItemCode, objEntry.PayMonth)
                End If
            End If
        Next

        Dim oCmd As New OleDb.OleDbCommand()
        Dim intReturn As Integer
        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "LIB_COLLPAY.SetTimes"
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.Parameters.Add(CConnection.StringParameter("CollpayID", Me.CollPay.CollPayID, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("time1", inHoursMins.ToString, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("time2", outHours.TotalMinutes.ToString, 10))

            If CConnection.ConnOpen Then            'Attempt to open reader
                intReturn = oCmd.ExecuteNonQuery()
            End If
            Me.CollPay.PayItems.Clear()
            Me.CollPay.PayItems.load(Me.CollPay.CollPayID)
            Dim strTos As String = CConnection.PackageStringList("Lib_collpay.TimeOnSite", Me.CollPay.CollPayID)
            Me.lblTimeOnSite.Text = "Time On Site : " & strTos
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollPayEntry-UpdateTimes-1797")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try


    End Sub



    Private Sub DteTimeArrOnSite_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeArrOnSite.Leave
        If Me.blnDrawing Then Exit Sub
        If DteTimeLeftSite.Value < Me.DteTimeArrOnSite.Value Then Me.DteTimeLeftSite.Value = Me.DteTimeArrOnSite.Value
        If DteTimeArrHome.Value < Me.DteTimeArrOnSite.Value Then Me.DteTimeArrHome.Value = Me.DteTimeArrOnSite.Value
        Me.txtHoursWorked.Text = CDbl(HoursMinutesWorked.TotalMinutes / 60).ToString("00.00")
        DrawPayItems(MyCollPayInHoursCost)
        'CalcPay()
    End Sub

    Private Sub DteTimeLeftSite_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeLeftSite.Leave
        If Me.blnDrawing Then Exit Sub
        If DteTimeArrHome.Value < Me.DteTimeLeftSite.Value Then Me.DteTimeArrHome.Value = Me.DteTimeLeftSite.Value
        Me.txtHoursWorked.Text = CDbl(HoursMinutesWorked.TotalMinutes / 60).ToString("00.00")
        DrawPayItems(MyCollPayInHoursCost)
        'CalcPay()
    End Sub


    ''' <summary>
    ''' Created by tonyw on ANDREW at 26/02/2007 15:01:40
    '''     Calculate actual pay
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>tonyw</Author><date> 26/02/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub CalcPay()


        If Me.MyCollPayEntry Is Nothing Then Exit Sub
        If Me.MyCollPayInHoursCost Is Nothing Then Exit Sub

        DrawPayItems(Me.MyCollPayInHoursCost)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get pay details
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	16/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub GetForm()
        If MyCollPayEntry Is Nothing Then Exit Sub
        With MyCollPayEntry
            .TimeLeftHome = Me.DteTimeSetOff.Value
            .TimeArrOnSite = Me.DteTimeArrOnSite.Value
            .TimeLeftSite = Me.DteTimeLeftSite.Value
            .TimeArrHome = Me.DteTimeArrHome.Value
            .Currency = Me.cbCurrency.SelectedValue
            .NoMinimum = Me.ckIgnoreMinimum.Checked
            'If Me.rButSamples.Checked Then
            '    .PayMethod = "S"
            'ElseIf Me.rButHourly.Checked Then
            '    .PayMethod = "H"
            'Else
            '    .PayMethod = "O"
            'End If

            .PaymentMonth = dtPayMonth.Value
            .Comment = Me.txtComment.Text
        End With
    End Sub

    Private Sub SetDateTimeFields()
        Dim outHours As TimeSpan = Medscreen.HoursOutOfHours(Me.DteTimeSetOff.Value, Me.DteTimeArrHome.Value)

        'If Not objCollOfficer.UkBased Then ' exit if Uk c/o 
        '    Exit Sub
        'End If

        If MyCollPayEntry.TimeLeftHome = DateField.ZeroDate Then
            Me.DteTimeSetOff.Value = objCM.CollectionDate
            Me.DteTimeArrOnSite.Value = objCM.CollectionDate
            Me.DteTimeLeftSite.Value = objCM.CollectionDate
            Me.DteTimeArrHome.Value = objCM.CollectionDate
        Else
            Me.DteTimeSetOff.Value = MyCollPayEntry.TimeLeftHome
            Me.DteTimeArrOnSite.Value = MyCollPayEntry.TimeArrOnSite
            Me.DteTimeLeftSite.Value = MyCollPayEntry.TimeLeftSite
            Me.DteTimeArrHome.Value = MyCollPayEntry.TimeArrHome
            Me.txtTotalInHours.Text = CDbl((HoursMinutesWorked.TotalMinutes - outHours.TotalMinutes) / 60).ToString("00.00")
            Me.txtOutHours.Text = CDbl(outHours.TotalMinutes / 60).ToString("00.00")
            Me.txtHoursWorked.Text = MyCollPayEntry.TimeArrHome.Subtract(MyCollPayEntry.TimeLeftHome).ToString
            Me.txtHoursWorked.Refresh()
        End If
        If Me.CollPay.PayMethod = "H" AndAlso Me.CollPay.PayStatus <> "L" Then
            Me.UpdateTimes(True)
            Dim strTW As String = CConnection.PackageStringList("lib_collpay.LongTimeWarning", Me.CollPay.CollPayID)
            Me.lblTimeWarning.Text = strTW
            Application.DoEvents()

        End If
    End Sub

    Private Const cstPayPerSample As String = "S"
    Private Const cstPayPerHour As String = "H"
    Private Const cstPayPerJob As String = "J"

    Private Sub SetCollectorSampleRate()


        If objCollOfficer Is Nothing Then Exit Sub

        'If objCollOfficer.UkBased Then
        If objCollOfficer.PayMethod = cstPayPerSample Then
            MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(SampleRatePayType) ' assign standard UK generic rate.
        ElseIf objCollOfficer.PayMethod = cstPayPerHour Then
            MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_CallOut_In_hours)
        ElseIf objCollOfficer.PayMethod = "C" Then
            MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_Fixed_Price_Charge)
        ElseIf objCollOfficer.PayMethod = "W" Then
            MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_PayWorkingTime)
            MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayTravelTime)
            If MyCollPayOutOfHoursCost Is Nothing Then
                MyCollPayOutOfHoursCost = Me.CollPay.PayItems.OutOfHoursPayItem
            End If
            Me.DTWorkingTime.Value = Today.AddMinutes(MyCollPayInHoursCost.Quantity * 60)
            Me.dtTravelTime.Value = Today.AddMinutes(MyCollPayOutOfHoursCost.Quantity * 60)
        ElseIf objCollOfficer.PayMethod = "O" Then
            MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_PayPerDAY)

        End If
        'Else
        'If ChkCollOffSamplePay(MyCollPayInHoursCost) Then
        '    MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_Per_Sample_Pay)
        'Else
        '    ' Overseas per hourly rates do not currently exist ---- MyCollPayInHoursCost = MyCollPayEntry.PayItems.Item(ExpenseTypeCollections.GCST_Routine_In_)
        'End If
        'End If
        If MyCollPayInHoursCost Is Nothing Then Exit Sub

        MyCollPayInHoursCost.DoUpdate()
    End Sub

    Private Function IsUkCollection()
        Dim boolIsUkColl As Boolean = False
        If objCM.UK Then
            Me.gBoxPayOpts.Show()
            Me.gBoxCollOpts.Show()
            Me.TimeSumBox.Show()
            Me.PaySummBox.Show()
            Me.gBoxTimeEntry.Show()

            Me.grpPerSampleRate.Hide()


            'Me.PaySummBox.Left = 160
            'Me.PaySummBox.Top = 96

            Me.Text = "UK Payment For " & objCollOfficer.Name & " - " & Me.CollPay.PayStatusString
            boolIsUkColl = True
        Else
            Me.gBoxCollOpts.Show()
            Me.gBoxPayOpts.Show()
            Me.gBoxTimeEntry.Show()
            Me.TimeSumBox.Show()
            Me.PaySummBox.Show()
            'Me.grpPerSampleRate.Left = 160
            'Me.grpPerSampleRate.Top = 16

            Me.grpPerSampleRate.Hide()
            Me.Text = "OVERSEAS Payment For " & objCollOfficer.Name & " - " & Me.CollPay.PayStatusString
        End If
        Return boolIsUkColl
    End Function


    Private Sub ChkCancelledSamples()

        If objCM.NoCancelled > 0 Then
            Me.lblCancSamples.ForeColor = System.Drawing.Color.RosyBrown
            Me.lblCancSamples.Text = Me.lblCancSamples.Text & " " & objCM.NoCancelled
        Else
            Me.lblCancSamples.Text = "No Samples Cancelled."
            Me.lblCancSamples.Enabled = False
            Me.lblCancSamples.ForeColor = System.Drawing.Color.Black
        End If

    End Sub

    Private Function ChkCollOffSamplePay(ByVal InHoursCosts) As Boolean
        Dim boolSampRate As Boolean = False


        If InHoursCosts Is Nothing Then Exit Function

        If InHoursCosts.PayType = SampleRatePayType() Then
            boolSampRate = True
            Me.rButSamples.Checked = True
        Else
            Me.rButRoutine.Checked = False
        End If

        Return boolSampRate

    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exit and commit data 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	16/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.CollPay.PayStatus <> CollPayItem.GCST_Pay_Locked Then Me.GetForm()
        If Me.FormType = enFormType.CollRelatedPay Then
            Me.CollPay.PayType = "C"
        ElseIf Me.FormType = enFormType.NonCollRelatedExp Then
            Me.CollPay.PayType = "E"
        ElseIf Me.FormType = enFormType.OCR Then
            Me.CollPay.PayType = "R"
        End If
        If Me.CollPay.PayStatus.Trim.Length = 0 Then Me.CollPay.PayStatus = "E"
        If Not Me.MyCollPayInHoursCost Is Nothing Then Me.MyCollPayInHoursCost.DoUpdate()
        Me.CollPay.DoUpdate()
        Dim oCPItem As CollPayItem
        'Make sure individual items are updated
        For Each oCPItem In Me.CollPay.PayItems
            If oCPItem.PayStatus <> CollPayItem.GCST_Pay_Locked Then oCPItem.DoUpdate()
        Next
    End Sub


    'Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    '    AddExpenseItem()
    '    DrawPayItems(MyCollPayInHoursCost)
    'End Sub


    'Private Sub cmdAddPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddPay.Click
    '    AddPayItem()
    '    DrawPayItems(MyCollPayInHoursCost)
    'End Sub

    'Private Sub AddExpenseItem()
    '    Dim frmEdit As New frmExpEntry()
    '    Dim objCollPayItem As Intranet.collectorpay.CollPayItem = Me.MyCollPayEntry.GetExpenseItem

    '    frmEdit.FormType = frmExpEntry.enumExpFormType.stdExpEntryFormat
    '    If objCollPayItem Is Nothing Then Exit Sub 'Fatal error

    '    objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '    objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '    If Me.FormType = enFormType.CollRelatedPay Then
    '        frmEdit.CollOfficer = Me.objCollOfficer
    '    Else
    '        frmEdit.CollOfficer = Me.cbCollOfficer.SelectedItem
    '    End If
    '    frmEdit.PayItem = objCollPayItem
    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        objCollPayItem = frmEdit.PayItem
    '        objCollPayItem.SetPriceInfo(frmEdit.Quantity.Value, CDbl(frmEdit.txtPrice.Text), frmEdit.Discount)
    '        objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '        objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '        If Me.FormType = enFormType.CollRelatedPay Then
    '            Me.ListView1.Items.Add(New ExpenseDescription(objCollPayItem, Me.objCollOfficer))
    '        Else
    '            Me.ListView1.Items.Add(New ExpenseDescription(objCollPayItem, Me.cbCollOfficer.SelectedItem))
    '        End If
    '    Else
    '        objCollPayItem.CancelItem() ' = True

    '    End If

    '    MyCollPayEntry.SetOfficer(Me.cbCollOfficer.SelectedValue, Me.dtPayMonth.Value, objCollPayItem.PayType)

    'End Sub


    'Private Sub AddPayItem()
    '    Dim frmEdit As New frmExpEntry()
    '    Dim objCollPayItem As Intranet.collectorpay.CollPayItem = Me.MyCollPayEntry.GetPayItem

    '    frmEdit.FormType = frmExpEntry.enumExpFormType.OtherPayItem
    '    If objCollPayItem Is Nothing Then Exit Sub 'Fatal error

    '    objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '    objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '    If Me.FormType = enFormType.CollRelatedPay Then
    '        frmEdit.CollOfficer = Me.objCollOfficer
    '    Else
    '        frmEdit.CollOfficer = Me.cbCollOfficer.SelectedItem
    '    End If
    '    frmEdit.PayItem = objCollPayItem
    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        objCollPayItem = frmEdit.PayItem
    '        objCollPayItem.SetPriceInfo(frmEdit.Quantity.Value, CDbl(frmEdit.txtPrice.Text))
    '        objCollPayItem.PayMonth = Me.dtPayMonth.Value
    '        objCollPayItem.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered

    '        If Me.FormType = enFormType.CollRelatedPay Then
    '            Me.ListView1.Items.Add(New ExpenseDescription(objCollPayItem, Me.objCollOfficer))
    '        Else
    '            Me.ListView1.Items.Add(New ExpenseDescription(objCollPayItem, Me.cbCollOfficer.SelectedItem))
    '        End If
    '    Else
    '        objCollPayItem.CancelItem() ' = True

    '    End If

    '    MyCollPayEntry.SetOfficer(Me.cbCollOfficer.SelectedValue, Me.dtPayMonth.Value, objCollPayItem.PayType)

    'End Sub


    Private Sub txtTotalInHours_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalInHours.Leave


    End Sub



    Private Sub DteTimeArrHome_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeArrHome.ValueChanged
        If Me.blnDrawing Then Exit Sub
        If Not Me.Visible = True Then Exit Sub

        If DteTimeArrHome.Value < DteTimeLeftSite.Value Then
            MsgBox("You Have Tried To Enter A Date Which Is Earlier Than The LEFT SITE Time/Date")
            DteTimeArrHome.Value = DteTimeLeftSite.Value
            DteTimeArrHome.Refresh()
        End If
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
        End If


    End Sub



    Private Sub DteTimeLeftSite_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeLeftSite.ValueChanged
        If Me.blnDrawing Then Exit Sub
        If Not Me.Visible Then Exit Sub

        If DteTimeLeftSite.Value < DteTimeArrOnSite.Value Then
            MsgBox("You Have Tried To Enter A Date Which Is Earlier Than The ARRIVE ON SITE Time/Date")
            DteTimeLeftSite.Value = DteTimeArrOnSite.Value
            DteTimeLeftSite.Refresh()
        End If
        'UpdateTimes()
        'DrawPayItems(MyCollPayInHoursCost)
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
        End If

    End Sub

    Private Sub DteTimeArrOnSite_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeArrOnSite.ValueChanged

        If Not Me.Visible Then Exit Sub

        If DteTimeArrOnSite.Value < DteTimeSetOff.Value Then
            MsgBox("You Have Tried To Enter A Date Which Is Earlier Than The LEFT HOME Time/Date")
            DteTimeArrOnSite.Value = DteTimeSetOff.Value
            DteTimeArrOnSite.Refresh()
        End If

        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
        End If

        'UpdateTimes()
        'DrawPayItems(MyCollPayInHoursCost)

    End Sub

    Private Sub DteTimeLeftSite_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeLeftSite.TextChanged

    End Sub

    Private Sub DteTimeLeftSite_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DteTimeLeftSite.KeyPress

    End Sub

    Private Sub DteTimeArrHome_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DteTimeArrHome.KeyPress


    End Sub

    Private Sub DteTimeSetOff_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DteTimeSetOff.ValueChanged
        If Me.blnDrawing Then Exit Sub
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
        End If

    End Sub

    'Private Sub Donors_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Donors.Leave
    '    If Me.rButSamples.Checked Then
    '        Me.FillForm()
    '    End If
    'End Sub

    '''  -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculates mileage expenses 
    ''' </summary>
    ''' <param name="blnDrawing"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	19/06/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateMileage(Optional ByVal blnDrawing As Boolean = False)
        If Me.blnDrawing Then Exit Sub
        If Me.CollPay Is Nothing Then Exit Sub
        If Me.CollPay.PayStatus <> ExpenseTypeCollections.GCST_Pay_Status_Entered Then Exit Sub

        If Me.MyMileageCost Is Nothing Then
            'See if it is present but not yet loaded
            If Me.CollPay.PayItems.IsMileagePresent Then
                MyMileageCost = CollPay.PayItems.MileagePay
            Else
                'Create an entry 
                MyMileageCost = Me.CollPay.PayItems.Create()
                If Me.udKMMiles.SelectedIndex = 0 Then
                    MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerMile
                Else
                    MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerKM
                End If
                MyMileageCost.Currency = CollPay.Currency
                MyMileageCost.Quantity = 0
                MyMileageCost.PayMonth = CollPay.PaymentMonth
            End If
        End If

        MyMileageCost.Currency = Me.cbCurrency.SelectedValue

        If blnDrawing Then Me.udMileage.Value = MyMileageCost.Quantity
        If Not MyMileageCost.PayStatus = "L" Then
            MyMileageCost.Quantity = Me.udMileage.Value
            MyMileageCost.Currency = objCollOfficer.PayCurrency
            MyMileageCost.PayMonth = Me.dtPayMonth.Value
            MyMileageCost.ItemPrice = objCollOfficer.PriceForItem(MyMileageCost.ItemCode, Me.dtPayMonth.Value)
            MyMileageCost.ItemDiscount = objCollOfficer.DiscountONItem(MyMileageCost.ItemCode, Me.dtPayMonth.Value)
            MyMileageCost.ItemValue = MyMileageCost.ItemPrice * (MyMileageCost.ItemDiscount) * MyMileageCost.Quantity
            Me.txtMileage.Text = MyMileageCost.ItemValue.ToString("c", Me.objCollOfficer.Culture)
            MyMileageCost.DoUpdate()
        End If
        RefreshCollCostListView()
        CalcPay()

    End Sub

    Private Sub UpDateCollPayItem()
        Me.DTWorkingTime.Hide()
        Me.dtTravelTime.Hide()
        Me.lblTravelTime.Hide()
        Me.lblCancSamples.Show()
        If Me.CollPay Is Nothing Then Exit Sub
        Me.Donors.Show()
        'set coll pay pay method 
        Me.lblHalfDays.Hide()
        Me.udHalfdays.Hide()
        Me.PanelSamplePlus.Hide()
        If Me.rbOther.Checked Then
            Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODDay
            Me.lblDonors.Text = "Days:"
            MyCollPayInHoursCost.Quantity = Donors.Value
            Me.lblHalfDays.Show()
            Me.udHalfdays.Show()


        ElseIf Me.rButSamples.Checked Then
            Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODSample
            Me.lblDonors.Text = "Samples:"
            MyCollPayInHoursCost.Quantity = Donors.Value
            PanelSamplePlus.Show()

        ElseIf Me.rButHourly.Checked Then
            Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODHourly
            Me.lblDonors.Text = "Samples:"
            UpdateTimes()
            DrawPayItems(MyCollPayInHoursCost)

            RefreshCollCostListView()
            Exit Sub
        ElseIf Me.rbCollection.Checked Then
            Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODCollection
            Me.lblDonors.Text = "Samples:"
            MyCollPayInHoursCost.Quantity = 1

        ElseIf Me.rbWorkingTime.Checked Then
            Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODWorkingTime
            Me.lblDonors.Text = "Working Time:"
            Me.Donors.Hide()
            Me.DTWorkingTime.Show()
            Me.lblTravelTime.Show()
            Me.dtTravelTime.Show()
            Me.lblCancSamples.Hide()
        End If

        MyCollPayInHoursCost.Currency = Me.cbCurrency.SelectedValue
        'MyCollPayInHoursCost.ItemPrice = Me.objCollOfficer.PriceForItem(MyCollPayInHoursCost.ItemCode, MyCollPayInHoursCost.PayMonth)
        'MyCollPayInHoursCost.ItemValue = Me.objCollOfficer.PayForItem(MyCollPayInHoursCost.ItemCode, MyCollPayInHoursCost.PayMonth) * MyCollPayInHoursCost.Quantity
        'MyCollPayInHoursCost.DoUpdate()

        'UpdateCollPayInHoursCosts(MyCollPayInHoursCost)


        'If MyCollPayOutOfHoursCost Is Nothing Then MyCollPayOutOfHoursCost = Me.CollPay.PayItems.OutOfHoursPayItem

        'UpdateCollPayOutOfHoursCost(MyCollPayOutOfHoursCost)
        Me.CheckMinumumValue()


        DrawPayItems(MyCollPayInHoursCost)

        RefreshCollCostListView()
        'UpdateTimes()
    End Sub
    ''' <summary>
    ''' Created by tonyw on ANDREW at 27/02/2007 13:35:32
    '''     
    ''' </summary>
    ''' <param name="InHoursCosts" type="Intranet.collectorpay.CollPayItem">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>tonyw</Author><date> 27/02/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub UpdateCollPayInHoursCostsx(ByVal InHoursCosts As Intranet.collectorpay.CollPayItem)
        If objCM Is Nothing Then Exit Sub 'All the rest is collection related so no CM no need to do anything 

        If Me.FormType = enFormType.NonCollRelatedExp Then Exit Sub
        'Check to see that it hasn't moved on 
        If Me.CollPay.PayStatus <> ExpenseTypeCollections.GCST_Pay_Status_Entered Then Exit Sub

        Dim outHours As TimeSpan = Medscreen.HoursOutOfHours(Me.DteTimeSetOff.Value, Me.DteTimeArrHome.Value)

        If InHoursCosts Is Nothing Then
            If Me.CollPay.PayItems.IsStandardPayItemPresent Then
                InHoursCosts = Me.CollPay.PayItems.StandardPayItem
            Else
                InHoursCosts = Me.CollPay.PayItems.Create()
                InHoursCosts.Currency = Me.CollPay.Currency     'Set the currency of the item

            End If
        End If
        InHoursCosts.Currency = Me.cbCurrency.SelectedValue

        If Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODSample Then
        ElseIf Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODHourly Then 'Dealing with hourly costs 

            ' must be hourly inhours pay, but could be call out 
            If objCM.CollectionType.Trim.Length = 0 Then objCM.CollectionType = "R"
            If Asc(objCM.CollectionType) = objCM.CollectionTypes.GCST_JobTypeCallout Or Asc(objCM.CollectionType) = objCM.CollectionTypes.GCST_JobTypeCalloutOverseas Then
                InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_CallOut_In_hours
            Else
                InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_Routine_In_hours
            End If

            If (CDbl((HoursMinutesWorked.TotalMinutes >= outHours.TotalMinutes))) Then
                InHoursCosts.Quantity = CDbl((HoursMinutesWorked.TotalMinutes - outHours.TotalMinutes) / 60).ToString("00.00")
            Else
                InHoursCosts.Quantity = CDbl((outHours.TotalMinutes - HoursMinutesWorked.TotalMinutes) / 60).ToString("00.00")
            End If
            InHoursCosts.ItemPrice = objCollOfficer.PriceForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)

            InHoursCosts.ItemValue = ChkMinCollCarge(InHoursCosts.ItemValue)

            InHoursCosts.ItemDiscount = objCollOfficer.DiscountONItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemValue = InHoursCosts.Quantity * InHoursCosts.ItemPrice * InHoursCosts.ItemDiscount

            InHoursCosts.CollPayID = Me.MyCollPayEntry.CollPayID
        ElseIf Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODDay Then
            InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerDAY
            InHoursCosts.Quantity = Me.Donors.Value
            InHoursCosts.ItemPrice = objCollOfficer.PriceForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemDiscount = objCollOfficer.DiscountONItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * InHoursCosts.Quantity
            'See if less than minimum charge
            Me.CheckMinumumValue()
            'Fixed Rate 
        ElseIf Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODCollection Then 'Pay per collection
            InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_Fixed_Price_Charge
            InHoursCosts.Quantity = 1
            InHoursCosts.ItemPrice = objCollOfficer.PriceForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemDiscount = objCollOfficer.DiscountONItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * InHoursCosts.Quantity
            'Working time 
        ElseIf Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODWorkingTime Then
            InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_PayWorkingTime
            Dim DummyTime As Date = Today
            MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayTravelTime)
            If MyCollPayOutOfHoursCost Is Nothing Then
                MyCollPayOutOfHoursCost = Me.CollPay.PayItems.OutOfHoursPayItem
            End If
            InHoursCosts.Quantity = Me.DTWorkingTime.Value.Subtract(DummyTime).TotalMinutes / 60
            InHoursCosts.ItemPrice = objCollOfficer.PriceForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemDiscount = objCollOfficer.DiscountONItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
            InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * InHoursCosts.Quantity
            MyCollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayTravelTime
            MyCollPayOutOfHoursCost.Quantity = Me.dtTravelTime.Value.Subtract(DummyTime).TotalMinutes / 60
            MyCollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
            MyCollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
            MyCollPayOutOfHoursCost.ItemValue = objCollOfficer.PayForItem(MyCollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value) * MyCollPayOutOfHoursCost.Quantity
            MyCollPayOutOfHoursCost.DoUpdate()
            Me.CheckMinumumValue()
        End If

        If InHoursCosts.PayStatus = "" Then InHoursCosts.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
        InHoursCosts.PayMonth = Me.dtPayMonth.Value
        InHoursCosts.DoUpdate()

    End Sub

    Private Sub SamplePay(ByVal InHoursCosts As Intranet.collectorpay.CollPayItem)
        ' Need to ensure that we have no hourly entry
        'If objCM.CollectionPort.IsUK Then
        '    InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_Std_UK_Sample_Pay
        'Else
        'InHoursCosts.ExpenseTypeID = ExpenseTypeCollections.GCST_Std_Sample_Pay
        Dim intIndex As Integer = Me.updLineitem.SelectedIndex + 1
        If intIndex < 1 Then Exit Sub
        Dim strLineItem = Me.cstPaySampleStub & CStr(intIndex).Trim
        InHoursCosts.ExpenseTypeID = strLineItem

        'End If
        InHoursCosts.Quantity = Me.Donors.Value
        InHoursCosts.ItemPrice = objCollOfficer.PriceForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
        InHoursCosts.ItemDiscount = objCollOfficer.DiscountONItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
        InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * InHoursCosts.Quantity
        'see if we have a volume discount
        Dim objVD As Object = CConnection.PackageStringList("lib_collpay.HasVolumeDiscount", Me.objCollOfficer.OfficerID)
        If Not objVD Is Nothing AndAlso objVD = "T" Then
            'Need to find the discount and number of samples at each stage
            Dim intsamples As Integer = 0
            Dim CurrentDiscount As Double = 0
            Dim i As Integer
            Dim blnPassOne As Boolean = True
            For i = 1 To Me.Donors.Value
                Dim objcoll As New Collection()
                objcoll.Add(Me.objCollOfficer.OfficerID)
                objcoll.Add(Me.objCM.ID)
                objcoll.Add("")
                objcoll.Add(InHoursCosts.ItemCode)
                objcoll.Add(i)
                objcoll.Add("")
                objcoll.Add(Me.dtPayMonth.Value.ToString("dd-MMM-yy"))
                objcoll.Add("Coll_Officer.Identity")
                Dim objDisc As Object = CConnection.PackageStringList("PckInvoicing.GetVolumeDiscount", objcoll)
                If CurrentDiscount = 0 Then CurrentDiscount = objDisc
                If CurrentDiscount <> objDisc Or i = Me.Donors.Value Then 'Completed first part 
                    InHoursCosts.ItemDiscount = CurrentDiscount
                    InHoursCosts.Quantity = intsamples + 1
                    If i = Me.Donors.Value Then intsamples += 1
                    InHoursCosts.ItemPrice = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
                    InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * CurrentDiscount * InHoursCosts.Quantity
                    InHoursCosts.Comments = "Pay for samples " & CStr(i - intsamples) & " to " & i
                    InHoursCosts.DoUpdate()
                    intsamples = 0
                    If blnPassOne Then
                        If CurrentDiscount <> objDisc Then 'is really banded 
                            If MyCollPayOutOfHoursCost Is Nothing Then
                                MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item("BANDED")
                                If MyCollPayOutOfHoursCost Is Nothing Then MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Create
                            End If
                            If Not MyCollPayOutOfHoursCost Is Nothing Then
                                MyCollPayOutOfHoursCost.ExpenseTypeID = "COLLCOSTO"
                                MyCollPayOutOfHoursCost.ItemDiscount = InHoursCosts.ItemDiscount
                                MyCollPayOutOfHoursCost.Quantity = InHoursCosts.Quantity - 1
                                MyCollPayOutOfHoursCost.Comments = InHoursCosts.Comments
                                MyCollPayOutOfHoursCost.ItemPrice = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value)
                                MyCollPayOutOfHoursCost.ItemValue = InHoursCosts.ItemPrice * MyCollPayOutOfHoursCost.Quantity * CurrentDiscount
                                If MyCollPayOutOfHoursCost.CollPayID.Trim.Length = 0 Then
                                    MyCollPayOutOfHoursCost.CollPayID = InHoursCosts.CollPayID
                                End If
                                If MyCollPayOutOfHoursCost.PayID.Trim.Length = 0 Then
                                    MyCollPayOutOfHoursCost.PayID = CConnection.NextSequence("SEQ_COLLECTORPAY_ENTRY")
                                End If
                                MyCollPayOutOfHoursCost.DoUpdate()
                                InHoursCosts.Quantity = 1
                                InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * objDisc * InHoursCosts.Quantity
                            End If
                        Else
                            If MyCollPayOutOfHoursCost Is Nothing Then
                                MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item("COLLCOSTO")
                            End If
                            If Not MyCollPayOutOfHoursCost Is Nothing Then
                                MyCollPayOutOfHoursCost.Quantity = 0
                                MyCollPayOutOfHoursCost.ItemValue = 0
                                MyCollPayOutOfHoursCost.DoUpdate()
                            End If
                            InHoursCosts.Quantity = Me.Donors.Value
                            InHoursCosts.ItemValue = objCollOfficer.PayForItem(InHoursCosts.ItemCode, Me.dtPayMonth.Value) * objDisc * InHoursCosts.Quantity

                        End If
                        blnPassOne = False
                    End If
                End If
                CurrentDiscount = objDisc
                intsamples += 1
            Next
        Else
            InHoursCosts.DoUpdate()
        End If
        'See if less than minimum charge
        Me.CheckMinumumValue()

    End Sub

    Private Sub UpdateCollPayOutOfHoursCostx(ByVal CollPayOutOfHoursCost As Intranet.collectorpay.CollPayItem)
        If objCM Is Nothing Then Exit Sub 'All the rest is collection related so no CM no need to do anything 
        If Me.CollPay.PayMethod = CollPay.GCST_PAYMETHODSample Then Exit Sub 'May be using this item for banded payments

        Dim OutHours As TimeSpan = Medscreen.HoursOutOfHours(Me.DteTimeSetOff.Value, Me.DteTimeArrHome.Value)
        'Check to see that it hasn't moved on 
        If Me.CollPay.PayStatus <> ExpenseTypeCollections.GCST_Pay_Status_Entered Then Exit Sub

        'If Not objCollOfficer.UkBased Then Exit Sub ' BAIL OUT IF OVERSEAS C/O.

        If CollPayOutOfHoursCost Is Nothing Then
            If Me.CollPay.PayItems.IsOutOfHoursItemPresent Then
                CollPayOutOfHoursCost = Me.CollPay.PayItems.OutOfHoursPayItem
            Else
                CollPayOutOfHoursCost = Me.MyCollPayEntry.PayItems.Create()
                CollPayOutOfHoursCost.Currency = Me.CollPay.Currency     'Set the currency of the item

            End If
            '
        End If
        CollPayOutOfHoursCost.Currency = Me.cbCurrency.SelectedValue

        'If OutHours.TotalMinutes = 0 Then Exit Sub
        If objCollOfficer Is Nothing Then Exit Sub
        'If objCollOfficer.UkBased Then
        If CollPayOutOfHoursCost.IsSampleTypeExpenseId(CollPayOutOfHoursCost.ExpenseTypeID) Then Exit Sub

        If Me.rButSamples.Checked Then ' For sample payments, remove the secondary out-of-hours payment
            CollPayOutOfHoursCost.Removed = True
            CollPayOutOfHoursCost.ItemPrice = 0
            CollPayOutOfHoursCost.ItemValue = 0
            CollPayOutOfHoursCost.ItemDiscount = 0
            CollPayOutOfHoursCost.Quantity = 0
            CollPayOutOfHoursCost.DoUpdate()
            Exit Sub
        Else
            CollPayOutOfHoursCost.Removed = False
        End If
        'End If

        If Me.CollPay.PayMethod = "H" Then
            If (objCM.CollectionType = Constants.GCST_JobTypeCallout) Or objCM.CollectionType = Constants.GCST_JobTypeCalloutOverseas Then
                CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_CallOut_Out_hours
            Else
                CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_Routine_Out_hours
            End If
        ElseIf Me.CollPay.PayMethod = "W" Then
            CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayTravelTime

        End If
        If objCM.CollectionType = ExpenseTypeCollections.GCST_UK_CallOut_Type_Coll Then 'If a UK callout collection type....
            If Me.rButHourly.Checked Then ' process hourly payment.
                CollPayOutOfHoursCost.Quantity = CDbl((OutHours.TotalMinutes) / 60).ToString("00.00")
                CollPayOutOfHoursCost.ItemValue = CDbl((OutHours.TotalMinutes) / 60).ToString * objCollOfficer.PayForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                CollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                CollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_CallOut_Out_hours
            Else ' Sample payment seleted.
                CollPayOutOfHoursCost.Removed = True
                'CollPayOutOfHoursCost.ItemValue = Me.Donors.Value * objCollOfficer.PayForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                'CollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                'CollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(CollPayOutOfHoursCost.ItemCode)
                'CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_Std_UK_Sample_Pay
            End If
        ElseIf objCM.CollectionType = ExpenseTypeCollections.GCST_Std_Routine_Type_Coll Then ' Routine - both UK and Overseas.

            'If objCollOfficer.UkBased Then
            If Me.CollPay.PayMethod = "H" Then ' process hourly payment.
                CollPayOutOfHoursCost.Quantity = CDbl((OutHours.TotalMinutes) / 60).ToString("00.00")
                CollPayOutOfHoursCost.ItemValue = CollPayOutOfHoursCost.Quantity * objCollOfficer.PayForItem(CollPayOutOfHoursCost.ItemCode)
                CollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                CollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
            ElseIf Me.CollPay.PayMethod = Intranet.collectorpay.CollectorPay.GCST_PAYMETHODWorkingTime Then
                CollPayOutOfHoursCost.Quantity = Me.dtTravelTime.Value.Subtract(Today).TotalMinutes / 60
                CollPayOutOfHoursCost.ItemValue = CollPayOutOfHoursCost.Quantity * objCollOfficer.PayForItem(CollPayOutOfHoursCost.ItemCode)
                CollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                CollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)

            Else ' Sample payment seleted.
                CollPayOutOfHoursCost.Removed = True
                'CollPayOutOfHoursCost.ItemValue = Me.Donors.Value * objCollOfficer.PayForItem(CollPayOutOfHoursCost.ItemCode)
                'CollPayOutOfHoursCost.ItemPrice = objCollOfficer.PriceForItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                'CollPayOutOfHoursCost.ItemDiscount = objCollOfficer.DiscountONItem(CollPayOutOfHoursCost.ItemCode, Me.dtPayMonth.Value)
                'CollPayOutOfHoursCost.ExpenseTypeID = ExpenseTypeCollections.GCST_Std_UK_Sample_Pay
                'End If
            End If
        End If

        CollPayOutOfHoursCost.PayStatus = ExpenseTypeCollections.GCST_Pay_Status_Entered
        CollPayOutOfHoursCost.PayMonth = Me.dtPayMonth.Value
        CollPayOutOfHoursCost.CollPayID = Me.MyCollPayEntry.CollPayID

        CollPayOutOfHoursCost.DoUpdate()

    End Sub

    Private Sub CheckMinumumValue()
        If Me.ckIgnoreMinimum.Checked Then Exit Sub
        Me.CollPay.MinimumCharge()
    End Sub

    Private Sub BuildItems()
        If objCM Is Nothing Then Exit Sub
        If CollPay Is Nothing Then Exit Sub
        Dim oCmd As New OleDb.OleDbCommand()
        Dim intReturn As Integer
        Try ' Protecting 
            'oCmd.CommandText = "LIB_COLLPAY.CreatePayEntries"
            'oCmd.CommandType = CommandType.StoredProcedure
            'oCmd.Parameters.Add(CConnection.StringParameter("CMJobNumber", Me.objCM.ID, 10))
            'oCmd.Parameters.Add(CConnection.StringParameter("Officer", Me.objCollOfficer.OfficerID, 10))
            'oCmd.Parameters.Add(CConnection.StringParameter("payMethod", Me.CollPay.PayMethod, 10))
            'oCmd.Connection = CConnection.DbConnection
            'If CConnection.ConnOpen Then            'Attempt to open reader
            '    intReturn = oCmd.ExecuteNonQuery()
            'End If
            Me.CollPay.PayItems.Clear()
            Me.CollPay.PayItems.load(Me.CollPay.CollPayID)
            Select Case Me.CollPay.PayMethod
                Case "S"
                    Dim strExpenseItem As String = CConnection.PackageStringList("Lib_Collpay.SampleExpenseType", Me.CollPay.CollPayID)
                    Dim intPos As Integer = CInt(Mid(strExpenseItem, strExpenseItem.Length - 2))
                    Me.updLineitem.SelectedIndex = intPos - 1
                    MyCollPayInHoursCost = Me.CollPay.PayItems.Item(strExpenseItem)
                    'MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Std_UK_Sample_Pay)
                    MyCollPayInHoursCost.ExpenseTypeID = strExpenseItem
                    Me.MyCollPayOutOfHoursCost = Nothing
                    If Not MyCollPayInHoursCost Is Nothing Then Me.Donors.Value = MyCollPayInHoursCost.Quantity
                Case "H"
                    If Not objCM Is Nothing Then
                        If objCM.CollectionType = "C" Or objCM.CollectionType = "O" Then
                            MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_CallOut_In_hours)
                            Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_CallOut_Out_hours)
                        Else
                            MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Routine_In_hours)
                            Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Routine_Out_hours)
                        End If
                    Else
                        MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Routine_In_hours)
                        Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Routine_Out_hours)
                    End If
                Case "O"
                    Me.TimeSumBox.Enabled = False
                    Me.PaySummBox.Enabled = True
                    MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayPerDAY)
                    Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayPerHalfDAY)
                    If Not MyCollPayInHoursCost Is Nothing Then Me.Donors.Value = MyCollPayInHoursCost.Quantity
                    If Not MyCollPayOutOfHoursCost Is Nothing Then Me.udHalfdays.Value = MyCollPayOutOfHoursCost.Quantity
                Case "C"
                    MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_Fixed_Price_Charge)
                    Me.MyCollPayOutOfHoursCost = Nothing
                Case "W"
                    MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayWorkingTime)
                    Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayTravelTime)
            End Select


        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollPayEntry-BuildItems-2631")
        Finally

            Dim strinfo As String = CConnection.PackageStringList("Lib_collpay.getdebuginfo", "")
            Debug.WriteLine(strinfo)
            CConnection.SetConnClosed()             'Close connection
        End Try
    End Sub

    Private Sub rButSamples_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rButSamples.CheckedChanged, rbOther.CheckedChanged

        If NewForm Then Exit Sub
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
            Me.BuildItems()
        End If


    End Sub

    Private Sub rButHourly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rButHourly.CheckedChanged, rbCollection.CheckedChanged, rbWorkingTime.CheckedChanged
        If Me.blnDrawing Then Exit Sub
        If NewForm Then Exit Sub
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            UpDateCollPayItem()
            BuildItems()
        End If


    End Sub

    Private Sub cmdRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Me.cmdRemove.Enabled = True

        Me.lvCurrentItem.Expense.Removed = Not Me.lvCurrentItem.Expense.Removed
        Me.lvCurrentItem.Expense.DoUpdate()

        Me.lvCurrentItem.Refresh()
        Me.MyCollPayEntry.DoUpdate()
        DrawPayItems(MyCollPayInHoursCost)
        RefreshCollCostListView()

    End Sub




    Private Sub Donors_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Donors.ValueChanged, Donors.Leave
        If Me.blnDrawing Then Exit Sub
        Me.lblTimeWarning.Text = ""
        If Me.CollPay.PayMethod = "S" Then
            Me.CollPay.NoSamples = Donors.Value
            Dim strExpenseItem As String = ExpenseTypeCollections.GCST_Per_Sample_PayStub & CStr(Me.updLineitem.SelectedIndex + 1).Trim
            MyCollPayInHoursCost = Me.CollPay.PayItems.Item(strExpenseItem)
            Me.SamplePay(MyCollPayInHoursCost)
        ElseIf Me.CollPay.PayMethod = "O" Then
            MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayPerDAY)
        ElseIf Me.CollPay.PayMethod = "H" Then
            Me.CollPay.NoSamples = Donors.Value
            Dim strTW As String = CConnection.PackageStringList("lib_collpay.LongTimeWarning", Me.CollPay.CollPayID)
            Me.lblTimeWarning.Text = strTW
            Application.DoEvents()
        Else
            Exit Sub
        End If
        If Me.MyCollPayEntry Is Nothing Then Exit Sub 'We must have a coll pay item 
        If Not Me.MyCollPayInHoursCost Is Nothing Then
            Me.MyCollPayInHoursCost.Quantity = Donors.Value
            MyCollPayInHoursCost.ItemValue = MyCollPayInHoursCost.ItemPrice * MyCollPayInHoursCost.ItemDiscount * MyCollPayInHoursCost.Quantity
            MyCollPayInHoursCost.DoUpdate()
        End If
        DrawPayItems(MyCollPayInHoursCost)
        RefreshCollCostListView()
        If Me.CollPay.PayMethod = "S" Or Me.CollPay.PayMethod = "H" Then

        End If
    End Sub

    Private Sub UDHalfDay_changed(ByVal sender As Object, ByVal e As System.EventArgs) Handles udHalfdays.Leave, udHalfdays.ValueChanged
        If Me.blnDrawing Then Exit Sub
        If Me.CollPay.PayMethod = "O" Then
            MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayPerHalfDAY)
        Else
            Exit Sub
        End If
        If Me.MyCollPayEntry Is Nothing Then Exit Sub 'We must have a coll pay item 
        If Not Me.MyCollPayOutOfHoursCost Is Nothing Then
            Me.MyCollPayOutOfHoursCost.Quantity = udHalfdays.Value
            MyCollPayOutOfHoursCost.ItemValue = MyCollPayOutOfHoursCost.ItemPrice * MyCollPayOutOfHoursCost.ItemDiscount * MyCollPayOutOfHoursCost.Quantity
            MyCollPayOutOfHoursCost.DoUpdate()
        End If
        DrawPayItems(MyCollPayInHoursCost)
        RefreshCollCostListView()

    End Sub




    Private Sub cbCollOfficer_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCollOfficer.SelectedValueChanged
        If Me.blnDrawing Then Exit Sub
        If Not Me.MyCollPayEntry Is Nothing Then
            MyCollPayEntry.CollOfficer = Me.cbCollOfficer.SelectedValue
            If Not Me.cbCollOfficer.SelectedItem Is Nothing Then
                Dim strXML As String = CType(Me.cbCollOfficer.SelectedItem, Intranet.intranet.jobs.CollectingOfficer).ToHtml
                Me.HtmlEditor1.DocumentText = (strXML)
            End If
        End If

    End Sub



    Private Sub udMileage_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udMileage.ValueChanged, udMileage.Leave
        UpdateMileage()
    End Sub

    Private Sub mnuLockPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLockPay.Click
        If Me.MyCollPayEntry Is Nothing Then Exit Sub
        'Check to see if we can lock pay 
        If Not Me.blnUserCanLockPay Then
            MsgBox("You are not in the Role which allows you to lock pay", MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim objCollPayItem As CollPayItem
        'Set header status
        Me.MyCollPayEntry.LockPay()
        SetStatus()
    End Sub

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
        Me.cmdRefresh.Enabled = True

        If Status = "L" Then
            Me.mnuLockPay.Enabled = False
            Me.mnuApproveItem.Enabled = False
            Me.mnuCancelItem.Enabled = False
            Me.MnuReferItem.Enabled = False
            Me.mnuDeclineItem.Enabled = False
            Me.mnuEntered.Enabled = False
            Me.cmdRefresh.Enabled = False
        ElseIf Status = "E" Then
            Me.mnuLockPay.Enabled = False
            Me.mnuApproveItem.Enabled = True
            Me.mnuCancelItem.Enabled = True
            Me.MnuReferItem.Enabled = True
            Me.mnuDeclineItem.Enabled = True
            Me.mnuEntered.Enabled = False
            Me.cmdRefresh.Enabled = True

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


    Private Sub SetStatus()
        If Not Me.CollPay Is Nothing Then
            Me.lblStatus.Text = "Status: " & Me.CollPay.PayStatusString
            LockControls((Me.CollPay.PayStatus = "E"), Me.CollPay.PayStatus)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lock the controls dependent on various things 
    ''' </summary>
    ''' <param name="LockState"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	02/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LockControls(ByVal LockState As Boolean, ByVal PayStatus As String)

        If LockState Then
            Me.mnuLockPay.Visible = False
            Me.mnuUnlockPay.Visible = True
            Me.mnuApprovePay.Visible = True

        Else
            Me.mnuLockPay.Visible = True
            Me.mnuUnlockPay.Visible = False
            Me.mnuApprovePay.Visible = False
        End If
        If Not (Me.blnCanEnterPay Or Me.blnUserCanLockPay) Then
            LockState = False           'Always locked if user can't enter pay 
            'Hide menus as well
            Me.mnuLockPay.Visible = False
            Me.mnuUnlockPay.Visible = False

        End If
        Me.Donors.Enabled = LockState
        Me.cbCollOfficer.Enabled = LockState
        Me.dtPayMonth.Enabled = (LockState OrElse PayStatus = "A")
        Me.cbCurrency.Enabled = (LockState OrElse PayStatus = "A")
        Me.gBoxPayOpts.Enabled = (LockState OrElse PayStatus = "A")
        Me.gBoxCollOpts.Enabled = (LockState OrElse PayStatus = "A")
        Me.gBoxTimeEntry.Enabled = (LockState OrElse PayStatus = "A")
        Me.cmdAdd.Enabled = (LockState OrElse PayStatus = "A")


    End Sub

    Private Sub mnuApproveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuApproveItem.Click
        If GetItemType(sender) = ItemType.Expense Then
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            lvCurrentItem.Expense.PayStatus = lvCurrentItem.Expense.GCST_Pay_Approved
            lvCurrentItem.Expense.DoUpdate()
            lvCurrentItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
        Else
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            lvCurrentPayItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Approved
            lvCurrentPayItem.Expense.DoUpdate()
            lvCurrentPayItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
        End If
        SetStatus()
    End Sub

    Private Sub mnuDeclineItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeclineItem.Click
        If GetItemType(sender) = ItemType.Pay Then
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            lvCurrentPayItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Declined
            lvCurrentPayItem.Expense.DoUpdate()
            lvCurrentPayItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        Else
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            lvCurrentItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Declined
            lvCurrentItem.Expense.DoUpdate()
            lvCurrentItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        End If

    End Sub

    Private Sub MnuReferItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuReferItem.Click
        If GetItemType(sender) = ItemType.Pay Then
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            lvCurrentPayItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Referred
            lvCurrentPayItem.Expense.DoUpdate()
            lvCurrentPayItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        Else
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            lvCurrentItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Referred
            lvCurrentItem.Expense.DoUpdate()
            lvCurrentItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        End If
    End Sub

    Private Sub mnuCancelItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancelItem.Click
        If GetItemType(sender) = ItemType.Pay Then
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            lvCurrentPayItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Declined
            lvCurrentPayItem.Expense.DoUpdate()
            lvCurrentPayItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        Else
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            lvCurrentItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Declined
            lvCurrentItem.Expense.DoUpdate()
            lvCurrentItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        End If
    End Sub

    Private Sub mnuEntered_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEntered.Click
        If GetItemType(sender) = ItemType.Pay Then
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            lvCurrentPayItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Entered
            lvCurrentPayItem.Expense.DoUpdate()
            lvCurrentPayItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        Else
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            lvCurrentItem.Expense.PayStatus = lvCurrentPayItem.Expense.GCST_Pay_Entered
            lvCurrentItem.Expense.DoUpdate()
            lvCurrentItem.Refresh()
            Me.CollPay.UpdateItemStatuses()
            SetStatus()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List view containing pay details has been clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	23/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub lvPayItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvPayItems.Click
        lvCurrentPayItem = Nothing
        If Me.lvPayItems.SelectedItems.Count > 0 Then
            Me.lvCurrentPayItem = Me.lvPayItems.SelectedItems.Item(0)
        End If
        If Not Me.lvCurrentPayItem Is Nothing Then
            SetStatusMenus(Me.lvCurrentPayItem.Expense.PayStatus)
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempting to edit pay item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	19/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub lvPayItems_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvPayItems.DoubleClick
    '    If Me.lvCurrentPayItem Is Nothing Then Exit Sub
    '    If Me.blnDrawing Then Exit Sub

    '    Dim frmEdit As New frmExpEntry()
    '    frmEdit.FormType = frmExpEntry.enumExpFormType.PayDetail

    '    frmEdit.CollOfficer = Me.lvCurrentPayItem.CollOfficer
    '    frmEdit.PayItem = Me.lvCurrentPayItem.Expense

    '    If frmEdit.ShowDialog = DialogResult.OK Then
    '        Me.lvCurrentPayItem.Expense = frmEdit.PayItem
    '        Me.lvCurrentPayItem.Expense.Fields.Update(MedscreenLib.CConnection.DbConnection)
    '        Me.lvCurrentPayItem.Refresh()
    '        DrawPayItems(MyCollPayInHoursCost)
    '    End If

    'End Sub

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    End Sub

    Private Sub udKMMiles_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udKMMiles.SelectedItemChanged
        If Me.blnDrawing Then Exit Sub
        If udKMMiles.SelectedIndex = 0 Then
            Label4.Text = "Mileage:"
            If Not Me.MyMileageCost Is Nothing Then
                MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerMile
                MyMileageCost.DoUpdate()
            End If
        Else
            Label4.Text = "Travel:"
            If Not Me.MyMileageCost Is Nothing Then
                MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerKM
                MyMileageCost.DoUpdate()
            End If
        End If
        UpdateMileage()
    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawing = False
    End Sub

    Private Sub ckIgnoreMinimum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckIgnoreMinimum.Click
        'need to update pay
        If Me.blnDrawing Then Exit Sub
        If Me.MyCollPayEntry Is Nothing Then Exit Sub 'We must have a coll pay item 
        If Not Me.MyCollPayInHoursCost Is Nothing AndAlso (Me.MyCollPayEntry.PayMethod = "S" Or Me.MyCollPayEntry.PayMethod = "O") Then
            MyCollPayInHoursCost.ExpenseTypeID = SampleRatePayType()
            UpDateCollPayItem()
            Me.CheckMinumumValue()
        End If
    End Sub

    Private Sub rbOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOther.CheckedChanged

    End Sub

    Private Sub DTWorkingTime_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTWorkingTime.Leave, dtTravelTime.Leave, dtTravelTime.ValueChanged, DTWorkingTime.ValueChanged
        If NewForm Then Exit Sub
        If Not TypeOf sender Is DateTimePicker Then Exit Sub
        Dim dte As DateTimePicker = sender
        If dte.Tag = "PAYWRKTIME" Then
            Me.MyCollPayInHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayWorkingTime)
            If Not MyCollPayInHoursCost Is Nothing Then
                MyCollPayInHoursCost.Quantity = DTWorkingTime.Value.Hour + DTWorkingTime.Value.Minute / 60
                MyCollPayInHoursCost.ItemValue = MyCollPayInHoursCost.ItemPrice * MyCollPayInHoursCost.ItemDiscount * MyCollPayInHoursCost.Quantity
                MyCollPayInHoursCost.DoUpdate()
                UpDateCollPayItem()
                Me.CheckMinumumValue()
            End If
        Else
            Me.MyCollPayOutOfHoursCost = Me.CollPay.PayItems.Item(ExpenseTypeCollections.GCST_PayTravelTime)
            If Not MyCollPayOutOfHoursCost Is Nothing Then
                MyCollPayOutOfHoursCost.Quantity = Me.dtTravelTime.Value.Hour + dtTravelTime.Value.Minute / 60
                MyCollPayOutOfHoursCost.ItemValue = MyCollPayOutOfHoursCost.ItemPrice * MyCollPayOutOfHoursCost.ItemDiscount * MyCollPayOutOfHoursCost.Quantity
                MyCollPayOutOfHoursCost.DoUpdate()
                UpDateCollPayItem()
                Me.CheckMinumumValue()
            End If
        End If
    End Sub

    Private Sub mnuPaystatement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaystatement.Click
        Dim objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        'This needs moving to config file
        objReport = Me.LogInCrystal("\\andrew\solutions\worktemp\PayDetail41.rpt")

        'Set up report formula

        Dim FF As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        For Each FF In objReport.DataDefinition.FormulaFields
            If FF.FormulaName.ToUpper = "{@COLLPAYID}" Then
                FF.Text = """" & Me.CollPay.CollPayID & """"
            End If
        Next

        OutputReport(objReport, "Detailed pay statement " & Me.objCollOfficer.Name & " for " & Me.CollPay.CMidentity)
        objReport.Close()



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

    '''
    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        If Me.CollPay Is Nothing Then Exit Sub
        CollPay.RefreshEntries()
        'redraw entries
        DrawPayItems(MyCollPayInHoursCost)
        RefreshCollCostListView()
        If CollPay.PayMethod = "H" Then
            Me.UpdateTimes()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an additional collecting officer to the collection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	15/08/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddOfficer.Click
        Dim objCS As New frmSelectCollector()
        If objCM Is Nothing Then Exit Sub

        If objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim objColl As Intranet.intranet.jobs.CollectingOfficer = objCS.Officer
            'Clear payment boxes
            Me.txtCostsTotal.Text = ""
            Me.txtCostTotal.Text = ""
            Me.txtTotalExp.Text = ""
            Me.txtMileage.Text = ""
            'Clear mileage value 
            MyMileageCost = Nothing

            'We should now have an officer need to see if there is a payment for them for this collection
            Dim oColl As New Collection()
            oColl.Add(objCM.ID)
            oColl.Add(objColl.OfficerID)
            Dim strCollPayId As String = CConnection.PackageStringList("Lib_collpay.FindPayHeader", oColl)
            If strCollPayId.Trim.Length > 0 Then            'Already exists
                MsgBox("Pay statement exists will load it!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
                Dim ocp As Intranet.collectorpay.CollectorPay = New Intranet.collectorpay.CollectorPay(CollectorPayCollection.PayTypes.COLLECTION)
                ocp.CollPayID = strCollPayId
                ocp.Load(CConnection.DbConnection)
                Me.CollPay = ocp
            Else
                Dim ocp As Intranet.collectorpay.CollectorPay = CollectorPayCollection.CreateHeader(CollectorPayCollection.PayTypes.COLLECTION)
                ocp.CMidentity = objCM.ID
                ocp.CollOfficer = objColl.OfficerID
                ocp.PaymentMonth = DateSerial(Today.Year, Today.Month, 1)
                If Not objCM.CollectingOfficer Is Nothing Then
                    If objCM.CollectionType = "O" Or objCM.CollectionType = "C" Then
                        ocp.PayMethod = "H"
                    Else
                        ocp.PayMethod = objCM.CollectingOfficer.PayMethod
                    End If
                End If
                ocp.DoUpdate()
                'Get the entries 
                Me.CollPay = ocp


            End If
        End If

    End Sub

    Private Sub lvOfficers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvOfficers.SelectedIndexChanged
        If lvOfficers.SelectedItems.Count > 0 Then
            If Not Me.CollPay Is Nothing AndAlso Me.CollPay.PayStatus <> "L" Then
                Me.CollPay.DoUpdate()
            End If
            Dim stroff As String = lvOfficers.SelectedItems.Item(0).Tag
            Dim oColl As New Collection()
            oColl.Add(objCM.ID)
            oColl.Add(stroff)
            Dim strCollPayId As String = CConnection.PackageStringList("Lib_collpay.FindPayHeader", oColl)
            If strCollPayId.Trim.Length > 0 Then
                'Clear payment boxes
                Me.txtCostsTotal.Text = ""
                Me.txtCostTotal.Text = ""
                Me.txtTotalExp.Text = ""
                Me.txtMileage.Text = ""
                'Clear mileage value 
                MyMileageCost = Nothing
                Debug.WriteLine(strCollPayId, "INFO")
                Dim ocp As Intranet.collectorpay.CollectorPay = New Intranet.collectorpay.CollectorPay(CollectorPayCollection.PayTypes.COLLECTION)
                ocp.CollPayID = strCollPayId
                ocp.Load(CConnection.DbConnection)
                Me.CollPay = ocp
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with LineItemChanging
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TAYLOR]	02/10/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub updLineitem_SelectedItemChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles updLineitem.SelectedItemChanged
        If Me.blnDrawing Then Exit Sub
        Dim intIndex As Integer = Me.updLineitem.SelectedIndex + 1
        If intIndex < 1 Then Exit Sub
        Dim strLineItem = Me.cstPaySampleStub & CStr(intIndex).Trim
        If Me.MyCollPayEntry Is Nothing Then Exit Sub
        If MyCollPayInHoursCost Is Nothing Then Exit Sub
        'set expense id
        Me.MyCollPayInHoursCost.ExpenseTypeID = strLineItem
        Me.SamplePay(MyCollPayInHoursCost)
        DrawPayItems(MyCollPayInHoursCost)
        RefreshCollCostListView()

    End Sub

    Function SampleRatePayType() As String
        Dim strLineItem As String = cstPaySampleStub & "1"
        Dim intIndex As Integer = Me.updLineitem.SelectedIndex + 1
        If intIndex >= 1 Then
            strLineItem = Me.cstPaySampleStub & CStr(intIndex).Trim
        End If
        Return strLineItem

    End Function

    Private Sub mnuUnlockPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUnlockPay.Click
        If Me.MyCollPayEntry Is Nothing Then Exit Sub
        'Check to see if we can lock pay 
        If Not Me.blnUserCanLockPay Then
            MsgBox("You are not in the Role which allows you to unlock pay", MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim objCollPayItem As CollPayItem
        'Set header status
        Me.MyCollPayEntry.UnLockPay()
        SetStatus()

    End Sub

    Private Sub mnuApprovePay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuApprovePay.Click
        If Me.MyCollPayEntry Is Nothing Then Exit Sub
        'Check to see if we can lock pay 
        If Not Me.blnUserCanLockPay Then
            MsgBox("You are not in the Role which allows you to set pay status", MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim objCollPayItem As CollPayItem
        'Set header status
        Me.MyCollPayEntry.ApprovePay()
        SetStatus()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Function GetItemType(ByVal sender As Object) As ItemType
        If TypeOf sender Is MenuItem Then
            Dim Parent As MenuItem = sender
            If Not Parent.Parent Is Nothing AndAlso TypeOf Parent.Parent Is ContextMenu Then
                Dim CtxMenu As ContextMenu = Parent.Parent
                Dim cnt As Control = CtxMenu.SourceControl()
                If cnt.Tag = "PAY" Then
                    Return ItemType.Pay
                Else
                    Return ItemType.Expense
                End If
            End If
        End If

    End Function

    Private Sub mnuEditItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditItem.Click
        If GetItemType(sender) = ItemType.Pay Then
            If Me.lvCurrentPayItem Is Nothing Then Exit Sub
            'Me.lvPayItems_DoubleClick(Nothing, Nothing)
        Else
            If Me.lvCurrentItem Is Nothing Then Exit Sub
            ' Me.ListView1_DoubleClick(Nothing, Nothing)
        End If

    End Sub
End Class
