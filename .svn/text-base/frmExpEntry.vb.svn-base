Imports System.Windows.Forms
Imports Intranet.intranet.jobs
Public Class frmExpEntry
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim objExpList As MedscreenLib.Glossary.PhraseCollection
    Dim objPayList As MedscreenLib.Glossary.PhraseCollection
    Private Sub RebindExpenseTypes()
        Me.cboxExpType.DataSource = objExpList
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).SuspendBinding()
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).ResumeBinding()
    End Sub
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        objExpList = New MedscreenLib.Glossary.PhraseCollection("select e.EXPTYPEID,l.DESCRIPTION from lineitems l,expense_type e where l.itemcode = e.itemcode and e.expensegroup = 'EXPENSE' AND l.SERVICE = 'COLLECPAY'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        objExpList.Load()
        objPayList = New MedscreenLib.Glossary.PhraseCollection("select e.EXPTYPEID,l.DESCRIPTION from lineitems l,expense_type e where l.itemcode = e.itemcode and e.expensegroup = 'PAY' AND l.SERVICE = 'COLLECPAY' order by description", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        objPayList.Load()

        'Dim objExpList As New MedscreenLib.Glossary.PhraseCollection("select   ITEMCODE,DESCRIPTION from lineitems where TYPE = 'EXPENSE' AND SERVICE = 'COLLECPAY'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        'objExpList.Load()
        Dim objPayStatus As New MedscreenLib.Glossary.PhraseCollection("PAY_STATUS")
        objPayStatus.Load()
        Dim objExpClass As New MedscreenLib.Glossary.PhraseCollection("OFFEXPCLAS")
        objExpClass.Load()

        RebindExpenseTypes()
        Me.cboxExpType.ValueMember = "PhraseID"
        Me.cboxExpType.DisplayMember = "PhraseText"
        cbPayItems.DataSource = objPayList
        Me.cbPayItems.ValueMember = "PhraseID"
        Me.cbPayItems.DisplayMember = "PhraseText"
        Me.cbExpenseClass.DataSource = objExpClass
        Me.cbExpenseClass.ValueMember = "PhraseID"
        Me.cbExpenseClass.DisplayMember = "PhraseText"


        Me.cbPayStatus.DataSource = objPayStatus
        Me.cbPayStatus.ValueMember = "PhraseID"
        Me.cbPayStatus.DisplayMember = "PhraseText"

        Dim objCurr As New MedscreenLib.Glossary.PhraseCollection("Select Currencyid,description from currency", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        objCurr.Load()
        Me.cbCurrency.DataSource = objCurr
        Me.cbCurrency.ValueMember = "PhraseID"
        Me.cbCurrency.DisplayMember = "PhraseText"

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
    Friend WithEvents butReturn As System.Windows.Forms.Button
    Friend WithEvents butAddExp As System.Windows.Forms.Button
    Friend WithEvents cboxExpType As System.Windows.Forms.ComboBox
    Friend WithEvents lblExpType As System.Windows.Forms.Label
    Friend WithEvents lblItemValue As System.Windows.Forms.Label
    Friend WithEvents txtExpValue As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpComments As System.Windows.Forms.GroupBox
    Friend WithEvents txtExpComments As System.Windows.Forms.TextBox
    Friend WithEvents txtPrice As System.Windows.Forms.TextBox
    Friend WithEvents lblPrice As System.Windows.Forms.Label
    Friend WithEvents Quantity As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblCurrComments As System.Windows.Forms.Label
    Friend WithEvents lblOCRtext As System.Windows.Forms.Label
    Friend WithEvents dtPayMonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPaymentMonth As System.Windows.Forms.Label
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbPayStatus As System.Windows.Forms.ComboBox
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents grpAddPay As System.Windows.Forms.GroupBox
    Friend WithEvents ckTransfer As System.Windows.Forms.CheckBox
    Friend WithEvents ckCoordinator As System.Windows.Forms.CheckBox
    Friend WithEvents cbPayItems As System.Windows.Forms.ComboBox
    Friend WithEvents cbExpenseClass As System.Windows.Forms.ComboBox
    Friend WithEvents lblExpClass As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExpEntry))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.butAddExp = New System.Windows.Forms.Button
        Me.butReturn = New System.Windows.Forms.Button
        Me.lblExpType = New System.Windows.Forms.Label
        Me.cboxExpType = New System.Windows.Forms.ComboBox
        Me.lblItemValue = New System.Windows.Forms.Label
        Me.txtExpValue = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Quantity = New System.Windows.Forms.NumericUpDown
        Me.grpComments = New System.Windows.Forms.GroupBox
        Me.txtExpComments = New System.Windows.Forms.TextBox
        Me.txtPrice = New System.Windows.Forms.TextBox
        Me.lblPrice = New System.Windows.Forms.Label
        Me.lblCurrComments = New System.Windows.Forms.Label
        Me.lblOCRtext = New System.Windows.Forms.Label
        Me.dtPayMonth = New System.Windows.Forms.DateTimePicker
        Me.lblPaymentMonth = New System.Windows.Forms.Label
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbPayStatus = New System.Windows.Forms.ComboBox
        Me.cbCurrency = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.grpAddPay = New System.Windows.Forms.GroupBox
        Me.ckCoordinator = New System.Windows.Forms.CheckBox
        Me.ckTransfer = New System.Windows.Forms.CheckBox
        Me.cbPayItems = New System.Windows.Forms.ComboBox
        Me.cbExpenseClass = New System.Windows.Forms.ComboBox
        Me.lblExpClass = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        CType(Me.Quantity, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpComments.SuspendLayout()
        Me.grpAddPay.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.butAddExp)
        Me.Panel1.Controls.Add(Me.butReturn)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 421)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(440, 64)
        Me.Panel1.TabIndex = 0
        '
        'butAddExp
        '
        Me.butAddExp.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.butAddExp.Image = CType(resources.GetObject("butAddExp.Image"), System.Drawing.Image)
        Me.butAddExp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.butAddExp.Location = New System.Drawing.Point(120, 16)
        Me.butAddExp.Name = "butAddExp"
        Me.butAddExp.Size = New System.Drawing.Size(72, 32)
        Me.butAddExp.TabIndex = 2
        Me.butAddExp.Text = "&Ok"
        Me.butAddExp.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'butReturn
        '
        Me.butReturn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butReturn.Image = CType(resources.GetObject("butReturn.Image"), System.Drawing.Image)
        Me.butReturn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.butReturn.Location = New System.Drawing.Point(208, 16)
        Me.butReturn.Name = "butReturn"
        Me.butReturn.Size = New System.Drawing.Size(72, 32)
        Me.butReturn.TabIndex = 0
        Me.butReturn.Text = "&Cancel"
        Me.butReturn.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExpType
        '
        Me.lblExpType.Location = New System.Drawing.Point(8, 60)
        Me.lblExpType.Name = "lblExpType"
        Me.lblExpType.Size = New System.Drawing.Size(88, 24)
        Me.lblExpType.TabIndex = 1
        Me.lblExpType.Text = "Expense:"
        '
        'cboxExpType
        '
        Me.cboxExpType.Items.AddRange(New Object() {"type1", "type2", "type3"})
        Me.cboxExpType.Location = New System.Drawing.Point(128, 60)
        Me.cboxExpType.Name = "cboxExpType"
        Me.cboxExpType.Size = New System.Drawing.Size(184, 21)
        Me.cboxExpType.TabIndex = 2
        '
        'lblItemValue
        '
        Me.lblItemValue.Location = New System.Drawing.Point(8, 139)
        Me.lblItemValue.Name = "lblItemValue"
        Me.lblItemValue.Size = New System.Drawing.Size(112, 16)
        Me.lblItemValue.TabIndex = 3
        Me.lblItemValue.Text = "Value:"
        '
        'txtExpValue
        '
        Me.txtExpValue.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtExpValue.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtExpValue.Enabled = False
        Me.txtExpValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExpValue.Location = New System.Drawing.Point(128, 138)
        Me.txtExpValue.Name = "txtExpValue"
        Me.txtExpValue.Size = New System.Drawing.Size(120, 13)
        Me.txtExpValue.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 92)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 24)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Quantity:"
        '
        'Quantity
        '
        Me.Quantity.Location = New System.Drawing.Point(264, 92)
        Me.Quantity.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.Quantity.Name = "Quantity"
        Me.Quantity.Size = New System.Drawing.Size(48, 20)
        Me.Quantity.TabIndex = 6
        Me.Quantity.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'grpComments
        '
        Me.grpComments.Controls.Add(Me.txtExpComments)
        Me.grpComments.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpComments.Location = New System.Drawing.Point(0, 309)
        Me.grpComments.Name = "grpComments"
        Me.grpComments.Size = New System.Drawing.Size(440, 112)
        Me.grpComments.TabIndex = 9
        Me.grpComments.TabStop = False
        Me.grpComments.Text = "Comments Section"
        '
        'txtExpComments
        '
        Me.txtExpComments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtExpComments.Location = New System.Drawing.Point(3, 16)
        Me.txtExpComments.Multiline = True
        Me.txtExpComments.Name = "txtExpComments"
        Me.txtExpComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtExpComments.Size = New System.Drawing.Size(434, 93)
        Me.txtExpComments.TabIndex = 0
        '
        'txtPrice
        '
        Me.txtPrice.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtPrice.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPrice.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrice.Location = New System.Drawing.Point(128, 115)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(96, 13)
        Me.txtPrice.TabIndex = 11
        '
        'lblPrice
        '
        Me.lblPrice.Location = New System.Drawing.Point(8, 115)
        Me.lblPrice.Name = "lblPrice"
        Me.lblPrice.Size = New System.Drawing.Size(112, 24)
        Me.lblPrice.TabIndex = 10
        Me.lblPrice.Text = "Unit Price:"
        '
        'lblCurrComments
        '
        Me.lblCurrComments.ForeColor = System.Drawing.Color.Red
        Me.lblCurrComments.Location = New System.Drawing.Point(8, 8)
        Me.lblCurrComments.Name = "lblCurrComments"
        Me.lblCurrComments.Size = New System.Drawing.Size(280, 16)
        Me.lblCurrComments.TabIndex = 12
        '
        'lblOCRtext
        '
        Me.lblOCRtext.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOCRtext.Location = New System.Drawing.Point(8, 8)
        Me.lblOCRtext.Name = "lblOCRtext"
        Me.lblOCRtext.Size = New System.Drawing.Size(128, 24)
        Me.lblOCRtext.TabIndex = 13
        Me.lblOCRtext.Text = "On-Call Rota Entry"
        '
        'dtPayMonth
        '
        Me.dtPayMonth.CustomFormat = "MMMM-yyyy"
        Me.dtPayMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPayMonth.Location = New System.Drawing.Point(136, 172)
        Me.dtPayMonth.Name = "dtPayMonth"
        Me.dtPayMonth.Size = New System.Drawing.Size(112, 20)
        Me.dtPayMonth.TabIndex = 14
        '
        'lblPaymentMonth
        '
        Me.lblPaymentMonth.Location = New System.Drawing.Point(8, 172)
        Me.lblPaymentMonth.Name = "lblPaymentMonth"
        Me.lblPaymentMonth.Size = New System.Drawing.Size(96, 16)
        Me.lblPaymentMonth.TabIndex = 16
        Me.lblPaymentMonth.Text = "Payment Month:"
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.MonthCalendar1.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.MonthCalendar1.Location = New System.Drawing.Point(104, 166)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 15
        Me.MonthCalendar1.TitleBackColor = System.Drawing.SystemColors.InactiveBorder
        Me.MonthCalendar1.TitleForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.MonthCalendar1.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 203)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 24)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Status:"
        '
        'cbPayStatus
        '
        Me.cbPayStatus.Location = New System.Drawing.Point(128, 203)
        Me.cbPayStatus.Name = "cbPayStatus"
        Me.cbPayStatus.Size = New System.Drawing.Size(184, 21)
        Me.cbPayStatus.TabIndex = 18
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(320, 115)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(121, 21)
        Me.cbCurrency.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(216, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 24)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Currency:"
        '
        'grpAddPay
        '
        Me.grpAddPay.Controls.Add(Me.ckCoordinator)
        Me.grpAddPay.Controls.Add(Me.ckTransfer)
        Me.grpAddPay.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpAddPay.Location = New System.Drawing.Point(0, 245)
        Me.grpAddPay.Name = "grpAddPay"
        Me.grpAddPay.Size = New System.Drawing.Size(440, 64)
        Me.grpAddPay.TabIndex = 21
        Me.grpAddPay.TabStop = False
        Me.grpAddPay.Text = "Region Manager Payments"
        '
        'ckCoordinator
        '
        Me.ckCoordinator.Location = New System.Drawing.Point(224, 24)
        Me.ckCoordinator.Name = "ckCoordinator"
        Me.ckCoordinator.Size = New System.Drawing.Size(200, 24)
        Me.ckCoordinator.TabIndex = 1
        Me.ckCoordinator.Text = "Coordinator Fees"
        '
        'ckTransfer
        '
        Me.ckTransfer.Location = New System.Drawing.Point(16, 24)
        Me.ckTransfer.Name = "ckTransfer"
        Me.ckTransfer.Size = New System.Drawing.Size(152, 24)
        Me.ckTransfer.TabIndex = 0
        Me.ckTransfer.Text = "Transfer Fees"
        '
        'cbPayItems
        '
        Me.cbPayItems.Items.AddRange(New Object() {"type1", "type2", "type3"})
        Me.cbPayItems.Location = New System.Drawing.Point(128, 60)
        Me.cbPayItems.Name = "cbPayItems"
        Me.cbPayItems.Size = New System.Drawing.Size(184, 21)
        Me.cbPayItems.TabIndex = 22
        '
        'cbExpenseClass
        '
        Me.cbExpenseClass.Items.AddRange(New Object() {"type1", "type2", "type3"})
        Me.cbExpenseClass.Location = New System.Drawing.Point(128, 32)
        Me.cbExpenseClass.Name = "cbExpenseClass"
        Me.cbExpenseClass.Size = New System.Drawing.Size(184, 21)
        Me.cbExpenseClass.TabIndex = 24
        '
        'lblExpClass
        '
        Me.lblExpClass.Location = New System.Drawing.Point(8, 32)
        Me.lblExpClass.Name = "lblExpClass"
        Me.lblExpClass.Size = New System.Drawing.Size(88, 24)
        Me.lblExpClass.TabIndex = 23
        Me.lblExpClass.Text = "Expense Type:"
        '
        'frmExpEntry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 485)
        Me.Controls.Add(Me.cbExpenseClass)
        Me.Controls.Add(Me.lblExpClass)
        Me.Controls.Add(Me.cbPayItems)
        Me.Controls.Add(Me.grpAddPay)
        Me.Controls.Add(Me.cbCurrency)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbPayStatus)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtPayMonth)
        Me.Controls.Add(Me.lblPaymentMonth)
        Me.Controls.Add(Me.MonthCalendar1)
        Me.Controls.Add(Me.lblOCRtext)
        Me.Controls.Add(Me.lblCurrComments)
        Me.Controls.Add(Me.txtPrice)
        Me.Controls.Add(Me.lblPrice)
        Me.Controls.Add(Me.grpComments)
        Me.Controls.Add(Me.Quantity)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtExpValue)
        Me.Controls.Add(Me.lblItemValue)
        Me.Controls.Add(Me.cboxExpType)
        Me.Controls.Add(Me.lblExpType)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmExpEntry"
        Me.Text = "Expense Pay Entry"
        Me.Panel1.ResumeLayout(False)
        CType(Me.Quantity, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpComments.ResumeLayout(False)
        Me.grpComments.PerformLayout()
        Me.grpAddPay.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private MyCollPayExpense As CollPay
    Private MyCollOfficer As Intranet.intranet.jobs.CollectingOfficer
    Private myFormType As enumExpFormType = enumExpFormType.stdExpEntryFormat

    Public Enum enumExpFormType
        stdExpEntryFormat
        OnCallRotaEntry
        OtherPayItem
        PayDetail
    End Enum

    Private Sub RebindPaytypes()
        Me.cboxExpType.DataSource = objPayList
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).SuspendBinding()
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).ResumeBinding()
    End Sub
    Public Property FormType() As enumExpFormType
        Get
            Return Me.myFormType
        End Get
        Set(ByVal Value As enumExpFormType)
            Me.myFormType = Value
            Select Case myFormType
                Case enumExpFormType.stdExpEntryFormat
                    RebindExpenseTypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.cbPayItems.Hide()
                    Me.lblOCRtext.Visible = False
                    Me.Label1.Text = "Quantity"
                    RebindExpenseTypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.cbExpenseClass.Show()
                    Me.lblExpType.Show()
                    Me.lblExpType.Text = "Expense:"
                    Me.lblExpClass.Show()
                    Me.grpAddPay.Hide()
                Case enumExpFormType.OnCallRotaEntry
                    RebindPaytypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.cbPayItems.Show()
                    Me.cboxExpType.Visible = False
                    Me.txtExpComments.Visible = True
                    Me.grpComments.Visible = True
                    Me.lblOCRtext.Visible = True
                    Me.Label1.Text = "No. Of Shifts."
                    Me.grpAddPay.Show()
                    Me.cbExpenseClass.Hide()
                    Me.lblExpType.Hide()
                Case enumExpFormType.OtherPayItem
                    RebindPaytypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.cboxExpType.Visible = True
                    Me.txtExpComments.Visible = True
                    Me.grpComments.Visible = True
                    Me.lblOCRtext.Visible = False
                    Me.Label1.Text = "No. Of Days/Half Days"
                    Me.grpAddPay.Hide()
                    Me.cbPayItems.Hide()
                    Me.cbExpenseClass.Hide()
                    Me.lblExpType.Show()
                    Me.lblExpType.Text = "Pay for "
                    Me.lblExpClass.Hide()
                Case enumExpFormType.PayDetail
                    RebindPaytypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.cbPayItems.Show()
                    Me.lblOCRtext.Visible = False
                    Me.Label1.Text = "Quantity"
                    RebindExpenseTypes()
                    Me.cboxExpType.ValueMember = "PhraseID"
                    Me.cboxExpType.DisplayMember = "PhraseText"
                    Me.grpAddPay.Hide()
                    Me.cbExpenseClass.Hide()
                    Me.lblExpType.Show()
                    Me.lblExpType.Text = "Pay for "
                    Me.lblExpClass.Hide()

            End Select
        End Set
    End Property

    Public WriteOnly Property CollOfficer() As Intranet.intranet.jobs.CollectingOfficer
        Set(ByVal Value As Intranet.intranet.jobs.CollectingOfficer)
            MyCollOfficer = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collector pay entry item
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	19/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property PayItem() As collpay
        Get
            Return Me.MyCollPayExpense
        End Get
        Set(ByVal Value As CollPay)
            MyCollPayExpense = Value
            If MyCollPayExpense Is Nothing Then Exit Property
            FillForm()
        End Set
    End Property

    Private myDiscount As Double = 0
    Public Property Discount() As Double
        Get
            Return myDiscount
        End Get
        Set(ByVal Value As Double)
            myDiscount = Value
        End Set
    End Property

    Private Sub FillForm()
        If MyCollPayExpense Is Nothing Then Exit Sub
        Dim objCollExp As Intranet.intranet.jobs.CMJob

        'MyCollPayExpense.CollCurrency(MyCollOfficer.PayCurrency)

        With MyCollPayExpense
            If .Quantity < 1 Then ' make the default Quantity value to equal 1 
                .Quantity = 1
            End If

            Quantity.Value = .Quantity
            txtExpComments.Text = .Comments
            If .ItemPrice * .ItemDiscount > 0 Then
                txtPrice.Text = .ItemPrice * .ItemDiscount
            Else
                txtPrice.Text = .CollTotal
            End If

            If .ItemPrice * .Quantity * .ItemDiscount > 0 Then
                txtExpValue.Text = .ItemPrice * .Quantity * .ItemDiscount
            Else
                txtExpValue.Text = .CollTotal
            End If
            Dim strExpClass As String
            If FormType = enumExpFormType.stdExpEntryFormat AndAlso .ExpenseTypeID.Trim.Length > 0 Then
                strExpClass = MedscreenLib.CConnection.PackageStringList("Lib_collpay.ExpenseClass", .ExpenseTypeID)
                If strExpClass IsNot Nothing Then Me.cbExpenseClass.SelectedValue = strExpClass
            End If
            Me.cbCurrency.SelectedValue = MyCollOfficer.PayCurrency
            Me.cboxExpType.SelectedValue = .ExpenseTypeID
            Me.cbPayStatus.SelectedValue = .PayStatus

            Me.txtExpValue.Text = MyCollPayExpense.CollCurrency(CDbl(.ItemPrice * .Quantity), MyCollOfficer.PayCurrency)
            Me.Discount = 0
            If MyCollOfficer.PayForItem(.ItemCode) > 0 Then
                txtPrice.Text = MyCollOfficer.PayForItem(.ItemCode, Me.dtPayMonth.Value)
                .ItemPrice = MyCollOfficer.PriceForItem(.ItemCode, Me.dtPayMonth.Value)
                .ItemDiscount = MyCollOfficer.DiscountONItem(.ItemCode, Me.dtPayMonth.Value)
                Me.Discount = .ItemDiscount
                txtPrice.Enabled = False ' DO NOT ALLOW USER TO AMEND DATA FROM PRICES TABLE
            Else
                txtPrice.Enabled = True ' ENABLE CONTRAOL TO ALLOW USER TO AMEND PRICE
                txtPrice.BackColor = System.Drawing.SystemColors.Window
                txtPrice.BorderStyle = BorderStyle.Fixed3D

                If MyCollPayExpense.ItemPrice.ToString.Trim = "-1" Then
                    txtPrice.Text = ""
                Else
                    txtPrice.Text = MyCollPayExpense.ItemPrice
                End If
            End If
            tmpExpTypeId = MyCollPayExpense.ExpenseTypeID
            Me.cboxExpType.SelectedValue = MyCollPayExpense.ExpenseTypeID
            cbPayItems.SelectedValue = MyCollPayExpense.ExpenseTypeID
            txtPrice.Refresh()
        End With
    End Sub

    Private tmpExpTypeId As String = ""
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><modified>30-Nov-2011 12:00 collection pay value zeroed out for expense items</modified><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
    Private Sub GetForm()
        If MyCollPayExpense Is Nothing Then Exit Sub
        Dim objCollExp As Intranet.intranet.jobs.CMJob

        With MyCollPayExpense
            If .ItemPrice <> "-1" Then
                txtPrice.Enabled = True
                txtPrice.Refresh()
            End If
            .Quantity = Quantity.Value
            .Comments = txtExpComments.Text

            .ExpenseTypeID = Me.cboxExpType.SelectedValue
            If .ExpenseTypeID.Trim.Length = 0 Then
                .ExpenseTypeID = tmpExpTypeId
            End If
            .PayStatus = Me.cbPayStatus.SelectedValue
            '.Currency = Me.cbCurrency.SelectedValue
            If Me.FormType <> enumExpFormType.OtherPayItem Then
                .ItemPrice = Me.txtPrice.Text
            End If
            .ItemValue = Quantity.Value * Me.txtPrice.Text
            .PayMonth = Me.dtPayMonth.Value
            Dim strSubClass As String = MedscreenLib.CConnection.PackageStringList("Lib_collPay.GetExpenseSubClass", .ExpenseTypeID)
            If strSubClass IsNot Nothing AndAlso strSubClass.Trim.Length > 0 Then
                .PayType = strSubClass
                If strSubClass = CollPay.CST_PAYTYPE_DEDUCTION Then
                    .Collection = "DED" & Mid(.CID, 4)
                    .TotalCost = .ItemValue
                    .ExpensesTotal = 0
                Else
                    .ExpensesTotal = .ItemValue
                End If
            Else
                'deal with special pay items
                If .PayType = CollPay.CST_PAYTYPE_OCR Or .PayType = CollPay.CST_PAYTYPE_AREA Or .PayType = CollPay.CST_PAYTYPE_TRANSFER Then
                    .TotalCost = .ItemValue
                    .ExpensesTotal = 0
                Else
                    .PayType = CollPay.CST_PAYTYPE_OTHERCOSTSEXPENSES
                    .ExpensesTotal = .ItemValue
                    .TotalCost = .ItemValue
                    .CollTotal = 0         '                                         'Zero out collection costs
                    'Modified : 30-Nov-2011 by CONCATENO\Taylor code added to zero out collection costs for expenses
                End If

            End If
        End With
    End Sub

    Private Sub txtExpValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExpValue.TextChanged
    End Sub

    Private Sub butReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butReturn.Click

    End Sub

    Private Sub butAddExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butAddExp.Click
        Me.GetForm()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if transfer checkbox is ticked
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/01/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property AddTransferFee() As Boolean
        Get
            Return (Me.ckTransfer.Checked)
        End Get
        Set(ByVal Value As Boolean)
            Me.ckTransfer.Checked = Value
        End Set
    End Property

    Public Property AddCoordinatorFee() As Boolean
        Get
            Return Me.ckCoordinator.Checked
        End Get
        Set(ByVal Value As Boolean)
            Me.ckCoordinator.Checked = Value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Quantity_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Quantity.KeyUp
        If Quantity.Text.Trim.Length > 0 Then
            Me.txtExpValue.Text = CDbl(Me.txtPrice.Text) * CDbl(Me.Quantity.Text)
            Me.Refresh()
            Me.txtExpValue.Text = MyCollPayExpense.CollCurrency(CDbl(txtExpValue.Text), MyCollOfficer.PayCurrency)
            Me.Refresh()
        End If

    End Sub

    Private Sub Quantity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Quantity.KeyDown
        If Quantity.Text.Trim.Length > 0 Then
            Me.txtExpValue.Text = Me.txtPrice.Text * Me.Quantity.Text
            Me.Refresh()
            Me.txtExpValue.Text = MyCollPayExpense.CollCurrency(CDbl(txtExpValue.Text), MyCollOfficer.PayCurrency)
            Me.Refresh()
        End If

    End Sub

    Private Sub Quantity_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Quantity.Click, Quantity.ValueChanged, Quantity.Leave
        If Quantity.Text.Trim.Length > 0 AndAlso Me.txtPrice.Text.Trim.Length > 0 Then
            Try
                Me.txtExpValue.Text = CStr(CDbl(Me.txtPrice.Text) * CInt(Me.Quantity.Text))
                Me.Refresh()
                Me.txtExpValue.Text = MyCollPayExpense.CollCurrency(CDbl(txtExpValue.Text), MyCollOfficer.PayCurrency)
                Me.Refresh()
            Catch
            End Try
        End If
    End Sub

    Private Sub cboxExpType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If MyCollPayExpense Is Nothing Then Exit Sub
        If MyCollOfficer Is Nothing Then Exit Sub
        MyCollPayExpense.ExpenseTypeID = Me.cboxExpType.SelectedValue
        MyCollPayExpense.DoUpdate()
        Dim dblPayForItem As Double = MyCollOfficer.PayForItem(MyCollPayExpense.ExpenseTypeID)
        If dblPayForItem > 0 Then
            txtPrice.Text = dblPayForItem
            txtPrice.Enabled = False ' DO NOT ALLOW USER TO AMEND DATA FROM PRICES TABLE
            Me.txtExpValue.Text = MyCollPayExpense.CollCurrency(CDbl(dblPayForItem) * Me.Quantity.Value, MyCollOfficer.PayCurrency)
        Else
            txtPrice.Enabled = True ' ENABLE CONTRAOL TO ALLOW USER TO AMEND PRICE
            txtPrice.BackColor = System.Drawing.SystemColors.Window
            txtPrice.BorderStyle = BorderStyle.Fixed3D

            If MyCollPayExpense.ItemPrice.ToString.Trim = "-1" Then
                txtPrice.Text = ""
            Else
                txtPrice.Text = MyCollPayExpense.ItemPrice
            End If
        End If
        txtPrice.Refresh()

    End Sub



    Private Sub cbExpenseClass_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbExpenseClass.SelectedValueChanged
        If MyCollPayExpense Is Nothing Then Exit Sub
        If MyCollOfficer Is Nothing Then Exit Sub
        objExpList = New MedscreenLib.Glossary.PhraseCollection("select e.EXPTYPEID,l.DESCRIPTION from lineitems l,expense_type e where l.itemcode = e.itemcode and e.expensegroup = 'EXPENSE' AND l.SERVICE = 'COLLECPAY' and e.expenseclass = '" & Me.cbExpenseClass.SelectedValue & "'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        objExpList.Load()
        Me.cboxExpType.DataSource = objExpList
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).SuspendBinding()
        Me.cboxExpType.BindingContext(Me.cboxExpType.DataSource).ResumeBinding()
        Me.cboxExpType.ValueMember = "PhraseID"
        Me.cboxExpType.DisplayMember = "PhraseText"

    End Sub

    
End Class
