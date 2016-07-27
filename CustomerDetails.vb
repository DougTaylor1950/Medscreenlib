''''$Revision: 1.0 $ <para/>
''''$Author: taylor $ <para/>
''''$Date: 2006-01-17 14:34:21+00 $ <para/>
''''$Log: CustomerDetails.vb,v $ck
''''Revision 1.0  2006-01-17 14:34:21+00  taylor
''''Checked afte rmaking minor mods to fix options screens
'''' <para/>


Imports intranet.intranet
Imports Intranet.intranet.customerns
Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports System.ComponentModel
Imports MedscreenCommonGui
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCustomerDetails
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for maintaining customers
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [05-Dec-2006]</date>
'''<Action>Added panels to deal with attendence fee charging and minimum charges</Action></revision>
''' <revision><Author>[taylor]</Author><date> [17-Jan-2006]</date>
'''<Action>Resizing of options fixed</Action></revision>
''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action>Two invoicing addresses added</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCustomerDetails
    Inherits System.Windows.Forms.Form
    Private Const STR_NOMRO As String = "NONE"

    Private tabPages As New ArrayList()
    Private myPricebooks As MedscreenLib.Glossary.PhraseCollection
    Private myParents As MedscreenLib.Glossary.PhraseCollection

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'Logging = True              'Swicth on timing logging for testing 
        LogTiming("Opening Form calling new")

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Try  ' Protecting form initialise
            'Me.dtContractDate.Name = "dtNextActionDate"
            'Add any initialization after the InitializeComponent() call
            'Me.AddressControl1.Glossaries = ModTpanel.Glossaries
            'Set up Confirmation Method 
            LogTiming("Opening Form initialised")
            'Set customer type combo box up
            Me.CBCustomerType.DataSource = MedscreenLib.Glossary.Glossary.AccountType
            Me.CBCustomerType.DisplayMember = "PhraseText"
            Me.CBCustomerType.ValueMember = "PhraseID"

            'set industry type combo box up
            Me.cbIndustry.DataSource = MedscreenLib.Glossary.Glossary.Industry
            Me.cbIndustry.DisplayMember = "PhraseText"
            Me.cbIndustry.ValueMember = "PhraseID"

            'Set reporting method combobox up
            Me.cbReportMethod.DataSource = MedscreenLib.Glossary.Glossary.ReportMethod
            Me.cbReportMethod.DisplayMember = "PhraseText"
            Me.cbReportMethod.ValueMember = "PhraseID"

            'Set fax format combobox up
            Me.cbFaxFormat.DataSource = MedscreenLib.Glossary.Glossary.FaxFormat
            Me.cbFaxFormat.DisplayMember = "PhraseText"
            Me.cbFaxFormat.ValueMember = "PhraseID"

            'Set platinum account type combobox up
            Me.cbPlatAccType.DataSource = MedscreenLib.Glossary.Glossary.Platinum
            Me.cbPlatAccType.DisplayMember = "PhraseText"
            Me.cbPlatAccType.ValueMember = "PhraseID"

            'Set invoice frequency combobox up
            Me.cbInvFreq.DataSource = MedscreenLib.Glossary.Glossary.Invoicing
            Me.cbInvFreq.DisplayMember = "PhraseText"
            Me.cbInvFreq.ValueMember = "PhraseID"

            'Set payinterval combobox up
            Me.cbPayInterval.DataSource = MedscreenLib.Glossary.Glossary.PaymentIntervals
            Me.cbPayInterval.DisplayMember = "PhraseText"
            Me.cbPayInterval.ValueMember = "PhraseID"

            'SetMRO combobox up
            Me.cbMROID.DataSource = Intranet.intranet.Support.cSupport.MROs
            Me.cbMROID.DisplayMember = "MROName"
            Me.cbMROID.ValueMember = "MROID"

            LogTiming("Combo boxes done")
            'Me.cbLineItems.DataSource = CCToolAppMessages.RoutineLineItems
            'Me.cbLineItems.DisplayMember = "value"
            'Me.cbLineItems.ValueMember = "key"

            'Get lineitems from config file
            Dim i As Integer
            For i = 0 To CCToolAppMessages.RoutineLineItems.Count - 1
                Me.cbLineItems.Items.Add(CCToolAppMessages.RoutineLineItems.Item(i))
            Next

            'Dim objCustStatus As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("CUST_STAT")
            'objCustStatus.Load()
            Me.cbStatus.PhraseType = "CUST_STAT"
            'Me.cbStatus.DisplayMember = "PhraseText"
            'Me.cbStatus.ValueMember = "PhraseID"

            LogTiming("Routine line items done")
            'Set up programme manager and salesperson comboboxes from roles
            Me.cbProgMan.DataSource = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.RoleHolders("PROGRAMMEMANAGER")
            Me.cbSalesPerson.DataSource = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.RoleHolders("SALESPERSON")
            Me.cbSalesPerson.DisplayMember = "User"
            Me.cbSalesPerson.ValueMember = "User"
            Me.cbProgMan.DisplayMember = "User"
            Me.cbProgMan.ValueMember = "User"

            'Set form to be maximised
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

            'Check authority levels for options as check boxes

            Dim objOption As MedscreenLib.Glossary.Options

            Me.ckStopCerts.Enabled = True
            Me.CKStopSamples.Enabled = True

            LogTiming("Set user")
            'Code modified 11-Apr-2007 07:09 by taylor
            Try  ' Protecting Authority checks
                'Dim intAuthority As Integer = 0
                'If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
                '    intAuthority = MedscreenLib.Glossary.Glossary.CurrentSMUser.OptionAuthority
                'End If
                '  check stop sample authority level, user has to have a higher authority level than 
                ' authority associated with option, this code should allow levels to change
                objOption = MedscreenLib.Glossary.Glossary.Options.Item("STOP_SAMP")
                If Not objOption.UserCanEdit(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity) Then
                    Me.CKStopSamples.Enabled = False
                End If
                'Same check with stop certs
                objOption = MedscreenLib.Glossary.Glossary.Options.Item("STOP_CERT")
                If Not objOption.UserCanEdit(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity) Then
                    Me.ckStopCerts.Enabled = False
                End If
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, True, "frmCustomerDetails-New-Authority checking")
            Finally
            End Try
            For i = 0 To Me.TabOptions.TabPages.Count - 1
                Me.tabPages.Add(Me.TabOptions.TabPages(i))
            Next

            ' Dim i As Integer
            For i = 0 To 20
                ColumnDirection(i) = 1
            Next

            'Setup price security
            Me.blnDisableAll = True
            Dim objUser As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.CurrentSMUser
            If Not objUser Is Nothing Then
                If MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserHasRole(objUser.Identity, "MED_EDIT_PRICES") Then
                    Me.blnDisableAll = False
                End If
            End If
            LogTiming("About to initialise panels")

            'setup pricing panels 
            Me.pnlRoutine.Service = "ROUTINE"
            Me.pnlRoutine.ContextMenu = Me.ctxAddService

            Me.pnlAO.Service = "ANALONLY"
            Me.pnlAO.ContextMenu = Me.ctxAddService

            Me.pnlCallOut.Service = "CALLOUT"
            Me.pnlCallOut.ContextMenu = Me.ctxAddService


            Me.pnlExternal.Service = "EXTERNAL"
            Me.pnlExternal.ContextMenu = Me.ctxAddService

            Me.pnlFixedSite.Service = "FIXED"
            Me.pnlFixedSite.ContextMenu = Me.ctxAddService

            Me.udCallOut.SelectedIndex = 0
            Me.udRoutine.SelectedIndex = 0

            Me.cbCurrency.DataSource = MedscreenLib.Glossary.Glossary.Currencies
            Me.cbCurrency.DisplayMember = "Description"
            Dim intCurrLength As Integer = System.Configuration.ConfigurationSettings.AppSettings("CurrencyCodeLength")
            If intCurrLength = 3 Then
                Me.cbCurrency.ValueMember = "CurrencyId"
            Else
                Me.cbCurrency.ValueMember = "SingleCharId"
            End If

            '<Added code modified 23-Apr-2007 06:33 by taylor> 
            'Me.cbVolScheme.DataSource = MedscreenLib.Glossary.Glossary.VolumeDiscountSchemes
            'Me.cbVolScheme.DisplayMember = "Description"
            'Me.cbVolScheme.ValueMember = "DiscountSchemeId"

            '</Added code modified 23-Apr-2007 06:33 by taylor> 

            'objContTypePhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_TYP")
            'objContTypePhrase.Load()
            'objContMethodPhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_METH")
            'objContMethodPhrase.Load()
            objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("REP_OPTION")
            objRepOptionPhrase.Load()


            '<Added code modified 05-Jul-2007 09:18 by taylor> 
            'objContStatusPhrase = New MedscreenLib.Glossary.PhraseCollection("CONT_STAT")
            'objContStatusPhrase.Load()

            objDivisionPhrase = New MedscreenLib.Glossary.PhraseCollection("DIVISION")
            objDivisionPhrase.Load()

            Me.cbContactType.PhraseType = "CNTCT_TYP"
            'Me.cbContactType.DisplayMember = "PhraseText"
            'Me.cbContactType.ValueMember = "PhraseID"

            Me.cbContMethod.PhraseType = "CNTCT_METH"
            'Me.cbContMethod.DisplayMember = "PhraseText"
            'Me.cbContMethod.ValueMember = "PhraseID"

            Me.cbReportOptions.PhraseType = "REP_OPTION"
            'Me.cbReportOptions.DisplayMember = "PhraseText"
            'Me.cbReportOptions.ValueMember = "PhraseID"


            Me.cbContactStatus.PhraseType = "CONT_STAT"
            'Me.cbContactStatus.DisplayMember = "PhraseText"
            'Me.cbContactStatus.ValueMember = "PhraseID"

            Me.cbDivision.DataSource = objDivisionPhrase
            Me.cbDivision.DisplayMember = "PhraseText"
            Me.cbDivision.ValueMember = "PhraseID"


            '</Added code modified 23-Apr-2007 06:33 by taylor> 

            LogTiming("new complete")
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "frmCustomerDetails-New-47")
        Finally
        End Try

        Try  ' Protecting Building Tab pages 
            Dim strTabpages As String = CConnection.PackageStringList("lib_cctool.GetOptionTypes", "")
            Dim StrTabs As String() = strTabpages.Split(New Char() {","})

            Dim i As Integer
            For i = 0 To StrTabs.Length - 2
                Dim strHeader As String = CStr(StrTabs.GetValue(i))
                If strHeader.Trim.Length > 0 Then
                    Dim objTabPage As TabPage = New TabPage()
                    Me.TabOptions.Controls.Add(objTabPage)
                    objTabPage.Text = strHeader.Chars(0) & Mid(strHeader, 2).ToLower

                    Dim objLabel As Label = New Label()
                    objLabel.BackColor = System.Drawing.SystemColors.Highlight
                    objLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    objLabel.ForeColor = System.Drawing.SystemColors.HighlightText
                    objLabel.Location = New System.Drawing.Point(24, 8)
                    objLabel.Name = "lbl" & strHeader
                    objLabel.Size = New System.Drawing.Size(320, 23)
                    objLabel.TabIndex = 8
                    objLabel.Text = "Double-click Row to Add/Edit, Ctrl-delete to remove option"

                    objTabPage.Controls.Add(objLabel)
                    objTabPage.Location = New System.Drawing.Point(4, 4)
                    objTabPage.Name = strHeader
                    objTabPage.Size = New System.Drawing.Size(753, 470)
                    objTabPage.TabIndex = 2
                    'Me.objTabPage.Text = "Certificates"

                    Dim lvOpList As MedscreenCommonGui.OptionsListView = New MedscreenCommonGui.OptionsListView()

                    'Build control
                    lvOpList.ClientOptions = Nothing
                    lvOpList.FullRowSelect = True
                    lvOpList.Location = New System.Drawing.Point(24, 40)
                    lvOpList.Name = "optList" & strHeader
                    lvOpList.OptionClass = strHeader
                    lvOpList.Size = New System.Drawing.Size(672, 320)
                    lvOpList.TabIndex = 0
                    lvOpList.UserID = ""
                    lvOpList.View = System.Windows.Forms.View.Details
                    lvOpList.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top

                    objTabPage.Controls.Add(lvOpList)
                    objOptionTabs.Add(lvOpList)
                    AddHandler lvOpList.OptionRemoved, AddressOf lvAccounts_OptionAdded
                End If
            Next
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "frmCustomerDetails-New-247")
        Finally
        End Try

        If Not MedscreenLib.Glossary.Glossary.UserHasRole("MED_EDIT_PRICES") Then
            Me.PictureBox2.Show()
            Me.lblNoEdit.Show()
        End If
        'Set up price books
        myPricebooks = New MedscreenLib.Glossary.PhraseCollection("PRICEBOOK")
        myPricebooks.Load()
        Me.cbPricing.DataSource = myPricebooks
        Me.cbPricing.DisplayMember = "PhraseText"
        Me.cbPricing.ValueMember = "PhraseID"


        'dynamic pricing panels
        Try
            myPanels.Load()
            'Debug.WriteLine(mypanels.Count)
            Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet
            For Each myPnlColl In myPanels
                myPnlColl.BuildControls(Me)
            Next
            'Set up header controls so they will be at the top.
            Me.pnlMinCharge.Controls.Add(Me.pnlMinchargeheader)
            Me.pnlMinCharge.Controls.Add(Me.pnlSelectCharging)
            Me.panel2.Controls.Add(Me.Panel5)
        Catch ex As Exception
        End Try



    End Sub

    Private objOptionTabs As New Collection()

    'Private objContTypePhrase As MedscreenLib.Glossary.PhraseCollection
    'Private objContMethodPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objRepOptionPhrase As MedscreenLib.Glossary.PhraseCollection
    'Private objContStatusPhrase As MedscreenLib.Glossary.PhraseCollection
    Friend WithEvents lblCountry As System.Windows.Forms.Label
    Friend WithEvents Label97 As System.Windows.Forms.Label
    Friend WithEvents lblWorkflow As System.Windows.Forms.Label
    Private objDivisionPhrase As MedscreenLib.Glossary.PhraseCollection


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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtIdentity As System.Windows.Forms.TextBox
    Friend WithEvents Customer As System.Windows.Forms.TabPage
    Friend WithEvents TabAddresses As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RBAddress As System.Windows.Forms.RadioButton
    Friend WithEvents RBSecondary As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ckNewCustomer As System.Windows.Forms.CheckBox
    Friend WithEvents CKSuspended As System.Windows.Forms.CheckBox
    Friend WithEvents CkRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents CKStopSamples As System.Windows.Forms.CheckBox
    Friend WithEvents ckStopCerts As System.Windows.Forms.CheckBox
    'Friend WithEvents txtParent As System.Windows.Forms.TextBox
    Friend WithEvents cbParent As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CBCustomerType As System.Windows.Forms.ComboBox
    Friend WithEvents cbIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSubContOf As System.Windows.Forms.TextBox
    Friend WithEvents txtPinNumber As System.Windows.Forms.TextBox
    Friend WithEvents dtContractDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TabReporting As System.Windows.Forms.TabPage
    Friend WithEvents TabInvoicing As System.Windows.Forms.TabPage
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ckPORequired As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCustomerOptions As System.Windows.Forms.Button
    Friend WithEvents ckPhoneBeforeFaxing As System.Windows.Forms.CheckBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cbReportMethod As System.Windows.Forms.ComboBox
    Friend WithEvents cbFaxFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cbPlatAccType As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtSageID As System.Windows.Forms.TextBox
    Friend WithEvents ckPositivesOnly As System.Windows.Forms.CheckBox
    Friend WithEvents ckEmailWeeklyReport As System.Windows.Forms.CheckBox
    Friend WithEvents ckReportInBits As System.Windows.Forms.CheckBox
    Friend WithEvents ckAddTests As System.Windows.Forms.CheckBox
    Friend WithEvents ckInvalidPanel As System.Windows.Forms.CheckBox
    Friend WithEvents ckFerment As System.Windows.Forms.CheckBox
    Friend WithEvents ckEmailPosNeg As System.Windows.Forms.CheckBox
    Friend WithEvents ckCreditCard As System.Windows.Forms.CheckBox
    Friend WithEvents lblCurrency As System.Windows.Forms.Label
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtVAT As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents cbStatus As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents cbPricing As System.Windows.Forms.ComboBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents panelStd As System.Windows.Forms.Panel
    Friend WithEvents cbProgMan As System.Windows.Forms.ComboBox
    Friend WithEvents cmdServices As System.Windows.Forms.Button
    Friend WithEvents lbServices As System.Windows.Forms.ListBox
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TabOptions As System.Windows.Forms.TabControl
    Friend WithEvents tabPageAccounts As System.Windows.Forms.TabPage
    Friend WithEvents tabPagePricing As System.Windows.Forms.TabPage
    Friend WithEvents TabPageCertificates As System.Windows.Forms.TabPage
    Friend WithEvents tabPageMisc As System.Windows.Forms.TabPage
    'Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label16 As System.Windows.Forms.Label
    'Friend WithEvents LvPricing As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvAccounts As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvCertificates As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvCustomerType As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvContractor As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvPrisons As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvSamples As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvReporting As MedscreenCommonGui.OptionsListView
    'Friend WithEvents lvRepOptions As MedscreenCommonGui.OptionsListView
    'Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader17 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader18 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents tabPageSampleHandling As System.Windows.Forms.TabPage
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    'Friend WithEvents ColumnHeader19 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader20 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader21 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader22 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents cbInvFreq As System.Windows.Forms.ComboBox
    Friend WithEvents TabContracts As System.Windows.Forms.TabPage
    Friend WithEvents ColContNum As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColItemCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColAnnPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColNextAmount As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColStartDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColConEndDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColPayInterval As System.Windows.Forms.ColumnHeader
    Friend WithEvents colDueNext As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents DTContractStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtAnnualPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents DTNextDue As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents cbPayInterval As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents cbLineItems As System.Windows.Forms.ComboBox
    Friend WithEvents DTContractEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents RBInvoicing As System.Windows.Forms.RadioButton
    Friend WithEvents rbInvShipping As System.Windows.Forms.RadioButton
    Friend WithEvents RBInvMail As System.Windows.Forms.RadioButton
    Friend WithEvents txtSMID As System.Windows.Forms.TextBox
    Friend WithEvents txtProfile As System.Windows.Forms.TextBox
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtEmailPassword As System.Windows.Forms.TextBox
    Friend WithEvents rbGoldmine As System.Windows.Forms.RadioButton
    Friend WithEvents TabPageReporting As System.Windows.Forms.TabPage
    Friend WithEvents cmdDeleteContract As System.Windows.Forms.Button
    Friend WithEvents ckContractRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents TabTestPanels As System.Windows.Forms.TabPage
    Friend WithEvents TabLoginTemplate As System.Windows.Forms.TabPage
    Friend WithEvents lvTestPanels As System.Windows.Forms.ListView
    Friend WithEvents chTestpanelID As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHDefaultPanel As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTPDescription As System.Windows.Forms.ColumnHeader
    Friend WithEvents chMatrix As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAlcoholCutoff As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCannabisCutoff As System.Windows.Forms.ColumnHeader
    Friend WithEvents chBreath As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents chAnalysis As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvTestSchedules As System.Windows.Forms.ListView
    Friend WithEvents chTSDescription As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTSConf As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTSStdTest As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CmdCreateTestPanels As System.Windows.Forms.Button
    Friend WithEvents cbMROID As System.Windows.Forms.ComboBox
    Friend WithEvents cmdPassword As System.Windows.Forms.Button
    Friend WithEvents lblParent As System.Windows.Forms.Label
    Friend WithEvents lblCompanyName As System.Windows.Forms.Label
    Friend WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents lblIdentity As System.Windows.Forms.Label
    Friend WithEvents cbSalesPerson As System.Windows.Forms.ComboBox
    Friend WithEvents lblSalesPerson As System.Windows.Forms.Label
    Friend WithEvents lblMRO As System.Windows.Forms.Label
    Friend WithEvents lblPinNumber As System.Windows.Forms.Label
    Friend WithEvents lblProfile As System.Windows.Forms.Label
    Friend WithEvents lblContractNumber As System.Windows.Forms.Label
    Friend WithEvents txtContractNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtContractNumber1 As System.Windows.Forms.TextBox
    Friend WithEvents lblContractor As System.Windows.Forms.Label
    Friend WithEvents lblContractDate As System.Windows.Forms.Label
    Friend WithEvents lblVATRegno As System.Windows.Forms.Label
    Friend WithEvents txtVATRegno As System.Windows.Forms.TextBox
    Friend WithEvents lblSageID As System.Windows.Forms.Label
    Friend WithEvents lvRoutinePayments As System.Windows.Forms.ListView
    Friend WithEvents colLastInvoice As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblFirstPayment As System.Windows.Forms.Label
    Friend WithEvents dtFirstPayment As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblThen As System.Windows.Forms.Label
    Friend WithEvents txtNextPayment As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents tabComments As System.Windows.Forms.TabPage
    Friend WithEvents tabHistory As System.Windows.Forms.TabPage
    Friend WithEvents pnlComments As System.Windows.Forms.Panel
    Friend WithEvents lvComments As System.Windows.Forms.ListView
    Friend WithEvents chCommDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommType As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHComment As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommBy As System.Windows.Forms.ColumnHeader
    Friend WithEvents ChCommAction As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddComment As System.Windows.Forms.Button
    Friend WithEvents pnlHistory As System.Windows.Forms.Panel
    Friend WithEvents TabDiscSchemes As System.Windows.Forms.TabPage
    'Friend WithEvents pnlNormalDiscSheme As System.Windows.Forms.Panel
    'Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents SplitterAddr As System.Windows.Forms.Splitter
    Friend WithEvents txtSameAddress As System.Windows.Forms.RichTextBox
    'Friend WithEvents pnlVolDiscountScheme As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chNVDID As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNVDType As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNVDApplyto As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNVDDiscount As System.Windows.Forms.ColumnHeader
    'Friend WithEvents lvRoutineDiscount As System.Windows.Forms.ListView
    'Friend WithEvents lvVolumeDiscounts As System.Windows.Forms.ListView
    Friend WithEvents chVolScheme As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolType As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolApply As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolMin As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolMaxSample As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolMinCharge As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolDiscount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chVolExample As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNVDPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents chNVDDisPrice As System.Windows.Forms.ColumnHeader

    Friend WithEvents pnlNVDScheme As System.Windows.Forms.Panel
    'Friend WithEvents pnlVolumeDiscounts As System.Windows.Forms.Panel
    Friend WithEvents cmdFindScheme As System.Windows.Forms.Button
    Friend WithEvents tabPricing As System.Windows.Forms.TabPage
    Friend WithEvents tabTypes As System.Windows.Forms.TabControl
    Friend WithEvents tabpSample As System.Windows.Forms.TabPage
    Friend WithEvents panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel

    Friend WithEvents lblPCurr As System.Windows.Forms.Label
    Friend WithEvents lblLPCurr As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pnlAO As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents txtAOPrice As System.Windows.Forms.TextBox
    'Friend WithEvents txtAOListPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlCallOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlFixedSite As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents pnlExternal As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents pnlMRO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents lblMROComment As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents tabpMinCharge As System.Windows.Forms.TabPage
    Friend WithEvents pnlMinCharge As System.Windows.Forms.Panel
    Friend WithEvents pnlOutHoursMC As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents txtouthrmcDiscount As System.Windows.Forms.TextBox
    'Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents pnlInHourMinCharge As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents lblpCollCurr As System.Windows.Forms.Label
    Friend WithEvents lbllpCollCurr As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents tabpAddTime As System.Windows.Forms.TabPage
    Friend WithEvents pnlAddTime As System.Windows.Forms.Panel
    Friend WithEvents pnlOutHoursAddTimeCO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlINHoursAddTimeCO As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents lblihadtrpricephr As System.Windows.Forms.Label
    Friend WithEvents pnlOutHoursAddTimeRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlINHoursAddTimeRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlADHeader As System.Windows.Forms.Panel
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents lbladpCurr As System.Windows.Forms.Label
    Friend WithEvents lbllPADCurr As System.Windows.Forms.Label
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents Label67 As System.Windows.Forms.Label
    Friend WithEvents Label68 As System.Windows.Forms.Label
    Friend WithEvents pnlTimeAllowance As System.Windows.Forms.Panel
    Friend WithEvents lblTimeAllowance As System.Windows.Forms.Label
    Friend WithEvents txtTimeAllowance As System.Windows.Forms.TextBox
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents ckTimeAllowance As System.Windows.Forms.CheckBox
    Friend WithEvents tabpTests As System.Windows.Forms.TabPage
    Friend WithEvents pnlTests As System.Windows.Forms.Panel
    Friend WithEvents ckChargeAdditionalTests As System.Windows.Forms.CheckBox
    Friend WithEvents cbEditor As System.Windows.Forms.ComboBox
    Friend WithEvents lvTests As ListViewEx.ListViewEx
    Friend WithEvents chTest As System.Windows.Forms.ColumnHeader
    Friend WithEvents chBar1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chListPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents chbar2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDiscount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chBar3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label70 As System.Windows.Forms.Label
    Friend WithEvents pnlTimeAllowHead As System.Windows.Forms.Panel
    Friend WithEvents ctxAddService As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuctxAddService As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCTXDiscountDates As System.Windows.Forms.MenuItem
    Friend WithEvents Label73 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tabpKits As System.Windows.Forms.TabPage
    Friend WithEvents lvExKits As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtKitEdit As System.Windows.Forms.TextBox
    Friend WithEvents chListPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents tabpCancellations As System.Windows.Forms.TabPage
    Friend WithEvents pnlCancellations As System.Windows.Forms.Panel
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents dpOutOfhoursRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpInHoursCancel As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpFixedSiteCancellation As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpCallOutCancel As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents chCount As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvExEQP As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chKitSold As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAvgPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAvgPrice2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents dpRoutineCollCharge As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpCallOutCollCharge As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpRoutineCollChargeOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpCallOutCollChargeOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents Label74 As System.Windows.Forms.Label
    Friend WithEvents lvTestConfirm As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader17 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader18 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader19 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader20 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader21 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader22 As System.Windows.Forms.ColumnHeader
    Friend WithEvents dpMCCallOutIn As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpMCCallOutOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents cmdPasswordLetter As System.Windows.Forms.Button
    Friend WithEvents HtmlEditor2 As System.Windows.Forms.WebBrowser
    Friend WithEvents pnlSelectCharging As System.Windows.Forms.Panel
    Friend WithEvents pnlMinchargeheader As System.Windows.Forms.Panel
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents udRoutine As System.Windows.Forms.DomainUpDown
    Friend WithEvents udCallOut As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents lblReportingMethod As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents AddressControl1 As MedscreenCommonGui.AddressPanel
    Friend WithEvents lblFaxError As System.Windows.Forms.Label
    Friend WithEvents lblEmailError As System.Windows.Forms.Label
    Friend WithEvents txtEmailPositives As System.Windows.Forms.TextBox
    Friend WithEvents lblPosEmail As System.Windows.Forms.Label
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents HtmlEditor3 As System.Windows.Forms.WebBrowser
    Friend WithEvents ckDivertPos As System.Windows.Forms.CheckBox
    Friend WithEvents chkEmailAndPrint As System.Windows.Forms.CheckBox
    Friend WithEvents cmdAddTest As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteTest As System.Windows.Forms.Button
    Friend WithEvents txtAccountType As System.Windows.Forms.TextBox
    Friend WithEvents Label78 As System.Windows.Forms.Label
    Friend WithEvents Label79 As System.Windows.Forms.Label
    'Friend WithEvents cmdAddScheme As System.Windows.Forms.Button
    'Friend WithEvents cbVolScheme As System.Windows.Forms.ComboBox
    'Friend WithEvents Label80 As System.Windows.Forms.Label
    'Friend WithEvents cmdDelDiscount As System.Windows.Forms.Button
    Friend WithEvents pnlDiscounts As MedscreenCommonGui.Panels.VolDiscountPanel
    Friend WithEvents tabContacts As System.Windows.Forms.TabPage
    Friend WithEvents pnlContactList As System.Windows.Forms.Panel
    Friend WithEvents Splitter3 As System.Windows.Forms.Splitter
    Friend WithEvents pnlContactDetails As System.Windows.Forms.Panel
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ckContactRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents ColumnHeader55 As System.Windows.Forms.ColumnHeader

    Friend WithEvents cmdAddContact As System.Windows.Forms.Button
    Friend WithEvents cmdDelContact As System.Windows.Forms.Button
    Friend WithEvents lvContacts As System.Windows.Forms.ListView
    Friend WithEvents AdrContact As MedscreenCommonGui.AddressPanel
    Friend WithEvents Label81 As System.Windows.Forms.Label
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents cmdFindContactAdr As System.Windows.Forms.Button
    Friend WithEvents cmdNewContactAdr As System.Windows.Forms.Button
    Friend WithEvents Splitter4 As System.Windows.Forms.Splitter
    Friend WithEvents pnlContact As System.Windows.Forms.Panel
    Friend WithEvents Label83 As System.Windows.Forms.Label
    Friend WithEvents cbContactType As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents cbContMethod As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label84 As System.Windows.Forms.Label
    Friend WithEvents cbReportOptions As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label85 As System.Windows.Forms.Label
    Friend WithEvents cbReportFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label86 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdateContact As System.Windows.Forms.Button
    Friend WithEvents cbContactStatus As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label87 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label88 As System.Windows.Forms.Label
    Friend WithEvents cbDivision As System.Windows.Forms.ComboBox
    Friend WithEvents lblNoEdit As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents lblSerNotSold As System.Windows.Forms.Label
    Friend WithEvents lblSerNotSoldYet As System.Windows.Forms.Label
    Friend WithEvents lblSerWasSold As System.Windows.Forms.Label
    Friend WithEvents lblSerSold As System.Windows.Forms.Label
    Friend WithEvents LblItemComment As System.Windows.Forms.Label
    Friend WithEvents txtItemComment As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents cmdChangeDate As System.Windows.Forms.Button
    Friend WithEvents DTPricing As System.Windows.Forms.DateTimePicker

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerDetails))
        Dim DiscountPanelBase1 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase2 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase3 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase4 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase5 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase6 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase7 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase8 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase9 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase10 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase11 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase12 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase13 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase14 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase15 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase16 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase17 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase18 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase19 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase20 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase21 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Dim DiscountPanelBase22 As MedscreenCommonGui.Panels.DiscountPanelBase = New MedscreenCommonGui.Panels.DiscountPanelBase
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.Customer = New System.Windows.Forms.TabPage
        Me.lblCountry = New System.Windows.Forms.Label
        Me.cbDivision = New System.Windows.Forms.ComboBox
        Me.Label88 = New System.Windows.Forms.Label
        Me.Label79 = New System.Windows.Forms.Label
        Me.Label78 = New System.Windows.Forms.Label
        Me.txtAccountType = New System.Windows.Forms.TextBox
        Me.cmdPasswordLetter = New System.Windows.Forms.Button
        Me.cmdPassword = New System.Windows.Forms.Button
        Me.cbMROID = New System.Windows.Forms.ComboBox
        Me.lblMRO = New System.Windows.Forms.Label
        Me.txtEmailPassword = New System.Windows.Forms.TextBox
        Me.Label42 = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtProfile = New System.Windows.Forms.TextBox
        Me.lblProfile = New System.Windows.Forms.Label
        Me.txtSMID = New System.Windows.Forms.TextBox
        Me.cbSalesPerson = New System.Windows.Forms.ComboBox
        Me.cbProgMan = New System.Windows.Forms.ComboBox
        Me.cbStatus = New MedscreenCommonGui.PhraseComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.lblContractDate = New System.Windows.Forms.Label
        Me.dtContractDate = New System.Windows.Forms.DateTimePicker
        Me.txtPinNumber = New System.Windows.Forms.TextBox
        Me.lblPinNumber = New System.Windows.Forms.Label
        Me.txtSubContOf = New System.Windows.Forms.TextBox
        Me.cbIndustry = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.CBCustomerType = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.lblContractNumber = New System.Windows.Forms.Label
        Me.txtContractNumber = New System.Windows.Forms.TextBox
        Me.lblSalesPerson = New System.Windows.Forms.Label
        Me.cbParent = New System.Windows.Forms.ComboBox
        Me.lblParent = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ckStopCerts = New System.Windows.Forms.CheckBox
        Me.CKStopSamples = New System.Windows.Forms.CheckBox
        Me.CkRemoved = New System.Windows.Forms.CheckBox
        Me.CKSuspended = New System.Windows.Forms.CheckBox
        Me.ckNewCustomer = New System.Windows.Forms.CheckBox
        Me.lblVATRegno = New System.Windows.Forms.Label
        Me.txtVATRegno = New System.Windows.Forms.TextBox
        Me.lblCompanyName = New System.Windows.Forms.Label
        Me.txtCompanyName = New System.Windows.Forms.TextBox
        Me.lblIdentity = New System.Windows.Forms.Label
        Me.txtIdentity = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.tabContacts = New System.Windows.Forms.TabPage
        Me.pnlContactDetails = New System.Windows.Forms.Panel
        Me.pnlContact = New System.Windows.Forms.Panel
        Me.ckContactRemoved = New System.Windows.Forms.CheckBox
        Me.cbContactStatus = New MedscreenCommonGui.PhraseComboBox
        Me.Label87 = New System.Windows.Forms.Label
        Me.cmdUpdateContact = New System.Windows.Forms.Button
        Me.cbReportOptions = New MedscreenCommonGui.PhraseComboBox
        Me.Label85 = New System.Windows.Forms.Label
        Me.cbReportFormat = New System.Windows.Forms.ComboBox
        Me.Label86 = New System.Windows.Forms.Label
        Me.cbContMethod = New MedscreenCommonGui.PhraseComboBox
        Me.Label84 = New System.Windows.Forms.Label
        Me.cbContactType = New MedscreenCommonGui.PhraseComboBox
        Me.Label83 = New System.Windows.Forms.Label
        Me.Splitter4 = New System.Windows.Forms.Splitter
        Me.AdrContact = New MedscreenCommonGui.AddressPanel
        Me.Label81 = New System.Windows.Forms.Label
        Me.Label82 = New System.Windows.Forms.Label
        Me.cmdFindContactAdr = New System.Windows.Forms.Button
        Me.cmdNewContactAdr = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lvContacts = New System.Windows.Forms.ListView
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader55 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader
        Me.tabHistory = New System.Windows.Forms.TabPage
        Me.pnlHistory = New System.Windows.Forms.Panel
        Me.HtmlEditor2 = New System.Windows.Forms.WebBrowser
        Me.cmdCustomerOptions = New System.Windows.Forms.Button
        Me.Splitter3 = New System.Windows.Forms.Splitter
        Me.pnlContactList = New System.Windows.Forms.Panel
        Me.cmdDelContact = New System.Windows.Forms.Button
        Me.cmdAddContact = New System.Windows.Forms.Button
        Me.TabAddresses = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblReportingMethod = New System.Windows.Forms.Label
        Me.rbGoldmine = New System.Windows.Forms.RadioButton
        Me.rbInvShipping = New System.Windows.Forms.RadioButton
        Me.RBInvMail = New System.Windows.Forms.RadioButton
        Me.RBInvoicing = New System.Windows.Forms.RadioButton
        Me.RBSecondary = New System.Windows.Forms.RadioButton
        Me.RBAddress = New System.Windows.Forms.RadioButton
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.txtSameAddress = New System.Windows.Forms.RichTextBox
        Me.SplitterAddr = New System.Windows.Forms.Splitter
        Me.AddressControl1 = New MedscreenCommonGui.AddressPanel
        Me.Label20 = New System.Windows.Forms.Label
        Me.lblFaxError = New System.Windows.Forms.Label
        Me.lblEmailError = New System.Windows.Forms.Label
        Me.txtEmailPositives = New System.Windows.Forms.TextBox
        Me.lblPosEmail = New System.Windows.Forms.Label
        Me.tabPricing = New System.Windows.Forms.TabPage
        Me.cmdChangeDate = New System.Windows.Forms.Button
        Me.DTPricing = New System.Windows.Forms.DateTimePicker
        Me.Label61 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.lblSerSold = New System.Windows.Forms.Label
        Me.lblSerNotSoldYet = New System.Windows.Forms.Label
        Me.lblSerWasSold = New System.Windows.Forms.Label
        Me.lblSerNotSold = New System.Windows.Forms.Label
        Me.tabTypes = New System.Windows.Forms.TabControl
        Me.tabpSample = New System.Windows.Forms.TabPage
        Me.Label71 = New System.Windows.Forms.Label
        Me.panel2 = New System.Windows.Forms.Panel
        Me.pnlMRO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.lblMROComment = New System.Windows.Forms.Label
        Me.pnlExternal = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.Label6 = New System.Windows.Forms.Label
        Me.pnlFixedSite = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.Label73 = New System.Windows.Forms.Label
        Me.pnlCallOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlAO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.tabpMinCharge = New System.Windows.Forms.TabPage
        Me.pnlMinCharge = New System.Windows.Forms.Panel
        Me.dpCallOutCollChargeOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.ctxAddService = New System.Windows.Forms.ContextMenu
        Me.mnuctxAddService = New System.Windows.Forms.MenuItem
        Me.mnuCTXDiscountDates = New System.Windows.Forms.MenuItem
        Me.dpCallOutCollCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpRoutineCollChargeOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpRoutineCollCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpMCCallOutOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpMCCallOutIn = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlOutHoursMC = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlInHourMinCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.tabpAddTime = New System.Windows.Forms.TabPage
        Me.pnlAddTime = New System.Windows.Forms.Panel
        Me.pnlOutHoursAddTimeCO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlINHoursAddTimeCO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlOutHoursAddTimeRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlINHoursAddTimeRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlADHeader = New System.Windows.Forms.Panel
        Me.Label63 = New System.Windows.Forms.Label
        Me.lbladpCurr = New System.Windows.Forms.Label
        Me.lbllPADCurr = New System.Windows.Forms.Label
        Me.Label64 = New System.Windows.Forms.Label
        Me.Label65 = New System.Windows.Forms.Label
        Me.Label66 = New System.Windows.Forms.Label
        Me.Label67 = New System.Windows.Forms.Label
        Me.Label68 = New System.Windows.Forms.Label
        Me.pnlTimeAllowance = New System.Windows.Forms.Panel
        Me.pnlTimeAllowHead = New System.Windows.Forms.Panel
        Me.Label70 = New System.Windows.Forms.Label
        Me.lblTimeAllowance = New System.Windows.Forms.Label
        Me.txtTimeAllowance = New System.Windows.Forms.TextBox
        Me.Label69 = New System.Windows.Forms.Label
        Me.ckTimeAllowance = New System.Windows.Forms.CheckBox
        Me.tabpTests = New System.Windows.Forms.TabPage
        Me.pnlTests = New System.Windows.Forms.Panel
        Me.Label74 = New System.Windows.Forms.Label
        Me.lvTestConfirm = New ListViewEx.ListViewEx
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader17 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader18 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader19 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader20 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader21 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader22 = New System.Windows.Forms.ColumnHeader
        Me.Label62 = New System.Windows.Forms.Label
        Me.ckChargeAdditionalTests = New System.Windows.Forms.CheckBox
        Me.cbEditor = New System.Windows.Forms.ComboBox
        Me.lvTests = New ListViewEx.ListViewEx
        Me.chTest = New System.Windows.Forms.ColumnHeader
        Me.chBar1 = New System.Windows.Forms.ColumnHeader
        Me.chListPrice = New System.Windows.Forms.ColumnHeader
        Me.chbar2 = New System.Windows.Forms.ColumnHeader
        Me.chDiscount = New System.Windows.Forms.ColumnHeader
        Me.chBar3 = New System.Windows.Forms.ColumnHeader
        Me.chPrice = New System.Windows.Forms.ColumnHeader
        Me.tabpKits = New System.Windows.Forms.TabPage
        Me.lvExEQP = New ListViewEx.ListViewEx
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.chAvgPrice2 = New System.Windows.Forms.ColumnHeader
        Me.txtKitEdit = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.lvExKits = New ListViewEx.ListViewEx
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.chListPriceKit = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.chPriceKit = New System.Windows.Forms.ColumnHeader
        Me.chCount = New System.Windows.Forms.ColumnHeader
        Me.chKitSold = New System.Windows.Forms.ColumnHeader
        Me.chAvgPrice = New System.Windows.Forms.ColumnHeader
        Me.tabpCancellations = New System.Windows.Forms.TabPage
        Me.pnlCancellations = New System.Windows.Forms.Panel
        Me.dpCallOutCancel = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpFixedSiteCancellation = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpInHoursCancel = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpOutOfhoursRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.Label52 = New System.Windows.Forms.Label
        Me.Label53 = New System.Windows.Forms.Label
        Me.Label59 = New System.Windows.Forms.Label
        Me.Label60 = New System.Windows.Forms.Label
        Me.TabDiscSchemes = New System.Windows.Forms.TabPage
        Me.pnlDiscounts = New MedscreenCommonGui.Panels.VolDiscountPanel
        Me.TabInvoicing = New System.Windows.Forms.TabPage
        Me.cbInvFreq = New System.Windows.Forms.ComboBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.lbServices = New System.Windows.Forms.ListBox
        Me.cbPricing = New System.Windows.Forms.ComboBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtVAT = New System.Windows.Forms.TextBox
        Me.cbCurrency = New System.Windows.Forms.ComboBox
        Me.lblCurrency = New System.Windows.Forms.Label
        Me.ckCreditCard = New System.Windows.Forms.CheckBox
        Me.lblSageID = New System.Windows.Forms.Label
        Me.txtSageID = New System.Windows.Forms.TextBox
        Me.cbPlatAccType = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.ckPORequired = New System.Windows.Forms.CheckBox
        Me.panelStd = New System.Windows.Forms.Panel
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.TabReporting = New System.Windows.Forms.TabPage
        Me.chkEmailAndPrint = New System.Windows.Forms.CheckBox
        Me.TabOptions = New System.Windows.Forms.TabControl
        Me.ckEmailPosNeg = New System.Windows.Forms.CheckBox
        Me.ckPositivesOnly = New System.Windows.Forms.CheckBox
        Me.cbFaxFormat = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.cbReportMethod = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.ckPhoneBeforeFaxing = New System.Windows.Forms.CheckBox
        Me.ckEmailWeeklyReport = New System.Windows.Forms.CheckBox
        Me.ckReportInBits = New System.Windows.Forms.CheckBox
        Me.ckDivertPos = New System.Windows.Forms.CheckBox
        Me.TabContracts = New System.Windows.Forms.TabPage
        Me.HtmlEditor3 = New System.Windows.Forms.WebBrowser
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.lblThen = New System.Windows.Forms.Label
        Me.dtFirstPayment = New System.Windows.Forms.DateTimePicker
        Me.lblFirstPayment = New System.Windows.Forms.Label
        Me.ckContractRemoved = New System.Windows.Forms.CheckBox
        Me.cmdDeleteContract = New System.Windows.Forms.Button
        Me.DTContractEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label41 = New System.Windows.Forms.Label
        Me.cbLineItems = New System.Windows.Forms.ComboBox
        Me.Label40 = New System.Windows.Forms.Label
        Me.cbPayInterval = New System.Windows.Forms.ComboBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.DTNextDue = New System.Windows.Forms.DateTimePicker
        Me.Label38 = New System.Windows.Forms.Label
        Me.txtNextPayment = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.txtAnnualPrice = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.DTContractStart = New System.Windows.Forms.DateTimePicker
        Me.Label35 = New System.Windows.Forms.Label
        Me.txtContractNumber1 = New System.Windows.Forms.TextBox
        Me.Label34 = New System.Windows.Forms.Label
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.lvRoutinePayments = New System.Windows.Forms.ListView
        Me.ColContNum = New System.Windows.Forms.ColumnHeader
        Me.ColItemCode = New System.Windows.Forms.ColumnHeader
        Me.ColStartDate = New System.Windows.Forms.ColumnHeader
        Me.ColConEndDate = New System.Windows.Forms.ColumnHeader
        Me.ColAnnPrice = New System.Windows.Forms.ColumnHeader
        Me.ColNextAmount = New System.Windows.Forms.ColumnHeader
        Me.ColPayInterval = New System.Windows.Forms.ColumnHeader
        Me.colDueNext = New System.Windows.Forms.ColumnHeader
        Me.colLastInvoice = New System.Windows.Forms.ColumnHeader
        Me.LblItemComment = New System.Windows.Forms.Label
        Me.txtItemComment = New System.Windows.Forms.TextBox
        Me.tabComments = New System.Windows.Forms.TabPage
        Me.cmdAddComment = New System.Windows.Forms.Button
        Me.pnlComments = New System.Windows.Forms.Panel
        Me.lvComments = New System.Windows.Forms.ListView
        Me.chCommDate = New System.Windows.Forms.ColumnHeader
        Me.chCommType = New System.Windows.Forms.ColumnHeader
        Me.CHComment = New System.Windows.Forms.ColumnHeader
        Me.chCommBy = New System.Windows.Forms.ColumnHeader
        Me.ChCommAction = New System.Windows.Forms.ColumnHeader
        Me.TabTestPanels = New System.Windows.Forms.TabPage
        Me.cmdDeleteTest = New System.Windows.Forms.Button
        Me.cmdAddTest = New System.Windows.Forms.Button
        Me.CmdCreateTestPanels = New System.Windows.Forms.Button
        Me.lvTestSchedules = New System.Windows.Forms.ListView
        Me.chAnalysis = New System.Windows.Forms.ColumnHeader
        Me.chTSDescription = New System.Windows.Forms.ColumnHeader
        Me.chTSConf = New System.Windows.Forms.ColumnHeader
        Me.chTSStdTest = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.Label43 = New System.Windows.Forms.Label
        Me.lvTestPanels = New System.Windows.Forms.ListView
        Me.chTestpanelID = New System.Windows.Forms.ColumnHeader
        Me.chTPDescription = New System.Windows.Forms.ColumnHeader
        Me.CHDefaultPanel = New System.Windows.Forms.ColumnHeader
        Me.chMatrix = New System.Windows.Forms.ColumnHeader
        Me.chAlcoholCutoff = New System.Windows.Forms.ColumnHeader
        Me.chCannabisCutoff = New System.Windows.Forms.ColumnHeader
        Me.chBreath = New System.Windows.Forms.ColumnHeader
        Me.ckAddTests = New System.Windows.Forms.CheckBox
        Me.ckInvalidPanel = New System.Windows.Forms.CheckBox
        Me.TabLoginTemplate = New System.Windows.Forms.TabPage
        Me.chVolScheme = New System.Windows.Forms.ColumnHeader
        Me.chVolType = New System.Windows.Forms.ColumnHeader
        Me.chVolApply = New System.Windows.Forms.ColumnHeader
        Me.chVolMin = New System.Windows.Forms.ColumnHeader
        Me.chVolMaxSample = New System.Windows.Forms.ColumnHeader
        Me.chVolMinCharge = New System.Windows.Forms.ColumnHeader
        Me.chVolDiscount = New System.Windows.Forms.ColumnHeader
        Me.chVolExample = New System.Windows.Forms.ColumnHeader
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlNVDScheme = New System.Windows.Forms.Panel
        Me.cmdFindScheme = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.chNVDID = New System.Windows.Forms.ColumnHeader
        Me.chNVDType = New System.Windows.Forms.ColumnHeader
        Me.chNVDApplyto = New System.Windows.Forms.ColumnHeader
        Me.chNVDDiscount = New System.Windows.Forms.ColumnHeader
        Me.chNVDPrice = New System.Windows.Forms.ColumnHeader
        Me.chNVDDisPrice = New System.Windows.Forms.ColumnHeader
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblPCurr = New System.Windows.Forms.Label
        Me.lblLPCurr = New System.Windows.Forms.Label
        Me.Label46 = New System.Windows.Forms.Label
        Me.Label47 = New System.Windows.Forms.Label
        Me.Label48 = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.pnlMinchargeheader = New System.Windows.Forms.Panel
        Me.lblpCollCurr = New System.Windows.Forms.Label
        Me.lbllpCollCurr = New System.Windows.Forms.Label
        Me.Label54 = New System.Windows.Forms.Label
        Me.Label55 = New System.Windows.Forms.Label
        Me.Label56 = New System.Windows.Forms.Label
        Me.Label57 = New System.Windows.Forms.Label
        Me.Label58 = New System.Windows.Forms.Label
        Me.pnlSelectCharging = New System.Windows.Forms.Panel
        Me.Label77 = New System.Windows.Forms.Label
        Me.udCallOut = New System.Windows.Forms.DomainUpDown
        Me.Label76 = New System.Windows.Forms.Label
        Me.udRoutine = New System.Windows.Forms.DomainUpDown
        Me.Label75 = New System.Windows.Forms.Label
        Me.TabPageCertificates = New System.Windows.Forms.TabPage
        Me.Label30 = New System.Windows.Forms.Label
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.Label32 = New System.Windows.Forms.Label
        Me.tabPageAccounts = New System.Windows.Forms.TabPage
        Me.Label28 = New System.Windows.Forms.Label
        Me.TabPageReporting = New System.Windows.Forms.TabPage
        Me.ckFerment = New System.Windows.Forms.CheckBox
        Me.tabPageSampleHandling = New System.Windows.Forms.TabPage
        Me.Label31 = New System.Windows.Forms.Label
        Me.tabPagePricing = New System.Windows.Forms.TabPage
        Me.Label29 = New System.Windows.Forms.Label
        Me.tabPageMisc = New System.Windows.Forms.TabPage
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.lblContractor = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label44 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblNoEdit = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.cmdServices = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.Customer.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.tabContacts.SuspendLayout()
        Me.pnlContactDetails.SuspendLayout()
        Me.pnlContact.SuspendLayout()
        Me.AdrContact.SuspendLayout()
        Me.tabHistory.SuspendLayout()
        Me.pnlHistory.SuspendLayout()
        Me.pnlContactList.SuspendLayout()
        Me.TabAddresses.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.AddressControl1.SuspendLayout()
        Me.tabPricing.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabTypes.SuspendLayout()
        Me.tabpSample.SuspendLayout()
        Me.panel2.SuspendLayout()
        Me.pnlMRO.SuspendLayout()
        Me.pnlExternal.SuspendLayout()
        Me.pnlFixedSite.SuspendLayout()
        Me.tabpMinCharge.SuspendLayout()
        Me.pnlMinCharge.SuspendLayout()
        Me.tabpAddTime.SuspendLayout()
        Me.pnlAddTime.SuspendLayout()
        Me.pnlADHeader.SuspendLayout()
        Me.pnlTimeAllowance.SuspendLayout()
        Me.pnlTimeAllowHead.SuspendLayout()
        Me.tabpTests.SuspendLayout()
        Me.pnlTests.SuspendLayout()
        Me.tabpKits.SuspendLayout()
        Me.tabpCancellations.SuspendLayout()
        Me.pnlCancellations.SuspendLayout()
        Me.TabDiscSchemes.SuspendLayout()
        Me.TabInvoicing.SuspendLayout()
        Me.panelStd.SuspendLayout()
        Me.TabReporting.SuspendLayout()
        Me.TabContracts.SuspendLayout()
        Me.tabComments.SuspendLayout()
        Me.pnlComments.SuspendLayout()
        Me.TabTestPanels.SuspendLayout()
        Me.pnlNVDScheme.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.pnlMinchargeheader.SuspendLayout()
        Me.pnlSelectCharging.SuspendLayout()
        Me.TabPageCertificates.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        Me.tabPageAccounts.SuspendLayout()
        Me.tabPageSampleHandling.SuspendLayout()
        Me.tabPagePricing.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.Customer)
        Me.TabControl1.Controls.Add(Me.tabContacts)
        Me.TabControl1.Controls.Add(Me.TabAddresses)
        Me.TabControl1.Controls.Add(Me.tabPricing)
        Me.TabControl1.Controls.Add(Me.TabDiscSchemes)
        Me.TabControl1.Controls.Add(Me.TabInvoicing)
        Me.TabControl1.Controls.Add(Me.TabReporting)
        Me.TabControl1.Controls.Add(Me.TabContracts)
        Me.TabControl1.Controls.Add(Me.tabComments)
        Me.TabControl1.Controls.Add(Me.TabTestPanels)
        Me.TabControl1.Controls.Add(Me.TabLoginTemplate)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(928, 653)
        Me.TabControl1.TabIndex = 0
        '
        'Customer
        '
        Me.Customer.Controls.Add(Me.lblCountry)
        Me.Customer.Controls.Add(Me.cbDivision)
        Me.Customer.Controls.Add(Me.Label88)
        Me.Customer.Controls.Add(Me.Label79)
        Me.Customer.Controls.Add(Me.Label78)
        Me.Customer.Controls.Add(Me.txtAccountType)
        Me.Customer.Controls.Add(Me.cmdPasswordLetter)
        Me.Customer.Controls.Add(Me.cmdPassword)
        Me.Customer.Controls.Add(Me.cbMROID)
        Me.Customer.Controls.Add(Me.lblMRO)
        Me.Customer.Controls.Add(Me.txtEmailPassword)
        Me.Customer.Controls.Add(Me.Label42)
        Me.Customer.Controls.Add(Me.txtComments)
        Me.Customer.Controls.Add(Me.Label19)
        Me.Customer.Controls.Add(Me.txtProfile)
        Me.Customer.Controls.Add(Me.lblProfile)
        Me.Customer.Controls.Add(Me.txtSMID)
        Me.Customer.Controls.Add(Me.cbSalesPerson)
        Me.Customer.Controls.Add(Me.cbProgMan)
        Me.Customer.Controls.Add(Me.cbStatus)
        Me.Customer.Controls.Add(Me.Label13)
        Me.Customer.Controls.Add(Me.lblContractDate)
        Me.Customer.Controls.Add(Me.dtContractDate)
        Me.Customer.Controls.Add(Me.txtPinNumber)
        Me.Customer.Controls.Add(Me.lblPinNumber)
        Me.Customer.Controls.Add(Me.txtSubContOf)
        Me.Customer.Controls.Add(Me.cbIndustry)
        Me.Customer.Controls.Add(Me.Label5)
        Me.Customer.Controls.Add(Me.CBCustomerType)
        Me.Customer.Controls.Add(Me.Label8)
        Me.Customer.Controls.Add(Me.lblContractNumber)
        Me.Customer.Controls.Add(Me.txtContractNumber)
        Me.Customer.Controls.Add(Me.lblSalesPerson)
        Me.Customer.Controls.Add(Me.cbParent)
        Me.Customer.Controls.Add(Me.lblParent)
        Me.Customer.Controls.Add(Me.GroupBox2)
        Me.Customer.Controls.Add(Me.lblVATRegno)
        Me.Customer.Controls.Add(Me.txtVATRegno)
        Me.Customer.Controls.Add(Me.lblCompanyName)
        Me.Customer.Controls.Add(Me.txtCompanyName)
        Me.Customer.Controls.Add(Me.lblIdentity)
        Me.Customer.Controls.Add(Me.txtIdentity)
        Me.Customer.Controls.Add(Me.Label10)
        Me.Customer.Controls.Add(Me.Label12)
        Me.Customer.Controls.Add(Me.Label22)
        Me.Customer.Location = New System.Drawing.Point(4, 22)
        Me.Customer.Name = "Customer"
        Me.Customer.Size = New System.Drawing.Size(920, 627)
        Me.Customer.TabIndex = 0
        Me.Customer.Text = "Customer Details"
        '
        'lblCountry
        '
        Me.lblCountry.AutoSize = True
        Me.lblCountry.Location = New System.Drawing.Point(705, 108)
        Me.lblCountry.Name = "lblCountry"
        Me.lblCountry.Size = New System.Drawing.Size(19, 13)
        Me.lblCountry.TabIndex = 67
        Me.lblCountry.Text = "----"
        '
        'cbDivision
        '
        Me.cbDivision.Location = New System.Drawing.Point(608, 14)
        Me.cbDivision.Name = "cbDivision"
        Me.cbDivision.Size = New System.Drawing.Size(121, 21)
        Me.cbDivision.TabIndex = 66
        Me.ToolTip1.SetToolTip(Me.cbDivision, "Division of group who manages customer")
        '
        'Label88
        '
        Me.Label88.Location = New System.Drawing.Point(544, 16)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(48, 16)
        Me.Label88.TabIndex = 65
        Me.Label88.Text = "Division"
        Me.Label88.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label79
        '
        Me.Label79.Location = New System.Drawing.Point(24, 320)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(112, 24)
        Me.Label79.TabIndex = 64
        Me.Label79.Text = "Cust Centre Comments"
        Me.Label79.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label78
        '
        Me.Label78.Location = New System.Drawing.Point(24, 432)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(112, 23)
        Me.Label78.TabIndex = 63
        Me.Label78.Text = "Account Type"
        Me.Label78.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAccountType
        '
        Me.txtAccountType.Location = New System.Drawing.Point(144, 433)
        Me.txtAccountType.Name = "txtAccountType"
        Me.txtAccountType.Size = New System.Drawing.Size(544, 20)
        Me.txtAccountType.TabIndex = 62
        Me.ToolTip1.SetToolTip(Me.txtAccountType, "Account information for certificates etc")
        '
        'cmdPasswordLetter
        '
        Me.cmdPasswordLetter.Image = CType(resources.GetObject("cmdPasswordLetter.Image"), System.Drawing.Image)
        Me.cmdPasswordLetter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPasswordLetter.Location = New System.Drawing.Point(624, 360)
        Me.cmdPasswordLetter.Name = "cmdPasswordLetter"
        Me.cmdPasswordLetter.Size = New System.Drawing.Size(112, 23)
        Me.cmdPasswordLetter.TabIndex = 61
        Me.cmdPasswordLetter.Text = "Password Letter"
        Me.cmdPasswordLetter.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPassword
        '
        Me.cmdPassword.Image = CType(resources.GetObject("cmdPassword.Image"), System.Drawing.Image)
        Me.cmdPassword.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPassword.Location = New System.Drawing.Point(504, 360)
        Me.cmdPassword.Name = "cmdPassword"
        Me.cmdPassword.Size = New System.Drawing.Size(112, 23)
        Me.cmdPassword.TabIndex = 60
        Me.cmdPassword.Text = "Add &Password"
        Me.cmdPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbMROID
        '
        Me.cbMROID.Location = New System.Drawing.Point(144, 400)
        Me.cbMROID.Name = "cbMROID"
        Me.cbMROID.Size = New System.Drawing.Size(544, 21)
        Me.cbMROID.TabIndex = 59
        Me.ToolTip1.SetToolTip(Me.cbMROID, "Medical review Officer who will review results")
        '
        'lblMRO
        '
        Me.lblMRO.Location = New System.Drawing.Point(32, 400)
        Me.lblMRO.Name = "lblMRO"
        Me.lblMRO.Size = New System.Drawing.Size(100, 23)
        Me.lblMRO.TabIndex = 58
        Me.lblMRO.Text = "MRO"
        Me.lblMRO.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmailPassword
        '
        Me.txtEmailPassword.Font = New System.Drawing.Font("Monotype Corsiva", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte), True)
        Me.txtEmailPassword.Location = New System.Drawing.Point(392, 360)
        Me.txtEmailPassword.Name = "txtEmailPassword"
        Me.txtEmailPassword.Size = New System.Drawing.Size(100, 21)
        Me.txtEmailPassword.TabIndex = 57
        Me.txtEmailPassword.Text = "1liO0"
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(304, 360)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(88, 16)
        Me.Label42.TabIndex = 56
        Me.Label42.Text = "PDF Password"
        Me.Label42.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(112, 48)
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(608, 20)
        Me.txtComments.TabIndex = 55
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(8, 48)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(96, 16)
        Me.Label19.TabIndex = 54
        Me.Label19.Text = "Comments"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtProfile
        '
        Me.txtProfile.Location = New System.Drawing.Point(432, 16)
        Me.txtProfile.Name = "txtProfile"
        Me.txtProfile.Size = New System.Drawing.Size(100, 20)
        Me.txtProfile.TabIndex = 53
        Me.ToolTip1.SetToolTip(Me.txtProfile, "Customer Profile within SMID")
        '
        'lblProfile
        '
        Me.lblProfile.Location = New System.Drawing.Point(376, 16)
        Me.lblProfile.Name = "lblProfile"
        Me.lblProfile.Size = New System.Drawing.Size(96, 16)
        Me.lblProfile.TabIndex = 52
        Me.lblProfile.Text = "Profile"
        Me.lblProfile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSMID
        '
        Me.txtSMID.Location = New System.Drawing.Point(264, 14)
        Me.txtSMID.Name = "txtSMID"
        Me.txtSMID.Size = New System.Drawing.Size(100, 20)
        Me.txtSMID.TabIndex = 51
        Me.ToolTip1.SetToolTip(Me.txtSMID, "Sample Manager ID")
        '
        'cbSalesPerson
        '
        Me.cbSalesPerson.Location = New System.Drawing.Point(312, 158)
        Me.cbSalesPerson.Name = "cbSalesPerson"
        Me.cbSalesPerson.Size = New System.Drawing.Size(120, 21)
        Me.cbSalesPerson.TabIndex = 50
        '
        'cbProgMan
        '
        Me.cbProgMan.Location = New System.Drawing.Point(144, 360)
        Me.cbProgMan.Name = "cbProgMan"
        Me.cbProgMan.Size = New System.Drawing.Size(121, 21)
        Me.cbProgMan.TabIndex = 49
        '
        'cbStatus
        '
        Me.cbStatus.DisplayMember = "PhraseText"
        Me.cbStatus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbStatus.Location = New System.Drawing.Point(112, 158)
        Me.cbStatus.Name = "cbStatus"
        Me.cbStatus.PhraseType = Nothing
        Me.cbStatus.Size = New System.Drawing.Size(121, 21)
        Me.cbStatus.TabIndex = 48
        Me.ToolTip1.SetToolTip(Me.cbStatus, "Status of this account in the system")
        Me.cbStatus.ValueMember = "PhraseID"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 360)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 23)
        Me.Label13.TabIndex = 45
        Me.Label13.Text = "Programme Manager"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblContractDate
        '
        Me.lblContractDate.Location = New System.Drawing.Point(264, 104)
        Me.lblContractDate.Name = "lblContractDate"
        Me.lblContractDate.Size = New System.Drawing.Size(86, 23)
        Me.lblContractDate.TabIndex = 44
        Me.lblContractDate.Text = "Contract Expiry"
        Me.lblContractDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtContractDate
        '
        Me.dtContractDate.CustomFormat = "dd-MMM-yyyy"
        Me.dtContractDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtContractDate.Location = New System.Drawing.Point(352, 104)
        Me.dtContractDate.Name = "dtContractDate"
        Me.dtContractDate.Size = New System.Drawing.Size(96, 20)
        Me.dtContractDate.TabIndex = 43
        '
        'txtPinNumber
        '
        Me.txtPinNumber.Location = New System.Drawing.Point(496, 158)
        Me.txtPinNumber.Name = "txtPinNumber"
        Me.txtPinNumber.Size = New System.Drawing.Size(100, 20)
        Me.txtPinNumber.TabIndex = 42
        '
        'lblPinNumber
        '
        Me.lblPinNumber.Location = New System.Drawing.Point(400, 158)
        Me.lblPinNumber.Name = "lblPinNumber"
        Me.lblPinNumber.Size = New System.Drawing.Size(100, 23)
        Me.lblPinNumber.TabIndex = 41
        Me.lblPinNumber.Text = "Pin Number"
        Me.lblPinNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSubContOf
        '
        Me.txtSubContOf.Location = New System.Drawing.Point(144, 320)
        Me.txtSubContOf.Name = "txtSubContOf"
        Me.txtSubContOf.Size = New System.Drawing.Size(544, 20)
        Me.txtSubContOf.TabIndex = 40
        '
        'cbIndustry
        '
        Me.cbIndustry.Location = New System.Drawing.Point(144, 288)
        Me.cbIndustry.Name = "cbIndustry"
        Me.cbIndustry.Size = New System.Drawing.Size(544, 21)
        Me.cbIndustry.TabIndex = 38
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(32, 288)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Industry"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CBCustomerType
        '
        Me.CBCustomerType.Location = New System.Drawing.Point(144, 264)
        Me.CBCustomerType.Name = "CBCustomerType"
        Me.CBCustomerType.Size = New System.Drawing.Size(544, 21)
        Me.CBCustomerType.TabIndex = 36
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(32, 264)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 23)
        Me.Label8.TabIndex = 35
        Me.Label8.Text = "Profile Type"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblContractNumber
        '
        Me.lblContractNumber.Location = New System.Drawing.Point(16, 112)
        Me.lblContractNumber.Name = "lblContractNumber"
        Me.lblContractNumber.Size = New System.Drawing.Size(96, 23)
        Me.lblContractNumber.TabIndex = 16
        Me.lblContractNumber.Text = "Customer's Ref."
        Me.lblContractNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtContractNumber
        '
        Me.txtContractNumber.Location = New System.Drawing.Point(112, 104)
        Me.txtContractNumber.Name = "txtContractNumber"
        Me.txtContractNumber.Size = New System.Drawing.Size(136, 20)
        Me.txtContractNumber.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtContractNumber, "Reference supplied by the customer," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Examples are Purchase Order number," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Site ID" & _
                "," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Contract Number ")
        '
        'lblSalesPerson
        '
        Me.lblSalesPerson.Location = New System.Drawing.Point(208, 158)
        Me.lblSalesPerson.Name = "lblSalesPerson"
        Me.lblSalesPerson.Size = New System.Drawing.Size(100, 23)
        Me.lblSalesPerson.TabIndex = 13
        Me.lblSalesPerson.Text = "Sales Person"
        Me.lblSalesPerson.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParent
        '
        Me.cbParent.Location = New System.Drawing.Point(112, 130)
        Me.cbParent.Name = "cbParent"
        Me.cbParent.Size = New System.Drawing.Size(380, 21)
        Me.cbParent.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.cbParent, "This SMID Profile is a child account of the account shown")
        '
        'lblParent
        '
        Me.lblParent.Location = New System.Drawing.Point(18, 128)
        Me.lblParent.Name = "lblParent"
        Me.lblParent.Size = New System.Drawing.Size(88, 23)
        Me.lblParent.TabIndex = 11
        Me.lblParent.Text = "Parent "
        Me.lblParent.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ckStopCerts)
        Me.GroupBox2.Controls.Add(Me.CKStopSamples)
        Me.GroupBox2.Controls.Add(Me.CkRemoved)
        Me.GroupBox2.Controls.Add(Me.CKSuspended)
        Me.GroupBox2.Controls.Add(Me.ckNewCustomer)
        Me.GroupBox2.Enabled = False
        Me.GroupBox2.Location = New System.Drawing.Point(24, 184)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(688, 64)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Flags"
        '
        'ckStopCerts
        '
        Me.ckStopCerts.Location = New System.Drawing.Point(520, 24)
        Me.ckStopCerts.Name = "ckStopCerts"
        Me.ckStopCerts.Size = New System.Drawing.Size(144, 24)
        Me.ckStopCerts.TabIndex = 4
        Me.ckStopCerts.Text = "Stop &Certificates"
        '
        'CKStopSamples
        '
        Me.CKStopSamples.Location = New System.Drawing.Point(416, 24)
        Me.CKStopSamples.Name = "CKStopSamples"
        Me.CKStopSamples.Size = New System.Drawing.Size(104, 24)
        Me.CKStopSamples.TabIndex = 3
        Me.CKStopSamples.Text = "&Stop Samples"
        '
        'CkRemoved
        '
        Me.CkRemoved.Enabled = False
        Me.CkRemoved.Location = New System.Drawing.Point(296, 24)
        Me.CkRemoved.Name = "CkRemoved"
        Me.CkRemoved.Size = New System.Drawing.Size(144, 24)
        Me.CkRemoved.TabIndex = 2
        Me.CkRemoved.Text = "Account &Removed"
        '
        'CKSuspended
        '
        Me.CKSuspended.Location = New System.Drawing.Point(144, 24)
        Me.CKSuspended.Name = "CKSuspended"
        Me.CKSuspended.Size = New System.Drawing.Size(136, 24)
        Me.CKSuspended.TabIndex = 1
        Me.CKSuspended.Text = "&Account Suspended"
        '
        'ckNewCustomer
        '
        Me.ckNewCustomer.Location = New System.Drawing.Point(24, 24)
        Me.ckNewCustomer.Name = "ckNewCustomer"
        Me.ckNewCustomer.Size = New System.Drawing.Size(104, 24)
        Me.ckNewCustomer.TabIndex = 0
        Me.ckNewCustomer.Text = "&New Account"
        '
        'lblVATRegno
        '
        Me.lblVATRegno.Location = New System.Drawing.Point(472, 104)
        Me.lblVATRegno.Name = "lblVATRegno"
        Me.lblVATRegno.Size = New System.Drawing.Size(96, 16)
        Me.lblVATRegno.TabIndex = 9
        Me.lblVATRegno.Text = "Company VAT No"
        Me.lblVATRegno.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVATRegno
        '
        Me.txtVATRegno.Location = New System.Drawing.Point(568, 104)
        Me.txtVATRegno.Name = "txtVATRegno"
        Me.txtVATRegno.Size = New System.Drawing.Size(136, 20)
        Me.txtVATRegno.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtVATRegno, "VAT Registration number set to Country code of Company's country of registration")
        '
        'lblCompanyName
        '
        Me.lblCompanyName.Location = New System.Drawing.Point(16, 77)
        Me.lblCompanyName.Name = "lblCompanyName"
        Me.lblCompanyName.Size = New System.Drawing.Size(96, 16)
        Me.lblCompanyName.TabIndex = 5
        Me.lblCompanyName.Text = "Customer Name"
        Me.lblCompanyName.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCompanyName.Location = New System.Drawing.Point(112, 77)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(576, 20)
        Me.txtCompanyName.TabIndex = 4
        '
        'lblIdentity
        '
        Me.lblIdentity.Location = New System.Drawing.Point(8, 16)
        Me.lblIdentity.Name = "lblIdentity"
        Me.lblIdentity.Size = New System.Drawing.Size(96, 16)
        Me.lblIdentity.TabIndex = 3
        Me.lblIdentity.Text = "Customer Identity"
        Me.lblIdentity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtIdentity
        '
        Me.txtIdentity.BackColor = System.Drawing.SystemColors.Control
        Me.txtIdentity.Enabled = False
        Me.txtIdentity.Location = New System.Drawing.Point(112, 16)
        Me.txtIdentity.Name = "txtIdentity"
        Me.txtIdentity.Size = New System.Drawing.Size(88, 20)
        Me.txtIdentity.TabIndex = 2
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 16)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "Customer Identity"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(206, 16)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 16)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "SMID"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(56, 158)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(48, 23)
        Me.Label22.TabIndex = 47
        Me.Label22.Text = "Status"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabContacts
        '
        Me.tabContacts.Controls.Add(Me.pnlContactDetails)
        Me.tabContacts.Controls.Add(Me.Splitter3)
        Me.tabContacts.Controls.Add(Me.pnlContactList)
        Me.tabContacts.Location = New System.Drawing.Point(4, 22)
        Me.tabContacts.Name = "tabContacts"
        Me.tabContacts.Size = New System.Drawing.Size(920, 627)
        Me.tabContacts.TabIndex = 10
        Me.tabContacts.Text = "Contacts"
        '
        'pnlContactDetails
        '
        Me.pnlContactDetails.Controls.Add(Me.pnlContact)
        Me.pnlContactDetails.Controls.Add(Me.Splitter4)
        Me.pnlContactDetails.Controls.Add(Me.AdrContact)
        Me.pnlContactDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContactDetails.Location = New System.Drawing.Point(0, 259)
        Me.pnlContactDetails.Name = "pnlContactDetails"
        Me.pnlContactDetails.Size = New System.Drawing.Size(920, 368)
        Me.pnlContactDetails.TabIndex = 2
        '
        'pnlContact
        '
        Me.pnlContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlContact.Controls.Add(Me.ckContactRemoved)
        Me.pnlContact.Controls.Add(Me.cbContactStatus)
        Me.pnlContact.Controls.Add(Me.Label87)
        Me.pnlContact.Controls.Add(Me.cmdUpdateContact)
        Me.pnlContact.Controls.Add(Me.cbReportOptions)
        Me.pnlContact.Controls.Add(Me.Label85)
        Me.pnlContact.Controls.Add(Me.cbReportFormat)
        Me.pnlContact.Controls.Add(Me.Label86)
        Me.pnlContact.Controls.Add(Me.cbContMethod)
        Me.pnlContact.Controls.Add(Me.Label84)
        Me.pnlContact.Controls.Add(Me.cbContactType)
        Me.pnlContact.Controls.Add(Me.Label83)
        Me.pnlContact.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContact.Location = New System.Drawing.Point(712, 0)
        Me.pnlContact.Name = "pnlContact"
        Me.pnlContact.Size = New System.Drawing.Size(208, 368)
        Me.pnlContact.TabIndex = 238
        '
        'ckContactRemoved
        '
        Me.ckContactRemoved.Location = New System.Drawing.Point(40, 296)
        Me.ckContactRemoved.Name = "ckContactRemoved"
        Me.ckContactRemoved.Size = New System.Drawing.Size(1200, 24)
        Me.ckContactRemoved.TabIndex = 236
        Me.ckContactRemoved.Text = "Removed"
        '
        'cbContactStatus
        '
        Me.cbContactStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContactStatus.DisplayMember = "PhraseText"
        Me.cbContactStatus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContactStatus.Location = New System.Drawing.Point(44, 264)
        Me.cbContactStatus.Name = "cbContactStatus"
        Me.cbContactStatus.PhraseType = Nothing
        Me.cbContactStatus.Size = New System.Drawing.Size(136, 21)
        Me.cbContactStatus.TabIndex = 235
        Me.cbContactStatus.ValueMember = "PhraseID"
        '
        'Label87
        '
        Me.Label87.Location = New System.Drawing.Point(44, 240)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(100, 23)
        Me.Label87.TabIndex = 234
        Me.Label87.Text = "Contact Status"
        '
        'cmdUpdateContact
        '
        Me.cmdUpdateContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdateContact.Location = New System.Drawing.Point(40, 328)
        Me.cmdUpdateContact.Name = "cmdUpdateContact"
        Me.cmdUpdateContact.Size = New System.Drawing.Size(120, 23)
        Me.cmdUpdateContact.TabIndex = 233
        Me.cmdUpdateContact.Text = "Save"
        Me.cmdUpdateContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbReportOptions
        '
        Me.cbReportOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReportOptions.DisplayMember = "PhraseText"
        Me.cbReportOptions.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbReportOptions.Location = New System.Drawing.Point(40, 212)
        Me.cbReportOptions.Name = "cbReportOptions"
        Me.cbReportOptions.PhraseType = Nothing
        Me.cbReportOptions.Size = New System.Drawing.Size(136, 21)
        Me.cbReportOptions.TabIndex = 7
        Me.cbReportOptions.ValueMember = "PhraseID"
        '
        'Label85
        '
        Me.Label85.Location = New System.Drawing.Point(40, 188)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(100, 23)
        Me.Label85.TabIndex = 6
        Me.Label85.Text = "Report Options"
        '
        'cbReportFormat
        '
        Me.cbReportFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReportFormat.Location = New System.Drawing.Point(40, 156)
        Me.cbReportFormat.Name = "cbReportFormat"
        Me.cbReportFormat.Size = New System.Drawing.Size(136, 21)
        Me.cbReportFormat.TabIndex = 5
        '
        'Label86
        '
        Me.Label86.Location = New System.Drawing.Point(40, 132)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(100, 23)
        Me.Label86.TabIndex = 4
        Me.Label86.Text = "Report Format"
        '
        'cbContMethod
        '
        Me.cbContMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContMethod.DisplayMember = "PhraseText"
        Me.cbContMethod.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContMethod.Location = New System.Drawing.Point(40, 96)
        Me.cbContMethod.Name = "cbContMethod"
        Me.cbContMethod.PhraseType = Nothing
        Me.cbContMethod.Size = New System.Drawing.Size(136, 21)
        Me.cbContMethod.TabIndex = 3
        Me.cbContMethod.ValueMember = "PhraseID"
        '
        'Label84
        '
        Me.Label84.Location = New System.Drawing.Point(40, 72)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(100, 23)
        Me.Label84.TabIndex = 2
        Me.Label84.Text = "Contact Method"
        '
        'cbContactType
        '
        Me.cbContactType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContactType.DisplayMember = "PhraseText"
        Me.cbContactType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContactType.Location = New System.Drawing.Point(40, 40)
        Me.cbContactType.Name = "cbContactType"
        Me.cbContactType.PhraseType = Nothing
        Me.cbContactType.Size = New System.Drawing.Size(136, 21)
        Me.cbContactType.TabIndex = 1
        Me.cbContactType.ValueMember = "PhraseID"
        '
        'Label83
        '
        Me.Label83.Location = New System.Drawing.Point(40, 16)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(100, 23)
        Me.Label83.TabIndex = 0
        Me.Label83.Text = "Contact Type"
        '
        'Splitter4
        '
        Me.Splitter4.Location = New System.Drawing.Point(704, 0)
        Me.Splitter4.Name = "Splitter4"
        Me.Splitter4.Size = New System.Drawing.Size(8, 368)
        Me.Splitter4.TabIndex = 18
        Me.Splitter4.TabStop = False
        '
        'AdrContact
        '
        Me.AdrContact.Address = Nothing
        Me.AdrContact.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AdrContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AdrContact.Controls.Add(Me.Label81)
        Me.AdrContact.Controls.Add(Me.Label82)
        Me.AdrContact.Controls.Add(Me.cmdFindContactAdr)
        Me.AdrContact.Controls.Add(Me.cmdNewContactAdr)
        Me.AdrContact.Dock = System.Windows.Forms.DockStyle.Left
        Me.AdrContact.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AdrContact.InfoVisible = False
        Me.AdrContact.isReadOnly = False
        Me.AdrContact.Location = New System.Drawing.Point(0, 0)
        Me.AdrContact.Name = "AdrContact"
        Me.AdrContact.ShowCancelCmd = False
        Me.AdrContact.ShowCloneCmd = False
        Me.AdrContact.ShowContact = True
        Me.AdrContact.ShowEditCmd = False
        Me.AdrContact.ShowFindCmd = False
        Me.AdrContact.ShowNewCmd = False
        Me.AdrContact.ShowPanel = False
        Me.AdrContact.ShowSaveCmd = False
        Me.AdrContact.ShowStatus = False
        Me.AdrContact.Size = New System.Drawing.Size(704, 368)
        Me.AdrContact.TabIndex = 17
        Me.AdrContact.Tooltip = Me.ToolTip1
        Me.ToolTip1.SetToolTip(Me.AdrContact, "Contacts address")
        '
        'Label81
        '
        Me.Label81.Location = New System.Drawing.Point(392, 72)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(16, 16)
        Me.Label81.TabIndex = 235
        '
        'Label82
        '
        Me.Label82.Location = New System.Drawing.Point(48, 184)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(16, 16)
        Me.Label82.TabIndex = 234
        '
        'cmdFindContactAdr
        '
        Me.cmdFindContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindContactAdr.Location = New System.Drawing.Point(192, 256)
        Me.cmdFindContactAdr.Name = "cmdFindContactAdr"
        Me.cmdFindContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdFindContactAdr.TabIndex = 233
        Me.cmdFindContactAdr.Text = "Find Address"
        Me.cmdFindContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdNewContactAdr
        '
        Me.cmdNewContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewContactAdr.Location = New System.Drawing.Point(64, 256)
        Me.cmdNewContactAdr.Name = "cmdNewContactAdr"
        Me.cmdNewContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewContactAdr.TabIndex = 232
        Me.cmdNewContactAdr.Text = "New Address"
        Me.cmdNewContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'lvContacts
        '
        Me.lvContacts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader23, Me.ColumnHeader55, Me.ColumnHeader24, Me.ColumnHeader28, Me.ColumnHeader25, Me.ColumnHeader26, Me.ColumnHeader27})
        Me.lvContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvContacts.FullRowSelect = True
        Me.lvContacts.Location = New System.Drawing.Point(0, 0)
        Me.lvContacts.Name = "lvContacts"
        Me.lvContacts.Size = New System.Drawing.Size(920, 200)
        Me.lvContacts.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.lvContacts, "List of contacts for this customer, click to edit")
        Me.lvContacts.UseCompatibleStateImageBehavior = False
        Me.lvContacts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "Contact"
        Me.ColumnHeader23.Width = 180
        '
        'ColumnHeader55
        '
        Me.ColumnHeader55.Text = "AddressID"
        '
        'ColumnHeader24
        '
        Me.ColumnHeader24.Text = "Contact Type"
        Me.ColumnHeader24.Width = 100
        '
        'ColumnHeader28
        '
        Me.ColumnHeader28.Text = "Status"
        Me.ColumnHeader28.Width = 80
        '
        'ColumnHeader25
        '
        Me.ColumnHeader25.Text = "Contact Method"
        Me.ColumnHeader25.Width = 121
        '
        'ColumnHeader26
        '
        Me.ColumnHeader26.Text = "Report Format"
        Me.ColumnHeader26.Width = 120
        '
        'ColumnHeader27
        '
        Me.ColumnHeader27.Text = "Report Options"
        Me.ColumnHeader27.Width = 144
        '
        'tabHistory
        '
        Me.tabHistory.Controls.Add(Me.pnlHistory)
        Me.tabHistory.Location = New System.Drawing.Point(4, 22)
        Me.tabHistory.Name = "tabHistory"
        Me.tabHistory.Size = New System.Drawing.Size(920, 627)
        Me.tabHistory.TabIndex = 7
        Me.tabHistory.Text = "History"
        Me.ToolTip1.SetToolTip(Me.tabHistory, "History of changes to the customer")
        '
        'pnlHistory
        '
        Me.pnlHistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlHistory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlHistory.Controls.Add(Me.HtmlEditor2)
        Me.pnlHistory.Location = New System.Drawing.Point(16, 8)
        Me.pnlHistory.Name = "pnlHistory"
        Me.pnlHistory.Size = New System.Drawing.Size(896, 568)
        Me.pnlHistory.TabIndex = 0
        '
        'HtmlEditor2
        '
        Me.HtmlEditor2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlEditor2.Location = New System.Drawing.Point(0, 0)
        Me.HtmlEditor2.Name = "HtmlEditor2"
        Me.HtmlEditor2.Size = New System.Drawing.Size(892, 564)
        Me.HtmlEditor2.TabIndex = 2
        '
        'cmdCustomerOptions
        '
        Me.cmdCustomerOptions.Location = New System.Drawing.Point(32, 8)
        Me.cmdCustomerOptions.Name = "cmdCustomerOptions"
        Me.cmdCustomerOptions.Size = New System.Drawing.Size(75, 23)
        Me.cmdCustomerOptions.TabIndex = 14
        Me.cmdCustomerOptions.Text = "Options"
        Me.ToolTip1.SetToolTip(Me.cmdCustomerOptions, "Edit customer options")
        '
        'Splitter3
        '
        Me.Splitter3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter3.Location = New System.Drawing.Point(0, 256)
        Me.Splitter3.Name = "Splitter3"
        Me.Splitter3.Size = New System.Drawing.Size(920, 3)
        Me.Splitter3.TabIndex = 1
        Me.Splitter3.TabStop = False
        '
        'pnlContactList
        '
        Me.pnlContactList.Controls.Add(Me.cmdDelContact)
        Me.pnlContactList.Controls.Add(Me.cmdAddContact)
        Me.pnlContactList.Controls.Add(Me.lvContacts)
        Me.pnlContactList.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactList.Location = New System.Drawing.Point(0, 0)
        Me.pnlContactList.Name = "pnlContactList"
        Me.pnlContactList.Size = New System.Drawing.Size(920, 256)
        Me.pnlContactList.TabIndex = 0
        '
        'cmdDelContact
        '
        Me.cmdDelContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDelContact.Location = New System.Drawing.Point(152, 216)
        Me.cmdDelContact.Name = "cmdDelContact"
        Me.cmdDelContact.Size = New System.Drawing.Size(96, 23)
        Me.cmdDelContact.TabIndex = 2
        Me.cmdDelContact.Text = "Del Contact"
        Me.cmdDelContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAddContact
        '
        Me.cmdAddContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddContact.Location = New System.Drawing.Point(40, 216)
        Me.cmdAddContact.Name = "cmdAddContact"
        Me.cmdAddContact.Size = New System.Drawing.Size(96, 23)
        Me.cmdAddContact.TabIndex = 1
        Me.cmdAddContact.Text = "Add Contact"
        Me.cmdAddContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabAddresses
        '
        Me.TabAddresses.Controls.Add(Me.GroupBox1)
        Me.TabAddresses.Controls.Add(Me.Panel4)
        Me.TabAddresses.Location = New System.Drawing.Point(4, 22)
        Me.TabAddresses.Name = "TabAddresses"
        Me.TabAddresses.Size = New System.Drawing.Size(920, 627)
        Me.TabAddresses.TabIndex = 1
        Me.TabAddresses.Text = "Addresses"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblReportingMethod)
        Me.GroupBox1.Controls.Add(Me.rbGoldmine)
        Me.GroupBox1.Controls.Add(Me.rbInvShipping)
        Me.GroupBox1.Controls.Add(Me.RBInvMail)
        Me.GroupBox1.Controls.Add(Me.RBInvoicing)
        Me.GroupBox1.Controls.Add(Me.RBSecondary)
        Me.GroupBox1.Controls.Add(Me.RBAddress)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 344)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(920, 160)
        Me.GroupBox1.TabIndex = 31
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Address displayed"
        '
        'lblReportingMethod
        '
        Me.lblReportingMethod.Location = New System.Drawing.Point(24, 80)
        Me.lblReportingMethod.Name = "lblReportingMethod"
        Me.lblReportingMethod.Size = New System.Drawing.Size(448, 23)
        Me.lblReportingMethod.TabIndex = 9
        '
        'rbGoldmine
        '
        Me.rbGoldmine.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbGoldmine.Location = New System.Drawing.Point(24, 56)
        Me.rbGoldmine.Name = "rbGoldmine"
        Me.rbGoldmine.Size = New System.Drawing.Size(120, 24)
        Me.rbGoldmine.TabIndex = 8
        Me.rbGoldmine.Text = "Goldmine Address"
        '
        'rbInvShipping
        '
        Me.rbInvShipping.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbInvShipping.Location = New System.Drawing.Point(536, 24)
        Me.rbInvShipping.Name = "rbInvShipping"
        Me.rbInvShipping.Size = New System.Drawing.Size(120, 24)
        Me.rbInvShipping.TabIndex = 7
        Me.rbInvShipping.Text = "Invoicing Shipping"
        '
        'RBInvMail
        '
        Me.RBInvMail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBInvMail.Location = New System.Drawing.Point(408, 24)
        Me.RBInvMail.Name = "RBInvMail"
        Me.RBInvMail.Size = New System.Drawing.Size(104, 24)
        Me.RBInvMail.TabIndex = 6
        Me.RBInvMail.Text = "Invoicing Mail "
        '
        'RBInvoicing
        '
        Me.RBInvoicing.Location = New System.Drawing.Point(280, 24)
        Me.RBInvoicing.Name = "RBInvoicing"
        Me.RBInvoicing.Size = New System.Drawing.Size(104, 24)
        Me.RBInvoicing.TabIndex = 4
        Me.RBInvoicing.Text = "Invoicing"
        '
        'RBSecondary
        '
        Me.RBSecondary.Location = New System.Drawing.Point(152, 24)
        Me.RBSecondary.Name = "RBSecondary"
        Me.RBSecondary.Size = New System.Drawing.Size(104, 24)
        Me.RBSecondary.TabIndex = 1
        Me.RBSecondary.Text = "Secondary"
        '
        'RBAddress
        '
        Me.RBAddress.Checked = True
        Me.RBAddress.Location = New System.Drawing.Point(24, 24)
        Me.RBAddress.Name = "RBAddress"
        Me.RBAddress.Size = New System.Drawing.Size(104, 24)
        Me.RBAddress.TabIndex = 0
        Me.RBAddress.TabStop = True
        Me.RBAddress.Text = "Results "
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.txtSameAddress)
        Me.Panel4.Controls.Add(Me.SplitterAddr)
        Me.Panel4.Controls.Add(Me.AddressControl1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(920, 344)
        Me.Panel4.TabIndex = 32
        '
        'txtSameAddress
        '
        Me.txtSameAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSameAddress.Location = New System.Drawing.Point(635, 0)
        Me.txtSameAddress.Name = "txtSameAddress"
        Me.txtSameAddress.Size = New System.Drawing.Size(285, 344)
        Me.txtSameAddress.TabIndex = 0
        Me.txtSameAddress.Text = ""
        '
        'SplitterAddr
        '
        Me.SplitterAddr.Location = New System.Drawing.Point(632, 0)
        Me.SplitterAddr.Name = "SplitterAddr"
        Me.SplitterAddr.Size = New System.Drawing.Size(3, 344)
        Me.SplitterAddr.TabIndex = 1
        Me.SplitterAddr.TabStop = False
        '
        'AddressControl1
        '
        Me.AddressControl1.Address = Nothing
        Me.AddressControl1.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AddressControl1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AddressControl1.Controls.Add(Me.Label20)
        Me.AddressControl1.Controls.Add(Me.lblFaxError)
        Me.AddressControl1.Controls.Add(Me.lblEmailError)
        Me.AddressControl1.Controls.Add(Me.txtEmailPositives)
        Me.AddressControl1.Controls.Add(Me.lblPosEmail)
        Me.AddressControl1.Dock = System.Windows.Forms.DockStyle.Left
        Me.AddressControl1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AddressControl1.InfoVisible = False
        Me.AddressControl1.isReadOnly = False
        Me.AddressControl1.Location = New System.Drawing.Point(0, 0)
        Me.AddressControl1.Name = "AddressControl1"
        Me.AddressControl1.ShowCancelCmd = True
        Me.AddressControl1.ShowCloneCmd = True
        Me.AddressControl1.ShowContact = True
        Me.AddressControl1.ShowEditCmd = False
        Me.AddressControl1.ShowFindCmd = True
        Me.AddressControl1.ShowNewCmd = True
        Me.AddressControl1.ShowPanel = True
        Me.AddressControl1.ShowSaveCmd = True
        Me.AddressControl1.ShowStatus = False
        Me.AddressControl1.Size = New System.Drawing.Size(632, 344)
        Me.AddressControl1.TabIndex = 16
        Me.AddressControl1.Tooltip = Nothing
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(40, 280)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(304, 23)
        Me.Label20.TabIndex = 236
        Me.Label20.Text = "Positive result address is the secondary address"
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
        'txtEmailPositives
        '
        Me.txtEmailPositives.Enabled = False
        Me.txtEmailPositives.Location = New System.Drawing.Point(112, 250)
        Me.txtEmailPositives.Name = "txtEmailPositives"
        Me.txtEmailPositives.Size = New System.Drawing.Size(272, 20)
        Me.txtEmailPositives.TabIndex = 33
        '
        'lblPosEmail
        '
        Me.lblPosEmail.Location = New System.Drawing.Point(8, 250)
        Me.lblPosEmail.Name = "lblPosEmail"
        Me.lblPosEmail.Size = New System.Drawing.Size(96, 23)
        Me.lblPosEmail.TabIndex = 32
        Me.lblPosEmail.Text = "Positives Email"
        Me.lblPosEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabPricing
        '
        Me.tabPricing.Controls.Add(Me.cmdChangeDate)
        Me.tabPricing.Controls.Add(Me.DTPricing)
        Me.tabPricing.Controls.Add(Me.Label61)
        Me.tabPricing.Controls.Add(Me.PictureBox1)
        Me.tabPricing.Controls.Add(Me.Label7)
        Me.tabPricing.Controls.Add(Me.lblSerSold)
        Me.tabPricing.Controls.Add(Me.lblSerNotSoldYet)
        Me.tabPricing.Controls.Add(Me.lblSerWasSold)
        Me.tabPricing.Controls.Add(Me.lblSerNotSold)
        Me.tabPricing.Controls.Add(Me.tabTypes)
        Me.tabPricing.Location = New System.Drawing.Point(4, 22)
        Me.tabPricing.Name = "tabPricing"
        Me.tabPricing.Size = New System.Drawing.Size(920, 627)
        Me.tabPricing.TabIndex = 9
        Me.tabPricing.Text = "Pricing"
        '
        'cmdChangeDate
        '
        Me.cmdChangeDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdChangeDate.Image = CType(resources.GetObject("cmdChangeDate.Image"), System.Drawing.Image)
        Me.cmdChangeDate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdChangeDate.Location = New System.Drawing.Point(576, 528)
        Me.cmdChangeDate.Name = "cmdChangeDate"
        Me.cmdChangeDate.Size = New System.Drawing.Size(72, 24)
        Me.cmdChangeDate.TabIndex = 66
        Me.cmdChangeDate.Text = "Set Date"
        Me.cmdChangeDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTPricing
        '
        Me.DTPricing.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DTPricing.Location = New System.Drawing.Point(424, 528)
        Me.DTPricing.Name = "DTPricing"
        Me.DTPricing.Size = New System.Drawing.Size(144, 20)
        Me.DTPricing.TabIndex = 65
        '
        'Label61
        '
        Me.Label61.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label61.Location = New System.Drawing.Point(608, 560)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(136, 23)
        Me.Label61.TabIndex = 63
        Me.Label61.Text = "Volume Discount in place"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(584, 560)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Label7.Location = New System.Drawing.Point(400, 560)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(168, 23)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Service previously sold to client "
        '
        'lblSerSold
        '
        Me.lblSerSold.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSerSold.BackColor = System.Drawing.Color.PaleGreen
        Me.lblSerSold.Location = New System.Drawing.Point(224, 560)
        Me.lblSerSold.Name = "lblSerSold"
        Me.lblSerSold.Size = New System.Drawing.Size(168, 23)
        Me.lblSerSold.TabIndex = 14
        Me.lblSerSold.Text = "Service sold to client "
        '
        'lblSerNotSoldYet
        '
        Me.lblSerNotSoldYet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSerNotSoldYet.BackColor = System.Drawing.Color.Orange
        Me.lblSerNotSoldYet.Location = New System.Drawing.Point(48, 530)
        Me.lblSerNotSoldYet.Name = "lblSerNotSoldYet"
        Me.lblSerNotSoldYet.Size = New System.Drawing.Size(168, 23)
        Me.lblSerNotSoldYet.TabIndex = 13
        Me.lblSerNotSoldYet.Text = "Service will be sold to client "
        '
        'lblSerWasSold
        '
        Me.lblSerWasSold.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSerWasSold.BackColor = System.Drawing.Color.CadetBlue
        Me.lblSerWasSold.Location = New System.Drawing.Point(224, 530)
        Me.lblSerWasSold.Name = "lblSerWasSold"
        Me.lblSerWasSold.Size = New System.Drawing.Size(168, 23)
        Me.lblSerWasSold.TabIndex = 64
        Me.lblSerWasSold.Text = "Service expired "
        '
        'lblSerNotSold
        '
        Me.lblSerNotSold.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSerNotSold.BackColor = System.Drawing.Color.Beige
        Me.lblSerNotSold.Location = New System.Drawing.Point(48, 560)
        Me.lblSerNotSold.Name = "lblSerNotSold"
        Me.lblSerNotSold.Size = New System.Drawing.Size(168, 23)
        Me.lblSerNotSold.TabIndex = 13
        Me.lblSerNotSold.Text = "Service not sold to client "
        '
        'tabTypes
        '
        Me.tabTypes.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.tabTypes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabTypes.Controls.Add(Me.tabpSample)
        Me.tabTypes.Controls.Add(Me.tabpMinCharge)
        Me.tabTypes.Controls.Add(Me.tabpAddTime)
        Me.tabTypes.Controls.Add(Me.tabpTests)
        Me.tabTypes.Controls.Add(Me.tabpKits)
        Me.tabTypes.Controls.Add(Me.tabpCancellations)
        Me.tabTypes.Location = New System.Drawing.Point(0, 0)
        Me.tabTypes.Name = "tabTypes"
        Me.tabTypes.SelectedIndex = 0
        Me.tabTypes.Size = New System.Drawing.Size(920, 520)
        Me.tabTypes.TabIndex = 12
        '
        'tabpSample
        '
        Me.tabpSample.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabpSample.Controls.Add(Me.Label71)
        Me.tabpSample.Controls.Add(Me.panel2)
        Me.tabpSample.Location = New System.Drawing.Point(4, 4)
        Me.tabpSample.Name = "tabpSample"
        Me.tabpSample.Size = New System.Drawing.Size(912, 494)
        Me.tabpSample.TabIndex = 0
        Me.tabpSample.Text = "Sample"
        '
        'Label71
        '
        Me.Label71.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label71.BackColor = System.Drawing.SystemColors.Info
        Me.Label71.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label71.Font = New System.Drawing.Font("Times New Roman", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label71.Location = New System.Drawing.Point(88, 400)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(696, 128)
        Me.Label71.TabIndex = 58
        Me.Label71.Text = "Please use the separate pricing form this will be removed in the next release"
        Me.Label71.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panel2
        '
        Me.panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panel2.Controls.Add(Me.pnlMRO)
        Me.panel2.Controls.Add(Me.pnlExternal)
        Me.panel2.Controls.Add(Me.pnlFixedSite)
        Me.panel2.Controls.Add(Me.pnlCallOut)
        Me.panel2.Controls.Add(Me.pnlRoutine)
        Me.panel2.Controls.Add(Me.pnlAO)
        Me.panel2.Location = New System.Drawing.Point(8, 16)
        Me.panel2.Name = "panel2"
        Me.panel2.Size = New System.Drawing.Size(888, 448)
        Me.panel2.TabIndex = 0
        '
        'pnlMRO
        '
        Me.pnlMRO.AllowLeave = True
        Me.pnlMRO.BackColor = System.Drawing.SystemColors.Control
        DiscountPanelBase1.AllowLeave = False
        DiscountPanelBase1.DiscountScheme = ""
        DiscountPanelBase1.Height = 64
        DiscountPanelBase1.Label = "MRO Analysis"
        DiscountPanelBase1.Left = 0
        DiscountPanelBase1.LineItem = ""
        DiscountPanelBase1.Name = Nothing
        DiscountPanelBase1.Panel = Nothing
        DiscountPanelBase1.PanelType = 0
        DiscountPanelBase1.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase1.Service = ""
        DiscountPanelBase1.Top = 0
        DiscountPanelBase1.Width = 884
        Me.pnlMRO.BasePanel = DiscountPanelBase1
        Me.pnlMRO.Client = Nothing
        Me.pnlMRO.Controls.Add(Me.lblMROComment)
        Me.pnlMRO.Discount = ""
        Me.pnlMRO.DiscountScheme = ""
        Me.pnlMRO.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMRO.LineItem = ""
        Me.pnlMRO.ListPrice = ""
        Me.pnlMRO.Location = New System.Drawing.Point(0, 232)
        Me.pnlMRO.Name = "pnlMRO"
        Me.pnlMRO.PanelLabel = "MRO Analysis"
        Me.pnlMRO.PanelService = Nothing
        Me.pnlMRO.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.MRO
        Me.pnlMRO.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlMRO.PriceEnabled = True
        Me.pnlMRO.Scheme = ""
        Me.pnlMRO.Service = ""
        Me.pnlMRO.size = New System.Drawing.Size(884, 64)
        Me.pnlMRO.size = New System.Drawing.Size(884, 64)
        Me.pnlMRO.TabIndex = 56
        '
        'lblMROComment
        '
        Me.lblMROComment.Location = New System.Drawing.Point(16, 32)
        Me.lblMROComment.Name = "lblMROComment"
        Me.lblMROComment.Size = New System.Drawing.Size(624, 23)
        Me.lblMROComment.TabIndex = 41
        '
        'pnlExternal
        '
        Me.pnlExternal.AllowLeave = True
        Me.pnlExternal.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase2.AllowLeave = False
        DiscountPanelBase2.DiscountScheme = ""
        DiscountPanelBase2.Height = 48
        DiscountPanelBase2.Label = "External Samples"
        DiscountPanelBase2.Left = 0
        DiscountPanelBase2.LineItem = ""
        DiscountPanelBase2.Name = Nothing
        DiscountPanelBase2.Panel = Nothing
        DiscountPanelBase2.PanelType = 0
        DiscountPanelBase2.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase2.Service = "EXTERNAL"
        DiscountPanelBase2.Top = 0
        DiscountPanelBase2.Width = 884
        Me.pnlExternal.BasePanel = DiscountPanelBase2
        Me.pnlExternal.Client = Nothing
        Me.pnlExternal.Controls.Add(Me.Label6)
        Me.pnlExternal.Discount = ""
        Me.pnlExternal.DiscountScheme = ""
        Me.pnlExternal.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlExternal.LineItem = ""
        Me.pnlExternal.ListPrice = ""
        Me.pnlExternal.Location = New System.Drawing.Point(0, 184)
        Me.pnlExternal.Name = "pnlExternal"
        Me.pnlExternal.PanelLabel = "External Samples"
        Me.pnlExternal.PanelService = Nothing
        Me.pnlExternal.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Sample
        Me.pnlExternal.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlExternal.PriceEnabled = True
        Me.pnlExternal.Scheme = ""
        Me.pnlExternal.Service = "EXTERNAL"
        Me.pnlExternal.size = New System.Drawing.Size(884, 48)
        Me.pnlExternal.size = New System.Drawing.Size(884, 48)
        Me.pnlExternal.TabIndex = 48
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(32, 320)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 23)
        Me.Label6.TabIndex = 39
        Me.Label6.Text = "Account Info"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlFixedSite
        '
        Me.pnlFixedSite.AllowLeave = True
        Me.pnlFixedSite.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase3.AllowLeave = False
        DiscountPanelBase3.DiscountScheme = ""
        DiscountPanelBase3.Height = 56
        DiscountPanelBase3.Label = "Fixed Site Samples"
        DiscountPanelBase3.Left = 0
        DiscountPanelBase3.LineItem = ""
        DiscountPanelBase3.Name = Nothing
        DiscountPanelBase3.Panel = Nothing
        DiscountPanelBase3.PanelType = 0
        DiscountPanelBase3.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase3.Service = Nothing
        DiscountPanelBase3.Top = 0
        DiscountPanelBase3.Width = 884
        Me.pnlFixedSite.BasePanel = DiscountPanelBase3
        Me.pnlFixedSite.Client = Nothing
        Me.pnlFixedSite.Controls.Add(Me.Label73)
        Me.pnlFixedSite.Discount = ""
        Me.pnlFixedSite.DiscountScheme = ""
        Me.pnlFixedSite.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFixedSite.LineItem = ""
        Me.pnlFixedSite.ListPrice = ""
        Me.pnlFixedSite.Location = New System.Drawing.Point(0, 128)
        Me.pnlFixedSite.Name = "pnlFixedSite"
        Me.pnlFixedSite.PanelLabel = "Fixed Site Samples"
        Me.pnlFixedSite.PanelService = Nothing
        Me.pnlFixedSite.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Sample
        Me.pnlFixedSite.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlFixedSite.PriceEnabled = True
        Me.pnlFixedSite.Scheme = ""
        Me.pnlFixedSite.Service = Nothing
        Me.pnlFixedSite.size = New System.Drawing.Size(884, 56)
        Me.pnlFixedSite.size = New System.Drawing.Size(884, 56)
        Me.pnlFixedSite.TabIndex = 49
        '
        'Label73
        '
        Me.Label73.Location = New System.Drawing.Point(32, 32)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(528, 23)
        Me.Label73.TabIndex = 48
        Me.Label73.Text = "There will be a site handling charge in addition this can vary from site to site " & _
            "and is paid to the site"
        '
        'pnlCallOut
        '
        Me.pnlCallOut.AllowLeave = True
        Me.pnlCallOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase4.AllowLeave = False
        DiscountPanelBase4.DiscountScheme = ""
        DiscountPanelBase4.Height = 48
        DiscountPanelBase4.Label = "Call Out Samples"
        DiscountPanelBase4.Left = 0
        DiscountPanelBase4.LineItem = ""
        DiscountPanelBase4.Name = Nothing
        DiscountPanelBase4.Panel = Nothing
        DiscountPanelBase4.PanelType = 0
        DiscountPanelBase4.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase4.Service = "CALLOUT"
        DiscountPanelBase4.Top = 0
        DiscountPanelBase4.Width = 884
        Me.pnlCallOut.BasePanel = DiscountPanelBase4
        Me.pnlCallOut.Client = Nothing
        Me.pnlCallOut.Discount = ""
        Me.pnlCallOut.DiscountScheme = ""
        Me.pnlCallOut.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCallOut.LineItem = ""
        Me.pnlCallOut.ListPrice = ""
        Me.pnlCallOut.Location = New System.Drawing.Point(0, 80)
        Me.pnlCallOut.Name = "pnlCallOut"
        Me.pnlCallOut.PanelLabel = "Call Out Samples"
        Me.pnlCallOut.PanelService = Nothing
        Me.pnlCallOut.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Sample
        Me.pnlCallOut.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlCallOut.PriceEnabled = True
        Me.pnlCallOut.Scheme = ""
        Me.pnlCallOut.Service = "CALLOUT"
        Me.pnlCallOut.size = New System.Drawing.Size(884, 48)
        Me.pnlCallOut.size = New System.Drawing.Size(884, 48)
        Me.pnlCallOut.TabIndex = 50
        '
        'pnlRoutine
        '
        Me.pnlRoutine.AllowLeave = True
        Me.pnlRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase5.AllowLeave = False
        DiscountPanelBase5.DiscountScheme = ""
        DiscountPanelBase5.Height = 40
        DiscountPanelBase5.Label = "Routine  Samples"
        DiscountPanelBase5.Left = 0
        DiscountPanelBase5.LineItem = ""
        DiscountPanelBase5.Name = Nothing
        DiscountPanelBase5.Panel = Nothing
        DiscountPanelBase5.PanelType = 0
        DiscountPanelBase5.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase5.Service = "ROUTINE"
        DiscountPanelBase5.Top = 0
        DiscountPanelBase5.Width = 884
        Me.pnlRoutine.BasePanel = DiscountPanelBase5
        Me.pnlRoutine.Client = Nothing
        Me.pnlRoutine.Discount = ""
        Me.pnlRoutine.DiscountScheme = ""
        Me.pnlRoutine.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRoutine.LineItem = ""
        Me.pnlRoutine.ListPrice = ""
        Me.pnlRoutine.Location = New System.Drawing.Point(0, 40)
        Me.pnlRoutine.Name = "pnlRoutine"
        Me.pnlRoutine.PanelLabel = "Routine  Samples"
        Me.pnlRoutine.PanelService = Nothing
        Me.pnlRoutine.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Sample
        Me.pnlRoutine.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlRoutine.PriceEnabled = True
        Me.pnlRoutine.Scheme = ""
        Me.pnlRoutine.Service = "ROUTINE"
        Me.pnlRoutine.size = New System.Drawing.Size(884, 40)
        Me.pnlRoutine.size = New System.Drawing.Size(884, 40)
        Me.pnlRoutine.TabIndex = 51
        '
        'pnlAO
        '
        Me.pnlAO.AllowLeave = True
        Me.pnlAO.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase6.AllowLeave = False
        DiscountPanelBase6.DiscountScheme = ""
        DiscountPanelBase6.Height = 40
        DiscountPanelBase6.Label = "Analysis Only"
        DiscountPanelBase6.Left = 0
        DiscountPanelBase6.LineItem = ""
        DiscountPanelBase6.Name = Nothing
        DiscountPanelBase6.Panel = Nothing
        DiscountPanelBase6.PanelType = 0
        DiscountPanelBase6.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase6.Service = "ANALONLY"
        DiscountPanelBase6.Top = 0
        DiscountPanelBase6.Width = 884
        Me.pnlAO.BasePanel = DiscountPanelBase6
        Me.pnlAO.Client = Nothing
        Me.pnlAO.Discount = ""
        Me.pnlAO.DiscountScheme = ""
        Me.pnlAO.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAO.LineItem = ""
        Me.pnlAO.ListPrice = ""
        Me.pnlAO.Location = New System.Drawing.Point(0, 0)
        Me.pnlAO.Name = "pnlAO"
        Me.pnlAO.PanelLabel = "Analysis Only"
        Me.pnlAO.PanelService = Nothing
        Me.pnlAO.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Sample
        Me.pnlAO.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlAO.PriceEnabled = True
        Me.pnlAO.Scheme = ""
        Me.pnlAO.Service = "ANALONLY"
        Me.pnlAO.size = New System.Drawing.Size(884, 40)
        Me.pnlAO.size = New System.Drawing.Size(884, 40)
        Me.pnlAO.TabIndex = 52
        '
        'tabpMinCharge
        '
        Me.tabpMinCharge.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabpMinCharge.Controls.Add(Me.pnlMinCharge)
        Me.tabpMinCharge.Location = New System.Drawing.Point(4, 4)
        Me.tabpMinCharge.Name = "tabpMinCharge"
        Me.tabpMinCharge.Size = New System.Drawing.Size(912, 494)
        Me.tabpMinCharge.TabIndex = 1
        Me.tabpMinCharge.Text = "Minimum Charge"
        Me.tabpMinCharge.Visible = False
        '
        'pnlMinCharge
        '
        Me.pnlMinCharge.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlMinCharge.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMinCharge.Controls.Add(Me.dpCallOutCollChargeOut)
        Me.pnlMinCharge.Controls.Add(Me.dpCallOutCollCharge)
        Me.pnlMinCharge.Controls.Add(Me.dpRoutineCollChargeOut)
        Me.pnlMinCharge.Controls.Add(Me.dpRoutineCollCharge)
        Me.pnlMinCharge.Controls.Add(Me.dpMCCallOutOut)
        Me.pnlMinCharge.Controls.Add(Me.dpMCCallOutIn)
        Me.pnlMinCharge.Controls.Add(Me.pnlOutHoursMC)
        Me.pnlMinCharge.Controls.Add(Me.pnlInHourMinCharge)
        Me.pnlMinCharge.Location = New System.Drawing.Point(8, 8)
        Me.pnlMinCharge.Name = "pnlMinCharge"
        Me.pnlMinCharge.Size = New System.Drawing.Size(896, 507)
        Me.pnlMinCharge.TabIndex = 0
        '
        'dpCallOutCollChargeOut
        '
        Me.dpCallOutCollChargeOut.AllowLeave = True
        Me.dpCallOutCollChargeOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase7.AllowLeave = False
        DiscountPanelBase7.DiscountScheme = ""
        DiscountPanelBase7.Height = 40
        DiscountPanelBase7.Label = "Call Out Attendance Fee (Out Of Hours)"
        DiscountPanelBase7.Left = 0
        DiscountPanelBase7.LineItem = ""
        DiscountPanelBase7.Name = Nothing
        DiscountPanelBase7.Panel = Nothing
        DiscountPanelBase7.PanelType = 0
        DiscountPanelBase7.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase7.Service = "CALLOUTOUT"
        DiscountPanelBase7.Top = 0
        DiscountPanelBase7.Width = 892
        Me.dpCallOutCollChargeOut.BasePanel = DiscountPanelBase7
        Me.dpCallOutCollChargeOut.Client = Nothing
        Me.dpCallOutCollChargeOut.ContextMenu = Me.ctxAddService
        Me.dpCallOutCollChargeOut.Discount = ""
        Me.dpCallOutCollChargeOut.DiscountScheme = ""
        Me.dpCallOutCollChargeOut.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpCallOutCollChargeOut.LineItem = ""
        Me.dpCallOutCollChargeOut.ListPrice = ""
        Me.dpCallOutCollChargeOut.Location = New System.Drawing.Point(0, 280)
        Me.dpCallOutCollChargeOut.Name = "dpCallOutCollChargeOut"
        Me.dpCallOutCollChargeOut.PanelLabel = "Call Out Attendance Fee (Out Of Hours)"
        Me.dpCallOutCollChargeOut.PanelService = Nothing
        Me.dpCallOutCollChargeOut.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.CollectionCharge
        Me.dpCallOutCollChargeOut.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpCallOutCollChargeOut.PriceEnabled = True
        Me.dpCallOutCollChargeOut.Scheme = ""
        Me.dpCallOutCollChargeOut.Service = "CALLOUTOUT"
        Me.dpCallOutCollChargeOut.size = New System.Drawing.Size(892, 40)
        Me.dpCallOutCollChargeOut.size = New System.Drawing.Size(892, 40)
        Me.dpCallOutCollChargeOut.TabIndex = 65
        '
        'ctxAddService
        '
        Me.ctxAddService.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuctxAddService, Me.mnuCTXDiscountDates})
        '
        'mnuctxAddService
        '
        Me.mnuctxAddService.Index = 0
        Me.mnuctxAddService.Text = "Add Service"
        '
        'mnuCTXDiscountDates
        '
        Me.mnuCTXDiscountDates.Index = 1
        Me.mnuCTXDiscountDates.Text = "Set Discount Dates"
        '
        'dpCallOutCollCharge
        '
        Me.dpCallOutCollCharge.AllowLeave = True
        Me.dpCallOutCollCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase8.AllowLeave = False
        DiscountPanelBase8.DiscountScheme = ""
        DiscountPanelBase8.Height = 40
        DiscountPanelBase8.Label = "Call Out Attendance Fee"
        DiscountPanelBase8.Left = 0
        DiscountPanelBase8.LineItem = ""
        DiscountPanelBase8.Name = Nothing
        DiscountPanelBase8.Panel = Nothing
        DiscountPanelBase8.PanelType = 0
        DiscountPanelBase8.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase8.Service = "CALLOUT"
        DiscountPanelBase8.Top = 0
        DiscountPanelBase8.Width = 892
        Me.dpCallOutCollCharge.BasePanel = DiscountPanelBase8
        Me.dpCallOutCollCharge.Client = Nothing
        Me.dpCallOutCollCharge.ContextMenu = Me.ctxAddService
        Me.dpCallOutCollCharge.Discount = ""
        Me.dpCallOutCollCharge.DiscountScheme = ""
        Me.dpCallOutCollCharge.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpCallOutCollCharge.LineItem = ""
        Me.dpCallOutCollCharge.ListPrice = ""
        Me.dpCallOutCollCharge.Location = New System.Drawing.Point(0, 240)
        Me.dpCallOutCollCharge.Name = "dpCallOutCollCharge"
        Me.dpCallOutCollCharge.PanelLabel = "Call Out Attendance Fee"
        Me.dpCallOutCollCharge.PanelService = Nothing
        Me.dpCallOutCollCharge.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.CollectionCharge
        Me.dpCallOutCollCharge.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpCallOutCollCharge.PriceEnabled = True
        Me.dpCallOutCollCharge.Scheme = ""
        Me.dpCallOutCollCharge.Service = "CALLOUT"
        Me.dpCallOutCollCharge.size = New System.Drawing.Size(892, 40)
        Me.dpCallOutCollCharge.size = New System.Drawing.Size(892, 40)
        Me.dpCallOutCollCharge.TabIndex = 63
        '
        'dpRoutineCollChargeOut
        '
        Me.dpRoutineCollChargeOut.AllowLeave = True
        Me.dpRoutineCollChargeOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase9.AllowLeave = False
        DiscountPanelBase9.DiscountScheme = ""
        DiscountPanelBase9.Height = 40
        DiscountPanelBase9.Label = "Collection Attendance Fee (Out Of Hours)"
        DiscountPanelBase9.Left = 0
        DiscountPanelBase9.LineItem = ""
        DiscountPanelBase9.Name = Nothing
        DiscountPanelBase9.Panel = Nothing
        DiscountPanelBase9.PanelType = 0
        DiscountPanelBase9.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase9.Service = "ROUTINEOUT"
        DiscountPanelBase9.Top = 0
        DiscountPanelBase9.Width = 892
        Me.dpRoutineCollChargeOut.BasePanel = DiscountPanelBase9
        Me.dpRoutineCollChargeOut.Client = Nothing
        Me.dpRoutineCollChargeOut.ContextMenu = Me.ctxAddService
        Me.dpRoutineCollChargeOut.Discount = ""
        Me.dpRoutineCollChargeOut.DiscountScheme = ""
        Me.dpRoutineCollChargeOut.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpRoutineCollChargeOut.LineItem = ""
        Me.dpRoutineCollChargeOut.ListPrice = ""
        Me.dpRoutineCollChargeOut.Location = New System.Drawing.Point(0, 200)
        Me.dpRoutineCollChargeOut.Name = "dpRoutineCollChargeOut"
        Me.dpRoutineCollChargeOut.PanelLabel = "Collection Attendance Fee (Out Of Hours)"
        Me.dpRoutineCollChargeOut.PanelService = Nothing
        Me.dpRoutineCollChargeOut.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.CollectionCharge
        Me.dpRoutineCollChargeOut.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpRoutineCollChargeOut.PriceEnabled = True
        Me.dpRoutineCollChargeOut.Scheme = ""
        Me.dpRoutineCollChargeOut.Service = "ROUTINEOUT"
        Me.dpRoutineCollChargeOut.size = New System.Drawing.Size(892, 40)
        Me.dpRoutineCollChargeOut.size = New System.Drawing.Size(892, 40)
        Me.dpRoutineCollChargeOut.TabIndex = 64
        '
        'dpRoutineCollCharge
        '
        Me.dpRoutineCollCharge.AllowLeave = True
        Me.dpRoutineCollCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase10.AllowLeave = False
        DiscountPanelBase10.DiscountScheme = ""
        DiscountPanelBase10.Height = 40
        DiscountPanelBase10.Label = "Collection Attendance Fee"
        DiscountPanelBase10.Left = 0
        DiscountPanelBase10.LineItem = ""
        DiscountPanelBase10.Name = Nothing
        DiscountPanelBase10.Panel = Nothing
        DiscountPanelBase10.PanelType = 0
        DiscountPanelBase10.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase10.Service = "ROUTINE"
        DiscountPanelBase10.Top = 0
        DiscountPanelBase10.Width = 892
        Me.dpRoutineCollCharge.BasePanel = DiscountPanelBase10
        Me.dpRoutineCollCharge.Client = Nothing
        Me.dpRoutineCollCharge.ContextMenu = Me.ctxAddService
        Me.dpRoutineCollCharge.Discount = ""
        Me.dpRoutineCollCharge.DiscountScheme = ""
        Me.dpRoutineCollCharge.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpRoutineCollCharge.LineItem = ""
        Me.dpRoutineCollCharge.ListPrice = ""
        Me.dpRoutineCollCharge.Location = New System.Drawing.Point(0, 160)
        Me.dpRoutineCollCharge.Name = "dpRoutineCollCharge"
        Me.dpRoutineCollCharge.PanelLabel = "Collection Attendance Fee"
        Me.dpRoutineCollCharge.PanelService = Nothing
        Me.dpRoutineCollCharge.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.CollectionCharge
        Me.dpRoutineCollCharge.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpRoutineCollCharge.PriceEnabled = True
        Me.dpRoutineCollCharge.Scheme = ""
        Me.dpRoutineCollCharge.Service = "ROUTINE"
        Me.dpRoutineCollCharge.size = New System.Drawing.Size(892, 40)
        Me.dpRoutineCollCharge.size = New System.Drawing.Size(892, 40)
        Me.dpRoutineCollCharge.TabIndex = 62
        '
        'dpMCCallOutOut
        '
        Me.dpMCCallOutOut.AllowLeave = True
        Me.dpMCCallOutOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase11.AllowLeave = False
        DiscountPanelBase11.DiscountScheme = ""
        DiscountPanelBase11.Height = 40
        DiscountPanelBase11.Label = "Call Out Minimum Charge (out hours)"
        DiscountPanelBase11.Left = 0
        DiscountPanelBase11.LineItem = "APCOLC007"
        DiscountPanelBase11.Name = Nothing
        DiscountPanelBase11.Panel = Nothing
        DiscountPanelBase11.PanelType = 0
        DiscountPanelBase11.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase11.Service = "CALLOUTOUT"
        DiscountPanelBase11.Top = 0
        DiscountPanelBase11.Width = 892
        Me.dpMCCallOutOut.BasePanel = DiscountPanelBase11
        Me.dpMCCallOutOut.Client = Nothing
        Me.dpMCCallOutOut.ContextMenu = Me.ctxAddService
        Me.dpMCCallOutOut.Discount = ""
        Me.dpMCCallOutOut.DiscountScheme = ""
        Me.dpMCCallOutOut.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpMCCallOutOut.LineItem = "APCOLC007"
        Me.dpMCCallOutOut.ListPrice = ""
        Me.dpMCCallOutOut.Location = New System.Drawing.Point(0, 120)
        Me.dpMCCallOutOut.Name = "dpMCCallOutOut"
        Me.dpMCCallOutOut.PanelLabel = "Call Out Minimum Charge (out hours)"
        Me.dpMCCallOutOut.PanelService = Nothing
        Me.dpMCCallOutOut.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.MinimumCharge
        Me.dpMCCallOutOut.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpMCCallOutOut.PriceEnabled = True
        Me.dpMCCallOutOut.Scheme = ""
        Me.dpMCCallOutOut.Service = "CALLOUTOUT"
        Me.dpMCCallOutOut.size = New System.Drawing.Size(892, 40)
        Me.dpMCCallOutOut.size = New System.Drawing.Size(892, 40)
        Me.dpMCCallOutOut.TabIndex = 67
        '
        'dpMCCallOutIn
        '
        Me.dpMCCallOutIn.AllowLeave = True
        Me.dpMCCallOutIn.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase12.AllowLeave = False
        DiscountPanelBase12.DiscountScheme = ""
        DiscountPanelBase12.Height = 40
        DiscountPanelBase12.Label = "Call Out minimum charge (in hours)"
        DiscountPanelBase12.Left = 0
        DiscountPanelBase12.LineItem = "APCOLC006"
        DiscountPanelBase12.Name = Nothing
        DiscountPanelBase12.Panel = Nothing
        DiscountPanelBase12.PanelType = 0
        DiscountPanelBase12.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase12.Service = "CALLOUT"
        DiscountPanelBase12.Top = 0
        DiscountPanelBase12.Width = 892
        Me.dpMCCallOutIn.BasePanel = DiscountPanelBase12
        Me.dpMCCallOutIn.Client = Nothing
        Me.dpMCCallOutIn.ContextMenu = Me.ctxAddService
        Me.dpMCCallOutIn.Discount = ""
        Me.dpMCCallOutIn.DiscountScheme = ""
        Me.dpMCCallOutIn.Dock = System.Windows.Forms.DockStyle.Top
        Me.dpMCCallOutIn.LineItem = "APCOLC006"
        Me.dpMCCallOutIn.ListPrice = ""
        Me.dpMCCallOutIn.Location = New System.Drawing.Point(0, 80)
        Me.dpMCCallOutIn.Name = "dpMCCallOutIn"
        Me.dpMCCallOutIn.PanelLabel = "Call Out minimum charge (in hours)"
        Me.dpMCCallOutIn.PanelService = Nothing
        Me.dpMCCallOutIn.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.MinimumCharge
        Me.dpMCCallOutIn.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpMCCallOutIn.PriceEnabled = True
        Me.dpMCCallOutIn.Scheme = ""
        Me.dpMCCallOutIn.Service = "CALLOUT"
        Me.dpMCCallOutIn.size = New System.Drawing.Size(892, 40)
        Me.dpMCCallOutIn.size = New System.Drawing.Size(892, 40)
        Me.dpMCCallOutIn.TabIndex = 66
        '
        'pnlOutHoursMC
        '
        Me.pnlOutHoursMC.AllowLeave = True
        Me.pnlOutHoursMC.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase13.AllowLeave = False
        DiscountPanelBase13.DiscountScheme = ""
        DiscountPanelBase13.Height = 40
        DiscountPanelBase13.Label = "Out of Hours minimum charge"
        DiscountPanelBase13.Left = 0
        DiscountPanelBase13.LineItem = ""
        DiscountPanelBase13.Name = Nothing
        DiscountPanelBase13.Panel = Nothing
        DiscountPanelBase13.PanelType = 0
        DiscountPanelBase13.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase13.Service = "ROUTINEOUT"
        DiscountPanelBase13.Top = 0
        DiscountPanelBase13.Width = 892
        Me.pnlOutHoursMC.BasePanel = DiscountPanelBase13
        Me.pnlOutHoursMC.Client = Nothing
        Me.pnlOutHoursMC.ContextMenu = Me.ctxAddService
        Me.pnlOutHoursMC.Discount = ""
        Me.pnlOutHoursMC.DiscountScheme = ""
        Me.pnlOutHoursMC.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOutHoursMC.LineItem = ""
        Me.pnlOutHoursMC.ListPrice = ""
        Me.pnlOutHoursMC.Location = New System.Drawing.Point(0, 40)
        Me.pnlOutHoursMC.Name = "pnlOutHoursMC"
        Me.pnlOutHoursMC.PanelLabel = "Out of Hours minimum charge"
        Me.pnlOutHoursMC.PanelService = Nothing
        Me.pnlOutHoursMC.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.MinimumCharge
        Me.pnlOutHoursMC.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlOutHoursMC.PriceEnabled = True
        Me.pnlOutHoursMC.Scheme = ""
        Me.pnlOutHoursMC.Service = "ROUTINEOUT"
        Me.pnlOutHoursMC.size = New System.Drawing.Size(892, 40)
        Me.pnlOutHoursMC.size = New System.Drawing.Size(892, 40)
        Me.pnlOutHoursMC.TabIndex = 59
        '
        'pnlInHourMinCharge
        '
        Me.pnlInHourMinCharge.AllowLeave = True
        Me.pnlInHourMinCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase14.AllowLeave = False
        DiscountPanelBase14.DiscountScheme = ""
        DiscountPanelBase14.Height = 40
        DiscountPanelBase14.Label = "In hours minimum charge"
        DiscountPanelBase14.Left = 0
        DiscountPanelBase14.LineItem = ""
        DiscountPanelBase14.Name = Nothing
        DiscountPanelBase14.Panel = Nothing
        DiscountPanelBase14.PanelType = 0
        DiscountPanelBase14.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase14.Service = "ROUTINE"
        DiscountPanelBase14.Top = 0
        DiscountPanelBase14.Width = 892
        Me.pnlInHourMinCharge.BasePanel = DiscountPanelBase14
        Me.pnlInHourMinCharge.Client = Nothing
        Me.pnlInHourMinCharge.ContextMenu = Me.ctxAddService
        Me.pnlInHourMinCharge.Discount = ""
        Me.pnlInHourMinCharge.DiscountScheme = ""
        Me.pnlInHourMinCharge.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInHourMinCharge.LineItem = ""
        Me.pnlInHourMinCharge.ListPrice = ""
        Me.pnlInHourMinCharge.Location = New System.Drawing.Point(0, 0)
        Me.pnlInHourMinCharge.Name = "pnlInHourMinCharge"
        Me.pnlInHourMinCharge.PanelLabel = "In hours minimum charge"
        Me.pnlInHourMinCharge.PanelService = Nothing
        Me.pnlInHourMinCharge.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.MinimumCharge
        Me.pnlInHourMinCharge.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlInHourMinCharge.PriceEnabled = True
        Me.pnlInHourMinCharge.Scheme = ""
        Me.pnlInHourMinCharge.Service = "ROUTINE"
        Me.pnlInHourMinCharge.size = New System.Drawing.Size(892, 40)
        Me.pnlInHourMinCharge.size = New System.Drawing.Size(892, 40)
        Me.pnlInHourMinCharge.TabIndex = 53
        '
        'tabpAddTime
        '
        Me.tabpAddTime.Controls.Add(Me.pnlAddTime)
        Me.tabpAddTime.Location = New System.Drawing.Point(4, 4)
        Me.tabpAddTime.Name = "tabpAddTime"
        Me.tabpAddTime.Size = New System.Drawing.Size(912, 494)
        Me.tabpAddTime.TabIndex = 2
        Me.tabpAddTime.Text = "Additional Time"
        Me.tabpAddTime.Visible = False
        '
        'pnlAddTime
        '
        Me.pnlAddTime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlAddTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlAddTime.Controls.Add(Me.pnlOutHoursAddTimeCO)
        Me.pnlAddTime.Controls.Add(Me.pnlINHoursAddTimeCO)
        Me.pnlAddTime.Controls.Add(Me.pnlOutHoursAddTimeRoutine)
        Me.pnlAddTime.Controls.Add(Me.pnlINHoursAddTimeRoutine)
        Me.pnlAddTime.Controls.Add(Me.pnlADHeader)
        Me.pnlAddTime.Controls.Add(Me.pnlTimeAllowance)
        Me.pnlAddTime.Location = New System.Drawing.Point(16, 8)
        Me.pnlAddTime.Name = "pnlAddTime"
        Me.pnlAddTime.Size = New System.Drawing.Size(896, 320)
        Me.pnlAddTime.TabIndex = 0
        '
        'pnlOutHoursAddTimeCO
        '
        Me.pnlOutHoursAddTimeCO.AllowLeave = True
        Me.pnlOutHoursAddTimeCO.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase15.AllowLeave = False
        DiscountPanelBase15.DiscountScheme = ""
        DiscountPanelBase15.Height = 40
        DiscountPanelBase15.Label = "Out hours Add Time (Call Out)"
        DiscountPanelBase15.Left = 0
        DiscountPanelBase15.LineItem = ""
        DiscountPanelBase15.Name = Nothing
        DiscountPanelBase15.Panel = Nothing
        DiscountPanelBase15.PanelType = 0
        DiscountPanelBase15.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase15.Service = "CALLOUT"
        DiscountPanelBase15.Top = 0
        DiscountPanelBase15.Width = 894
        Me.pnlOutHoursAddTimeCO.BasePanel = DiscountPanelBase15
        Me.pnlOutHoursAddTimeCO.Client = Nothing
        Me.pnlOutHoursAddTimeCO.ContextMenu = Me.ctxAddService
        Me.pnlOutHoursAddTimeCO.Discount = ""
        Me.pnlOutHoursAddTimeCO.DiscountScheme = ""
        Me.pnlOutHoursAddTimeCO.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOutHoursAddTimeCO.LineItem = ""
        Me.pnlOutHoursAddTimeCO.ListPrice = ""
        Me.pnlOutHoursAddTimeCO.Location = New System.Drawing.Point(0, 248)
        Me.pnlOutHoursAddTimeCO.Name = "pnlOutHoursAddTimeCO"
        Me.pnlOutHoursAddTimeCO.PanelLabel = "Out hours Add Time (Call Out)"
        Me.pnlOutHoursAddTimeCO.PanelService = Nothing
        Me.pnlOutHoursAddTimeCO.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.AdditionalTime
        Me.pnlOutHoursAddTimeCO.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlOutHoursAddTimeCO.PriceEnabled = True
        Me.pnlOutHoursAddTimeCO.Scheme = ""
        Me.pnlOutHoursAddTimeCO.Service = "CALLOUT"
        Me.pnlOutHoursAddTimeCO.size = New System.Drawing.Size(894, 40)
        Me.pnlOutHoursAddTimeCO.size = New System.Drawing.Size(894, 40)
        Me.pnlOutHoursAddTimeCO.TabIndex = 59
        '
        'pnlINHoursAddTimeCO
        '
        Me.pnlINHoursAddTimeCO.AllowLeave = True
        Me.pnlINHoursAddTimeCO.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase16.AllowLeave = False
        DiscountPanelBase16.DiscountScheme = ""
        DiscountPanelBase16.Height = 40
        DiscountPanelBase16.Label = "In hours Add Time (Call Out)"
        DiscountPanelBase16.Left = 0
        DiscountPanelBase16.LineItem = ""
        DiscountPanelBase16.Name = Nothing
        DiscountPanelBase16.Panel = Nothing
        DiscountPanelBase16.PanelType = 0
        DiscountPanelBase16.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase16.Service = "CALLOUT"
        DiscountPanelBase16.Top = 0
        DiscountPanelBase16.Width = 894
        Me.pnlINHoursAddTimeCO.BasePanel = DiscountPanelBase16
        Me.pnlINHoursAddTimeCO.Client = Nothing
        Me.pnlINHoursAddTimeCO.ContextMenu = Me.ctxAddService
        Me.pnlINHoursAddTimeCO.Discount = ""
        Me.pnlINHoursAddTimeCO.DiscountScheme = ""
        Me.pnlINHoursAddTimeCO.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlINHoursAddTimeCO.LineItem = ""
        Me.pnlINHoursAddTimeCO.ListPrice = ""
        Me.pnlINHoursAddTimeCO.Location = New System.Drawing.Point(0, 208)
        Me.pnlINHoursAddTimeCO.Name = "pnlINHoursAddTimeCO"
        Me.pnlINHoursAddTimeCO.PanelLabel = "In hours Add Time (Call Out)"
        Me.pnlINHoursAddTimeCO.PanelService = Nothing
        Me.pnlINHoursAddTimeCO.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.AdditionalTime
        Me.pnlINHoursAddTimeCO.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlINHoursAddTimeCO.PriceEnabled = True
        Me.pnlINHoursAddTimeCO.Scheme = ""
        Me.pnlINHoursAddTimeCO.Service = "CALLOUT"
        Me.pnlINHoursAddTimeCO.size = New System.Drawing.Size(894, 40)
        Me.pnlINHoursAddTimeCO.size = New System.Drawing.Size(894, 40)
        Me.pnlINHoursAddTimeCO.TabIndex = 58
        '
        'pnlOutHoursAddTimeRoutine
        '
        Me.pnlOutHoursAddTimeRoutine.AllowLeave = True
        Me.pnlOutHoursAddTimeRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase17.AllowLeave = False
        DiscountPanelBase17.DiscountScheme = ""
        DiscountPanelBase17.Height = 40
        DiscountPanelBase17.Label = "Out hours Add Time (Routine)"
        DiscountPanelBase17.Left = 0
        DiscountPanelBase17.LineItem = ""
        DiscountPanelBase17.Name = Nothing
        DiscountPanelBase17.Panel = Nothing
        DiscountPanelBase17.PanelType = 0
        DiscountPanelBase17.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase17.Service = "ROUTINE"
        DiscountPanelBase17.Top = 0
        DiscountPanelBase17.Width = 894
        Me.pnlOutHoursAddTimeRoutine.BasePanel = DiscountPanelBase17
        Me.pnlOutHoursAddTimeRoutine.Client = Nothing
        Me.pnlOutHoursAddTimeRoutine.ContextMenu = Me.ctxAddService
        Me.pnlOutHoursAddTimeRoutine.Discount = ""
        Me.pnlOutHoursAddTimeRoutine.DiscountScheme = ""
        Me.pnlOutHoursAddTimeRoutine.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOutHoursAddTimeRoutine.LineItem = ""
        Me.pnlOutHoursAddTimeRoutine.ListPrice = ""
        Me.pnlOutHoursAddTimeRoutine.Location = New System.Drawing.Point(0, 168)
        Me.pnlOutHoursAddTimeRoutine.Name = "pnlOutHoursAddTimeRoutine"
        Me.pnlOutHoursAddTimeRoutine.PanelLabel = "Out hours Add Time (Routine)"
        Me.pnlOutHoursAddTimeRoutine.PanelService = Nothing
        Me.pnlOutHoursAddTimeRoutine.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.AdditionalTime
        Me.pnlOutHoursAddTimeRoutine.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlOutHoursAddTimeRoutine.PriceEnabled = True
        Me.pnlOutHoursAddTimeRoutine.Scheme = ""
        Me.pnlOutHoursAddTimeRoutine.Service = "ROUTINE"
        Me.pnlOutHoursAddTimeRoutine.size = New System.Drawing.Size(894, 40)
        Me.pnlOutHoursAddTimeRoutine.size = New System.Drawing.Size(894, 40)
        Me.pnlOutHoursAddTimeRoutine.TabIndex = 57
        '
        'pnlINHoursAddTimeRoutine
        '
        Me.pnlINHoursAddTimeRoutine.AllowLeave = True
        Me.pnlINHoursAddTimeRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase18.AllowLeave = False
        DiscountPanelBase18.DiscountScheme = ""
        DiscountPanelBase18.Height = 40
        DiscountPanelBase18.Label = "In hours Add Time (Routine)"
        DiscountPanelBase18.Left = 0
        DiscountPanelBase18.LineItem = ""
        DiscountPanelBase18.Name = Nothing
        DiscountPanelBase18.Panel = Nothing
        DiscountPanelBase18.PanelType = 0
        DiscountPanelBase18.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase18.Service = "ROUTINE"
        DiscountPanelBase18.Top = 0
        DiscountPanelBase18.Width = 894
        Me.pnlINHoursAddTimeRoutine.BasePanel = DiscountPanelBase18
        Me.pnlINHoursAddTimeRoutine.Client = Nothing
        Me.pnlINHoursAddTimeRoutine.ContextMenu = Me.ctxAddService
        Me.pnlINHoursAddTimeRoutine.Discount = ""
        Me.pnlINHoursAddTimeRoutine.DiscountScheme = ""
        Me.pnlINHoursAddTimeRoutine.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlINHoursAddTimeRoutine.LineItem = ""
        Me.pnlINHoursAddTimeRoutine.ListPrice = ""
        Me.pnlINHoursAddTimeRoutine.Location = New System.Drawing.Point(0, 128)
        Me.pnlINHoursAddTimeRoutine.Name = "pnlINHoursAddTimeRoutine"
        Me.pnlINHoursAddTimeRoutine.PanelLabel = "In hours Add Time (Routine)"
        Me.pnlINHoursAddTimeRoutine.PanelService = Nothing
        Me.pnlINHoursAddTimeRoutine.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.AdditionalTime
        Me.pnlINHoursAddTimeRoutine.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.pnlINHoursAddTimeRoutine.PriceEnabled = True
        Me.pnlINHoursAddTimeRoutine.Scheme = ""
        Me.pnlINHoursAddTimeRoutine.Service = "ROUTINE"
        Me.pnlINHoursAddTimeRoutine.size = New System.Drawing.Size(894, 40)
        Me.pnlINHoursAddTimeRoutine.size = New System.Drawing.Size(894, 40)
        Me.pnlINHoursAddTimeRoutine.TabIndex = 56
        '
        'pnlADHeader
        '
        Me.pnlADHeader.Controls.Add(Me.Label63)
        Me.pnlADHeader.Controls.Add(Me.lbladpCurr)
        Me.pnlADHeader.Controls.Add(Me.lbllPADCurr)
        Me.pnlADHeader.Controls.Add(Me.Label64)
        Me.pnlADHeader.Controls.Add(Me.Label65)
        Me.pnlADHeader.Controls.Add(Me.Label66)
        Me.pnlADHeader.Controls.Add(Me.Label67)
        Me.pnlADHeader.Controls.Add(Me.Label68)
        Me.pnlADHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlADHeader.Location = New System.Drawing.Point(0, 80)
        Me.pnlADHeader.Name = "pnlADHeader"
        Me.pnlADHeader.Size = New System.Drawing.Size(894, 48)
        Me.pnlADHeader.TabIndex = 1
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Location = New System.Drawing.Point(520, 12)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(45, 13)
        Me.Label63.TabIndex = 63
        Me.Label63.Text = "Price/hr"
        '
        'lbladpCurr
        '
        Me.lbladpCurr.AutoSize = True
        Me.lbladpCurr.Location = New System.Drawing.Point(492, 12)
        Me.lbladpCurr.Name = "lbladpCurr"
        Me.lbladpCurr.Size = New System.Drawing.Size(19, 13)
        Me.lbladpCurr.TabIndex = 62
        Me.lbladpCurr.Text = "(£)"
        '
        'lbllPADCurr
        '
        Me.lbllPADCurr.AutoSize = True
        Me.lbllPADCurr.Location = New System.Drawing.Point(302, 12)
        Me.lbllPADCurr.Name = "lbllPADCurr"
        Me.lbllPADCurr.Size = New System.Drawing.Size(19, 13)
        Me.lbllPADCurr.TabIndex = 61
        Me.lbllPADCurr.Text = "(£)"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(582, 12)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(46, 13)
        Me.Label64.TabIndex = 60
        Me.Label64.Text = "Scheme"
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Location = New System.Drawing.Point(454, 12)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(31, 13)
        Me.Label65.TabIndex = 59
        Me.Label65.Text = "Price"
        '
        'Label66
        '
        Me.Label66.Location = New System.Drawing.Point(360, 12)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(100, 23)
        Me.Label66.TabIndex = 58
        Me.Label66.Text = "Discount %"
        '
        'Label67
        '
        Me.Label67.AutoSize = True
        Me.Label67.Location = New System.Drawing.Point(216, 12)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(87, 13)
        Me.Label67.TabIndex = 57
        Me.Label67.Text = "List Price per min"
        '
        'Label68
        '
        Me.Label68.Location = New System.Drawing.Point(80, 12)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(100, 23)
        Me.Label68.TabIndex = 56
        Me.Label68.Text = "Min Charge Type "
        '
        'pnlTimeAllowance
        '
        Me.pnlTimeAllowance.Controls.Add(Me.pnlTimeAllowHead)
        Me.pnlTimeAllowance.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTimeAllowance.Location = New System.Drawing.Point(0, 0)
        Me.pnlTimeAllowance.Name = "pnlTimeAllowance"
        Me.pnlTimeAllowance.Size = New System.Drawing.Size(894, 80)
        Me.pnlTimeAllowance.TabIndex = 60
        '
        'pnlTimeAllowHead
        '
        Me.pnlTimeAllowHead.Controls.Add(Me.Label70)
        Me.pnlTimeAllowHead.Controls.Add(Me.lblTimeAllowance)
        Me.pnlTimeAllowHead.Controls.Add(Me.txtTimeAllowance)
        Me.pnlTimeAllowHead.Controls.Add(Me.Label69)
        Me.pnlTimeAllowHead.Controls.Add(Me.ckTimeAllowance)
        Me.pnlTimeAllowHead.Location = New System.Drawing.Point(0, 8)
        Me.pnlTimeAllowHead.Name = "pnlTimeAllowHead"
        Me.pnlTimeAllowHead.Size = New System.Drawing.Size(696, 136)
        Me.pnlTimeAllowHead.TabIndex = 62
        '
        'Label70
        '
        Me.Label70.BackColor = System.Drawing.SystemColors.Info
        Me.Label70.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label70.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label70.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label70.Location = New System.Drawing.Point(288, 56)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(352, 23)
        Me.Label70.TabIndex = 61
        Me.Label70.Text = "Time allowance can be removed by removing all text from box "
        '
        'lblTimeAllowance
        '
        Me.lblTimeAllowance.Location = New System.Drawing.Point(344, 16)
        Me.lblTimeAllowance.Name = "lblTimeAllowance"
        Me.lblTimeAllowance.Size = New System.Drawing.Size(344, 23)
        Me.lblTimeAllowance.TabIndex = 60
        '
        'txtTimeAllowance
        '
        Me.txtTimeAllowance.Location = New System.Drawing.Point(280, 16)
        Me.txtTimeAllowance.Name = "txtTimeAllowance"
        Me.txtTimeAllowance.Size = New System.Drawing.Size(40, 20)
        Me.txtTimeAllowance.TabIndex = 59
        '
        'Label69
        '
        Me.Label69.AutoSize = True
        Me.Label69.Location = New System.Drawing.Point(176, 16)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(100, 13)
        Me.Label69.TabIndex = 58
        Me.Label69.Text = "Time Allowance (hr)"
        '
        'ckTimeAllowance
        '
        Me.ckTimeAllowance.Location = New System.Drawing.Point(32, 8)
        Me.ckTimeAllowance.Name = "ckTimeAllowance"
        Me.ckTimeAllowance.Size = New System.Drawing.Size(160, 32)
        Me.ckTimeAllowance.TabIndex = 0
        Me.ckTimeAllowance.Text = "Charge Customer Collection Time"
        '
        'tabpTests
        '
        Me.tabpTests.Controls.Add(Me.pnlTests)
        Me.tabpTests.Location = New System.Drawing.Point(4, 4)
        Me.tabpTests.Name = "tabpTests"
        Me.tabpTests.Size = New System.Drawing.Size(912, 494)
        Me.tabpTests.TabIndex = 3
        Me.tabpTests.Text = "Tests"
        Me.tabpTests.Visible = False
        '
        'pnlTests
        '
        Me.pnlTests.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTests.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTests.Controls.Add(Me.Label74)
        Me.pnlTests.Controls.Add(Me.lvTestConfirm)
        Me.pnlTests.Controls.Add(Me.Label62)
        Me.pnlTests.Controls.Add(Me.ckChargeAdditionalTests)
        Me.pnlTests.Controls.Add(Me.cbEditor)
        Me.pnlTests.Controls.Add(Me.lvTests)
        Me.pnlTests.Location = New System.Drawing.Point(16, 8)
        Me.pnlTests.Name = "pnlTests"
        Me.pnlTests.Size = New System.Drawing.Size(888, 480)
        Me.pnlTests.TabIndex = 0
        '
        'Label74
        '
        Me.Label74.Location = New System.Drawing.Point(32, 312)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(224, 16)
        Me.Label74.TabIndex = 6
        Me.Label74.Text = "Charges for Confirmation Tests"
        '
        'lvTestConfirm
        '
        Me.lvTestConfirm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTestConfirm.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader16, Me.ColumnHeader17, Me.ColumnHeader18, Me.ColumnHeader19, Me.ColumnHeader20, Me.ColumnHeader21, Me.ColumnHeader22})
        Me.lvTestConfirm.FullRowSelect = True
        Me.lvTestConfirm.GridLines = True
        Me.lvTestConfirm.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvTestConfirm.Location = New System.Drawing.Point(24, 344)
        Me.lvTestConfirm.Name = "lvTestConfirm"
        Me.lvTestConfirm.Size = New System.Drawing.Size(864, 152)
        Me.lvTestConfirm.TabIndex = 5
        Me.lvTestConfirm.UseCompatibleStateImageBehavior = False
        Me.lvTestConfirm.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "Description/Test"
        Me.ColumnHeader16.Width = 150
        '
        'ColumnHeader17
        '
        Me.ColumnHeader17.Text = ""
        Me.ColumnHeader17.Width = 40
        '
        'ColumnHeader18
        '
        Me.ColumnHeader18.Text = "List Price"
        Me.ColumnHeader18.Width = 100
        '
        'ColumnHeader19
        '
        Me.ColumnHeader19.Text = ""
        '
        'ColumnHeader20
        '
        Me.ColumnHeader20.Text = "Discount (%)"
        Me.ColumnHeader20.Width = 100
        '
        'ColumnHeader21
        '
        Me.ColumnHeader21.Text = ""
        '
        'ColumnHeader22
        '
        Me.ColumnHeader22.Text = "Price"
        '
        'Label62
        '
        Me.Label62.Location = New System.Drawing.Point(40, 32)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(224, 16)
        Me.Label62.TabIndex = 4
        Me.Label62.Text = "Charges for Screening Tests"
        '
        'ckChargeAdditionalTests
        '
        Me.ckChargeAdditionalTests.Location = New System.Drawing.Point(32, 0)
        Me.ckChargeAdditionalTests.Name = "ckChargeAdditionalTests"
        Me.ckChargeAdditionalTests.Size = New System.Drawing.Size(232, 32)
        Me.ckChargeAdditionalTests.TabIndex = 3
        Me.ckChargeAdditionalTests.Text = "Charge Customer Additional Tests"
        '
        'cbEditor
        '
        Me.cbEditor.Items.AddRange(New Object() {"0", "100"})
        Me.cbEditor.Location = New System.Drawing.Point(480, 8)
        Me.cbEditor.Name = "cbEditor"
        Me.cbEditor.Size = New System.Drawing.Size(121, 21)
        Me.cbEditor.TabIndex = 2
        Me.cbEditor.Visible = False
        '
        'lvTests
        '
        Me.lvTests.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTests.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chTest, Me.chBar1, Me.chListPrice, Me.chbar2, Me.chDiscount, Me.chBar3, Me.chPrice})
        Me.lvTests.FullRowSelect = True
        Me.lvTests.GridLines = True
        Me.lvTests.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvTests.Location = New System.Drawing.Point(16, 64)
        Me.lvTests.Name = "lvTests"
        Me.lvTests.Size = New System.Drawing.Size(864, 232)
        Me.lvTests.TabIndex = 0
        Me.lvTests.UseCompatibleStateImageBehavior = False
        Me.lvTests.View = System.Windows.Forms.View.Details
        '
        'chTest
        '
        Me.chTest.Text = "Description/Test"
        Me.chTest.Width = 150
        '
        'chBar1
        '
        Me.chBar1.Text = ""
        Me.chBar1.Width = 40
        '
        'chListPrice
        '
        Me.chListPrice.Text = "List Price"
        Me.chListPrice.Width = 100
        '
        'chbar2
        '
        Me.chbar2.Text = ""
        '
        'chDiscount
        '
        Me.chDiscount.Text = "Discount (%)"
        Me.chDiscount.Width = 100
        '
        'chBar3
        '
        Me.chBar3.Text = ""
        '
        'chPrice
        '
        Me.chPrice.Text = "Price"
        '
        'tabpKits
        '
        Me.tabpKits.Controls.Add(Me.lvExEQP)
        Me.tabpKits.Controls.Add(Me.txtKitEdit)
        Me.tabpKits.Controls.Add(Me.Label9)
        Me.tabpKits.Controls.Add(Me.lvExKits)
        Me.tabpKits.Location = New System.Drawing.Point(4, 4)
        Me.tabpKits.Name = "tabpKits"
        Me.tabpKits.Size = New System.Drawing.Size(912, 494)
        Me.tabpKits.TabIndex = 4
        Me.tabpKits.Text = "Kits"
        '
        'lvExEQP
        '
        Me.lvExEQP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvExEQP.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.ColumnHeader13, Me.ColumnHeader14, Me.ColumnHeader15, Me.chAvgPrice2})
        Me.lvExEQP.FullRowSelect = True
        Me.lvExEQP.GridLines = True
        Me.lvExEQP.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvExEQP.Location = New System.Drawing.Point(24, 288)
        Me.lvExEQP.Name = "lvExEQP"
        Me.lvExEQP.Size = New System.Drawing.Size(864, 208)
        Me.lvExEQP.TabIndex = 4
        Me.lvExEQP.UseCompatibleStateImageBehavior = False
        Me.lvExEQP.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Description/Kit"
        Me.ColumnHeader4.Width = 300
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = ""
        Me.ColumnHeader8.Width = 40
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "List Price"
        Me.ColumnHeader9.Width = 100
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = ""
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Discount (%)"
        Me.ColumnHeader11.Width = 100
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = ""
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Price"
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "#Invoices"
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "#Sold"
        '
        'chAvgPrice2
        '
        Me.chAvgPrice2.Text = "Avg Price"
        '
        'txtKitEdit
        '
        Me.txtKitEdit.Location = New System.Drawing.Point(336, 8)
        Me.txtKitEdit.Name = "txtKitEdit"
        Me.txtKitEdit.Size = New System.Drawing.Size(100, 20)
        Me.txtKitEdit.TabIndex = 3
        Me.txtKitEdit.Visible = False
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(24, 248)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(576, 64)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Kit pricing for information only, no discount schemes yet in existence, pricing i" & _
            "s handled by Manual setting of Invoice Price."
        '
        'lvExKits
        '
        Me.lvExKits.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvExKits.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.chListPriceKit, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.chPriceKit, Me.chCount, Me.chKitSold, Me.chAvgPrice})
        Me.lvExKits.FullRowSelect = True
        Me.lvExKits.GridLines = True
        Me.lvExKits.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvExKits.Location = New System.Drawing.Point(16, 24)
        Me.lvExKits.Name = "lvExKits"
        Me.lvExKits.Size = New System.Drawing.Size(864, 208)
        Me.lvExKits.TabIndex = 1
        Me.lvExKits.UseCompatibleStateImageBehavior = False
        Me.lvExKits.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Description/Kit"
        Me.ColumnHeader2.Width = 300
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = ""
        Me.ColumnHeader3.Width = 40
        '
        'chListPriceKit
        '
        Me.chListPriceKit.Text = "List Price"
        Me.chListPriceKit.Width = 100
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = ""
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Discount (%)"
        Me.ColumnHeader6.Width = 100
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = ""
        '
        'chPriceKit
        '
        Me.chPriceKit.Text = "Price"
        '
        'chCount
        '
        Me.chCount.Text = "#Invoices"
        '
        'chKitSold
        '
        Me.chKitSold.Text = "#Sold"
        '
        'chAvgPrice
        '
        Me.chAvgPrice.Text = "Avg Price"
        '
        'tabpCancellations
        '
        Me.tabpCancellations.Controls.Add(Me.pnlCancellations)
        Me.tabpCancellations.Location = New System.Drawing.Point(4, 4)
        Me.tabpCancellations.Name = "tabpCancellations"
        Me.tabpCancellations.Size = New System.Drawing.Size(912, 494)
        Me.tabpCancellations.TabIndex = 5
        Me.tabpCancellations.Text = "Cancellations"
        '
        'pnlCancellations
        '
        Me.pnlCancellations.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCancellations.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCancellations.Controls.Add(Me.dpCallOutCancel)
        Me.pnlCancellations.Controls.Add(Me.dpFixedSiteCancellation)
        Me.pnlCancellations.Controls.Add(Me.dpInHoursCancel)
        Me.pnlCancellations.Controls.Add(Me.dpOutOfhoursRoutine)
        Me.pnlCancellations.Controls.Add(Me.Label18)
        Me.pnlCancellations.Controls.Add(Me.Label25)
        Me.pnlCancellations.Controls.Add(Me.Label45)
        Me.pnlCancellations.Controls.Add(Me.Label52)
        Me.pnlCancellations.Controls.Add(Me.Label53)
        Me.pnlCancellations.Controls.Add(Me.Label59)
        Me.pnlCancellations.Controls.Add(Me.Label60)
        Me.pnlCancellations.Location = New System.Drawing.Point(16, 24)
        Me.pnlCancellations.Name = "pnlCancellations"
        Me.pnlCancellations.Size = New System.Drawing.Size(888, 216)
        Me.pnlCancellations.TabIndex = 0
        '
        'dpCallOutCancel
        '
        Me.dpCallOutCancel.AllowLeave = True
        Me.dpCallOutCancel.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase19.AllowLeave = False
        DiscountPanelBase19.DiscountScheme = ""
        DiscountPanelBase19.Height = 40
        DiscountPanelBase19.Label = "Call Out Cancellation"
        DiscountPanelBase19.Left = 0
        DiscountPanelBase19.LineItem = "APCOLC005"
        DiscountPanelBase19.Name = Nothing
        DiscountPanelBase19.Panel = Nothing
        DiscountPanelBase19.PanelType = 0
        DiscountPanelBase19.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase19.Service = "CALLOUT"
        DiscountPanelBase19.Top = 0
        DiscountPanelBase19.Width = 884
        Me.dpCallOutCancel.BasePanel = DiscountPanelBase19
        Me.dpCallOutCancel.Client = Nothing
        Me.dpCallOutCancel.ContextMenu = Me.ctxAddService
        Me.dpCallOutCancel.Discount = ""
        Me.dpCallOutCancel.DiscountScheme = ""
        Me.dpCallOutCancel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dpCallOutCancel.LineItem = "APCOLC005"
        Me.dpCallOutCancel.ListPrice = ""
        Me.dpCallOutCancel.Location = New System.Drawing.Point(0, 52)
        Me.dpCallOutCancel.Name = "dpCallOutCancel"
        Me.dpCallOutCancel.PanelLabel = "Call Out Cancellation"
        Me.dpCallOutCancel.PanelService = Nothing
        Me.dpCallOutCancel.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Cancellation
        Me.dpCallOutCancel.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpCallOutCancel.PriceEnabled = True
        Me.dpCallOutCancel.Scheme = ""
        Me.dpCallOutCancel.Service = "CALLOUT"
        Me.dpCallOutCancel.size = New System.Drawing.Size(884, 40)
        Me.dpCallOutCancel.size = New System.Drawing.Size(884, 40)
        Me.dpCallOutCancel.TabIndex = 66
        '
        'dpFixedSiteCancellation
        '
        Me.dpFixedSiteCancellation.AllowLeave = True
        Me.dpFixedSiteCancellation.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase20.AllowLeave = False
        DiscountPanelBase20.DiscountScheme = ""
        DiscountPanelBase20.Height = 40
        DiscountPanelBase20.Label = "Fixed Site Cancellation"
        DiscountPanelBase20.Left = 0
        DiscountPanelBase20.LineItem = "APCOLF003"
        DiscountPanelBase20.Name = Nothing
        DiscountPanelBase20.Panel = Nothing
        DiscountPanelBase20.PanelType = 0
        DiscountPanelBase20.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase20.Service = "FIXED"
        DiscountPanelBase20.Top = 0
        DiscountPanelBase20.Width = 884
        Me.dpFixedSiteCancellation.BasePanel = DiscountPanelBase20
        Me.dpFixedSiteCancellation.Client = Nothing
        Me.dpFixedSiteCancellation.ContextMenu = Me.ctxAddService
        Me.dpFixedSiteCancellation.Discount = ""
        Me.dpFixedSiteCancellation.DiscountScheme = ""
        Me.dpFixedSiteCancellation.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dpFixedSiteCancellation.LineItem = "APCOLF003"
        Me.dpFixedSiteCancellation.ListPrice = ""
        Me.dpFixedSiteCancellation.Location = New System.Drawing.Point(0, 92)
        Me.dpFixedSiteCancellation.Name = "dpFixedSiteCancellation"
        Me.dpFixedSiteCancellation.PanelLabel = "Fixed Site Cancellation"
        Me.dpFixedSiteCancellation.PanelService = Nothing
        Me.dpFixedSiteCancellation.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Cancellation
        Me.dpFixedSiteCancellation.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpFixedSiteCancellation.PriceEnabled = True
        Me.dpFixedSiteCancellation.Scheme = ""
        Me.dpFixedSiteCancellation.Service = "FIXED"
        Me.dpFixedSiteCancellation.size = New System.Drawing.Size(884, 40)
        Me.dpFixedSiteCancellation.size = New System.Drawing.Size(884, 40)
        Me.dpFixedSiteCancellation.TabIndex = 65
        '
        'dpInHoursCancel
        '
        Me.dpInHoursCancel.AllowLeave = True
        Me.dpInHoursCancel.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase21.AllowLeave = False
        DiscountPanelBase21.DiscountScheme = ""
        DiscountPanelBase21.Height = 40
        DiscountPanelBase21.Label = "In Hours (Routine)"
        DiscountPanelBase21.Left = 0
        DiscountPanelBase21.LineItem = "APCOLS005"
        DiscountPanelBase21.Name = Nothing
        DiscountPanelBase21.Panel = Nothing
        DiscountPanelBase21.PanelType = 0
        DiscountPanelBase21.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase21.Service = "ROUTINE"
        DiscountPanelBase21.Top = 0
        DiscountPanelBase21.Width = 884
        Me.dpInHoursCancel.BasePanel = DiscountPanelBase21
        Me.dpInHoursCancel.Client = Nothing
        Me.dpInHoursCancel.ContextMenu = Me.ctxAddService
        Me.dpInHoursCancel.Discount = ""
        Me.dpInHoursCancel.DiscountScheme = ""
        Me.dpInHoursCancel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dpInHoursCancel.LineItem = "APCOLS005"
        Me.dpInHoursCancel.ListPrice = ""
        Me.dpInHoursCancel.Location = New System.Drawing.Point(0, 132)
        Me.dpInHoursCancel.Name = "dpInHoursCancel"
        Me.dpInHoursCancel.PanelLabel = "In Hours (Routine)"
        Me.dpInHoursCancel.PanelService = Nothing
        Me.dpInHoursCancel.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Cancellation
        Me.dpInHoursCancel.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpInHoursCancel.PriceEnabled = True
        Me.dpInHoursCancel.Scheme = ""
        Me.dpInHoursCancel.Service = "ROUTINE"
        Me.dpInHoursCancel.size = New System.Drawing.Size(884, 40)
        Me.dpInHoursCancel.size = New System.Drawing.Size(884, 40)
        Me.dpInHoursCancel.TabIndex = 64
        '
        'dpOutOfhoursRoutine
        '
        Me.dpOutOfhoursRoutine.AllowLeave = True
        Me.dpOutOfhoursRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase22.AllowLeave = False
        DiscountPanelBase22.DiscountScheme = ""
        DiscountPanelBase22.Height = 40
        DiscountPanelBase22.Label = "Out of Hours (Routine)"
        DiscountPanelBase22.Left = 0
        DiscountPanelBase22.LineItem = "APCOLS006"
        DiscountPanelBase22.Name = Nothing
        DiscountPanelBase22.Panel = Nothing
        DiscountPanelBase22.PanelType = 0
        DiscountPanelBase22.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        DiscountPanelBase22.Service = "ROUTINE"
        DiscountPanelBase22.Top = 0
        DiscountPanelBase22.Width = 884
        Me.dpOutOfhoursRoutine.BasePanel = DiscountPanelBase22
        Me.dpOutOfhoursRoutine.Client = Nothing
        Me.dpOutOfhoursRoutine.ContextMenu = Me.ctxAddService
        Me.dpOutOfhoursRoutine.Discount = ""
        Me.dpOutOfhoursRoutine.DiscountScheme = ""
        Me.dpOutOfhoursRoutine.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dpOutOfhoursRoutine.LineItem = "APCOLS006"
        Me.dpOutOfhoursRoutine.ListPrice = ""
        Me.dpOutOfhoursRoutine.Location = New System.Drawing.Point(0, 172)
        Me.dpOutOfhoursRoutine.Name = "dpOutOfhoursRoutine"
        Me.dpOutOfhoursRoutine.PanelLabel = "Out of Hours (Routine)"
        Me.dpOutOfhoursRoutine.PanelService = Nothing
        Me.dpOutOfhoursRoutine.PanelType = MedscreenCommonGui.Panels.DiscountPanel.TypesOfPanel.Cancellation
        Me.dpOutOfhoursRoutine.PriceDate = New Date(2010, 3, 4, 0, 0, 0, 0)
        Me.dpOutOfhoursRoutine.PriceEnabled = True
        Me.dpOutOfhoursRoutine.Scheme = ""
        Me.dpOutOfhoursRoutine.Service = "ROUTINE"
        Me.dpOutOfhoursRoutine.size = New System.Drawing.Size(884, 40)
        Me.dpOutOfhoursRoutine.size = New System.Drawing.Size(884, 40)
        Me.dpOutOfhoursRoutine.TabIndex = 63
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(488, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(19, 13)
        Me.Label18.TabIndex = 62
        Me.Label18.Text = "(£)"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(272, 8)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(19, 13)
        Me.Label25.TabIndex = 61
        Me.Label25.Text = "(£)"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(552, 8)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(46, 13)
        Me.Label45.TabIndex = 60
        Me.Label45.Text = "Scheme"
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Location = New System.Drawing.Point(448, 8)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(31, 13)
        Me.Label52.TabIndex = 59
        Me.Label52.Text = "Price"
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(328, 8)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(100, 23)
        Me.Label53.TabIndex = 58
        Me.Label53.Text = "Discount %"
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Location = New System.Drawing.Point(216, 8)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(50, 13)
        Me.Label59.TabIndex = 57
        Me.Label59.Text = "List Price"
        '
        'Label60
        '
        Me.Label60.Location = New System.Drawing.Point(48, 8)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(100, 23)
        Me.Label60.TabIndex = 56
        Me.Label60.Text = "Sample Type "
        '
        'TabDiscSchemes
        '
        Me.TabDiscSchemes.Controls.Add(Me.pnlDiscounts)
        Me.TabDiscSchemes.Location = New System.Drawing.Point(4, 22)
        Me.TabDiscSchemes.Name = "TabDiscSchemes"
        Me.TabDiscSchemes.Size = New System.Drawing.Size(920, 627)
        Me.TabDiscSchemes.TabIndex = 8
        Me.TabDiscSchemes.Text = "Discount Schemes"
        '
        'pnlDiscounts
        '
        Me.pnlDiscounts.CustomerID = Nothing
        Me.pnlDiscounts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDiscounts.Location = New System.Drawing.Point(0, 0)
        Me.pnlDiscounts.Name = "pnlDiscounts"
        Me.pnlDiscounts.Size = New System.Drawing.Size(920, 627)
        Me.pnlDiscounts.TabIndex = 0
        '
        'TabInvoicing
        '
        Me.TabInvoicing.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabInvoicing.Controls.Add(Me.cbInvFreq)
        Me.TabInvoicing.Controls.Add(Me.Label33)
        Me.TabInvoicing.Controls.Add(Me.Label24)
        Me.TabInvoicing.Controls.Add(Me.lbServices)
        Me.TabInvoicing.Controls.Add(Me.cbPricing)
        Me.TabInvoicing.Controls.Add(Me.Label23)
        Me.TabInvoicing.Controls.Add(Me.Label21)
        Me.TabInvoicing.Controls.Add(Me.txtVAT)
        Me.TabInvoicing.Controls.Add(Me.cbCurrency)
        Me.TabInvoicing.Controls.Add(Me.lblCurrency)
        Me.TabInvoicing.Controls.Add(Me.ckCreditCard)
        Me.TabInvoicing.Controls.Add(Me.lblSageID)
        Me.TabInvoicing.Controls.Add(Me.txtSageID)
        Me.TabInvoicing.Controls.Add(Me.cbPlatAccType)
        Me.TabInvoicing.Controls.Add(Me.Label17)
        Me.TabInvoicing.Controls.Add(Me.ckPORequired)
        Me.TabInvoicing.Controls.Add(Me.panelStd)
        Me.TabInvoicing.Location = New System.Drawing.Point(4, 22)
        Me.TabInvoicing.Name = "TabInvoicing"
        Me.TabInvoicing.Size = New System.Drawing.Size(920, 627)
        Me.TabInvoicing.TabIndex = 3
        Me.TabInvoicing.Text = "Invoicing"
        '
        'cbInvFreq
        '
        Me.cbInvFreq.Location = New System.Drawing.Point(149, 120)
        Me.cbInvFreq.Name = "cbInvFreq"
        Me.cbInvFreq.Size = New System.Drawing.Size(419, 21)
        Me.cbInvFreq.TabIndex = 30
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(7, 120)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(100, 23)
        Me.Label33.TabIndex = 29
        Me.Label33.Text = "Invoice Frequency"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label24
        '
        Me.Label24.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label24.Location = New System.Drawing.Point(592, 8)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(100, 23)
        Me.Label24.TabIndex = 28
        Me.Label24.Text = "Services"
        '
        'lbServices
        '
        Me.lbServices.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbServices.Location = New System.Drawing.Point(592, 40)
        Me.lbServices.Name = "lbServices"
        Me.lbServices.Size = New System.Drawing.Size(136, 82)
        Me.lbServices.TabIndex = 27
        '
        'cbPricing
        '
        Me.cbPricing.Items.AddRange(New Object() {"Standard Pricing ", "Discount Pricing "})
        Me.cbPricing.Location = New System.Drawing.Point(149, 96)
        Me.cbPricing.Name = "cbPricing"
        Me.cbPricing.Size = New System.Drawing.Size(419, 21)
        Me.cbPricing.TabIndex = 25
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(7, 96)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(41, 23)
        Me.Label23.TabIndex = 24
        Me.Label23.Text = "P&ricing"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(328, 64)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(56, 23)
        Me.Label21.TabIndex = 23
        Me.Label21.Text = "&VAT Rate"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVAT
        '
        Me.txtVAT.Enabled = False
        Me.txtVAT.Location = New System.Drawing.Point(400, 64)
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(64, 20)
        Me.txtVAT.TabIndex = 22
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(424, 10)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(121, 21)
        Me.cbCurrency.TabIndex = 21
        '
        'lblCurrency
        '
        Me.lblCurrency.Location = New System.Drawing.Point(360, 10)
        Me.lblCurrency.Name = "lblCurrency"
        Me.lblCurrency.Size = New System.Drawing.Size(56, 50)
        Me.lblCurrency.TabIndex = 20
        Me.lblCurrency.Text = "C&urrency"
        Me.lblCurrency.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ckCreditCard
        '
        Me.ckCreditCard.Location = New System.Drawing.Point(40, 40)
        Me.ckCreditCard.Name = "ckCreditCard"
        Me.ckCreditCard.Size = New System.Drawing.Size(168, 24)
        Me.ckCreditCard.TabIndex = 19
        Me.ckCreditCard.Text = "&Card Payment Required"
        '
        'lblSageID
        '
        Me.lblSageID.Location = New System.Drawing.Point(208, 10)
        Me.lblSageID.Name = "lblSageID"
        Me.lblSageID.Size = New System.Drawing.Size(48, 23)
        Me.lblSageID.TabIndex = 18
        Me.lblSageID.Text = "&Sage ID"
        Me.lblSageID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSageID
        '
        Me.txtSageID.Location = New System.Drawing.Point(256, 10)
        Me.txtSageID.Name = "txtSageID"
        Me.txtSageID.Size = New System.Drawing.Size(64, 20)
        Me.txtSageID.TabIndex = 17
        '
        'cbPlatAccType
        '
        Me.cbPlatAccType.Location = New System.Drawing.Point(152, 64)
        Me.cbPlatAccType.Name = "cbPlatAccType"
        Me.cbPlatAccType.Size = New System.Drawing.Size(176, 21)
        Me.cbPlatAccType.TabIndex = 8
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(16, 64)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(128, 23)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "Platinum &Account Type"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckPORequired
        '
        Me.ckPORequired.Location = New System.Drawing.Point(40, 10)
        Me.ckPORequired.Name = "ckPORequired"
        Me.ckPORequired.Size = New System.Drawing.Size(168, 24)
        Me.ckPORequired.TabIndex = 0
        Me.ckPORequired.Text = "&Purchase Order Required"
        '
        'panelStd
        '
        Me.panelStd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panelStd.Controls.Add(Me.HtmlEditor1)
        Me.panelStd.Location = New System.Drawing.Point(8, 144)
        Me.panelStd.Name = "panelStd"
        Me.panelStd.Size = New System.Drawing.Size(916, 440)
        Me.panelStd.TabIndex = 26
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 0)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(916, 440)
        Me.HtmlEditor1.TabIndex = 1
        '
        'TabReporting
        '
        Me.TabReporting.Controls.Add(Me.ckDivertPos)
        Me.TabReporting.Controls.Add(Me.chkEmailAndPrint)
        Me.TabReporting.Controls.Add(Me.TabOptions)
        Me.TabReporting.Controls.Add(Me.ckEmailPosNeg)
        Me.TabReporting.Controls.Add(Me.ckPositivesOnly)
        Me.TabReporting.Controls.Add(Me.cbFaxFormat)
        Me.TabReporting.Controls.Add(Me.Label15)
        Me.TabReporting.Controls.Add(Me.cbReportMethod)
        Me.TabReporting.Controls.Add(Me.Label14)
        Me.TabReporting.Controls.Add(Me.ckPhoneBeforeFaxing)
        Me.TabReporting.Controls.Add(Me.ckEmailWeeklyReport)
        Me.TabReporting.Controls.Add(Me.ckReportInBits)
        Me.TabReporting.Location = New System.Drawing.Point(4, 22)
        Me.TabReporting.Name = "TabReporting"
        Me.TabReporting.Size = New System.Drawing.Size(920, 627)
        Me.TabReporting.TabIndex = 2
        Me.TabReporting.Text = "Options"
        '
        'chkEmailAndPrint
        '
        Me.chkEmailAndPrint.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkEmailAndPrint.Location = New System.Drawing.Point(616, 50)
        Me.chkEmailAndPrint.Name = "chkEmailAndPrint"
        Me.chkEmailAndPrint.Size = New System.Drawing.Size(128, 24)
        Me.chkEmailAndPrint.TabIndex = 14
        Me.chkEmailAndPrint.Text = "Email and &Print"
        Me.chkEmailAndPrint.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'TabOptions
        '
        Me.TabOptions.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabOptions.Location = New System.Drawing.Point(0, 120)
        Me.TabOptions.Name = "TabOptions"
        Me.TabOptions.SelectedIndex = 0
        Me.TabOptions.Size = New System.Drawing.Size(921, 472)
        Me.TabOptions.TabIndex = 9
        '
        'ckEmailPosNeg
        '
        Me.ckEmailPosNeg.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckEmailPosNeg.Location = New System.Drawing.Point(376, 50)
        Me.ckEmailPosNeg.Name = "ckEmailPosNeg"
        Me.ckEmailPosNeg.Size = New System.Drawing.Size(234, 37)
        Me.ckEmailPosNeg.TabIndex = 8
        Me.ckEmailPosNeg.Text = "&Email Positives and Negatives to separate addresses"
        Me.ckEmailPosNeg.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ckPositivesOnly
        '
        Me.ckPositivesOnly.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckPositivesOnly.Location = New System.Drawing.Point(184, 50)
        Me.ckPositivesOnly.Name = "ckPositivesOnly"
        Me.ckPositivesOnly.Size = New System.Drawing.Size(160, 24)
        Me.ckPositivesOnly.TabIndex = 7
        Me.ckPositivesOnly.Text = "&Report Positives Only"
        Me.ckPositivesOnly.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'cbFaxFormat
        '
        Me.cbFaxFormat.Location = New System.Drawing.Point(464, 21)
        Me.cbFaxFormat.Name = "cbFaxFormat"
        Me.cbFaxFormat.Size = New System.Drawing.Size(176, 21)
        Me.cbFaxFormat.TabIndex = 4
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(344, 21)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(100, 23)
        Me.Label15.TabIndex = 3
        Me.Label15.Text = "&Fax Format"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbReportMethod
        '
        Me.cbReportMethod.Location = New System.Drawing.Point(152, 21)
        Me.cbReportMethod.Name = "cbReportMethod"
        Me.cbReportMethod.Size = New System.Drawing.Size(176, 21)
        Me.cbReportMethod.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(32, 21)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(100, 23)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = "Report &Method"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckPhoneBeforeFaxing
        '
        Me.ckPhoneBeforeFaxing.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckPhoneBeforeFaxing.Location = New System.Drawing.Point(32, 50)
        Me.ckPhoneBeforeFaxing.Name = "ckPhoneBeforeFaxing"
        Me.ckPhoneBeforeFaxing.Size = New System.Drawing.Size(136, 24)
        Me.ckPhoneBeforeFaxing.TabIndex = 0
        Me.ckPhoneBeforeFaxing.Text = "&Phone before faxing "
        Me.ckPhoneBeforeFaxing.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ckEmailWeeklyReport
        '
        Me.ckEmailWeeklyReport.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckEmailWeeklyReport.Location = New System.Drawing.Point(32, 85)
        Me.ckEmailWeeklyReport.Name = "ckEmailWeeklyReport"
        Me.ckEmailWeeklyReport.Size = New System.Drawing.Size(150, 24)
        Me.ckEmailWeeklyReport.TabIndex = 10
        Me.ckEmailWeeklyReport.Text = "&Mail Weekly Report"
        Me.ckEmailWeeklyReport.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ckReportInBits
        '
        Me.ckReportInBits.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckReportInBits.Location = New System.Drawing.Point(184, 85)
        Me.ckReportInBits.Name = "ckReportInBits"
        Me.ckReportInBits.Size = New System.Drawing.Size(160, 24)
        Me.ckReportInBits.TabIndex = 11
        Me.ckReportInBits.Text = "Report incomplete &Jobs"
        Me.ckReportInBits.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ckDivertPos
        '
        Me.ckDivertPos.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ckDivertPos.Location = New System.Drawing.Point(376, 85)
        Me.ckDivertPos.Name = "ckDivertPos"
        Me.ckDivertPos.Size = New System.Drawing.Size(224, 24)
        Me.ckDivertPos.TabIndex = 13
        Me.ckDivertPos.Text = "&Send Positives to secondary address"
        Me.ckDivertPos.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'TabContracts
        '
        Me.TabContracts.Controls.Add(Me.HtmlEditor3)
        Me.TabContracts.Controls.Add(Me.Splitter2)
        Me.TabContracts.Controls.Add(Me.lblThen)
        Me.TabContracts.Controls.Add(Me.dtFirstPayment)
        Me.TabContracts.Controls.Add(Me.lblFirstPayment)
        Me.TabContracts.Controls.Add(Me.ckContractRemoved)
        Me.TabContracts.Controls.Add(Me.cmdDeleteContract)
        Me.TabContracts.Controls.Add(Me.DTContractEndDate)
        Me.TabContracts.Controls.Add(Me.Label41)
        Me.TabContracts.Controls.Add(Me.cbLineItems)
        Me.TabContracts.Controls.Add(Me.Label40)
        Me.TabContracts.Controls.Add(Me.cbPayInterval)
        Me.TabContracts.Controls.Add(Me.Label39)
        Me.TabContracts.Controls.Add(Me.cmdUpdate)
        Me.TabContracts.Controls.Add(Me.DTNextDue)
        Me.TabContracts.Controls.Add(Me.Label38)
        Me.TabContracts.Controls.Add(Me.txtNextPayment)
        Me.TabContracts.Controls.Add(Me.Label37)
        Me.TabContracts.Controls.Add(Me.txtAnnualPrice)
        Me.TabContracts.Controls.Add(Me.Label36)
        Me.TabContracts.Controls.Add(Me.DTContractStart)
        Me.TabContracts.Controls.Add(Me.Label35)
        Me.TabContracts.Controls.Add(Me.txtContractNumber1)
        Me.TabContracts.Controls.Add(Me.Label34)
        Me.TabContracts.Controls.Add(Me.cmdAdd)
        Me.TabContracts.Controls.Add(Me.lvRoutinePayments)
        Me.TabContracts.Controls.Add(Me.LblItemComment)
        Me.TabContracts.Controls.Add(Me.txtItemComment)
        Me.TabContracts.Location = New System.Drawing.Point(4, 22)
        Me.TabContracts.Name = "TabContracts"
        Me.TabContracts.Size = New System.Drawing.Size(920, 627)
        Me.TabContracts.TabIndex = 4
        Me.TabContracts.Text = "Contracts"
        '
        'HtmlEditor3
        '
        Me.HtmlEditor3.Dock = System.Windows.Forms.DockStyle.Top
        Me.HtmlEditor3.Location = New System.Drawing.Point(0, 131)
        Me.HtmlEditor3.Name = "HtmlEditor3"
        Me.HtmlEditor3.Size = New System.Drawing.Size(920, 229)
        Me.HtmlEditor3.TabIndex = 25
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter2.Location = New System.Drawing.Point(0, 128)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(920, 3)
        Me.Splitter2.TabIndex = 24
        Me.Splitter2.TabStop = False
        '
        'lblThen
        '
        Me.lblThen.AutoSize = True
        Me.lblThen.Location = New System.Drawing.Point(280, 562)
        Me.lblThen.Name = "lblThen"
        Me.lblThen.Size = New System.Drawing.Size(31, 13)
        Me.lblThen.TabIndex = 23
        Me.lblThen.Text = "then "
        '
        'dtFirstPayment
        '
        Me.dtFirstPayment.Location = New System.Drawing.Point(144, 562)
        Me.dtFirstPayment.Name = "dtFirstPayment"
        Me.dtFirstPayment.Size = New System.Drawing.Size(120, 20)
        Me.dtFirstPayment.TabIndex = 22
        '
        'lblFirstPayment
        '
        Me.lblFirstPayment.AutoSize = True
        Me.lblFirstPayment.Location = New System.Drawing.Point(40, 562)
        Me.lblFirstPayment.Name = "lblFirstPayment"
        Me.lblFirstPayment.Size = New System.Drawing.Size(87, 13)
        Me.lblFirstPayment.TabIndex = 21
        Me.lblFirstPayment.Text = "First Payment On"
        '
        'ckContractRemoved
        '
        Me.ckContractRemoved.Location = New System.Drawing.Point(40, 594)
        Me.ckContractRemoved.Name = "ckContractRemoved"
        Me.ckContractRemoved.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ckContractRemoved.Size = New System.Drawing.Size(104, 24)
        Me.ckContractRemoved.TabIndex = 20
        Me.ckContractRemoved.Text = "Removed"
        Me.ckContractRemoved.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdDeleteContract
        '
        Me.cmdDeleteContract.Image = CType(resources.GetObject("cmdDeleteContract.Image"), System.Drawing.Image)
        Me.cmdDeleteContract.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDeleteContract.Location = New System.Drawing.Point(152, 384)
        Me.cmdDeleteContract.Name = "cmdDeleteContract"
        Me.cmdDeleteContract.Size = New System.Drawing.Size(75, 23)
        Me.cmdDeleteContract.TabIndex = 19
        Me.cmdDeleteContract.Text = "Delete"
        Me.cmdDeleteContract.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDeleteContract.Visible = False
        '
        'DTContractEndDate
        '
        Me.DTContractEndDate.Location = New System.Drawing.Point(536, 408)
        Me.DTContractEndDate.Name = "DTContractEndDate"
        Me.DTContractEndDate.Size = New System.Drawing.Size(120, 20)
        Me.DTContractEndDate.TabIndex = 18
        '
        'Label41
        '
        Me.Label41.Location = New System.Drawing.Point(536, 376)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(100, 23)
        Me.Label41.TabIndex = 17
        Me.Label41.Text = "End Date"
        Me.Label41.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbLineItems
        '
        Me.cbLineItems.Location = New System.Drawing.Point(384, 480)
        Me.cbLineItems.Name = "cbLineItems"
        Me.cbLineItems.Size = New System.Drawing.Size(328, 21)
        Me.cbLineItems.TabIndex = 16
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(320, 480)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(56, 23)
        Me.Label40.TabIndex = 15
        Me.Label40.Text = "Line Item"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbPayInterval
        '
        Me.cbPayInterval.Location = New System.Drawing.Point(504, 562)
        Me.cbPayInterval.Name = "cbPayInterval"
        Me.cbPayInterval.Size = New System.Drawing.Size(121, 21)
        Me.cbPayInterval.TabIndex = 14
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(472, 562)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(28, 13)
        Me.Label39.TabIndex = 13
        Me.Label39.Text = "and "
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Image = CType(resources.GetObject("cmdUpdate.Image"), System.Drawing.Image)
        Me.cmdUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdate.Location = New System.Drawing.Point(544, 594)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(75, 23)
        Me.cmdUpdate.TabIndex = 12
        Me.cmdUpdate.Text = "Update"
        Me.cmdUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTNextDue
        '
        Me.DTNextDue.Location = New System.Drawing.Point(328, 562)
        Me.DTNextDue.Name = "DTNextDue"
        Me.DTNextDue.Size = New System.Drawing.Size(136, 20)
        Me.DTNextDue.TabIndex = 11
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(632, 562)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(52, 13)
        Me.Label38.TabIndex = 10
        Me.Label38.Text = "thereafter"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtNextPayment
        '
        Me.txtNextPayment.Location = New System.Drawing.Point(384, 448)
        Me.txtNextPayment.Name = "txtNextPayment"
        Me.txtNextPayment.Size = New System.Drawing.Size(100, 20)
        Me.txtNextPayment.TabIndex = 9
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(256, 448)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(112, 13)
        Me.Label37.TabIndex = 8
        Me.Label37.Text = "Next Payment Amount"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAnnualPrice
        '
        Me.txtAnnualPrice.Location = New System.Drawing.Point(144, 448)
        Me.txtAnnualPrice.Name = "txtAnnualPrice"
        Me.txtAnnualPrice.Size = New System.Drawing.Size(100, 20)
        Me.txtAnnualPrice.TabIndex = 7
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(40, 448)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(67, 13)
        Me.Label36.TabIndex = 6
        Me.Label36.Text = "Annual Price"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTContractStart
        '
        Me.DTContractStart.Location = New System.Drawing.Point(384, 408)
        Me.DTContractStart.Name = "DTContractStart"
        Me.DTContractStart.Size = New System.Drawing.Size(120, 20)
        Me.DTContractStart.TabIndex = 5
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(384, 376)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(100, 23)
        Me.Label35.TabIndex = 4
        Me.Label35.Text = "Start Date"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtContractNumber1
        '
        Me.txtContractNumber1.Location = New System.Drawing.Point(144, 416)
        Me.txtContractNumber1.Name = "txtContractNumber1"
        Me.txtContractNumber1.Size = New System.Drawing.Size(100, 20)
        Me.txtContractNumber1.TabIndex = 3
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(40, 416)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(100, 23)
        Me.Label34.TabIndex = 2
        Me.Label34.Text = "Contract Number"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAdd
        '
        Me.cmdAdd.Image = CType(resources.GetObject("cmdAdd.Image"), System.Drawing.Image)
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(40, 384)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvRoutinePayments
        '
        Me.lvRoutinePayments.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColContNum, Me.ColItemCode, Me.ColStartDate, Me.ColConEndDate, Me.ColAnnPrice, Me.ColNextAmount, Me.ColPayInterval, Me.colDueNext, Me.colLastInvoice})
        Me.lvRoutinePayments.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvRoutinePayments.FullRowSelect = True
        Me.lvRoutinePayments.Location = New System.Drawing.Point(0, 0)
        Me.lvRoutinePayments.Name = "lvRoutinePayments"
        Me.lvRoutinePayments.Size = New System.Drawing.Size(920, 128)
        Me.lvRoutinePayments.TabIndex = 0
        Me.lvRoutinePayments.UseCompatibleStateImageBehavior = False
        Me.lvRoutinePayments.View = System.Windows.Forms.View.Details
        '
        'ColContNum
        '
        Me.ColContNum.Text = "Contract number"
        Me.ColContNum.Width = 100
        '
        'ColItemCode
        '
        Me.ColItemCode.Text = "Item Code"
        Me.ColItemCode.Width = 80
        '
        'ColStartDate
        '
        Me.ColStartDate.Text = "Start Date"
        Me.ColStartDate.Width = 80
        '
        'ColConEndDate
        '
        Me.ColConEndDate.Text = "Contract End"
        Me.ColConEndDate.Width = 80
        '
        'ColAnnPrice
        '
        Me.ColAnnPrice.Text = "Annual Price"
        Me.ColAnnPrice.Width = 80
        '
        'ColNextAmount
        '
        Me.ColNextAmount.Text = "Next Amount"
        Me.ColNextAmount.Width = 90
        '
        'ColPayInterval
        '
        Me.ColPayInterval.Text = "Pay Interval"
        Me.ColPayInterval.Width = 80
        '
        'colDueNext
        '
        Me.colDueNext.Text = "Due Next"
        Me.colDueNext.Width = 80
        '
        'colLastInvoice
        '
        Me.colLastInvoice.Text = "Last Invoice"
        Me.colLastInvoice.Width = 100
        '
        'LblItemComment
        '
        Me.LblItemComment.Location = New System.Drawing.Point(248, 512)
        Me.LblItemComment.Name = "LblItemComment"
        Me.LblItemComment.Size = New System.Drawing.Size(128, 23)
        Me.LblItemComment.TabIndex = 26
        Me.LblItemComment.Text = "Comment on Line item"
        Me.LblItemComment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtItemComment
        '
        Me.txtItemComment.Location = New System.Drawing.Point(384, 512)
        Me.txtItemComment.Name = "txtItemComment"
        Me.txtItemComment.Size = New System.Drawing.Size(328, 20)
        Me.txtItemComment.TabIndex = 27
        '
        'tabComments
        '
        Me.tabComments.Controls.Add(Me.cmdAddComment)
        Me.tabComments.Controls.Add(Me.pnlComments)
        Me.tabComments.Location = New System.Drawing.Point(4, 22)
        Me.tabComments.Name = "tabComments"
        Me.tabComments.Size = New System.Drawing.Size(920, 627)
        Me.tabComments.TabIndex = 6
        Me.tabComments.Text = "Comments"
        '
        'cmdAddComment
        '
        Me.cmdAddComment.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAddComment.Image = CType(resources.GetObject("cmdAddComment.Image"), System.Drawing.Image)
        Me.cmdAddComment.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddComment.Location = New System.Drawing.Point(760, 560)
        Me.cmdAddComment.Name = "cmdAddComment"
        Me.cmdAddComment.Size = New System.Drawing.Size(80, 23)
        Me.cmdAddComment.TabIndex = 1
        Me.cmdAddComment.Text = "Comment"
        Me.cmdAddComment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlComments
        '
        Me.pnlComments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlComments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlComments.Controls.Add(Me.lvComments)
        Me.pnlComments.Location = New System.Drawing.Point(16, 8)
        Me.pnlComments.Name = "pnlComments"
        Me.pnlComments.Size = New System.Drawing.Size(888, 528)
        Me.pnlComments.TabIndex = 0
        '
        'lvComments
        '
        Me.lvComments.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chCommDate, Me.chCommType, Me.CHComment, Me.chCommBy, Me.ChCommAction})
        Me.lvComments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvComments.Location = New System.Drawing.Point(0, 0)
        Me.lvComments.Name = "lvComments"
        Me.lvComments.Size = New System.Drawing.Size(884, 524)
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
        'TabTestPanels
        '
        Me.TabTestPanels.Controls.Add(Me.cmdDeleteTest)
        Me.TabTestPanels.Controls.Add(Me.cmdAddTest)
        Me.TabTestPanels.Controls.Add(Me.CmdCreateTestPanels)
        Me.TabTestPanels.Controls.Add(Me.lvTestSchedules)
        Me.TabTestPanels.Controls.Add(Me.Label43)
        Me.TabTestPanels.Controls.Add(Me.lvTestPanels)
        Me.TabTestPanels.Controls.Add(Me.ckAddTests)
        Me.TabTestPanels.Controls.Add(Me.ckInvalidPanel)
        Me.TabTestPanels.Location = New System.Drawing.Point(4, 22)
        Me.TabTestPanels.Name = "TabTestPanels"
        Me.TabTestPanels.Size = New System.Drawing.Size(920, 627)
        Me.TabTestPanels.TabIndex = 5
        Me.TabTestPanels.Text = "Test Panels"
        '
        'cmdDeleteTest
        '
        Me.cmdDeleteTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDeleteTest.Location = New System.Drawing.Point(112, 328)
        Me.cmdDeleteTest.Name = "cmdDeleteTest"
        Me.cmdDeleteTest.Size = New System.Drawing.Size(88, 23)
        Me.cmdDeleteTest.TabIndex = 12
        Me.cmdDeleteTest.Text = "Delete Test"
        Me.cmdDeleteTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDeleteTest.Visible = False
        '
        'cmdAddTest
        '
        Me.cmdAddTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddTest.Location = New System.Drawing.Point(16, 328)
        Me.cmdAddTest.Name = "cmdAddTest"
        Me.cmdAddTest.Size = New System.Drawing.Size(88, 23)
        Me.cmdAddTest.TabIndex = 11
        Me.cmdAddTest.Text = "Add Test"
        Me.cmdAddTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAddTest.Visible = False
        '
        'CmdCreateTestPanels
        '
        Me.CmdCreateTestPanels.Image = CType(resources.GetObject("CmdCreateTestPanels.Image"), System.Drawing.Image)
        Me.CmdCreateTestPanels.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.CmdCreateTestPanels.Location = New System.Drawing.Point(184, 120)
        Me.CmdCreateTestPanels.Name = "CmdCreateTestPanels"
        Me.CmdCreateTestPanels.Size = New System.Drawing.Size(75, 23)
        Me.CmdCreateTestPanels.TabIndex = 3
        Me.CmdCreateTestPanels.Text = "&Create"
        Me.CmdCreateTestPanels.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CmdCreateTestPanels.Visible = False
        '
        'lvTestSchedules
        '
        Me.lvTestSchedules.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTestSchedules.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chAnalysis, Me.chTSDescription, Me.chTSConf, Me.chTSStdTest, Me.ColumnHeader1})
        Me.lvTestSchedules.FullRowSelect = True
        Me.lvTestSchedules.Location = New System.Drawing.Point(0, 152)
        Me.lvTestSchedules.Name = "lvTestSchedules"
        Me.lvTestSchedules.Size = New System.Drawing.Size(960, 152)
        Me.lvTestSchedules.TabIndex = 2
        Me.lvTestSchedules.UseCompatibleStateImageBehavior = False
        Me.lvTestSchedules.View = System.Windows.Forms.View.Details
        '
        'chAnalysis
        '
        Me.chAnalysis.Text = "Analysis"
        Me.chAnalysis.Width = 80
        '
        'chTSDescription
        '
        Me.chTSDescription.Text = "Description"
        Me.chTSDescription.Width = 200
        '
        'chTSConf
        '
        Me.chTSConf.Text = "Confirmatory"
        Me.chTSConf.Width = 100
        '
        'chTSStdTest
        '
        Me.chTSStdTest.Text = "Standard"
        Me.chTSStdTest.Width = 80
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Replicates"
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(16, 120)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(144, 23)
        Me.Label43.TabIndex = 1
        Me.Label43.Text = "Test Schedule Entries "
        '
        'lvTestPanels
        '
        Me.lvTestPanels.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chTestpanelID, Me.chTPDescription, Me.CHDefaultPanel, Me.chMatrix, Me.chAlcoholCutoff, Me.chCannabisCutoff, Me.chBreath})
        Me.lvTestPanels.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvTestPanels.Location = New System.Drawing.Point(0, 0)
        Me.lvTestPanels.Name = "lvTestPanels"
        Me.lvTestPanels.Size = New System.Drawing.Size(920, 97)
        Me.lvTestPanels.TabIndex = 0
        Me.lvTestPanels.UseCompatibleStateImageBehavior = False
        Me.lvTestPanels.View = System.Windows.Forms.View.Details
        '
        'chTestpanelID
        '
        Me.chTestpanelID.Text = "ID"
        Me.chTestpanelID.Width = 80
        '
        'chTPDescription
        '
        Me.chTPDescription.Text = "Description"
        Me.chTPDescription.Width = 200
        '
        'CHDefaultPanel
        '
        Me.CHDefaultPanel.Text = "Default Panel"
        '
        'chMatrix
        '
        Me.chMatrix.Text = "Matrix"
        '
        'chAlcoholCutoff
        '
        Me.chAlcoholCutoff.Text = "Alcohol Cutoff"
        Me.chAlcoholCutoff.Width = 80
        '
        'chCannabisCutoff
        '
        Me.chCannabisCutoff.Text = "Cannabis"
        Me.chCannabisCutoff.Width = 80
        '
        'chBreath
        '
        Me.chBreath.Text = "Breath"
        Me.chBreath.Width = 80
        '
        'ckAddTests
        '
        Me.ckAddTests.Location = New System.Drawing.Point(280, 120)
        Me.ckAddTests.Name = "ckAddTests"
        Me.ckAddTests.Size = New System.Drawing.Size(150, 24)
        Me.ckAddTests.TabIndex = 10
        Me.ckAddTests.Text = "&Allow additional tests"
        '
        'ckInvalidPanel
        '
        Me.ckInvalidPanel.Location = New System.Drawing.Point(450, 120)
        Me.ckInvalidPanel.Name = "ckInvalidPanel"
        Me.ckInvalidPanel.Size = New System.Drawing.Size(150, 24)
        Me.ckInvalidPanel.TabIndex = 10
        Me.ckInvalidPanel.Text = "&Allow invalid panels"
        '
        'TabLoginTemplate
        '
        Me.TabLoginTemplate.Location = New System.Drawing.Point(4, 22)
        Me.TabLoginTemplate.Name = "TabLoginTemplate"
        Me.TabLoginTemplate.Size = New System.Drawing.Size(920, 627)
        Me.TabLoginTemplate.TabIndex = 5
        Me.TabLoginTemplate.Text = "Login Templates"
        '
        'chVolScheme
        '
        Me.chVolScheme.Text = "Scheme-Entry"
        Me.chVolScheme.Width = 100
        '
        'chVolType
        '
        Me.chVolType.Text = "Type"
        Me.chVolType.Width = 100
        '
        'chVolApply
        '
        Me.chVolApply.Text = "Applies To"
        Me.chVolApply.Width = 100
        '
        'chVolMin
        '
        Me.chVolMin.Text = "Min Samples"
        Me.chVolMin.Width = 100
        '
        'chVolMaxSample
        '
        Me.chVolMaxSample.Text = "Max Sample"
        Me.chVolMaxSample.Width = 100
        '
        'chVolMinCharge
        '
        Me.chVolMinCharge.Text = "Min Charge"
        Me.chVolMinCharge.Width = 100
        '
        'chVolDiscount
        '
        Me.chVolDiscount.Text = "Sample Discount"
        Me.chVolDiscount.Width = 100
        '
        'chVolExample
        '
        Me.chVolExample.Text = "Example"
        Me.chVolExample.Width = 400
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(134, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Volume Discount Schemes"
        '
        'pnlNVDScheme
        '
        Me.pnlNVDScheme.Controls.Add(Me.cmdFindScheme)
        Me.pnlNVDScheme.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlNVDScheme.Location = New System.Drawing.Point(0, 240)
        Me.pnlNVDScheme.Name = "pnlNVDScheme"
        Me.pnlNVDScheme.Size = New System.Drawing.Size(920, 32)
        Me.pnlNVDScheme.TabIndex = 2
        '
        'cmdFindScheme
        '
        Me.cmdFindScheme.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdFindScheme.Image = CType(resources.GetObject("cmdFindScheme.Image"), System.Drawing.Image)
        Me.cmdFindScheme.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindScheme.Location = New System.Drawing.Point(752, 8)
        Me.cmdFindScheme.Name = "cmdFindScheme"
        Me.cmdFindScheme.Size = New System.Drawing.Size(112, 23)
        Me.cmdFindScheme.TabIndex = 0
        Me.cmdFindScheme.Text = "Prices"
        Me.cmdFindScheme.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdFindScheme.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(157, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Non Volume Discount Schemes"
        '
        'chNVDID
        '
        Me.chNVDID.Text = "Scheme-Entry"
        Me.chNVDID.Width = 150
        '
        'chNVDType
        '
        Me.chNVDType.Text = "Type"
        Me.chNVDType.Width = 200
        '
        'chNVDApplyto
        '
        Me.chNVDApplyto.Text = "Applys To (Line item)"
        Me.chNVDApplyto.Width = 300
        '
        'chNVDDiscount
        '
        Me.chNVDDiscount.Text = "Discount"
        Me.chNVDDiscount.Width = 120
        '
        'chNVDPrice
        '
        Me.chNVDPrice.Text = "Price"
        '
        'chNVDDisPrice
        '
        Me.chNVDDisPrice.Text = "Discount Price"
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.Label3)
        Me.Panel5.Controls.Add(Me.lblPCurr)
        Me.Panel5.Controls.Add(Me.lblLPCurr)
        Me.Panel5.Controls.Add(Me.Label46)
        Me.Panel5.Controls.Add(Me.Label47)
        Me.Panel5.Controls.Add(Me.Label48)
        Me.Panel5.Controls.Add(Me.Label49)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(692, 48)
        Me.Panel5.TabIndex = 64
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(560, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "Scheme"
        '
        'lblPCurr
        '
        Me.lblPCurr.AutoSize = True
        Me.lblPCurr.Location = New System.Drawing.Point(496, 12)
        Me.lblPCurr.Name = "lblPCurr"
        Me.lblPCurr.Size = New System.Drawing.Size(19, 13)
        Me.lblPCurr.TabIndex = 55
        Me.lblPCurr.Text = "(£)"
        '
        'lblLPCurr
        '
        Me.lblLPCurr.AutoSize = True
        Me.lblLPCurr.Location = New System.Drawing.Point(280, 13)
        Me.lblLPCurr.Name = "lblLPCurr"
        Me.lblLPCurr.Size = New System.Drawing.Size(19, 13)
        Me.lblLPCurr.TabIndex = 54
        Me.lblLPCurr.Text = "(£)"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(458, 13)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(31, 13)
        Me.Label46.TabIndex = 32
        Me.Label46.Text = "Price"
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(338, 13)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(100, 23)
        Me.Label47.TabIndex = 31
        Me.Label47.Text = "Discount %"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(226, 13)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(50, 13)
        Me.Label48.TabIndex = 30
        Me.Label48.Text = "List Price"
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(58, 13)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(100, 23)
        Me.Label49.TabIndex = 24
        Me.Label49.Text = "Sample Type "
        '
        'pnlMinchargeheader
        '
        Me.pnlMinchargeheader.Controls.Add(Me.lblpCollCurr)
        Me.pnlMinchargeheader.Controls.Add(Me.lbllpCollCurr)
        Me.pnlMinchargeheader.Controls.Add(Me.Label54)
        Me.pnlMinchargeheader.Controls.Add(Me.Label55)
        Me.pnlMinchargeheader.Controls.Add(Me.Label56)
        Me.pnlMinchargeheader.Controls.Add(Me.Label57)
        Me.pnlMinchargeheader.Controls.Add(Me.Label58)
        Me.pnlMinchargeheader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMinchargeheader.Location = New System.Drawing.Point(0, 40)
        Me.pnlMinchargeheader.Name = "pnlMinchargeheader"
        Me.pnlMinchargeheader.Size = New System.Drawing.Size(892, 48)
        Me.pnlMinchargeheader.TabIndex = 0
        '
        'lblpCollCurr
        '
        Me.lblpCollCurr.AutoSize = True
        Me.lblpCollCurr.Location = New System.Drawing.Point(518, 12)
        Me.lblpCollCurr.Name = "lblpCollCurr"
        Me.lblpCollCurr.Size = New System.Drawing.Size(19, 13)
        Me.lblpCollCurr.TabIndex = 62
        Me.lblpCollCurr.Text = "(£)"
        '
        'lbllpCollCurr
        '
        Me.lbllpCollCurr.AutoSize = True
        Me.lbllpCollCurr.Location = New System.Drawing.Point(302, 13)
        Me.lbllpCollCurr.Name = "lbllpCollCurr"
        Me.lbllpCollCurr.Size = New System.Drawing.Size(19, 13)
        Me.lbllpCollCurr.TabIndex = 61
        Me.lbllpCollCurr.Text = "(£)"
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Location = New System.Drawing.Point(582, 13)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(46, 13)
        Me.Label54.TabIndex = 60
        Me.Label54.Text = "Scheme"
        '
        'Label55
        '
        Me.Label55.AutoSize = True
        Me.Label55.Location = New System.Drawing.Point(480, 13)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(31, 13)
        Me.Label55.TabIndex = 59
        Me.Label55.Text = "Price"
        '
        'Label56
        '
        Me.Label56.Location = New System.Drawing.Point(360, 13)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(100, 23)
        Me.Label56.TabIndex = 58
        Me.Label56.Text = "Discount %"
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Location = New System.Drawing.Point(248, 13)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(50, 13)
        Me.Label57.TabIndex = 57
        Me.Label57.Text = "List Price"
        '
        'Label58
        '
        Me.Label58.Location = New System.Drawing.Point(80, 13)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(100, 23)
        Me.Label58.TabIndex = 56
        Me.Label58.Text = "Min Charge Type "
        '
        'pnlSelectCharging
        '
        Me.pnlSelectCharging.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlSelectCharging.Controls.Add(Me.Label77)
        Me.pnlSelectCharging.Controls.Add(Me.udCallOut)
        Me.pnlSelectCharging.Controls.Add(Me.Label76)
        Me.pnlSelectCharging.Controls.Add(Me.udRoutine)
        Me.pnlSelectCharging.Controls.Add(Me.Label75)
        Me.pnlSelectCharging.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSelectCharging.Location = New System.Drawing.Point(0, 0)
        Me.pnlSelectCharging.Name = "pnlSelectCharging"
        Me.pnlSelectCharging.Size = New System.Drawing.Size(892, 40)
        Me.pnlSelectCharging.TabIndex = 68
        '
        'Label77
        '
        Me.Label77.Location = New System.Drawing.Point(488, 0)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(200, 24)
        Me.Label77.TabIndex = 4
        Me.Label77.Text = "Tab away from control to activate the new pricing model"
        '
        'udCallOut
        '
        Me.udCallOut.Items.Add("Sample Cost only")
        Me.udCallOut.Items.Add("Attendance Fee")
        Me.udCallOut.Items.Add("Minimum charge")
        Me.udCallOut.Items.Add("Service not sold")
        Me.udCallOut.Location = New System.Drawing.Point(360, 8)
        Me.udCallOut.Name = "udCallOut"
        Me.udCallOut.Size = New System.Drawing.Size(120, 20)
        Me.udCallOut.TabIndex = 3
        Me.udCallOut.Tag = "CO"
        Me.udCallOut.Text = "DomainUpDown1"
        Me.udCallOut.Wrap = True
        '
        'Label76
        '
        Me.Label76.Location = New System.Drawing.Point(256, 8)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(128, 23)
        Me.Label76.TabIndex = 2
        Me.Label76.Text = "Call Out Collections"
        '
        'udRoutine
        '
        Me.udRoutine.Items.Add("Sample Cost only")
        Me.udRoutine.Items.Add("Attendance Fee")
        Me.udRoutine.Items.Add("Minimum charge")
        Me.udRoutine.Items.Add("Service Not Sold")
        Me.udRoutine.Location = New System.Drawing.Point(128, 8)
        Me.udRoutine.Name = "udRoutine"
        Me.udRoutine.Size = New System.Drawing.Size(120, 20)
        Me.udRoutine.TabIndex = 1
        Me.udRoutine.Tag = "R"
        Me.udRoutine.Text = "DomainUpDown1"
        Me.udRoutine.Wrap = True
        '
        'Label75
        '
        Me.Label75.Location = New System.Drawing.Point(16, 8)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(128, 23)
        Me.Label75.TabIndex = 0
        Me.Label75.Text = "Routine Collections"
        '
        'TabPageCertificates
        '
        Me.TabPageCertificates.Controls.Add(Me.Label30)
        Me.TabPageCertificates.Location = New System.Drawing.Point(4, 4)
        Me.TabPageCertificates.Name = "TabPageCertificates"
        Me.TabPageCertificates.Size = New System.Drawing.Size(753, 446)
        Me.TabPageCertificates.TabIndex = 2
        Me.TabPageCertificates.Text = "Certificates"
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label30.Location = New System.Drawing.Point(24, 8)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(320, 23)
        Me.Label30.TabIndex = 8
        Me.Label30.Text = "Double-click Row to Add/Edit, Ctrl-delete to remove option"
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.Label32)
        Me.TabPage6.Location = New System.Drawing.Point(4, 4)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(753, 446)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Collections"
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label32.Location = New System.Drawing.Point(24, 8)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(320, 23)
        Me.Label32.TabIndex = 8
        Me.Label32.Text = "Double Click Row to Edit/Add Ctrl-Delete to remove option"
        '
        'tabPageAccounts
        '
        Me.tabPageAccounts.Controls.Add(Me.Label28)
        Me.tabPageAccounts.Location = New System.Drawing.Point(4, 4)
        Me.tabPageAccounts.Name = "tabPageAccounts"
        Me.tabPageAccounts.Size = New System.Drawing.Size(753, 446)
        Me.tabPageAccounts.TabIndex = 0
        Me.tabPageAccounts.Text = "Accounts"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label28.Location = New System.Drawing.Point(24, 8)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(320, 23)
        Me.Label28.TabIndex = 8
        Me.Label28.Text = "Double Click Row to Edit/Add Ctrl-Delete to remove option"
        '
        'TabPageReporting
        '
        Me.TabPageReporting.Location = New System.Drawing.Point(4, 4)
        Me.TabPageReporting.Name = "TabPageReporting"
        Me.TabPageReporting.Size = New System.Drawing.Size(753, 446)
        Me.TabPageReporting.TabIndex = 6
        Me.TabPageReporting.Text = "Reporting"
        '
        'ckFerment
        '
        Me.ckFerment.Location = New System.Drawing.Point(20, 20)
        Me.ckFerment.Name = "ckFerment"
        Me.ckFerment.Size = New System.Drawing.Size(300, 24)
        Me.ckFerment.TabIndex = 8
        Me.ckFerment.Text = "&Report Fermentation Alcohol as Negative"
        '
        'tabPageSampleHandling
        '
        Me.tabPageSampleHandling.Controls.Add(Me.Label31)
        Me.tabPageSampleHandling.Location = New System.Drawing.Point(4, 4)
        Me.tabPageSampleHandling.Name = "tabPageSampleHandling"
        Me.tabPageSampleHandling.Size = New System.Drawing.Size(753, 446)
        Me.tabPageSampleHandling.TabIndex = 4
        Me.tabPageSampleHandling.Text = "Sample Handling"
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label31.Location = New System.Drawing.Point(16, 8)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(320, 23)
        Me.Label31.TabIndex = 8
        Me.Label31.Text = "Double Click Row to Edit/Add Ctrl-Delete to remove option"
        '
        'tabPagePricing
        '
        Me.tabPagePricing.Controls.Add(Me.Label29)
        Me.tabPagePricing.Location = New System.Drawing.Point(4, 4)
        Me.tabPagePricing.Name = "tabPagePricing"
        Me.tabPagePricing.Size = New System.Drawing.Size(753, 446)
        Me.tabPagePricing.TabIndex = 5
        Me.tabPagePricing.Text = "Pricing"
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label29.Location = New System.Drawing.Point(24, 8)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(320, 23)
        Me.Label29.TabIndex = 8
        Me.Label29.Text = "Double Click Row to Edit/Add Ctrl-Delete to remove option"
        '
        'tabPageMisc
        '
        Me.tabPageMisc.Location = New System.Drawing.Point(4, 4)
        Me.tabPageMisc.Name = "tabPageMisc"
        Me.tabPageMisc.Size = New System.Drawing.Size(753, 446)
        Me.tabPageMisc.TabIndex = 6
        Me.tabPageMisc.Text = "Misc"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label27.Location = New System.Drawing.Point(160, 8)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(320, 23)
        Me.Label27.TabIndex = 7
        Me.Label27.Text = "Double Click Row to Edit/Add Ctrl-Delete to remove option"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(24, 304)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(100, 23)
        Me.Label26.TabIndex = 5
        Me.Label26.Text = "Screen and Hold"
        '
        'lblContractor
        '
        Me.lblContractor.Location = New System.Drawing.Point(24, 152)
        Me.lblContractor.Name = "lblContractor"
        Me.lblContractor.Size = New System.Drawing.Size(100, 23)
        Me.lblContractor.TabIndex = 3
        Me.lblContractor.Text = "Contractor"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(32, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 23)
        Me.Label16.TabIndex = 0
        Me.Label16.Text = "Customer Type"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(77, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "Analysis Only "
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(32, 37)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(528, 23)
        Me.Label11.TabIndex = 47
        Me.Label11.Text = "There will be a site handling charge in addition this can vary from site to site " & _
            "and is paid to the site"
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(80, 11)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(100, 23)
        Me.Label44.TabIndex = 45
        Me.Label44.Text = "External"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblNoEdit)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.cmdServices)
        Me.Panel1.Controls.Add(Me.cmdCustomerOptions)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 613)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(928, 40)
        Me.Panel1.TabIndex = 1
        '
        'lblNoEdit
        '
        Me.lblNoEdit.Location = New System.Drawing.Point(264, 8)
        Me.lblNoEdit.Name = "lblNoEdit"
        Me.lblNoEdit.Size = New System.Drawing.Size(100, 23)
        Me.lblNoEdit.TabIndex = 71
        Me.lblNoEdit.Text = "Can't Edit Prices"
        Me.lblNoEdit.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(216, 0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 70
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'cmdServices
        '
        Me.cmdServices.Location = New System.Drawing.Point(128, 8)
        Me.cmdServices.Name = "cmdServices"
        Me.cmdServices.Size = New System.Drawing.Size(75, 23)
        Me.cmdServices.TabIndex = 15
        Me.cmdServices.Text = "Services"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(836, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(748, 5)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(75, 23)
        Me.cmdOk.TabIndex = 12
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmCustomerDetails
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(928, 653)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCustomerDetails"
        Me.Text = "Customer Details"
        Me.TabControl1.ResumeLayout(False)
        Me.Customer.ResumeLayout(False)
        Me.Customer.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.tabContacts.ResumeLayout(False)
        Me.pnlContactDetails.ResumeLayout(False)
        Me.pnlContact.ResumeLayout(False)
        Me.AdrContact.ResumeLayout(False)
        Me.AdrContact.PerformLayout()
        Me.tabHistory.ResumeLayout(False)
        Me.pnlHistory.ResumeLayout(False)
        Me.pnlContactList.ResumeLayout(False)
        Me.TabAddresses.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.AddressControl1.ResumeLayout(False)
        Me.AddressControl1.PerformLayout()
        Me.tabPricing.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabTypes.ResumeLayout(False)
        Me.tabpSample.ResumeLayout(False)
        Me.panel2.ResumeLayout(False)
        Me.pnlMRO.ResumeLayout(False)
        Me.pnlMRO.PerformLayout()
        Me.pnlExternal.ResumeLayout(False)
        Me.pnlExternal.PerformLayout()
        Me.pnlFixedSite.ResumeLayout(False)
        Me.pnlFixedSite.PerformLayout()
        Me.tabpMinCharge.ResumeLayout(False)
        Me.pnlMinCharge.ResumeLayout(False)
        Me.tabpAddTime.ResumeLayout(False)
        Me.pnlAddTime.ResumeLayout(False)
        Me.pnlADHeader.ResumeLayout(False)
        Me.pnlADHeader.PerformLayout()
        Me.pnlTimeAllowance.ResumeLayout(False)
        Me.pnlTimeAllowHead.ResumeLayout(False)
        Me.pnlTimeAllowHead.PerformLayout()
        Me.tabpTests.ResumeLayout(False)
        Me.pnlTests.ResumeLayout(False)
        Me.tabpKits.ResumeLayout(False)
        Me.tabpKits.PerformLayout()
        Me.tabpCancellations.ResumeLayout(False)
        Me.pnlCancellations.ResumeLayout(False)
        Me.pnlCancellations.PerformLayout()
        Me.TabDiscSchemes.ResumeLayout(False)
        Me.TabInvoicing.ResumeLayout(False)
        Me.TabInvoicing.PerformLayout()
        Me.panelStd.ResumeLayout(False)
        Me.TabReporting.ResumeLayout(False)
        Me.TabContracts.ResumeLayout(False)
        Me.TabContracts.PerformLayout()
        Me.tabComments.ResumeLayout(False)
        Me.pnlComments.ResumeLayout(False)
        Me.TabTestPanels.ResumeLayout(False)
        Me.pnlNVDScheme.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.pnlMinchargeheader.ResumeLayout(False)
        Me.pnlMinchargeheader.PerformLayout()
        Me.pnlSelectCharging.ResumeLayout(False)
        Me.TabPageCertificates.ResumeLayout(False)
        Me.TabPage6.ResumeLayout(False)
        Me.tabPageAccounts.ResumeLayout(False)
        Me.tabPageSampleHandling.ResumeLayout(False)
        Me.tabPagePricing.ResumeLayout(False)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Client As Intranet.intranet.customerns.Client
    Private myCustComment As MedscreenCommonGui.ListViewItems.lvComment
    Private myVolDiscount As ListViewItems.DiscountSchemes.lvVolumeDiscountHeader
    Private myRoutineDiscount As ListViewItems.DiscountSchemes.lvCustDiscountScheme

    Private blnNew As Boolean
    Private blnForceClose As Boolean = False
    Private blnPriceChanged As Boolean = False          'Indicator that a price has changed

    Private AccTab As Integer
    Private AddressUsed As MedscreenLib.Constants.AddressType = Constants.AddressType.AddressMain

    Private Const cstStandardPricing = 0
    'Private Const cstOldPricing = 2
    Private Const cstAlternatePricing = 1

    Private intPricingOption As Integer = -1
    Private blnDrawing As Boolean = True

    'Private objRP As Intranet.intranet.customerns.RoutinePayment
    Private myCustomerTestPanel As Intranet.intranet.TestPanels.CustomerTestPanel
    Dim objTS As Intranet.intranet.TestResult.TestScheduleEntry
    Private passwordState As ClientReport.PasswordLetters = ClientReport.PasswordLetters.NewPassword
    Private ColumnDirection(20) As Integer
    Private WithEvents myService As Intranet.intranet.customerns.CustomerService

    Private MCActionChar As Char = "I"
    Private COActionChar As Char = "I"
    Private strContractAmendment As String = ""
    'dynamic pricing panels
    Private myPanels As New Panels.DiscPanelCollCollection()

    '<Added code modified 15-Apr-2007 11:39 by taylor> 
    'Constants for panels
    Private Const cstGeneralPanel As Integer = 0
    Private Const cstDiscountSchemes As Integer = 1
    Private Const cstInvoicingPanel As Integer = 2
    Private Const cstAddressesPanel As Integer = 3
    Private Const cstPricingPanel As Integer = 4
    Private Const cstOptionsPanel As Integer = 5
    Private Const cstContractsPanel As Integer = 6
    Private Const cstHistoryPanel As Integer = 7
    Private Const cstCommentsPanel As Integer = 8
    Private Const cstTestPanelsPanel As Integer = 9


    '</Added code modified 15-Apr-2007 11:39 by taylor> 


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub

    'Public WriteOnly Property CustomerID() As Integer
    '    Set(ByVal Value As Integer)
    '        SetCustomer(Value)
    '    End Set
    'End Property

    Private blnClientPanelFilled As Boolean = False

    Public Property cClient() As Intranet.intranet.customerns.Client
        Get
            Return Client
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            Try
                Client = Value
                'initialise the service used to respond to events
                myService = New CustomerService()
                PricingSupport.ChangeDate = Today                           'Initialise price change date
                Me.blnPriceChanged = False                                  'Clear price change flag
                Me.ErrorProvider1.SetError(Me.txtSageID, "")                'Clear errors on sage number
                'Set discount panel controls client property 
                blnClientPanelFilled = False
                LogTiming("Start panel fill")
                Me.pnlAO.Client = Value
                Me.pnlRoutine.Client = Value
                Me.pnlExternal.Client = Value
                Me.pnlFixedSite.Client = Value
                Me.pnlCallOut.Client = Value
                Me.pnlINHoursAddTimeCO.Client = Value
                Me.pnlOutHoursAddTimeCO.Client = Value
                Me.pnlINHoursAddTimeRoutine.Client = Value
                Me.pnlOutHoursAddTimeRoutine.Client = Value
                Me.pnlMRO.Client = Value
                Me.pnlOutHoursMC.Client = Value
                Me.dpMCCallOutIn.Client = Value
                Me.dpMCCallOutOut.Client = Value
                Me.pnlInHourMinCharge.Client = Value
                LogTiming("End Panel Fill")
                If Not Client Is Nothing Then
                    'Me.cbMROID.DataBindings.Clear()
                    'Me.cbMROID.DataBindings.Add("SelectedValue", Client.ReportingInfo, "MROID")
                    'Me.cbSalesPerson.DataBindings.Clear()
                    'Me.cbSalesPerson.DataBindings.Add("SelectedValue", Client, "SalesPerson")
                    'Me.cbProgMan.DataBindings.Clear()
                    'Me.cbProgMan.DataBindings.Add("SelectedValue", Client, "ProgrammeManager")
                End If
                blnDrawing = True
                LogTiming("Start Set customer")
                SetCustomer()
                LogTiming("End set customer")
                blnDrawing = False

                'Start using the template to fill out controls that will chnage function in 2004 release
                Dim objTemplate As MedscreenLib.Template
                Try
                    'Cant 4.0 version templates
                    objTemplate = MedscreenLib.Template.GetTemplate(Template.TemplateType.Med, "CUSTOMER")

                Catch ex As Exception
                    LogError(ex, True, "")
                End Try
                LogTiming("Setup form")
                SetupForm(objTemplate)
                LogTiming("Setup Form finish")
                If Client Is Nothing Then Exit Property

                'Finalise by drawing prices
                DrawPrices()
                'Get the currency in use and tidy display for that currency
                Dim strCurr As String = (5).ToString("c", Client.Culture).Chars(0)
                Me.lblLPCurr.Text = "(" & strCurr & ")"
                Me.lblPCurr.Text = "(" & strCurr & ")"
                Me.lbllpCollCurr.Text = "(" & strCurr & ")"
                Me.lblpCollCurr.Text = "(" & strCurr & ")"
                Me.lbllPADCurr.Text = "(" & strCurr & ")"
                Me.lbladpCurr.Text = "(" & strCurr & ")"
                Me.chListPrice.Text = "List Price (" & strCurr & ")"
                Me.chPrice.Text = "Price (" & strCurr & ")"
                Me.chListPriceKit.Text = "List Price (" & strCurr & ")"
                Me.chPriceKit.Text = "Price (" & strCurr & ")"

                'See if there is an MRO
                If Client.HasMRO Then
                    Me.lblMROComment.Text = "MRO is " & Client.ReportingInfo.MROID
                Else
                    Me.lblMROComment.Text = "No MRO set for this customer"
                End If

                'Collection Time allowance
                LogTiming("SEtup Time allowance")
                SetUpTimeAllowance()

                'Deal with test schedules
                LogTiming("Fill test list view")
                'Me.FillTestListView()

                'Deal with Kit and equipment
                LogTiming("Fill kit list view")
                'Me.FillKitListview()

                LogTiming("Finish")
                blnClientPanelFilled = True
                Dim strAddressMismatch As String = "lib_Customer.SharedMailAddress"
                Dim strXML As String = CConnection.PackageStringList(strAddressMismatch, Client.Identity)
                If Not strXML Is Nothing Then
                    Dim strText As String = Medscreen.ResolveStyleSheet(strXML, "PairedAddresstxt1.xsl", 0)
                    Try
                        Me.txtSameAddress.Rtf = strText
                    Catch
                    End Try
                End If
            Catch ex As Exception
                    LogError(ex, True, )
                End Try
                'dynamic panels
                Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet
                For Each myPnlColl In myPanels
                    Dim BasePanel As Panels.DiscountPanelBase
                    For Each BasePanel In myPnlColl.Panels
                        CType(BasePanel.Panel, Panels.DiscountPanel).Client = Value
                    Next
                Next


        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw out the prices for this customer
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawPrices()
        LogTiming("Draw Prices")
        DoPanelFill()
        LogTiming("Draw Prices Finish")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Go through Form setting it up 
    ''' </summary>
    ''' <param name="Template"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetupForm(ByVal Template As MedscreenLib.Template)
        If Template Is Nothing Then Exit Sub

        Me.Text = Template.Title & " - " & Client.CompanyName
        Dim objTempField As MedscreenLib.Template.TemplateField

        For Each objTempField In Template.Fields
            Dim objControl As Control
            For Each objControl In Me.Controls
                FindControl(objTempField, objControl)
            Next
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Routine to bind contreols and set their visibility state dependent on values in database
    ''' </summary>
    ''' <param name="ControlName"></param>
    ''' <param name="aControl"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [17/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function FindControl(ByVal ControlName As Template.TemplateField, _
    ByVal aControl As Control) As Boolean

        Dim blnReturn As Boolean = False
        Dim aPropertyName As String = ControlName.PropertyName
        Dim clientProperties As System.ComponentModel.PropertyDescriptorCollection = _
        System.ComponentModel.TypeDescriptor.GetProperties(Me.Client)
        'Find details on the property 
        Dim aProperty As PropertyDescriptor = clientProperties.Item(aPropertyName)
        'Debug.WriteLine(Client.VATRegno)
        If aProperty Is Nothing Then Exit Function
        'Adjust property name for dependent properties
        Dim intPos As Integer = InStr(aPropertyName, ".")
        While intPos > 0
            aPropertyName = Mid(aPropertyName, intPos + 1)
            intPos = InStr(aPropertyName, ".")
        End While

        'Look at container controls they may contain the control in their controls so call recursively
        If TypeOf aControl Is Panel Or TypeOf aControl Is TabControl Or TypeOf aControl Is GroupBox Then
            Dim objControl As Control
            For Each objControl In aControl.Controls
                blnReturn = blnReturn Or FindControl(ControlName, objControl)
            Next
            Exit Function
        End If

        'Deal with labels
        If aControl.Name = "lbl" & aPropertyName AndAlso TypeOf aControl Is Label Then
            aControl.Text = ControlName.DisplayName
            aControl.Visible = ControlName.IsDisplay
            aControl.Enabled = Not ControlName.IsReadOnly

            blnReturn = True
        End If
        'Deal with text boxes
        If aControl.Name = "txt" & aPropertyName AndAlso TypeOf aControl Is TextBox Then
            aControl.Visible = ControlName.IsDisplay
            aControl.Enabled = Not ControlName.IsReadOnly

            'Dim atextbox As TextBox = aControl
            aControl.DataBindings.Clear()
            aControl.Text = aProperty.GetValue(Me.Client)
            'aControl.Text = ControlName.CurrentValue
            'If aControl.DataBindings.Count = 0 Then aControl.DataBindings.Add("Text", Client, ControlName.PropertyName)
            blnReturn = True
        End If

        'Deal with Combo boxes
        If aControl.Name = "cb" & aPropertyName Then
            aControl.Visible = ControlName.IsDisplay
            'aControl.DataBindings.Clear()
            'If aControl.DataBindings.Count = 0 Then aControl.DataBindings.Add("SelectedValue", Client, ControlName.PropertyName)

            blnReturn = True
        End If

        'Deal with date time pickers
        If aControl.Name = "dt" & aPropertyName AndAlso TypeOf aControl Is DateTimePicker Then
            aControl.Visible = ControlName.IsDisplay
            Dim dtControl As DateTimePicker = aControl
            aControl.Enabled = Not ControlName.IsReadOnly

            aControl.DataBindings.Clear()
            dtcontrol.Value = aProperty.GetValue(Me.Client)
            'If dtControl.DataBindings.Count = 0 Then dtControl.DataBindings.Add("Value", Client, ControlName.PropertyName)

            blnReturn = True
        End If

        'Deal with check boxes
        If aControl.Name = "ck" & aPropertyName AndAlso TypeOf aControl Is CheckBox Then
            aControl.Visible = ControlName.IsDisplay
            Dim dtControl As CheckBox = aControl
            aControl.DataBindings.Clear()

            dtcontrol.Checked = aProperty.GetValue(Me.Client)
            'If dtControl.DataBindings.Count = 0 Then dtControl.DataBindings.Add("Checked", Client, ControlName.PropertyName)
            dtcontrol.Text = ControlName.DisplayName
            aControl.Enabled = Not ControlName.IsReadOnly
            blnReturn = True
        End If

        'Deal with List views
        If aControl.Name = "lv" & aPropertyName AndAlso TypeOf aControl Is MedscreenCommonGui.OptionsListView Then
            aControl.Visible = ControlName.IsDisplay
            Dim dtControl As MedscreenCommonGui.OptionsListView = aControl
            aControl.DataBindings.Clear()
            'If dtControl.DataBindings.Count = 0 Then dtControl.DataBindings.Add("Checked", Client, ControlName.PropertyName)
            dtcontrol.Text = ControlName.DisplayName
            dtcontrol.ForeColor = Color.LightGray
            dtcontrol.BackColor = Color.White
            aControl.Enabled = Not ControlName.IsReadOnly
            blnReturn = True
        End If


    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if the user has sufficent rights to edit ids 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function CanEditIDs() As Boolean
        'Code modified 11-Apr-2007 06:47 by taylor
        Try  ' Protecting getting ids
            If MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
                Return False
                Exit Function
            End If
            Dim user As String = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity

            Dim blnRet As Boolean = False
            blnRet = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserHasRole(user, "MED_CREATE_CUSTOMER")
            blnRet = blnRet Or MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserHasRole(user, "MED_EDIT_CUSTOMERID")
            Return blnRet
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCustomerDetails-CanEditIDs-4692")
        Finally
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill all of the controls on the form, handing off to particular routines for some tasks
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub SetCustomer()
        Try

            Me.ErrorProvider1.SetError(txtCompanyName, "")
            Me.ErrorProvider1.SetError(Me.cmdOk, "")
            Me.cmdPassword.Enabled = False

            If Not Client Is Nothing Then
                Try
                    Me.Text = "Customer Details - " & Client.CompanyName

                    'check user rights 
                    LogTiming("Enabling")
                    Me.GroupBox2.Enabled = Me.CanEditIDs
                    Me.txtSMID.Enabled = Me.CanEditIDs
                    Me.txtProfile.Enabled = Me.CanEditIDs
                    Me.txtEmailPassword.Enabled = Me.CanEditIDs
                    Me.cmdPassword.Enabled = Me.CanEditIDs
                    LogTiming("Enabling done")

                    'Me.txtCompanyName.Text = Client.CompanyName
                    'Me.txtIdentity.Text = Client.Identity
                    'do new fields
                    Me.txtSMID.Text = Client.SMID
                    'Me.txtProfile.Text = Client.Profile
                    Me.txtComments.Text = Client.Comments
                    Me.txtAccountType.Text = Client.CustomerAccountType

                    LogTiming("Email password")
                    Me.txtEmailPassword.Text = Client.ReportingInfo.EmailPassword
                    Me.ToolTip1.SetToolTip(Me.txtEmailPassword, Client.ReportingInfo.EmailPassword.ToUpper)
                    If Client.EmailResults.Trim.Length > 0 Then 'Can email results
                        If Client.ReportingInfo.EmailPassword.Trim.Length = 0 Then ' no password set 
                            Me.cmdPassword.Text = "Add &Password"
                            passwordState = ClientReport.PasswordLetters.NewPassword
                        Else
                            Me.cmdPassword.Text = "Change &Password"
                            passwordState = ClientReport.PasswordLetters.ChangedPassword
                        End If
                    Else
                        Me.cmdPassword.Enabled = False
                    End If
                    LogTiming("Results Address")
                    'If Not Client.InvoiceInfo.REGNO Is System.DBNull.Value Then Me.txtCustReg.Text = Client.REGNO
                    Me.txtVATRegno.Text = Client.VATRegno
                    Me.txtVATRegno_Validating(Nothing, Nothing)
                    If Not Client.ResultsAddress Is Nothing Then
                        Me.AddressControl1.Address = Client.ResultsAddress
                    Else
                        Me.AddressControl1.Address = Nothing
                        'Me.txtContact.Text = Client.ResultsAddress.Contact
                    End If
                    'flags
                    'Me.ckNewCustomer.Checked = Client.NewCustomer
                    LogTiming("Various controls")
                    Me.CkRemoved.Checked = Client.Removed
                    Me.ckStopCerts.Checked = Client.StopCerts
                    Me.CKStopSamples.Checked = Client.StopSamples
                    Me.ckPositivesOnly.Checked = Client.ReportingInfo.PositivesOnly
                    Me.ckDivertPos.Checked = Client.ReportingInfo.DivertPositives
                    Me.ckEmailPosNeg.Checked = Client.ReportingInfo.PositivesAndNegatives
                    Me.ckEmailWeeklyReport.Checked = Client.ReportingInfo.EmailWeeklyReport
                    Me.ckCreditCard.Checked = Client.InvoiceInfo.CreditCard_pay
                    Me.ckAddTests.Checked = Client.InvoiceInfo.AddTests
                    Me.ckInvalidPanel.Checked = Client.ReportingInfo.AllowInvalidPanel
                    Me.ckReportInBits.Checked = Client.ReportingInfo.Bits
                    Me.ckEmailWeeklyReport.Checked = Client.ReportingInfo.EmailWeeklyReport
                    Me.ckFerment.Checked = Client.ReportingInfo.FermentationNegative

                    'Me.txtParent.Text = Client.Parent
                    Me.cbParent.SelectedValue = Client.Parent
                    'Me.txtSalesPerson.Text = Client.SalesPerson
                    Me.txtContractNumber.Text = Client.ContractNumber
                    Me.cbDivision.SelectedValue = Client.Division
#If r2004 Then
#Else
                    Me.txtSubContOf.Text = Client.SubContractorOf
                    Me.CKSuspended.Checked = Client.AccountSuspended

#End If
                    Me.txtPinNumber.Text = Client.PinNumber
                    'Me.txtProgrammeManager.Text = Client.ReportingInfo.ProgrammeManager

                    Me.cbProgMan.SelectedValue = Me.Client.ProgrammeManager
                    Me.cbSalesPerson.SelectedValue = Me.Client.SalesPerson
                    Me.txtSageID.Text = Client.SageID
                    Try
                        'If CConnection.CompareSchema("1.1.0") < 0 Then
                        '    'use customer table
                        '    Me.txtEmailPositives.Text = Client.ReportingInfo.EmailPositives
                        'Else
                        'use secondary address if present
                        If Not Client.SecondaryAddress Is Nothing Then
                            'set email address
                            Me.txtEmailPositives.Text = Client.SecondaryAddress.Email
                        Else
                            Me.txtEmailPositives.Text = "No Secondary address set."
                        End If
                        'End If
                    Catch ex As Exception
                        LogError(ex, True, )
                    End Try



                    'Me.txtVAT.Text = Client.InvoiceInfo.VATRate

                    Me.ckPORequired.Checked = Client.InvoiceInfo.PoRequired
                    Me.ckPhoneBeforeFaxing.Checked = Client.ReportingInfo.PhoneBeforeFaxing
                    Me.chkEmailAndPrint.Checked = Client.ReportingInfo.EmailAndPrint
                    'Phrases
                    'Account type 
                    Dim I1 As Integer
                    LogTiming("Phrases")
                    Dim objPhrase As MedscreenLib.Glossary.Phrase

                    Me.CBCustomerType.SelectedIndex = -1
                    If Client.AccountCollectionType.Trim.Length > 0 Then
                        I1 = MedscreenLib.Glossary.Glossary.AccountType.IndexOf(Client.AccountCollectionType)
                        Me.CBCustomerType.SelectedIndex = I1
                    End If

                    Try
                        'Industry Code
                        Me.cbIndustry.SelectedIndex = -1
                        If Client.IndustryCode.Trim.Length > 0 Then
                            I1 = MedscreenLib.Glossary.Glossary.Industry.IndexOf(Client.IndustryCode)
                            Me.cbIndustry.SelectedIndex = I1
                        End If
                        'report method
                        If Client.ReportingInfo.ReportType.Trim.Length > 0 Then
                            For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ReportMethod).Count - 1
                                objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ReportMethod).Current()
                                If objPhrase.PhraseID = Client.ReportingInfo.ReportType Then
                                    Exit For
                                End If
                                Me.BindingContext(MedscreenLib.Glossary.Glossary.ReportMethod).Position += 1
                            Next
                        End If
                        'fax format
                        If Client.ReportingInfo.FaxFormat.Trim.Length > 0 Then
                            For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.FaxFormat).Count - 1
                                objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.FaxFormat).Current()
                                If objPhrase.PhraseID = Client.ReportingInfo.FaxFormat Then
                                    Exit For
                                End If
                                Me.BindingContext(MedscreenLib.Glossary.Glossary.FaxFormat).Position += 1
                            Next
                        End If
                        'invoice frequency
                        If Client.InvoiceInfo.InvFrequency.Trim.Length > 0 Then
                            Me.cbInvFreq.SelectedValue = Client.InvoiceInfo.InvFrequency
                        End If
                        'Platinum Account Type 
                        Me.cbPlatAccType.SelectedIndex = -1
                        If Client.InvoiceInfo.PlatAccType.Trim.Length > 0 Then
                            For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Count - 1
                                objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Current()
                                If objPhrase.PhraseID = Client.InvoiceInfo.PlatAccType Then
                                    Exit For
                                End If
                                Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Position += 1
                            Next
                        End If
                        Me.cbReportMethod.SelectedIndex = -1
                        'Do reporting method
                        Dim strPhrase As String = Client.ReportingInfo.ReportType
                        objPhrase = MedscreenLib.Glossary.Glossary.ReportMethod.Item(strPhrase)
                        'Empty controls
                        Me.ErrorProvider1.SetError(Me.cbReportMethod, "")
                        Me.AddressControl1.SetErrorText(MedscreenCommonGui.AddressPanel.HighlightedControl.Fax, "")
                        Me.AddressControl1.SetErrorText(MedscreenCommonGui.AddressPanel.HighlightedControl.Email, "")
                        Me.lblEmailError.Text = ""
                        Me.lblFaxError.Text = ""
                        Me.AddressControl1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
                        'See what we have got 
                        If objPhrase Is Nothing Then 'Error should be set 
                            Me.ErrorProvider1.SetError(Me.cbReportMethod, "Reporting method needs to be set")
                        Else
                            Me.lblReportingMethod.Text = "Results by " & objPhrase.PhraseText
                            I1 = MedscreenLib.Glossary.Glossary.ReportMethod.IndexOf(strPhrase)
                            Me.cbReportMethod.SelectedIndex = I1
                            'check we match address data.

                            If strPhrase = "F" Then ' A fax is required 
                                If Not IsNumber(Client.ResultsAddress.Fax) Then
                                    Me.AddressControl1.SetErrorText(MedscreenCommonGui.AddressPanel.HighlightedControl.Fax, "We have no fax number")
                                End If
                                Me.lblFaxError.Text = "*"
                                Me.AddressControl1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.Fax
                            ElseIf strPhrase = "E" Then
                                If Not Medscreen.ValidateEmail(Client.ResultsAddress.Email) = Constants.EmailErrors.None Then
                                    Me.AddressControl1.SetErrorText(MedscreenCommonGui.AddressPanel.HighlightedControl.Email, "There is an error in the email address provided or it is missing")

                                End If
                                Me.lblEmailError.Text = "*"
                                Me.AddressControl1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.Email
                            ElseIf strPhrase = "P" Then
                                Me.AddressControl1.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.Line1
                            End If
                        End If
                        'Fax format 
                        strPhrase = Client.ReportingInfo.FaxFormat
                        objPhrase = MedscreenLib.Glossary.Glossary.FaxFormat.Item(strPhrase)
                        Me.ErrorProvider1.SetError(Me.cbFaxFormat, "")
                        If objPhrase Is Nothing Then 'Error should be set 
                            Me.ErrorProvider1.SetError(Me.cbFaxFormat, "Fax Format needs to be set")
                        Else
                            I1 = MedscreenLib.Glossary.Glossary.FaxFormat.IndexOf(strPhrase)
                            Me.cbFaxFormat.SelectedIndex = I1
                        End If
                        'Me.dtContractDate.Value = Client.ContractDate
                        myParents = New MedscreenLib.Glossary.PhraseCollection("Select identity,smid||'-'||profile as SMIDProfile from customer", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query, "SMID = '" & Client.SMID & "'")
                        myParents.Add(New MedscreenLib.Glossary.Phrase("", ""))
                        cbParent.DataSource = myParents
                        cbParent.DisplayMember = "PhraseText"
                        cbParent.ValueMember = "PhraseID"
                        Me.cbParent.SelectedValue = Client.Parent
                    Catch ex As Exception
                        LogError(ex, True, )
                    End Try

                    Try

                        Me.cbCurrency.SelectedIndex = -1
                        'Currency 
                        'Fix odd currencies
                        If Client.Currency = "G" Then Client.Currency = "S"
                        If Client.Currency = "U" Then Client.Currency = "D"

                        Dim i As Integer
                        Dim objCurr As MedscreenLib.Glossary.Currency
                        Try  ' Protecting
                            For i = 0 To MedscreenLib.Glossary.Glossary.Currencies.Count - 1
                                objCurr = MedscreenLib.Glossary.Glossary.Currencies.Item(i)
                                'Deal with Single Char
                                If Client.Currency.Trim.Length = 1 AndAlso objCurr.SingleCharId = Client.Currency Then
                                    Me.cbCurrency.SelectedIndex = i
                                    Exit For
                                End If
                                If Client.Currency.Trim.Length > 2 AndAlso objCurr.CurrencyId = Client.Currency Then
                                    Me.cbCurrency.SelectedIndex = i
                                    Exit For
                                End If
                                'Deal with ISO ID

                            Next
                        Catch ex As Exception
                            MedscreenLib.Medscreen.LogError(ex, True, "frmCustomerDetails-SetCustomer-4946 currency pos " & i)
                        Finally
                        End Try
                        LogTiming("Status")
                        Try
                            Me.cbStatus.SelectedValue = Client.AccStatus
                        Catch
                        End Try

                        'Pricing 

                        Me.cbPricing.SelectedIndex = -1
                        Me.cbPricing.SelectedValue = Client.PriceBook

                        'If Client.StandardPricing Then
                        '    Me.cbPricing.SelectedIndex = Me.cstStandardPricing
                        '    'FillStandardPricing()
                        'ElseIf Client.AlternatePricing Then
                        '    Me.cbPricing.SelectedIndex = Me.cstAlternatePricing
                        '    'FillDiscountedPricing()
                        '    'ElseIf Client.OldPricing Then
                        '    ' Me.cbPricing.SelectedIndex = Me.cstOldPricing
                        '    'FillOldPricing()
                        'End If
                        LogTiming("Pricing")
                        Me.intPricingOption = Me.cbPricing.SelectedIndex
                        Me.drawContactList()
                        'deal with MROS 
                        Me.cbMROID.SelectedValue = Me.Client.ReportingInfo.MROID
                        'SetMinChargePanels(actionChar, Tag)
                        'Code modified 10-Apr-2007 14:37 by taylor
                        Try  ' Protecting load document which is giving an error
                            Me.HtmlEditor3.DocumentText = ("<HTML><BODY>No Info</BODY></HTML>")
                        Catch
                        Finally
                        End Try

                    Catch ex As Exception
                        LogError(ex, True, )
                    End Try



                    Me.cbPricing.SelectedIndex = -1
                    Me.cbPricing.SelectedValue = Client.PriceBook

                    'If Client.StandardPricing Then
                    '    Me.cbPricing.SelectedIndex = Me.cstStandardPricing
                    '    'FillStandardPricing()
                    'ElseIf Client.AlternatePricing Then
                    '    Me.cbPricing.SelectedIndex = Me.cstAlternatePricing
                    '    'FillDiscountedPricing()
                    '    'ElseIf Client.OldPricing Then
                    '    ' Me.cbPricing.SelectedIndex = Me.cstOldPricing
                    '    'FillOldPricing()
                    'End If
                    LogTiming("Pricing")
                    Me.intPricingOption = Me.cbPricing.SelectedIndex
                    FillPrices()
                    'DrawOptions()
                    Me.DrawPayments()
                    'Me.DrawDiscounts()
                    'Me.DrawVolumeDiscounts()
                    Me.DrawPrices()
                    Me.CheckMinCharges()
                    LogTiming("Test PAnels")
                    Me.DrawTestPanels()
                    LogTiming("Comments")
                    Me.DrawComments()
                    LogTiming("History")
                    Me.DrawHistory()
                    LogTiming("Goldmine info")
                    'See if we have info from Goldmine
                    If Me.Client.GoldMineId.Trim.Length = 0 Then
                        Me.rbGoldmine.Hide()
                    Else
                        Me.rbGoldmine.Show()
                    End If
                Catch ex As Exception
                    LogError(ex, True, "Fill customer form")
                End Try
            Else
                Me.txtIdentity.Text = "NEW"
                Me.txtCompanyName.Text = "New Client"
                'End If
            End If
            Me.AddressUsed = 1
            Me.RBAddress.Checked = True
            SetControlHighlighted(Me.RBAddress)
        Catch ex As Exception
            LogError(ex, True, )
        End Try


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check that the minimum charges are correct
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/05/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub CheckMinCharges()
        'DEal with setting up minimum charges but only for discounted pricing 
        If Not Client.HasStdPricing Then
            Me.udCallOut.Enabled = True
            Me.udRoutine.Enabled = True
            If Client.HasService("ROUTINE") Then
                'Dim tag As String = "R"
                'Dim actionChar As Char
                Dim RetVal As customerns.Client.MinChargePricingBy = Client.RoutineMinChargeMethod(Me.pnlInHourMinCharge.LineItem, Me.dpRoutineCollCharge.LineItem)

                If RetVal = customerns.Client.MinChargePricingBy.SamplesOnly Then 'Sample based
                    Me.udRoutine.SelectedIndex = 0
                    MCActionChar = "S"
                ElseIf RetVal = customerns.Client.MinChargePricingBy.AttendenceFee Then
                    Me.udRoutine.SelectedIndex = 1
                    MCActionChar = "A"
                ElseIf RetVal = customerns.Client.MinChargePricingBy.MinimumCharge Then
                    Me.udRoutine.SelectedIndex = 2
                    MCActionChar = "M"
                Else
                    Me.udRoutine.SelectedIndex = -1
                    MsgBox("Routine Minimum charge and Attendance Fee pricing needs to be fixed", MsgBoxStyle.Exclamation)
                End If
            Else
                Me.udRoutine.SelectedIndex = 3
                Me.udRoutine.Enabled = False
            End If
            'see if customer has the call out sevice so we need to worry about CallOut pricing 
            If Client.HasService("CALLOUT") Then
                Tag = "CO"
                Dim RetVal As customerns.Client.MinChargePricingBy = Client.CallOutMinChargeMethod(Me.dpMCCallOutIn.LineItem, Me.dpCallOutCollCharge.LineItem)
                If RetVal = customerns.Client.MinChargePricingBy.SamplesOnly Then 'Sample based
                    Me.udCallOut.SelectedIndex = 0
                    COActionChar = "S"
                ElseIf retval = customerns.Client.MinChargePricingBy.AttendenceFee Then
                    Me.udCallOut.SelectedIndex = 1
                    COActionChar = "A"
                ElseIf retval = customerns.Client.MinChargePricingBy.MinimumCharge Then
                    Me.udCallOut.SelectedIndex = 2
                    COActionChar = "M"
                Else
                    Me.udCallOut.SelectedIndex = -1
                    MsgBox("Call out Minimum charge and Attendance Fee pricing needs to be fixed", MsgBoxStyle.Exclamation)

                End If
            Else
                Me.udCallOut.SelectedIndex = 3
                Me.udCallOut.Enabled = False
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw the payments
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawPayments()
        Me.cmdDeleteContract.Hide()

        Dim objRp As Intranet.intranet.customerns.RoutinePayment
        blnUpdateingContract = False
        Me.txtAnnualPrice.Text = ""
        Me.txtNextPayment.Text = ""
        Me.txtContractNumber1.Text = ""
        Me.DTContractStart.Value = Now
        Me.dtFirstPayment.Value = Now

        'Me.Client.InvoiceInfo.RoutinePayments.Clear()

        Me.lvRoutinePayments.Items.Clear()
        For Each objRp In Me.Client.InvoiceInfo.RoutinePayments
            Dim lvit As LVRoutinePayment = New LVRoutinePayment(objRp)
            Me.lvRoutinePayments.Items.Add(lvit)
        Next
    End Sub


    ' ********************************************************************************
    '  Routine     :  DrawVolumeDiscounts [Sub]
    '  Created by  :  taylor, at 19/01/2007-08:36:05 on machine ANDREW
    '  Description :  [type_description_here]
    '  Parameters  :  
    ' ********************************************************************************
    ''' <summary>
    ''' taylor 19/01/2007 08:40:13
    '''     Deal with volume discounts
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    'Private Sub DrawVolumeDiscounts()
    '    Me.lvVolumeDiscounts.Items.Clear()
    '    Dim objCustDisnt As Intranet.intranet.customerns.CustomerVolumeDiscount

    '    'Ensure data is refreshed before being drawn
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Clear()
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Load()
    '    'Go through all customer volume discounts
    '    For Each objCustDisnt In Me.Client.InvoiceInfo.VolumeDiscountSchemes
    '        Dim objlvCitem As MedscreenCommonGui.ListViewItems.DiscountSchemes.lvVolumeDiscountHeader
    '        objlvCitem = New ListViewItems.DiscountSchemes.lvVolumeDiscountHeader(objCustDisnt)

    '        'Add header to list view
    '        'Me.lvVolumeDiscounts.Items.Add(objlvCitem)
    '        Dim objEntry As MedscreenLib.Glossary.VolumeDiscountEntry

    '        'Go through each entry and add it
    '        For Each objEntry In objCustDisnt.DiscountScheme.Entries
    '            Dim objLvItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.lvVolumeDiscountEntry(objEntry, Me.Client)
    '            'Me.lvVolumeDiscounts.Items.Add(objLvItem)
    '        Next
    '    Next
    'End Sub


    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:45:10
    '''     Draw the various non volume discounts
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    'Private Sub DrawDiscounts()
    '    Me.lvRoutineDiscount.Items.Clear()
    '    Dim objCustDisnt As Intranet.intranet.customerns.CustomerDiscount

    '    'Go through each discount scheme
    '    For Each objCustDisnt In Me.Client.InvoiceInfo.DiscountSchemes
    '        Dim objlvCitem As MedscreenCommonGui.ListViewItems.DiscountSchemes.lvCustDiscountScheme
    '        objlvCitem = New ListViewItems.DiscountSchemes.lvCustDiscountScheme(objCustDisnt)
    '        Me.lvRoutineDiscount.Items.Add(objlvCitem)
    '        Dim objEntry As MedscreenLib.Glossary.DiscountSchemeEntry

    '        'Go through each entry
    '        For Each objEntry In objCustDisnt.DiscountScheme.Entries
    '            'Create line item 
    '            Dim objLvItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.lvDiscountScheme(objEntry, Me.Client)
    '            Me.lvRoutineDiscount.Items.Add(objLvItem)
    '        Next
    '    Next
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw out any test panels 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawTestPanels()
        Me.lvTestPanels.Items.Clear()
        Dim objTP As TestPanels.CustomerTestPanel

        myCustomerTestPanel = Nothing               'Clear item out 
        For Each objTP In Me.Client.TestPanels
            Dim lvCTP As New MedscreenCommonGui.ListViewItems.lvCustomerTestPanel(objTP)
            Me.lvTestPanels.Items.Add(lvCTP)
        Next
        If Me.Client.TestPanels.Count > 0 Then
            myCustomerTestPanel = Me.Client.TestPanels.Item(0)
        End If
        Me.DrawTestSchedules()
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:47:00
    '''     Fill out test schedule information 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub DrawTestSchedules()
        Me.lvTestSchedules.Items.Clear()
        If Me.Client.TestSchedule Is Nothing Then Exit Sub
        Me.cmdAddTest.Show()

        'Me.Client.sc()
        'No test schedules then show create button
        If Not myCustomerTestPanel Is Nothing AndAlso myCustomerTestPanel.TestSchedules.Count = 0 Then
            Me.CmdCreateTestPanels.Show()
        End If

        Me.lvTestSchedules.Show()
        Me.lvTestSchedules.Items.Clear()

        'For each test schedule add it to the list view 
        For Each objTS In Me.Client.TestSchedule
            Dim objLvTSched As New MedscreenCommonGui.ListViewItems.lvTestSchedule(objTS)
            Me.lvTestSchedules.Items.Add(objLvTSched)
        Next
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:48:35
    '''     Deal with Routine payments 
    ''' </summary>
    ''' <param name="objRP" type="Intranet.intranet.customerns.RoutinePayment">
    '''     <para>
    '''         Routine payment (Retainer) to draw
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub FillPayments(ByVal objRP As Intranet.intranet.customerns.RoutinePayment)
        Me.txtNextPayment.Text = ""
        Me.txtAnnualPrice.Text = ""
        If objRP Is Nothing Then Exit Sub
        With objRP
            .Load()
            Me.txtContractNumber1.Text = .ContractNumber 'Contract ID
            Me.DTContractStart.Value = .ContractStart   'start date for contract period
            Me.DTContractEndDate.Value = .ContractEnd   'End date for contract period
            Me.txtAnnualPrice.Text = .AnnualPrice       'Amount charged
            'Me.DTNextDue.Value = .NextPayment           'Date for next payment
            Me.ckContractRemoved.Checked = .Removed     'Has this contract been removed 
            Me.dtFirstPayment.Value = .FirstPayment     'Date of first payment
            'output next payment in customers culture
            Me.txtNextPayment.Text = .IntervalPrice.ToString("c", Me.Client.Culture)
            Dim objPhrase As MedscreenLib.Glossary.Phrase = _
                MedscreenLib.Glossary.Glossary.PaymentIntervals.Item(.InvoiceFrequency)
            If Not objPhrase Is Nothing Then
                Dim intPos As Integer = MedscreenLib.Glossary.Glossary.PaymentIntervals.IndexOf(.InvoiceFrequency)
                Me.cbPayInterval.SelectedItem = objPhrase
                Me.cbPayInterval.SelectedIndex = intPos

                'Deal with payment intervals
                Select Case objPhrase.PhraseID
                    Case "ANN"
                        Me.DTNextDue.Value = .FirstPayment.AddYears(1)
                    Case "WEEK"
                        Me.DTNextDue.Value = .FirstPayment.AddDays(7)
                    Case "FORT"
                        Me.DTNextDue.Value = .FirstPayment.AddDays(14)
                    Case "QTR"
                        Me.DTNextDue.Value = .FirstPayment.AddMonths(3)
                    Case "MON"
                        Me.DTNextDue.Value = .FirstPayment.AddMonths(1)

                End Select
            End If
            If .NextPayment > Me.DTNextDue.Value Then
                Me.DTNextDue.Value = .NextPayment
            End If
            Me.txtItemComment.Text = .ItemComment
            'Fill out the line items, stored in config file
            Dim I As Integer
            Me.cbLineItems.SelectedIndex = -1
            For I = 0 To CCToolAppMessages.RoutineLineItems.Keys.Count - 1
                If CCToolAppMessages.RoutineLineItems.Keys.Item(I) = objRP.LineItem Then
                    Me.cbLineItems.SelectedIndex = I
                    Exit For
                End If
            Next

            'Get the payment history
            Dim strXML As String = objRP.ToXmlWithHistory
            Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "CustPricing2.xsl", 0)
            Me.HtmlEditor3.DocumentText = (strHTML)
            Me.DTContractEndDate.Value = .ContractEnd   'End date for contract period
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(100)
        End With
    End Sub

    Private blnOptionsDraw As Boolean = False

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:51:29
    '''     Deal with customer options
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    ''' 
    Private Sub DrawOptions()
        'Sub routine to display the options and to show which ones selected for theis customer
        'Start With Accounts
        'Dim intAuthLevel As Integer = 0
        If blnOptionsDraw Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        blnOptionsDraw = True
        LogTiming("Draw Options")
        'If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
        '    intAuthLevel = MedscreenLib.Glossary.Glossary.CurrentSMUser.OptionAuthority
        'End If
        'Code modified 11-Apr-2007 06:53 by taylor
        Try  ' Protecting
            Dim i As Integer
            For i = 1 To Me.objOptionTabs.Count
                Dim lvoptList As MedscreenCommonGui.OptionsListView = Me.objOptionTabs.Item(i)
                LogTiming("  - Draw Options - Get List view" & lvoptList.OptionClass)
                lvoptList.ClientOptions = Me.Client.Options
                LogTiming("  - Draw Options - Set Client Options" & lvoptList.OptionClass)
                lvoptList.MYclientID = Me.Client.Identity
                LogTiming("  - Draw Options - Set Client " & lvoptList.OptionClass)
                lvoptList.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                LogTiming("  - Draw Options - Set user id" & lvoptList.OptionClass)
                lvoptList.DrawOptionListView()
                LogTiming("  - Draw Options - " & lvoptList.OptionClass)
            Next
            'LogTiming("  - Draw Options - Accounts")
            'Me.lvAccounts.ClientOptions = Me.Client.Options
            'Me.lvAccounts.OptionClass = "ACCOUNTS"
            'Me.lvAccounts.MYclientID = Me.Client.Identity
            'Me.lvAccounts.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvAccounts.DrawOptionListView()

            'LogTiming("  - Draw Options - Certificates")
            'Deal with certificate result based options
            'Me.lvCertificates.ClientOptions = Me.Client.Options
            'Me.lvCertificates.OptionClass = "CERTIFICATES"
            'Me.lvCertificates.MYclientID = Me.Client.Identity
            'Me.lvCertificates.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvCertificates.DrawOptionListView()

            'LogTiming("  - Draw Options - Pricing")

            'Deal with pricing related options 
            'Me.LvPricing.ClientOptions = Me.Client.Options
            'Me.LvPricing.OptionClass = "PRICING"
            'Me.LvPricing.MYclientID = Me.Client.Identity
            'Me.LvPricing.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.LvPricing.DrawOptionListView()
            'Me.DrawOptionListView("ACCOUNTS", Me.lvAccounts)
            'Me.DrawOptionListView("CERTIFICATES", Me.lvCertificates)
            '        Me.DrawOptionListView("PRICING", Me.LvPricing)

            'Deal with customer type options
            'LogTiming("  - Draw Options - CustomerType")

            'Me.lvCustomerType.ClientOptions = Me.Client.Options
            'Me.lvCustomerType.OptionClass = "CUSTOMERTYPE"
            'Me.lvCustomerType.MYclientID = Me.Client.Identity
            'Me.lvCustomerType.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvCustomerType.DrawOptionListView()
            'Me.DrawOptionListView("CUSTOMERTYPE", Me.lvCustomerType)

            'Deal with contractor options
            'LogTiming("  - Draw Options - Contractor")

            'Me.lvContractor.ClientOptions = Me.Client.Options
            'Me.lvContractor.OptionClass = "CONTRACTOR"
            'Me.lvContractor.MYclientID = Me.Client.Identity
            'Me.lvContractor.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvContractor.DrawOptionListView()
            'Me.DrawOptionListView("CONTRACTOR", Me.lvContractor)

            'Deal with Prison options
            'LogTiming("  - Draw Options - HMP")

            'Me.lvPrisons.ClientOptions = Me.Client.Options
            'Me.lvPrisons.OptionClass = "HMP"
            'Me.lvPrisons.MYclientID = Me.Client.Identity
            'Me.lvPrisons.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvPrisons.DrawOptionListView()
            'Me.DrawOptionListView("PRISONS", Me.lvPrisons)

            'Deal with collection options
            LogTiming("  - Draw Options - Collections")

            'Me.lvReporting.ClientOptions = Me.Client.Options
            'Me.lvReporting.OptionClass = "COLLECTIONS"
            'Me.lvReporting.MYclientID = Me.Client.Identity
            'Me.lvReporting.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvReporting.DrawOptionListView()

            'Deal with sample handling options
            'Me.DrawOptionListView("REPORTING", Me.lvReporting)
            LogTiming("  - Draw Options - Handling")

            'Me.lvSamples.ClientOptions = Me.Client.Options
            'Me.lvSamples.OptionClass = "HANDLING"
            'Me.lvSamples.MYclientID = Me.Client.Identity
            'Me.lvSamples.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvSamples.DrawOptionListView()

            'Deal with reporting options 
            '           Me.lvRepOptions.ClientOptions = Me.Client.Options
            LogTiming("  - Draw Options - Reporting")

            'Me.lvRepOptions.OptionClass = "Reporting"
            'Me.lvRepOptions.MYclientID = Me.Client.Identity
            'Me.lvRepOptions.UserID = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            'Me.lvRepOptions.DrawOptionListView()

            Me.ckStopCerts.Checked = Client.StopCerts
            Me.CKStopSamples.Checked = Client.StopSamples
            Me.ckInvalidPanel.Checked = Client.ReportingInfo.AllowInvalidPanel

            'Me.DrawOptionListView("SAMPLES", Me.lvSamples)
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "frmCustomerDetails-DrawOptions-5348")
        Finally
            Cursor = Cursors.Default
        End Try

    End Sub

    'Private Sub DrawOptionListView(ByVal OptionClass As String, ByVal myListView As ListView)
    '    myListView.Items.Clear()
    '    Dim aSerrRow As Groups._OptionRow
    '    For Each aSerrRow In ModTpanel.OptionGroups(OptionClass)
    '        Dim optStr As String = aSerrRow.Option_Text
    '        Dim oOpt As Intranet.intranet.glossary.Options = ModTpanel.Options.Item(optStr)
    '        If Not oOpt Is Nothing Then

    '            Dim ocOpt As Intranet.intranet.customerns.SmClientOption = Me.Client.Options.Item(optStr)
    '            Dim lvItem As LvOptItem = New LvOptItem(oOpt, ocOpt)
    '            myListView.Items.Add(lvItem)
    '        End If
    '    Next

    'End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:54:48
    '''     Show the account history
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub DrawHistory()
        'Dim objComm As Intranet.intranet.customerns.ClientHistory
        'Me.lvHistory.Items.Clear()
        'For Each objComm In Me.Client.History
        '    Me.lvHistory.Items.Add(New MedscreenCommonGui.ListViewItems.LVCustSelectItem(objComm))
        'Next
        Try
            Dim strHistory As String = Me.Client.HistoryXML
            Dim strXML As String = Medscreen.ResolveStyleSheet(strHistory, "custhistory.xsl", 0)
            Me.HtmlEditor2.DocumentText = (strXML)
        Catch
        End Try
    End Sub

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
        For Each objComm In Me.Client.CustomerComments
            Me.lvComments.Items.Add(New MedscreenCommonGui.ListViewItems.lvComment(objComm))
        Next
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 08:58:13
    '''     Draw out pricing info
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub FillPrices()
        LogTiming("   Fill Prices - Start")

        Try

            Me.lbServices.Items.Clear()
            If Client.Services.Count > 0 Then
                Dim I As Integer
                Dim objService As Intranet.intranet.customerns.CustomerService
                For I = 0 To Client.Services.Count - 1
                    objService = Client.Services.Item(I)
                    If Not objService.Removed And _
                     (objService.EndDate > Now Or objService.EndDate = DateField.ZeroDate) Then
                        Me.lbServices.Items.Add(objService.Service)
                    End If
                Next
            End If
        Catch
        End Try

        LogTiming("   Services done")
        'Client.InvoiceInfo.ResetTimeStamp()

        'Draw the XML price list
        Try
            Dim strHeader As String = "<Customer><CustomerHeader>"
            Dim strFormat As String = "<format>" & CConnection.PackageStringList("Lib_customer.GetCustomerFormat", Me.Client.Identity) & "</format>"

            Dim strFooter As String = "</CustomerHeader></Customer>"
            Dim strXML As String = Client.InvoiceInfo.PriceListXML
            strXML = strHeader & strFormat & strXML & strFooter
            LogTiming("   XML Got")
            Try
                Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "CustPricingBaser.xsl", 0)
                Try
                    Me.HtmlEditor1.DocumentText = (strHTML)
                Catch
                End Try
            Catch ex As Exception
                LogError(ex, True, "Fill prices - customer form ")
            Finally
                strXML = Nothing
            End Try

        Catch
        End Try
        LogTiming("   XML Resolved")

        'Based on the pricing model fill the panels
        'If Me.Client.InvoiceInfo.PricingMethod = Intranet.intranet.customerns.ClientInvoice.InvoicePricingType.StandardPricing Then
        '    Me.FillStdPanels()
        'Else
        '    Me.FillPanels()
        'End If
    End Sub

    ' <Removed by: taylor at: 19/01/2007-08:59:42 on machine: ANDREW>
    ' 
    ' </Removed by: taylor at: 19/01/2007-08:59:42 on machine: ANDREW>
    'Private Sub FillOldPricing()
    '    Dim i As Integer
    '    Dim objLI As Intranet.intranet.jobs.LineItem
    '    Dim objPrice As Intranet.intranet.jobs.LineItemPrice

    '    Dim Lprice As Double
    '    Dim disc As Double
    '    Dim APrice As Double
    '    For i = 0 To ModTpanel.PriceList.Count - 1
    '        objPrice = ModTpanel.PriceList.Item(i)
    '        If objPrice.Currency = Client.InvoiceInfo.CurrencyCode Then
    '            objLI = ModTpanel.LineItems.Item(objPrice.Lineitem)
    '            Lprice = objPrice.Price
    '            If Not objLI Is Nothing Then
    '                If objLI.ItemCode = "APCOLS002" Then
    '                    If Client.InvoiceInfo.RoutineSampleCharge > 0 Then
    '                        APrice = Client.InvoiceInfo.RoutineSampleCharge
    '                    Else
    '                        APrice = Client.InvoiceInfo.ChargePerSample
    '                    End If
    '                ElseIf objLI.ItemCode = "APANO0001" Then
    '                    APrice = Client.InvoiceInfo.ChargePerSample

    '                ElseIf objLI.ItemCode = "APCOL0008" Then
    '                    APrice = Client.InvoiceInfo.MROFee
    '                ElseIf objLI.ItemCode = "APCOL0001" Then
    '                    APrice = Client.InvoiceInfo.HourlyInHours
    '                ElseIf objLI.ItemCode = "APCOL0002" Then
    '                    APrice = Client.InvoiceInfo.HourlyOutHours
    '                ElseIf objLI.ItemCode = "APCOL0003" Then
    '                    APrice = Client.InvoiceInfo.MinOutHoursCharge
    '                ElseIf objLI.ItemCode = "APCOL0004" Then
    '                    APrice = Client.InvoiceInfo.MinimumCharge
    '                ElseIf objLI.ItemCode = "ATP0001" Then
    '                    APrice = Client.InvoiceInfo.AddCharge
    '                ElseIf objLI.ItemCode = "COLC003" Then
    '                    APrice = Client.InvoiceInfo.OnCallHourlyInHours
    '                ElseIf objLI.ItemCode = "COLC002" Then
    '                    APrice = Client.InvoiceInfo.OnCallSampleCharge
    '                Else
    '                    APrice = Lprice     'Not a special case 
    '                End If
    '                If APrice = 0 Then APrice = Lprice 'No price defined use price list 
    '                If Lprice > 0 Then
    '                    disc = ((Lprice - APrice) / Lprice) * 100
    '                Else
    '                    disc = 0
    '                End If

    '                Dim liItem As ListViewItem = New ListViewItem(objLI.ItemCode)
    '                'Me.ListView1.Items.Add(liItem)
    '                Dim iType As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Type)
    '                liItem.SubItems.Add(iType)
    '                Dim iDesc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Description)
    '                liItem.SubItems.Add(iDesc)
    '                Dim iPrice As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, Lprice)
    '                liItem.SubItems.Add(iPrice)
    '                Dim iDisc As ListViewItem.ListViewSubItem

    '                If disc <> 0 Then
    '                    iDisc = New ListViewItem.ListViewSubItem(liItem, disc)
    '                Else
    '                    iDisc = New ListViewItem.ListViewSubItem(liItem, "-")
    '                End If
    '                liItem.SubItems.Add(iDisc)

    '                Dim iPriceAct As ListViewItem.ListViewSubItem
    '                'If APrice = Lprice Then
    '                '    iPriceAct = New ListViewItem.ListViewSubItem(liItem, "-")
    '                'Else
    '                iPriceAct = New ListViewItem.ListViewSubItem(liItem, APrice)
    '                'End If
    '                liItem.SubItems.Add(iPriceAct)
    '            End If
    '        End If
    '    Next

    '    ListView1.EndUpdate()
    'End Sub

    'Private Sub FillDiscountedPricing()
    '    Dim i As Integer
    '    Dim objLI As Intranet.intranet.jobs.LineItem
    '    Dim objPrice As Intranet.intranet.jobs.LineItemPrice
    '    Dim blnService As Boolean = (Me.Client.Services.Count > 0)

    '    ListView1.BeginUpdate()
    '    ListView1.Items.Clear()
    '    For i = 0 To ModTpanel.PriceList.Count - 1
    '        objPrice = ModTpanel.PriceList.Item(i)
    '        If objPrice.Currency = Client.InvoiceInfo.CurrencyCode Then
    '            objLI = ModTpanel.LineItems.Item(objPrice.Lineitem)
    '            If Not objLI Is Nothing Then
    '                Dim objService As CustomerService = Me.Client.Services.Item(objLI.Service)
    '                If ((blnService And Not objService Is Nothing) Or Not blnService) Then
    '                    'Check for scheme entry type CODE then type GROUP then type TYPE
    '                    'if a match then apply that discount
    '                    Dim tmpDisc As Double = Client.InvoiceInfo.DiscountSchemes.Discount(objLI.ItemCode, objLI.Type)
    '                    Dim liItem As ListViewItem = New ListViewItem(objLI.ItemCode)
    '                    Me.ListView1.Items.Add(liItem)
    '                    Dim iType As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Type)
    '                    liItem.SubItems.Add(iType)
    '                    Dim iDesc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Description)
    '                    liItem.SubItems.Add(iDesc)
    '                    Dim iPrice As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objPrice.Price)
    '                    liItem.SubItems.Add(iPrice)
    '                    Dim iDisc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, tmpDisc)
    '                    liItem.SubItems.Add(iDisc)
    '                    Dim iPriceAct As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objPrice.Price - (objPrice.Price * tmpDisc / 100))
    '                    liItem.SubItems.Add(iPriceAct)
    '                    Dim iPriceser As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Service)
    '                    liItem.SubItems.Add(iPriceser)
    '                End If
    '            End If
    '        End If
    '    Next

    '    ListView1.EndUpdate()

    'End Sub

    'Private Sub FillStandardPricing()
    '    Dim i As Integer
    '    Dim objLI As Intranet.intranet.jobs.LineItem
    '    Dim objPrice As Intranet.intranet.jobs.LineItemPrice
    '    Dim blnService As Boolean = (Me.Client.Services.Count > 0)

    '    ListView1.BeginUpdate()
    '    ListView1.Items.Clear()
    '    For i = 0 To ModTpanel.PriceList.Count - 1
    '        objPrice = ModTpanel.PriceList.Item(i)
    '        If objPrice.Currency = Client.InvoiceInfo.CurrencyCode Then
    '            objLI = ModTpanel.LineItems.Item(objPrice.Lineitem)
    '            Dim objService As CustomerService = Me.Client.Services.Item(objLI.Service)
    '            If Not objLI Is Nothing Then
    '                If ((blnService And Not objService Is Nothing) Or Not blnService) Then
    '                    Dim liItem As ListViewItem = New ListViewItem(objLI.ItemCode)
    '                    Me.ListView1.Items.Add(liItem)
    '                    Dim iType As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Type)
    '                    liItem.SubItems.Add(iType)
    '                    Dim iDesc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Description)
    '                    liItem.SubItems.Add(iDesc)
    '                    Dim iPrice As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objPrice.Price)
    '                    liItem.SubItems.Add(iPrice)
    '                    Dim iDisc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, "-")
    '                    liItem.SubItems.Add(iDisc)
    '                    Dim iPriceAct As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objPrice.Price)
    '                    liItem.SubItems.Add(iPriceAct)
    '                    Dim iPriceSer As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(liItem, objLI.Service)
    '                    liItem.SubItems.Add(iPriceSer)
    '                End If
    '            End If
    '        End If

    '    Next

    '    ListView1.EndUpdate()
    'End Sub

    'Private Sub MyEMAILValidatingCode()
    '    ' Confirm there is text in the control.
    '    If txtEmailResults.Text.Length = 0 Then
    '        'Throw New Exception("Email address is a required field")
    '    Else
    '        ' Confirm that there is a "." and an "@" in the e-mail address.
    '        If txtEmailResults.Text.IndexOf(".") = -1 Or txtEmailResults.Text.IndexOf("@") = -1 Then
    '            Throw New Exception("E-mail address must be valid e-mail address format." + _
    '                  ControlChars.Cr + "For example 'someone@microsoft.com'")
    '        End If
    '    End If
    'End Sub


    Private Sub MyCCValidatingCode()
        ' Confirm there is text in the control.
        If Me.txtCompanyName.Text.Length = 0 Then
            Throw New Exception("Customer Name must be entered")
        Else
        End If
    End Sub

    'Private Sub txtEmailResults_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '    Try
    '        MyEMAILValidatingCode()

    '    Catch ex As Exception
    '        ' Cancel the event and select the text to be corrected by the user.
    '        e.Cancel = True
    '        txtEmailResults.Select(0, txtEmailResults.Text.Length)

    '        ' Set the ErrorProvider error with the text to display. 
    '        Me.ErrorProvider1.SetError(txtEmailResults, ex.Message)
    '    End Try
    'End Sub

    'Private Sub textBox1_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ' If all conditions have been met, clear the error provider of errors.
    '    ErrorProvider1.SetError(txtEmailResults, "")
    'End Sub

    'Private Sub DisplayAccounts()
    '    Dim objPage As TabPage
    '    Dim i As Integer
    '    Dim lAccount As oneStopClasses.Customer.ClientAccount

    '    Try
    '        Me.TabControl2.TabPages.Clear()

    '        For i = 0 To CustAccounts.Count - 1
    '            lAccount = New oneStopClasses.Customer.ClientAccount(CustAccounts.Item(i).Instance, AccountList, CustAccounts.Client)
    '            objPage = New TabPage()
    '            If Mid(lAccount.AccountId, 1, 1) = "<" Then
    '                lAccount.AccountId = GetSageName(Client.Identity)
    '                lAccount.AccountType = GetAccountType(Client.Identity)
    '                lAccount.Update()
    '            End If
    '            If lAccount.AccountType Is DBNull.Value Then
    '                lAccount.AccountType = GetAccountType(Client.Identity)
    '                lAccount.Update()
    '            End If
    '            'End If
    '        Next
    '    Catch ex As Exception
    '    End Try
    'End Sub


    Private Function GetSageName(ByVal strId As String) As String
        Dim datareader As OleDb.OleDbDataReader
        Dim dc As New OleDb.OleDbConnection()
        Dim Comm As New OleDb.OleDbCommand()

        Dim myQuery As String
        Dim myReader As OleDb.OleDbDataReader
        Dim myRet As String

        dc.ConnectionString = "Provider=MSDAORA;Password=live;User ID=live;Data Source=john-live"

        myQuery = "Select Sage_name from customer where identity = '" & strId & "'"

        Comm.CommandText = myQuery
        Comm.Connection = dc

        Try
            dc.Open()
            myReader = Comm.ExecuteReader
            If myReader.Read() Then
                myRet = myReader.GetString(0)
            Else
                myRet = "<New Account>"
            End If


        Catch ex As OleDb.OleDbException
            LogError(ex, True, "Get sage name")
        Finally
            dc.Close()
        End Try
        Return myRet

    End Function

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 09:00:50
    '''     Get the account type
    ''' </summary>
    ''' <param name="strId" type="String">
    '''     <para>
    '''         Customer ID
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     A System.String value...
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Function GetAccountType(ByVal strId As String) As String
        Dim datareader As OleDb.OleDbDataReader
        Dim dc As New OleDb.OleDbConnection()
        Dim Comm As New OleDb.OleDbCommand()

        Dim myQuery As String
        Dim myReader As OleDb.OleDbDataReader
        Dim myRet As String

        dc.ConnectionString = "Provider=MSDAORA;Password=live;User ID=live;Data Source=john-live"

        myQuery = "Select Customer_note from customer where identity = '" & strId & "'"

        Comm.CommandText = myQuery
        Comm.Connection = dc

        Try
            dc.Open()
            myReader = Comm.ExecuteReader
            If myReader.Read() Then
                myRet = myReader.GetString(0)
            Else
                myRet = "F"
            End If


        Catch ex As OleDb.OleDbException
            LogError(ex, , "Get accounttype")
        Finally
            dc.Close()
        End Try
        If myRet = "T" Then
            Return "ANALONLY"
        Else
            Return "COLLANAL"
        End If

    End Function


    'Private Function GetPaymentType(ByVal strId As String) As String
    '    Dim datareader As OleDb.OleDbDataReader
    '    Dim dc As New OleDb.OleDbConnection()
    '    Dim Comm As New OleDb.OleDbCommand()

    '    Dim myQuery As String
    '    Dim myReader As OleDb.OleDbDataReader
    '    Dim cPay As String
    '    Dim invFreq As String

    '    dc.ConnectionString = "Provider=MSDAORA;Password=live;User ID=live;Data Source=john-live"

    '    myQuery = "Select creditcard_pay,invoice_frequency from customer where identity = '" & strId & "'"

    '    Comm.CommandText = myQuery
    '    Comm.Connection = dc

    '    Try
    '        dc.Open()
    '        myReader = Comm.ExecuteReader
    '        If myReader.Read() Then
    '            cPay = myReader.GetString(0)
    '            invFreq = myReader.GetString(1)
    '        Else
    '            invFreq = "N"
    '        End If


    '    Catch ex As OleDb.OleDbException
    '    Finally
    '        dc.Close()
    '    End Try
    '    If invFreq = "N" Then
    '        Return "CCONLY"
    '    Else
    '        If cPay = "T" Then
    '            Return "INVCC"
    '        Else
    '            Return "INVONLY"
    '        End If
    '    End If

    'End Function

    'Private Function GetInvoicingType(ByVal strId As String) As String
    '    Dim datareader As OleDb.OleDbDataReader
    '    Dim dc As New OleDb.OleDbConnection()
    '    Dim Comm As New OleDb.OleDbCommand()

    '    Dim myQuery As String
    '    Dim myReader As OleDb.OleDbDataReader
    '    Dim cPay As String
    '    Dim invFreq As String

    '    dc.ConnectionString = "Provider=MSDAORA;Password=live;User ID=live;Data Source=john-live"

    '    myQuery = "Select creditcard_pay,invoice_frequency from customer where identity = '" & strId & "'"

    '    Comm.CommandText = myQuery
    '    Comm.Connection = dc

    '    Try
    '        dc.Open()
    '        myReader = Comm.ExecuteReader
    '        If myReader.Read() Then
    '            cPay = myReader.GetString(0)
    '            invFreq = myReader.GetString(1)
    '        Else
    '            invFreq = "N"
    '        End If

    '    Catch ex As OleDb.OleDbException
    '    Finally
    '        dc.Close()
    '    End Try
    '    If invFreq = "N" Then
    '        Return "PREPAY"
    '    ElseIf invFreq = "M" Then
    '        Return "MONTH"
    '    ElseIf invFreq = "T" Then
    '        Return "BIMONTH"
    '    Else
    '        Return "JOB"
    '    End If

    'End Function


    'Private Function AddAccount()
    '    CustAccount = AccountList.CreateClientAccount(Client.ClientId, Client.ClientIdentity)
    '    CustAccounts = AccountList.Items(Client)
    '    DisplayAccounts()
    '    Invalidate()
    'End Function


    'Private Sub accControl_NewCustomer() Handles accControl.NewCustomer
    '    AddAccount()
    'End Sub

    'Private Sub accControl_AccountNameChanged(ByVal strAccNAme As String) Handles accControl.AccountNameChanged
    '    TabControl2.TabPages(AccTab).Text = strAccNAme
    '    Invalidate()
    'End Sub

    'Private Sub TabControl2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    AccTab = TabControl2.SelectedIndex
    'End Sub


    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 09:01:53
    '''     Accept the form
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
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub CmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        'see if contact has changed
        CheckForChangedContact()

        Dim blnOk As Boolean = True                     'Set Exit flag 
        Dim intPanel As Integer = Me.cstGeneralPanel

        Client.CompanyName = Me.txtCompanyName.Text
        'Client.REGNO = Me.txtCustReg.Text
        Client.VATRegno = Me.txtVATRegno.Text

        'Deal with new controls

        Me.ErrorProvider1.SetError(Me.txtSMID, "")
        If Me.txtSMID.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtSMID, "You must specify a SMID for this customer")
            Me.txtSMID.Focus()
            blnOk = False
            intPanel = cstGeneralPanel
        End If

        If Me.txtSMID.Text.Trim.Length > 10 Then
            Me.ErrorProvider1.SetError(Me.txtSMID, "The SMID can't be more than 10 characters")
            Me.txtSMID.Focus()
            blnOk = False
            intPanel = cstGeneralPanel

        End If

        If blnMROIssue Then
            blnOk = False
            intPanel = cstGeneralPanel
        End If

        'Check the SMID is a valid identity
        Dim strCheck As String = CConnection.PackageStringList("Lib_Utils.IsValidGenericIdentityChar", txtSMID.Text)
        If strCheck = "F" Then
            Me.ErrorProvider1.SetError(Me.txtProfile, "SMID has invalid characters can be only [0-9][A-Z]+-_, with no spaces")
            blnOk = False
            intPanel = cstGeneralPanel

        End If
        Client.SMID = Me.txtSMID.Text

        'Check for null profiles
        Me.ErrorProvider1.SetError(Me.txtProfile, "")
        If Me.txtProfile.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtProfile, "You must specify a profile for this customer")
            Me.txtProfile.Focus()
            blnOk = False
            intPanel = cstGeneralPanel

        End If

        If Me.txtProfile.Text.Trim.Length > 10 Then
            Me.ErrorProvider1.SetError(Me.txtProfile, "The profile can't be more than 10 characters")
            Me.txtProfile.Focus()
            blnOk = False
            intPanel = cstGeneralPanel

        End If
        ' Check profile is a valid identity
        strCheck = CConnection.PackageStringList("Lib_Utils.IsValidGenericIdentityChar", txtProfile.Text)
        If strCheck = "F" Then
            Me.ErrorProvider1.SetError(Me.txtProfile, "The Profile has invalid characters can be only [0-9][A-Z]+-_, with no spaces")
            blnOk = False
            intPanel = cstGeneralPanel

        End If
        Client.Profile = Me.txtProfile.Text

        Client.Comments = Me.txtComments.Text
        Client.CustomerAccountType = Me.txtAccountType.Text
        Dim strPassword As String = Me.txtEmailPassword.Text
        ErrorProvider1.SetError(txtEmailPassword, "")
        If MedscreenCommonGui.CustomerSupport.ValidatePassword(strPassword) Then
            Client.ReportingInfo.EmailPassword = Me.txtEmailPassword.Text
        Else
            ErrorProvider1.SetError(txtEmailPassword, "Invalid characters in password")
            blnOk = False
            intPanel = cstGeneralPanel
        End If
        'Client.Parent = Me.txtParent.Text
        If Me.AddressControl1.Address.Fields.Changed Then Me.AddressControl1.Address.Update()
        If Me.RBAddress.Checked Then
            Client.ResultsAddress = Me.AddressControl1.Address
            Client.ResultsAddress = Me.AddressControl1.Address
            'Client.Contact = Me.AddressControl1.Address.Contact
        End If

        Client.ContractNumber = Me.txtContractNumber.Text
        Client.PinNumber = Me.txtPinNumber.Text
        Client.VATRegno = Me.txtVATRegno.Text

        'check we have a sage id
        If Not Medscreen.IsValidSageID(Me.txtSageID.Text) Then
            Me.ErrorProvider1.SetError(Me.txtSageID, "You must have a valid sage ID for this customer, credit card payment customers can be assigned one of the MSCXX ids")
            Me.txtSageID.Focus()
            intPanel = cstInvoicingPanel
            blnOk = False
        End If
        Client.InvoiceInfo.SageID = Me.txtSageID.Text
        Client.Division = Me.cbDivision.SelectedValue
        Client.SubContractorOf = Me.txtSubContOf.Text
        Client.AccountSuspended = Me.CKSuspended.Checked
        Client.ContractDate = Me.dtContractDate.Value

        'Client.ReportingInfo.EmailPositives = Me.txtEmailPositives.Text
        If Not Client.SecondaryAddress Is Nothing Then
            'Client.SecondaryAddress.Email = Me.txtEmailPositives.Text
        Else
            'Create a secondary address
            Dim Addr As MedscreenLib.Address.Caddress
            Addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "CSHP", Me.Client.Identity, MedscreenLib.Constants.AddressType.AddressSec)
            Me.Client.ShippingAddressId = Addr.AddressId
            MedscreenLib.Address.AddressCollection.AddLink(Addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
            Client.SecondaryAddress = Addr
        End If
        'End If
        If Me.txtVAT.Text.Trim.Length = 0 Then
            'Client.Fields.Item("VAT_RATE").SetNull()
        Else
            'Client.InvoiceInfo.VATRate = Me.txtVAT.Text
        End If
        '

        Client.InvoiceInfo.PoRequired = Me.ckPORequired.Checked
        Client.ReportingInfo.PhoneBeforeFaxing = Me.ckPhoneBeforeFaxing.Checked
        Client.ReportingInfo.PositivesOnly = Me.ckPositivesOnly.Checked
        Client.ReportingInfo.DivertPositives = Me.ckDivertPos.Checked
        Client.ReportingInfo.PositivesAndNegatives = Me.ckEmailPosNeg.Checked
        Client.ReportingInfo.EmailWeeklyReport = Me.ckEmailWeeklyReport.Checked     'Email Weekly Report
        Client.ReportingInfo.EmailAndPrint = Me.chkEmailAndPrint.Checked

        Client.ReportingInfo.Bits = Me.ckReportInBits.Checked
        Client.ReportingInfo.EmailWeeklyReport = Me.ckEmailWeeklyReport.Checked
        Client.InvoiceInfo.CreditCard_pay = Me.ckCreditCard.Checked
        '<Added code modified 27-Apr-2007 11:41 by taylor> Credit card check
        If Not Client.InvoiceInfo.CreditCard_pay AndAlso Mid(Client.InvoiceInfo.SageID, 1, 3) = "MSC" Then
            Me.ErrorProvider1.SetError(Me.ckCreditCard, "MSCXX account needs to have card pay option set")
            Me.ckCreditCard.Checked = True
            Me.ckCreditCard.Focus()
            intPanel = cstInvoicingPanel
            blnOk = False

        End If
        '</Added code modified 27-Apr-2007 11:41 by taylor> 

        Client.ReportingInfo.FermentationNegative = Me.ckFerment.Checked
        Client.InvoiceInfo.AddTests = Me.ckAddTests.Checked
        Client.ReportingInfo.AllowInvalidPanel = Me.ckInvalidPanel.Checked
        Client.Removed = Me.CkRemoved.Checked
        Client.NewCustomer = Me.ckNewCustomer.Checked
        Client.StopCerts = Me.ckStopCerts.Checked
        Client.StopSamples = Me.CKStopSamples.Checked
        Client.ProgrammeManager = Me.cbProgMan.SelectedValue

        Client.SalesPerson = Me.cbSalesPerson.SelectedValue

        If Me.cbStatus.Text.Trim.Length > 0 Then
            Client.AccStatus = Me.cbStatus.SelectedValue
        Else
            Client.AccStatus = "A"
        End If

        Dim objPhrase As MedscreenLib.Glossary.Phrase
        ''Dim i As Integer
        If Me.CBCustomerType.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.AccountType.Item(Me.CBCustomerType.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.AccountCollectionType = objPhrase.PhraseID
                If Me.Client.Services.Count = 0 Then
                    If MsgBox("Do you want the default services set up for this customer type?", _
                        MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        Dim aSerList As AccountTypeServices.ServiceRow() = AccountServices.AccountTypeServicesLt(objPhrase.PhraseID)
                        If Not aSerList Is Nothing Then
                            Dim aSerItem As AccountTypeServices.ServiceRow
                            For Each aSerItem In aSerList
                                Dim aService As New Intranet.intranet.customerns.CustomerService()
                                aService.CustomerId = Me.Client.Identity
                                aService.Service = aSerItem.Service_Text
                                aService.[Operator] = Intranet.intranet.Support.cSupport.User.Identity
                                aService.update()
                                Me.Client.Services.Add(aService)
                            Next
                        End If
                    End If
                End If
            End If
        End If

        If Me.cbIndustry.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.Industry.Item(Me.cbIndustry.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.IndustryCode = objPhrase.PhraseID
                If objPhrase.PhraseID = "OTH-OTH" Then
                    blnOk = False
                    Me.ErrorProvider1.SetError(Me.cbIndustry, "The 'unknown' Industry code must be replace with the correct code")

                End If
            Else
                    blnOk = False
                    Me.ErrorProvider1.SetError(Me.cbIndustry, "An Industry code must be provided")
                End If
            Else
            blnOk = False
            Me.ErrorProvider1.SetError(Me.cbIndustry, "An Industry code must be provided")

        End If

        If Me.cbPlatAccType.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.Platinum.Item(Me.cbPlatAccType.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.InvoiceInfo.PlatAccType = objPhrase.PhraseID
            End If
        End If


        If Me.cbReportMethod.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.ReportMethod.Item(Me.cbReportMethod.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.ReportingInfo.ReportType = objPhrase.PhraseID
            End If
        End If

        If cbFaxFormat.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.FaxFormat.Item(Me.cbFaxFormat.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.ReportingInfo.FaxFormat = objPhrase.PhraseID
            End If
        End If

        If Me.cbInvFreq.SelectedIndex > -1 Then
            Client.InvoiceInfo.InvFrequency = Me.cbInvFreq.SelectedValue
        End If

        'Currency
        If Me.cbCurrency.SelectedIndex > -1 Then
            'Client.Currency = Mid(Me.cbCurrency.Text, 1, 1)
            Client.Currency = Me.cbCurrency.SelectedValue
        End If

        'MRO
        Client.ReportingInfo.MROID = Me.cbMROID.SelectedValue()
        'If Me.cbMROID.SelectedIndex > -1 Then
        '    Client.ReportingInfo.MROID = Me.cbMROID.SelectedValue
        'End If
        Client.Parent = Me.cbParent.SelectedValue

        If Not blnOk Then
            Dim dlgRet As MsgBoxResult = MsgBoxResult.Yes
            '<Added code modified 01-May-2007 10:26 by taylor> 
            If Client.SMID = "TEMPLATES" Then
                dlgRet = MsgBox("Do you want to examine errors?", MsgBoxStyle.YesNo Or MsgBoxStyle.Information)
                If dlgRet = MsgBoxResult.Yes Then
                    Me.DialogResult = DialogResult.None
                End If
            Else
                Me.DialogResult = DialogResult.None
            End If
            '</Added code modified 01-May-2007 10:26 by taylor> 

            Me.TabControl1.SelectedIndex = intPanel
            If dlgRet = MsgBoxResult.Yes Then MsgBox("There are errors on the form," & vbCrLf & _
            "they will be marked by red exclamation marks which will explain the error")
            Exit Sub
        End If

        'Deal with any price changes that may have occurred by sending a warning 
        If Me.blnPriceChanged Then

        End If

    End Sub

    Public Property NewClient() As Boolean
        Get
            Return blnNew
        End Get
        Set(ByVal Value As Boolean)
            blnNew = Value
        End Set
    End Property





    Private Sub Customer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Customer.Click

    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If blnForceClose Then
            e.Cancel = False
            Exit Sub
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If Not Me.Client Is Nothing Then Me.Client.Rollback()
        blnForceClose = True
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the address and make sure the right one is updated, address used indicates which address. 
    ''' The two addresses only in the extra table need special handling
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub CheckAddressChanged()

        If Me.Client Is Nothing Then Exit Sub
        ' Make the commands invisible
        'Me.cmdNewAddress.Visible = False
        Me.lblPosEmail.Visible = False
        Me.txtEmailPositives.Visible = False

        If Not Me.AddressControl1.Address Is Nothing AndAlso _
        Me.AddressControl1.Address.AddressId > 0 AndAlso _
        Me.AddressControl1.Address.Fields.Changed Then
            Me.AddressControl1.Address.Update()
            Me.AddressControl1.Address.SetAddressLink(Me.Client.Identity, Me.AddressUsed)
        End If
        '<Added code modified 13-Apr-2007 09:27 by taylor> 
        'Check addresslink can only do if we have a valid address.
        If Me.AddressControl1.Address Is Nothing Then Exit Sub
        If Me.AddressControl1.Address.GetAddressLink(Me.Client.Identity, Me.AddressUsed) <= 0 Then
            Me.AddressControl1.Address.SetAddressLink(Me.Client.Identity, Me.AddressUsed)
        End If
        'End If
        '</Added code modified 13-Apr-2007 09:27 by taylor> 

        ' check through adddress in both
        If Me.AddressUsed = Constants.AddressType.AddressMain Then
            'If Not Me.AddressControl1.IsValid Then          'Check that the address is fine
            '    MsgBox("There are issues with the address please fix them.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            'End If
            Me.Client.ResultsAddress = Me.AddressControl1.Address
            If Not Me.AddressControl1.Address Is Nothing Then ' can't do anything without an address
                If Me.AddressControl1.Address.Email.Trim.Length > 0 Then 'We have a results address
                    Me.cmdPassword.Enabled = True
                Else
                    Me.cmdPassword.Enabled = False
                End If
                If Me.Client.MailAddressId = 0 Then
                    'Me.Client.MailAddressId = Me.AddressControl1.Address.AddressId
                    Me.Client.MailAddressId = Me.AddressControl1.Address.AddressId
                End If
            Else
                Me.cmdPassword.Enabled = False

                Me.Client.MailAddressId = -1
            End If
            'With the main address we also have to worry about positives
            If Me.Client.SecondaryAddress Is Nothing AndAlso Me.txtEmailPositives.Text.Trim.Length > 0 Then
                'we need to create a secondary address
                Dim Addr As MedscreenLib.Address.Caddress
                Addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "CSHP", Me.Client.Identity, MedscreenLib.Constants.AddressType.AddressSec)
                Me.Client.ShippingAddressId = Addr.AddressId
                MedscreenLib.Address.AddressCollection.AddLink(Addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
                Client.SecondaryAddress = Addr
                'Addr.Email = Me.txtEmailPositives.Text
            ElseIf Not Me.Client.SecondaryAddress Is Nothing Then
                Me.txtEmailPositives.Text = Me.Client.SecondaryAddress.Email
                'see if Id is 0 
                If Me.Client.SecondaryAddress.AddressId <= 0 Then
                    Client.SecondaryAddress.AddressId = MedscreenLib.Glossary.Glossary.SMAddresses.NextAddressId
                    Client.ShippingAddressId = Client.SecondaryAddress.AddressId
                    Client.SecondaryAddress.Update()
                    Dim ocmd As New OleDb.OleDbCommand("update customer set shipaddressid = " & Client.SecondaryAddress.AddressId & _
                      " where identity = '" & Client.Identity & "'", CConnection.DbConnection)
                    If CConnection.ConnOpen Then
                        Dim intRet As Integer = ocmd.ExecuteNonQuery()
                        CConnection.SetConnClosed()
                    End If
                End If
            End If
        ElseIf Me.AddressUsed = Constants.AddressType.AddressSec Then
            Me.Client.SecondaryAddress = Me.AddressControl1.Address
            If Not Me.AddressControl1.Address Is Nothing Then
                If Me.Client.ShippingAddressId = 0 Then
                    Me.Client.ShippingAddressId = Me.AddressControl1.Address.AddressId
                End If
                Me.Client.SecondaryAddress.Update()
                'update the text version of the secondary address
                Me.txtEmailPositives.Text = Me.Client.SecondaryAddress.Email
            Else
                'Me.Client.ExtraInfo.ShipAddressID = -1
                Me.Client.ShippingAddressId = -1
            End If
        ElseIf Me.AddressUsed = Constants.AddressType.AddressInv Then
            Me.Client.InvoiceInfo.InvoiceAddress = Me.AddressControl1.Address
            'now deal with the ones only in trhe extra table
            If Not Me.AddressControl1.Address Is Nothing Then
                If Me.Client.InvoiceAddressId = 0 Then
                    'Me.Client.ExtraInfo.InvoiceAddressID = Me.AddressControl1.Address.AddressId

                    Me.Client.InvoiceAddressId = Me.AddressControl1.Address.AddressId
                End If
            Else
                'Me.Client.ExtraInfo.InvoiceAddressID = -1
                Me.Client.InvoiceAddressId = -1
            End If

        ElseIf Me.AddressUsed = Constants.AddressType.InvoiceMail Then
            If Not Me.AddressControl1.Address Is Nothing Then
                Me.AddressControl1.Address.Update()
                If Me.Client.InvoiceMailAddressID = 0 Then
                    'Me.Client.ExtraInfo.InvoiceMailAddressID = Me.AddressControl1.Address.AddressId
                    Me.Client.InvoiceMailAddressID = Me.AddressControl1.Address.AddressId()
                End If
            Else
                'Me.Client.ExtraInfo.InvoiceMailAddressID = -1
                Me.Client.InvoiceMailAddressID = -1
            End If
        ElseIf Me.AddressUsed = Constants.AddressType.InvoiceShipping Then
            If Not Me.AddressControl1.Address Is Nothing Then
                Me.AddressControl1.Address.Update()
                If Me.Client.InvoiceShippingAddressId = 0 Then
                    'Me.Client.ExtraInfo.InvoiceShippingAddressID = Me.AddressControl1.Address.AddressId
                    Me.Client.InvoiceShippingAddressId = Me.AddressControl1.Address.AddressId
                End If
            Else
                'Me.Client.ExtraInfo.InvoiceShippingAddressID = -1
                Me.Client.InvoiceShippingAddressId = -1
            End If
        End If

        Dim RBControl As Control

        For Each RBControl In Me.GroupBox1.Controls
            If TypeOf RBControl Is RadioButton Then
                RBControl.BackColor = Drawing.SystemColors.Control
                RBControl.ForeColor = SystemColors.ControlText
            End If
        Next
        'TODO add check on address links
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main address selected 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RBAddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBAddress.CheckedChanged, RBAddress.Leave
        If Client Is Nothing Then Exit Sub

        CheckAddressChanged()
        If RBAddress.Checked Then
            Me.AddressUsed = Constants.AddressType.AddressMain
            Dim strAddressMismatch As String = "lib_Customer.SharedMailAddress"
            Dim strXML As String = CConnection.PackageStringList(strAddressMismatch, Client.Identity)
            Dim strText As String = Medscreen.ResolveStyleSheet(strXML, "PairedAddresstxt1.xsl", 0)
            Me.txtSameAddress.Rtf = strText
            If Not Client.ResultsAddress Is Nothing Then
                Me.AddressControl1.Address = Client.ResultsAddress
            Else
                Me.AddressControl1.Address = Nothing
            End If
            Me.lblPosEmail.Visible = True
            Me.txtEmailPositives.Visible = True
            SetControlHighlighted(Me.RBAddress)
            Me.AddressControl1.AddressType = Me.AddressUsed

        End If
    End Sub

    Private Sub RBSecondary_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBSecondary.CheckedChanged, RBSecondary.Leave
        If Client Is Nothing Then Exit Sub
        CheckAddressChanged()
        If RBSecondary.Checked Then
            Me.txtSameAddress.Rtf = ""

            Me.AddressUsed = Constants.AddressType.AddressSec
            If Not Client.SecondaryAddress Is Nothing Then
                'Check for existing email positives 
                If Client.SecondaryAddress.Email.Trim.Length = 0 Then 'If we don't have one then set it from the emailpositives
                    'Client.SecondaryAddress.Email = Me.Client.ReportingInfo.EmailPositives
                    ' Client.SecondaryAddress.Update()
                End If
                Me.AddressControl1.Address = Client.SecondaryAddress
            Else
                Me.AddressControl1.Address = Nothing
            End If
            SetControlHighlighted(Me.RBSecondary)
            Me.AddressControl1.AddressType = Me.AddressUsed
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main invoicing address changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action>restored event handler</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RBInvoicing_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBInvoicing.CheckedChanged, RBInvoicing.Leave
        If Client Is Nothing Then Exit Sub

        CheckAddressChanged()       'Sort out old address

        If RBInvoicing.Checked Then     'Have we  been checked?
            Dim strAddressMismatch As String = "lib_Customer.SharedInvMailAddress"
            Dim strXML As String = CConnection.PackageStringList(strAddressMismatch, Client.Identity)
            If Not strXML Is Nothing AndAlso strXML.Trim.Length > 0 Then
                Dim strText As String = Medscreen.ResolveStyleSheet(strXML, "PairedInvAddresstxt1.xsl", 0)
                Me.txtSameAddress.Rtf = strText
            End If
            Me.AddressUsed = Constants.AddressType.AddressInv
            If Not Client.InvoiceInfo.InvoiceAddress Is Nothing Then
                Me.AddressControl1.Address = Client.InvoiceInfo.InvoiceAddress
            Else
                Me.AddressControl1.Address = Nothing
            End If
            SetControlHighlighted(Me.RBInvoicing)
            Me.AddressControl1.AddressType = Me.AddressUsed
        End If
    End Sub

    Private Sub cmdCustomerOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomerOptions.Click
        Dim frm As New MedscreenCommonGui.frmCustomerOptions()

        With frm
            .Client = Client
            If .ShowDialog = DialogResult.OK Then
                Client = .Client
                'client.Options
                blnOptionsDraw = False
                Me.DrawOptions()
            End If
        End With
    End Sub

    Private WithEvents CustOptions As MedscreenCommonGui.frmCustomerOptions
    Private Sub cmdServices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdServices.Click
        CustOptions = New MedscreenCommonGui.frmCustomerOptions()

        Try
            With CustOptions
                .FormType = MedscreenCommonGui.frmCustomerOptions.COFormType.Services
                .Client = Me.Client
                .ShowDialog()
            End With
            'Refesh the services
            Me.Client.Services.Refresh()
            FillPrices()
            DrawPrices()
            Me.CheckMinCharges()
        Catch ex As Exception
            LogError(ex, True, "Do services")
        End Try

    End Sub

    'Private Sub cbPricing_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPricing.SelectedIndexChanged
    '    If Me.blnDrawing Then Exit Sub

    '    If Me.intPricingOption = Me.cstStandardPricing Then
    '        Client.StandardPricing = False
    '        'ElseIf Me.intPricingOption = Me.cstOldPricing Then
    '        '   Me.Client.Options.DeleteOption("OLD_PRICES")
    '    ElseIf Me.intPricingOption = Me.cstAlternatePricing Then
    '        Client.StandardPricing = True
    '    End If
    '    Me.intPricingOption = cbPricing.SelectedIndex
    '    Client.InvoiceInfo.ResetTimeStamp()
    '    FillPrices()
    '    Me.DrawHistory()
    'End Sub

    Private strPriceBook As String = ""
    Private Sub cbPricing_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPricing.SelectedIndexChanged
        If Me.blnDrawing Then Exit Sub

        'Set up to check price book change
        Dim strPriceBook As String = Client.PriceBook
        Me.intPricingOption = cbPricing.SelectedIndex
        If cbPricing.SelectedValue <> strPriceBook Then
            Client.PriceBook = cbPricing.SelectedValue
            If Client.PriceBook.Substring(0, 3) = "STD" Then
                'Client.StandardPricing = True
            Else
                'Client.StandardPricing = False
            End If
            strPriceBook = Client.PriceBook
            Client.InvoiceInfo.ResetTimeStamp()
            FillPrices()
        End If
    End Sub


    Private Sub TabOptions_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabOptions.TabIndexChanged

    End Sub


    Private Sub lvAccounts_OptionAdded(ByVal Sender As MedscreenCommonGui.OptionsListView, ByVal option_id As String) 'Handles lvAccounts.OptionAdded, lvAccounts.OptionRemoved, _
        'lvCertificates.OptionAdded(, lvCertificates.OptionRemoved, lvContractor.OptionAdded, lvContractor.OptionRemoved, _
        'lvCustomerType.OptionAdded, lvCustomerType.OptionRemoved, LvPricing.OptionAdded, LvPricing.OptionRemoved, _
        'lvPrisons.OptionAdded, lvPrisons.OptionRemoved, lvRepOptions.OptionAdded, lvRepOptions.OptionRemoved, _
        'lvReporting.OptionAdded, lvReporting.OptionRemoved, lvSamples.OptionAdded, lvSamples.OptionRemoved)

        Me.Client.Options = Sender.ClientOptions
        blnOptionsDraw = False
        Me.DrawOptions()
    End Sub




    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim objRP As Intranet.intranet.customerns.RoutinePayment = Me.Client.InvoiceInfo.RoutinePayments.CreatePayment
        objRP.ContractStart = Now
        Dim tmpDate As Date = Now.AddYears(1)
        tmpDate = tmpDate.AddDays(-1)
        objRP.ContractEnd = tmpDate
        objRP.LineItem = "RTN0001"
        objRP.InvoiceFrequency = "ANN"
        objRP.NextPayment = Now.AddYears(1)
        objRP.FirstPayment = Now
        objRP.LastPayment = tmpDate
        objRP.ContractNumber = Hex(Now.Year) & "-" & CConnection.NextSequence("SEQ_CONTRACT_NO", 6)
        'objRP.ContractNumber = Medscreen.GetParameter(MyTypes.typString, "Contract Number", "Contract Number", "")
        Me.currentPayment = New LVRoutinePayment(objRP)
        'Me.DrawPayments()
        Me.lvRoutinePayments.Items.Add(Me.currentPayment)
        Me.FillPayments(objRP)
        DTContractStart.Focus()
    End Sub

    Private Class LVRoutinePayment
        Inherits ListViewItem

        Private objRP As Intranet.intranet.customerns.RoutinePayment

        Public Sub New(ByVal RoutinePayment As Intranet.intranet.customerns.RoutinePayment)
            objRP = RoutinePayment
            Draw()
        End Sub

        Public Sub Draw()
            Me.SubItems.Clear()
            Dim SubItem As ListViewItem.ListViewSubItem
            Me.Text = objRP.ContractNumber
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.LineItem)
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.ContractStart.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.ContractEnd.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.AnnualPrice)
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.IntervalPrice)
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.InvoiceFrequency)
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.NextPayment.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(SubItem)
            SubItem = New ListViewItem.ListViewSubItem(Me, objRP.LastInvoice)
            Me.SubItems.Add(SubItem)
        End Sub

        Public Property RoutinePayment() As Intranet.intranet.customerns.RoutinePayment
            Get
                Return objRP
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.RoutinePayment)
                objRP = Value
            End Set
        End Property

    End Class

    Dim currentPayment As LVRoutinePayment
    Private Sub lvRoutinePayments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvRoutinePayments.Click
        Dim lvItem As LVRoutinePayment
        currentPayment = Nothing
        'objRP = Nothing
        If Me.lvRoutinePayments.SelectedItems.Count = 0 Then Exit Sub

        currentPayment = Me.lvRoutinePayments.SelectedItems.Item(0)
        'objRP = lvItem.RoutinePayment
        Me.FillPayments(currentPayment.RoutinePayment)
        Me.cmdDeleteContract.Show()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update the contract information 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [31/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        With currentPayment.RoutinePayment
            .AnnualPrice = Me.txtAnnualPrice.Text              'Contrcat total price 
            .ContractNumber = Me.txtContractNumber1.Text        'set Contract number
            .ContractStart = Me.DTContractStart.Value          'Set contract start date
            .ContractEnd = Me.DTContractEndDate.Value          'Set contract end date
            .Removed = Me.ckContractRemoved.Checked            'Has this contract been removed 
            If Me.cbLineItems.SelectedIndex > -1 Then
                .LineItem = CCToolAppMessages.RoutineLineItems.Keys.Item(Me.cbLineItems.SelectedIndex)
            End If
            If Me.cbPayInterval.SelectedIndex <> -1 Then
                Dim objPhrase As MedscreenLib.Glossary.Phrase = _
                     MedscreenLib.Glossary.Glossary.PaymentIntervals.Item(Me.cbPayInterval.SelectedIndex)
                If Not objPhrase Is Nothing Then
                    '    objRP.InvoiceFrequency = objPhrase.PhraseID
                    '    Select Case objPhrase.PhraseID
                    '        Case "ANN"
                    '            objRP.NextPayment = objRP.ContractStart.AddYears(1)
                    '        Case "QTR"
                    '            objRP.NextPayment = objRP.ContractStart
                    '            While objRP.NextPayment < Today
                    '                objRP.NextPayment = objRP.NextPayment.AddMonths(3)
                    '            End While
                    '            objRP.IntervalPrice = objRP.AnnualPrice / 4
                    '        Case "MON"
                    '            objRP.NextPayment = objRP.ContractStart
                    '            While objRP.NextPayment < Today
                    '                objRP.NextPayment = objRP.NextPayment.AddMonths(1)
                    '            End While
                    '            objRP.IntervalPrice = objRP.AnnualPrice / 12
                    '        Case "WEEK"
                    '            objRP.NextPayment = objRP.ContractStart
                    '            While objRP.NextPayment < Today
                    '                objRP.NextPayment = objRP.NextPayment.AddDays(7)
                    '            End While
                    '            objRP.IntervalPrice = (objRP.AnnualPrice / 365.25) * 7

                    '    End Select

                    .InvoiceFrequency = objPhrase.PhraseID
                End If
            End If
            If Me.DTNextDue.Value > .FirstPayment Then .NextPayment = Me.DTNextDue.Value
            .FirstPayment = Me.dtFirstPayment.Value
            If Char.IsDigit(Me.txtNextPayment.Text.Trim.Chars(0)) Then  'No currency code
                .IntervalPrice = Me.txtNextPayment.Text
            Else
                .IntervalPrice = Mid(Me.txtNextPayment.Text, 2)
            End If
            .ItemComment = Me.txtItemComment.Text

            .Update()
            currentPayment.Draw()
            'Me.DrawPayments()
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoice mail radio button changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub RBInvMail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBInvMail.CheckedChanged
        If Client Is Nothing Then Exit Sub

        CheckAddressChanged()                   ' sort out old address in control

        If Me.RBInvMail.Checked Then
            Me.txtSameAddress.Rtf = ""

            Me.AddressUsed = Constants.AddressType.InvoiceMail
            If Not Client.InvoiceMailAddress Is Nothing Then      'Check valid address
                Me.AddressControl1.Address = Client.InvoiceMailAddress    'set it 
                'Me.cmdNewAddress.Visible = True
            Else
                Me.AddressControl1.Address = Nothing                        'clear control
                'Me.cmdNewAddress.Visible = True
            End If
            SetControlHighlighted(Me.RBInvMail)
            Me.AddressControl1.AddressType = Me.AddressUsed
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoice Shipping address radio button changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub rbInvShipping_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInvShipping.CheckedChanged
        If Client Is Nothing Then Exit Sub

        CheckAddressChanged()

        If Me.rbInvShipping.Checked Then        'Are we in a checked state
            Me.txtSameAddress.ResetText()

            Me.AddressUsed = Constants.AddressType.InvoiceShipping      ' Set address used
            If Not Client.InvoiceShippingAddress Is Nothing Then  'Do we have an address
                Me.AddressControl1.Address = Client.InvoiceShippingAddress    'Set it 
            Else
                Me.AddressControl1.Address = Nothing                    ' Null address control
                'Me.cmdNewAddress.Visible = True                         ' Show controls
            End If
            SetControlHighlighted(Me.rbInvShipping)
            Me.AddressControl1.AddressType = Me.AddressUsed
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create a new address and display it 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNewAddress_Click(ByVal AddressID As Integer) Handles AddressControl1.NewAddress

        If Me.AddressUsed = Constants.AddressType.InvoiceShipping Then
            Me.Client.InvoiceShippingAddressId = AddressID
            Me.AddressControl1.Address.AddressType = "INSA"
            MedscreenLib.Address.AddressCollection.AddLink(AddressID, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
        ElseIf Me.AddressUsed = Constants.AddressType.InvoiceMail Then
            Me.Client.InvoiceMailAddressID = AddressID
            Me.AddressControl1.Address.AddressType = "INMA"
            MedscreenLib.Address.AddressCollection.AddLink(AddressID, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
        ElseIf Me.AddressUsed = Constants.AddressType.AddressMain Then
            Me.Client.MailAddressId = AddressID
            Me.AddressControl1.Address.AddressType = "INMA"
            MedscreenLib.Address.AddressCollection.AddLink(AddressID, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
        ElseIf Me.AddressUsed = Constants.AddressType.AddressSec Then
            Me.Client.ShippingAddressId = AddressID
            Me.AddressControl1.Address.AddressType = "CSHP"
            MedscreenLib.Address.AddressCollection.AddLink(AddressID, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
        ElseIf Me.AddressUsed = Constants.AddressType.AddressInv Then
            Me.Client.InvoiceAddressId = AddressID
            Me.AddressControl1.Address.AddressType = "INSA"
            MedscreenLib.Address.AddressCollection.AddLink(AddressID, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)

        End If

        'Me.cmdNewAddress.Visible = False

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find the address so it can be set up
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdFindAddress_Click(ByVal addressId As Integer) Handles AddressControl1.FindAddress

        Dim intAddrID As Integer
        intAddrID = addressId                          ' a valid address put it in 
        If Me.AddressUsed = Constants.AddressType.InvoiceShipping Then
            Me.Client.InvoiceShippingAddressId = intAddrID
        ElseIf Me.AddressUsed = Constants.AddressType.InvoiceMail Then
            Me.Client.InvoiceMailAddressID = intAddrID
        ElseIf Me.AddressUsed = Constants.AddressType.AddressMain Then
            Me.Client.MailAddressId = intAddrID
        ElseIf Me.AddressUsed = Constants.AddressType.AddressInv Then
            Me.Client.InvoiceAddressId = intAddrID
        ElseIf Me.AddressUsed = Constants.AddressType.AddressSec Then
            Me.Client.ShippingAddressId = intAddrID
        End If

        'get address and fill control
        Dim addr As MedscreenLib.Address.Caddress
        If intAddrID > 0 Then
            addr = MedscreenLib.Glossary.Glossary.SMAddresses.item(intAddrID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
            Me.AddressControl1.Address = addr
            MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
            'TODO: Need to change address links
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' set selected contro to higlighted 
    ''' </summary>
    ''' <param name="rbControl"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [14/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetControlHighlighted(ByVal rbControl As Control)
        rbControl.BackColor = Drawing.SystemColors.ActiveCaption
        rbControl.ForeColor = SystemColors.ActiveCaptionText
    End Sub

    Private Sub TabAddresses_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabAddresses.Enter
        Me.RBAddress.Checked = True
    End Sub

    Private Sub rbGoldmine_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbGoldmine.CheckedChanged
        If Client Is Nothing Then Exit Sub

        CheckAddressChanged()
        If rbGoldmine.Checked Then
            Me.AddressUsed = Constants.AddressType.GoldmineAddress
            If Client.GoldMineId.Trim.Length > 0 Then
                Dim GMContact As MedscreenCommonGui.GoldMineContactRO

                GMContact = Me.Client.GoldMineContact
                If Not GMContact Is Nothing Then
                    If Not GMContact.Address Is Nothing Then
                        Dim objAddress As New MedscreenLib.Address.Caddress()
                        objAddress.AddressId = -2
                        With GMContact.Address
                            objAddress.AdrLine1 = .Line1
                            objAddress.AdrLine2 = .Line2
                            objAddress.AdrLine2 = .Line3
                            objAddress.Country = .Country
                            objAddress.City = .City
                            objAddress.PostCode = .PostCode
                            objAddress.District = .District
                            objAddress.Contact = GMContact.Contact
                        End With
                        Me.AddressControl1.Address = objAddress
                    Else
                        Me.AddressControl1.Address = Nothing
                    End If
                Else
                    Me.AddressControl1.Address = Nothing
                End If

            Else
                Me.AddressControl1.Address = Nothing
            End If
            Me.lblPosEmail.Visible = True
            Me.txtEmailPositives.Visible = True
            SetControlHighlighted(Me.RBAddress)
        End If

    End Sub

    Private Sub cmdDeleteContract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteContract.Click
        If Not currentPayment Is Nothing Then
            If MsgBox("Delete contract?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                currentPayment.RoutinePayment.Delete()
            End If
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvRoutinePayments.SelectedIndexChanged

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with a new Address being created 
    ''' </summary>
    ''' <param name="AddressId"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub AddressControl1_NewAddress(ByVal AddressId As Integer)
        CheckAddressChanged()
        'Me.Client.Update()
    End Sub

    Private Sub CmdCreateTestPanels_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCreateTestPanels.Click
        If Me.Client Is Nothing Then
            Exit Sub
        End If
        If Me.myCustomerTestPanel Is Nothing Then Exit Sub
        Me.myCustomerTestPanel.CreateTestSchedule()

    End Sub

    Private myTestSchedule As Intranet.intranet.TestResult.TestScheduleEntry

    Private Sub lvTestSchedules_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvTestSchedules.Click
        Dim lvCTP As MedscreenCommonGui.ListViewItems.lvTestSchedule
        myTestSchedule = Nothing
        Me.cmdDeleteTest.Hide()

        'See if we have a test schedule 
        If Me.lvTestSchedules Is Nothing Then
            Me.cmdAddTest.Hide()
            Exit Sub
        Else
            Me.cmdAddTest.Show()
        End If


        If Me.lvTestSchedules.SelectedItems.Count = 1 Then
            lvCTP = Me.lvTestSchedules.SelectedItems.Item(0)
            myTestSchedule = lvCTP.Schedule
            If Not myTestSchedule Is Nothing Then Me.cmdDeleteTest.Show()
            'Me.myCustomerTestPanel = lvCTP.CustomerTestPanel
        End If
        'Me.DrawTestSchedules()
    End Sub

    Private Sub cmdPassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPassword.Click
        Me.txtEmailPassword.Text = MedscreenLib.Medscreen.GeneratePassword(6, 8, True, True, 2)
        If passwordState = ClientReport.PasswordLetters.NewPassword Then
            Me.Client.ReportingInfo.PasswordLetter(ClientReport.PasswordLetters.NewPassword, Me.txtEmailPassword.Text, Now)
        Else
            Me.Client.ReportingInfo.PasswordLetter(ClientReport.PasswordLetters.ChangedPassword, Me.txtEmailPassword.Text, Now)
        End If
    End Sub

    Private blnUpdateingContract As Boolean = False
    Private Sub UpdateContractDetails()
        If currentPayment Is Nothing Then Exit Sub 'Bog off if there is no payment object 
        If blnUpdateingContract Then Exit Sub
        blnUpdateingContract = True
        If Me.txtAnnualPrice.Text.Trim.Length > 0 Then currentPayment.RoutinePayment.AnnualPrice = Me.txtAnnualPrice.Text 'Contrcat total price 
        currentPayment.RoutinePayment.ContractNumber = Me.txtContractNumber1.Text        'set Contract number
        currentPayment.RoutinePayment.ContractStart = Me.DTContractStart.Value          'Set contract start date
        currentPayment.RoutinePayment.ContractEnd = Me.DTContractEndDate.Value          'Set contract end date
        currentPayment.RoutinePayment.Removed = Me.ckContractRemoved.Checked            'Has this contract been removed 
        UpdateContractDates()
        blnUpdateingContract = False
    End Sub

    Private Sub UpdateContractDates()
        Dim objPhrase As MedscreenLib.Glossary.Phrase = Me.cbPayInterval.SelectedItem
        If Me.DTContractEndDate.Value <= Me.DTContractStart.Value Then _
            Me.DTContractEndDate.Value = Me.DTContractStart.Value.AddYears(1).AddDays(-1)
        If Me.dtFirstPayment.Value >= Me.DTNextDue.Value Then _
            Me.DTNextDue.Value = Me.dtFirstPayment.Value.AddYears(1).AddDays(-1)
        If Not objPhrase Is Nothing Then
            Select Case objPhrase.PhraseID
                Case "ANN"
                    Me.DTNextDue.Value = currentPayment.RoutinePayment.FirstPayment.AddYears(1)

                    currentPayment.RoutinePayment.IntervalPrice = currentPayment.RoutinePayment.AnnualPrice
                    Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)
                Case "WEEK"
                    Me.DTNextDue.Value = currentPayment.RoutinePayment.FirstPayment.AddDays(7)
                    currentPayment.RoutinePayment.IntervalPrice = currentPayment.RoutinePayment.AnnualPrice * 7 / (365.25)
                    Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)

                Case "FORT"
                    Me.DTNextDue.Value = currentPayment.RoutinePayment.FirstPayment.AddDays(14)
                    currentPayment.RoutinePayment.IntervalPrice = currentPayment.RoutinePayment.AnnualPrice * 14 / (365.25)
                    Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)
                Case "QTR"
                    Me.DTNextDue.Value = currentPayment.RoutinePayment.FirstPayment.AddMonths(3)
                    currentPayment.RoutinePayment.IntervalPrice = currentPayment.RoutinePayment.AnnualPrice / 4
                    Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)
                Case "MON"
                    Me.DTNextDue.Value = currentPayment.RoutinePayment.FirstPayment.AddMonths(1)
                    currentPayment.RoutinePayment.IntervalPrice = currentPayment.RoutinePayment.AnnualPrice / 12
                    Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)

            End Select
            Me.txtNextPayment.Text = currentPayment.RoutinePayment.IntervalPrice.ToString("c", Me.Client.Culture)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' React to changes to control 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbPayInterval_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPayInterval.SelectedIndexChanged
        UpdateContractDetails()
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
        afrm.FormType = FrmCollComment.FCCFormTypes.customerComment
        afrm.CommentType = MedscreenLib.Constants.GCST_COMM_CollCOOfficerInfo
        afrm.CMNumber = Me.Client.Identity
        If Me.Client.ProgrammeManager.Trim.Length > 0 Then
            afrm.ProgamManager = Me.Client.SalesPerson
        End If
        If afrm.ShowDialog = DialogResult.OK Then
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, Me.Client.Identity, afrm.CommentType, afrm.Comment, _
            MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            'If Not aComm Is Nothing Then
            Me.DrawComments()
            'End If
        End If

    End Sub

    'Private Sub lvHistory_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs)

    '    Dim intdirection As Integer

    '    intdirection = Me.ColumnDirection(e.Column)     'Get last direction 
    '    'if we have user option information update what has been selected
    '    If Not CommonForms.UserOptions Is Nothing Then
    '        CommonForms.UserOptions.CustSelectColumn = e.Column
    '        CommonForms.UserOptions.CustSelectSortDirection = intdirection
    '    End If
    '    'flip direction 
    '    Me.ColumnDirection(e.Column) *= -1
    '    'Set up a new sorter
    '    Me.lvHistory.ListViewItemSorter = New CustSortViewItemComparer(e.Column, intdirection)
    'End Sub

    Private Sub txtAnnualPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAnnualPrice.Leave
        Me.UpdateContractDetails()
    End Sub

    Private Sub dtFirstPayment_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFirstPayment.ValueChanged
        If Me.currentPayment Is Nothing Then Exit Sub
        If Me.currentPayment.RoutinePayment Is Nothing Then Exit Sub

        Me.currentPayment.RoutinePayment.FirstPayment = Me.dtFirstPayment.Value
        UpdateContractDetails()
    End Sub

    Private Sub DTContractStart_valuechanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTContractStart.ValueChanged
        If Me.currentPayment Is Nothing Then Exit Sub
        If Me.currentPayment.RoutinePayment Is Nothing Then Exit Sub
        Dim tmpDate As Date = Me.DTContractStart.Value
        tmpDate = tmpDate.AddYears(1)
        tmpDate = tmpDate.AddDays(-1)
        If tmpDate > Me.DTContractStart.Value Then Me.DTContractEndDate.Value = tmpDate
        If Me.DTContractStart.Value < Me.dtFirstPayment.Value Then Me.dtFirstPayment.Value = Me.DTContractStart.Value
        If currentPayment Is Nothing OrElse Me.currentPayment.RoutinePayment Is Nothing Then Exit Sub
        Me.currentPayment.RoutinePayment.FirstPayment = Me.dtFirstPayment.Value
        UpdateContractDates()
    End Sub

    Private Sub lvComments_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvComments.SelectedIndexChanged
        Me.myCustComment = Nothing
        If Me.lvComments.SelectedItems.Count = 1 Then
            Me.myCustComment = Me.lvComments.SelectedItems.Item(0)
            If Not Me.myCustComment Is Nothing AndAlso Not Me.myCustComment.Comment Is Nothing AndAlso myCustComment.Comment.CommentType = "COLL" Then

            End If
        End If
    End Sub

    Private Sub tabHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabHistory.Click
        If Me.Client Is Nothing Then Exit Sub
        Me.Client.History.RefreshAppend()
    End Sub

    

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        'See if contacts have changed
        CheckForChangedContact()

        If Me.TabControl1.SelectedIndex = Me.TabControl1.TabPages.IndexOf(Me.tabHistory) Then
            DrawHistory()
            If Me.Client Is Nothing Then Exit Sub
            'Me.Client.History.RefreshAppend()
        ElseIf Me.TabControl1.SelectedIndex = Me.TabControl1.TabPages.IndexOf(TabReporting) Then

            DrawOptions()
        ElseIf Me.TabControl1.SelectedIndex = Me.TabControl1.TabPages.IndexOf(Me.tabpKits) Then
            Me.FillKitListview()
        ElseIf Me.TabControl1.SelectedIndex = Me.TabControl1.TabPages.IndexOf(tabpTests) Then
            Me.FillTestListView()
        ElseIf Me.TabControl1.SelectedIndex = Me.TabControl1.TabPages.IndexOf(TabDiscSchemes) Then
            Me.pnlDiscounts.CustomerID = Me.Client.Identity
            Me.pnlDiscounts.DrawDiscounts()
            Me.pnlDiscounts.DrawVolumeDiscounts()
        End If
    End Sub

    'Private Sub cmdFindScheme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindScheme.Click
    '    Dim objFrm As New FrmPricing()
    '    Me.Cursor = Cursors.WaitCursor
    '    Me.cmdFindScheme.Enabled = False
    '    objFrm.Client = Me.Client
    '    If objFrm.ShowDialog = DialogResult.OK Then
    '        Me.DrawDiscounts()
    '    End If
    '    Me.Cursor = Cursors.Default
    '    Me.cmdFindScheme.Enabled = True
    'End Sub

    'Declarations for pricing
    Private Const cstNoAddTimeChargeScheme As Integer = 19
    Private Const cstNoTestChargeScheme As Integer = 359
    Private blnDisableAll As Boolean = True
    Private lvDItem As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry
    Private lvKitItem As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry


    Private Sub FillAPanel(ByVal Panel As MedscreenCommonGui.Panels.DiscountPanel)
        Dim strService As String
        Dim strLineItem As String
        Dim blnVisible As Boolean
        Dim strPanelLabel As String
        strService = CCToolAppMessages.PanelServices.Item(Panel.Name)
        strLineItem = CCToolAppMessages.PanelDiscountLineItems.Item(Panel.Name)
        strPanelLabel = CCToolAppMessages.PanelLabel.Item(Panel.Name)
        blnVisible = (CCToolAppMessages.PanelVisibility.Item(Panel.Name) = "TRUE")
        FillPanel(Panel, strLineItem, strService, blnVisible, strPanelLabel)
        'See if current user can edit prices 
        If Panel.PriceEnabled Then
            If Not MedscreenLib.Glossary.Glossary.UserHasRole("MED_EDIT_PRICES") Then
                Panel.PriceEnabled = False
                Panel.InfoBitmap = Me.PictureBox2.Image
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the individual panels with pricing information
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillPanels()
        'LogTiming("    Fill Panel AO")
        Dim strService As String
        Dim strLineItem As String
        Dim blnVisible As Boolean
        Dim strPanelLabel As String
        pnlSelectCharging.Show()
        FillAPanel(Me.pnlAO)
        FillAPanel(Me.pnlRoutine) ', strLineItem, strService, blnVisible) '"APCOLS002", "ROUTINE")
        'LogTiming("    Fill Panel Callout")

        FillAPanel(Me.pnlCallOut) ', strLineItem, strService, blnVisible) ' "APCOLC002", "CALLOUT")
        'LogTiming("    Fill Panel Fixed")


        FillAPanel(Me.pnlFixedSite) ', strLineItem, strService, blnVisible) ' "APCOLF002", "FIXED")
        Me.pnlExternal.Show()
        'LogTiming("    Fill Panel External")

        FillAPanel(Me.pnlExternal) ', strLineItem, strService, blnVisible) ' "APCOLE002", "EXTERNAL")
        Me.pnlMRO.Show()
        'LogTiming("    Fill Panel 008 Routine")

        FillAPanel(Me.pnlMRO) ', strLineItem, strService, blnVisible) ' "APCOL0008", "ROUTINE")
        Me.pnlInHourMinCharge.Show()
        'LogTiming("    Fill Panel 003 Routine")

        FillAPanel(Me.pnlInHourMinCharge) ', strLineItem, strService, blnVisible) '"APCOL0003", "ROUTINE,CALLOUT")

        Me.pnlOutHoursMC.Show()

        FillAPanel(Me.pnlOutHoursMC) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        Me.dpMCCallOutIn.Show()

        FillAPanel(Me.dpMCCallOutIn) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        Me.dpMCCallOutOut.Show()

        FillAPanel(Me.dpMCCallOutOut) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        'LogTiming("    Fill Panel add time Routine")


        FillAPanel(Me.pnlINHoursAddTimeRoutine) ', strLineItem, strService, blnVisible) '"APCOLS003", "ROUTINE")
        FillAPanel(Me.pnlINHoursAddTimeCO) ', strLineItem, strService, blnVisible) '"APCOLC003", "CALLOUT")
        Me.pnlOutHoursAddTimeCO.Show()
        'LogTiming("    Fill Panel add time Routine")

        FillAPanel(Me.pnlOutHoursAddTimeRoutine) ', strLineItem, strService, blnVisible) '"APCOLS004", "ROUTINE")
        Me.pnlOutHoursAddTimeRoutine.Show()
        'LogTiming("    Fill Panel out co Routine")

        FillAPanel(Me.pnlOutHoursAddTimeCO) ', strLineItem, strService, blnVisible) '"APCOLC004", "CALLOUT")
        'Me.pnlCallOutCharge.Show()
        'LogTiming("    Fill Panel on call ")

        'blnVisible = (CCToolAppMessages.PanelVisibility.Item(Me.pnlCallOutCharge.Name) = "TRUE")
        'FillOptionPanel(Me.pnlCallOutCharge, "ONCALLCHGE", blnVisible)
        'Me.pnlRCOOption.Show()
        'LogTiming("    Fill  option Panel Routine")

        'blnVisible = (CCToolAppMessages.PanelVisibility.Item(Me.pnlRCOOption.Name) = "TRUE")
        'FillOptionPanel(Me.pnlRCOOption, "RTNCOLCHGE", blnVisible)
        Me.pnlTimeAllowHead.Show()
        'LogTiming("    Fill Panel dp call")


        FillAPanel(Me.dpCallOutCancel) ', strLineItem, strService, blnVisible) ' "APCOLC005", "CALLOUT")
        FillAPanel(Me.dpFixedSiteCancellation) ', strLineItem, strService, blnVisible) ' "APCOLF003", "FIXED")

        FillAPanel(Me.dpInHoursCancel) ', strLineItem, strService, blnVisible) '"APCOLS005", "ROUTINE")

        FillAPanel(Me.dpOutOfhoursRoutine) ', strLineItem, strService, blnVisible) ' "APCOLS006", "ROUTINE")
        'LogTiming("    Fill Panel dp in h cancel")
        dpInHoursCancel.PanelLabel = "Routine In-Hours Cancellation"
        dpOutOfhoursRoutine.Show()

        'Routine Collection Charges 

        FillAPanel(Me.dpRoutineCollCharge) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpRoutineCollChargeOut) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpCallOutCollCharge) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpCallOutCollChargeOut) ', strLineItem, strService, blnVisible)
        SetMinChargePanels(MCActionChar, "R")
        SetMinChargePanels(COActionChar, "CO")
        'dynamic panels
        Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet
        For Each myPnlColl In myPanels
            Dim BasePanel As Panels.DiscountPanelBase
            For Each BasePanel In myPnlColl.Panels
                FillAPanel(BasePanel.Panel)
            Next
        Next

        Application.DoEvents()
    End Sub

    Private Sub FillStdPanels()
        Dim strService As String
        Dim strLineItem As String
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlAO.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlAO.Name)
        FillPanel(Me.pnlAO, strLineItem, strService, True) ' "ANOS002", "ANALONLY")
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlRoutine.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlRoutine.Name)
        FillPanel(Me.pnlRoutine, strLineItem, strService, True) ' "COLS002", "ROUTINE")
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlCallOut.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlCallOut.Name)
        FillPanel(Me.pnlCallOut, strLineItem, strService, True) '"COLC002", "CALLOUT")
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlFixedSite.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlFixedSite.Name)
        FillPanel(Me.pnlFixedSite, strLineItem, strService, True) '"COLF002", "FIXED")
        Me.pnlExternal.Hide()
        Me.pnlMRO.Hide()
        Me.pnlInHourMinCharge.Show()
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlInHourMinCharge.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlInHourMinCharge.Name)
        FillPanel(Me.pnlInHourMinCharge, strLineItem, strService, True) ' "COLS001", "ROUTINE")
        Me.pnlOutHoursMC.Hide()
        Me.pnlOutHoursAddTimeCO.Hide()
        Me.pnlOutHoursAddTimeRoutine.Hide()
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlINHoursAddTimeRoutine.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlINHoursAddTimeRoutine.Name)
        FillPanel(Me.pnlINHoursAddTimeRoutine, strLineItem, strService, True) '"COLS003", "ROUTINE")
        strService = CCToolAppMessages.PanelServices.Item(Me.pnlINHoursAddTimeCO.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.pnlINHoursAddTimeCO.Name)
        FillPanel(Me.pnlINHoursAddTimeCO, strLineItem, strService, True) '"COLC003", "ROUTINE")
        'Me.pnlCallOutCharge.Hide()
        'Me.pnlRCOOption.Hide()
        Me.pnlTimeAllowHead.Hide()
        strService = CCToolAppMessages.PanelServices.Item(Me.dpCallOutCancel.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.dpCallOutCancel.Name)
        FillPanel(Me.dpCallOutCancel, strLineItem, strService, True) '"COLC004", "CALLOUT")
        strService = CCToolAppMessages.PanelServices.Item(Me.dpFixedSiteCancellation.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.dpFixedSiteCancellation.Name)
        FillPanel(Me.dpFixedSiteCancellation, strLineItem, strService, True) '"COLF002", "FIXED")
        strService = CCToolAppMessages.PanelServices.Item(Me.dpInHoursCancel.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(Me.dpInHoursCancel.Name)
        FillPanel(Me.dpInHoursCancel, strLineItem, strService, True) '"COLS004", "ROUTINE")
        dpInHoursCancel.PanelLabel = "Routine Cancellation"
        'FillPanel(Me.dpOutOfhoursRoutine, "APCOLS006", "ROUTINE")
        dpOutOfhoursRoutine.Hide()
        pnlSelectCharging.Hide()
        pnlOutHoursMC.Hide()
        strService = CCToolAppMessages.PanelServices.Item(Me.dpCallOutCollCharge.Name)
        strLineItem = CCToolAppMessages.PanelStandardLineItems.Item(dpCallOutCollCharge.Name)
        FillPanel(Me.dpCallOutCollCharge, strLineItem, strService, True, CCToolAppMessages.PanelLabel(dpCallOutCollCharge.Name))  '"COLS003", "ROUTINE")
        dpMCCallOutIn.Hide()
        dpMCCallOutOut.Hide()
        dpRoutineCollCharge.Show()
        dpRoutineCollChargeOut.Hide()
        dpCallOutCollCharge.Hide()
        dpCallOutCollChargeOut.Hide()
        'pnlRCOOption.Hide()
        'pnlCallOutCharge.Hide()
        Application.DoEvents()
        'FillPanel(Me.pnlExternal, "APCOLE002")
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill a specific panel with information 
    ''' </summary>
    ''' <param name="Panel"></param>
    ''' <param name="LineItem"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/06/2006]</date><Action></Action></revision>
    '''     ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillPanel(ByVal Panel As Windows.Forms.Panel, _
        ByVal LineItem As String, ByVal Service As String, _
        ByVal SetVisible As Boolean, Optional ByVal PanelLabel As String = "")


        Dim objTextBox As Windows.Forms.TextBox
        Dim objControl As Windows.Forms.Control
        Dim discount As Double


        Dim objService As Intranet.intranet.customerns.CustomerService
        Dim services As String() = Service.Split(New Char() {","})
        Dim k As Integer
        LogTiming("       Fill Panel Services")
        For k = 0 To services.Length - 1
            objService = Me.Client.Services.Item(services(k))
            If Not objService Is Nothing Then Exit For
        Next

        If TypeOf Panel Is MedscreenCommonGui.Panels.DiscountPanel Then
            If Not SetVisible Then
                Panel.Hide()

                Exit Sub
            Else
                Panel.Show()
            End If
            Dim objPanel As MedscreenCommonGui.Panels.DiscountPanel = Panel
            objPanel.LineItem = LineItem
            'LogTiming("       Fill Panel - Client")
            objPanel.Client = Me.Client
            objPanel.AllowLeave = False
            objPanel.Service = Service
            objPanel.PriceDate = Me.DTPricing.Value
            If PanelLabel.Trim.Length > 0 Then objPanel.PanelLabel = PanelLabel
            'LogTiming("       Fill Panel - redraw")

            objPanel.reDraw(Me.blnDisableAll)
            objPanel.AllowLeave = True
            'LogTiming("       Fill Panel - image")

            objPanel.PictureBoxBitmap = Me.PictureBox1.Image
        End If

        Panel.Tag = objService

        For Each objControl In Panel.Controls
            If TypeOf objControl Is Label Then
                Dim objLabel As Label
                objLabel = objControl
                If objLabel.Name = Me.lblMROComment.Name Then
                    If Client.HasMRO Then
                        Me.lblMROComment.Text = "MRO is " & Client.ReportingInfo.MROID
                        discount = CDbl(Client.InvoiceInfo.Discount(LineItem))
                        If discount = 100 Then
                            Me.lblMROComment.Text += " - No MRO Charge to be applied"
                        Else
                            Me.lblMROComment.Text += " - MRO Charge applied to Referred Samples Only"
                        End If
                    Else
                        Me.lblMROComment.Text = "No MRO set for this customer"
                        Panel.Enabled = False
                    End If
                End If
            End If

        Next
    End Sub

    Private Sub FillOptionPanel(ByVal Panel As Windows.Forms.Panel, _
    ByVal _Option As String, ByVal SetVisible As Boolean)

        'check visiblity of control if invisible no need to waste time on filling it
        If SetVisible Then
            Panel.Show()
        Else
            Panel.Hide()
            Exit Sub
        End If
        Dim objSMOption As Intranet.intranet.customerns.SmClientOption
        If Me.Client Is Nothing Then Exit Sub

        'get option 
        objSMOption = Me.Client.Options.Item(_Option)
        Dim strOptText As String = ""
        'See if we have the option if so display it
        If Not objSMOption Is Nothing Then
            strOptText = objSMOption.Value
        End If
        Dim objControl As Windows.Forms.Control
        Dim objTextBox As Windows.Forms.TextBox
        For Each objControl In Panel.Controls
            If TypeOf objControl Is TextBox Then    'Only one text box present 
                objTextBox = objControl
                objTextBox.Text = strOptText
                objTextBox.Enabled = (Not Me.blnDisableAll)
            End If
        Next

    End Sub




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
    'Private Sub CreateScheme(ByVal Discount As Double, ByVal strType As String, _
    'ByVal ExistingScheme As String, ByVal ItemCode As String, ByVal DiscountType As String)
    '    'We need to create a scheme
    '    Dim strSchemeTitle As String = Math.Abs(Discount).ToString("0.00") & "% "
    '    If Discount < 0 Then
    '        strSchemeTitle += "Surcharge on "
    '    Else
    '        strSchemeTitle += "Discount on "
    '    End If
    '    'Check for discount type 
    '    If DiscountType = "TYPE" Then
    '        strSchemeTitle += "All "
    '    End If
    '    strSchemeTitle += strType
    '    If MsgBox("create discount scheme for " & strSchemeTitle, MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
    '        Dim objNewScheme As New MedscreenLib.Glossary.DiscountScheme()
    '        objNewScheme.Description = strSchemeTitle
    '        objNewScheme.DiscountSchemeId = MedscreenLib.CConnection.NextID("DISCOUNTSCHEME_HEADER", "IDENTITY")
    '        objNewScheme.Insert()

    '        objNewScheme.Update()
    '        MedscreenLib.Glossary.Glossary.DiscountSchemes.Add(objNewScheme)
    '        Dim objNewEntry As New MedscreenLib.Glossary.DiscountSchemeEntry()
    '        objNewEntry.ApplyTo = ItemCode
    '        objNewEntry.Discount = Discount
    '        objNewEntry.DiscountSchemeId = objNewScheme.DiscountSchemeId
    '        objNewEntry.DiscountType = DiscountType
    '        objNewEntry.EntryNumber = 1
    '        objNewEntry.Insert()
    '        objNewEntry.Update()
    '        objNewScheme.Entries.Add(objNewEntry)
    '        PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
    '        PricingSupport.CreateCustomerScheme(objNewScheme, Client)
    '    End If



    'End Sub

    'Private Sub RemoveExistingScheme(ByVal ExistingScheme As String)
    '    If ExistingScheme.Trim.Length > 0 Then      'We need to get rid of existing schemes
    '        Dim Schemes As String() = ExistingScheme.Split(New Char() {",", "-"})
    '        Dim k As Integer
    '        For k = 0 To Schemes.Length - 1     'Could be more than one value 
    '            Dim cd As Intranet.intranet.customerns.CustomerDiscount
    '            Dim strDiscId As String = Schemes.GetValue(k)
    '            If strDiscId.Trim.Length <> 0 Then       'Check it is not a dummy string 
    '                If Char.IsNumber(strDiscId.Chars(0)) Then
    '                    cd = Client.InvoiceInfo.DiscountSchemes.Item(strDiscId, 0)
    '                    If Not cd Is Nothing Then       'We've found the item remove it 
    '                        cd.RemoveScheme()              'Delete from database
    '                        'remove from collection 
    '                        Client.InvoiceInfo.DiscountSchemes.Remove(cd)
    '                    End If
    '                End If
    '            End If
    '        Next
    '    End If
    '    Me.DrawDiscounts()
    'End Sub

    'Private Sub CreateCustomerScheme(ByVal Ds As MedscreenLib.Glossary.DiscountScheme)
    '    Dim CDexist As CustomerDiscount = Client.InvoiceInfo.DiscountSchemes.GetDiscountByScheme(Ds.DiscountSchemeId)
    '    If cdExist Is Nothing Then
    '        Dim cd As New Intranet.intranet.customerns.CustomerDiscount() 'Scheme not in list 
    '        cd.CustomerID = Client.Identity               'Initialise scheme set customer Id
    '        cd.DiscountScheme = Ds                          'Set scheme object
    '        cd.DiscountSchemeId = Ds.DiscountSchemeId       'Set scheme ID
    '        cd.Priority = Me.Client.InvoiceInfo.DiscountSchemes.Count + 1         'Set Priority to 1 + number of schemes
    '        cd.LinkField = "Customer.Identity"

    '        Dim Remembered As Boolean
    '        Dim tmpContAmend As String = Me.strContractAmendment
    '        tmpContAmend = MedscreenLib.Medscreen.GetContractAmendment(Remembered, tmpContAmend)

    '        'Dim objConAmmend As Object = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typString, "Contract Ammendment")
    '        'If Not objConAmmend Is Nothing Then
    '        'cd.ContractAmendment = CStr(objConAmmend)
    '        cd.ContractAmendment = tmpContAmend
    '        'End If
    '        If Remembered Then Me.strContractAmendment = tmpContAmend 'Remember the contract amendment 
    '        cd.Insert()                                     'Insert and update scheme

    '        cd.Update()
    '        Me.ChangeAudit()
    '        Me.Client.InvoiceInfo.DiscountSchemes.Add(cd) 'Add to customer's discount schemes
    '    Else 'if we have a discount we need to remove its end date
    '        CDexist.SetEndDateNull()
    '        CDexist.Update()
    '    End If
    '    Me.DrawDiscounts()
    'End Sub

    Private Sub MROPrice_Leave(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlMRO.PriceChanged
        Dim objTextBox As TextBox
        Dim discount As Double
        Dim Price As Double
        Dim ExistingScheme As String = ""
        If PricingSupport.GetChangeDate Then
            Price = Sender.ListPrice
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
            Else

                Dim ds As MedscreenLib.Glossary.DiscountScheme
                ds = PricingSupport.FindDiscountScheme(discount, "", "MED_REVIEW", Client)
                If Not ds Is Nothing Then 'We have a scheme
                    'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
                    PricingSupport.CreateCustomerScheme(ds, Client)
                    'End If
                Else
                    PricingSupport.CreateScheme(discount, "MRO ", ExistingScheme, "MED_REVIEW", "TYPE", Client)
                End If
            End If
            If blnClientPanelFilled Then
                DoPanelFill
            End If
        End If
    End Sub

    'Private Function Pricingsupport.FindDiscountScheme(ByVal Discount As Double, ByVal ItemCode As String, ByVal SampleType As String) As MedscreenLib.Glossary.DiscountScheme
    '    Dim ds As MedscreenLib.Glossary.DiscountScheme
    '    Dim From As Integer = 0
    '    ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.Pricingsupport.FindDiscountScheme(Discount, ItemCode, SampleType, From)
    '    If ds Is Nothing Then
    '        Return Nothing
    '    Else
    '        Dim strRet As String = ds.FullDescription
    '        Dim intRet As MsgBoxResult = MsgBox("Do you want to use this scheme?" & vbCrLf & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
    '        While intRet = MsgBoxResult.No And Not ds Is Nothing
    '            From += 1
    '            ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.Pricingsupport.FindDiscountScheme(Discount, ItemCode, SampleType, From)
    '            If Not ds Is Nothing Then                   'We may kick out with no scheme
    '                strRet = ds.FullDescription
    '                intRet = MsgBox("Do you want to use this scheme?" & vbCrLf & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
    '            End If
    '        End While
    '        Return ds
    '    End If
    'End Function

    Private Sub ChangeAudit()
        If Me.Client Is Nothing Then Exit Sub 'Exit if no client
        Dim objAudit As Intranet.intranet.customerns.CustomerDataAudit = Me.Client.Audits.Item("PRICES")
        'Check to see if prices are audited
        If objAudit Is Nothing Then     'Doesn't exist craete and populate
            objAudit = New Intranet.intranet.customerns.CustomerDataAudit()
            objAudit.CustomerID = Me.Client.Identity
            objAudit.DataAudited = "PRICES"
            objAudit.DataFields.Insert(CConnection.DbConnection)
        End If
        'Set all values and update 
        objAudit.Audited = False
        objAudit.AuditDate = Now
        objAudit.Comment = "Pricing changed in CCTool"
        objAudit.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
        objAudit.Update()
    End Sub

    Private Sub txtinhrmcPrice_leave(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlInHourMinCharge.PriceChanged, pnlOutHoursMC.PriceChanged, dpMCCallOutIn.PriceChanged, dpMCCallOutOut.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String
            Price = Sender.ListPrice
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            Select Case Sender.Service
                Case "ROUTINE"
                    strType = "Min Charge"
                Case "ROUTINEOUT"
                    strType = "Out hours Min Charge"
                Case "CALLOUT"
                    strType = "Call Out Min Charge"
                Case "CALLOUTOUT"
                    strType = "Out hours Call Out Min Charge"
            End Select
            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
            Else

                Dim ds As MedscreenLib.Glossary.DiscountScheme
                Dim DiscType As String
                Dim myItemCode As String = Sender.LineItem
                'do we want a scheme for this sample type or all samples
                'If MsgBox("Do you want to apply this scheme to all Minimum charge types", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '    DiscType = "TYPE"
                '    myItemCode = "MINCHARGE"
                '    ds = PricingSupport.FindDiscountScheme(discount, "", "MINCHARGE", Client)
                '    ExistingScheme = Me.pnlInHourMinCharge.Scheme & "," & Me.pnlOutHoursMC.Scheme
                '    strType = "Minimum Charges"
                'Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", Client)

                '                End If
                If Not ds Is Nothing Then 'We have a scheme
                    'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
                    PricingSupport.CreateCustomerScheme(ds, Client)
                    'End If
                Else
                    PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, Client)
                End If
            End If
            DoPanelFill

            ChangeAudit()
        End If
    End Sub

    Private Sub txtRoutineCharge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim objTextBox As TextBox
        Dim _Option As String
        If TypeOf sender Is TextBox Then
            objTextBox = sender
            Dim strName As String = Mid(objTextBox.Name.ToUpper, 4)
            Dim newPrice As Double
            If objTextBox.Text.Trim.Length > 0 Then newPrice = CDbl(objTextBox.Text)
            Select Case strName
                Case "ROUTINECHARGE"
                    _Option = "RTNCOLCHGE"

                Case "CALLOUTCHARGE"
                    _Option = "ONCALLCHGE"
            End Select
            Dim objSMOption As Intranet.intranet.customerns.SmClientOption
            If Me.Client Is Nothing Then Exit Sub

            'get option 
            objSMOption = Me.Client.Options.Item(_Option)
            If objSMOption Is Nothing Then
                objSMOption = Client.Options.AddOption(Client.Identity, _Option, System.DBNull.Value, newPrice.ToString("0.00"))
            End If
            objSMOption.Value = newPrice.ToString("0.00")
            objSMOption.Update()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The time allowance option FXDCOLTIME decides whether the customer pays collection time
    '''  If Allowance option set to 0
    '''    Charges are based on time away from home
    '''If Allowance option > 0
    '''    Charges are based on time on site less customer allowance
    '''Else
    '''    No time charge
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetUpTimeAllowance()
        Dim objSMOption As Intranet.intranet.customerns.SmClientOption
        If Me.Client Is Nothing Then Exit Sub

        'get option 
        objSMOption = Me.Client.Options.Item("FXDCOLTIME")
        Me.ckTimeAllowance.Enabled = (Not Me.blnDisableAll)
        Me.txtTimeAllowance.Enabled = True
        If objSMOption Is Nothing Then
            'Me.ckTimeAllowance.Checked = False
            'Me.txtTimeAllowance.Enabled = False

            Me.lblTimeAllowance.Text = "Charges are based on time away from home"
        Else
            'Me.ckTimeAllowance.Checked = True

            Me.txtTimeAllowance.Text = CDbl(objSMOption.Value).ToString("0.0")
            If CDbl(objSMOption.Value) = 0 Then
                Me.lblTimeAllowance.Text = "Charges are based on all time spent on site"
            Else
                Me.lblTimeAllowance.Text = "Charges are based on time on site less customer allowance"
            End If
        End If

        'Check to see if we have discount scheme set
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount = Me.DiscountPresent(Me.cstNoAddTimeChargeScheme)

        If objSMDiscount Is Nothing Then
            Me.ckTimeAllowance.Checked = True
        Else
            Me.ckTimeAllowance.Checked = False
        End If
        'Check that we have sanity
        If Not objSMDiscount Is Nothing AndAlso Not objSMOption Is Nothing Then
            MsgBox("You can't have a customer on the no charges scheme and having a time allowance!", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo)
        End If
        'Set controls 
        Me.pnlINHoursAddTimeCO.PriceEnabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.pnlOutHoursAddTimeCO.PriceEnabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.pnlINHoursAddTimeRoutine.PriceEnabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.pnlOutHoursAddTimeRoutine.PriceEnabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))

    End Sub

    Private Sub FillKitListview()
        Me.lvExKits.BeginUpdate()
        Me.lvExKits.Items.Clear()
        Me.lvExEQP.BeginUpdate()
        Me.lvExEQP.Items.Clear()
        'Get kits sold 
        Dim oCmd As New OleDb.OleDbCommand()
        'Dim kitColl As New Collection()
        Dim objLvTest As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry
        Dim objLineItem As MedscreenLib.Glossary.LineItem

        Dim strKitList As String = CConnection.PackageStringList("pckwarehouse.KitsAndEquipmentWithPrices", "")
        Dim KItList As String() = strKitList.Split(New Char() {","})

        For Each item As String In KItList
            objLineItem = New MedscreenLib.Glossary.LineItem(item)
            'Add in all kits only displaying active ones 
            Dim KitSold As Integer
            Dim avgPrice As Double
            If (Mid(objLineItem.ItemCode, 1, 3) = "KIT") AndAlso (Not objLineItem.Removed) Then
                Dim i As Integer
                Dim objKit As CustomerInvoiceItem
                For i = 1 To Client.InvoiceLineItems.Count - 1
                    objKit = Client.InvoiceLineItems.Item(i)
                    If objKit.ItemCode.ToUpper = objLineItem.ItemCode.ToUpper Then
                        Exit For
                    End If
                    objKit = Nothing
                Next
                If Not objKit Is Nothing Then
                    i = objKit.NoInvoiced
                    KitSold = objKit.NoSold
                    avgPrice = objKit.TotalRevenue
                    If KitSold > 0 Then avgPrice = avgPrice / KitSold
                Else
                    i = 0
                    KitSold = 0
                    avgPrice = 0
                End If
                objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, objLineItem.ItemCode, i, KitSold, avgPrice)
                Me.lvExKits.Items.Add(objLvTest)
            End If
            If (Mid(objLineItem.ItemCode, 1, 3) = "EQP") AndAlso (Not objLineItem.Removed) Then
                Dim i As Integer
                Dim objKit As CustomerInvoiceItem
                For i = 1 To Client.InvoiceLineItems.Count - 1
                    objKit = Client.InvoiceLineItems.Item(i)
                    If objKit.ItemCode.ToUpper = objLineItem.ItemCode.ToUpper Then
                        Exit For
                    End If
                    objKit = Nothing
                Next
                If Not objKit Is Nothing Then
                    i = objKit.NoInvoiced
                    KitSold = objKit.NoSold
                    avgPrice = objKit.TotalRevenue
                    If KitSold > 0 Then avgPrice = avgPrice / KitSold
                Else
                    i = 0
                    KitSold = 0
                    avgPrice = 0
                End If
                objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, objLineItem.ItemCode, i, KitSold, avgPrice)
                Me.lvExEQP.Items.Add(objLvTest)
            End If
        Next
        Me.lvExKits.EndUpdate()
        Me.lvExEQP.EndUpdate()
    End Sub

    Private Sub FillTestListView()

        'Check to see if we have discount scheme set
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount = DiscountPresent(Me.cstNoTestChargeScheme)

        'Check enabling 
        Me.ckChargeAdditionalTests.Enabled = (Not Me.blnDisableAll)
        Me.lvTests.BeginUpdate()
        If (objSMDiscount Is Nothing) And (Not Me.blnDisableAll) Then
            Me.ckChargeAdditionalTests.Checked = True
            Me.lvTests.Enabled = True
            Me.lvTestConfirm.Enabled = True
        Else
            Me.ckChargeAdditionalTests.Checked = False
            Me.lvTests.Enabled = False
            Me.ckChargeAdditionalTests.Checked = False
        End If


        Me.lvTests.Items.Clear()
        Dim testString As String = MedscreenLib.Glossary.Glossary.LineItems.LineItemsByType("TEST")
        Dim testStringArray As String() = testString.Split(New Char() {","})
        Dim objLvTest As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry
        Dim i As Integer
        For i = 0 To testStringArray.Length - 1
            objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, testStringArray.GetValue(i))
            Me.lvTests.Items.Add(objLvTest)
        Next
        Me.lvTests.EndUpdate()

        Me.lvTestConfirm.BeginUpdate()
        Me.lvTestConfirm.Items.Clear()
        testString = MedscreenLib.Glossary.Glossary.LineItems.LineItemsByType("TEST_CONF")
        testStringArray = testString.Split(New Char() {","})
        For i = 0 To testStringArray.Length - 1
            objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, testStringArray.GetValue(i))
            Me.lvTestConfirm.Items.Add(objLvTest)
        Next
        Me.lvTestConfirm.EndUpdate()


        'objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, "APTST0002")
        'Me.lvTests.Items.Add(objLvTest)
        'objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, "APTST0003")
        'Me.lvTests.Items.Add(objLvTest)
        'objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, "APTST0004")
        'Me.lvTests.Items.Add(objLvTest)
        'objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, "APTST0005")
        'Me.lvTests.Items.Add(objLvTest)
        'objLvTest = New MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry(Me.Client, "APTST0006")
        'Me.lvTests.Items.Add(objLvTest)
    End Sub

    Private Sub LvTestsSubItem_clicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs) Handles lvTests.SubItemClicked, lvTestConfirm.SubItemClicked
        Dim lvItem As ListViewItem = e.Item

        Dim lvSubItem As Integer = e.SubItem
        lvDItem = Nothing
        'Dim textbox1 As TextBox = New TextBox()
        If TypeOf e.Item Is MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry Then
            lvDItem = e.Item
        Else
            Exit Sub
        End If
        'Clear list box items
        Me.cbEditor.Items.Clear()
        Me.cbEditor.Items.Add(0)
        Dim Editors As Control() = New Control() {Me.cbEditor}
        Dim i As Integer = 1
        Dim priceString As String
        If lvSubItem > 5 Then
            priceString = CCToolAppMessages.TestPrices.Item(lvDItem.LineItem.ItemCode & "-" & CStr(i).Trim)
            While Not priceString Is Nothing AndAlso priceString.Trim.Length > 0
                Me.cbEditor.Items.Add(priceString)
                i = i + 1
                priceString = CCToolAppMessages.TestPrices.Item(lvDItem.LineItem.ItemCode & "-" & CStr(i).Trim)
            End While
            'If lvDItem.LineItem.ItemCode = "APTST0002" Then
            '    Me.cbEditor.Items.Add("4.85")
            'End If
            'If lvDItem.LineItem.ItemCode = "APTST1000" Then
            '    Me.cbEditor.Items.Add("35")
            'End If
            Dim strPrice As String = Client.InvoiceInfo.Price(lvDItem.LineItem.ItemCode).ToString("0.00")

            Me.cbEditor.Items.Add(strPrice)

            sender.StartEditing(Editors(0), e.Item, e.SubItem)
        End If
    End Sub

    Private Sub LvKitsSubItem_clicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs) Handles lvExKits.SubItemClicked, lvExEQP.SubItemClicked
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
        If lvSubItem > 5 Then
            Dim strPrice As String = Client.InvoiceInfo.Price(lvKitItem.LineItem.ItemCode).ToString("0.00")

            Me.txtKitEdit.Text = strPrice

            sender.StartEditing(Editors(0), e.Item, e.SubItem)
        End If
    End Sub


    Private Function DiscountPresent(ByVal DiscountScheme As Integer) As Intranet.intranet.customerns.CustomerDiscount
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount
        For Each objSMDiscount In Me.Client.InvoiceInfo.DiscountSchemes
            If objSMDiscount.DiscountSchemeId = DiscountScheme Then
                Exit For
            End If
            objSMDiscount = Nothing
        Next
        Return objSMDiscount

    End Function

    Private Sub ckChargeAdditionalTests_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckChargeAdditionalTests.CheckedChanged
        'Need to see if the discount is present and add/delete as appropriate 
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount = DiscountPresent(Me.cstNoTestChargeScheme)
        Me.ckChargeAdditionalTests.Enabled = False
        Me.Cursor = Cursors.WaitCursor

        If Me.ckChargeAdditionalTests.Checked Then
            If Not objSMDiscount Is Nothing Then    'Need to remove scheme
                objSMDiscount.Delete()
                Me.Client.InvoiceInfo.DiscountSchemes.Remove(objSMDiscount)

            End If
        Else    'Don't charge we need to have the scheme
            If objSMDiscount Is Nothing Then
                objSMDiscount = New Intranet.intranet.customerns.CustomerDiscount()
                With objSMDiscount
                    .CustomerID = Me.Client.Identity
                    .DiscountSchemeId = Me.cstNoTestChargeScheme
                    .Priority = Me.Client.InvoiceInfo.DiscountSchemes.Count + 1
                    .Insert()
                    .Update()
                End With
                Me.Client.InvoiceInfo.DiscountSchemes.Add(objSMDiscount)
            End If
        End If
        Me.FillTestListView()
        Me.FillKitListview()
        Me.Cursor = Cursors.Default
        Me.ckChargeAdditionalTests.Enabled = (Not Me.blnDisableAll)
    End Sub


    Private Sub ckTimeAllowance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckTimeAllowance.CheckedChanged
        'Need to see if the discount is present and add remove it as appropriate 
        'Need to see if the discount is present and add/delete as appropriate 
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount = DiscountPresent(Me.cstNoAddTimeChargeScheme)
        Me.ckTimeAllowance.Enabled = False
        Me.Cursor = Cursors.WaitCursor

        If Me.ckTimeAllowance.Checked Then
            If Not objSMDiscount Is Nothing Then    'Need to remove scheme
                objSMDiscount.Delete()
                Me.Client.InvoiceInfo.DiscountSchemes.Remove(objSMDiscount)

            End If
        Else    'Don't charge we need to have the scheme
            If objSMDiscount Is Nothing Then
                objSMDiscount = New Intranet.intranet.customerns.CustomerDiscount()
                With objSMDiscount
                    .CustomerID = Me.Client.Identity
                    .DiscountSchemeId = Me.cstNoAddTimeChargeScheme
                    .Priority = Me.Client.InvoiceInfo.DiscountSchemes.Count + 1
                    .Insert()
                    .Update()
                End With
                Me.Client.InvoiceInfo.DiscountSchemes.Add(objSMDiscount)
            End If
        End If
        SetUpTimeAllowance()
        If blnClientPanelFilled Then DoPanelFill()
        Me.Cursor = Cursors.Default
        Me.ckTimeAllowance.Enabled = (Not Me.blnDisableAll)

    End Sub

    Private Sub cbEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbEditor.Leave
        Dim strText As String = Me.cbEditor.Text
        Me.Cursor = Cursors.WaitCursor
        If Not lvDItem Is Nothing Then
            Dim strTest = lvDItem.LineItem.ItemCode
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim Discount As Double
            Dim dblPrice As Double = Me.Client.InvoiceInfo.Price(lvDItem.LineItem.ItemCode)
            If CDbl(strText) = 0 Then
                Discount = 100

            ElseIf CDbl(strText) <> dblPrice Then
                Dim tmpPrice As Double = CDbl(strText)
                Discount = (dblPrice - tmpPrice) / dblPrice * 100
            Else
                PricingSupport.RemoveExistingScheme(lvDItem.DiscountScheme, Client)
                lvDItem.Refresh()
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            'do we want a scheme for this sample type or all samples
            DiscType = "CODE"
            ds = PricingSupport.FindDiscountScheme(Discount, strTest, "", Client)
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                PricingSupport.CreateCustomerScheme(ds, Client)
                'End 'If
            Else
                PricingSupport.CreateScheme(Discount, lvDItem.LineItem.Description, "", strTest, DiscType, Client)
            End If
            lvDItem.Refresh()
            Me.Cursor = Cursors.Default
        End If
    End Sub


    Private Sub txtKitEdit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKitEdit.Leave
        If PricingSupport.GetChangeDate Then
            Dim strText As String = Me.txtKitEdit.Text
            Me.Cursor = Cursors.WaitCursor
            If Not lvKitItem Is Nothing Then
                Dim strTest = lvKitItem.LineItem.ItemCode
                Dim ds As MedscreenLib.Glossary.DiscountScheme
                Dim DiscType As String
                Dim Discount As Double
                If CDbl(strText) = 0 Then
                    Discount = 100
                ElseIf CDbl(strText) = Me.Client.InvoiceInfo.Price(strTest) Then
                    PricingSupport.RemoveExistingScheme(lvKitItem.DiscountScheme, Client)
                    lvKitItem.Refresh()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                Else
                    Dim Price As Double = Me.Client.InvoiceInfo.Price(strTest)
                    Dim NewPrice As Double = CDbl(strText)
                    Discount = (Price - NewPrice) * 100 / Price

                End If
                'do we want a scheme for this sample type or all samples
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(Discount, strTest, "", Client)
                If Not ds Is Nothing Then 'We have a scheme
                    If lvKitItem.DiscountScheme = ds.DiscountSchemeId Then 'Matches existing scheme
                        Exit Sub

                    End If
                    'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    PricingSupport.CreateCustomerScheme(ds, Client)
                    'End If
                Else
                    PricingSupport.CreateScheme(Discount, lvKitItem.LineItem.Description, "", strTest, DiscType, Client)
                End If
                lvKitItem.Refresh()
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    'Deal with setting time allowance 
    Private Sub txtTimeAllowance_leave(ByVal Sender As Object, ByVal e As EventArgs) Handles txtTimeAllowance.Leave
        Dim objSMOption As Intranet.intranet.customerns.SmClientOption
        If Me.Client Is Nothing Then Exit Sub

        'get option 
        objSMOption = Me.Client.Options.Item("FXDCOLTIME")
        If objSMOption Is Nothing AndAlso Me.txtTimeAllowance.Text.Trim.Length > 0 Then
            'Need to create an option 
            objSMOption = Me.Client.Options.AddOption(Me.Client.Identity, "FXDCOLTIME", True, "")
            objSMOption.Value = Me.txtTimeAllowance.Text
            objSMOption.Update()
            If CDbl(objSMOption.Value) = 0 Then
                Me.lblTimeAllowance.Text = "Charges are based on time away from home"
            Else
                Me.lblTimeAllowance.Text = "Charges are based on time on site less customer allowance"
            End If
        ElseIf objSMOption Is Nothing Then
            Me.lblTimeAllowance.Text = ""
        Else
            If Me.txtTimeAllowance.Text.Trim.Length = 0 Then
                'Need to delete option 
                objSMOption.Delete()
                Me.Client.Options.Remove(objSMOption)
                Me.lblTimeAllowance.Text = ""
            Else
                'Me.ckTimeAllowance.Checked = True
                objSMOption.Value = Me.txtTimeAllowance.Text
                objSMOption.Update()
                If CDbl(objSMOption.Value) = 0 Then
                    Me.lblTimeAllowance.Text = "Charges are based on time away from home"
                Else
                    Me.lblTimeAllowance.Text = "Charges are based on time on site less customer allowance"
                End If
            End If
        End If
        SetUpTimeAllowance()
    End Sub

    Private strAddService As String = ""
    Dim strCurrScheme As String = "" 'Hold any scheme in use

    Private Sub ctxAddService_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctxAddService.Popup
        Dim objPanel As Panel
        Dim objctxMenu As ContextMenu

        strCurrScheme = ""
        If TypeOf sender Is ContextMenu Then
            objctxMenu = sender
            If TypeOf objctxMenu.SourceControl Is Panel Then
                objPanel = objctxMenu.SourceControl
                Dim objService As Intranet.intranet.customerns.CustomerService
                If TypeOf objPanel Is MedscreenCommonGui.Panels.DiscountPanel Then
                    strAddService = CType(objPanel, MedscreenCommonGui.Panels.DiscountPanel).Service
                    objService = CType(objPanel, MedscreenCommonGui.Panels.DiscountPanel).PanelService
                    strCurrScheme = CType(objPanel, MedscreenCommonGui.Panels.DiscountPanel).Scheme
                End If

                If objPanel.Tag Is Nothing Then
                    Me.mnuctxAddService.Text = "&Add Service"
                Else
                    If Not objService Is Nothing Then
                        If objService.Removed Then
                            Me.mnuctxAddService.Text = "&Add Service"
                        Else
                            Me.mnuctxAddService.Text = "&Remove Service"
                        End If
                    Else
                        Me.mnuctxAddService.Text = "&Remove Service"
                    End If
                End If

                If strCurrScheme.Trim.Length > 0 Then
                    Me.mnuCTXDiscountDates.Visible = True
                Else
                    Me.mnuCTXDiscountDates.Visible = False
                End If
            End If
        End If
    End Sub

    Private Sub mnuCTXDiscountDates_click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCTXDiscountDates.Click
        If strCurrScheme.Trim.Length = 0 Then Exit Sub
        Dim Schemes As String() = strCurrScheme.Split(New Char() {",", "-"})
        Dim k As Integer
        If Schemes.Length > 0 Then     'Could be more than one value 
            Dim cd As Intranet.intranet.customerns.CustomerDiscount
            Dim strDiscId As String = Schemes.GetValue(0)
            If strDiscId.Trim.Length <> 0 Then       'Check it is not a dummy string 
                cd = Client.InvoiceInfo.DiscountSchemes.Item(strDiscId, 0)
                If Not cd Is Nothing Then       'We've found the item remove it 
                    Dim objForm As New MedscreenCommonGui.frmCustomerService()
                    objForm.CustomerDiscount = cd
                    If objForm.ShowDialog = DialogResult.OK Then
                        cd = objForm.CustomerDiscount
                        cd.Update()
                        Me.DoPanelFill()
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub mnuctxAddService_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuctxAddService.Click
        Dim objMenu As MenuItem
        objMenu = sender
        MsgBox(Me.strAddService & " " & objMenu.Text)
        Dim objSer As Intranet.intranet.customerns.CustomerService
        If InStr(objMenu.Text.ToUpper, "ADD ") > 0 Then 'Adding Service 
            objSer = Client.Services.Item(Me.strAddService)
            If objSer Is Nothing Then
                objSer = New Intranet.intranet.customerns.CustomerService()
                objSer.CustomerId = Me.Client.Identity
                objSer.Service = Me.strAddService
            Else
                objSer.Removed = False
            End If
            objSer.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            Me.Client.Services.Add(objSer)
            Dim frmSer As New frmCustomerService()
            frmSer.CustomerService = objSer
            If frmSer.ShowDialog = DialogResult.OK Then
                objSer = frmSer.CustomerService
                objSer.ModifiedOn = Now
                objSer.update()
            End If
            CommonForms.NotifyServiceChange(Client, objSer.Service, "Add")
        ElseIf InStr(objMenu.Text.ToUpper, "RESET ") > 0 Then
            objSer = Client.Services.Item(Me.strAddService)
            objSer.NullDates()
            objSer.ModifiedOn = Now
            objSer.update()
            CommonForms.NotifyServiceChange(Client, objSer.Service, "Reinstated")

        Else                                                'Removing Service
            objSer = Client.Services.Item(Me.strAddService)
            If Not objSer Is Nothing Then
                objSer.Removed = True
                objSer.update()
            End If
            CommonForms.NotifyServiceChange(Client, objSer.Service, "Removed")

        End If
        'Redraw the prices 
        DrawPrices()
    End Sub



    Private Sub pnlAO_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlAO.PriceChanged, pnlCallOut.PriceChanged, _
          pnlExternal.PriceChanged, pnlFixedSite.PriceChanged, pnlRoutine.PriceChanged


        'Dim newPrice As Double = objTextBox.Text
        Dim discount As Double
        Dim Price As Double
        Dim ExistingScheme As String = ""
        Dim strType As String
        Price = Sender.ListPrice
        discount = (Price - NewPrice) * 100 / Price
        Sender.Discount = discount.ToString("0.00")
        ExistingScheme = Sender.DiscountScheme
        strType = PricingSupport.IdentifySenderService(Sender.Service)

        'Check to see if killing scheme
        Dim ds As MedscreenLib.Glossary.DiscountScheme = Nothing
        Dim DiscType As String = ""
        Dim myItemCode As String = Sender.LineItem
        If PricingSupport.GetChangeDate Then
            Debug.WriteLine(PricingSupport.ChangeDate)
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
                ExistingScheme = ""

            Else

                'do we want a scheme for this sample type or all samples
                'If MsgBox("Do you want to use this price for all Sample types", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '    DiscType = "TYPE"
                '    myItemCode = "SAMPLE"
                '    ds = PricingSupport.FindDiscountScheme(discount, "", "SAMPLE", Client)
                '    ExistingScheme = Me.pnlAO.Scheme & "," & Me.pnlCallOut.Scheme & "," & _
                '        Me.pnlExternal.Scheme & "," & Me.pnlFixedSite.Scheme & "," & Me.pnlRoutine.Scheme
                'Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", Client)

                'End If
            End If
            PricingSupport.PriceHasChanged(discount, ExistingScheme, myItemCode, Client, strType, ds, DiscType)

            Me.blnPriceChanged = True                   'Record that a price has changed 

            DoPanelFill
        End If
    End Sub

    Private Sub dpOutOfhoursRoutine_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles dpInHoursCancel.PriceChanged, dpCallOutCancel.PriceChanged, dpFixedSiteCancellation.PriceChanged, dpOutOfhoursRoutine.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String = " Collection Cancellation"
            Price = Sender.ListPrice
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim myItemCode As String = Sender.LineItem

            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
            Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", Client)

            End If
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
                PricingSupport.CreateCustomerScheme(ds, Client)
                'End If
            Else
                PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, Client)
            End If
            If blnClientPanelFilled Then
               DoPanelFill
            End If
        End If
    End Sub

    'Collection charge management
    Private Sub dpRoutineCollCharge_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles dpRoutineCollCharge.PriceChanged, dpRoutineCollChargeOut.PriceChanged, dpCallOutCollCharge.PriceChanged, dpCallOutCollChargeOut.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String = " Attendance Fee"
            Price = Sender.ListPrice
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim myItemCode As String = Sender.LineItem
            Dim objLineItem As MedscreenLib.Glossary.LineItem = MedscreenLib.Glossary.Glossary.LineItems.Item(Sender.LineItem)
            ' if we have a line item then change the type text
            If Not objLineItem Is Nothing Then
                strType = objLineItem.Description & " Charge"
            End If

            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
            Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", Client)

            End If
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
                PricingSupport.CreateCustomerScheme(ds, Client)
                'End If
            Else
                PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, Client)
            End If
            If blnClientPanelFilled Then
                DoPanelFill
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Send password letter to customer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdPasswordLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPasswordLetter.Click
        'Send pasword letter to customer
        Dim frmPass As New frmPrintPassword()
        frmPass.Client = Me.Client
        Dim myTextFile As Reporter.clsTextFile
        Dim objReporter As Reporter.Report = New Reporter.Report()

        'objReporter.PurgeOldFiles()
        'Need to set email body file from config settings
        Dim mySeetings As System.Configuration.ConfigurationSettings
        objReporter.ReportSettings.EmailBodyFile = MedscreenLib.Constants.PasswordBodyFile
        If frmPass.ShowDialog = DialogResult.OK Then
            myTextFile = Me.Client.ReportingInfo.PasswordLetter(frmPass.ConfirmMethod, Me.txtEmailPassword.Text, Now, frmPass.txtOutput.Text)
            Dim strPath As String = "X:\Lab Programs\DBReports\Templates|X:\Lab Programs\DBReports\Templates\Scanned Documents"


            Dim PathArray As String() = strPath.Split(New Char() {"|"})
            objReporter.ReportSettings.SetTemplatePaths(PathArray)

            objReporter.SetToCustomerDirectory()
            objReporter.ReportSettings.DocumentPath += Me.Client.SMID & "\"
            If Not IO.Directory.Exists(objReporter.ReportSettings.DocumentPath) Then
                IO.Directory.CreateDirectory(objReporter.ReportSettings.DocumentPath)
            End If
            Dim strDocName As String = frmPass.ConfirmMethod.ToString & "-" & Me.Client.Profile & "-" & Now.ToString("yyyyMMddHHmm") & ".doc"
            Dim objCReport As Reporter.clsWordReport = objReporter.GetDatafile(myTextFile.ExportStringArray, strDocName)
            objReporter.WordApplication.Visible = (frmPass.OutputAs = frmPrintPassword.OutputOption.Printer)
            Intranet.Support.AdditionalDocumentColl.CreateCustomerDocument(Me.Client.Identity, _
             objReporter.ReportSettings.DocumentPath & strDocName, "PASSWORD")
            objReporter.WordApplication.ActiveDocument.Save()

            If frmPass.OutputAs = frmPrintPassword.OutputOption.Email Then
                objReporter.WordApplication.ActiveDocument.Save()

                Dim objRecip As Reporter.clsRecipient = New Reporter.clsRecipient()
                objRecip.address = Me.Client.EmailResults.Replace(";", ",")
                objReporter.WordApplication.ActiveDocument.Save()

                objCReport.SendByEmail(objRecip)
            End If


            'Okay deal with send log
            Dim objSendLog As MedscreenLib.SendLogInfo = MedscreenLib.SendLogCollection.CreateSendLog()
            objSendLog.Filename = strDocName
            objSendLog.ReferenceID = Me.Client.Identity
            objSendLog.ReferenceTo = "PASSWORD"
            objSendLog.SentOn = Now
            objSendLog.SendingProgram = "CCTOOL"
            objSendLog.SendMethod = "DISK"
            objSendLog.Status = "C"
            objSendLog.Update()
            'Now we need to set up a comment.
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, Me.Client.Identity, "DOC", frmPass.ConfirmMethod.ToString & " created", _
MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)

            'Me.Client.CustomerComments.CreateComment(frmPass.ConfirmMethod.ToString & " created", "DOC", jobs.CollectionCommentList.CommentType.Customer)

        End If
    End Sub

    Private Sub udRoutine_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udRoutine.Leave, udCallOut.Leave
        PricingSupport.ChangeDate = Today
        Dim udControl As DomainUpDown = sender
        Dim tag As String = udControl.Tag
        Dim ActionChar As Char = udControl.Text.Chars(0)
        'Decide on action
        If Client Is Nothing Then Exit Sub
        Dim blnSchemeneedsChanging As Boolean = False
        Dim strChanges As String = "Discounts for "
        Dim strExistingSchemes As String = ""
        Dim strNewSchemes As String = ""
        If tag = "R" Then
            Me.MCActionChar = ActionChar
            'See what we have in the way of discounts already then check that they match the action character.
            Dim inRoutineMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.pnlInHourMinCharge.LineItem)
            Dim outRoutineMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.pnlOutHoursMC.LineItem)
            Dim inRoutineCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpRoutineCollCharge.LineItem)
            Dim outRoutineCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpRoutineCollChargeOut.LineItem)
            'If sample pricing all discounts need to be set to zero 
            'which means scheme_ids 504,536,498,499
            'If minimum charge scheme ids 498,499
            'If collection charge ids 504,536
            If Me.MCActionChar = "M" Then 'Dealing with Minimum charges
                If InStr(inRoutineCCDisc, "498") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Attendance fee"
                    Dim Schemes As String() = inRoutineCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "498,"
                End If
                If InStr(outRoutineCCDisc, "499") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Attendance fee"
                    Dim Schemes As String() = outRoutineCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "499,"
                End If
                strChanges += " need 100% discounts"
            ElseIf Me.MCActionChar = "A" Then             'Dealing with Minimum charges
                If InStr(inRoutineMCDisc, "504") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Minimum charge"
                    Dim Schemes As String() = inRoutineMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "504,"
                End If
                If InStr(outRoutineMCDisc, "536") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Minimum charge"
                    Dim Schemes As String() = outRoutineMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "536,"
                End If
                strChanges += " need 100% discounts"
            ElseIf Me.MCActionChar = "S" Then             'Dealing with Minimum charges
                If InStr(inRoutineCCDisc, "498") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Attendance fee"
                    Dim Schemes As String() = inRoutineCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "498,"
                End If
                If InStr(outRoutineCCDisc, "499") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Attendance fee"
                    Dim Schemes As String() = outRoutineCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "499,"
                End If
                If InStr(inRoutineMCDisc, "504") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Minimum charge"
                    Dim Schemes As String() = inRoutineMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "504,"
                End If
                If InStr(outRoutineMCDisc, "536") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Minimum charge"
                    Dim Schemes As String() = outRoutineMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "536,"
                End If
                strChanges += " need 100% discounts"




            End If
        Else
            Me.COActionChar = ActionChar
            Dim inCallOutMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpMCCallOutIn.LineItem)
            Dim outCallOutMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpMCCallOutOut.LineItem)
            Dim inCallOutCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpCallOutCollCharge.LineItem)
            Dim outCallOutCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpCallOutCollChargeOut.LineItem)
            'If sample pricing all discounts need to be set to zero 
            'which means scheme_ids 507,508,529,503
            'If minimum charge scheme ids 529,503
            'If collection charge ids 507,508
            If Me.COActionChar = "M" Then 'Dealing with Minimum charges
                If InStr(inCallOutCCDisc, "529") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Attendance fee"
                    Dim Schemes As String() = inCallOutCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "529,"
                End If
                If InStr(outCallOutCCDisc, "503") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Attendance fee"
                    Dim Schemes As String() = outCallOutCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","

                    strNewSchemes += "503,"
                End If
                strChanges += " need 100% discounts"

            ElseIf Me.COActionChar = "A" Then 'Dealing with Minimum charges
                If InStr(inCallOutMCDisc, "507") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Minimum charge"
                    Dim Schemes As String() = inCallOutMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "507,"
                End If
                If InStr(outCallOutMCDisc, "508") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Minimum charge"
                    Dim Schemes As String() = outCallOutMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","

                    strNewSchemes += "508,"
                End If
                strChanges += " need 100% discounts"
            ElseIf Me.COActionChar = "S" Then         'Dealing with Minimum charges
                If InStr(inCallOutCCDisc, "529") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Attendance fee"
                    Dim Schemes As String() = inCallOutCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "529,"
                End If
                If InStr(outCallOutCCDisc, "503") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Attendance fee"
                    Dim Schemes As String() = outCallOutCCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","

                    strNewSchemes += "503,"
                End If
                If InStr(inCallOutMCDisc, "507") = 0 Then
                    blnSchemeneedsChanging = True
                    strChanges += "in hours Minimum charge"
                    Dim Schemes As String() = inCallOutMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","
                    strNewSchemes += "507,"
                End If
                If InStr(outCallOutMCDisc, "508") = 0 Then
                    blnSchemeneedsChanging = True
                    If strChanges.Length > 15 Then strChanges += ", "
                    strChanges += "out hours Minimum charge"
                    Dim Schemes As String() = outCallOutMCDisc.Split(New Char() {",", "-"})
                    strExistingSchemes += Schemes.GetValue(0) & ","

                    strNewSchemes += "508,"
                End If
                strChanges += " need 100% discounts"

            End If
        End If
        If blnSchemeneedsChanging Then
            Dim dlgRet As MsgBoxResult = MsgBox(strChanges & " create them?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If dlgRet = MsgBoxResult.Yes Then
                PricingSupport.RemoveExistingScheme(strExistingSchemes, Client)
                Dim newIds As String() = strNewSchemes.Split(New Char() {","})
                Dim i As Integer
                Dim ds As MedscreenLib.Glossary.DiscountScheme
                For i = 0 To newIds.Length - 1
                    Dim newID As String = newIds.GetValue(i)
                    If newID.Trim.Length > 0 Then
                        ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.Item(CInt(newID), 1)
                        If Not ds Is Nothing Then PricingSupport.CreateCustomerScheme(ds, Client) 'Check that discount scheme exists
                    End If
                Next
                Me.DoPanelFill()
            End If
        End If
        Debug.WriteLine("")
        SetMinChargePanels(ActionChar, tag)

    End Sub

    Private Sub SetMinChargePanels(ByVal ActionChar As Char, ByVal tag As String)
        Select Case ActionChar
            Case "A"        'Collectioncharge
                If tag = "R" Then 'Routine payments set all coll charge panels to enabled mincharge disabled
                    Me.dpRoutineCollCharge.Enabled = True
                    Me.dpRoutineCollChargeOut.Enabled = True
                    Me.pnlInHourMinCharge.Enabled = False
                    Me.pnlOutHoursMC.Enabled = False
                ElseIf tag = "CO" Then 'Routine payments set all coll charge panels to enabled mincharge disabled
                    Me.dpMCCallOutIn.Enabled = False
                    Me.dpMCCallOutOut.Enabled = False
                    Me.dpCallOutCollCharge.Enabled = True
                    Me.dpCallOutCollChargeOut.Enabled = True
                End If
            Case "M"        'Minimum charge
                If tag = "R" Then 'Routine payments set all coll charge panels to disabled mincharge enabled
                    Me.dpRoutineCollCharge.Enabled = False
                    Me.dpRoutineCollChargeOut.Enabled = False
                    Me.pnlInHourMinCharge.Enabled = True
                    Me.pnlOutHoursMC.Enabled = True
                ElseIf tag = "CO" Then 'Call out payments set all coll charge panels to disabled mincharge enabled
                    Me.dpMCCallOutIn.Enabled = True
                    Me.dpMCCallOutOut.Enabled = True
                    Me.dpCallOutCollCharge.Enabled = False
                    Me.dpCallOutCollChargeOut.Enabled = False
                End If
            Case "S"        'sample based
                If tag = "R" Then       'Routine payments set all routine panels to disabled
                    Me.dpRoutineCollCharge.Enabled = False
                    Me.dpRoutineCollChargeOut.Enabled = False
                    Me.pnlInHourMinCharge.Enabled = False
                    Me.pnlOutHoursMC.Enabled = False
                ElseIf tag = "CO" Then  'Set all call out panels to disabled
                    Me.dpMCCallOutIn.Enabled = False
                    Me.dpMCCallOutOut.Enabled = False
                    Me.dpCallOutCollCharge.Enabled = False
                    Me.dpCallOutCollChargeOut.Enabled = False
                End If
        End Select

    End Sub

    Private Sub cmdAddTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddTest.Click
        Dim myAnalysis As New MedscreenLib.Glossary.PhraseCollection("lib_cctool.AnalysisList", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction)
        Dim oAnalysis As Object = Medscreen.GetParameter(myAnalysis, "Test")
        If Not oAnalysis Is Nothing Then
            'We have a test we need to see if it is already in list
            If Me.Client.TestSchedule.IndexofTest(CStr(oAnalysis)) > -1 Then
                MsgBox("Test " & CStr(oAnalysis) & " is already in schedule", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                Exit Sub
            End If
            'Ok we can create a new test
            RemoveCustTestPanels()
            Dim objTest As New Intranet.intranet.TestResult.TestScheduleEntry()
            objTest.AnalysisID = CStr(oAnalysis)
            objTest.StandardTest = True
            objTest.IsAnalysis = True
            objTest.OrderNumber = Me.Client.TestSchedule.Count + 1
            objTest.ReplicateCount = 1
            objTest.TestPanelID = Me.Client.TestSchedule.TestScheduleId
            objTest.Fields.Insert(CConnection.DbConnection)
            Me.Client.TestSchedule.Clear()
            Me.Client.TestSchedule.Load()
            Me.Client.TestSchedule.Renumber()
            Me.DrawTestPanels()
            Me.DrawTestSchedules()

        End If
    End Sub

    Private Sub RemoveCustTestPanels()

        Dim oCmd As New OleDb.OleDbCommand()
        Dim intRet As Integer
        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "Delete from cust_test_panel where customer_id = '" & Me.Client.Identity & "'"

            If CConnection.ConnOpen Then            'Attempt to open reader
                intRet = oCmd.ExecuteNonQuery
                oCmd.CommandText = "delete from tos_cust_panel where customer_id = '" & Me.Client.Identity & "'"
                intRet = oCmd.ExecuteNonQuery
            End If
            Me.Client.TestPanels.Clear()
            Me.Client.TestPanels = Nothing
            If Me.Client.TestPanelData.Count = 1 Then
                Me.Client.TestPanelData.Item(0).Update()
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "LVRoutinePayment-RemoveCustTestPanels-8903")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try
    End Sub

    Private Sub lvTestSchedules_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvTestSchedules.DoubleClick
        If Me.myTestSchedule Is Nothing Then Exit Sub
        'Now we can edit the test schedule, will be able to change the standard and no of replicates
        Dim testForm As New MedscreenCommonGui.frmTestShedEntry()
        testForm.TestScheduleEntry = Me.myTestSchedule
        If testForm.ShowDialog = DialogResult.OK Then
            Me.myTestSchedule = testForm.TestScheduleEntry
            Me.myTestSchedule.DoUpdate()
            Me.DrawTestSchedules()
        End If
    End Sub

    Private Sub lvTestSchedules_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvTestSchedules.SelectedIndexChanged

    End Sub

    Private Sub cmdDeleteTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteTest.Click
        If Me.myTestSchedule Is Nothing Then Exit Sub
        Dim intIndex As Integer = Me.Client.TestSchedule.IndexofTest(myTestSchedule.AnalysisID)
        If intIndex > -1 Then
            Dim objTp As Intranet.intranet.TestResult.TestScheduleEntry = Me.Client.TestSchedule.Item(intIndex)
            objTp.Fields.Delete(CConnection.DbConnection)
            Me.Client.TestSchedule.RemoveAt(intIndex)
            Me.Client.TestSchedule.Renumber()
            Me.RemoveCustTestPanels()
            Me.DrawTestSchedules()
            Me.DrawTestPanels()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a discount scheme
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	23/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdAddScheme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddScheme.Click
    '    Dim objSchemeId As Integer
    '    objSchemeId = Me.cbVolScheme.SelectedValue
    '    Dim objScheme As New Intranet.intranet.customerns.CustomerVolumeDiscount()
    '    objScheme.CustomerID = Me.Client.Identity
    '    objScheme.DiscountSchemeId = objSchemeId
    '    objScheme.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
    '    objScheme.LinkField = "Customer.Identity"
    '    objScheme.Insert()
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Add(objScheme)
    '    'Me.DrawVolumeDiscounts()

    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Remove a discount scheme
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' Scheme removed by setting end_date to current date
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub cmdDelDiscount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelDiscount.Click
    '    If Not myVolDiscount Is Nothing Then
    '        myVolDiscount.Scheme.EndDate = Today.AddDays(-1)
    '        myVolDiscount.Scheme.Update()
    '        'Me.DrawVolumeDiscounts()
    '    End If
    '    If Not Me.myRoutineDiscount Is Nothing Then
    '        myRoutineDiscount.DiscountScheme.EndDate = Today.AddDays(-1)
    '        myRoutineDiscount.DiscountScheme.Update()
    '        'Me.DrawDiscounts()
    '    End If
    'End Sub

    'Private Sub lvVolumeDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvVolumeDiscounts.Click
    '    Me.myVolDiscount = Nothing
    '    Me.myRoutineDiscount = Nothing
    '    If Me.lvVolumeDiscounts.SelectedItems.Count = 1 Then
    '        Dim objLvItem As ListViewItem
    '        objLvItem = Me.lvVolumeDiscounts.SelectedItems(0)
    '        While TypeOf objLvItem Is ListViewItems.DiscountSchemes.lvVolumeDiscountEntry
    '            objLvItem = Me.lvVolumeDiscounts.Items(objLvItem.Index - 1)

    '        End While
    '        Me.myVolDiscount = objLvItem
    '    End If
    'End Sub

    'Private Sub lvRoutineDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRoutineDiscount.Click
    '    Me.myVolDiscount = Nothing
    '    Me.myRoutineDiscount = Nothing
    '    If Me.lvRoutineDiscount.SelectedItems.Count = 1 Then
    '        Dim objLvItem As ListViewItem
    '        objLvItem = Me.lvRoutineDiscount.SelectedItems(0)
    '        While TypeOf objLvItem Is ListViewItems.DiscountSchemes.lvDiscountScheme
    '            objLvItem = Me.lvRoutineDiscount.Items(objLvItem.Index - 1)

    '        End While
    '        Me.myRoutineDiscount = objLvItem
    '    End If


    'End Sub

    '''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with editing volume discount schemes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    'Private Sub lvVolumeDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvVolumeDiscounts.DoubleClick
    '    If myVolDiscount Is Nothing Then Exit Sub
    '    Dim objFrm As New MedscreenCommonGui.frmCustomerOption()
    '    With objFrm
    '        .Type = frmCustomerOption.FormType.CustomerDiscount
    '        .StartDate = myVolDiscount.Scheme.Startdate
    '        .EndDate = myVolDiscount.Scheme.EndDate
    '        If .ShowDialog = DialogResult.OK Then
    '            myVolDiscount.Scheme.Startdate = .StartDate
    '            myVolDiscount.Scheme.EndDate = .EndDate
    '            myVolDiscount.Scheme.ContractAmendment = .ContractAmmedment
    '            myVolDiscount.Scheme.Update()
    '        End If
    '    End With
    '    Me.DrawVolumeDiscounts()
    'End Sub


    'Private Sub lvRoutineDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRoutineDiscount.DoubleClick
    '    If Me.myRoutineDiscount Is Nothing Then Exit Sub
    '    Dim objFrm As New MedscreenCommonGui.frmCustomerOption()
    '    With objFrm
    '        .Type = frmCustomerOption.FormType.CustomerDiscount
    '        .StartDate = myRoutineDiscount.DiscountScheme.Startdate
    '        .EndDate = myRoutineDiscount.DiscountScheme.EndDate
    '        If .ShowDialog = DialogResult.OK Then
    '            myRoutineDiscount.DiscountScheme.Startdate = .StartDate
    '            myRoutineDiscount.DiscountScheme.EndDate = .EndDate
    '            myRoutineDiscount.DiscountScheme.ContractAmendment = .ContractAmmedment
    '            myRoutineDiscount.DiscountScheme.Update()
    '        End If
    '    End With
    '    'Me.DrawDiscounts()
    'End Sub

    Private Sub CkRemoved_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CkRemoved.CheckedChanged
        If Me.CkRemoved.Checked Then
            Me.cbStatus.SelectedIndex = 1
        Else
            Me.cbStatus.SelectedIndex = 0
        End If
    End Sub

    Private Sub drawContactList()

        Me.lvContacts.Items.Clear()
        Me.AdrContact.Address = Nothing                         'cLEAR OUT ADDRESS

        Dim objContact As MedscreenLib.contacts.CustomerContact
        For Each objContact In Me.Client.CustomerContacts
            Dim objLvCont As New MedscreenCommonGui.ListViewItems.LVCustomerContact(objContact)
            Me.lvContacts.Items.Add(objLvCont)
        Next
        If Me.Client.CustomerContacts.Count > 0 Then
            Me.lvContacts.Focus()
            Me.lvContacts.Items(0).Selected = True
            Me.lvContacts_Click(Nothing, Nothing)
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
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddContact.Click
        If Me.Client Is Nothing Then Exit Sub

        Me.Client.CustomerContacts.CreateCustomerContact()
        Me.drawContactList()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Variable to hold contact row
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private myLvContact As MedscreenCommonGui.ListViewItems.LVCustomerContact

    Private Sub CheckForChangedContact()
        Try
            If myLvContact IsNot Nothing AndAlso myLvContact.Contact IsNot Nothing AndAlso (myLvContact.Contact.Fields.Changed OrElse AddressChanged()) Then
                Dim objret As MsgBoxResult = MsgBox("The Contact's details have changed, do you want to save the changes (yes), loose the changes (no)?", MsgBoxStyle.YesNo)
                If objret = MsgBoxResult.Yes Then
                    myLvContact.Contact.Update()
                    UpdateAddress()
                Else
                    myLvContact.Contact.Fields.Rollback()
                End If

            End If
        Catch ex As Exception
            Medscreen.LogError(ex, False, "Check contact for change")
        End Try
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
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub lvContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvContacts.Click
        'Check to see if we have an existing contact
        If myLvContact IsNot Nothing Then ' if so we need to warn the user if changed
            UpdateContactInfo()           'Get contact info

            CheckForChangedContact()
        End If
        myLvContact = Nothing
        Me.cmdDelContact.Enabled = False
        If Me.lvContacts.SelectedItems.Count = 1 Then
            myLvContact = Me.lvContacts.SelectedItems.Item(0)

            If Not myLvContact Is Nothing Then
                Me.AdrContact.Address = myLvContact.Contact.Address
                Me.cbContactType.SelectedValue = myLvContact.Contact.ContactType
                Me.cbContMethod.SelectedValue = myLvContact.Contact.ContactMethod
                Me.cbReportOptions.SelectedValue = myLvContact.Contact.ReportOptions
                Me.cbContactStatus.SelectedValue = myLvContact.Contact.ContactStatus
                Me.cmdDelContact.Enabled = True
                Me.ckContactRemoved.Checked = myLvContact.Contact.Removed
            Else
                Me.AdrContact.Address = Nothing
                Me.cbContactType.SelectedValue = ""
                Me.cbContMethod.SelectedValue = ""
                Me.cbReportOptions.SelectedValue = ""

            End If
        End If

    End Sub

    Private Sub UpdateAddress()
        If Not Me.AdrContact.Address Is Nothing Then
            myLvContact.Contact.Address.Update()
        End If
    End Sub

    Private Function AddressChanged() As Boolean
        Dim BLNrET As Boolean = False
        If Not Me.AdrContact.Address Is Nothing Then
            BLNrET = Me.AdrContact.Address.Fields.Changed
        End If
        Return BLNrET
    End Function


    Private Sub UpdateContactInfo()
        If Not Me.AdrContact.Address Is Nothing Then
            myLvContact.Contact.AddressId = Me.AdrContact.Address.AddressId
            myLvContact.Contact.Address = Me.AdrContact.Address

        End If
        myLvContact.Contact.ContactMethod = Me.cbContMethod.SelectedValue
        myLvContact.Contact.Removed = Me.ckContactRemoved.Checked
        myLvContact.Contact.ContactStatus = Me.cbContactStatus.SelectedValue

        myLvContact.Contact.ContactType = Me.cbContactType.SelectedValue
        myLvContact.Contact.ReportOptions = Me.cbReportOptions.SelectedValue
        'check that we have a status
        If myLvContact.Contact.ContactStatus.Trim.Length = 0 Then myLvContact.Contact.ContactStatus = contacts.CustomerContact.GCST_STATUS_ACTIVE
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
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUpdateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateContact.Click
        'Update info 

        If myLvContact Is Nothing Then Exit Sub

        UpdateContactInfo()
        UpdateAddress()
        
        myLvContact.Contact.Update()
        myLvContact.Refresh()
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
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNewContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewContactAdr.Click
        If Me.myLvContact Is Nothing Then Exit Sub
        Dim addr As MedscreenLib.Address.Caddress
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "", Me.Client.Identity, MedscreenLib.Constants.AddressType.AddressMain)
        Me.AdrContact.Address = addr
        Me.myLvContact.Contact.Address = addr
        Me.myLvContact.Contact.AddressId = addr.AddressId


    End Sub

    Private Sub txtVATRegno_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVATRegno.Validating
        Me.ErrorProvider1.SetError(Me.txtVATRegno, "")

        Dim myregex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(System.Configuration.ConfigurationSettings.AppSettings("VATRegMask"))
        Dim mymatch As Match = myregex.Match(Me.txtVATRegno.Text.ToUpper)
        If mymatch.Index <> 0 OrElse Not mymatch.Success Then
            Me.ErrorProvider1.SetError(Me.txtVATRegno, "VAT number must be entered")
            Me.txtVATRegno.Focus()
            If Not e Is Nothing Then e.Cancel = True
        Else
            If mymatch.Length = 2 Then
                If Me.txtVATRegno.Text.Length = 2 Then
                    Dim strCountry As String = CConnection.PackageStringList("Lib_address.GetCountryNameByCode", Me.txtVATRegno.Text.Substring(0, 2))
                    If strCountry.Trim.Length = 0 Then
                        Me.ErrorProvider1.SetError(Me.txtVATRegno, "Country unknown ")
                        If Not e Is Nothing Then e.Cancel = True
                    Else
                        Me.lblCountry.Text = strCountry
                    End If
                Else
                    Me.lblCountry.Text = "Not EU"
                End If
            End If
        End If

    End Sub

    Private Sub cbCurrency_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCurrency.SelectedValueChanged
        If Me.cbCurrency.SelectedValue Is Nothing Then Exit Sub
        Dim strCulture As String
        If TypeOf Me.cbCurrency.SelectedValue Is String Then
            strCulture = MedscreenLib.Glossary.Glossary.Currencies.Culture(Me.cbCurrency.SelectedValue)
        Else
            strCulture = MedscreenLib.Glossary.Glossary.Currencies.Culture(Me.cbCurrency.SelectedValue.currencyid)
        End If

        Dim strPrice As String = CDbl("10.0").ToString("C", New Globalization.CultureInfo(strCulture.ToUpper)).Replace("0", " ").Replace("1", " ")
        strPrice = strPrice.Replace(",", " ").Replace(".", " ").Trim
        Me.lblCurrency.Text = "C&urrency (" & strPrice & ")"
    End Sub



    Private Sub tabTypes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabTypes.SelectedIndexChanged
        If Me.tabTypes.SelectedIndex = Me.tabTypes.TabPages.IndexOf(Me.tabpKits) Then
            Me.FillKitListview()
        ElseIf Me.tabTypes.SelectedIndex = Me.tabTypes.TabPages.IndexOf(tabpTests) Then
            Me.FillTestListView()
        End If
    End Sub


    Private Sub DTPricing_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangeDate.Click
        DoPanelFill()

    End Sub

    Private Sub myService_NotifyServiceChange(ByVal Service As String, ByVal Change As String) Handles CustOptions.NotifyServiceChange
        CommonForms.NotifyServiceChange(Client, Service, Change)
    End Sub

    Private Sub SetReportOptions()
        Me.cbReportOptions.Text = ""
        Me.cbReportOptions.SelectedIndex = -1
        objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("select phrase_id,phrase_text from phrase where phrase_type = '" & cbContactType.SelectedValue & "'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        cbReportOptions.DataSource = objRepOptionPhrase

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' changing the contact type will change the report options
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/12/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbContactType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbContactType.SelectedIndexChanged
        If (blnDrawing) Then Exit Sub
        SetReportOptions()
    End Sub

    Private Sub cmdFindContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindContactAdr.Click
        If Me.myLvContact Is Nothing Then Exit Sub
        Dim addr As MedscreenLib.Address.Caddress
        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                addr = New MedscreenLib.Address.Caddress(intAddrID)
                Me.myLvContact.Contact.AddressId = intAddrID
                Me.myLvContact.Contact.Address = addr
                Me.myLvContact.Contact.Update()
                Me.AdrContact.Address = addr
            End If
        End With

    End Sub

    Private Sub DoPanelFill()
        If Me.Client.InvoiceInfo.PricingMethod = Intranet.intranet.customerns.ClientInvoice.InvoicePricingType.StandardPricing Then
            Me.FillStdPanels()
        Else
            Me.FillPanels()
        End If

    End Sub

    'Private Sub txtParent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParent.TextChanged

    'End Sub
    Private blnMROIssue As Boolean = False
    Private myMRO As MRO
    Private Sub cbMROID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMROID.SelectedIndexChanged
        Dim strWithMRO As String = MedscreenCommonGUIConfig.NodePLSQL("WithMRO")
        Dim strMROSamplesCount As String = "0"
        If Not Client Is Nothing Then strMROSamplesCount = CConnection.PackageStringList(strWithMRO, Me.Client.Identity)
        ErrorProvider1.SetError(cbMROID, "")
        blnMROIssue = False
        If TypeOf cbMROID.SelectedValue Is MRO Then
            myMRO = cbMROID.SelectedValue

            If myMRO.MROID = STR_NOMRO AndAlso CInt(strMROSamplesCount) > 0 Then    'We have samples with MRO 
                ErrorProvider1.SetError(cbMROID, "There are samples with MRO for this customer")
                blnMROIssue = True
            End If
        ElseIf TypeOf cbMROID.SelectedValue Is String Then
            If cbMROID.SelectedValue = STR_NOMRO AndAlso CInt(strMROSamplesCount) > 0 Then    'We have samples with MRO 
                ErrorProvider1.SetError(cbMROID, "There are samples with MRO for this customer")
                blnMROIssue = True
            End If
        End If

    End Sub
End Class
'Friend WithEvents LvPricing As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvAccounts As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvCertificates As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvCustomerType As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvContractor As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvPrisons As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvSamples As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvReporting As MedscreenCommonGui.OptionsListView
'Friend WithEvents lvRepOptions As MedscreenCommonGui.OptionsListView
'Private Sub lvAccounts_OptionAdded(ByVal Sender As MedscreenCommonGui.OptionsListView, ByVal option_id As String) Handles lvAccounts.OptionAdded, lvAccounts.OptionRemoved, _
'    lvCertificates.OptionAdded, lvCertificates.OptionRemoved, lvContractor.OptionAdded, lvContractor.OptionRemoved, _
'    lvCustomerType.OptionAdded, lvCustomerType.OptionRemoved, LvPricing.OptionAdded, LvPricing.OptionRemoved, _
'    lvPrisons.OptionAdded, lvPrisons.OptionRemoved, lvRepOptions.OptionAdded, lvRepOptions.OptionRemoved, _
'    lvReporting.OptionAdded, lvReporting.OptionRemoved, lvSamples.OptionAdded, lvSamples.OptionRemoved

'    Me.Client.Options = Sender.ClientOptions
'    Me.DrawOptions()
'End Sub
