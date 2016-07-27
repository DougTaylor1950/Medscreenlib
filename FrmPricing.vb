Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports System.Text.RegularExpressions
Imports MedscreenCommonGui
Imports Intranet.intranet
Imports System.Windows.Forms
Public Class FrmPricing
    Inherits System.Windows.Forms.Form

    Private myClient As Intranet.intranet.customerns.Client
    Private myPricebooks As MedscreenLib.Glossary.PhraseCollection
    Private myVolDiscount As ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader
    Private myVolDiscountEntry As ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry
    Friend WithEvents VolDiscountPanel1 As MedscreenCommonGui.Panels.VolDiscountPanel
    Friend WithEvents lvTestConfirm As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader63 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader64 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader65 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader66 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader67 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader68 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader69 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdDeleteDiscounts As System.Windows.Forms.Button
    Private myRoutineDiscount As ListViewItems.DiscountSchemes.lvCustDiscountScheme

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.blnDisableAll = True
        Dim objUser As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.CurrentSMUser
        If Not objUser Is Nothing Then
            If MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserHasRole(objUser.Identity, "MED_EDIT_PRICES") Then
                Me.blnDisableAll = False
            End If
        End If

        Me.cbCurrency.DataSource = MedscreenLib.Glossary.Glossary.Currencies
        Me.cbCurrency.DisplayMember = "Description"
        Dim intCurrLength As Integer = System.Configuration.ConfigurationSettings.AppSettings("CurrencyCodeLength")
        If intCurrLength = 3 Then
            Me.cbCurrency.ValueMember = "CurrencyId"
        Else
            Me.cbCurrency.ValueMember = "SingleCharId"
        End If

        'Set platinum account type combobox up
        Me.cbPlatAccType.DataSource = MedscreenLib.Glossary.Glossary.Platinum
        Me.cbPlatAccType.DisplayMember = "PhraseText"
        Me.cbPlatAccType.ValueMember = "PhraseID"

        'Set invoice frequency combobox up
        Me.cbInvFreq.DataSource = MedscreenLib.Glossary.Glossary.Invoicing
        Me.cbInvFreq.DisplayMember = "PhraseText"
        Me.cbInvFreq.ValueMember = "PhraseID"

        'Set up price books
        myPricebooks = New MedscreenLib.Glossary.PhraseCollection("PRICEBOOK")
        myPricebooks.Load()
        Me.cbPricing.DataSource = myPricebooks
        Me.cbPricing.DisplayMember = "PhraseText"
        Me.cbPricing.ValueMember = "PhraseID"
        ''<Added code modified 23-Apr-2007 06:33 by taylor> 
        'Me.cbVolScheme.DataSource = MedscreenLib.Glossary.Glossary.VolumeDiscountSchemes
        'Me.cbVolScheme.DisplayMember = "Description"
        'Me.cbVolScheme.ValueMember = "DiscountSchemeId"


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
            Me.Panel2.Controls.Add(Me.Panel5)
        Catch ex As Exception
        End Try
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
    Friend WithEvents tabTypes As System.Windows.Forms.TabControl
    Friend WithEvents tabpSample As System.Windows.Forms.TabPage
    Friend WithEvents pnlExternal As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents tabpMinCharge As System.Windows.Forms.TabPage
    Friend WithEvents pnlMinCharge As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents lblpCollCurr As System.Windows.Forms.Label
    Friend WithEvents lbllpCollCurr As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents ctxAddService As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuctxAddService As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCTXDiscountDates As System.Windows.Forms.MenuItem
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents tabpAddTime As System.Windows.Forms.TabPage
    Friend WithEvents pnlAddTime As System.Windows.Forms.Panel
    Friend WithEvents pnlADHeader As System.Windows.Forms.Panel
    Friend WithEvents lbladpCurr As System.Windows.Forms.Label
    Friend WithEvents lbllPADCurr As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents pnlINHoursAddTimeRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents txtinhratrtScheme As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratrtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratrtDiscount As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratrtListPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents pnlOutHoursAddTimeRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents txtouthratrtScheme As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratrtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratrtDiscount As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratrtListPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents pnlINHoursAddTimeCO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents txtinhratcoScheme As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratcoPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratcoDiscount As System.Windows.Forms.TextBox
    Friend WithEvents txtinhratcoListPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents pnlOutHoursAddTimeCO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents txtouthratcoScheme As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratcoPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratcoDiscount As System.Windows.Forms.TextBox
    Friend WithEvents txtouthratcoListPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents pnlPriceInfo As System.Windows.Forms.Panel
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents lblihadtpricephr As System.Windows.Forms.Label
    Friend WithEvents lblohadtpricephr As System.Windows.Forms.Label
    Friend WithEvents lblihadtrpricephr As System.Windows.Forms.Label
    Friend WithEvents lblohadtrpricephr As System.Windows.Forms.Label
    Friend WithEvents tabpTests As System.Windows.Forms.TabPage
    Friend WithEvents pnlTests As System.Windows.Forms.Panel
    Friend WithEvents lvTests As ListViewEx.ListViewEx
    Friend WithEvents chTest As System.Windows.Forms.ColumnHeader
    Friend WithEvents chBar1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chListPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents chbar2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDiscount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chBar3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbEditor As System.Windows.Forms.ComboBox
    Friend WithEvents ckChargeAdditionalTests As System.Windows.Forms.CheckBox
    Friend WithEvents pnlRoutine As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents pnlBenzene As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlCallOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlFixedSite As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Label73 As System.Windows.Forms.Label
    Friend WithEvents pnlMRO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents lblMROComment As System.Windows.Forms.Label
    Friend WithEvents dpCallOutCollChargeOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpCallOutCollCharge As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpRoutineCollChargeOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpRoutineCollCharge As MedscreenCommonGui.Panels.DiscountPanel
    'Friend WithEvents pnlCallOutCharge As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents txtCallOutCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents dpMCCallOutOut As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpMCCallOutIn As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlOutHoursMC As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents pnlInHourMinCharge As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader17 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader18 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader19 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader20 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader21 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader22 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAvgPrice2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chListPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chPriceKit As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCount As System.Windows.Forms.ColumnHeader
    Friend WithEvents chKitSold As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAvgPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader29 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader30 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader31 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader32 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader33 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader34 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader35 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader36 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader37 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader38 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader39 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader40 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader41 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader42 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader43 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader44 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader45 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader46 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader47 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader48 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader49 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader50 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader51 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader52 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader53 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader54 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader55 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader56 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader57 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader58 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader59 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader60 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader61 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader62 As System.Windows.Forms.ColumnHeader
    Friend WithEvents tabpCancellations As System.Windows.Forms.TabPage
    Friend WithEvents pnlCancellations As System.Windows.Forms.Panel
    Friend WithEvents dpCallOutCancel As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpFixedSiteCancellation As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpInHoursCancel As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents dpOutOfhoursRoutine As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents pnlSelectCharging As System.Windows.Forms.Panel
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents udCallOut As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents udRoutine As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents pnlMinchargeheader As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents pnlTimeAllowHead As System.Windows.Forms.Panel
    Friend WithEvents Label70 As System.Windows.Forms.Label
    Friend WithEvents txtTimeAllowance As System.Windows.Forms.TextBox
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents pnlTimeAllowance As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblTimeAllowance As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ckTimeAllowance As System.Windows.Forms.CheckBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents pnlAO As MedscreenCommonGui.Panels.DiscountPanel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents lblLPCurr As System.Windows.Forms.Label
    Friend WithEvents lblPCurr As System.Windows.Forms.Label
    Friend WithEvents tabInvoicing As System.Windows.Forms.TabPage
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents cbInvFreq As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents lbServices As System.Windows.Forms.ListBox
    Friend WithEvents cbPricing As System.Windows.Forms.ComboBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents lblCurrency As System.Windows.Forms.Label
    Friend WithEvents ckCreditCard As System.Windows.Forms.CheckBox
    Friend WithEvents lblSageID As System.Windows.Forms.Label
    Friend WithEvents txtSageID As System.Windows.Forms.TextBox
    Friend WithEvents cbPlatAccType As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents ckPORequired As System.Windows.Forms.CheckBox
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents lblCountry As System.Windows.Forms.Label
    Friend WithEvents lblVATRegno As System.Windows.Forms.Label
    Friend WithEvents txtVATRegno As System.Windows.Forms.TextBox
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label72 As System.Windows.Forms.Label
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label74 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuSavePanels As System.Windows.Forms.MenuItem
    Friend WithEvents cmdClearDiscounts As System.Windows.Forms.Button
    Friend WithEvents TabDiscounts As System.Windows.Forms.TabPage
    Friend WithEvents pnlVolDiscountScheme As System.Windows.Forms.Panel
    Friend WithEvents DTPricing As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmdChangeDate As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPricing))
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdDeleteDiscounts = New System.Windows.Forms.Button
        Me.cmdChangeDate = New System.Windows.Forms.Button
        Me.DTPricing = New System.Windows.Forms.DateTimePicker
        Me.cmdClearDiscounts = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.tabTypes = New System.Windows.Forms.TabControl
        Me.tabpSample = New System.Windows.Forms.TabPage
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.tabpCancellations = New System.Windows.Forms.TabPage
        Me.pnlCancellations = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.Label52 = New System.Windows.Forms.Label
        Me.Label53 = New System.Windows.Forms.Label
        Me.Label59 = New System.Windows.Forms.Label
        Me.Label60 = New System.Windows.Forms.Label
        Me.tabpMinCharge = New System.Windows.Forms.TabPage
        Me.pnlMinCharge = New System.Windows.Forms.Panel
        Me.tabpAddTime = New System.Windows.Forms.TabPage
        Me.pnlAddTime = New System.Windows.Forms.Panel
        Me.pnlADHeader = New System.Windows.Forms.Panel
        Me.Label34 = New System.Windows.Forms.Label
        Me.lbladpCurr = New System.Windows.Forms.Label
        Me.lbllPADCurr = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.pnlTimeAllowance = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label39 = New System.Windows.Forms.Label
        Me.lblTimeAllowance = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.ckTimeAllowance = New System.Windows.Forms.CheckBox
        Me.tabpTests = New System.Windows.Forms.TabPage
        Me.pnlTests = New System.Windows.Forms.Panel
        Me.lvTestConfirm = New ListViewEx.ListViewEx
        Me.ColumnHeader63 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader64 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader65 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader66 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader67 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader68 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader69 = New System.Windows.Forms.ColumnHeader
        Me.Label74 = New System.Windows.Forms.Label
        Me.ckChargeAdditionalTests = New System.Windows.Forms.CheckBox
        Me.cbEditor = New System.Windows.Forms.ComboBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.lvTests = New ListViewEx.ListViewEx
        Me.chTest = New System.Windows.Forms.ColumnHeader
        Me.chBar1 = New System.Windows.Forms.ColumnHeader
        Me.chListPrice = New System.Windows.Forms.ColumnHeader
        Me.chbar2 = New System.Windows.Forms.ColumnHeader
        Me.chDiscount = New System.Windows.Forms.ColumnHeader
        Me.chBar3 = New System.Windows.Forms.ColumnHeader
        Me.chPrice = New System.Windows.Forms.ColumnHeader
        Me.tabInvoicing = New System.Windows.Forms.TabPage
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.lblCountry = New System.Windows.Forms.Label
        Me.lblVATRegno = New System.Windows.Forms.Label
        Me.txtVATRegno = New System.Windows.Forms.TextBox
        Me.cbPlatAccType = New System.Windows.Forms.ComboBox
        Me.cbCurrency = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ckCreditCard = New System.Windows.Forms.CheckBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label38 = New System.Windows.Forms.Label
        Me.cbInvFreq = New System.Windows.Forms.ComboBox
        Me.lbServices = New System.Windows.Forms.ListBox
        Me.cbPricing = New System.Windows.Forms.ComboBox
        Me.txtSageID = New System.Windows.Forms.TextBox
        Me.lblCurrency = New System.Windows.Forms.Label
        Me.lblSageID = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.ckPORequired = New System.Windows.Forms.CheckBox
        Me.TabDiscounts = New System.Windows.Forms.TabPage
        Me.pnlVolDiscountScheme = New System.Windows.Forms.Panel
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.lblPCurr = New System.Windows.Forms.Label
        Me.lblLPCurr = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.txtCallOutCharge = New System.Windows.Forms.TextBox
        Me.Label50 = New System.Windows.Forms.Label
        Me.pnlMinchargeheader = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
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
        Me.ctxAddService = New System.Windows.Forms.ContextMenu
        Me.mnuctxAddService = New System.Windows.Forms.MenuItem
        Me.mnuCTXDiscountDates = New System.Windows.Forms.MenuItem
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblohadtrpricephr = New System.Windows.Forms.Label
        Me.txtouthratcoScheme = New System.Windows.Forms.TextBox
        Me.txtouthratcoPrice = New System.Windows.Forms.TextBox
        Me.txtouthratcoDiscount = New System.Windows.Forms.TextBox
        Me.txtouthratcoListPrice = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.lblihadtrpricephr = New System.Windows.Forms.Label
        Me.txtinhratcoScheme = New System.Windows.Forms.TextBox
        Me.txtinhratcoPrice = New System.Windows.Forms.TextBox
        Me.txtinhratcoDiscount = New System.Windows.Forms.TextBox
        Me.txtinhratcoListPrice = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.lblohadtpricephr = New System.Windows.Forms.Label
        Me.txtouthratrtScheme = New System.Windows.Forms.TextBox
        Me.txtouthratrtPrice = New System.Windows.Forms.TextBox
        Me.txtouthratrtDiscount = New System.Windows.Forms.TextBox
        Me.txtouthratrtListPrice = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.lblihadtpricephr = New System.Windows.Forms.Label
        Me.txtinhratrtScheme = New System.Windows.Forms.TextBox
        Me.txtinhratrtPrice = New System.Windows.Forms.TextBox
        Me.txtinhratrtDiscount = New System.Windows.Forms.TextBox
        Me.txtinhratrtListPrice = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.lblpCollCurr = New System.Windows.Forms.Label
        Me.lbllpCollCurr = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.pnlPriceInfo = New System.Windows.Forms.Panel
        Me.Label42 = New System.Windows.Forms.Label
        Me.Label41 = New System.Windows.Forms.Label
        Me.Label61 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label72 = New System.Windows.Forms.Label
        Me.Label71 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader17 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader18 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader19 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader20 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader21 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader22 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.chAvgPrice2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader
        Me.chListPriceKit = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader
        Me.chPriceKit = New System.Windows.Forms.ColumnHeader
        Me.chCount = New System.Windows.Forms.ColumnHeader
        Me.chKitSold = New System.Windows.Forms.ColumnHeader
        Me.chAvgPrice = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader29 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader30 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader31 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader32 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader33 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader34 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader35 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader36 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader37 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader38 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader39 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader40 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader41 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader42 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader43 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader44 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader45 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader46 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader47 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader48 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader49 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader50 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader51 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader52 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader53 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader54 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader55 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader56 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader57 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader58 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader59 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader60 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader61 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader62 = New System.Windows.Forms.ColumnHeader
        Me.pnlTimeAllowHead = New System.Windows.Forms.Panel
        Me.Label70 = New System.Windows.Forms.Label
        Me.txtTimeAllowance = New System.Windows.Forms.TextBox
        Me.Label69 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuSavePanels = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlMRO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.lblMROComment = New System.Windows.Forms.Label
        Me.pnlExternal = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.Label6 = New System.Windows.Forms.Label
        Me.pnlFixedSite = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.Label73 = New System.Windows.Forms.Label
        Me.pnlCallOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlAO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpCallOutCancel = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpFixedSiteCancellation = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpInHoursCancel = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpOutOfhoursRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpCallOutCollChargeOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpCallOutCollCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpRoutineCollChargeOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpRoutineCollCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpMCCallOutOut = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.dpMCCallOutIn = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlOutHoursMC = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlInHourMinCharge = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlOutHoursAddTimeCO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlINHoursAddTimeCO = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlOutHoursAddTimeRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.pnlINHoursAddTimeRoutine = New MedscreenCommonGui.Panels.DiscountPanel(Me.components)
        Me.VolDiscountPanel1 = New MedscreenCommonGui.Panels.VolDiscountPanel
        Me.Panel1.SuspendLayout()
        Me.tabTypes.SuspendLayout()
        Me.tabpSample.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.tabpCancellations.SuspendLayout()
        Me.pnlCancellations.SuspendLayout()
        Me.tabpMinCharge.SuspendLayout()
        Me.pnlMinCharge.SuspendLayout()
        Me.tabpAddTime.SuspendLayout()
        Me.pnlAddTime.SuspendLayout()
        Me.pnlADHeader.SuspendLayout()
        Me.pnlTimeAllowance.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.tabpTests.SuspendLayout()
        Me.pnlTests.SuspendLayout()
        Me.tabInvoicing.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.TabDiscounts.SuspendLayout()
        Me.pnlVolDiscountScheme.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.pnlMinchargeheader.SuspendLayout()
        Me.pnlSelectCharging.SuspendLayout()
        Me.pnlPriceInfo.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTimeAllowHead.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMRO.SuspendLayout()
        Me.pnlExternal.SuspendLayout()
        Me.pnlFixedSite.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdDeleteDiscounts)
        Me.Panel1.Controls.Add(Me.cmdChangeDate)
        Me.Panel1.Controls.Add(Me.DTPricing)
        Me.Panel1.Controls.Add(Me.cmdClearDiscounts)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 553)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(771, 61)
        Me.Panel1.TabIndex = 10
        '
        'cmdDeleteDiscounts
        '
        Me.cmdDeleteDiscounts.Location = New System.Drawing.Point(118, 17)
        Me.cmdDeleteDiscounts.Name = "cmdDeleteDiscounts"
        Me.cmdDeleteDiscounts.Size = New System.Drawing.Size(96, 23)
        Me.cmdDeleteDiscounts.TabIndex = 5
        Me.cmdDeleteDiscounts.Text = "Delete Discounts"
        Me.ToolTip1.SetToolTip(Me.cmdDeleteDiscounts, "This will permanetly delete all simple discounts resetting the account back the p" & _
                "rices in the price book." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "It can't be undone and is immediate.  You will have to" & _
                " reconstruct the pricing from scratch.")
        '
        'cmdChangeDate
        '
        Me.cmdChangeDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdChangeDate.Image = CType(resources.GetObject("cmdChangeDate.Image"), System.Drawing.Image)
        Me.cmdChangeDate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdChangeDate.Location = New System.Drawing.Point(432, 21)
        Me.cmdChangeDate.Name = "cmdChangeDate"
        Me.cmdChangeDate.Size = New System.Drawing.Size(72, 24)
        Me.cmdChangeDate.TabIndex = 4
        Me.cmdChangeDate.Text = "Set Date"
        Me.cmdChangeDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTPricing
        '
        Me.DTPricing.Location = New System.Drawing.Point(280, 17)
        Me.DTPricing.Name = "DTPricing"
        Me.DTPricing.Size = New System.Drawing.Size(144, 20)
        Me.DTPricing.TabIndex = 3
        '
        'cmdClearDiscounts
        '
        Me.cmdClearDiscounts.Location = New System.Drawing.Point(16, 17)
        Me.cmdClearDiscounts.Name = "cmdClearDiscounts"
        Me.cmdClearDiscounts.Size = New System.Drawing.Size(96, 23)
        Me.cmdClearDiscounts.TabIndex = 2
        Me.cmdClearDiscounts.Text = "Clear Discounts"
        Me.ToolTip1.SetToolTip(Me.cmdClearDiscounts, "This will reset the discounts back to a particular date")
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(683, 21)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(603, 21)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tabTypes
        '
        Me.tabTypes.Controls.Add(Me.tabpSample)
        Me.tabTypes.Controls.Add(Me.tabpCancellations)
        Me.tabTypes.Controls.Add(Me.tabpMinCharge)
        Me.tabTypes.Controls.Add(Me.tabpAddTime)
        Me.tabTypes.Controls.Add(Me.tabpTests)
        Me.tabTypes.Controls.Add(Me.tabInvoicing)
        Me.tabTypes.Controls.Add(Me.TabDiscounts)
        Me.tabTypes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabTypes.Location = New System.Drawing.Point(0, 0)
        Me.tabTypes.Name = "tabTypes"
        Me.tabTypes.SelectedIndex = 0
        Me.tabTypes.Size = New System.Drawing.Size(771, 441)
        Me.tabTypes.TabIndex = 11
        '
        'tabpSample
        '
        Me.tabpSample.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabpSample.Controls.Add(Me.Panel2)
        Me.tabpSample.Location = New System.Drawing.Point(4, 22)
        Me.tabpSample.Name = "tabpSample"
        Me.tabpSample.Size = New System.Drawing.Size(763, 415)
        Me.tabpSample.TabIndex = 0
        Me.tabpSample.Text = "Sample"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.pnlMRO)
        Me.Panel2.Controls.Add(Me.pnlExternal)
        Me.Panel2.Controls.Add(Me.pnlFixedSite)
        Me.Panel2.Controls.Add(Me.pnlCallOut)
        Me.Panel2.Controls.Add(Me.pnlRoutine)
        Me.Panel2.Controls.Add(Me.pnlAO)
        Me.Panel2.Location = New System.Drawing.Point(16, 16)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(739, 488)
        Me.Panel2.TabIndex = 0
        '
        'tabpCancellations
        '
        Me.tabpCancellations.Controls.Add(Me.pnlCancellations)
        Me.tabpCancellations.Location = New System.Drawing.Point(4, 22)
        Me.tabpCancellations.Name = "tabpCancellations"
        Me.tabpCancellations.Size = New System.Drawing.Size(763, 435)
        Me.tabpCancellations.TabIndex = 6
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
        Me.pnlCancellations.Controls.Add(Me.Label2)
        Me.pnlCancellations.Controls.Add(Me.Label45)
        Me.pnlCancellations.Controls.Add(Me.Label52)
        Me.pnlCancellations.Controls.Add(Me.Label53)
        Me.pnlCancellations.Controls.Add(Me.Label59)
        Me.pnlCancellations.Controls.Add(Me.Label60)
        Me.pnlCancellations.Location = New System.Drawing.Point(16, 24)
        Me.pnlCancellations.Name = "pnlCancellations"
        Me.pnlCancellations.Size = New System.Drawing.Size(739, 216)
        Me.pnlCancellations.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(488, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(19, 13)
        Me.Label2.TabIndex = 62
        Me.Label2.Text = "(£)"
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
        'tabpMinCharge
        '
        Me.tabpMinCharge.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabpMinCharge.Controls.Add(Me.pnlMinCharge)
        Me.tabpMinCharge.Location = New System.Drawing.Point(4, 22)
        Me.tabpMinCharge.Name = "tabpMinCharge"
        Me.tabpMinCharge.Size = New System.Drawing.Size(763, 435)
        Me.tabpMinCharge.TabIndex = 1
        Me.tabpMinCharge.Text = "Minimum Charge"
        '
        'pnlMinCharge
        '
        Me.pnlMinCharge.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
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
        Me.pnlMinCharge.Size = New System.Drawing.Size(755, 436)
        Me.pnlMinCharge.TabIndex = 0
        '
        'tabpAddTime
        '
        Me.tabpAddTime.Controls.Add(Me.pnlAddTime)
        Me.tabpAddTime.Location = New System.Drawing.Point(4, 22)
        Me.tabpAddTime.Name = "tabpAddTime"
        Me.tabpAddTime.Size = New System.Drawing.Size(763, 435)
        Me.tabpAddTime.TabIndex = 2
        Me.tabpAddTime.Text = "Additional Time"
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
        Me.pnlAddTime.Size = New System.Drawing.Size(739, 320)
        Me.pnlAddTime.TabIndex = 0
        '
        'pnlADHeader
        '
        Me.pnlADHeader.Controls.Add(Me.Label34)
        Me.pnlADHeader.Controls.Add(Me.lbladpCurr)
        Me.pnlADHeader.Controls.Add(Me.lbllPADCurr)
        Me.pnlADHeader.Controls.Add(Me.Label28)
        Me.pnlADHeader.Controls.Add(Me.Label29)
        Me.pnlADHeader.Controls.Add(Me.Label30)
        Me.pnlADHeader.Controls.Add(Me.Label31)
        Me.pnlADHeader.Controls.Add(Me.Label32)
        Me.pnlADHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlADHeader.Location = New System.Drawing.Point(0, 80)
        Me.pnlADHeader.Name = "pnlADHeader"
        Me.pnlADHeader.Size = New System.Drawing.Size(737, 48)
        Me.pnlADHeader.TabIndex = 1
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(520, 12)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(45, 13)
        Me.Label34.TabIndex = 63
        Me.Label34.Text = "Price/hr"
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
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(582, 12)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(46, 13)
        Me.Label28.TabIndex = 60
        Me.Label28.Text = "Scheme"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(454, 12)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(31, 13)
        Me.Label29.TabIndex = 59
        Me.Label29.Text = "Price"
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(360, 12)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(100, 23)
        Me.Label30.TabIndex = 58
        Me.Label30.Text = "Discount %"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(216, 12)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(87, 13)
        Me.Label31.TabIndex = 57
        Me.Label31.Text = "List Price per min"
        '
        'Label32
        '
        Me.Label32.Location = New System.Drawing.Point(80, 12)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(100, 23)
        Me.Label32.TabIndex = 56
        Me.Label32.Text = "Min Charge Type "
        '
        'pnlTimeAllowance
        '
        Me.pnlTimeAllowance.Controls.Add(Me.Panel4)
        Me.pnlTimeAllowance.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTimeAllowance.Location = New System.Drawing.Point(0, 0)
        Me.pnlTimeAllowance.Name = "pnlTimeAllowance"
        Me.pnlTimeAllowance.Size = New System.Drawing.Size(737, 80)
        Me.pnlTimeAllowance.TabIndex = 61
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label39)
        Me.Panel4.Controls.Add(Me.lblTimeAllowance)
        Me.Panel4.Controls.Add(Me.TextBox2)
        Me.Panel4.Controls.Add(Me.Label12)
        Me.Panel4.Controls.Add(Me.ckTimeAllowance)
        Me.Panel4.Location = New System.Drawing.Point(0, 8)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(696, 136)
        Me.Panel4.TabIndex = 62
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.SystemColors.Info
        Me.Label39.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label39.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label39.Location = New System.Drawing.Point(172, 57)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(352, 23)
        Me.Label39.TabIndex = 62
        Me.Label39.Text = "Time allowance can be removed by removing all text from box "
        '
        'lblTimeAllowance
        '
        Me.lblTimeAllowance.Location = New System.Drawing.Point(344, 16)
        Me.lblTimeAllowance.Name = "lblTimeAllowance"
        Me.lblTimeAllowance.Size = New System.Drawing.Size(344, 23)
        Me.lblTimeAllowance.TabIndex = 60
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(280, 16)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(40, 20)
        Me.TextBox2.TabIndex = 59
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(176, 16)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 13)
        Me.Label12.TabIndex = 58
        Me.Label12.Text = "Time Allowance (hr)"
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
        Me.tabpTests.Location = New System.Drawing.Point(4, 22)
        Me.tabpTests.Name = "tabpTests"
        Me.tabpTests.Size = New System.Drawing.Size(763, 435)
        Me.tabpTests.TabIndex = 3
        Me.tabpTests.Text = "Tests"
        '
        'pnlTests
        '
        Me.pnlTests.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTests.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTests.Controls.Add(Me.lvTestConfirm)
        Me.pnlTests.Controls.Add(Me.Label74)
        Me.pnlTests.Controls.Add(Me.ckChargeAdditionalTests)
        Me.pnlTests.Controls.Add(Me.cbEditor)
        Me.pnlTests.Controls.Add(Me.TextBox1)
        Me.pnlTests.Controls.Add(Me.lvTests)
        Me.pnlTests.Location = New System.Drawing.Point(16, 8)
        Me.pnlTests.Name = "pnlTests"
        Me.pnlTests.Size = New System.Drawing.Size(739, 424)
        Me.pnlTests.TabIndex = 0
        '
        'lvTestConfirm
        '
        Me.lvTestConfirm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTestConfirm.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader63, Me.ColumnHeader64, Me.ColumnHeader65, Me.ColumnHeader66, Me.ColumnHeader67, Me.ColumnHeader68, Me.ColumnHeader69})
        Me.lvTestConfirm.FullRowSelect = True
        Me.lvTestConfirm.GridLines = True
        Me.lvTestConfirm.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvTestConfirm.Location = New System.Drawing.Point(3, 238)
        Me.lvTestConfirm.Name = "lvTestConfirm"
        Me.lvTestConfirm.Size = New System.Drawing.Size(729, 171)
        Me.lvTestConfirm.TabIndex = 9
        Me.lvTestConfirm.UseCompatibleStateImageBehavior = False
        Me.lvTestConfirm.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader63
        '
        Me.ColumnHeader63.Text = "Description/Test"
        Me.ColumnHeader63.Width = 150
        '
        'ColumnHeader64
        '
        Me.ColumnHeader64.Text = ""
        Me.ColumnHeader64.Width = 40
        '
        'ColumnHeader65
        '
        Me.ColumnHeader65.Text = "List Price"
        Me.ColumnHeader65.Width = 100
        '
        'ColumnHeader66
        '
        Me.ColumnHeader66.Text = ""
        '
        'ColumnHeader67
        '
        Me.ColumnHeader67.Text = "Discount (%)"
        Me.ColumnHeader67.Width = 100
        '
        'ColumnHeader68
        '
        Me.ColumnHeader68.Text = ""
        '
        'ColumnHeader69
        '
        Me.ColumnHeader69.Text = "Price"
        '
        'Label74
        '
        Me.Label74.Location = New System.Drawing.Point(9, 206)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(224, 16)
        Me.Label74.TabIndex = 8
        Me.Label74.Text = "Charges for Confirmation Tests"
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
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(360, 0)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "TextBox1"
        Me.TextBox1.Visible = False
        '
        'lvTests
        '
        Me.lvTests.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTests.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chTest, Me.chBar1, Me.chListPrice, Me.chbar2, Me.chDiscount, Me.chBar3, Me.chPrice})
        Me.lvTests.FullRowSelect = True
        Me.lvTests.GridLines = True
        Me.lvTests.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvTests.Location = New System.Drawing.Point(16, 48)
        Me.lvTests.Name = "lvTests"
        Me.lvTests.Size = New System.Drawing.Size(716, 141)
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
        'tabInvoicing
        '
        Me.tabInvoicing.Controls.Add(Me.HtmlEditor1)
        Me.tabInvoicing.Controls.Add(Me.Splitter1)
        Me.tabInvoicing.Controls.Add(Me.Panel6)
        Me.tabInvoicing.Location = New System.Drawing.Point(4, 22)
        Me.tabInvoicing.Name = "tabInvoicing"
        Me.tabInvoicing.Size = New System.Drawing.Size(763, 435)
        Me.tabInvoicing.TabIndex = 7
        Me.tabInvoicing.Tag = "Invoicing"
        Me.tabInvoicing.Text = "Invoicing"
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 179)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(763, 256)
        Me.HtmlEditor1.TabIndex = 2
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 176)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(763, 3)
        Me.Splitter1.TabIndex = 1
        Me.Splitter1.TabStop = False
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.lblCountry)
        Me.Panel6.Controls.Add(Me.lblVATRegno)
        Me.Panel6.Controls.Add(Me.txtVATRegno)
        Me.Panel6.Controls.Add(Me.cbPlatAccType)
        Me.Panel6.Controls.Add(Me.cbCurrency)
        Me.Panel6.Controls.Add(Me.Label3)
        Me.Panel6.Controls.Add(Me.ckCreditCard)
        Me.Panel6.Controls.Add(Me.Label26)
        Me.Panel6.Controls.Add(Me.Label38)
        Me.Panel6.Controls.Add(Me.cbInvFreq)
        Me.Panel6.Controls.Add(Me.lbServices)
        Me.Panel6.Controls.Add(Me.cbPricing)
        Me.Panel6.Controls.Add(Me.txtSageID)
        Me.Panel6.Controls.Add(Me.lblCurrency)
        Me.Panel6.Controls.Add(Me.lblSageID)
        Me.Panel6.Controls.Add(Me.Label40)
        Me.Panel6.Controls.Add(Me.ckPORequired)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel6.Location = New System.Drawing.Point(0, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(763, 176)
        Me.Panel6.TabIndex = 0
        '
        'lblCountry
        '
        Me.lblCountry.AutoSize = True
        Me.lblCountry.Location = New System.Drawing.Point(344, 64)
        Me.lblCountry.Name = "lblCountry"
        Me.lblCountry.Size = New System.Drawing.Size(19, 13)
        Me.lblCountry.TabIndex = 70
        Me.lblCountry.Text = "----"
        '
        'lblVATRegno
        '
        Me.lblVATRegno.Location = New System.Drawing.Point(336, 80)
        Me.lblVATRegno.Name = "lblVATRegno"
        Me.lblVATRegno.Size = New System.Drawing.Size(96, 16)
        Me.lblVATRegno.TabIndex = 69
        Me.lblVATRegno.Text = "Company VAT No"
        Me.lblVATRegno.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtVATRegno
        '
        Me.txtVATRegno.Location = New System.Drawing.Point(432, 80)
        Me.txtVATRegno.Name = "txtVATRegno"
        Me.txtVATRegno.Size = New System.Drawing.Size(136, 20)
        Me.txtVATRegno.TabIndex = 68
        '
        'cbPlatAccType
        '
        Me.cbPlatAccType.Location = New System.Drawing.Point(149, 75)
        Me.cbPlatAccType.Name = "cbPlatAccType"
        Me.cbPlatAccType.Size = New System.Drawing.Size(176, 21)
        Me.cbPlatAccType.TabIndex = 33
        '
        'cbCurrency
        '
        Me.cbCurrency.Location = New System.Drawing.Point(421, 21)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(121, 21)
        Me.cbCurrency.TabIndex = 38
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "Invoice Frequency"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ckCreditCard
        '
        Me.ckCreditCard.Location = New System.Drawing.Point(37, 51)
        Me.ckCreditCard.Name = "ckCreditCard"
        Me.ckCreditCard.Size = New System.Drawing.Size(168, 24)
        Me.ckCreditCard.TabIndex = 36
        Me.ckCreditCard.Text = "&Card Payment Required"
        '
        'Label26
        '
        Me.Label26.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.Location = New System.Drawing.Point(624, 19)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(100, 23)
        Me.Label26.TabIndex = 44
        Me.Label26.Text = "Services"
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(14, 107)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(124, 23)
        Me.Label38.TabIndex = 41
        Me.Label38.Text = "P&rice Book"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbInvFreq
        '
        Me.cbInvFreq.Location = New System.Drawing.Point(149, 144)
        Me.cbInvFreq.Name = "cbInvFreq"
        Me.cbInvFreq.Size = New System.Drawing.Size(419, 21)
        Me.cbInvFreq.TabIndex = 46
        '
        'lbServices
        '
        Me.lbServices.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbServices.Location = New System.Drawing.Point(624, 51)
        Me.lbServices.Name = "lbServices"
        Me.lbServices.Size = New System.Drawing.Size(136, 82)
        Me.lbServices.TabIndex = 43
        '
        'cbPricing
        '
        Me.cbPricing.Items.AddRange(New Object() {"Standard Pricing ", "Discount Pricing "})
        Me.cbPricing.Location = New System.Drawing.Point(149, 112)
        Me.cbPricing.Name = "cbPricing"
        Me.cbPricing.Size = New System.Drawing.Size(419, 21)
        Me.cbPricing.TabIndex = 42
        '
        'txtSageID
        '
        Me.txtSageID.Location = New System.Drawing.Point(253, 21)
        Me.txtSageID.Name = "txtSageID"
        Me.txtSageID.Size = New System.Drawing.Size(64, 20)
        Me.txtSageID.TabIndex = 34
        '
        'lblCurrency
        '
        Me.lblCurrency.Location = New System.Drawing.Point(357, 21)
        Me.lblCurrency.Name = "lblCurrency"
        Me.lblCurrency.Size = New System.Drawing.Size(56, 50)
        Me.lblCurrency.TabIndex = 37
        Me.lblCurrency.Text = "C&urrency"
        Me.lblCurrency.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSageID
        '
        Me.lblSageID.Location = New System.Drawing.Point(205, 21)
        Me.lblSageID.Name = "lblSageID"
        Me.lblSageID.Size = New System.Drawing.Size(48, 23)
        Me.lblSageID.TabIndex = 35
        Me.lblSageID.Text = "&Sage ID"
        Me.lblSageID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(10, 75)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(128, 23)
        Me.Label40.TabIndex = 32
        Me.Label40.Text = "Platinum &Account Type"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckPORequired
        '
        Me.ckPORequired.Location = New System.Drawing.Point(37, 21)
        Me.ckPORequired.Name = "ckPORequired"
        Me.ckPORequired.Size = New System.Drawing.Size(168, 24)
        Me.ckPORequired.TabIndex = 31
        Me.ckPORequired.Text = "&Purchase Order Required"
        '
        'TabDiscounts
        '
        Me.TabDiscounts.Controls.Add(Me.pnlVolDiscountScheme)
        Me.TabDiscounts.Controls.Add(Me.Splitter2)
        Me.TabDiscounts.Location = New System.Drawing.Point(4, 22)
        Me.TabDiscounts.Name = "TabDiscounts"
        Me.TabDiscounts.Size = New System.Drawing.Size(763, 435)
        Me.TabDiscounts.TabIndex = 8
        Me.TabDiscounts.Tag = "Discounts"
        Me.TabDiscounts.Text = "Discounts"
        '
        'pnlVolDiscountScheme
        '
        Me.pnlVolDiscountScheme.Controls.Add(Me.VolDiscountPanel1)
        Me.pnlVolDiscountScheme.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlVolDiscountScheme.Location = New System.Drawing.Point(0, 3)
        Me.pnlVolDiscountScheme.Name = "pnlVolDiscountScheme"
        Me.pnlVolDiscountScheme.Size = New System.Drawing.Size(763, 432)
        Me.pnlVolDiscountScheme.TabIndex = 2
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter2.Location = New System.Drawing.Point(0, 0)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(763, 3)
        Me.Splitter2.TabIndex = 1
        Me.Splitter2.TabStop = False
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.lblPCurr)
        Me.Panel5.Controls.Add(Me.lblLPCurr)
        Me.Panel5.Controls.Add(Me.Label27)
        Me.Panel5.Controls.Add(Me.Label33)
        Me.Panel5.Controls.Add(Me.Label35)
        Me.Panel5.Controls.Add(Me.Label36)
        Me.Panel5.Controls.Add(Me.Label37)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(692, 48)
        Me.Panel5.TabIndex = 64
        '
        'lblPCurr
        '
        Me.lblPCurr.AutoSize = True
        Me.lblPCurr.Location = New System.Drawing.Point(518, 12)
        Me.lblPCurr.Name = "lblPCurr"
        Me.lblPCurr.Size = New System.Drawing.Size(19, 13)
        Me.lblPCurr.TabIndex = 62
        Me.lblPCurr.Text = "(£)"
        '
        'lblLPCurr
        '
        Me.lblLPCurr.AutoSize = True
        Me.lblLPCurr.Location = New System.Drawing.Point(302, 13)
        Me.lblLPCurr.Name = "lblLPCurr"
        Me.lblLPCurr.Size = New System.Drawing.Size(19, 13)
        Me.lblLPCurr.TabIndex = 61
        Me.lblLPCurr.Text = "(£)"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(582, 13)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(46, 13)
        Me.Label27.TabIndex = 60
        Me.Label27.Text = "Scheme"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(480, 13)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(31, 13)
        Me.Label33.TabIndex = 59
        Me.Label33.Text = "Price"
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(360, 13)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(100, 23)
        Me.Label35.TabIndex = 58
        Me.Label35.Text = "Discount %"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(248, 13)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(50, 13)
        Me.Label36.TabIndex = 57
        Me.Label36.Text = "List Price"
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(16, 13)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(100, 23)
        Me.Label37.TabIndex = 56
        Me.Label37.Text = "Sample Type"
        '
        'txtCallOutCharge
        '
        Me.txtCallOutCharge.Location = New System.Drawing.Point(463, 9)
        Me.txtCallOutCharge.Name = "txtCallOutCharge"
        Me.txtCallOutCharge.Size = New System.Drawing.Size(72, 20)
        Me.txtCallOutCharge.TabIndex = 39
        '
        'Label50
        '
        Me.Label50.Location = New System.Drawing.Point(32, 9)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(312, 23)
        Me.Label50.TabIndex = 36
        Me.Label50.Text = "Call-Out Attendance Fee (Option)"
        '
        'pnlMinchargeheader
        '
        Me.pnlMinchargeheader.Controls.Add(Me.Label4)
        Me.pnlMinchargeheader.Controls.Add(Me.Label5)
        Me.pnlMinchargeheader.Controls.Add(Me.Label54)
        Me.pnlMinchargeheader.Controls.Add(Me.Label55)
        Me.pnlMinchargeheader.Controls.Add(Me.Label56)
        Me.pnlMinchargeheader.Controls.Add(Me.Label57)
        Me.pnlMinchargeheader.Controls.Add(Me.Label58)
        Me.pnlMinchargeheader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMinchargeheader.Location = New System.Drawing.Point(0, 40)
        Me.pnlMinchargeheader.Name = "pnlMinchargeheader"
        Me.pnlMinchargeheader.Size = New System.Drawing.Size(708, 48)
        Me.pnlMinchargeheader.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(518, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(19, 13)
        Me.Label4.TabIndex = 62
        Me.Label4.Text = "(£)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(302, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(19, 13)
        Me.Label5.TabIndex = 61
        Me.Label5.Text = "(£)"
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
        Me.pnlSelectCharging.Size = New System.Drawing.Size(708, 40)
        Me.pnlSelectCharging.TabIndex = 69
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
        Me.udCallOut.Items.Add("Service not available")
        Me.udCallOut.Location = New System.Drawing.Point(360, 8)
        Me.udCallOut.Name = "udCallOut"
        Me.udCallOut.Size = New System.Drawing.Size(120, 20)
        Me.udCallOut.TabIndex = 3
        Me.udCallOut.Tag = "CO"
        Me.udCallOut.Text = "Service not available"
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
        Me.udRoutine.Items.Add("Service not available")
        Me.udRoutine.Location = New System.Drawing.Point(128, 8)
        Me.udRoutine.Name = "udRoutine"
        Me.udRoutine.Size = New System.Drawing.Size(120, 20)
        Me.udRoutine.TabIndex = 1
        Me.udRoutine.Tag = "R"
        Me.udRoutine.Text = "Service not available"
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
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(480, 8)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(18, 13)
        Me.Label19.TabIndex = 62
        Me.Label19.Text = "(£)"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(264, 8)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(18, 13)
        Me.Label20.TabIndex = 61
        Me.Label20.Text = "(£)"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(544, 8)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(46, 13)
        Me.Label21.TabIndex = 60
        Me.Label21.Text = "Scheme"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(560, 13)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(46, 13)
        Me.Label10.TabIndex = 53
        Me.Label10.Text = "Scheme"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(458, 13)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 13)
        Me.Label9.TabIndex = 32
        Me.Label9.Text = "Price"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(338, 13)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 23)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Discount %"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(226, 13)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 13)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "List Price"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(58, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Sample Type "
        '
        'lblohadtrpricephr
        '
        Me.lblohadtrpricephr.Location = New System.Drawing.Point(528, 8)
        Me.lblohadtrpricephr.Name = "lblohadtrpricephr"
        Me.lblohadtrpricephr.Size = New System.Drawing.Size(40, 23)
        Me.lblohadtrpricephr.TabIndex = 42
        '
        'txtouthratcoScheme
        '
        Me.txtouthratcoScheme.Enabled = False
        Me.txtouthratcoScheme.Location = New System.Drawing.Point(578, 8)
        Me.txtouthratcoScheme.Name = "txtouthratcoScheme"
        Me.txtouthratcoScheme.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratcoScheme.TabIndex = 40
        '
        'txtouthratcoPrice
        '
        Me.txtouthratcoPrice.Location = New System.Drawing.Point(463, 9)
        Me.txtouthratcoPrice.Name = "txtouthratcoPrice"
        Me.txtouthratcoPrice.Size = New System.Drawing.Size(52, 20)
        Me.txtouthratcoPrice.TabIndex = 39
        '
        'txtouthratcoDiscount
        '
        Me.txtouthratcoDiscount.BackColor = System.Drawing.SystemColors.Info
        Me.txtouthratcoDiscount.Location = New System.Drawing.Point(343, 9)
        Me.txtouthratcoDiscount.Name = "txtouthratcoDiscount"
        Me.txtouthratcoDiscount.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratcoDiscount.TabIndex = 38
        '
        'txtouthratcoListPrice
        '
        Me.txtouthratcoListPrice.Enabled = False
        Me.txtouthratcoListPrice.Location = New System.Drawing.Point(231, 9)
        Me.txtouthratcoListPrice.Name = "txtouthratcoListPrice"
        Me.txtouthratcoListPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratcoListPrice.TabIndex = 37
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(32, 9)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(176, 23)
        Me.Label25.TabIndex = 36
        Me.Label25.Text = "Out hours Add Time (Call Out)"
        '
        'lblihadtrpricephr
        '
        Me.lblihadtrpricephr.Location = New System.Drawing.Point(528, 8)
        Me.lblihadtrpricephr.Name = "lblihadtrpricephr"
        Me.lblihadtrpricephr.Size = New System.Drawing.Size(40, 23)
        Me.lblihadtrpricephr.TabIndex = 42
        '
        'txtinhratcoScheme
        '
        Me.txtinhratcoScheme.Enabled = False
        Me.txtinhratcoScheme.Location = New System.Drawing.Point(578, 8)
        Me.txtinhratcoScheme.Name = "txtinhratcoScheme"
        Me.txtinhratcoScheme.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratcoScheme.TabIndex = 40
        '
        'txtinhratcoPrice
        '
        Me.txtinhratcoPrice.Location = New System.Drawing.Point(463, 9)
        Me.txtinhratcoPrice.Name = "txtinhratcoPrice"
        Me.txtinhratcoPrice.Size = New System.Drawing.Size(52, 20)
        Me.txtinhratcoPrice.TabIndex = 39
        '
        'txtinhratcoDiscount
        '
        Me.txtinhratcoDiscount.BackColor = System.Drawing.SystemColors.Info
        Me.txtinhratcoDiscount.Location = New System.Drawing.Point(343, 9)
        Me.txtinhratcoDiscount.Name = "txtinhratcoDiscount"
        Me.txtinhratcoDiscount.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratcoDiscount.TabIndex = 38
        '
        'txtinhratcoListPrice
        '
        Me.txtinhratcoListPrice.Enabled = False
        Me.txtinhratcoListPrice.Location = New System.Drawing.Point(231, 9)
        Me.txtinhratcoListPrice.Name = "txtinhratcoListPrice"
        Me.txtinhratcoListPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratcoListPrice.TabIndex = 37
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(32, 9)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(184, 23)
        Me.Label24.TabIndex = 36
        Me.Label24.Text = "In hours Add Time (Call Out)"
        '
        'lblohadtpricephr
        '
        Me.lblohadtpricephr.Location = New System.Drawing.Point(528, 9)
        Me.lblohadtpricephr.Name = "lblohadtpricephr"
        Me.lblohadtpricephr.Size = New System.Drawing.Size(40, 23)
        Me.lblohadtpricephr.TabIndex = 42
        '
        'txtouthratrtScheme
        '
        Me.txtouthratrtScheme.Enabled = False
        Me.txtouthratrtScheme.Location = New System.Drawing.Point(578, 8)
        Me.txtouthratrtScheme.Name = "txtouthratrtScheme"
        Me.txtouthratrtScheme.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratrtScheme.TabIndex = 40
        '
        'txtouthratrtPrice
        '
        Me.txtouthratrtPrice.Location = New System.Drawing.Point(463, 9)
        Me.txtouthratrtPrice.Name = "txtouthratrtPrice"
        Me.txtouthratrtPrice.Size = New System.Drawing.Size(52, 20)
        Me.txtouthratrtPrice.TabIndex = 39
        '
        'txtouthratrtDiscount
        '
        Me.txtouthratrtDiscount.BackColor = System.Drawing.SystemColors.Info
        Me.txtouthratrtDiscount.Location = New System.Drawing.Point(343, 9)
        Me.txtouthratrtDiscount.Name = "txtouthratrtDiscount"
        Me.txtouthratrtDiscount.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratrtDiscount.TabIndex = 38
        '
        'txtouthratrtListPrice
        '
        Me.txtouthratrtListPrice.Enabled = False
        Me.txtouthratrtListPrice.Location = New System.Drawing.Point(231, 9)
        Me.txtouthratrtListPrice.Name = "txtouthratrtListPrice"
        Me.txtouthratrtListPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtouthratrtListPrice.TabIndex = 37
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(32, 9)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(176, 23)
        Me.Label23.TabIndex = 36
        Me.Label23.Text = "Out hours Add Time (Routine)"
        '
        'lblihadtpricephr
        '
        Me.lblihadtpricephr.Location = New System.Drawing.Point(528, 8)
        Me.lblihadtpricephr.Name = "lblihadtpricephr"
        Me.lblihadtpricephr.Size = New System.Drawing.Size(40, 23)
        Me.lblihadtpricephr.TabIndex = 41
        '
        'txtinhratrtScheme
        '
        Me.txtinhratrtScheme.Enabled = False
        Me.txtinhratrtScheme.Location = New System.Drawing.Point(578, 8)
        Me.txtinhratrtScheme.Name = "txtinhratrtScheme"
        Me.txtinhratrtScheme.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratrtScheme.TabIndex = 40
        '
        'txtinhratrtPrice
        '
        Me.txtinhratrtPrice.Location = New System.Drawing.Point(463, 9)
        Me.txtinhratrtPrice.Name = "txtinhratrtPrice"
        Me.txtinhratrtPrice.Size = New System.Drawing.Size(52, 20)
        Me.txtinhratrtPrice.TabIndex = 39
        '
        'txtinhratrtDiscount
        '
        Me.txtinhratrtDiscount.BackColor = System.Drawing.SystemColors.Info
        Me.txtinhratrtDiscount.Location = New System.Drawing.Point(343, 9)
        Me.txtinhratrtDiscount.Name = "txtinhratrtDiscount"
        Me.txtinhratrtDiscount.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratrtDiscount.TabIndex = 38
        '
        'txtinhratrtListPrice
        '
        Me.txtinhratrtListPrice.Enabled = False
        Me.txtinhratrtListPrice.Location = New System.Drawing.Point(231, 9)
        Me.txtinhratrtListPrice.Name = "txtinhratrtListPrice"
        Me.txtinhratrtListPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtinhratrtListPrice.TabIndex = 37
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(32, 9)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(184, 23)
        Me.Label22.TabIndex = 36
        Me.Label22.Text = "In hours Add Time (Routine)"
        '
        'Panel3
        '
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(708, 48)
        Me.Panel3.TabIndex = 0
        '
        'lblpCollCurr
        '
        Me.lblpCollCurr.AutoSize = True
        Me.lblpCollCurr.Location = New System.Drawing.Point(518, 12)
        Me.lblpCollCurr.Name = "lblpCollCurr"
        Me.lblpCollCurr.Size = New System.Drawing.Size(18, 13)
        Me.lblpCollCurr.TabIndex = 62
        Me.lblpCollCurr.Text = "(£)"
        '
        'lbllpCollCurr
        '
        Me.lbllpCollCurr.AutoSize = True
        Me.lbllpCollCurr.Location = New System.Drawing.Point(302, 13)
        Me.lbllpCollCurr.Name = "lbllpCollCurr"
        Me.lbllpCollCurr.Size = New System.Drawing.Size(18, 13)
        Me.lbllpCollCurr.TabIndex = 61
        Me.lbllpCollCurr.Text = "(£)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(582, 13)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(46, 13)
        Me.Label14.TabIndex = 60
        Me.Label14.Text = "Scheme"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(480, 13)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(30, 13)
        Me.Label15.TabIndex = 59
        Me.Label15.Text = "Price"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(360, 13)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 23)
        Me.Label16.TabIndex = 58
        Me.Label16.Text = "Discount %"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(248, 13)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(51, 13)
        Me.Label17.TabIndex = 57
        Me.Label17.Text = "List Price"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(80, 13)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(100, 23)
        Me.Label18.TabIndex = 56
        Me.Label18.Text = "Min Charge Type "
        '
        'pnlPriceInfo
        '
        Me.pnlPriceInfo.Controls.Add(Me.Label42)
        Me.pnlPriceInfo.Controls.Add(Me.Label41)
        Me.pnlPriceInfo.Controls.Add(Me.Label61)
        Me.pnlPriceInfo.Controls.Add(Me.PictureBox2)
        Me.pnlPriceInfo.Controls.Add(Me.PictureBox1)
        Me.pnlPriceInfo.Controls.Add(Me.Label11)
        Me.pnlPriceInfo.Controls.Add(Me.Label72)
        Me.pnlPriceInfo.Controls.Add(Me.Label71)
        Me.pnlPriceInfo.Controls.Add(Me.Label13)
        Me.pnlPriceInfo.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPriceInfo.Location = New System.Drawing.Point(0, 441)
        Me.pnlPriceInfo.Name = "pnlPriceInfo"
        Me.pnlPriceInfo.Size = New System.Drawing.Size(771, 112)
        Me.pnlPriceInfo.TabIndex = 12
        '
        'Label42
        '
        Me.Label42.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label42.BackColor = System.Drawing.Color.Orange
        Me.Label42.Location = New System.Drawing.Point(192, 51)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(168, 23)
        Me.Label42.TabIndex = 70
        Me.Label42.Text = "Service to be available"
        '
        'Label41
        '
        Me.Label41.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label41.BackColor = System.Drawing.Color.CadetBlue
        Me.Label41.Location = New System.Drawing.Point(368, 51)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(168, 23)
        Me.Label41.TabIndex = 69
        Me.Label41.Text = "Service expired"
        '
        'Label61
        '
        Me.Label61.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label61.Location = New System.Drawing.Point(576, 75)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(136, 23)
        Me.Label61.TabIndex = 68
        Me.Label61.Text = "Volume Discount in place"
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(712, 75)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 67
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(552, 75)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'Label11
        '
        Me.Label11.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label11.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Label11.Location = New System.Drawing.Point(368, 75)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(168, 23)
        Me.Label11.TabIndex = 66
        Me.Label11.Text = "Service previously sold to client "
        '
        'Label72
        '
        Me.Label72.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label72.BackColor = System.Drawing.Color.PaleGreen
        Me.Label72.Location = New System.Drawing.Point(192, 75)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(168, 23)
        Me.Label72.TabIndex = 65
        Me.Label72.Text = "Service sold to client "
        '
        'Label71
        '
        Me.Label71.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label71.BackColor = System.Drawing.Color.Beige
        Me.Label71.Location = New System.Drawing.Point(16, 75)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(168, 23)
        Me.Label71.TabIndex = 64
        Me.Label71.Text = "Service not sold to client "
        '
        'Label13
        '
        Me.Label13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label13.BackColor = System.Drawing.Color.AntiqueWhite
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.DeepPink
        Me.Label13.Location = New System.Drawing.Point(32, 19)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(648, 32)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "All Prices are based on the published List Price and are discounted/surcharged fr" & _
            "om that, any changes to the list price will be reflected in a change to the pric" & _
            "e for ALL customers."
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Description/Test"
        Me.ColumnHeader1.Width = 150
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = ""
        Me.ColumnHeader2.Width = 40
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "List Price"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = ""
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Discount (%)"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = ""
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Price"
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Description/Kit"
        Me.ColumnHeader8.Width = 300
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = ""
        Me.ColumnHeader9.Width = 40
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "List Price"
        Me.ColumnHeader10.Width = 100
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = ""
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Discount (%)"
        Me.ColumnHeader12.Width = 100
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = ""
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Price"
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "#Invoices"
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "#Sold"
        '
        'chAvgPrice2
        '
        Me.chAvgPrice2.Text = "Avg Price"
        '
        'ColumnHeader24
        '
        Me.ColumnHeader24.Text = "Description/Kit"
        Me.ColumnHeader24.Width = 300
        '
        'ColumnHeader25
        '
        Me.ColumnHeader25.Text = ""
        Me.ColumnHeader25.Width = 40
        '
        'chListPriceKit
        '
        Me.chListPriceKit.Text = "List Price"
        Me.chListPriceKit.Width = 100
        '
        'ColumnHeader26
        '
        Me.ColumnHeader26.Text = ""
        '
        'ColumnHeader27
        '
        Me.ColumnHeader27.Text = "Discount (%)"
        Me.ColumnHeader27.Width = 100
        '
        'ColumnHeader28
        '
        Me.ColumnHeader28.Text = ""
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
        'ColumnHeader29
        '
        Me.ColumnHeader29.Text = "Description/Test"
        Me.ColumnHeader29.Width = 150
        '
        'ColumnHeader30
        '
        Me.ColumnHeader30.Text = ""
        Me.ColumnHeader30.Width = 40
        '
        'ColumnHeader31
        '
        Me.ColumnHeader31.Text = "List Price"
        Me.ColumnHeader31.Width = 100
        '
        'ColumnHeader32
        '
        Me.ColumnHeader32.Text = ""
        '
        'ColumnHeader33
        '
        Me.ColumnHeader33.Text = "Discount (%)"
        Me.ColumnHeader33.Width = 100
        '
        'ColumnHeader34
        '
        Me.ColumnHeader34.Text = ""
        '
        'ColumnHeader35
        '
        Me.ColumnHeader35.Text = "Price"
        '
        'ColumnHeader36
        '
        Me.ColumnHeader36.Text = "Description/Test"
        Me.ColumnHeader36.Width = 150
        '
        'ColumnHeader37
        '
        Me.ColumnHeader37.Text = ""
        Me.ColumnHeader37.Width = 40
        '
        'ColumnHeader38
        '
        Me.ColumnHeader38.Text = "List Price"
        Me.ColumnHeader38.Width = 100
        '
        'ColumnHeader39
        '
        Me.ColumnHeader39.Text = ""
        '
        'ColumnHeader40
        '
        Me.ColumnHeader40.Text = "Discount (%)"
        Me.ColumnHeader40.Width = 100
        '
        'ColumnHeader41
        '
        Me.ColumnHeader41.Text = ""
        '
        'ColumnHeader42
        '
        Me.ColumnHeader42.Text = "Price"
        '
        'ColumnHeader43
        '
        Me.ColumnHeader43.Text = "Description/Kit"
        Me.ColumnHeader43.Width = 300
        '
        'ColumnHeader44
        '
        Me.ColumnHeader44.Text = ""
        Me.ColumnHeader44.Width = 40
        '
        'ColumnHeader45
        '
        Me.ColumnHeader45.Text = "List Price"
        Me.ColumnHeader45.Width = 100
        '
        'ColumnHeader46
        '
        Me.ColumnHeader46.Text = ""
        '
        'ColumnHeader47
        '
        Me.ColumnHeader47.Text = "Discount (%)"
        Me.ColumnHeader47.Width = 100
        '
        'ColumnHeader48
        '
        Me.ColumnHeader48.Text = ""
        '
        'ColumnHeader49
        '
        Me.ColumnHeader49.Text = "Price"
        '
        'ColumnHeader50
        '
        Me.ColumnHeader50.Text = "#Invoices"
        '
        'ColumnHeader51
        '
        Me.ColumnHeader51.Text = "#Sold"
        '
        'ColumnHeader52
        '
        Me.ColumnHeader52.Text = "Avg Price"
        '
        'ColumnHeader53
        '
        Me.ColumnHeader53.Text = "Description/Kit"
        Me.ColumnHeader53.Width = 300
        '
        'ColumnHeader54
        '
        Me.ColumnHeader54.Text = ""
        Me.ColumnHeader54.Width = 40
        '
        'ColumnHeader55
        '
        Me.ColumnHeader55.Text = "List Price"
        Me.ColumnHeader55.Width = 100
        '
        'ColumnHeader56
        '
        Me.ColumnHeader56.Text = ""
        '
        'ColumnHeader57
        '
        Me.ColumnHeader57.Text = "Discount (%)"
        Me.ColumnHeader57.Width = 100
        '
        'ColumnHeader58
        '
        Me.ColumnHeader58.Text = ""
        '
        'ColumnHeader59
        '
        Me.ColumnHeader59.Text = "Price"
        '
        'ColumnHeader60
        '
        Me.ColumnHeader60.Text = "#Invoices"
        '
        'ColumnHeader61
        '
        Me.ColumnHeader61.Text = "#Sold"
        '
        'ColumnHeader62
        '
        Me.ColumnHeader62.Text = "Avg Price"
        '
        'pnlTimeAllowHead
        '
        Me.pnlTimeAllowHead.Controls.Add(Me.Label70)
        Me.pnlTimeAllowHead.Controls.Add(Me.txtTimeAllowance)
        Me.pnlTimeAllowHead.Controls.Add(Me.Label69)
        Me.pnlTimeAllowHead.Location = New System.Drawing.Point(3, 91)
        Me.pnlTimeAllowHead.Name = "pnlTimeAllowHead"
        Me.pnlTimeAllowHead.Size = New System.Drawing.Size(696, 136)
        Me.pnlTimeAllowHead.TabIndex = 63
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
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSavePanels})
        '
        'mnuSavePanels
        '
        Me.mnuSavePanels.Index = 0
        Me.mnuSavePanels.Text = "save panels"
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'pnlMRO
        '
        Me.pnlMRO.AllowLeave = True
        Me.pnlMRO.BackColor = System.Drawing.SystemColors.Control
        DiscountPanelBase1.AllowLeave = False
        DiscountPanelBase1.DiscountScheme = ""
        DiscountPanelBase1.Height = 120
        DiscountPanelBase1.Label = "MRO Analysis"
        DiscountPanelBase1.Left = 0
        DiscountPanelBase1.LineItem = ""
        DiscountPanelBase1.Name = Nothing
        DiscountPanelBase1.Panel = Nothing
        DiscountPanelBase1.PanelType = 0
        DiscountPanelBase1.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase1.Service = ""
        DiscountPanelBase1.Top = 0
        DiscountPanelBase1.Width = 735
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
        Me.pnlMRO.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlMRO.PriceEnabled = True
        Me.pnlMRO.Scheme = ""
        Me.pnlMRO.Service = ""
        Me.pnlMRO.size = New System.Drawing.Size(735, 120)
        Me.pnlMRO.size = New System.Drawing.Size(735, 120)
        Me.pnlMRO.TabIndex = 61
        '
        'lblMROComment
        '
        Me.lblMROComment.Location = New System.Drawing.Point(16, 64)
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
        DiscountPanelBase2.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase2.Service = "EXTERNAL"
        DiscountPanelBase2.Top = 0
        DiscountPanelBase2.Width = 735
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
        Me.pnlExternal.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlExternal.PriceEnabled = True
        Me.pnlExternal.Scheme = ""
        Me.pnlExternal.Service = "EXTERNAL"
        Me.pnlExternal.size = New System.Drawing.Size(735, 48)
        Me.pnlExternal.size = New System.Drawing.Size(735, 48)
        Me.pnlExternal.TabIndex = 48
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Info
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label6.Location = New System.Drawing.Point(288, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(352, 23)
        Me.Label6.TabIndex = 61
        Me.Label6.Text = "Time allowance can be removed by removing all text from box "
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
        DiscountPanelBase3.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase3.Service = Nothing
        DiscountPanelBase3.Top = 0
        DiscountPanelBase3.Width = 735
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
        Me.pnlFixedSite.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlFixedSite.PriceEnabled = True
        Me.pnlFixedSite.Scheme = ""
        Me.pnlFixedSite.Service = Nothing
        Me.pnlFixedSite.size = New System.Drawing.Size(735, 56)
        Me.pnlFixedSite.size = New System.Drawing.Size(735, 56)
        Me.pnlFixedSite.TabIndex = 57
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
        DiscountPanelBase4.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase4.Service = "CALLOUT"
        DiscountPanelBase4.Top = 0
        DiscountPanelBase4.Width = 735
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
        Me.pnlCallOut.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlCallOut.PriceEnabled = True
        Me.pnlCallOut.Scheme = ""
        Me.pnlCallOut.Service = "CALLOUT"
        Me.pnlCallOut.size = New System.Drawing.Size(735, 48)
        Me.pnlCallOut.size = New System.Drawing.Size(735, 48)
        Me.pnlCallOut.TabIndex = 58
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
        DiscountPanelBase5.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase5.Service = "ROUTINE"
        DiscountPanelBase5.Top = 0
        DiscountPanelBase5.Width = 735
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
        Me.pnlRoutine.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlRoutine.PriceEnabled = True
        Me.pnlRoutine.Scheme = ""
        Me.pnlRoutine.Service = "ROUTINE"
        Me.pnlRoutine.size = New System.Drawing.Size(735, 40)
        Me.pnlRoutine.size = New System.Drawing.Size(735, 40)
        Me.pnlRoutine.TabIndex = 59
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
        DiscountPanelBase6.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase6.Service = "ANALONLY"
        DiscountPanelBase6.Top = 0
        DiscountPanelBase6.Width = 735
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
        Me.pnlAO.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlAO.PriceEnabled = True
        Me.pnlAO.Scheme = ""
        Me.pnlAO.Service = "ANALONLY"
        Me.pnlAO.size = New System.Drawing.Size(735, 40)
        Me.pnlAO.size = New System.Drawing.Size(735, 40)
        Me.pnlAO.TabIndex = 63
        '
        'dpCallOutCancel
        '
        Me.dpCallOutCancel.AllowLeave = True
        Me.dpCallOutCancel.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase7.AllowLeave = False
        DiscountPanelBase7.DiscountScheme = ""
        DiscountPanelBase7.Height = 40
        DiscountPanelBase7.Label = "Call Out Cancellation"
        DiscountPanelBase7.Left = 0
        DiscountPanelBase7.LineItem = "APCOLC005"
        DiscountPanelBase7.Name = Nothing
        DiscountPanelBase7.Panel = Nothing
        DiscountPanelBase7.PanelType = 0
        DiscountPanelBase7.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase7.Service = "CALLOUT"
        DiscountPanelBase7.Top = 0
        DiscountPanelBase7.Width = 735
        Me.dpCallOutCancel.BasePanel = DiscountPanelBase7
        Me.dpCallOutCancel.Client = Nothing
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
        Me.dpCallOutCancel.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpCallOutCancel.PriceEnabled = True
        Me.dpCallOutCancel.Scheme = ""
        Me.dpCallOutCancel.Service = "CALLOUT"
        Me.dpCallOutCancel.size = New System.Drawing.Size(735, 40)
        Me.dpCallOutCancel.size = New System.Drawing.Size(735, 40)
        Me.dpCallOutCancel.TabIndex = 66
        '
        'dpFixedSiteCancellation
        '
        Me.dpFixedSiteCancellation.AllowLeave = True
        Me.dpFixedSiteCancellation.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase8.AllowLeave = False
        DiscountPanelBase8.DiscountScheme = ""
        DiscountPanelBase8.Height = 40
        DiscountPanelBase8.Label = "Fixed Site Cancellation"
        DiscountPanelBase8.Left = 0
        DiscountPanelBase8.LineItem = "APCOLF003"
        DiscountPanelBase8.Name = Nothing
        DiscountPanelBase8.Panel = Nothing
        DiscountPanelBase8.PanelType = 0
        DiscountPanelBase8.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase8.Service = "FIXED"
        DiscountPanelBase8.Top = 0
        DiscountPanelBase8.Width = 735
        Me.dpFixedSiteCancellation.BasePanel = DiscountPanelBase8
        Me.dpFixedSiteCancellation.Client = Nothing
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
        Me.dpFixedSiteCancellation.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpFixedSiteCancellation.PriceEnabled = True
        Me.dpFixedSiteCancellation.Scheme = ""
        Me.dpFixedSiteCancellation.Service = "FIXED"
        Me.dpFixedSiteCancellation.size = New System.Drawing.Size(735, 40)
        Me.dpFixedSiteCancellation.size = New System.Drawing.Size(735, 40)
        Me.dpFixedSiteCancellation.TabIndex = 65
        '
        'dpInHoursCancel
        '
        Me.dpInHoursCancel.AllowLeave = True
        Me.dpInHoursCancel.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase9.AllowLeave = False
        DiscountPanelBase9.DiscountScheme = ""
        DiscountPanelBase9.Height = 40
        DiscountPanelBase9.Label = "In Hours (Routine)"
        DiscountPanelBase9.Left = 0
        DiscountPanelBase9.LineItem = "APCOLS005"
        DiscountPanelBase9.Name = Nothing
        DiscountPanelBase9.Panel = Nothing
        DiscountPanelBase9.PanelType = 0
        DiscountPanelBase9.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase9.Service = "ROUTINE"
        DiscountPanelBase9.Top = 0
        DiscountPanelBase9.Width = 735
        Me.dpInHoursCancel.BasePanel = DiscountPanelBase9
        Me.dpInHoursCancel.Client = Nothing
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
        Me.dpInHoursCancel.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpInHoursCancel.PriceEnabled = True
        Me.dpInHoursCancel.Scheme = ""
        Me.dpInHoursCancel.Service = "ROUTINE"
        Me.dpInHoursCancel.size = New System.Drawing.Size(735, 40)
        Me.dpInHoursCancel.size = New System.Drawing.Size(735, 40)
        Me.dpInHoursCancel.TabIndex = 64
        '
        'dpOutOfhoursRoutine
        '
        Me.dpOutOfhoursRoutine.AllowLeave = True
        Me.dpOutOfhoursRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase10.AllowLeave = False
        DiscountPanelBase10.DiscountScheme = ""
        DiscountPanelBase10.Height = 40
        DiscountPanelBase10.Label = "Out of Hours (Routine)"
        DiscountPanelBase10.Left = 0
        DiscountPanelBase10.LineItem = "APCOLS006"
        DiscountPanelBase10.Name = Nothing
        DiscountPanelBase10.Panel = Nothing
        DiscountPanelBase10.PanelType = 0
        DiscountPanelBase10.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase10.Service = "ROUTINE"
        DiscountPanelBase10.Top = 0
        DiscountPanelBase10.Width = 735
        Me.dpOutOfhoursRoutine.BasePanel = DiscountPanelBase10
        Me.dpOutOfhoursRoutine.Client = Nothing
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
        Me.dpOutOfhoursRoutine.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpOutOfhoursRoutine.PriceEnabled = True
        Me.dpOutOfhoursRoutine.Scheme = ""
        Me.dpOutOfhoursRoutine.Service = "ROUTINE"
        Me.dpOutOfhoursRoutine.size = New System.Drawing.Size(735, 40)
        Me.dpOutOfhoursRoutine.size = New System.Drawing.Size(735, 40)
        Me.dpOutOfhoursRoutine.TabIndex = 63
        '
        'dpCallOutCollChargeOut
        '
        Me.dpCallOutCollChargeOut.AllowLeave = True
        Me.dpCallOutCollChargeOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase11.AllowLeave = False
        DiscountPanelBase11.DiscountScheme = ""
        DiscountPanelBase11.Height = 40
        DiscountPanelBase11.Label = "Call Out Attendance Fee (Out Of Hours)"
        DiscountPanelBase11.Left = 0
        DiscountPanelBase11.LineItem = ""
        DiscountPanelBase11.Name = Nothing
        DiscountPanelBase11.Panel = Nothing
        DiscountPanelBase11.PanelType = 0
        DiscountPanelBase11.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase11.Service = "CALLOUTOUT"
        DiscountPanelBase11.Top = 0
        DiscountPanelBase11.Width = 751
        Me.dpCallOutCollChargeOut.BasePanel = DiscountPanelBase11
        Me.dpCallOutCollChargeOut.Client = Nothing
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
        Me.dpCallOutCollChargeOut.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpCallOutCollChargeOut.PriceEnabled = True
        Me.dpCallOutCollChargeOut.Scheme = ""
        Me.dpCallOutCollChargeOut.Service = "CALLOUTOUT"
        Me.dpCallOutCollChargeOut.size = New System.Drawing.Size(751, 40)
        Me.dpCallOutCollChargeOut.size = New System.Drawing.Size(751, 40)
        Me.dpCallOutCollChargeOut.TabIndex = 75
        '
        'dpCallOutCollCharge
        '
        Me.dpCallOutCollCharge.AllowLeave = True
        Me.dpCallOutCollCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase12.AllowLeave = False
        DiscountPanelBase12.DiscountScheme = ""
        DiscountPanelBase12.Height = 40
        DiscountPanelBase12.Label = "Call Out Attendance Fee"
        DiscountPanelBase12.Left = 0
        DiscountPanelBase12.LineItem = ""
        DiscountPanelBase12.Name = Nothing
        DiscountPanelBase12.Panel = Nothing
        DiscountPanelBase12.PanelType = 0
        DiscountPanelBase12.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase12.Service = "CALLOUT"
        DiscountPanelBase12.Top = 0
        DiscountPanelBase12.Width = 751
        Me.dpCallOutCollCharge.BasePanel = DiscountPanelBase12
        Me.dpCallOutCollCharge.Client = Nothing
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
        Me.dpCallOutCollCharge.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpCallOutCollCharge.PriceEnabled = True
        Me.dpCallOutCollCharge.Scheme = ""
        Me.dpCallOutCollCharge.Service = "CALLOUT"
        Me.dpCallOutCollCharge.size = New System.Drawing.Size(751, 40)
        Me.dpCallOutCollCharge.size = New System.Drawing.Size(751, 40)
        Me.dpCallOutCollCharge.TabIndex = 73
        '
        'dpRoutineCollChargeOut
        '
        Me.dpRoutineCollChargeOut.AllowLeave = True
        Me.dpRoutineCollChargeOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase13.AllowLeave = False
        DiscountPanelBase13.DiscountScheme = ""
        DiscountPanelBase13.Height = 40
        DiscountPanelBase13.Label = "Collection Attendance Fee (Out Of Hours)"
        DiscountPanelBase13.Left = 0
        DiscountPanelBase13.LineItem = ""
        DiscountPanelBase13.Name = Nothing
        DiscountPanelBase13.Panel = Nothing
        DiscountPanelBase13.PanelType = 0
        DiscountPanelBase13.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase13.Service = "ROUTINEOUT"
        DiscountPanelBase13.Top = 0
        DiscountPanelBase13.Width = 751
        Me.dpRoutineCollChargeOut.BasePanel = DiscountPanelBase13
        Me.dpRoutineCollChargeOut.Client = Nothing
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
        Me.dpRoutineCollChargeOut.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpRoutineCollChargeOut.PriceEnabled = True
        Me.dpRoutineCollChargeOut.Scheme = ""
        Me.dpRoutineCollChargeOut.Service = "ROUTINEOUT"
        Me.dpRoutineCollChargeOut.size = New System.Drawing.Size(751, 40)
        Me.dpRoutineCollChargeOut.size = New System.Drawing.Size(751, 40)
        Me.dpRoutineCollChargeOut.TabIndex = 74
        '
        'dpRoutineCollCharge
        '
        Me.dpRoutineCollCharge.AllowLeave = True
        Me.dpRoutineCollCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase14.AllowLeave = False
        DiscountPanelBase14.DiscountScheme = ""
        DiscountPanelBase14.Height = 40
        DiscountPanelBase14.Label = "Collection Attendance Fee"
        DiscountPanelBase14.Left = 0
        DiscountPanelBase14.LineItem = ""
        DiscountPanelBase14.Name = Nothing
        DiscountPanelBase14.Panel = Nothing
        DiscountPanelBase14.PanelType = 0
        DiscountPanelBase14.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase14.Service = "ROUTINE"
        DiscountPanelBase14.Top = 0
        DiscountPanelBase14.Width = 751
        Me.dpRoutineCollCharge.BasePanel = DiscountPanelBase14
        Me.dpRoutineCollCharge.Client = Nothing
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
        Me.dpRoutineCollCharge.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpRoutineCollCharge.PriceEnabled = True
        Me.dpRoutineCollCharge.Scheme = ""
        Me.dpRoutineCollCharge.Service = "ROUTINE"
        Me.dpRoutineCollCharge.size = New System.Drawing.Size(751, 40)
        Me.dpRoutineCollCharge.size = New System.Drawing.Size(751, 40)
        Me.dpRoutineCollCharge.TabIndex = 72
        '
        'dpMCCallOutOut
        '
        Me.dpMCCallOutOut.AllowLeave = True
        Me.dpMCCallOutOut.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase15.AllowLeave = False
        DiscountPanelBase15.DiscountScheme = ""
        DiscountPanelBase15.Height = 40
        DiscountPanelBase15.Label = "Call Out Minimum Charge (out hours)"
        DiscountPanelBase15.Left = 0
        DiscountPanelBase15.LineItem = "APCOLC007"
        DiscountPanelBase15.Name = Nothing
        DiscountPanelBase15.Panel = Nothing
        DiscountPanelBase15.PanelType = 0
        DiscountPanelBase15.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase15.Service = "CALLOUTOUT"
        DiscountPanelBase15.Top = 0
        DiscountPanelBase15.Width = 751
        Me.dpMCCallOutOut.BasePanel = DiscountPanelBase15
        Me.dpMCCallOutOut.Client = Nothing
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
        Me.dpMCCallOutOut.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpMCCallOutOut.PriceEnabled = True
        Me.dpMCCallOutOut.Scheme = ""
        Me.dpMCCallOutOut.Service = "CALLOUTOUT"
        Me.dpMCCallOutOut.size = New System.Drawing.Size(751, 40)
        Me.dpMCCallOutOut.size = New System.Drawing.Size(751, 40)
        Me.dpMCCallOutOut.TabIndex = 77
        '
        'dpMCCallOutIn
        '
        Me.dpMCCallOutIn.AllowLeave = True
        Me.dpMCCallOutIn.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase16.AllowLeave = False
        DiscountPanelBase16.DiscountScheme = ""
        DiscountPanelBase16.Height = 40
        DiscountPanelBase16.Label = "Call Out minimum charge (in hours)"
        DiscountPanelBase16.Left = 0
        DiscountPanelBase16.LineItem = "APCOLC006"
        DiscountPanelBase16.Name = Nothing
        DiscountPanelBase16.Panel = Nothing
        DiscountPanelBase16.PanelType = 0
        DiscountPanelBase16.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase16.Service = "CALLOUT"
        DiscountPanelBase16.Top = 0
        DiscountPanelBase16.Width = 751
        Me.dpMCCallOutIn.BasePanel = DiscountPanelBase16
        Me.dpMCCallOutIn.Client = Nothing
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
        Me.dpMCCallOutIn.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.dpMCCallOutIn.PriceEnabled = True
        Me.dpMCCallOutIn.Scheme = ""
        Me.dpMCCallOutIn.Service = "CALLOUT"
        Me.dpMCCallOutIn.size = New System.Drawing.Size(751, 40)
        Me.dpMCCallOutIn.size = New System.Drawing.Size(751, 40)
        Me.dpMCCallOutIn.TabIndex = 76
        '
        'pnlOutHoursMC
        '
        Me.pnlOutHoursMC.AllowLeave = True
        Me.pnlOutHoursMC.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase17.AllowLeave = False
        DiscountPanelBase17.DiscountScheme = ""
        DiscountPanelBase17.Height = 40
        DiscountPanelBase17.Label = "Out of Hours minimum charge"
        DiscountPanelBase17.Left = 0
        DiscountPanelBase17.LineItem = ""
        DiscountPanelBase17.Name = Nothing
        DiscountPanelBase17.Panel = Nothing
        DiscountPanelBase17.PanelType = 0
        DiscountPanelBase17.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase17.Service = "ROUTINEOUT"
        DiscountPanelBase17.Top = 0
        DiscountPanelBase17.Width = 751
        Me.pnlOutHoursMC.BasePanel = DiscountPanelBase17
        Me.pnlOutHoursMC.Client = Nothing
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
        Me.pnlOutHoursMC.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlOutHoursMC.PriceEnabled = True
        Me.pnlOutHoursMC.Scheme = ""
        Me.pnlOutHoursMC.Service = "ROUTINEOUT"
        Me.pnlOutHoursMC.size = New System.Drawing.Size(751, 40)
        Me.pnlOutHoursMC.size = New System.Drawing.Size(751, 40)
        Me.pnlOutHoursMC.TabIndex = 69
        '
        'pnlInHourMinCharge
        '
        Me.pnlInHourMinCharge.AllowLeave = True
        Me.pnlInHourMinCharge.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase18.AllowLeave = False
        DiscountPanelBase18.DiscountScheme = ""
        DiscountPanelBase18.Height = 40
        DiscountPanelBase18.Label = "In hours minimum charge"
        DiscountPanelBase18.Left = 0
        DiscountPanelBase18.LineItem = ""
        DiscountPanelBase18.Name = Nothing
        DiscountPanelBase18.Panel = Nothing
        DiscountPanelBase18.PanelType = 0
        DiscountPanelBase18.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase18.Service = "ROUTINE"
        DiscountPanelBase18.Top = 0
        DiscountPanelBase18.Width = 751
        Me.pnlInHourMinCharge.BasePanel = DiscountPanelBase18
        Me.pnlInHourMinCharge.Client = Nothing
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
        Me.pnlInHourMinCharge.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlInHourMinCharge.PriceEnabled = True
        Me.pnlInHourMinCharge.Scheme = ""
        Me.pnlInHourMinCharge.Service = "ROUTINE"
        Me.pnlInHourMinCharge.size = New System.Drawing.Size(751, 40)
        Me.pnlInHourMinCharge.size = New System.Drawing.Size(751, 40)
        Me.pnlInHourMinCharge.TabIndex = 68
        '
        'pnlOutHoursAddTimeCO
        '
        Me.pnlOutHoursAddTimeCO.AllowLeave = True
        Me.pnlOutHoursAddTimeCO.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase19.AllowLeave = False
        DiscountPanelBase19.DiscountScheme = ""
        DiscountPanelBase19.Height = 40
        DiscountPanelBase19.Label = "Out hours Add Time (Call Out)"
        DiscountPanelBase19.Left = 0
        DiscountPanelBase19.LineItem = ""
        DiscountPanelBase19.Name = Nothing
        DiscountPanelBase19.Panel = Nothing
        DiscountPanelBase19.PanelType = 0
        DiscountPanelBase19.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase19.Service = "CALLOUT"
        DiscountPanelBase19.Top = 0
        DiscountPanelBase19.Width = 737
        Me.pnlOutHoursAddTimeCO.BasePanel = DiscountPanelBase19
        Me.pnlOutHoursAddTimeCO.Client = Nothing
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
        Me.pnlOutHoursAddTimeCO.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlOutHoursAddTimeCO.PriceEnabled = True
        Me.pnlOutHoursAddTimeCO.Scheme = ""
        Me.pnlOutHoursAddTimeCO.Service = "CALLOUT"
        Me.pnlOutHoursAddTimeCO.size = New System.Drawing.Size(737, 40)
        Me.pnlOutHoursAddTimeCO.size = New System.Drawing.Size(737, 40)
        Me.pnlOutHoursAddTimeCO.TabIndex = 59
        '
        'pnlINHoursAddTimeCO
        '
        Me.pnlINHoursAddTimeCO.AllowLeave = True
        Me.pnlINHoursAddTimeCO.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase20.AllowLeave = False
        DiscountPanelBase20.DiscountScheme = ""
        DiscountPanelBase20.Height = 40
        DiscountPanelBase20.Label = "In hours Add Time (Call Out)"
        DiscountPanelBase20.Left = 0
        DiscountPanelBase20.LineItem = ""
        DiscountPanelBase20.Name = Nothing
        DiscountPanelBase20.Panel = Nothing
        DiscountPanelBase20.PanelType = 0
        DiscountPanelBase20.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase20.Service = "CALLOUT"
        DiscountPanelBase20.Top = 0
        DiscountPanelBase20.Width = 737
        Me.pnlINHoursAddTimeCO.BasePanel = DiscountPanelBase20
        Me.pnlINHoursAddTimeCO.Client = Nothing
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
        Me.pnlINHoursAddTimeCO.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlINHoursAddTimeCO.PriceEnabled = True
        Me.pnlINHoursAddTimeCO.Scheme = ""
        Me.pnlINHoursAddTimeCO.Service = "CALLOUT"
        Me.pnlINHoursAddTimeCO.size = New System.Drawing.Size(737, 40)
        Me.pnlINHoursAddTimeCO.size = New System.Drawing.Size(737, 40)
        Me.pnlINHoursAddTimeCO.TabIndex = 58
        '
        'pnlOutHoursAddTimeRoutine
        '
        Me.pnlOutHoursAddTimeRoutine.AllowLeave = True
        Me.pnlOutHoursAddTimeRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase21.AllowLeave = False
        DiscountPanelBase21.DiscountScheme = ""
        DiscountPanelBase21.Height = 40
        DiscountPanelBase21.Label = "Out hours Add Time (Routine)"
        DiscountPanelBase21.Left = 0
        DiscountPanelBase21.LineItem = ""
        DiscountPanelBase21.Name = Nothing
        DiscountPanelBase21.Panel = Nothing
        DiscountPanelBase21.PanelType = 0
        DiscountPanelBase21.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase21.Service = "ROUTINE"
        DiscountPanelBase21.Top = 0
        DiscountPanelBase21.Width = 737
        Me.pnlOutHoursAddTimeRoutine.BasePanel = DiscountPanelBase21
        Me.pnlOutHoursAddTimeRoutine.Client = Nothing
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
        Me.pnlOutHoursAddTimeRoutine.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlOutHoursAddTimeRoutine.PriceEnabled = True
        Me.pnlOutHoursAddTimeRoutine.Scheme = ""
        Me.pnlOutHoursAddTimeRoutine.Service = "ROUTINE"
        Me.pnlOutHoursAddTimeRoutine.size = New System.Drawing.Size(737, 40)
        Me.pnlOutHoursAddTimeRoutine.size = New System.Drawing.Size(737, 40)
        Me.pnlOutHoursAddTimeRoutine.TabIndex = 57
        '
        'pnlINHoursAddTimeRoutine
        '
        Me.pnlINHoursAddTimeRoutine.AllowLeave = True
        Me.pnlINHoursAddTimeRoutine.BackColor = System.Drawing.Color.Beige
        DiscountPanelBase22.AllowLeave = False
        DiscountPanelBase22.DiscountScheme = ""
        DiscountPanelBase22.Height = 40
        DiscountPanelBase22.Label = "In hours Add Time (Routine)"
        DiscountPanelBase22.Left = 0
        DiscountPanelBase22.LineItem = ""
        DiscountPanelBase22.Name = Nothing
        DiscountPanelBase22.Panel = Nothing
        DiscountPanelBase22.PanelType = 0
        DiscountPanelBase22.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        DiscountPanelBase22.Service = "ROUTINE"
        DiscountPanelBase22.Top = 0
        DiscountPanelBase22.Width = 737
        Me.pnlINHoursAddTimeRoutine.BasePanel = DiscountPanelBase22
        Me.pnlINHoursAddTimeRoutine.Client = Nothing
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
        Me.pnlINHoursAddTimeRoutine.PriceDate = New Date(2009, 4, 27, 0, 0, 0, 0)
        Me.pnlINHoursAddTimeRoutine.PriceEnabled = True
        Me.pnlINHoursAddTimeRoutine.Scheme = ""
        Me.pnlINHoursAddTimeRoutine.Service = "ROUTINE"
        Me.pnlINHoursAddTimeRoutine.size = New System.Drawing.Size(737, 40)
        Me.pnlINHoursAddTimeRoutine.size = New System.Drawing.Size(737, 40)
        Me.pnlINHoursAddTimeRoutine.TabIndex = 56
        '
        'VolDiscountPanel1
        '
        Me.VolDiscountPanel1.CustomerID = Nothing
        Me.VolDiscountPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.VolDiscountPanel1.Location = New System.Drawing.Point(0, 0)
        Me.VolDiscountPanel1.Name = "VolDiscountPanel1"
        Me.VolDiscountPanel1.Size = New System.Drawing.Size(763, 432)
        Me.VolDiscountPanel1.TabIndex = 0
        '
        'FrmPricing
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(771, 614)
        Me.Controls.Add(Me.tabTypes)
        Me.Controls.Add(Me.pnlPriceInfo)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmPricing"
        Me.Text = "Pricing "
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.tabTypes.ResumeLayout(False)
        Me.tabpSample.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.tabpCancellations.ResumeLayout(False)
        Me.pnlCancellations.ResumeLayout(False)
        Me.pnlCancellations.PerformLayout()
        Me.tabpMinCharge.ResumeLayout(False)
        Me.pnlMinCharge.ResumeLayout(False)
        Me.tabpAddTime.ResumeLayout(False)
        Me.pnlAddTime.ResumeLayout(False)
        Me.pnlADHeader.ResumeLayout(False)
        Me.pnlADHeader.PerformLayout()
        Me.pnlTimeAllowance.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.tabpTests.ResumeLayout(False)
        Me.pnlTests.ResumeLayout(False)
        Me.pnlTests.PerformLayout()
        Me.tabInvoicing.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.TabDiscounts.ResumeLayout(False)
        Me.pnlVolDiscountScheme.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.pnlMinchargeheader.ResumeLayout(False)
        Me.pnlMinchargeheader.PerformLayout()
        Me.pnlSelectCharging.ResumeLayout(False)
        Me.pnlPriceInfo.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTimeAllowHead.ResumeLayout(False)
        Me.pnlTimeAllowHead.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMRO.ResumeLayout(False)
        Me.pnlMRO.PerformLayout()
        Me.pnlExternal.ResumeLayout(False)
        Me.pnlExternal.PerformLayout()
        Me.pnlFixedSite.ResumeLayout(False)
        Me.pnlFixedSite.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private MCActionChar As Char = "I"
    Private COActionChar As Char = "I"
    Private Const cstNoAddTimeChargeScheme As Integer = 19
    Private Const cstNoTestChargeScheme As Integer = 359
    Private blnDisableAll As Boolean = True
    Private lvDItem As MedscreenCommonGui.ListViewItems.DiscountSchemes.TestDiscountEntry
    Private blnDrawing As Boolean = True
    Private blnPriceChanged As Boolean = False          'Indicator that a price has changed
    Private blnClientPanelFilled As Boolean = False
    Private myPanels As New Panels.DiscPanelCollCollection()

    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            myClient = Value
            blnDrawing = True
            blnClientPanelFilled = False
            PricingSupport.ChangeDate = Today                           'Initialise price change date
            If myClient Is Nothing Then Exit Property
            Me.pnlAO.Client = Value
            Me.pnlRoutine.Client = Value
            'Me.pnlBenzene.Client = Value
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

            Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet
            For Each myPnlColl In myPanels
                Dim BasePanel As Panels.DiscountPanelBase
                For Each BasePanel In myPnlColl.Panels
                    CType(BasePanel.Panel, Panels.DiscountPanel).Client = Value
                Next
            Next
            SetCustomer()

            DoPanelFill()
            Dim strCurr As String = (5).ToString("c", myClient.Culture).Chars(0)
            Me.lblLPCurr.Text = "(" & strCurr & ")"
            Me.lblPCurr.Text = "(" & strCurr & ")"
            Me.lbllpCollCurr.Text = "(" & strCurr & ")"
            Me.lblpCollCurr.Text = "(" & strCurr & ")"
            Me.lbllPADCurr.Text = "(" & strCurr & ")"
            Me.lbladpCurr.Text = "(" & strCurr & ")"
            Me.Label2.Text = "(" & strCurr & ")"
            Me.Label4.Text = "(" & strCurr & ")"
            Me.Label5.Text = "(" & strCurr & ")"
            Me.Label19.Text = "(" & strCurr & ")"
            Me.Label20.Text = "(" & strCurr & ")"

            Me.chListPrice.Text = "List Price (" & strCurr & ")"
            Me.chPrice.Text = "Price (" & strCurr & ")"

            If myClient.HasMRO Then
                Me.lblMROComment.Text = "MRO is " & myClient.ReportingInfo.MROID
            Else
                Me.lblMROComment.Text = "No MRO set for this customer"
            End If
            SetUpTimeAllowance()
            Me.FillTestListView()
            Me.Text = "Pricing setup for " & Me.myClient.SMIDProfile
            'blnDrawing = False
        End Set
    End Property

    Private Const cstStandardPricing = 0
    'Private Const cstOldPricing = 2
    Private Const cstAlternatePricing = 1

    Private intPricingOption As Integer = -1

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


    Public Sub SetCustomer()
        If Me.myClient Is Nothing Then Exit Sub
        Me.ckPORequired.Checked = Client.InvoiceInfo.PoRequired
        Me.ErrorProvider1.SetError(Me.txtSageID, "")                'Clear errors on sage number
        Me.txtSageID.Text = Client.SageID

        Try
            Me.txtVATRegno.Text = Client.VATRegno
            Me.txtVATRegno_Validating(Nothing, Nothing)

            Me.cbCurrency.SelectedIndex = -1
            'Currency 
            'Fix odd currencies
            If Client.Currency = "G" Then Client.Currency = "S"
            If Client.Currency = "U" Then Client.Currency = "D"

            Dim i As Integer
            Dim objCurr As MedscreenLib.Glossary.Currency
            Dim inRoutineMCDisc As Double = 0
            Dim inCallOutMCDisc As Double = 0
            Dim inRoutineCCDisc As Double = 0
            Dim inCallOutCCDisc As Double = 0

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
                MedscreenLib.Medscreen.LogError(ex, , "frmCustomerDetails-SetCustomer-4946 currency pos " & i)
            Finally
            End Try
            LogTiming("Status")

            If Client.InvoiceInfo.InvFrequency.Trim.Length > 0 Then
                Me.cbInvFreq.SelectedValue = Client.InvoiceInfo.InvFrequency
            End If
            Try  ' Protecting load document which is giving an error
                'Me.HtmlEditor1.DocumentText =("<HTML><BODY>No Info</BODY></HTML>")
            Catch
            Finally
            End Try
            FillPrices()

            'Pricing 

            Me.cbPricing.SelectedIndex = -1
            'If Client.StandardPricing Then
            '    Me.cbPricing.SelectedIndex = Me.cstStandardPricing
            '    'FillStandardPricing()
            'ElseIf Client.AlternatePricing Then
            '    Me.cbPricing.SelectedIndex = Me.cstAlternatePricing
            'End If
            Me.cbPricing.SelectedValue = Client.PriceBook
            LogTiming("Pricing")
            Me.intPricingOption = Me.cbPricing.SelectedIndex
            'deal with MROS 
            'DEal with setting up minimum charges but only for discounted pricing 
            'SetMinChargePanels(actionChar, Tag)
            'Code modified 10-Apr-2007 14:37 by taylor

            Me.CheckMinCharges()
            VolDiscountPanel1.DrawDiscounts()
            VolDiscountPanel1.DrawVolumeDiscounts()

        Catch ex As Exception
            LogError(ex, True, )
        End Try

        Me.ckCreditCard.Checked = Client.InvoiceInfo.CreditCard_pay
        Me.cbPlatAccType.SelectedIndex = -1

        Dim I1 As Integer
        Dim objPhrase As MedscreenLib.Glossary.Phrase
        If Client.InvoiceInfo.PlatAccType.Trim.Length > 0 Then
            For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Count - 1
                objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Current()
                If objPhrase.PhraseID = Client.InvoiceInfo.PlatAccType Then
                    Exit For
                End If
                Me.BindingContext(MedscreenLib.Glossary.Glossary.Platinum).Position += 1
            Next
        End If

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
                ElseIf RetVal = customerns.Client.MinChargePricingBy.AttendenceFee Then
                    Me.udCallOut.SelectedIndex = 1
                    COActionChar = "A"
                ElseIf RetVal = customerns.Client.MinChargePricingBy.MinimumCharge Then
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
    ''' Fill the individual panels with pricing information
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillPanels()
        LogTiming("    Fill Panel AO")
        Dim strService As String
        Dim strLineItem As String
        Dim blnVisible As Boolean
        Dim strPanelLabel As String
        pnlSelectCharging.Show()
        FillAPanel(Me.pnlAO)
        FillAPanel(Me.pnlRoutine) ', strLineItem, strService, blnVisible) '"APCOLS002", "ROUTINE")
        LogTiming("    Fill Panel Callout")
        'FillAPanel(Me.pnlBenzene)

        FillAPanel(Me.pnlCallOut) ', strLineItem, strService, blnVisible) ' "APCOLC002", "CALLOUT")
        LogTiming("    Fill Panel Fixed")


        FillAPanel(Me.pnlFixedSite) ', strLineItem, strService, blnVisible) ' "APCOLF002", "FIXED")
        Me.pnlExternal.Show()
        LogTiming("    Fill Panel External")

        FillAPanel(Me.pnlExternal) ', strLineItem, strService, blnVisible) ' "APCOLE002", "EXTERNAL")
        Me.pnlMRO.Show()
        LogTiming("    Fill Panel 008 Routine")

        FillAPanel(Me.pnlMRO) ', strLineItem, strService, blnVisible) ' "APCOL0008", "ROUTINE")
        Me.pnlInHourMinCharge.Show()
        LogTiming("    Fill Panel 003 Routine")

        FillAPanel(Me.pnlInHourMinCharge) ', strLineItem, strService, blnVisible) '"APCOL0003", "ROUTINE,CALLOUT")

        Me.pnlOutHoursMC.Show()

        FillAPanel(Me.pnlOutHoursMC) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        Me.dpMCCallOutIn.Show()

        FillAPanel(Me.dpMCCallOutIn) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        Me.dpMCCallOutOut.Show()

        FillAPanel(Me.dpMCCallOutOut) ', strLineItem, strService, blnVisible) '"APCOL0004", "ROUTINE,CALLOUT")

        LogTiming("    Fill Panel add time Routine")


        FillAPanel(Me.pnlINHoursAddTimeRoutine) ', strLineItem, strService, blnVisible) '"APCOLS003", "ROUTINE")
        FillAPanel(Me.pnlINHoursAddTimeCO) ', strLineItem, strService, blnVisible) '"APCOLC003", "CALLOUT")
        Me.pnlOutHoursAddTimeCO.Show()
        LogTiming("    Fill Panel add time Routine")

        FillAPanel(Me.pnlOutHoursAddTimeRoutine) ', strLineItem, strService, blnVisible) '"APCOLS004", "ROUTINE")
        Me.pnlOutHoursAddTimeRoutine.Show()
        LogTiming("    Fill Panel out co Routine")

        FillAPanel(Me.pnlOutHoursAddTimeCO) ', strLineItem, strService, blnVisible) '"APCOLC004", "CALLOUT")
        'Me.pnlCallOutCharge.Show()
        LogTiming("    Fill Panel on call ")
        Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet
        For Each myPnlColl In myPanels
            Dim BasePanel As Panels.DiscountPanelBase
            For Each BasePanel In myPnlColl.Panels
                FillAPanel(BasePanel.Panel)
            Next
        Next


        'blnVisible = (CCToolAppMessages.PanelVisibility.Item(Me.pnlCallOutCharge.Name) = "TRUE")
        'FillAPanel(Me.pnlCallOutCharge) ', "ONCALLCHGE", blnVisible)
        'Me.pnlRCOOption.Show()
        LogTiming("    Fill  option Panel Routine")

        'blnVisible = (CCToolAppMessages.PanelVisibility.Item(Me.pnlRCOOption.Name) = "TRUE")
        'FillOptionPanel(Me.pnlRCOOption, "RTNCOLCHGE", blnVisible)
        Me.pnlTimeAllowHead.Show()
        LogTiming("    Fill Panel dp call")


        FillAPanel(Me.dpCallOutCancel) ', strLineItem, strService, blnVisible) ' "APCOLC005", "CALLOUT")
        FillAPanel(Me.dpFixedSiteCancellation) ', strLineItem, strService, blnVisible) ' "APCOLF003", "FIXED")

        FillAPanel(Me.dpInHoursCancel) ', strLineItem, strService, blnVisible) '"APCOLS005", "ROUTINE")

        FillAPanel(Me.dpOutOfhoursRoutine) ', strLineItem, strService, blnVisible) ' "APCOLS006", "ROUTINE")
        LogTiming("    Fill Panel dp in h cancel")
        dpInHoursCancel.PanelLabel = "Routine In-Hours Cancellation"
        dpOutOfhoursRoutine.Show()

        'Routine Collection Charges 

        FillAPanel(Me.dpRoutineCollCharge) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpRoutineCollChargeOut) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpCallOutCollCharge) ', strLineItem, strService, blnVisible)

        FillAPanel(Me.dpCallOutCollChargeOut) ', strLineItem, strService, blnVisible)
        SetMinChargePanels(MCActionChar, "R")
        SetMinChargePanels(COActionChar, "CO")

        Application.DoEvents()
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
        dpRoutineCollCharge.Hide()
        dpRoutineCollChargeOut.Hide()
        dpCallOutCollCharge.Show()
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
            LogTiming("       Fill Panel - Client")
            objPanel.Client = Me.Client
            objPanel.AllowLeave = False
            objPanel.Service = Service
            objPanel.PriceDate = Me.DTPricing.Value
            If PanelLabel.Trim.Length > 0 Then objPanel.PanelLabel = PanelLabel
            LogTiming("       Fill Panel - redraw")

            objPanel.reDraw(Me.blnDisableAll)
            objPanel.AllowLeave = True
            LogTiming("       Fill Panel - image")

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
        Panel.ContextMenu = Me.ctxAddService
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

    Private Sub txtinhrmcPrice_leave(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlInHourMinCharge.PriceChanged, pnlOutHoursMC.PriceChanged, dpMCCallOutIn.PriceChanged, dpMCCallOutOut.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String
            Price = Sender.ListPrice
            If Price = 0 Then
                MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
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
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
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

                'End If
                If Not ds Is Nothing Then 'We have a scheme
                    'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                    PricingSupport.CreateCustomerScheme(ds, myClient)
                    'End If
                Else
                    PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, myClient)
                End If
            End If
            DoPanelFill()

            ChangeAudit()
        End If
    End Sub


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

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    '''' Create a scheme and add one entry 
    '''' </summary>
    '''' <param name="Discount"></param>
    '''' <param name="strType"></param>
    '''' <param name="ExistingScheme"></param>
    '''' <param name="ItemCode"></param>
    '''' <param name="DiscountType"></param>
    '''' <remarks>
    '''' </remarks>
    '''' <revisionHistory>
    '''' <revision><Author>[taylor]</Author><date> [08/06/2006]</date><Action></Action></revision>
    '''' </revisionHistory>
    '''' -----------------------------------------------------------------------------
    'Private Sub pricingsupport.CreateScheme(ByVal Discount As Double, ByVal strType As String, _
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
    '        PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
    '        CreateCustomerScheme(objNewScheme)
    '    End If



    'End Sub

    'Private Sub RemoveExistingScheme(ByVal ExistingScheme As String)
    '    If ExistingScheme.Trim.Length > 0 Then      'We need to get rid of existing schemes
    '        Dim Schemes As String() = ExistingScheme.Split(New Char() {","})
    '        Dim k As Integer
    '        For k = 0 To Schemes.Length - 1     'Could be more than one value 
    '            Dim cd As Intranet.intranet.customerns.CustomerDiscount
    '            Dim strDiscId As String = Schemes.GetValue(k)
    '            If strDiscId.Trim.Length <> 0 Then       'Check it is not a dummy string 
    '                cd = myClient.InvoiceInfo.DiscountSchemes.GetDiscountByScheme(strDiscId)
    '                If Not cd Is Nothing Then       'We've found the item remove it 

    '                    cd.RemoveScheme()                  'expire item
    '                    'remove from collection 
    '                    myClient.InvoiceInfo.DiscountSchemes.Refresh()
    '                End If
    '            End If
    '        Next
    '    End If

    'End Sub

    'Private Sub CreateCustomerScheme(ByVal Ds As MedscreenLib.Glossary.DiscountScheme)
    '    Dim CDexist As Intranet.intranet.customerns.CustomerDiscount = Client.InvoiceInfo.DiscountSchemes.GetDiscountByScheme(Ds.DiscountSchemeId)
    '    If CDexist Is Nothing Then
    '        Dim cd As New Intranet.intranet.customerns.CustomerDiscount() 'Scheme not in list 
    '        cd.CustomerID = myClient.Identity               'Initialise scheme set customer Id
    '        cd.DiscountScheme = Ds                          'Set scheme object
    '        cd.DiscountSchemeId = Ds.DiscountSchemeId       'Set scheme ID
    '        cd.Priority = Me.myClient.InvoiceInfo.DiscountSchemes.Count + 1         'Set Priority to 1 + number of schemes
    '        Dim objConAmmend As Object = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typString, "Contract Ammendment")
    '        If Not objConAmmend Is Nothing Then
    '            cd.ContractAmendment = CStr(objConAmmend)
    '        End If
    '        cd.Insert()                                     'Insert and update scheme
    '        cd.Update()
    '        Me.myClient.InvoiceInfo.DiscountSchemes.Add(cd) 'Add to customer's discount schemes
    '    Else 'if we have a discount we need to remove its end date
    '        CDexist.SetEndDateNull()
    '        CDexist.Update()

    '    End If

    'End Sub

    'Private Sub MROPrice_Leave(ByVal Sender As System.Object, ByVal e As System.EventArgs)
    '    Dim objTextBox As TextBox
    '    If TypeOf Sender Is TextBox Then
    '        objTextBox = Sender
    '        Dim newPrice As Double = objTextBox.Text
    '        Dim discount As Double
    '        Dim Price As Double
    '        Dim ExistingScheme As String = ""
    '        Price = Me.txtMROListPrice.Text
    '        discount = (Price - newPrice) * 100 / Price
    '        Me.txtMRODiscount.Text = discount.ToString("0.00")
    '        ExistingScheme = Me.txtMROScheme.Text
    '        'Check to see if killing scheme
    '        If discount = 0 And ExistingScheme.Trim.Length > 0 Then
    '            RemoveExistingScheme(ExistingScheme)
    '        Else

    '            Dim ds As MedscreenLib.Glossary.DiscountScheme
    '            ds = FindDiscountScheme(discount, "", "MED_REVIEW")
    '            If Not ds Is Nothing Then 'We have a scheme
    '                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
    '                RemoveExistingScheme(ExistingScheme)
    '                CreateCustomerScheme(ds)
    '                'End 'If
    '            Else
    '                pricingsupport.CreateScheme(discount, "MRO ", ExistingScheme, "MED_REVIEW", "TYPE")
    '            End If
    '        End If
    '        If Me.Client.InvoiceInfo.PricingMethod = Intranet.intranet.customerns.ClientInvoice.InvoicePricingType.StandardPricing Then
    '            Me.FillStdPanels()
    '        Else
    '            Me.FillPanels()
    '        End If
    '    End If
    'End Sub

    'Private Function FindDiscountScheme(ByVal Discount As Double, ByVal ItemCode As String, ByVal SampleType As String) As MedscreenLib.Glossary.DiscountScheme
    '    Dim ds As MedscreenLib.Glossary.DiscountScheme
    '    Dim From As Integer = 0
    '    ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
    '    If ds Is Nothing Then
    '        Return Nothing
    '    Else
    '        Dim strRet As String = ds.FullDescription
    '        Dim intRet As MsgBoxResult = MsgBox("Do you want to use this scheme?" & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
    '        While intRet = MsgBoxResult.No And Not ds Is Nothing
    '            From += 1
    '            ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.FindDiscountScheme(Discount, ItemCode, SampleType, From)
    '            If Not ds Is Nothing Then                   'We may kick out with no scheme
    '                strRet = ds.FullDescription
    '                intRet = MsgBox("Do you want to use this scheme?" & strRet, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
    '            End If
    '        End While
    '        Return ds
    '    End If
    'End Function

    Private Sub pnlOutHoursAddTimeCO_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlOutHoursAddTimeCO.PriceChanged, pnlINHoursAddTimeCO.PriceChanged, pnlOutHoursAddTimeRoutine.PriceChanged, pnlINHoursAddTimeRoutine.PriceChanged
        Dim discount As Double
        Dim Price As Double
        Dim ExistingScheme As String = ""
        Dim strType As String
        If PricingSupport.GetChangeDate Then

            Price = Sender.ListPrice
            If Price = 0 Then
                MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim myItemCode As String = Sender.LineItem

            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
            Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", myClient)

            End If
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                PricingSupport.CreateCustomerScheme(ds, myClient)
                'End If
            Else
                PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, myClient)
            End If


            Me.blnPriceChanged = True                   'Record that a price has changed 


            If blnClientPanelFilled Then
                DoPanelFill()
            End If
        End If

    End Sub

    Private Sub dpOutOfhoursRoutine_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles dpInHoursCancel.PriceChanged, dpCallOutCancel.PriceChanged, dpFixedSiteCancellation.PriceChanged, dpOutOfhoursRoutine.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String = " Collection Cancellation"
            Price = Sender.ListPrice
            If Price = 0 Then
                MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            Dim ds As MedscreenLib.Glossary.DiscountScheme
            Dim DiscType As String
            Dim myItemCode As String = Sender.LineItem

            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
            Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", myClient)
            End If
            If Not ds Is Nothing Then 'We have a scheme
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                PricingSupport.CreateCustomerScheme(ds, myClient)
                'End If
            Else
                PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, myClient)
            End If
            If blnClientPanelFilled Then
                DoPanelFill()
            End If
        End If
    End Sub

    'Collection charge management
    Private Sub dpRoutineCollCharge_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles dpRoutineCollCharge.PriceChanged, dpRoutineCollChargeOut.PriceChanged, dpCallOutCollCharge.PriceChanged, dpCallOutCollChargeOut.PriceChanged ', pnlCallOutCharge.PriceChanged
        If PricingSupport.GetChangeDate Then
            Dim discount As Double
            Dim Price As Double
            Dim ExistingScheme As String = ""
            Dim strType As String = " Attendance Fee"
            Price = Sender.ListPrice
            If Price = 0 Then
                MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
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
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
            Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", myClient)

            End If
            If Not ds Is Nothing Then 'We have a scheme
                'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                PricingSupport.CreateCustomerScheme(ds, myClient)
                'End If
            Else
                PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, myClient)
            End If
            If blnClientPanelFilled Then
                DoPanelFill()
            End If
        End If
    End Sub

    Private Sub pnlAO_PriceChanged(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlAO.PriceChanged, pnlCallOut.PriceChanged, pnlExternal.PriceChanged, pnlFixedSite.PriceChanged, pnlRoutine.PriceChanged ', pnlBenzene.PriceChanged


        'Dim newPrice As Double = objTextBox.Text
        Dim discount As Double
        Dim Price As Double
        Dim ExistingScheme As String = ""
        Dim strType As String
        Price = Sender.ListPrice
        discount = (Price - NewPrice) * 100 / Price
        Sender.Discount = discount.ToString("0.00")
        ExistingScheme = Sender.DiscountScheme
        'If Sender.Service = "ANALONLY" Then
        '    strType = "Analysis Only"
        'ElseIf Sender.Service = "EXTERNAL" Then
        '    strType = "External Sample"
        'ElseIf Sender.Service = "CALLOUT" Then
        '    strType = "Call-Out Sample"
        'ElseIf Sender.Service = "FIXED" Then
        '    strType = "Fixed-Site Sample"
        'ElseIf Sender.Service = "ROUTINE" Then
        '    strType = "Routine Sample"
        'End If
        strType = PricingSupport.IdentifySenderService(Sender.Service)

        'Check to see if killing scheme
        Dim ds As MedscreenLib.Glossary.DiscountScheme = Nothing
        Dim DiscType As String = ""
        Dim myItemCode As String = Sender.LineItem
        If PricingSupport.GetChangeDate Then
            Debug.WriteLine(PricingSupport.ChangeDate)
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                ExistingScheme = ""

            Else

                'do we want a scheme for this sample type or all samples
                'If MsgBox("Do you want to use this price for all non Benzene Sample types", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '    DiscType = "TYPE"
                '    myItemCode = "SAMPLE"
                '    ds = PricingSupport.FindDiscountScheme(discount, "", "SAMPLE", myClient)
                '    ExistingScheme = Me.pnlAO.Scheme & "," & Me.pnlCallOut.Scheme & "," & _
                '        Me.pnlExternal.Scheme & "," & Me.pnlFixedSite.Scheme & "," & Me.pnlRoutine.Scheme
                'Else
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(discount, myItemCode, "", myClient)

                'End If
                'If Not ds Is Nothing Then 'We have a scheme
                '    'If MsgBox(ds.Description & " scheme matches this discount use it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '    PricingSupport.RemoveExistingScheme(ExistingScheme, myClient)
                '    CreateCustomerScheme(ds)
                '    'End If
                'Else
                '    PricingSupport.CreateScheme(discount, strType, ExistingScheme, myItemCode, DiscType, myClient)
                'End If
            End If
            'check to see if we have lost the existing scheme
            If ExistingScheme.Trim.Length = 0 Then
                Dim ocoll1 As New Collection()
                ocoll1.Add(myClient.Identity)
                ocoll1.Add(Sender.LineItem)
                Dim strRetScheme As String = CConnection.PackageStringList("pckWarehouse.FindoldestDiscountScheme", ocoll1)
                If Not strRetScheme Is Nothing Then
                    ExistingScheme = strRetScheme
                End If
            End If
            PricingSupport.PriceHasChanged(discount, ExistingScheme, myItemCode, myClient, strType, ds, DiscType)

            Me.blnPriceChanged = True                   'Record that a price has changed 

            DoPanelFill()
            Sender.reDraw(False)
        End If
    End Sub


    Private Sub txtRoutineCharge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCallOutCharge.Leave
        If PricingSupport.GetChangeDate Then
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
                If Me.myClient Is Nothing Then Exit Sub

                'get option 
                objSMOption = Me.myClient.Options.Item(_Option)
                If objSMOption Is Nothing Then
                    objSMOption = myClient.Options.AddOption(myClient.Identity, _Option, System.DBNull.Value, newPrice.ToString("0.00"))
                End If
                objSMOption.Value = newPrice.ToString("0.00")
                objSMOption.Update()
            End If
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
        If Me.myClient Is Nothing Then Exit Sub

        'get option 
        objSMOption = Me.myClient.Options.Item("FXDCOLTIME")
        Me.ckTimeAllowance.Enabled = (Not Me.blnDisableAll)
        Me.txtTimeAllowance.Enabled = True
        If objSMOption Is Nothing Then
            'Me.ckTimeAllowance.Checked = False
            'Me.txtTimeAllowance.Enabled = False

            Me.lblTimeAllowance.Text = "Charges are based on time away from home"
        Else
            'Me.ckTimeAllowance.Checked = True
            Dim dblTimeAllow As Double = 0
            Try
                dblTimeAllow = CDbl(objSMOption.Value)
                Me.txtTimeAllowance.Text = dblTimeAllow.ToString("0.0")
            Catch ex As Exception
            End Try

            If dblTimeAllow = 0 Then
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
        Me.txtinhratcoPrice.Enabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.txtouthratcoPrice.Enabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.txtinhratrtPrice.Enabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))
        Me.txtouthratrtPrice.Enabled = (objSMDiscount Is Nothing And (Not Me.blnDisableAll))

    End Sub

    Private Sub FillTestListView()

        'Check to see if we have discount scheme set
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount = DiscountPresent(Me.cstNoTestChargeScheme)

        'Check enabling 
        Me.ckChargeAdditionalTests.Enabled = (Not Me.blnDisableAll)
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
        Me.lvTests.BeginUpdate()

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


    Private Function DiscountPresent(ByVal DiscountScheme As Integer) As Intranet.intranet.customerns.CustomerDiscount
        Dim objSMDiscount As Intranet.intranet.customerns.CustomerDiscount
        For Each objSMDiscount In Me.myClient.InvoiceInfo.DiscountSchemes
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
                Me.myClient.InvoiceInfo.DiscountSchemes.Remove(objSMDiscount)

            End If
        Else    'Don't charge we need to have the scheme
            If objSMDiscount Is Nothing Then
                objSMDiscount = New Intranet.intranet.customerns.CustomerDiscount()
                With objSMDiscount
                    .CustomerID = Me.myClient.Identity
                    .DiscountSchemeId = Me.cstNoTestChargeScheme
                    .Priority = Me.myClient.InvoiceInfo.DiscountSchemes.Count + 1
                    .Insert()
                    .Update()
                End With
                Me.myClient.InvoiceInfo.DiscountSchemes.Add(objSMDiscount)
            End If
        End If
        Me.FillTestListView()
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
                Me.myClient.InvoiceInfo.DiscountSchemes.Remove(objSMDiscount)

            End If
        Else    'Don't charge we need to have the scheme
            If objSMDiscount Is Nothing Then
                objSMDiscount = New Intranet.intranet.customerns.CustomerDiscount()
                With objSMDiscount
                    .CustomerID = Me.myClient.Identity
                    .DiscountSchemeId = Me.cstNoAddTimeChargeScheme
                    .Priority = Me.myClient.InvoiceInfo.DiscountSchemes.Count + 1
                    .Insert()
                    .Update()
                End With
                Me.myClient.InvoiceInfo.DiscountSchemes.Add(objSMDiscount)
            End If
        End If
        SetUpTimeAllowance()
        DoPanelFill()
        Me.Cursor = Cursors.Default
        Me.ckTimeAllowance.Enabled = (Not Me.blnDisableAll)

    End Sub
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><developer>Taylor</developer><modified>06-0ct-2011 added check on zero prices</modified></revision></revisionHistory>
    Private Sub cbEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbEditor.Leave
        Dim strText As String = Me.cbEditor.Text
        Me.Cursor = Cursors.WaitCursor
        If Not lvDItem Is Nothing Then
            If PricingSupport.GetChangeDate Then
                Dim strTest = lvDItem.LineItem.ItemCode
                Dim ds As MedscreenLib.Glossary.DiscountScheme
                Dim DiscType As String
                Dim Discount As Double
                Dim dblPrice As Double = Me.Client.InvoiceInfo.Price(lvDItem.LineItem.ItemCode)
                If CDbl(strText) = 0 Then
                    Discount = 100
                ElseIf CDbl(strText) <> dblPrice Then
                    'Deal with a zero price
                    If dblPrice = 0 Then
                        MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    Discount = (dblPrice - CDbl(strText)) / dblPrice * 100
                Else
                    PricingSupport.RemoveExistingScheme(lvDItem.DiscountScheme, myClient)
                    lvDItem.Refresh()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
                'do we want a scheme for this sample type or all samples
                DiscType = "CODE"
                ds = PricingSupport.FindDiscountScheme(Discount, strTest, "", myClient)
                If Not ds Is Nothing Then 'We have a scheme
                    PricingSupport.CreateCustomerScheme(ds, myClient)
                Else
                    PricingSupport.CreateScheme(Discount, lvDItem.LineItem.Description, "", strTest, DiscType, myClient)
                End If
                lvDItem.Refresh()
                Me.Cursor = Cursors.Default()
            End If
        End If
    End Sub

    Dim strContractAmendment As String
    Private Sub udRoutine_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udRoutine.Leave, udCallOut.Leave

        Dim udControl As DomainUpDown = sender
        Dim tag As String = udControl.Tag
        Dim ActionChar As Char = udControl.Text.Chars(0)
        'Decide on action
        If Client Is Nothing Then Exit Sub
        Dim blnSchemeneedsChanging As Boolean = False
        Dim strChanges As String = "Discounts for "
        Dim strExistingSchemes As String = ""
        Dim strNewSchemes As String = ""
        Try
            If tag = "R" Then
                Me.MCActionChar = ActionChar
                Dim oColl As New Collection()
                oColl.Add(Me.myClient.Identity)
                oColl.Add("ROUTINE")
                oColl.Add(ActionChar)
                Dim strResponse As String = CConnection.PackageStringList("lib_cctool.PricingModelCorrect", oColl)
                If strResponse = "F" Then 'We have the wrong pricing model need to get amendment code
                    Dim dlgRet As MsgBoxResult = MsgBox("We need to alter the prices to activate this pricing model ok?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If dlgRet = MsgBoxResult.Yes Then
                        Dim tmpContAmend As String = strContractAmendment
                        Dim Remembered As Boolean
                        tmpContAmend = MedscreenLib.Medscreen.GetContractAmendment(Remembered, tmpContAmend)
                        If Remembered Then strContractAmendment = tmpContAmend 'Remember the contract amendment 
                        ' we now need to setup the new model
                        Dim oCmd As New OleDb.OleDbCommand()
                        Dim intChange As Integer
                        Try
                            oCmd.Connection = CConnection.DbConnection
                            oCmd.CommandText = "Lib_CCTool.SetPricingModel"
                            oCmd.CommandType = CommandType.StoredProcedure
                            oCmd.Parameters.Add(CConnection.StringParameter("CustomerID", myClient.Identity, 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Service", "ROUTINE", 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Model", ActionChar, 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Ammendment", tmpContAmend, 100))
                            oCmd.Parameters.Add(CConnection.StringParameter("user", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                            If CConnection.ConnOpen Then
                                intChange = oCmd.ExecuteNonQuery
                            End If
                        Catch ex As Exception

                        End Try
                        VolDiscountPanel1.DrawDiscounts()
                        Me.DoPanelFill()
                    End If
                End If
                'See what we have in the way of discounts already then check that they match the action character.
                'Dim inRoutineMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.pnlInHourMinCharge.LineItem)
                'Dim outRoutineMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.pnlOutHoursMC.LineItem)
                'Dim inRoutineCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpRoutineCollCharge.LineItem)
                'Dim outRoutineCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpRoutineCollChargeOut.LineItem)
                'If sample pricing all discounts need to be set to zero 
                'which means scheme_ids 504,536,498,499
                'If minimum charge scheme ids 498,499
                'If collection charge ids 504,536
                'If Me.MCActionChar = "M" Then 'Dealing with Minimum charges
                '    If InStr(inRoutineCCDisc, "498") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Attendance fee"
                '        Dim Schemes As String() = inRoutineCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "498,"
                '    End If
                '    If InStr(outRoutineCCDisc, "499") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Attendance fee"
                '        Dim Schemes As String() = outRoutineCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "499,"
                '    End If
                '    strChanges += " need 100% discounts"
                'ElseIf Me.MCActionChar = "A" Then             'Dealing with Minimum charges
                '    If InStr(inRoutineMCDisc, "504") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Minimum charge"
                '        Dim Schemes As String() = inRoutineMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "504,"
                '    End If
                '    If InStr(outRoutineMCDisc, "536") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Minimum charge"
                '        Dim Schemes As String() = outRoutineMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "536,"
                '    End If
                '    strChanges += " need 100% discounts"
                'ElseIf Me.MCActionChar = "S" Then             'Dealing with Minimum charges
                '    If InStr(inRoutineCCDisc, "498") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Attendance fee"
                '        Dim Schemes As String() = inRoutineCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "498,"
                '    End If
                '    If InStr(outRoutineCCDisc, "499") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Attendance fee"
                '        Dim Schemes As String() = outRoutineCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "499,"
                '    End If
                '    If InStr(inRoutineMCDisc, "504") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Minimum charge"
                '        Dim Schemes As String() = inRoutineMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "504,"
                '    End If
                '    If InStr(outRoutineMCDisc, "536") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Minimum charge"
                '        Dim Schemes As String() = outRoutineMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "536,"
                '    End If
                '    strChanges += " need 100% discounts"




                'End If
            Else
                Me.COActionChar = ActionChar
                Dim oColl As New Collection()
                oColl.Add(Me.myClient.Identity)
                oColl.Add("CALLOUT")
                oColl.Add(ActionChar)
                Dim strResponse As String = CConnection.PackageStringList("lib_cctool.PricingModelCorrect", oColl)
                If strResponse = "F" Then 'We have the wrong pricing model need to get amendment code
                    Dim dlgRet As MsgBoxResult = MsgBox("We need to alter the prices to activate this pricing model ok?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If dlgRet = MsgBoxResult.Yes Then
                        Dim tmpContAmend As String = strContractAmendment
                        Dim Remembered As Boolean
                        tmpContAmend = MedscreenLib.Medscreen.GetContractAmendment(Remembered, tmpContAmend)
                        If Remembered Then strContractAmendment = tmpContAmend 'Remember the contract amendment 
                        ' we now need to setup the new model
                        Dim oCmd As New OleDb.OleDbCommand()
                        Dim intChange As Integer
                        Try
                            oCmd.Connection = CConnection.DbConnection
                            oCmd.CommandText = "Lib_CCTool.SetPricingModel"
                            oCmd.CommandType = CommandType.StoredProcedure
                            oCmd.Parameters.Add(CConnection.StringParameter("CustomerID", myClient.Identity, 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Service", "CALLOUT", 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Model", ActionChar, 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("Ammendment", tmpContAmend, 100))
                            oCmd.Parameters.Add(CConnection.StringParameter("user", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                            If CConnection.ConnOpen Then
                                intChange = oCmd.ExecuteNonQuery
                            End If
                        Catch ex As Exception

                        End Try
                        VolDiscountPanel1.DrawDiscounts()
                        Me.DoPanelFill()
                    End If
                End If

                'Dim inCallOutMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpMCCallOutIn.LineItem)
                'Dim outCallOutMCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpMCCallOutOut.LineItem)
                'Dim inCallOutCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpCallOutCollCharge.LineItem)
                'Dim outCallOutCCDisc As String = Client.InvoiceInfo.DiscountScheme(Me.dpCallOutCollChargeOut.LineItem)
                ''If sample pricing all discounts need to be set to zero 
                ''which means scheme_ids 507,508,529,503
                ''If minimum charge scheme ids 529,503
                ''If collection charge ids 507,508
                'If Me.COActionChar = "M" Then 'Dealing with Minimum charges
                '    If InStr(inCallOutCCDisc, "529") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Attendance fee"
                '        Dim Schemes As String() = inCallOutCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "529,"
                '    End If
                '    If InStr(outCallOutCCDisc, "503") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Attendance fee"
                '        Dim Schemes As String() = outCallOutCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","

                '        strNewSchemes += "503,"
                '    End If
                '    strChanges += " need 100% discounts"

                'ElseIf Me.COActionChar = "A" Then 'Dealing with Minimum charges
                '    If InStr(inCallOutMCDisc, "507") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Minimum charge"
                '        Dim Schemes As String() = inCallOutMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "507,"
                '    End If
                '    If InStr(outCallOutMCDisc, "508") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Minimum charge"
                '        Dim Schemes As String() = outCallOutMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","

                '        strNewSchemes += "508,"
                '    End If
                '    strChanges += " need 100% discounts"
                'ElseIf Me.COActionChar = "S" Then         'Dealing with Minimum charges
                '    If InStr(inCallOutCCDisc, "529") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Attendance fee"
                '        Dim Schemes As String() = inCallOutCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "529,"
                '    End If
                '    If InStr(outCallOutCCDisc, "503") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Attendance fee"
                '        Dim Schemes As String() = outCallOutCCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","

                '        strNewSchemes += "503,"
                '    End If
                '    If InStr(inCallOutMCDisc, "507") = 0 Then
                '        blnSchemeneedsChanging = True
                '        strChanges += "in hours Minimum charge"
                '        Dim Schemes As String() = inCallOutMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","
                '        strNewSchemes += "507,"
                '    End If
                '    If InStr(outCallOutMCDisc, "508") = 0 Then
                '        blnSchemeneedsChanging = True
                '        If strChanges.Length > 15 Then strChanges += ", "
                '        strChanges += "out hours Minimum charge"
                '        Dim Schemes As String() = outCallOutMCDisc.Split(New Char() {",", "-"})
                '        strExistingSchemes += Schemes.GetValue(0) & ","

                '        strNewSchemes += "508,"
                '    End If
                '    strChanges += " need 100% discounts"

                'End If
            End If
            'If blnSchemeneedsChanging Then
            '    Dim dlgRet As MsgBoxResult = MsgBox(strChanges & " create them?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            '    If dlgRet = MsgBoxResult.Yes Then
            '        PricingSupport.RemoveExistingScheme(strExistingSchemes, myClient)
            '        Dim newIds As String() = strNewSchemes.Split(New Char() {","})
            '        Dim i As Integer
            '        Dim ds As MedscreenLib.Glossary.DiscountScheme
            '        For i = 0 To newIds.Length - 1
            '            Dim newID As String = newIds.GetValue(i)
            '            If newID.Trim.Length > 0 Then
            '                ds = MedscreenLib.Glossary.Glossary.DiscountSchemes.Item(CInt(newID), 1)
            '                If Not ds Is Nothing Then PricingSupport.CreateCustomerScheme(ds, myClient) 'Check that discount scheme exists
            '            End If
            '        Next
            '        Me.FillPanels()
            '    End If
            'End If
        Catch ex As Exception
        End Try
        Debug.WriteLine("")
        SetMinChargePanels(ActionChar, tag)

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

    Private Sub FillPrices()

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
                LogError(ex, , "Fill prices - customer form ")
            Finally
                strXML = Nothing
            End Try
            DoPanelFill()
        Catch
        End Try
        LogTiming("   XML Resolved")

        'Based on the pricing model fill the panels
        VolDiscountPanel1.DrawDiscounts()

    End Sub

    Private Sub DoPanelFill()
        If Me.Client.InvoiceInfo.PricingMethod = Intranet.intranet.customerns.ClientInvoice.InvoicePricingType.StandardPricing Then
            Me.FillStdPanels()
        Else
            Me.FillPanels()
        End If

    End Sub

    Private strPriceBook As String = ""

    Private selected As Boolean = True

    ''' <developer></developer>
    ''' <summary>
    ''' Combo boxes will select on key down or key up events, this will make sure they only respond to space as a selector.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Private Sub cbPricing_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbPricing.KeyDown
        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
            selected = False
        ElseIf e.KeyCode = Keys.Space Then
            selected = True
            'e.Handled = True
        End If
    End Sub

    Private Sub cbPricing_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPricing.SelectedIndexChanged
        If Me.blnDrawing Then Exit Sub
        If Not selected Then Exit Sub 'Get around the annoying selection on key down up behaviour of combo box.
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
            Client.InvoiceInfo.Update()
            Client.InvoiceInfo.ResetTimeStamp()
            FillPrices()
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
                Me.lblTimeAllowance.Text = "Charges are based on all time on site"
            Else
                Me.lblTimeAllowance.Text = "Charges are based on time on site less customer allowance"
            End If
        ElseIf objSMOption Is Nothing Then
            Me.lblTimeAllowance.Text = "Charges are based on all collector time spent"
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
                    Me.lblTimeAllowance.Text = "Charges are based on all time on site"
                Else
                    Me.lblTimeAllowance.Text = "Charges are based on time on site less customer allowance"
                End If
            End If
        End If
        SetUpTimeAllowance()
    End Sub



    Private blnOk As Boolean = True
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim intPanel As Integer = 0
        Client.InvoiceInfo.PoRequired = Me.ckPORequired.Checked
        blnOk = True
        If Not Medscreen.IsValidSageID(Me.txtSageID.Text) Then
            Me.ErrorProvider1.SetError(Me.txtSageID, "You must have a valid sage ID for this customer")
            Me.txtSageID.Focus()
            intPanel = Me.tabTypes.TabPages.IndexOf(Me.tabInvoicing)
            blnOk = False
        Else
            Me.Client.SageID = Me.txtSageID.Text
        End If
        Client.VATRegno = Me.txtVATRegno.Text

        'Currency
        If Me.cbCurrency.SelectedIndex > -1 Then
            Client.Currency = Me.cbCurrency.SelectedValue
        End If
        Client.InvoiceInfo.CreditCard_pay = Me.ckCreditCard.Checked
        '<Added code modified 27-Apr-2007 11:41 by taylor> Credit card check
        If Not Client.InvoiceInfo.CreditCard_pay AndAlso Mid(Client.InvoiceInfo.SageID, 1, 3) = "MSC" Then
            Me.ErrorProvider1.SetError(Me.ckCreditCard, "MSCXX account needs to have card pay option set")
            Me.ckCreditCard.Checked = True
            Me.ckCreditCard.Focus()
            blnOk = False

        End If
        '</Added code modified 27-Apr-2007 11:41 by taylor> 
        Dim objPhrase As MedscreenLib.Glossary.Phrase
        If Me.cbPlatAccType.SelectedIndex > -1 Then
            objPhrase = MedscreenLib.Glossary.Glossary.Platinum.Item(Me.cbPlatAccType.SelectedIndex)
            If Not objPhrase Is Nothing Then
                Client.InvoiceInfo.PlatAccType = objPhrase.PhraseID
            End If
        End If

        If Me.cbInvFreq.SelectedIndex > -1 Then
            Client.InvoiceInfo.InvFrequency = Me.cbInvFreq.SelectedValue
        End If

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

            Me.tabTypes.SelectedIndex = intPanel
            If dlgRet = MsgBoxResult.Yes Then MsgBox("There are errors on the form," & vbCrLf & _
            "they will be marked by red exclamation marks which will explain the error")
            Exit Sub
        Else
            Client.Update()
        End If


    End Sub

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
                            If objService.EndDate > DateField.ZeroDate And objService.EndDate < Now _
                                Or objService.StartDate > DateField.ZeroDate And objService.StartDate > Now Then
                                Me.mnuctxAddService.Text = "&Reset expiry dates"
                            Else
                                Me.mnuctxAddService.Text = "&Remove Service"
                            End If
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
                        DoPanelFill()
                    End If
                End If
            End If
        End If

    End Sub

    Private strAddService As String = ""
    Dim strCurrScheme As String = "" 'Hold any scheme in use


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
        LogTiming("Draw Prices")
        DoPanelFill()
        Me.CheckMinCharges()
        LogTiming("Draw Prices Finish")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with MRO pricing
    ''' </summary>
    ''' <param name="Sender"></param>
    ''' <param name="NewPrice"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	14/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub MROPrice_Leave(ByVal Sender As MedscreenCommonGui.Panels.DiscountPanel, ByVal NewPrice As String) Handles pnlMRO.PriceChanged
        Dim objTextBox As TextBox
        Dim discount As Double
        Dim Price As Double
        Dim ExistingScheme As String = ""
        If PricingSupport.GetChangeDate Then
            Price = Sender.ListPrice
            If Price = 0 Then
                MsgBox("This would create an impossible pricing, this item has no price in the price list", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
            discount = (Price - NewPrice) * 100 / Price
            Sender.Discount = discount.ToString("0.00")
            ExistingScheme = Sender.DiscountScheme
            'Check to see if killing scheme
            If discount = 0 And ExistingScheme.Trim.Length > 0 Then
                PricingSupport.RemoveExistingScheme(ExistingScheme, Client)
            Else

                Dim ds As MedscreenLib.Glossary.DiscountScheme
                ds = PricingSupport.FindDiscountScheme(discount, "", "MED_REVIEW", myClient)
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
                DoPanelFill()
            End If
        End If
    End Sub


    'Private Sub mnuSavePanels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSavePanels.Click
    'Try
    'Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet = New Panels.DiscountPanelSet()
    'myPnlColl.Panels.Add(pnlCallOutCharge.BasePanel)
    'myPnlColl.Panels.Add(dpCallOutCollChargeOut.BasePanel)
    'myPnlColl.Panels.Add(Me.dpCallOutCollCharge.BasePanel)
    'myPnlColl.Panels.Add(dpRoutineCollChargeOut.BasePanel)
    'myPnlColl.Panels.Add(Me.dpRoutineCollCharge.BasePanel)
    ''myPnlColl.Add(pnlRCOOption.basepanel)
    'myPnlColl.Panels.Add(Me.dpMCCallOutOut.BasePanel)
    'myPnlColl.Panels.Add(Me.dpMCCallOutIn.BasePanel)
    'myPnlColl.Panels.Add(Me.pnlOutHoursMC.BasePanel)
    'myPnlColl.Panels.Add(Me.pnlInHourMinCharge.BasePanel)
    ''myPnlColl.Add(Me.pnlMinchargeheader.basepanel)
    ''myPnlColl.Add(Me.pnlSelectCharging.basepanel)
    'myPnlColl.ParentPanel = "pnlMinCharge"
    'myPnlColl.ActionHandler = "dpRoutineCollCharge_PriceChanged"
    'mypanels.Add(myPnlColl)
    'myPanels.Load()
    'Debug.WriteLine(myPanels.Count)
    'Dim myPnlColl As MedscreenCommonGui.Panels.DiscountPanelSet = myPanels.Item(0)
    'myPnlColl.BuildControls(Me)

    'Dim mypnl As Panels.DiscountPanelBase = New Panels.DiscountPanelBase()
    'mypnl.Save()
    'Catch ex As Exception
    ' End Try

    'End Sub

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
    'Private Sub cmdAddScheme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim objSchemeId As Integer
    '    objSchemeId = Me.cbVolScheme.SelectedValue
    '    Dim objScheme As New Intranet.intranet.customerns.CustomerVolumeDiscount()
    '    objScheme.CustomerID = Me.Client.Identity
    '    objScheme.DiscountSchemeId = objSchemeId
    '    objScheme.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
    '    objScheme.LinkField = "Customer.Identity"
    '    objScheme.Insert()
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Add(objScheme)
    '    Me.DrawVolumeDiscounts()

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
    Private Sub cmdDelDiscount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not myVolDiscount Is Nothing Then
            myVolDiscount.Scheme.EndDate = Today.AddDays(-1)
            myVolDiscount.Scheme.Update()
            VolDiscountPanel1.DrawVolumeDiscounts()
        End If
        If Not Me.myRoutineDiscount Is Nothing Then
            myRoutineDiscount.DiscountScheme.EndDate = Today.AddDays(-1)
            myRoutineDiscount.DiscountScheme.Update()
            VolDiscountPanel1.DrawDiscounts()
        End If
    End Sub

    'Private Sub lvVolumeDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Me.myVolDiscountEntry = Nothing
    '    myVolDiscount = Nothing
    '    Me.myRoutineDiscount = Nothing
    '    If Me.dcvVolumeDiscounts.SelectedRows.Count = 1 Then
    '        Dim objLvItem As DataGridViewRow
    '        objLvItem = Me.dcvVolumeDiscounts.SelectedRows(0)
    '        If TypeOf objLvItem Is ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader Then
    '            myVolDiscount = objLvItem
    '        ElseIf TypeOf objLvItem Is ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry Then
    '            myVolDiscountEntry = objLvItem
    '        End If
    '        'Me.myVolDiscountEntry = objLvItem
    '    End If
    'End Sub

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
    ' Private Sub DrawVolumeDiscounts()
    '    'Me.lvVolumeDiscounts.Items.Clear()
    '    Dim objCustDisnt As Intranet.intranet.customerns.CustomerVolumeDiscount
    '    dcvVolumeDiscounts.Rows.Clear()

    '    'Ensure data is refreshed before being drawn
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Clear()
    '    Me.Client.InvoiceInfo.VolumeDiscountSchemes.Load()
    '    'Go through all customer volume discounts
    '    For Each objCustDisnt In Me.Client.InvoiceInfo.VolumeDiscountSchemes
    '        Dim objlvCitem As MedscreenCommonGui.ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader
    '        objlvCitem = New ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader(objCustDisnt)

    '        'Add header to list view
    '        Me.dcvVolumeDiscounts.Rows.Add(objlvCitem)
    '        Dim objEntry As MedscreenLib.Glossary.VolumeDiscountEntry

    '        'Go through each entry and add it
    '        For Each objEntry In objCustDisnt.DiscountScheme.Entries
    '            'Dim objLvItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.lvVolumeDiscountEntry(objEntry, Me.Client)
    '            'Me.lvVolumeDiscounts.Items.Add(objLvItem)
    '            Dim objDCItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry(objEntry, Client)
    '            Me.dcvVolumeDiscounts.Rows.Add(objDCItem)
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
    '   Me.VolDiscountPanel1.CustomerID = Me.Client.Identity
    '    Me.lvRoutineDiscount.Items.Clear()
    '    Dim objCustDisnt As Intranet.intranet.customerns.CustomerDiscount

    '    'Go through each discount scheme having refreshed them first
    '    Me.Client.InvoiceInfo.DiscountSchemes.Clear()
    '    Me.Client.InvoiceInfo.DiscountSchemes.Load()
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


    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawing = False
    End Sub

    Private Sub DTPricing_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangeDate.Click
        DoPanelFill()

    End Sub

    Private Sub cmdClearDiscounts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearDiscounts.Click
        Dim oCmd As New OleDb.OleDbCommand()
        Dim oEndDate As Object = Medscreen.GetParameter(MyTypes.typDate, "Expire Date", , DateSerial(Today.Year, Today.Month, 1))
        If oEndDate Is Nothing Then Exit Sub
        Dim intChange As Integer
        Try
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "lib_cctool.SetDiscExpiryDates"
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.Parameters.Add(CConnection.StringParameter("CustomerID", Me.myClient.Identity, 10))
            oCmd.Parameters.Add(CConnection.DateParameter("onDate", CDate(oEndDate), True))
            If CConnection.ConnOpen Then
                intChange = oCmd.ExecuteNonQuery()
            End If
            VolDiscountPanel1.DrawDiscounts()
            Me.DrawPrices()
        Catch ex As Exception
            CConnection.SetConnClosed()
        End Try
    End Sub


    'Private Sub lvRoutineDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    ''' -----------------------------------------------------------------------------
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
    Private Sub lvVolumeDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPricing.DoubleClick
        If myVolDiscount Is Nothing Then Exit Sub
        Dim objFrm As New MedscreenCommonGui.frmCustomerOption()
        With objFrm
            .Type = frmCustomerOption.FormType.CustomerDiscount
            .StartDate = myVolDiscount.Scheme.Startdate
            .EndDate = myVolDiscount.Scheme.EndDate
            If .ShowDialog = DialogResult.OK Then
                myVolDiscount.Scheme.Startdate = .StartDate
                If .EndDate = DateField.ZeroDate Then
                    myVolDiscount.Scheme.SetEndDateNull()
                Else
                    myVolDiscount.Scheme.EndDate = .EndDate
                End If
                myVolDiscount.Scheme.ContractAmendment = .ContractAmmedment
                myVolDiscount.Scheme.Update()
            End If
        End With
        VolDiscountPanel1.DrawVolumeDiscounts()
    End Sub


    Private Sub lvRoutineDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.myRoutineDiscount Is Nothing Then Exit Sub
        Dim objFrm As New MedscreenCommonGui.frmCustomerOption()
        With objFrm
            .Type = frmCustomerOption.FormType.CustomerDiscount
            .StartDate = myRoutineDiscount.DiscountScheme.Startdate
            .EndDate = myRoutineDiscount.DiscountScheme.EndDate
            If .ShowDialog = DialogResult.OK Then
                myRoutineDiscount.DiscountScheme.Startdate = .StartDate
                myRoutineDiscount.DiscountScheme.EndDate = .EndDate
                myRoutineDiscount.DiscountScheme.ContractAmendment = .ContractAmmedment
                myRoutineDiscount.DiscountScheme.Update()
            End If
        End With
        VolDiscountPanel1.DrawDiscounts()
    End Sub


    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub pnlPriceInfo_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlPriceInfo.Paint

    End Sub

   
    Private Sub tabTypes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabTypes.SelectedIndexChanged
        If tabTypes.SelectedTab.Name = Me.TabDiscounts.Name Then
            Debug.Print("here")
            VolDiscountPanel1.CustomerID = myClient.Identity

        End If
    End Sub

    Private Sub cmdDeleteDiscounts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteDiscounts.Click
        Dim oCmd As New OleDb.OleDbCommand()
        Dim dlRet As MsgBoxResult = MsgBox("Do you want to delete all discounts and you understand the consequences?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2)
        If dlRet = MsgBoxResult.No Then Exit Sub
        Dim intChange As Integer
        Try
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "delete from customer_discountscheme where customer_id = ?"
            oCmd.Parameters.Add(CConnection.StringParameter("CustomerID", Me.myClient.Identity, 10))
            If CConnection.ConnOpen Then
                intChange = oCmd.ExecuteNonQuery()
            End If
            VolDiscountPanel1.DrawDiscounts()
            Me.DrawPrices()
        Catch ex As Exception
            CConnection.SetConnClosed()
        End Try
    End Sub
End Class
