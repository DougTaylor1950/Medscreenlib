'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 07:09:28+00 $
'$Log: frmCollectionForm.vb,v $
'Revision 1.0  2005-11-29 07:09:28+00  taylor
'checked in after commenting
'

Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports MedscreenLib.Constants
Imports MedscreenLib.Glossary.Glossary
Imports MedscreenCommonGui.ListViewItems
Imports Intranet.intranet.Support
Imports Intranet.intranet.jobs.CMJob
Imports Intranet.intranet.jobs
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCollectionForm
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to enter and maintain routine collections
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCollectionForm
    Inherits System.Windows.Forms.Form
#Region "Declarations"


    Private iniFile As New MedscreenLib.IniFile.IniFiles()
    Private strFileDir As String
    Private objVessels As Intranet.intranet.customerns.CVesselCollection
    Private objPort As Intranet.intranet.jobs.Port
    Private CustSites As Intranet.intranet.Address.CustomerSites
    Private blnFormLoaded As Boolean = False
    Private VesselFirstDate As Date = Now
    Private dblCost As Double
    Private ProGmanagers As MedscreenLib.Glossary.RoleAssignments = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserRoleHolders("PROGRAMMEMANAGER")
    Private cr As System.Globalization.CultureInfo
    Private CustomerCurrency As String = ""
    Private blnFormSet As Boolean = False
    Private blnCustSiteSelected As Boolean = False
    Dim objSibColl As Intranet.intranet.customerns.SMClientCollection

    Private myCollection As Intranet.intranet.jobs.CMJob
    Private myClient As Intranet.intranet.customerns.Client

    Private objSampleTemplates As Glossary.PhraseCollection
    Private objAnalalyticalLabs As Glossary.PhraseCollection
    Friend WithEvents cmdEditDoc As System.Windows.Forms.Button
    Friend WithEvents lblTemplate As System.Windows.Forms.Label
    Friend WithEvents lvColumnGroups As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents udEditor As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdCreateDocument As System.Windows.Forms.Button
    Friend WithEvents txtEditor As System.Windows.Forms.TextBox
    Friend WithEvents lvRandom As System.Windows.Forms.DataGridView
    Friend WithEvents cmdSendRandom As System.Windows.Forms.Button
    'Labs that can perform the analysis
    Friend WithEvents Process1 As System.Diagnostics.Process

#End Region


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Dim strPath As String
        Dim fi As System.IO.FileInfo

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Medscreen.LogTiming("Form Open")

        Me.cbPManager.DataSource = ProGmanagers
        Me.cbPManager.ValueMember = "User"
        Me.cbPManager.DisplayMember = "User"

        strPath = Application.ExecutablePath
        fi = New System.IO.FileInfo(strPath)
        strPath = fi.DirectoryName

        Me.cBLocality.SelectedIndex = -1

        Medscreen.LogTiming("Fill Lists")

        Try
            MedscreenLib.Medscreen.FillList(Me.cBLocality, MedscreenCommonGui.My.Settings.List_cmbUKOS)
            MedscreenLib.Medscreen.FillList(Me.cbBlnProc, MedscreenCommonGui.My.Settings.List_cmbProcedures)
            MedscreenLib.Medscreen.FillList(Me.cbPMAnaged, MedscreenCommonGui.My.Settings.List_cmbManaged)
            MedscreenLib.Medscreen.FillList(Me.cbAnnouced, MedscreenCommonGui.My.Settings.List_cmbAnnounced)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim strCurrencyList As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CurrencyList")

        Dim objCurrencyList As MedscreenLib.Glossary.PhraseCollection = New Glossary.PhraseCollection(strCurrencyList, Glossary.PhraseCollection.BuildBy.Query)
        Me.cbCurrency.DataSource = objCurrencyList
        Me.cbCurrency.DisplayMember = "PhraseText"
        Me.cbCurrency.ValueMember = "PhraseID"

        'MedscreenLib.Medscreen.FillList(Me.cbCurrency, strPath & _
        ' "\collman.ini", "List_cmbCurrency")


        Medscreen.LogTiming("Phrases")

        'Set up reason for test 
        Me.cbRandom.DataSource = MedscreenLib.Glossary.Glossary.ReasonForTestList
        Me.cbRandom.DisplayMember = "PhraseText"
        Me.cbRandom.ValueMember = "PhraseID"

        'Set up Confirmation Method 
        Me.cbConfMethod.DataSource = MedscreenLib.Glossary.Glossary.ConfMethodList
        Me.cbConfMethod.DisplayMember = "PhraseText"
        Me.cbConfMethod.ValueMember = "PhraseID"

        'Set up test type 
        Me.cbTestType.DataSource = MedscreenLib.Glossary.Glossary.TestTypeList
        Me.cbTestType.DisplayMember = "PhraseText"
        Me.cbTestType.ValueMember = "PhraseID"

        'Set up confirmation Type
        Me.cbConfType.DataSource = MedscreenLib.Glossary.Glossary.ConfTypeList
        Me.cbConfType.DisplayMember = "PhraseText"
        Me.cbConfType.ValueMember = "PhraseID"

        'Set up confirmation Type

        Me.cbSampleTemplate.DataSource = objSampleTemplates
        Me.cbSampleTemplate.DisplayMember = "PhraseText"
        Me.cbSampleTemplate.ValueMember = "PhraseID"

        Dim strSquery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("LaboratoryQuery")
        'Dim strSQuery As String = "select identity,description from instrument s " & _
        '"where s.removeflag='F' and insttype_id = 'LABORATORY'"
        Me.objAnalalyticalLabs = New Glossary.PhraseCollection(strSquery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)


        Me.cbAnalyticalLab.DataSource = objAnalalyticalLabs
        Me.cbAnalyticalLab.DisplayMember = "PhraseText"
        Me.cbAnalyticalLab.ValueMember = "PhraseID"


        iniFile.FileName = strPath & "\collman.ini"
        strFileDir = iniFile.ReadString("General", "AddDocPath", "")

        Me.NoDonorsActual.SelectedIndex = 0
        Me.cBLocality.Items.Clear()
        Me.cBLocality.Items.AddRange(New Object() {"UK", "Overseas"})

        'Add any initialization after the InitializeComponent() call
        Me.TabControl1.SelectedIndex = 0
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
    Friend WithEvents pnlCollection As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdSubmit As System.Windows.Forms.Button
    Friend WithEvents cmdClearForm As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtilities As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuViewInvoice2 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuEdAddDoc2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdAddDoc As System.Windows.Forms.MenuItem
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents btnImage As System.Windows.Forms.ImageList
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents mnuCustDefaults As System.Windows.Forms.MenuItem
    Friend WithEvents NWInfo As VbPowerPack.NotificationWindow
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents cmdPortCosts As System.Windows.Forms.Button
    Friend WithEvents cmdEditVessel As System.Windows.Forms.Button
    Friend WithEvents NoDonorsActual As System.Windows.Forms.DomainUpDown
    Friend WithEvents cmdSiteAddress As System.Windows.Forms.Button
    Friend WithEvents cmdViewPanel As System.Windows.Forms.Button
    Friend WithEvents cbTestType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdCustomerComment As System.Windows.Forms.Button
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents cmdAddVessel As System.Windows.Forms.Button
    Friend WithEvents lbllkAddDoc As System.Windows.Forms.LinkLabel
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents txtAddDoc As System.Windows.Forms.TextBox
    Friend WithEvents lblAddDoc As System.Windows.Forms.Label
    Friend WithEvents txtAddCosts As System.Windows.Forms.TextBox
    Friend WithEvents lblAddCosts As System.Windows.Forms.Label
    Friend WithEvents txtAgentphone As System.Windows.Forms.TextBox
    Friend WithEvents lblagentPHOne As System.Windows.Forms.Label
    Friend WithEvents txtAgentName As System.Windows.Forms.TextBox
    Friend WithEvents lblAgentNAme As System.Windows.Forms.Label
    Friend WithEvents cbSampleTemplate As System.Windows.Forms.ComboBox
    Friend WithEvents cbRandom As System.Windows.Forms.ComboBox
    Friend WithEvents cBLocality As System.Windows.Forms.ComboBox
    Friend WithEvents txtWhototest As System.Windows.Forms.TextBox
    Friend WithEvents updnNoDonors As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblTestPanel As System.Windows.Forms.Label
    Friend WithEvents lblCollType As System.Windows.Forms.Label
    Friend WithEvents lblWhototest As System.Windows.Forms.Label
    Friend WithEvents lblNDonors As System.Windows.Forms.Label
    Friend WithEvents cbAnnouced As System.Windows.Forms.ComboBox
    Friend WithEvents lblannouced As System.Windows.Forms.Label
    Friend WithEvents DTCollDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblCollDate As System.Windows.Forms.Label
    Friend WithEvents cbBlnProc As System.Windows.Forms.ComboBox
    Friend WithEvents lblProcedures As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblVessel As System.Windows.Forms.Label
    Friend WithEvents cbSiteVessel As System.Windows.Forms.ComboBox
    Friend WithEvents lblSite As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents cmdEditSite As System.Windows.Forms.Button
    Friend WithEvents cmdNextDueVessel As System.Windows.Forms.Button
    Friend WithEvents cmdNewSite As System.Windows.Forms.Button
    Friend WithEvents cbCustSites As System.Windows.Forms.ComboBox
    Friend WithEvents cbVessel As System.Windows.Forms.ComboBox
    Friend WithEvents gbCollection As System.Windows.Forms.GroupBox
    Friend WithEvents GBConfirm As System.Windows.Forms.GroupBox
    Friend WithEvents lblInvoice As System.Windows.Forms.LinkLabel
    Friend WithEvents txtConfAdddress As System.Windows.Forms.TextBox
    Friend WithEvents lblConfAddress As System.Windows.Forms.Label
    Friend WithEvents TxtConfrecip As System.Windows.Forms.TextBox
    Friend WithEvents cbConfMethod As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfMethod As System.Windows.Forms.Label
    Friend WithEvents cbConfType As System.Windows.Forms.ComboBox
    Friend WithEvents cbPManager As System.Windows.Forms.ComboBox
    Friend WithEvents lblBy As System.Windows.Forms.Label
    Friend WithEvents cbPMAnaged As System.Windows.Forms.ComboBox
    Friend WithEvents lblPm As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents mnuPricing As System.Windows.Forms.MenuItem
    Friend WithEvents mnuJobCost As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCustomerPricing As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPort As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewPort As System.Windows.Forms.MenuItem
    Friend WithEvents txtCustComment As System.Windows.Forms.TextBox
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents cmdFindPort As System.Windows.Forms.Button
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents cmdAssign As System.Windows.Forms.Button
    Friend WithEvents pnlOfficer As System.Windows.Forms.Panel
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents colID As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColForeName As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSurname As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents LVRedirect As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ctxOfficer As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Friend WithEvents ctxOfficerEmail As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuSendEmail As System.Windows.Forms.MenuItem
    Friend WithEvents multisendImg As System.Windows.Forms.ImageList
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents cmdGenerateRandom As System.Windows.Forms.Button
    Friend WithEvents txtMatrix As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbAnalyticalLab As System.Windows.Forms.ComboBox
    Friend WithEvents cmdLabAddress As System.Windows.Forms.Button
    Friend WithEvents cmdConvert As System.Windows.Forms.Button
    Friend WithEvents ckBenzene As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TabSubControls As System.Windows.Forms.TabControl
    Friend WithEvents tpDocuments As System.Windows.Forms.TabPage
    Friend WithEvents pnlDocuments As System.Windows.Forms.Panel
    Friend WithEvents cmdRemoveDocument As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdAddDoc As System.Windows.Forms.Button
    Friend WithEvents lvDocuments As System.Windows.Forms.ListView
    Friend WithEvents tpInfo As System.Windows.Forms.TabPage
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtAddress As System.Windows.Forms.RichTextBox
    Friend WithEvents lblLocality As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PnlConfirmations As System.Windows.Forms.Panel
    Friend WithEvents rtbContacts As System.Windows.Forms.RichTextBox
    Friend WithEvents txtPO As System.Windows.Forms.TextBox
    Friend WithEvents lblPONumber As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCollectionForm))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdAssign = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdClearForm = New System.Windows.Forms.Button
        Me.cmdSubmit = New System.Windows.Forms.Button
        Me.pnlCollection = New System.Windows.Forms.Panel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.gbCollection = New System.Windows.Forms.GroupBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.mnuViewInvoice2 = New System.Windows.Forms.MenuItem
        Me.mnuEdAddDoc2 = New System.Windows.Forms.MenuItem
        Me.txtPO = New System.Windows.Forms.TextBox
        Me.lblPONumber = New System.Windows.Forms.Label
        Me.TabSubControls = New System.Windows.Forms.TabControl
        Me.tpInfo = New System.Windows.Forms.TabPage
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtAddress = New System.Windows.Forms.RichTextBox
        Me.tpDocuments = New System.Windows.Forms.TabPage
        Me.lblLocality = New System.Windows.Forms.Label
        Me.pnlDocuments = New System.Windows.Forms.Panel
        Me.cmdRemoveDocument = New System.Windows.Forms.Button
        Me.cmdEdit = New System.Windows.Forms.Button
        Me.cmdAddDoc = New System.Windows.Forms.Button
        Me.lvDocuments = New System.Windows.Forms.ListView
        Me.ckBenzene = New System.Windows.Forms.CheckBox
        Me.cmdConvert = New System.Windows.Forms.Button
        Me.cmdLabAddress = New System.Windows.Forms.Button
        Me.cbAnalyticalLab = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtMatrix = New System.Windows.Forms.TextBox
        Me.txtCustomer = New System.Windows.Forms.TextBox
        Me.cmdFindPort = New System.Windows.Forms.Button
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.txtCustComment = New System.Windows.Forms.TextBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.cmdPortCosts = New System.Windows.Forms.Button
        Me.cmdEditVessel = New System.Windows.Forms.Button
        Me.NoDonorsActual = New System.Windows.Forms.DomainUpDown
        Me.cmdSiteAddress = New System.Windows.Forms.Button
        Me.cmdViewPanel = New System.Windows.Forms.Button
        Me.cbTestType = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdCustomerComment = New System.Windows.Forms.Button
        Me.lblInfo = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.cmdAddVessel = New System.Windows.Forms.Button
        Me.lbllkAddDoc = New System.Windows.Forms.LinkLabel
        Me.cbCurrency = New System.Windows.Forms.ComboBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.txtAddDoc = New System.Windows.Forms.TextBox
        Me.lblAddDoc = New System.Windows.Forms.Label
        Me.txtAddCosts = New System.Windows.Forms.TextBox
        Me.lblAddCosts = New System.Windows.Forms.Label
        Me.txtAgentphone = New System.Windows.Forms.TextBox
        Me.lblagentPHOne = New System.Windows.Forms.Label
        Me.txtAgentName = New System.Windows.Forms.TextBox
        Me.lblAgentNAme = New System.Windows.Forms.Label
        Me.cbSampleTemplate = New System.Windows.Forms.ComboBox
        Me.cbRandom = New System.Windows.Forms.ComboBox
        Me.cBLocality = New System.Windows.Forms.ComboBox
        Me.txtWhototest = New System.Windows.Forms.TextBox
        Me.updnNoDonors = New System.Windows.Forms.NumericUpDown
        Me.lblTestPanel = New System.Windows.Forms.Label
        Me.lblCollType = New System.Windows.Forms.Label
        Me.lblWhototest = New System.Windows.Forms.Label
        Me.lblNDonors = New System.Windows.Forms.Label
        Me.cbAnnouced = New System.Windows.Forms.ComboBox
        Me.lblannouced = New System.Windows.Forms.Label
        Me.DTCollDate = New System.Windows.Forms.DateTimePicker
        Me.lblCollDate = New System.Windows.Forms.Label
        Me.cbBlnProc = New System.Windows.Forms.ComboBox
        Me.lblProcedures = New System.Windows.Forms.Label
        Me.lblLocation = New System.Windows.Forms.Label
        Me.lblVessel = New System.Windows.Forms.Label
        Me.cbSiteVessel = New System.Windows.Forms.ComboBox
        Me.lblSite = New System.Windows.Forms.Label
        Me.lblID = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.cmdEditSite = New System.Windows.Forms.Button
        Me.cmdNextDueVessel = New System.Windows.Forms.Button
        Me.btnImage = New System.Windows.Forms.ImageList(Me.components)
        Me.cmdNewSite = New System.Windows.Forms.Button
        Me.cbCustSites = New System.Windows.Forms.ComboBox
        Me.cbVessel = New System.Windows.Forms.ComboBox
        Me.GBConfirm = New System.Windows.Forms.GroupBox
        Me.PnlConfirmations = New System.Windows.Forms.Panel
        Me.rtbContacts = New System.Windows.Forms.RichTextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.lblInvoice = New System.Windows.Forms.LinkLabel
        Me.txtConfAdddress = New System.Windows.Forms.TextBox
        Me.lblConfAddress = New System.Windows.Forms.Label
        Me.TxtConfrecip = New System.Windows.Forms.TextBox
        Me.cbConfMethod = New System.Windows.Forms.ComboBox
        Me.lblConfMethod = New System.Windows.Forms.Label
        Me.cbConfType = New System.Windows.Forms.ComboBox
        Me.cbPManager = New System.Windows.Forms.ComboBox
        Me.lblBy = New System.Windows.Forms.Label
        Me.cbPMAnaged = New System.Windows.Forms.ComboBox
        Me.lblPm = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.pnlOfficer = New System.Windows.Forms.Panel
        Me.LVRedirect = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ctxOfficerEmail = New System.Windows.Forms.ContextMenu
        Me.mnuSendEmail = New System.Windows.Forms.MenuItem
        Me.multisendImg = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.colID = New System.Windows.Forms.ColumnHeader
        Me.ColForeName = New System.Windows.Forms.ColumnHeader
        Me.colSurname = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ctxOfficer = New System.Windows.Forms.ContextMenu
        Me.mnuAddOfficer = New System.Windows.Forms.MenuItem
        Me.mnuRemoveOfficer = New System.Windows.Forms.MenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.cmdSendRandom = New System.Windows.Forms.Button
        Me.lvRandom = New System.Windows.Forms.DataGridView
        Me.txtEditor = New System.Windows.Forms.TextBox
        Me.cmdCreateDocument = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.udEditor = New System.Windows.Forms.NumericUpDown
        Me.lvColumnGroups = New ListViewEx.ListViewEx
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.lblTemplate = New System.Windows.Forms.Label
        Me.cmdEditDoc = New System.Windows.Forms.Button
        Me.cmdGenerateRandom = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuUtilities = New System.Windows.Forms.MenuItem
        Me.mnuViewInvoice = New System.Windows.Forms.MenuItem
        Me.mnuEdAddDoc = New System.Windows.Forms.MenuItem
        Me.mnuCustDefaults = New System.Windows.Forms.MenuItem
        Me.mnuPricing = New System.Windows.Forms.MenuItem
        Me.mnuJobCost = New System.Windows.Forms.MenuItem
        Me.mnuCustomerPricing = New System.Windows.Forms.MenuItem
        Me.mnuPort = New System.Windows.Forms.MenuItem
        Me.mnuNewPort = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.NWInfo = New VbPowerPack.NotificationWindow(Me.components)
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel1.SuspendLayout()
        Me.pnlCollection.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.gbCollection.SuspendLayout()
        Me.TabSubControls.SuspendLayout()
        Me.tpInfo.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.tpDocuments.SuspendLayout()
        Me.pnlDocuments.SuspendLayout()
        CType(Me.updnNoDonors, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBConfirm.SuspendLayout()
        Me.PnlConfirmations.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.pnlOfficer.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.lvRandom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdAssign)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdClearForm)
        Me.Panel1.Controls.Add(Me.cmdSubmit)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 565)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(824, 48)
        Me.Panel1.TabIndex = 0
        '
        'cmdAssign
        '
        Me.cmdAssign.Image = CType(resources.GetObject("cmdAssign.Image"), System.Drawing.Image)
        Me.cmdAssign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAssign.Location = New System.Drawing.Point(184, 16)
        Me.cmdAssign.Name = "cmdAssign"
        Me.cmdAssign.Size = New System.Drawing.Size(118, 23)
        Me.cmdAssign.TabIndex = 4
        Me.cmdAssign.Text = "&Assign Officer"
        Me.cmdAssign.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(686, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(118, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdClearForm
        '
        Me.cmdClearForm.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClearForm.Image = CType(resources.GetObject("cmdClearForm.Image"), System.Drawing.Image)
        Me.cmdClearForm.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClearForm.Location = New System.Drawing.Point(548, 16)
        Me.cmdClearForm.Name = "cmdClearForm"
        Me.cmdClearForm.Size = New System.Drawing.Size(118, 23)
        Me.cmdClearForm.TabIndex = 2
        Me.cmdClearForm.Text = "Clear &Form"
        Me.cmdClearForm.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdSubmit
        '
        Me.cmdSubmit.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSubmit.Image = CType(resources.GetObject("cmdSubmit.Image"), System.Drawing.Image)
        Me.cmdSubmit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSubmit.Location = New System.Drawing.Point(48, 16)
        Me.cmdSubmit.Name = "cmdSubmit"
        Me.cmdSubmit.Size = New System.Drawing.Size(118, 23)
        Me.cmdSubmit.TabIndex = 0
        Me.cmdSubmit.Text = "&Submit Collection"
        Me.cmdSubmit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCollection
        '
        Me.pnlCollection.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCollection.Controls.Add(Me.TabControl1)
        Me.pnlCollection.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCollection.Location = New System.Drawing.Point(0, 0)
        Me.pnlCollection.Name = "pnlCollection"
        Me.pnlCollection.Size = New System.Drawing.Size(824, 565)
        Me.pnlCollection.TabIndex = 1
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(820, 561)
        Me.TabControl1.TabIndex = 85
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.gbCollection)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(812, 535)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Collection"
        '
        'gbCollection
        '
        Me.gbCollection.ContextMenu = Me.ContextMenu1
        Me.gbCollection.Controls.Add(Me.txtPO)
        Me.gbCollection.Controls.Add(Me.lblPONumber)
        Me.gbCollection.Controls.Add(Me.TabSubControls)
        Me.gbCollection.Controls.Add(Me.ckBenzene)
        Me.gbCollection.Controls.Add(Me.cmdConvert)
        Me.gbCollection.Controls.Add(Me.cmdLabAddress)
        Me.gbCollection.Controls.Add(Me.cbAnalyticalLab)
        Me.gbCollection.Controls.Add(Me.Label8)
        Me.gbCollection.Controls.Add(Me.txtMatrix)
        Me.gbCollection.Controls.Add(Me.txtCustomer)
        Me.gbCollection.Controls.Add(Me.cmdFindPort)
        Me.gbCollection.Controls.Add(Me.txtLocation)
        Me.gbCollection.Controls.Add(Me.txtCustComment)
        Me.gbCollection.Controls.Add(Me.LinkLabel1)
        Me.gbCollection.Controls.Add(Me.cmdPortCosts)
        Me.gbCollection.Controls.Add(Me.cmdEditVessel)
        Me.gbCollection.Controls.Add(Me.NoDonorsActual)
        Me.gbCollection.Controls.Add(Me.cmdSiteAddress)
        Me.gbCollection.Controls.Add(Me.cmdViewPanel)
        Me.gbCollection.Controls.Add(Me.cbTestType)
        Me.gbCollection.Controls.Add(Me.Label2)
        Me.gbCollection.Controls.Add(Me.cmdCustomerComment)
        Me.gbCollection.Controls.Add(Me.lblInfo)
        Me.gbCollection.Controls.Add(Me.txtComments)
        Me.gbCollection.Controls.Add(Me.lblComments)
        Me.gbCollection.Controls.Add(Me.cmdAddVessel)
        Me.gbCollection.Controls.Add(Me.lbllkAddDoc)
        Me.gbCollection.Controls.Add(Me.cbCurrency)
        Me.gbCollection.Controls.Add(Me.cmdBrowse)
        Me.gbCollection.Controls.Add(Me.txtAddDoc)
        Me.gbCollection.Controls.Add(Me.lblAddDoc)
        Me.gbCollection.Controls.Add(Me.txtAddCosts)
        Me.gbCollection.Controls.Add(Me.lblAddCosts)
        Me.gbCollection.Controls.Add(Me.txtAgentphone)
        Me.gbCollection.Controls.Add(Me.lblagentPHOne)
        Me.gbCollection.Controls.Add(Me.txtAgentName)
        Me.gbCollection.Controls.Add(Me.lblAgentNAme)
        Me.gbCollection.Controls.Add(Me.cbSampleTemplate)
        Me.gbCollection.Controls.Add(Me.cbRandom)
        Me.gbCollection.Controls.Add(Me.cBLocality)
        Me.gbCollection.Controls.Add(Me.txtWhototest)
        Me.gbCollection.Controls.Add(Me.updnNoDonors)
        Me.gbCollection.Controls.Add(Me.lblTestPanel)
        Me.gbCollection.Controls.Add(Me.lblCollType)
        Me.gbCollection.Controls.Add(Me.lblWhototest)
        Me.gbCollection.Controls.Add(Me.lblNDonors)
        Me.gbCollection.Controls.Add(Me.cbAnnouced)
        Me.gbCollection.Controls.Add(Me.lblannouced)
        Me.gbCollection.Controls.Add(Me.DTCollDate)
        Me.gbCollection.Controls.Add(Me.lblCollDate)
        Me.gbCollection.Controls.Add(Me.cbBlnProc)
        Me.gbCollection.Controls.Add(Me.lblProcedures)
        Me.gbCollection.Controls.Add(Me.lblLocation)
        Me.gbCollection.Controls.Add(Me.lblVessel)
        Me.gbCollection.Controls.Add(Me.cbSiteVessel)
        Me.gbCollection.Controls.Add(Me.lblSite)
        Me.gbCollection.Controls.Add(Me.lblID)
        Me.gbCollection.Controls.Add(Me.Label1)
        Me.gbCollection.Controls.Add(Me.NumericUpDown1)
        Me.gbCollection.Controls.Add(Me.cmdEditSite)
        Me.gbCollection.Controls.Add(Me.cmdNextDueVessel)
        Me.gbCollection.Controls.Add(Me.cmdNewSite)
        Me.gbCollection.Controls.Add(Me.cbCustSites)
        Me.gbCollection.Controls.Add(Me.cbVessel)
        Me.gbCollection.Controls.Add(Me.GBConfirm)
        Me.gbCollection.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbCollection.Location = New System.Drawing.Point(0, 0)
        Me.gbCollection.Name = "gbCollection"
        Me.gbCollection.Size = New System.Drawing.Size(812, 535)
        Me.gbCollection.TabIndex = 0
        Me.gbCollection.TabStop = False
        Me.gbCollection.Text = "Routine Collection"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewInvoice2, Me.mnuEdAddDoc2})
        '
        'mnuViewInvoice2
        '
        Me.mnuViewInvoice2.Index = 0
        Me.mnuViewInvoice2.Text = "View &Invoice"
        '
        'mnuEdAddDoc2
        '
        Me.mnuEdAddDoc2.Index = 1
        Me.mnuEdAddDoc2.Text = "&Edit Additional Doc"
        '
        'txtPO
        '
        Me.txtPO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPO.Location = New System.Drawing.Point(568, 312)
        Me.txtPO.Name = "txtPO"
        Me.txtPO.Size = New System.Drawing.Size(136, 20)
        Me.txtPO.TabIndex = 101
        '
        'lblPONumber
        '
        Me.lblPONumber.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblPONumber.Location = New System.Drawing.Point(456, 312)
        Me.lblPONumber.Name = "lblPONumber"
        Me.lblPONumber.Size = New System.Drawing.Size(100, 23)
        Me.lblPONumber.TabIndex = 100
        Me.lblPONumber.Text = "Purchase Order"
        Me.lblPONumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabSubControls
        '
        Me.TabSubControls.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabSubControls.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabSubControls.Controls.Add(Me.tpInfo)
        Me.TabSubControls.Controls.Add(Me.tpDocuments)
        Me.TabSubControls.Location = New System.Drawing.Point(456, 344)
        Me.TabSubControls.Name = "TabSubControls"
        Me.TabSubControls.SelectedIndex = 0
        Me.TabSubControls.Size = New System.Drawing.Size(344, 192)
        Me.TabSubControls.TabIndex = 99
        '
        'tpInfo
        '
        Me.tpInfo.Controls.Add(Me.Panel2)
        Me.tpInfo.Location = New System.Drawing.Point(4, 4)
        Me.tpInfo.Name = "tpInfo"
        Me.tpInfo.Size = New System.Drawing.Size(336, 166)
        Me.tpInfo.TabIndex = 1
        Me.tpInfo.Text = "Info"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtAddress)
        Me.Panel2.Location = New System.Drawing.Point(16, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(312, 160)
        Me.Panel2.TabIndex = 77
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Collection Locality"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Visible = False
        '
        'txtAddress
        '
        Me.txtAddress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAddress.Location = New System.Drawing.Point(8, 24)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(296, 130)
        Me.txtAddress.TabIndex = 0
        Me.txtAddress.Text = ""
        '
        'tpDocuments
        '
        Me.tpDocuments.Controls.Add(Me.lblLocality)
        Me.tpDocuments.Controls.Add(Me.pnlDocuments)
        Me.tpDocuments.Location = New System.Drawing.Point(4, 4)
        Me.tpDocuments.Name = "tpDocuments"
        Me.tpDocuments.Size = New System.Drawing.Size(336, 166)
        Me.tpDocuments.TabIndex = 0
        Me.tpDocuments.Text = "Documents"
        '
        'lblLocality
        '
        Me.lblLocality.Location = New System.Drawing.Point(16, 0)
        Me.lblLocality.Name = "lblLocality"
        Me.lblLocality.Size = New System.Drawing.Size(144, 23)
        Me.lblLocality.TabIndex = 100
        Me.lblLocality.Text = "Documents to Officer"
        Me.lblLocality.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLocality.Visible = False
        '
        'pnlDocuments
        '
        Me.pnlDocuments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlDocuments.Controls.Add(Me.cmdRemoveDocument)
        Me.pnlDocuments.Controls.Add(Me.cmdEdit)
        Me.pnlDocuments.Controls.Add(Me.cmdAddDoc)
        Me.pnlDocuments.Controls.Add(Me.lvDocuments)
        Me.pnlDocuments.Location = New System.Drawing.Point(8, 32)
        Me.pnlDocuments.Name = "pnlDocuments"
        Me.pnlDocuments.Size = New System.Drawing.Size(322, 131)
        Me.pnlDocuments.TabIndex = 99
        '
        'cmdRemoveDocument
        '
        Me.cmdRemoveDocument.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRemoveDocument.Image = Global.MedscreenCommonGui.AdcResource.Cross16
        Me.cmdRemoveDocument.Location = New System.Drawing.Point(282, 104)
        Me.cmdRemoveDocument.Name = "cmdRemoveDocument"
        Me.cmdRemoveDocument.Size = New System.Drawing.Size(24, 23)
        Me.cmdRemoveDocument.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmdRemoveDocument, "Remove Collection Document")
        '
        'cmdEdit
        '
        Me.cmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdit.Image = CType(resources.GetObject("cmdEdit.Image"), System.Drawing.Image)
        Me.cmdEdit.Location = New System.Drawing.Point(282, 64)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(24, 23)
        Me.cmdEdit.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmdEdit, "Edit Collection Document")
        '
        'cmdAddDoc
        '
        Me.cmdAddDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAddDoc.Image = CType(resources.GetObject("cmdAddDoc.Image"), System.Drawing.Image)
        Me.cmdAddDoc.Location = New System.Drawing.Point(282, 24)
        Me.cmdAddDoc.Name = "cmdAddDoc"
        Me.cmdAddDoc.Size = New System.Drawing.Size(24, 23)
        Me.cmdAddDoc.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmdAddDoc, "Add Collection Document")
        '
        'lvDocuments
        '
        Me.lvDocuments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvDocuments.Location = New System.Drawing.Point(0, 0)
        Me.lvDocuments.Name = "lvDocuments"
        Me.lvDocuments.Size = New System.Drawing.Size(274, 127)
        Me.lvDocuments.TabIndex = 0
        Me.lvDocuments.UseCompatibleStateImageBehavior = False
        Me.lvDocuments.View = System.Windows.Forms.View.List
        '
        'ckBenzene
        '
        Me.ckBenzene.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckBenzene.Location = New System.Drawing.Point(128, 600)
        Me.ckBenzene.Name = "ckBenzene"
        Me.ckBenzene.Size = New System.Drawing.Size(176, 24)
        Me.ckBenzene.TabIndex = 97
        Me.ckBenzene.Text = "Test for Benzene and DoA"
        '
        'cmdConvert
        '
        Me.cmdConvert.Image = CType(resources.GetObject("cmdConvert.Image"), System.Drawing.Image)
        Me.cmdConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdConvert.Location = New System.Drawing.Point(272, 176)
        Me.cmdConvert.Name = "cmdConvert"
        Me.cmdConvert.Size = New System.Drawing.Size(80, 23)
        Me.cmdConvert.TabIndex = 96
        Me.cmdConvert.Text = "Convert"
        Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdLabAddress
        '
        Me.cmdLabAddress.Image = CType(resources.GetObject("cmdLabAddress.Image"), System.Drawing.Image)
        Me.cmdLabAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdLabAddress.Location = New System.Drawing.Point(248, 32)
        Me.cmdLabAddress.Name = "cmdLabAddress"
        Me.cmdLabAddress.Size = New System.Drawing.Size(24, 23)
        Me.cmdLabAddress.TabIndex = 95
        Me.cmdLabAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdLabAddress, "See details of the Laboratory address")
        '
        'cbAnalyticalLab
        '
        Me.cbAnalyticalLab.Location = New System.Drawing.Point(8, 32)
        Me.cbAnalyticalLab.Name = "cbAnalyticalLab"
        Me.cbAnalyticalLab.Size = New System.Drawing.Size(240, 21)
        Me.cbAnalyticalLab.TabIndex = 87
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 23)
        Me.Label8.TabIndex = 86
        Me.Label8.Text = "Analytical Lab"
        '
        'txtMatrix
        '
        Me.txtMatrix.Location = New System.Drawing.Point(128, 504)
        Me.txtMatrix.Multiline = True
        Me.txtMatrix.Name = "txtMatrix"
        Me.txtMatrix.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMatrix.Size = New System.Drawing.Size(320, 56)
        Me.txtMatrix.TabIndex = 85
        '
        'txtCustomer
        '
        Me.txtCustomer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomer.Location = New System.Drawing.Point(128, 64)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(672, 20)
        Me.txtCustomer.TabIndex = 84
        '
        'cmdFindPort
        '
        Me.cmdFindPort.Image = CType(resources.GetObject("cmdFindPort.Image"), System.Drawing.Image)
        Me.cmdFindPort.Location = New System.Drawing.Point(400, 144)
        Me.cmdFindPort.Name = "cmdFindPort"
        Me.cmdFindPort.Size = New System.Drawing.Size(48, 23)
        Me.cmdFindPort.TabIndex = 83
        Me.ToolTip1.SetToolTip(Me.cmdFindPort, "Find Port")
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(128, 144)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(256, 20)
        Me.txtLocation.TabIndex = 82
        '
        'txtCustComment
        '
        Me.txtCustComment.BackColor = System.Drawing.SystemColors.Info
        Me.txtCustComment.Location = New System.Drawing.Point(272, 16)
        Me.txtCustComment.Multiline = True
        Me.txtCustComment.Name = "txtCustComment"
        Me.txtCustComment.Size = New System.Drawing.Size(536, 40)
        Me.txtCustComment.TabIndex = 81
        Me.txtCustComment.Visible = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(208, 232)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(192, 23)
        Me.LinkLabel1.TabIndex = 75
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "LinkLabel1"
        Me.LinkLabel1.Visible = False
        '
        'cmdPortCosts
        '
        Me.cmdPortCosts.Image = CType(resources.GetObject("cmdPortCosts.Image"), System.Drawing.Image)
        Me.cmdPortCosts.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPortCosts.Location = New System.Drawing.Point(360, 176)
        Me.cmdPortCosts.Name = "cmdPortCosts"
        Me.cmdPortCosts.Size = New System.Drawing.Size(80, 23)
        Me.cmdPortCosts.TabIndex = 74
        Me.cmdPortCosts.Text = "Port Costs"
        Me.cmdPortCosts.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdEditVessel
        '
        Me.cmdEditVessel.Image = CType(resources.GetObject("cmdEditVessel.Image"), System.Drawing.Image)
        Me.cmdEditVessel.Location = New System.Drawing.Point(105, 120)
        Me.cmdEditVessel.Name = "cmdEditVessel"
        Me.cmdEditVessel.Size = New System.Drawing.Size(23, 23)
        Me.cmdEditVessel.TabIndex = 73
        Me.ToolTip1.SetToolTip(Me.cmdEditVessel, "Edit Site information")
        '
        'NoDonorsActual
        '
        Me.NoDonorsActual.Items.Add("-")
        Me.NoDonorsActual.Items.Add("Estimate")
        Me.NoDonorsActual.Items.Add("Actual")
        Me.NoDonorsActual.Location = New System.Drawing.Point(200, 280)
        Me.NoDonorsActual.Name = "NoDonorsActual"
        Me.NoDonorsActual.Size = New System.Drawing.Size(80, 20)
        Me.NoDonorsActual.TabIndex = 72
        Me.ToolTip1.SetToolTip(Me.NoDonorsActual, "Select whether the no of donors is an Estimate or known Accurately")
        Me.NoDonorsActual.Wrap = True
        '
        'cmdSiteAddress
        '
        Me.cmdSiteAddress.Image = CType(resources.GetObject("cmdSiteAddress.Image"), System.Drawing.Image)
        Me.cmdSiteAddress.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSiteAddress.Location = New System.Drawing.Point(376, 120)
        Me.cmdSiteAddress.Name = "cmdSiteAddress"
        Me.cmdSiteAddress.Size = New System.Drawing.Size(72, 23)
        Me.cmdSiteAddress.TabIndex = 67
        Me.cmdSiteAddress.Text = "Address"
        Me.cmdSiteAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSiteAddress, "See details of the address")
        '
        'cmdViewPanel
        '
        Me.cmdViewPanel.Image = CType(resources.GetObject("cmdViewPanel.Image"), System.Drawing.Image)
        Me.cmdViewPanel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdViewPanel.Location = New System.Drawing.Point(328, 568)
        Me.cmdViewPanel.Name = "cmdViewPanel"
        Me.cmdViewPanel.Size = New System.Drawing.Size(96, 23)
        Me.cmdViewPanel.TabIndex = 71
        Me.cmdViewPanel.Text = "View Panel"
        Me.cmdViewPanel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdViewPanel, "See details of the tests to be done")
        '
        'cbTestType
        '
        Me.cbTestType.ItemHeight = 13
        Me.cbTestType.Location = New System.Drawing.Point(128, 504)
        Me.cbTestType.Name = "cbTestType"
        Me.cbTestType.Size = New System.Drawing.Size(192, 21)
        Me.cbTestType.TabIndex = 16
        Me.cbTestType.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 504)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 69
        Me.Label2.Text = "Test Type"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCustomerComment
        '
        Me.cmdCustomerComment.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCustomerComment.Location = New System.Drawing.Point(320, 336)
        Me.cmdCustomerComment.Name = "cmdCustomerComment"
        Me.cmdCustomerComment.Size = New System.Drawing.Size(128, 23)
        Me.cmdCustomerComment.TabIndex = 68
        Me.cmdCustomerComment.Text = "Private Comment"
        Me.cmdCustomerComment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInfo
        '
        Me.lblInfo.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblInfo.Location = New System.Drawing.Point(288, 288)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(472, 16)
        Me.lblInfo.TabIndex = 66
        Me.lblInfo.Visible = False
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(128, 328)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComments.Size = New System.Drawing.Size(190, 40)
        Me.txtComments.TabIndex = 11
        '
        'lblComments
        '
        Me.lblComments.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblComments.Location = New System.Drawing.Point(0, 336)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(120, 23)
        Me.lblComments.TabIndex = 61
        Me.lblComments.Text = "Associated Comments"
        '
        'cmdAddVessel
        '
        Me.cmdAddVessel.Image = CType(resources.GetObject("cmdAddVessel.Image"), System.Drawing.Image)
        Me.cmdAddVessel.Location = New System.Drawing.Point(328, 120)
        Me.cmdAddVessel.Name = "cmdAddVessel"
        Me.cmdAddVessel.Size = New System.Drawing.Size(23, 23)
        Me.cmdAddVessel.TabIndex = 59
        Me.ToolTip1.SetToolTip(Me.cmdAddVessel, "Add New Vessel ")
        '
        'lbllkAddDoc
        '
        Me.lbllkAddDoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbllkAddDoc.Location = New System.Drawing.Point(456, 224)
        Me.lbllkAddDoc.Name = "lbllkAddDoc"
        Me.lbllkAddDoc.Size = New System.Drawing.Size(208, 14)
        Me.lbllkAddDoc.TabIndex = 54
        '
        'cbCurrency
        '
        Me.cbCurrency.ItemHeight = 13
        Me.cbCurrency.Location = New System.Drawing.Point(128, 176)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(48, 21)
        Me.cbCurrency.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.cbCurrency, "Currency to be used")
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
        Me.cmdBrowse.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdBrowse.Location = New System.Drawing.Point(320, 256)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(96, 23)
        Me.cmdBrowse.TabIndex = 49
        Me.cmdBrowse.Text = "Browse"
        Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddDoc
        '
        Me.txtAddDoc.Location = New System.Drawing.Point(128, 256)
        Me.txtAddDoc.Name = "txtAddDoc"
        Me.txtAddDoc.Size = New System.Drawing.Size(187, 20)
        Me.txtAddDoc.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtAddDoc, "Additional documents provided by the customer")
        '
        'lblAddDoc
        '
        Me.lblAddDoc.Location = New System.Drawing.Point(8, 256)
        Me.lblAddDoc.Name = "lblAddDoc"
        Me.lblAddDoc.Size = New System.Drawing.Size(120, 23)
        Me.lblAddDoc.TabIndex = 47
        Me.lblAddDoc.Text = "Additional Documents"
        Me.lblAddDoc.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtAddCosts
        '
        Me.txtAddCosts.Location = New System.Drawing.Point(184, 176)
        Me.txtAddCosts.Name = "txtAddCosts"
        Me.txtAddCosts.Size = New System.Drawing.Size(80, 20)
        Me.txtAddCosts.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtAddCosts, "Amount to be charged due to high collection costs")
        '
        'lblAddCosts
        '
        Me.lblAddCosts.Location = New System.Drawing.Point(40, 176)
        Me.lblAddCosts.Name = "lblAddCosts"
        Me.lblAddCosts.Size = New System.Drawing.Size(88, 23)
        Me.lblAddCosts.TabIndex = 45
        Me.lblAddCosts.Text = "Additional Costs"
        Me.lblAddCosts.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtAgentphone
        '
        Me.txtAgentphone.Location = New System.Drawing.Point(128, 440)
        Me.txtAgentphone.Name = "txtAgentphone"
        Me.txtAgentphone.Size = New System.Drawing.Size(184, 20)
        Me.txtAgentphone.TabIndex = 14
        '
        'lblagentPHOne
        '
        Me.lblagentPHOne.Location = New System.Drawing.Point(16, 440)
        Me.lblagentPHOne.Name = "lblagentPHOne"
        Me.lblagentPHOne.Size = New System.Drawing.Size(104, 23)
        Me.lblagentPHOne.TabIndex = 35
        Me.lblagentPHOne.Text = "Agent's Phone No."
        Me.lblagentPHOne.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtAgentName
        '
        Me.txtAgentName.Location = New System.Drawing.Point(128, 408)
        Me.txtAgentName.Name = "txtAgentName"
        Me.txtAgentName.Size = New System.Drawing.Size(184, 20)
        Me.txtAgentName.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtAgentName, "Name of the Agent (Maritime)")
        '
        'lblAgentNAme
        '
        Me.lblAgentNAme.Location = New System.Drawing.Point(32, 416)
        Me.lblAgentNAme.Name = "lblAgentNAme"
        Me.lblAgentNAme.Size = New System.Drawing.Size(88, 23)
        Me.lblAgentNAme.TabIndex = 33
        Me.lblAgentNAme.Text = "Agent's Name"
        Me.lblAgentNAme.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbSampleTemplate
        '
        Me.cbSampleTemplate.ItemHeight = 13
        Me.cbSampleTemplate.Location = New System.Drawing.Point(128, 568)
        Me.cbSampleTemplate.Name = "cbSampleTemplate"
        Me.cbSampleTemplate.Size = New System.Drawing.Size(192, 21)
        Me.cbSampleTemplate.TabIndex = 17
        '
        'cbRandom
        '
        Me.cbRandom.ItemHeight = 13
        Me.cbRandom.Location = New System.Drawing.Point(128, 472)
        Me.cbRandom.Name = "cbRandom"
        Me.cbRandom.Size = New System.Drawing.Size(280, 21)
        Me.cbRandom.TabIndex = 15
        '
        'cBLocality
        '
        Me.cBLocality.ItemHeight = 13
        Me.cBLocality.Items.AddRange(New Object() {"UK", "Overseas"})
        Me.cBLocality.Location = New System.Drawing.Point(216, 88)
        Me.cBLocality.Name = "cBLocality"
        Me.cBLocality.Size = New System.Drawing.Size(104, 21)
        Me.cBLocality.TabIndex = 0
        Me.cBLocality.TabStop = False
        '
        'txtWhototest
        '
        Me.txtWhototest.Location = New System.Drawing.Point(128, 304)
        Me.txtWhototest.Name = "txtWhototest"
        Me.txtWhototest.Size = New System.Drawing.Size(187, 20)
        Me.txtWhototest.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtWhototest, "Description of whom will be tested")
        '
        'updnNoDonors
        '
        Me.updnNoDonors.Location = New System.Drawing.Point(128, 280)
        Me.updnNoDonors.Name = "updnNoDonors"
        Me.updnNoDonors.Size = New System.Drawing.Size(64, 20)
        Me.updnNoDonors.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.updnNoDonors, "No of Donors from whom samples will be collected")
        '
        'lblTestPanel
        '
        Me.lblTestPanel.Location = New System.Drawing.Point(24, 568)
        Me.lblTestPanel.Name = "lblTestPanel"
        Me.lblTestPanel.Size = New System.Drawing.Size(100, 23)
        Me.lblTestPanel.TabIndex = 20
        Me.lblTestPanel.Text = "Sample Template"
        Me.lblTestPanel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCollType
        '
        Me.lblCollType.Location = New System.Drawing.Point(24, 472)
        Me.lblCollType.Name = "lblCollType"
        Me.lblCollType.Size = New System.Drawing.Size(100, 23)
        Me.lblCollType.TabIndex = 18
        Me.lblCollType.Text = "Reason for test"
        Me.lblCollType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblWhototest
        '
        Me.lblWhototest.Location = New System.Drawing.Point(24, 304)
        Me.lblWhototest.Name = "lblWhototest"
        Me.lblWhototest.Size = New System.Drawing.Size(100, 23)
        Me.lblWhototest.TabIndex = 16
        Me.lblWhototest.Text = "Who to test"
        Me.lblWhototest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNDonors
        '
        Me.lblNDonors.Location = New System.Drawing.Point(24, 280)
        Me.lblNDonors.Name = "lblNDonors"
        Me.lblNDonors.Size = New System.Drawing.Size(100, 23)
        Me.lblNDonors.TabIndex = 15
        Me.lblNDonors.Text = "No of Donors"
        Me.lblNDonors.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbAnnouced
        '
        Me.cbAnnouced.ItemHeight = 13
        Me.cbAnnouced.Items.AddRange(New Object() {"Annouced", "Unannouced"})
        Me.cbAnnouced.Location = New System.Drawing.Point(128, 384)
        Me.cbAnnouced.Name = "cbAnnouced"
        Me.cbAnnouced.Size = New System.Drawing.Size(128, 21)
        Me.cbAnnouced.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.cbAnnouced, "Whether the testees will know that a collection is being done")
        '
        'lblannouced
        '
        Me.lblannouced.Location = New System.Drawing.Point(24, 384)
        Me.lblannouced.Name = "lblannouced"
        Me.lblannouced.Size = New System.Drawing.Size(100, 23)
        Me.lblannouced.TabIndex = 13
        Me.lblannouced.Text = "Annouced"
        Me.lblannouced.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTCollDate
        '
        Me.DTCollDate.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.DTCollDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTCollDate.Location = New System.Drawing.Point(128, 200)
        Me.DTCollDate.Name = "DTCollDate"
        Me.DTCollDate.Size = New System.Drawing.Size(136, 20)
        Me.DTCollDate.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.DTCollDate, "Date of the collection")
        Me.DTCollDate.Value = New Date(2005, 4, 19, 16, 17, 55, 474)
        '
        'lblCollDate
        '
        Me.lblCollDate.Location = New System.Drawing.Point(24, 200)
        Me.lblCollDate.Name = "lblCollDate"
        Me.lblCollDate.Size = New System.Drawing.Size(100, 23)
        Me.lblCollDate.TabIndex = 11
        Me.lblCollDate.Text = "Collection Date"
        Me.lblCollDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbBlnProc
        '
        Me.cbBlnProc.ItemHeight = 13
        Me.cbBlnProc.Items.AddRange(New Object() {"No", "Yes"})
        Me.cbBlnProc.Location = New System.Drawing.Point(128, 224)
        Me.cbBlnProc.Name = "cbBlnProc"
        Me.cbBlnProc.Size = New System.Drawing.Size(61, 21)
        Me.cbBlnProc.TabIndex = 7
        '
        'lblProcedures
        '
        Me.lblProcedures.Location = New System.Drawing.Point(8, 224)
        Me.lblProcedures.Name = "lblProcedures"
        Me.lblProcedures.Size = New System.Drawing.Size(120, 23)
        Me.lblProcedures.TabIndex = 9
        Me.lblProcedures.Text = "Procedure required"
        Me.lblProcedures.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLocation
        '
        Me.lblLocation.Location = New System.Drawing.Point(24, 144)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(100, 23)
        Me.lblLocation.TabIndex = 7
        Me.lblLocation.Text = "Location Port/Site"
        Me.lblLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVessel
        '
        Me.lblVessel.Location = New System.Drawing.Point(3, 120)
        Me.lblVessel.Name = "lblVessel"
        Me.lblVessel.Size = New System.Drawing.Size(100, 23)
        Me.lblVessel.TabIndex = 5
        Me.lblVessel.Text = "Vessel Name"
        Me.lblVessel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbSiteVessel
        '
        Me.cbSiteVessel.ItemHeight = 13
        Me.cbSiteVessel.Items.AddRange(New Object() {"Site", "Vessel"})
        Me.cbSiteVessel.Location = New System.Drawing.Point(128, 88)
        Me.cbSiteVessel.Name = "cbSiteVessel"
        Me.cbSiteVessel.Size = New System.Drawing.Size(75, 21)
        Me.cbSiteVessel.TabIndex = 1
        '
        'lblSite
        '
        Me.lblSite.Location = New System.Drawing.Point(24, 88)
        Me.lblSite.Name = "lblSite"
        Me.lblSite.Size = New System.Drawing.Size(100, 23)
        Me.lblSite.TabIndex = 3
        Me.lblSite.Text = "Site or Vessel"
        Me.lblSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblID
        '
        Me.lblID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblID.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblID.Location = New System.Drawing.Point(106, 0)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(382, 16)
        Me.lblID.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer Reference"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(128, 280)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(64, 20)
        Me.NumericUpDown1.TabIndex = 21
        '
        'cmdEditSite
        '
        Me.cmdEditSite.Image = CType(resources.GetObject("cmdEditSite.Image"), System.Drawing.Image)
        Me.cmdEditSite.Location = New System.Drawing.Point(107, 120)
        Me.cmdEditSite.Name = "cmdEditSite"
        Me.cmdEditSite.Size = New System.Drawing.Size(23, 23)
        Me.cmdEditSite.TabIndex = 65
        Me.ToolTip1.SetToolTip(Me.cmdEditSite, "Edit Site information")
        '
        'cmdNextDueVessel
        '
        Me.cmdNextDueVessel.ImageIndex = 0
        Me.cmdNextDueVessel.ImageList = Me.btnImage
        Me.cmdNextDueVessel.Location = New System.Drawing.Point(352, 120)
        Me.cmdNextDueVessel.Name = "cmdNextDueVessel"
        Me.cmdNextDueVessel.Size = New System.Drawing.Size(23, 23)
        Me.cmdNextDueVessel.TabIndex = 64
        Me.ToolTip1.SetToolTip(Me.cmdNextDueVessel, "Scroll through vessels in next test order")
        '
        'btnImage
        '
        Me.btnImage.ImageStream = CType(resources.GetObject("btnImage.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.btnImage.TransparentColor = System.Drawing.Color.Transparent
        Me.btnImage.Images.SetKeyName(0, "")
        Me.btnImage.Images.SetKeyName(1, "")
        '
        'cmdNewSite
        '
        Me.cmdNewSite.Image = CType(resources.GetObject("cmdNewSite.Image"), System.Drawing.Image)
        Me.cmdNewSite.Location = New System.Drawing.Point(328, 120)
        Me.cmdNewSite.Name = "cmdNewSite"
        Me.cmdNewSite.Size = New System.Drawing.Size(23, 23)
        Me.cmdNewSite.TabIndex = 63
        Me.ToolTip1.SetToolTip(Me.cmdNewSite, "Add new Customer  Site")
        '
        'cbCustSites
        '
        Me.cbCustSites.ItemHeight = 13
        Me.cbCustSites.Location = New System.Drawing.Point(123, 120)
        Me.cbCustSites.Name = "cbCustSites"
        Me.cbCustSites.Size = New System.Drawing.Size(187, 21)
        Me.cbCustSites.TabIndex = 2
        Me.cbCustSites.Visible = False
        '
        'cbVessel
        '
        Me.HelpProvider1.SetHelpKeyword(Me.cbVessel, "Vessel")
        Me.HelpProvider1.SetHelpNavigator(Me.cbVessel, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.cbVessel.ItemHeight = 13
        Me.cbVessel.Location = New System.Drawing.Point(123, 120)
        Me.cbVessel.Name = "cbVessel"
        Me.HelpProvider1.SetShowHelp(Me.cbVessel, True)
        Me.cbVessel.Size = New System.Drawing.Size(186, 21)
        Me.cbVessel.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cbVessel, "Vessel Name")
        '
        'GBConfirm
        '
        Me.GBConfirm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GBConfirm.Controls.Add(Me.PnlConfirmations)
        Me.GBConfirm.Controls.Add(Me.Label9)
        Me.GBConfirm.Controls.Add(Me.lblInvoice)
        Me.GBConfirm.Controls.Add(Me.txtConfAdddress)
        Me.GBConfirm.Controls.Add(Me.lblConfAddress)
        Me.GBConfirm.Controls.Add(Me.TxtConfrecip)
        Me.GBConfirm.Controls.Add(Me.cbConfMethod)
        Me.GBConfirm.Controls.Add(Me.lblConfMethod)
        Me.GBConfirm.Controls.Add(Me.cbConfType)
        Me.GBConfirm.Controls.Add(Me.cbPManager)
        Me.GBConfirm.Controls.Add(Me.lblBy)
        Me.GBConfirm.Controls.Add(Me.cbPMAnaged)
        Me.GBConfirm.Controls.Add(Me.lblPm)
        Me.GBConfirm.Controls.Add(Me.Label3)
        Me.GBConfirm.Controls.Add(Me.Label4)
        Me.GBConfirm.Controls.Add(Me.ComboBox2)
        Me.GBConfirm.Location = New System.Drawing.Point(456, 96)
        Me.GBConfirm.Name = "GBConfirm"
        Me.GBConfirm.Size = New System.Drawing.Size(352, 80)
        Me.GBConfirm.TabIndex = 77
        Me.GBConfirm.TabStop = False
        '
        'PnlConfirmations
        '
        Me.PnlConfirmations.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PnlConfirmations.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PnlConfirmations.Controls.Add(Me.rtbContacts)
        Me.PnlConfirmations.Location = New System.Drawing.Point(0, 32)
        Me.PnlConfirmations.Name = "PnlConfirmations"
        Me.PnlConfirmations.Size = New System.Drawing.Size(344, 40)
        Me.PnlConfirmations.TabIndex = 78
        '
        'rtbContacts
        '
        Me.rtbContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbContacts.Location = New System.Drawing.Point(0, 0)
        Me.rtbContacts.Name = "rtbContacts"
        Me.rtbContacts.Size = New System.Drawing.Size(340, 36)
        Me.rtbContacts.TabIndex = 0
        Me.rtbContacts.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 144)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(296, 23)
        Me.Label9.TabIndex = 77
        Me.Label9.Text = "Confirmations will go to the current confirmation contacts"
        '
        'lblInvoice
        '
        Me.lblInvoice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInvoice.Location = New System.Drawing.Point(264, 184)
        Me.lblInvoice.Name = "lblInvoice"
        Me.lblInvoice.Size = New System.Drawing.Size(72, 23)
        Me.lblInvoice.TabIndex = 76
        '
        'txtConfAdddress
        '
        Me.txtConfAdddress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConfAdddress.Enabled = False
        Me.ErrorProvider1.SetIconAlignment(Me.txtConfAdddress, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtConfAdddress.Location = New System.Drawing.Point(118, 117)
        Me.txtConfAdddress.Name = "txtConfAdddress"
        Me.txtConfAdddress.Size = New System.Drawing.Size(216, 20)
        Me.txtConfAdddress.TabIndex = 63
        '
        'lblConfAddress
        '
        Me.lblConfAddress.Location = New System.Drawing.Point(8, 112)
        Me.lblConfAddress.Name = "lblConfAddress"
        Me.lblConfAddress.Size = New System.Drawing.Size(104, 27)
        Me.lblConfAddress.TabIndex = 74
        Me.lblConfAddress.Text = "Confirmation Address"
        Me.lblConfAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TxtConfrecip
        '
        Me.TxtConfrecip.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtConfrecip.Enabled = False
        Me.ErrorProvider1.SetIconAlignment(Me.TxtConfrecip, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.TxtConfrecip.Location = New System.Drawing.Point(118, 88)
        Me.TxtConfrecip.Name = "TxtConfrecip"
        Me.TxtConfrecip.Size = New System.Drawing.Size(216, 20)
        Me.TxtConfrecip.TabIndex = 62
        '
        'cbConfMethod
        '
        Me.cbConfMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConfMethod.Enabled = False
        Me.ErrorProvider1.SetIconAlignment(Me.cbConfMethod, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cbConfMethod.ItemHeight = 13
        Me.cbConfMethod.Location = New System.Drawing.Point(118, 64)
        Me.cbConfMethod.Name = "cbConfMethod"
        Me.cbConfMethod.Size = New System.Drawing.Size(216, 21)
        Me.cbConfMethod.TabIndex = 61
        Me.ToolTip1.SetToolTip(Me.cbConfMethod, "How the confirmation will be sent")
        Me.cbConfMethod.Visible = False
        '
        'lblConfMethod
        '
        Me.lblConfMethod.Location = New System.Drawing.Point(8, 64)
        Me.lblConfMethod.Name = "lblConfMethod"
        Me.lblConfMethod.Size = New System.Drawing.Size(109, 23)
        Me.lblConfMethod.TabIndex = 70
        Me.lblConfMethod.Text = "Confirmation Method"
        Me.lblConfMethod.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblConfMethod.Visible = False
        '
        'cbConfType
        '
        Me.cbConfType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConfType.Enabled = False
        Me.ErrorProvider1.SetIconAlignment(Me.cbConfType, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cbConfType.ItemHeight = 13
        Me.cbConfType.Location = New System.Drawing.Point(118, 32)
        Me.cbConfType.Name = "cbConfType"
        Me.cbConfType.Size = New System.Drawing.Size(216, 21)
        Me.cbConfType.TabIndex = 60
        Me.ToolTip1.SetToolTip(Me.cbConfType, "If the collection is to be confirmed indicate at what points it will be confirmed" & _
                "")
        Me.cbConfType.Visible = False
        '
        'cbPManager
        '
        Me.cbPManager.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbPManager.ItemHeight = 13
        Me.cbPManager.Location = New System.Drawing.Point(196, 8)
        Me.cbPManager.Name = "cbPManager"
        Me.cbPManager.Size = New System.Drawing.Size(136, 21)
        Me.cbPManager.TabIndex = 57
        '
        'lblBy
        '
        Me.lblBy.Location = New System.Drawing.Point(176, 8)
        Me.lblBy.Name = "lblBy"
        Me.lblBy.Size = New System.Drawing.Size(24, 23)
        Me.lblBy.TabIndex = 66
        Me.lblBy.Text = "by"
        '
        'cbPMAnaged
        '
        Me.cbPMAnaged.ItemHeight = 13
        Me.cbPMAnaged.Items.AddRange(New Object() {"No", "Yes"})
        Me.cbPMAnaged.Location = New System.Drawing.Point(117, 8)
        Me.cbPMAnaged.Name = "cbPMAnaged"
        Me.cbPMAnaged.Size = New System.Drawing.Size(61, 21)
        Me.cbPMAnaged.TabIndex = 56
        '
        'lblPm
        '
        Me.lblPm.Location = New System.Drawing.Point(0, 8)
        Me.lblPm.Name = "lblPm"
        Me.lblPm.Size = New System.Drawing.Size(115, 23)
        Me.lblPm.TabIndex = 65
        Me.lblPm.Text = "Programme Managed"
        Me.lblPm.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(20, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 23)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Confirmation Type"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.Label3.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 24)
        Me.Label4.TabIndex = 73
        Me.Label4.Text = "Confirmation Recipient"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ComboBox2
        '
        Me.ComboBox2.ItemHeight = 13
        Me.ComboBox2.Location = New System.Drawing.Point(118, 32)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(232, 21)
        Me.ComboBox2.TabIndex = 59
        Me.ToolTip1.SetToolTip(Me.ComboBox2, "If the collection is to be confirmed indicate at what points it will be confirmed" & _
                "")
        Me.ComboBox2.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.pnlOfficer)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(812, 535)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Officer"
        '
        'pnlOfficer
        '
        Me.pnlOfficer.Controls.Add(Me.LVRedirect)
        Me.pnlOfficer.Controls.Add(Me.Splitter1)
        Me.pnlOfficer.Controls.Add(Me.ListView1)
        Me.pnlOfficer.Controls.Add(Me.HtmlEditor1)
        Me.pnlOfficer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOfficer.Location = New System.Drawing.Point(0, 0)
        Me.pnlOfficer.Name = "pnlOfficer"
        Me.pnlOfficer.Size = New System.Drawing.Size(812, 535)
        Me.pnlOfficer.TabIndex = 0
        '
        'LVRedirect
        '
        Me.LVRedirect.BackColor = System.Drawing.SystemColors.Info
        Me.LVRedirect.CheckBoxes = True
        Me.LVRedirect.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.LVRedirect.ContextMenu = Me.ctxOfficerEmail
        Me.LVRedirect.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LVRedirect.FullRowSelect = True
        Me.LVRedirect.Location = New System.Drawing.Point(499, 0)
        Me.LVRedirect.Name = "LVRedirect"
        Me.LVRedirect.Size = New System.Drawing.Size(313, 255)
        Me.LVRedirect.SmallImageList = Me.multisendImg
        Me.LVRedirect.TabIndex = 14
        Me.LVRedirect.UseCompatibleStateImageBehavior = False
        Me.LVRedirect.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "From"
        Me.ColumnHeader4.Width = 80
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "To"
        Me.ColumnHeader5.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Address"
        Me.ColumnHeader6.Width = 200
        '
        'ctxOfficerEmail
        '
        Me.ctxOfficerEmail.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSendEmail})
        '
        'mnuSendEmail
        '
        Me.mnuSendEmail.Index = 0
        Me.mnuSendEmail.Text = "&Send Email"
        '
        'multisendImg
        '
        Me.multisendImg.ImageStream = CType(resources.GetObject("multisendImg.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.multisendImg.TransparentColor = System.Drawing.Color.Transparent
        Me.multisendImg.Images.SetKeyName(0, "")
        Me.multisendImg.Images.SetKeyName(1, "")
        Me.multisendImg.Images.SetKeyName(2, "")
        Me.multisendImg.Images.SetKeyName(3, "")
        Me.multisendImg.Images.SetKeyName(4, "")
        Me.multisendImg.Images.SetKeyName(5, "")
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(496, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 255)
        Me.Splitter1.TabIndex = 10
        Me.Splitter1.TabStop = False
        '
        'ListView1
        '
        Me.ListView1.AllowColumnReorder = True
        Me.ListView1.BackColor = System.Drawing.SystemColors.Info
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colID, Me.ColForeName, Me.colSurname, Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.ListView1.ContextMenu = Me.ctxOfficer
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Left
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HideSelection = False
        Me.ListView1.LargeImageList = Me.ImageList1
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(496, 255)
        Me.ListView1.TabIndex = 9
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'colID
        '
        Me.colID.Text = "Identity"
        Me.colID.Width = 80
        '
        'ColForeName
        '
        Me.ColForeName.Text = "Fore Name"
        Me.ColForeName.Width = 80
        '
        'colSurname
        '
        Me.colSurname.Text = "Surname"
        Me.colSurname.Width = 80
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Pin Person"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Cost"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Cost type"
        Me.ColumnHeader3.Width = 200
        '
        'ctxOfficer
        '
        Me.ctxOfficer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddOfficer, Me.mnuRemoveOfficer})
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
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 255)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(812, 280)
        Me.HtmlEditor1.TabIndex = 7
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.cmdSendRandom)
        Me.TabPage3.Controls.Add(Me.lvRandom)
        Me.TabPage3.Controls.Add(Me.txtEditor)
        Me.TabPage3.Controls.Add(Me.cmdCreateDocument)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.udEditor)
        Me.TabPage3.Controls.Add(Me.lvColumnGroups)
        Me.TabPage3.Controls.Add(Me.lblTemplate)
        Me.TabPage3.Controls.Add(Me.cmdEditDoc)
        Me.TabPage3.Controls.Add(Me.cmdGenerateRandom)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(812, 535)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Random Numbers"
        '
        'cmdSendRandom
        '
        Me.cmdSendRandom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSendRandom.Image = Global.MedscreenCommonGui.AdcResource.Mail1
        Me.cmdSendRandom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSendRandom.Location = New System.Drawing.Point(659, 285)
        Me.cmdSendRandom.Name = "cmdSendRandom"
        Me.cmdSendRandom.Size = New System.Drawing.Size(130, 23)
        Me.cmdSendRandom.TabIndex = 15
        Me.cmdSendRandom.Text = "Send Document"
        Me.cmdSendRandom.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSendRandom, "Send Document to collecting officer")
        Me.cmdSendRandom.UseVisualStyleBackColor = True
        '
        'lvRandom
        '
        Me.lvRandom.AllowUserToOrderColumns = True
        Me.lvRandom.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvRandom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.lvRandom.Location = New System.Drawing.Point(64, 193)
        Me.lvRandom.Name = "lvRandom"
        Me.lvRandom.Size = New System.Drawing.Size(576, 320)
        Me.lvRandom.TabIndex = 14
        '
        'txtEditor
        '
        Me.txtEditor.Location = New System.Drawing.Point(544, 4)
        Me.txtEditor.Name = "txtEditor"
        Me.txtEditor.Size = New System.Drawing.Size(100, 20)
        Me.txtEditor.TabIndex = 13
        Me.txtEditor.Visible = False
        '
        'cmdCreateDocument
        '
        Me.cmdCreateDocument.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCreateDocument.Image = Global.MedscreenCommonGui.AdcResource.NewDoc
        Me.cmdCreateDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCreateDocument.Location = New System.Drawing.Point(659, 206)
        Me.cmdCreateDocument.Name = "cmdCreateDocument"
        Me.cmdCreateDocument.Size = New System.Drawing.Size(130, 23)
        Me.cmdCreateDocument.TabIndex = 12
        Me.cmdCreateDocument.Text = "Create Document"
        Me.cmdCreateDocument.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdCreateDocument.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(61, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Column Groups"
        '
        'udEditor
        '
        Me.udEditor.Location = New System.Drawing.Point(394, 4)
        Me.udEditor.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.udEditor.Name = "udEditor"
        Me.udEditor.Size = New System.Drawing.Size(120, 20)
        Me.udEditor.TabIndex = 10
        Me.udEditor.Visible = False
        '
        'lvColumnGroups
        '
        Me.lvColumnGroups.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvColumnGroups.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12})
        Me.lvColumnGroups.FullRowSelect = True
        Me.lvColumnGroups.Location = New System.Drawing.Point(64, 53)
        Me.lvColumnGroups.Name = "lvColumnGroups"
        Me.lvColumnGroups.Size = New System.Drawing.Size(576, 118)
        Me.lvColumnGroups.TabIndex = 9
        Me.lvColumnGroups.UseCompatibleStateImageBehavior = False
        Me.lvColumnGroups.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Column Group"
        Me.ColumnHeader10.Width = 129
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "No of Donors"
        Me.ColumnHeader11.Width = 137
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "From"
        Me.ColumnHeader12.Width = 169
        '
        'lblTemplate
        '
        Me.lblTemplate.AutoSize = True
        Me.lblTemplate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemplate.Location = New System.Drawing.Point(113, 0)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.Size = New System.Drawing.Size(59, 15)
        Me.lblTemplate.TabIndex = 8
        Me.lblTemplate.Text = "Template :"
        '
        'cmdEditDoc
        '
        Me.cmdEditDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEditDoc.Image = Global.MedscreenCommonGui.AdcResource.NOTES_4
        Me.cmdEditDoc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEditDoc.Location = New System.Drawing.Point(659, 246)
        Me.cmdEditDoc.Name = "cmdEditDoc"
        Me.cmdEditDoc.Size = New System.Drawing.Size(130, 23)
        Me.cmdEditDoc.TabIndex = 7
        Me.cmdEditDoc.Text = "Edit Document"
        Me.cmdEditDoc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdEditDoc.UseVisualStyleBackColor = True
        '
        'cmdGenerateRandom
        '
        Me.cmdGenerateRandom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGenerateRandom.Image = Global.MedscreenCommonGui.AdcResource.CHECKMRK
        Me.cmdGenerateRandom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdGenerateRandom.Location = New System.Drawing.Point(725, 32)
        Me.cmdGenerateRandom.Name = "cmdGenerateRandom"
        Me.cmdGenerateRandom.Size = New System.Drawing.Size(75, 23)
        Me.cmdGenerateRandom.TabIndex = 4
        Me.cmdGenerateRandom.Text = "Generate"
        Me.cmdGenerateRandom.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 100
        Me.ToolTip1.AutoPopDelay = 10000
        Me.ToolTip1.InitialDelay = 20
        Me.ToolTip1.ReshowDelay = 20
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuUtilities, Me.mnuPricing, Me.mnuPort, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 0
        Me.mnuExit.Text = "E&xit"
        '
        'mnuUtilities
        '
        Me.mnuUtilities.Index = 1
        Me.mnuUtilities.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewInvoice, Me.mnuEdAddDoc, Me.mnuCustDefaults})
        Me.mnuUtilities.Text = "&Utilites"
        '
        'mnuViewInvoice
        '
        Me.mnuViewInvoice.Index = 0
        Me.mnuViewInvoice.Text = "View &Invoice"
        '
        'mnuEdAddDoc
        '
        Me.mnuEdAddDoc.Index = 1
        Me.mnuEdAddDoc.Text = "&Edit Additional Doc"
        '
        'mnuCustDefaults
        '
        Me.mnuCustDefaults.Index = 2
        Me.mnuCustDefaults.Text = "&Customer Defaults"
        '
        'mnuPricing
        '
        Me.mnuPricing.Index = 2
        Me.mnuPricing.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuJobCost, Me.mnuCustomerPricing})
        Me.mnuPricing.Text = "&Pricing"
        '
        'mnuJobCost
        '
        Me.mnuJobCost.Index = 0
        Me.mnuJobCost.Text = "&Job Cost"
        '
        'mnuCustomerPricing
        '
        Me.mnuCustomerPricing.Index = 1
        Me.mnuCustomerPricing.Text = "&Customer Pricing"
        '
        'mnuPort
        '
        Me.mnuPort.Index = 3
        Me.mnuPort.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewPort})
        Me.mnuPort.Text = "P&orts"
        '
        'mnuNewPort
        '
        Me.mnuNewPort.Index = 0
        Me.mnuNewPort.Text = "&New Port"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 4
        Me.mnuHelp.Text = "&Help"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Multiselect = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.AlwaysBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'NWInfo
        '
        Me.NWInfo.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.SystemColors.InactiveCaption, System.Drawing.SystemColors.Window)
        Me.NWInfo.DefaultText = Nothing
        Me.NWInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NWInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.NWInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NWInfo.ShowStyle = VbPowerPack.NotificationShowStyle.Slide
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "")
        Me.ImageList2.Images.SetKeyName(1, "")
        '
        'frmCollectionForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(824, 613)
        Me.Controls.Add(Me.pnlCollection)
        Me.Controls.Add(Me.Panel1)
        Me.HelpProvider1.SetHelpKeyword(Me, "Collection form")
        Me.HelpProvider1.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmCollectionForm"
        Me.HelpProvider1.SetShowHelp(Me, True)
        Me.Text = "Collection Form"
        Me.Panel1.ResumeLayout(False)
        Me.pnlCollection.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.gbCollection.ResumeLayout(False)
        Me.gbCollection.PerformLayout()
        Me.TabSubControls.ResumeLayout(False)
        Me.tpInfo.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.tpDocuments.ResumeLayout(False)
        Me.pnlDocuments.ResumeLayout(False)
        CType(Me.updnNoDonors, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBConfirm.ResumeLayout(False)
        Me.GBConfirm.PerformLayout()
        Me.PnlConfirmations.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.pnlOfficer.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.lvRandom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Define an enumerator of the possible modes that the form can be in 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum FormModes
        Create
        Edit
        ChangePort
        'AllowCustomerChange
    End Enum

    Private MyMode As FormModes = FormModes.Create  'Variable defining form mode


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 'Property To define what mode the form is in 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormMode() As FormModes
        Get
            Return MyMode
        End Get
        Set(ByVal Value As FormModes)
            MyMode = Value
            'blnFormSet = False
            blnCustSiteSelected = False
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Setup form enabling or disabling controls to suit mode 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Friend Sub SetFormMode()
        Select Case MyMode
            Case FormModes.Create           'if in create mode 
                'cbCustomer.Enabled = True
                cbSiteVessel.Enabled = True
                Me.GBConfirm.Enabled = True
                Me.cbVessel.Enabled = True
                Me.cbCustSites.Enabled = True
                Me.cbRandom.Enabled = True
                Me.cbAnnouced.Enabled = True
                cbTestType.Enabled = True
                Me.cBLocality.Enabled = True
                Me.cmdFindPort.Enabled = True
                Me.cmdGenerateRandom.Enabled = False

                Me.cmdEditDoc.Enabled = False
            Case FormModes.Edit                     'Edit mode
                'cbCustomer.Enabled = False
                'Me.SmidCombo1.Enabled = False
                Me.cmdFindPort.Enabled = True
                cbSiteVessel.Enabled = True
                GBConfirm.Enabled = True
                Me.cbVessel.Enabled = True
                cbCustSites.Enabled = True
                Me.cbRandom.Enabled = True
                cbAnnouced.Enabled = True
                cbTestType.Enabled = True
                Me.cBLocality.Enabled = True
                'Case FormModes.AllowCustomerChange      'Allowing customer change
                '    'cbCustomer.Enabled = True
                '    'Me.SmidCombo1.Enabled = True
                '    cbSiteVessel.Enabled = True
                '    GBConfirm.Enabled = True
                '    Me.cbVessel.Enabled = True
                '    cbCustSites.Enabled = True
                '    Me.cbRandom.Enabled = True
                '    cbAnnouced.Enabled = True
                '    cbTestType.Enabled = True
                '    Me.cBLocality.Enabled = True
            Case FormModes.ChangePort               'Allowing port change
                'cbCustomer.Enabled = False
                cbSiteVessel.Enabled = False
                GBConfirm.Enabled = False
                Me.cbVessel.Enabled = False
                cbCustSites.Enabled = False
                Me.cbRandom.Enabled = False
                Me.cbAnnouced.Enabled = False
                Me.cbTestType.Enabled = False
                Me.cmdFindPort.Enabled = True
                Me.cBLocality.Enabled = True
        End Select

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes collection being used on this form 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [25/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Collection() As Intranet.intranet.jobs.CMJob
        Get
            Return myCollection
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CMJob)
            Medscreen.LogTiming("Collection Start")

            myCollection = Value
            Me.TabControl1.SelectedIndex = 0
            blnFormSet = False
            blnCustSiteSelected = False
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
            Me.ErrorProvider1.SetError(Me.txtPO, "")

            ClearForm()
            If myCollection Is Nothing Then
                myClient = Nothing
                cmdGenerateRandom.Enabled = False
            Else
                cmdGenerateRandom.Enabled = True
                myClient = cSupport.CustList.Item(myCollection.ClientId)
                SetUpClient()
                'cbCustomer_SelectedIndexChanged(Nothing, Nothing)
                If myCollection.HasID Then
                    tmpReference = myCollection.ID
                Else
                    tmpReference = Now.ToString("ddHHmmssff")
                End If
                FillForm()
                blnFormLoaded = False
                Medscreen.LogTiming("Collection End")
                cmdGenerateRandom.Enabled = True
                'Me.SetFormMode()
            End If
        End Set
    End Property

    Private Sub RebindSampleTemplate()
        Me.cbSampleTemplate.DataSource = objSampleTemplates
        Me.cbSampleTemplate.BindingContext(Me.cbSampleTemplate.DataSource).SuspendBinding()
        Me.cbSampleTemplate.BindingContext(Me.cbSampleTemplate.DataSource).ResumeBinding()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the form set up for the current client
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [25/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetUpClient()

        Medscreen.LogTiming("Setup Client Start")

        If Not myClient Is Nothing Then
            If Not myClient Is Nothing Then
                Try
                    If Not myClient.CollectionInfo Is Nothing Then
                        Me.cbSampleTemplate.SelectedValue = myClient.CollectionInfo.SampleTemplate
                        If myClient.CollectionInfo.LaboratoryId.Trim.Length > 0 Then myCollection.LaboratoryId = myClient.CollectionInfo.LaboratoryId
                        'Me.Collection.TestScheduleID = myClient.CollectionInfo.SampleTemplate
                    Else
                        Me.cbSampleTemplate.SelectedValue = myClient.ReportingInfo.SampleTemplate
                        'Me.Collection.TestScheduleID = myClient.ReportingInfo.SampleTemplate
                    End If


                Catch ex As Exception
                    LogError(ex, , "Test PAnel")
                End Try
                Try
                    Me.TabControl1.SelectedIndex = 0

                    'Me.cbCustomer.DataSource = Nothing
                    Medscreen.LogTiming("  Setup Client combos")
                    'Me.SmidCombo1.SelectedValue = myClient.SMID
                    Me.txtCustomer.Text = myClient.Info
                    objVessels = myClient.Vessels
                    Me.lblInfo.Hide()
                    Me.lblInfo.Text = ""
                    If myClient.InvoiceInfo.CreditCard_pay Then
                        Me.lblInfo.Text = "This Client must pay by credit card"
                        Me.lblInfo.Show()
                    End If
                    Medscreen.LogTiming("  Setup Client vessel info")
                    Me.txtMatrix.Text = myClient.MatrixText(myClient.ReportingInfo.SampleTemplate)
                    'Get query from config file 

                    Dim strSQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("SampleTemplate")
                    'Dim strSQuery As String = "select identity,description from Samp_Tmpl_Header s " & _
                    '"where s.removeflag='F' and action_type in ('.NET','REFER')"
                    Me.objSampleTemplates = New Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)

                    ''Check to see we have a login template.
                    If objSampleTemplates.Count = 0 Then
                        MsgBox("The login templates have not been set up for this customer, this MUST be done before collections can be created!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                        Me.DialogResult = DialogResult.Cancel
                        Me.cmdCancel_Click(Nothing, Nothing)
                    End If
                    RebindSampleTemplate()
                    Me.cbSampleTemplate.DisplayMember = "PhraseText"
                    Me.cbSampleTemplate.ValueMember = "PhraseID"


                    If (myClient.Vessels.Count = 0 And myClient.Sites.Count > 1) _
                        Or Mid(myClient.IndustryCode, 1, 3) <> "MAR" Then
                        If Not Me.Collection Is Nothing Then Me.Collection.IsVessel = False
                        Me.cbVessel.Hide()
                        Me.cmdNextDueVessel.Hide()
                        myCollection.IsSiteVisit = True
                        If (myClient.Vessels.Count = 0 And myClient.Sites.Count > 1) Then _
                            Me.cbSiteVessel.Enabled = False
                        Me.cmdAddVessel.Hide()
                        Me.cmdNewSite.Show()
                        CustSites = Me.myClient.Sites
                        Me.cbCustSites.Show()
                        myClient.Sites.Refresh()
                        Me.myClient.Vessels.Sort()

                        RebindSites()
                        Me.cbCustSites.DisplayMember = "Info"
                        Me.cbCustSites.ValueMember = "SiteID"
                        Me.lblVessel.Text = "Customer Sites"
                    Else
                        Me.cmdNextDueVessel.Show()
                        Me.cmdNextDueVessel.ImageIndex = 0
                        VesselFirstDate = Now
                        myCollection.IsVessel = True
                        Me.cbVessel.Show()
                        Me.cbSiteVessel.Enabled = True
                        Me.cmdAddVessel.Show()
                        Me.cmdNewSite.Hide()
                        Me.cbCustSites.Hide()
                        Me.lblVessel.Text = "Vessel Name"
                        Me.cmdAddVessel.Show()
                        CustSites = Me.myClient.Sites
                        RebindSites()
                        Me.cbCustSites.DisplayMember = "Info"
                        Me.cbCustSites.ValueMember = "SiteID"
                    End If
                    cbCustomer_SelectedIndexChanged(Nothing, Nothing)
                Catch ex As Exception
                End Try
            End If
        Else
            'objVessels = ModTpanel.Vessels
        End If
        cbVessel.Items.Clear()
        RebindVessels()
        cbVessel.DisplayMember = "Info"
        cbVessel.ValueMember = "VesselID"

        Medscreen.LogTiming("  Setup Client Sites")

        Try
            cbCustSites.Text = ""
            'FillForm()

            If Not myClient Is Nothing Then         ' Check we have a client

                Dim blnCustSite As Boolean = False
                Dim objSite As Intranet.intranet.Address.CustomerSite = Nothing
                Try
                    If Me.myClient.Sites.Count > 1 Then      'Has customer sites
                        objSite = Me.myCollection.CustomerSite
                        blnCustSite = Not (objSite Is Nothing)
                        'Me.cbCustSites.SelectedItem = objSite
                    End If
                Catch ex As Exception
                    LogError(ex, , "Sites ")
                End Try
                Medscreen.LogTiming("  Setup Client programme management")

                Try
                    If myClient.ProgrammeManager.Trim.Length > 0 Then

                        Me.cbPMAnaged.SelectedIndex = 0

                        If Not Me.Collection Is Nothing Then
                            Me.Collection.ProgrammeManaged = True
                            Me.Collection.ProgrammeManager = myClient.ProgrammeManager
                            Dim i As Integer
                            For i = 0 To ProGmanagers.Count - 1
                                If ProGmanagers.Item(i).User = Me.Collection.ProgrammeManager Then
                                    Me.cbPManager.SelectedIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        Me.cbPMAnaged.SelectedIndex = 1
                    End If
                Catch ex As Exception
                    LogError(ex, , "Programme ,manager")
                End Try
                Medscreen.LogTiming("  Setup Client port info")

                Try
                    Dim tmpPort As String = myCollection.Port       'Temporarily save port seems to get destroyed 
                    myCollection.Port = tmpPort                     'Restore saved value
                    Dim objPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(myCollection.Port)       'Set the site 
                    If objPort Is Nothing Then
                        Me.txtLocation.Text = ""
                    Else
                        Me.txtLocation.Text = objPort.Info2
                    End If
                    Me.lblLocation.Text = "Location Port/Site"
                Catch ex As Exception
                    LogError(ex, , "Ports")
                End Try
                Try
                    If IO.File.Exists(Medscreen.LiveRoot & "procedures\" & myClient.Identity.Trim & ".doc") Then
                        Me.cbBlnProc.SelectedIndex = 0
                        Me.LinkLabel1.Text = myClient.Identity.Trim & ".doc"
                        Me.LinkLabel1.Links.Clear()
                        Me.LinkLabel1.Links.Add(0, Me.LinkLabel1.Text.Length, "\\john\live\procedures\" & myClient.Identity.Trim & ".doc")
                    Else
                        Me.cbBlnProc.SelectedIndex = 1
                        Me.LinkLabel1.Text = ""
                        Me.LinkLabel1.Links.Clear()
                    End If
                Catch ex As Exception
                    LogError(ex, , "Proc ")
                End Try

                Medscreen.LogTiming("  Setup Client test panels")


                'End If
                'Check to see if there are any collection comments 
                Me.txtCustComment.Hide()
                Medscreen.LogTiming("  Setup Client comments")

                Dim oPColl As New Collection()
                oPColl.Add(myClient.Identity)
                oPColl.Add(" ")
                Dim strComment As String = CConnection.PackageStringList("LIB_CCTOOL.GETCUSTOMERCOLLCOMMENTS", oPColl)
                If strComment.Length > 0 Then
                    strComment = Medscreen.ReplaceString(strComment, "|", vbCrLf)
                    Me.txtCustComment.Text = strComment.ToString
                    Me.txtCustComment.Show()
                End If
                'If myClient.CustomerComments.Count > 0 Then 'We have comments
                '    Dim objComm As Intranet.intranet.jobs.CollectionComment
                '    Dim strComment As New StringBuilder("")
                '    Dim ij As Integer
                '    For ij = 0 To myClient.CustomerComments.Count - 1
                '        objComm = myClient.CustomerComments.item(ij)
                '        If objComm.CommentType = "COLL" Then    'It is a collection comment
                '            strComment.Append(objComm.CommentText)
                '            strComment.Append(vbCrLf)

                '        End If
                '    Next
                '    If strComment.Length > 0 Then 'There was a comment present
                '        Me.txtCustComment.Text = strComment.ToString
                '        Me.txtCustComment.Show()
                '    End If
                'End If
            End If ' Not myClient Is Nothing
        Catch ex As Exception
            LogError(ex, , "Global")
        End Try
        Medscreen.LogTiming("Setup Client finish")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' empty all controls or set to default value 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ClearForm()
        'This routine will clear out the contents of the form and default the comboboxes to appropriate values 

        If Me.cbVessel.Items.Count > 0 Then
            Me.cbVessel.SelectedIndex = 0
        Else
            Me.cbVessel.SelectedIndex = -1
        End If

        If Me.cbCustSites.Items.Count > 0 Then
            Me.cbCustSites.SelectedIndex = 0
        Else
            Me.cbCustSites.SelectedIndex = -1
        End If

        'If Me.cbLocation.Items.Count > 0 Then
        '    Me.cbLocation.SelectedIndex = 0
        'Else
        '    Me.cbLocation.SelectedIndex = -1
        'End If

        'If Me.cbLocation.Items.Count > 0 Then
        '    Me.cbLocation.SelectedIndex = 0
        'Else
        '    Me.cbLocation.SelectedIndex = -1
        'End If
        Me.txtAgentName.Text = ""
        Me.txtAgentphone.Text = ""

        Me.txtAddDoc.Text = ""
    End Sub

    '
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Display the collection in the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date>
    ''' <Action>Altered the way vessels are grabbed from the vessel collection for display
    ''' </Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillForm()
        Try
            Me.cmdAssign.Enabled = False
            Medscreen.LogTiming("Fill Form Start")
            If Me.Collection.HasID > 0 AndAlso Me.Collection.Port.Trim.Length = 0 Then
                Me.Collection.Load()

            End If
            blnFormSet = True
            blnCustSiteSelected = False
            Me.Panel2.Hide()
            Dim strCust As String
            'Dim myClient As Intranet.intranet.customerns.Client
            With myCollection       'With command means that collection properties are accessed 
                '                   'with only a leading . 
                strCust = .ClientId     'These two lines could be rolled into one, but this will cut down on 
                '                       'far proc calls
                'myClient = Me.objSibColl.Item(strCust)

                Medscreen.LogTiming("    Fill Form Customer Details")

                'Me.cbCustomer.SelectedProfile = myClient
                Me.txtCustomer.Text = myClient.Info
                Me.lblID.Text = .ID                         'Set the ID of the collection
                If .Invoiceid.Trim.Length > 0 Then          'Has the collection been invoiced 
                    Me.lblInvoice.Text = "click " & .Invoiceid & " to view Invoice"
                    Me.lblInvoice.Links.Clear()
                    Me.lblInvoice.Links.Add(6, 5)
                    Me.lblInvoice.Show()
                    Me.lblID.Text += " - Invoiced"
                    Me.mnuViewInvoice.Visible = True
                    Me.mnuViewInvoice2.Visible = True
                    Me.ToolTip1.SetToolTip(Me.lblInvoice, "Double click to view Invoice")
                Else                                    'Not invoiced yet 
                    Me.lblInvoice.Hide()
                    Me.mnuViewInvoice.Visible = False
                    Me.mnuViewInvoice2.Visible = False
                    Me.ToolTip1.SetToolTip(Me.lblInvoice, "")
                    Me.lblID.Text += " - " & .StatusString
                End If

                'If we don't have client info (should never occur)
                If myClient Is Nothing Then
                    'Me.txtCustomer.Text = .ClientId
                    'TODO Align combo to correct customer
                    Me.CustomerCurrency = ""
                Else
                    'Client information
                    Dim strCurr As String
                    'If Not myClient Is Nothing Then
                    'Currency info use Irish Globalisation as it has everything the same as UK,
                    'except in the Euro zone
                    strCurr = myClient.Currency
                    Me.CustomerCurrency = strCurr
                    cr = myClient.Culture
                    cbCurrency.SelectedValue = strCurr

                    'Addtional Costs display as currency using the correct globalisation
                    Me.txtAddCosts.Text = .AgreedExpenses.ToString("c", cr)
                    'Check customer requirements Highlight if customer demands a PO Number
                    If myClient.InvoiceInfo.PoRequired Then
                        Me.lblPONumber.Text = "Purchase Order*"
                        Me.lblPONumber.ForeColor = System.Drawing.SystemColors.ActiveCaption
                    Else
                        Me.lblPONumber.Text = "Purchase Order"
                        Me.lblPONumber.ForeColor = System.Drawing.SystemColors.ControlText
                    End If
                End If

                Medscreen.LogTiming("    Fill Form Customer Vessel Site")
                'If it is a vessel dependent collection then say so and set up the controls
                'so that they are correct 
                If .IsVessel Then
                    Me.cbSiteVessel.SelectedIndex = 1
                    Me.cbVessel.Show()
                    Me.lblVessel.Show()
                    Me.cbCustSites.Hide()
                    Me.cmdEditSite.Hide()
                    Me.cmdEditVessel.Show()
                    Me.lblVessel.Text = "Vessel Name"

                    Dim i As Integer
                    Dim oVessel As Intranet.intranet.customerns.Vessel

                    If .vessel.Trim.Length > 0 Then     'Do we have vessel info yet?
                        If Not objVessels Is Nothing Then
                            If objVessels.Count > 0 And objVessels.Count < 800 Then
                                oVessel = objVessels.Item(.vessel)
                                Try
                                    If Not oVessel Is Nothing Then
                                        Dim intpos As Integer = objVessels.IndexOf(oVessel)
                                        If intpos >= 0 Then Me.cbVessel.SelectedValue = oVessel.VesselID
                                    End If
                                Catch
                                End Try

                                'IF we have vessel invoicing we can display the address
                                Me.Panel2.Hide()
                                If Not oVessel Is Nothing AndAlso oVessel.Invoiceable Then 'AndAlso Not oVessel.Invoicing Is Nothing Then
                                    If oVessel.MailingAddress > 0 Then
                                        'we have an address
                                        GetSiteAddress(oVessel.MailingAddress)

                                    End If

                                End If
                            Else
                                .IsVessel = False
                                .Update()
                            End If
                        End If
                    Else ' No vessel defined yet
                        Me.cbVessel.Text = ""
                    End If
                Else                                    'Customer site or ordinary collection 
                    Me.cbSiteVessel.SelectedIndex = 0
                    Me.cbCustSites.SelectedIndex = -1
                    Me.cbVessel.Hide()
                    Me.cmdEditVessel.Hide()
                    Me.cmdEditSite.Show()
                    Me.cbCustSites.Show()
                    Me.lblVessel.Text = "Site Name"
                    Dim i As Integer
                    Me.Panel2.Hide()
                    If Not myClient Is Nothing Then
                        Me.ToolTip1.SetToolTip(Me.cbCustSites, "")
                        If Not .SiteObject Is Nothing Then
                            'ensure that customer site collection and data source are synced

                            myClient.Sites.Refresh()
                            myClient.Sites.Sort()
                            RebindSites()
                            Me.cbCustSites.DisplayMember = "Info"
                            Me.cbCustSites.ValueMember = "SiteID"

                            For i = 0 To cbCustSites.DataSource.Count - 1
                                If cbCustSites.DataSource.Item(i).SiteID = .vessel Then     'found site 
                                    blnCustSiteSelected = True
                                    Dim custSite As Intranet.intranet.Address.CustomerSite
                                    Me.cbCustSites.SelectedIndex = i
                                    custSite = cbCustSites.DataSource.Item(i)
                                    Me.ToolTip1.SetToolTip(Me.cbCustSites, custSite.Info)
                                    'Get site address 
                                    'GetSiteAddress(custSite.AddressID)
                                    'A customer site is tied to a distinct port get that and align to it 
                                    Dim oPort As Intranet.intranet.jobs.Port = Intranet.intranet.Support.cSupport.Ports.Item(custSite.VesselSubType)
                                    If oPort.Identity.Trim.Length > 0 Then         'Found the port 
                                        .UK = oPort.IsUK
                                        Dim intPort As Integer
                                        If oPort.IsUK Then              'Get UK ports list 
                                            intPort = cSupport.UKPorts.IndexOf(oPort)
                                            'Me.cbLocation.DataSource = cSupport.UKPorts
                                        Else                            'Get overseas port list
                                            intPort = cSupport.OSPorts.IndexOf(oPort)
                                            'Me.cbLocation.DataSource = cSupport.OSPorts
                                        End If
                                        If intPort > -1 Then
                                            'Me.cbLocation.SelectedIndex = intPort
                                        End If
                                        Me.txtLocation.Text = oPort.Info2
                                    End If

                                    Exit For
                                End If
                            Next

                            Me.cbSiteVessel.Text = Me.cbSiteVessel.Items(Me.cbSiteVessel.SelectedIndex)
                        End If
                    End If
                    'Me.lblVessel.Hide()
                End If

                Medscreen.LogTiming("    Fill Form Port Info")

                'Is it a UK based collection 
                If .UK Then
                    'Me.cbLocation.DataSource = cSupport.UKPorts
                    Me.cBLocality.SelectedIndex = 0
                    Me.DTCollDate.CustomFormat = "dd-MMM-yyyy HH:mm"
                Else
                    'Me.cbLocation.DataSource = cSupport.OSPorts
                    Me.cBLocality.SelectedIndex = 1
                    Me.DTCollDate.CustomFormat = "dd-MMM-yyyy"
                End If

                'Do we have a port defined 
                If .Port.Trim.Length > 0 Then
                    objPort = cSupport.Ports.Item(.Port)       'Set the site 
                    'Me.cbLocation.SelectedItem = objPort
                    Me.txtLocation.Text = objPort.Info2
                    If .UK Then
                        'Me.cbLocation.SelectedIndex() = cSupport.UKPorts.IndexOf(objPort)
                    Else
                        'Me.cbLocation.SelectedIndex() = cSupport.OSPorts.IndexOf(objPort)
                    End If

                    '
                Else
                    Me.txtLocation.Text = ""
                End If

                Medscreen.LogTiming("    Fill Form Combos")

                'A id of zero indicates that it is a new collection 
                If .HasID = 0 Then                             'New collection
                    Me.DTCollDate.Value = Today '.AddDays(5) default to today
                Else
                    Me.DTCollDate.Value = .CollectionDate   'Set the collection date 
                End If

                Try
                    If cbAnnouced.Items.Count > 0 Then
                        If .Announced Then
                            Me.cbAnnouced.SelectedIndex = 1
                        Else
                            Me.cbAnnouced.SelectedIndex = 0
                        End If
                    End If
                Catch ex As Exception
                    LogError(ex, , "Fill Form Announced")
                End Try


                Try
                    'If we have special procedures then indicate it 
                    If .SpecialProcedures Then
                        Me.cbBlnProc.SelectedIndex = 0
                    Else
                        Me.cbBlnProc.SelectedIndex = 1
                    End If


                    Me.updnNoDonors.Value = .NoOfDonors         'Set the expected number of Donors
                Catch ex As Exception
                    LogError(ex, , "Fill Form Procedures")
                End Try

                'additional details 
                Try
                    'If .UK Then                                 'If it is a UK based or Overseas
                    '    Me.cBLocality.SelectedIndex = 0
                    'Else
                    '    Me.cBLocality.SelectedIndex = 1
                    'End If

                    Me.txtWhototest.Text = .WhoToTest           'Indcate Who to Test 
                    Me.ToolTip1.SetToolTip(Me.txtWhototest, .WhoToTest)
                Catch ex As Exception
                    LogError(ex, , "Fill Form Who to test " & .WhoToTest)
                End Try
                Medscreen.LogTiming("    Fill Form Phrases")

                'Set which of the items is selected set of phrase based variables, due to early design errors the phrase text is use dinstead of the phrase id
                Try
                    Dim objPhrase As MedscreenLib.Glossary.Phrase
                    'Dim i As Integer
                    Dim I1 As Integer
                    Me.cbConfType.SelectedValue = .ConfirmationType
                    'If .ConfirmationType.Trim.Length > 0 Then
                    '    For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Count - 1
                    '        objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Current()
                    '        If objPhrase.PhraseID = .ConfirmationType Or objPhrase.PhraseText = .ConfirmationType Then
                    '            Exit For
                    '        End If
                    '        Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Position += 1
                    '    Next
                    'End If

                    If .Tests.Trim.Length > 0 Then
                        Me.cbTestType.SelectedValue = .Tests
                        'Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Position = 0
                        'For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Count - 1
                        '    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Current()
                        '    If objPhrase.PhraseText = .Tests Then
                        '        Exit For
                        '    End If
                        '    Me.BindingContext(MedscreenLib.Glossary.Glossary.TestTypeList).Position += 1

                        'Next


                    End If


                    If .CollType.Length > 0 Then
                        'Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position = 0
                        'For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Count - 1
                        '    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Current()
                        '    If objPhrase.PhraseID = .CollType Or objPhrase.PhraseText = .CollType Then
                        '        Exit For
                        '    End If
                        '    Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position += 1
                        'Next
                        Me.cbRandom.SelectedValue = .CollType
                        'For Each objPhrase In MedscreenLib.Glossary.Glossary.ReasonForTestList
                        '    If objPhrase.PhraseID = .CollType Or objPhrase.PhraseText = .CollType Then
                        '        I1 = MedscreenLib.Glossary.Glossary.ReasonForTestList.IndexOf(objPhrase.PhraseID)
                        '        Me.cbRandom.SelectedIndex = I1
                        '        Exit For
                        '    End If
                        'Next
                    End If

                    If .ConfirmationMethod.Length > 0 Then
                        Me.cbConfMethod.SelectedValue = .ConfirmationMethod
                        'Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position = 0
                        'For I1 = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Count - 1
                        '    objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Current()
                        '    If objPhrase.PhraseID = .ConfirmationMethod Or objPhrase.PhraseText = .ConfirmationMethod Then
                        '        Exit For
                        '    End If
                        '    Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position += 1
                        'Next
                    End If
                Catch ex As Exception
                    LogError(ex, , "Fill Form Phrases")
                End Try

                Try
                    Me.cbConfMethod.SelectionLength = 0
                    Me.cbConfType.SelectionLength = 0
                    'Me.cbCustomer.SelectedIndex = objSibColl.IndexOf(myClient)

                    Me.txtPO.Text = .PurchaseOrder                      'Set purchase Order 
                    Me.txtComments.Text = .CollectionComments           'Set Comments about collection 
                    Me.ToolTip1.SetToolTip(Me.txtComments, .CollectionComments)
                Catch ex As Exception
                    LogError(ex, , "Fill Form Comments")
                End Try
                Try
                    'Is it a programme managed customer 
                    If .ProgrammeManaged Then
                        Me.cbPMAnaged.SelectedIndex = 0
                        Me.cbPManager.Show()
                        Me.lblBy.Show()
                    Else
                        Me.cbPMAnaged.SelectedIndex = 1
                        Me.cbPManager.Hide()
                        Me.lblBy.Hide()
                    End If

                    Me.cbPManager.SelectedValue = .ProgrammeManager

                    'Agents info 
                    Me.txtAgentName.Text = .Agent
                    Me.txtAgentphone.Text = .AgentsPhone
                Catch ex As Exception
                    LogError(ex, , "Fill Form Prog Man ")
                End Try
                'Customers comments 
                'Me.txtPrvComment.Text = .CustomerComment
                Me.txtAddDoc.Text = .Document                   'Additional documents 
                DrawDocuments()
                Try
                    If Me.txtAddDoc.Text.Trim.Length = 0 Then
                        Me.mnuEdAddDoc.Visible = False
                        Me.mnuEdAddDoc2.Visible = False
                        Me.lbllkAddDoc.Hide()
                    Else
                        Me.mnuEdAddDoc.Visible = True
                        Me.mnuEdAddDoc2.Visible = True
                        Me.lbllkAddDoc.Show()
                        Me.lbllkAddDoc.Text = "Click here to view Doc"
                        Me.lbllkAddDoc.Links.Clear()
                        Dim strFilenamet As String = Me.txtAddDoc.Text
                        If InStr(strFilenamet, ".") = 0 Then strFilenamet += ".doc"
                        Me.lbllkAddDoc.Links.Add(6, 4, Me.strFileDir & "\" & strFilenamet)
                    End If
                    Me.TxtConfrecip.Text = .ConfirmationContact
                    Me.txtConfAdddress.Text = .ConfirmationNumber
                Catch ex As Exception
                    LogError(ex, , "Fill Form Additional Doc")
                End Try
                'See if we can't display the site address.
                'Me.txtOther.Text = ""
                If .Addressid > 0 Then
                    Try

                        Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(.Addressid, Address.AddressCollection.ItemType.AddressId)
                        cmdSiteAddress.Enabled = Not (cadr Is Nothing)
                        If Not cadr Is Nothing Then
                            'Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
                        End If
                    Catch ex As Exception
                        LogError(ex, , "Fill Form Address Id" & .Addressid)
                    End Try
                Else
                    Try
                        Me.ToolTip1.SetToolTip(Me.txtLocation, "")
                        If Not objPort Is Nothing Then
                            Me.ToolTip1.SetToolTip(Me.txtLocation, objPort.Info2)
                            Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(objPort.AddressId)
                            cmdSiteAddress.Enabled = (cadr Is Nothing)
                            If Not cadr Is Nothing Then
                                'cmdSiteAddress.Enabled = (cadr Is Nothing)
                                'Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
                            End If

                        End If
                        If Not myCollection.VesselObject Is Nothing Then 'AndAlso myCollection.VesselObject.Invoicing Is Nothing Then
                            cmdSiteAddress.Enabled = True
                        End If
                    Catch ex As Exception
                        LogError(ex, , "Fill Form port site")
                    End Try
                End If


                Try


                    If Me.myCollection.HasID = 0 Then                      'No collection yet defined
                        'Me.cbCustomer.Enabled = True    'Only allow the customer to be set for new customers
                        'Me.SmidCombo1.Enabled = True
                        Me.NoDonorsActual.SelectedIndex = 0
                        Me.cBLocality.Enabled = True
                        Me.cmdGenerateRandom.Enabled = False
                    Else 'Editing collection
                        SetTitle()
                        Me.cmdGenerateRandom.Enabled = True
                        If InStr("AV", .Status) > 0 Then
                            'Me.cbCustomer.Enabled = True    'Only allow the customer to be set for new customers
                            'Me.SmidCombo1.Enabled = True
                            Me.cBLocality.Enabled = True
                            Me.FormMode = FormModes.Edit
                        Else
                            'Me.cbCustomer.Enabled = False
                            Me.cBLocality.Enabled = False
                            'Me.SmidCombo1.Enabled = False
                        End If
                        SetFormMode()
                        If .ActualNoOfDonors = "A" Then
                            Me.NoDonorsActual.SelectedIndex = 2
                        ElseIf .ActualNoOfDonors = "E" Then
                            Me.NoDonorsActual.SelectedIndex = 1
                        Else
                            Me.NoDonorsActual.SelectedIndex = 0
                        End If

                    End If
                Catch ex As Exception
                    LogError(ex, , "Fill Form id = 0 ")
                End Try
                Try
                    If .TestScheduleID.Trim.Length = 0 Then
                        Me.cbSampleTemplate.SelectedValue = Me.myClient.ReportingInfo.SampleTemplate()
                    Else
                        Me.cbSampleTemplate.SelectedValue = .TestScheduleID
                    End If
                    'deal with benezene testing 
                    Me.ckBenzene.Checked = .HasBenzene
                Catch ex As Exception
                    LogError(ex, , "Fill Form Test PAnel")
                End Try

                Me.cbAnalyticalLab.SelectedValue = .LaboratoryId

                Try
                    'lock vessel if status <> " " or "V"
                    If (.Status.Trim.Length = 0 Or .Status = "V" _
                        Or .Status = "P" Or .Status = "C") _
                        And (Me.MyMode = FormModes.Edit Or Me.MyMode = FormModes.Create) Then
                        Me.cbVessel.Enabled = True
                    Else
                        Me.cbVessel.Enabled = False
                    End If
                Catch ex As Exception
                    LogError(ex, , "Fill Form Vessel")
                End Try
            End With
            FillContacts()
            Dim strfilename As String = GetRandomFilename()
            If IO.File.Exists(strfilename) Then
                Me.cmdEditDoc.Enabled = True
                Me.cmdSendRandom.Enabled = True
            Else
                Me.cmdEditDoc.Enabled = False
                Me.cmdSendRandom.Enabled = False
            End If

        Catch ex As Exception

            MsgBox("Error in Fill Form - " & ex.ToString & " Please report to IT!", MsgBoxStyle.OkOnly)
            LogError(ex, , "Fill Form Overall error not specifiacally trapped")
        End Try
        blnFormSet = True

        Medscreen.LogTiming("Fill Form Finish")
    End Sub

    Private Sub FillContacts()
        Try
            If myClient Is Nothing Then Exit Sub
            Dim oColl As New Collection()
            oColl.Add(Me.myClient.Identity)
            oColl.Add("COLLCONF")
            Dim strContList As String = CConnection.PackageStringList("Lib_contacts.CustomerContactList", oColl)
            Dim strXML As String = "<Contacts>"
            If Not strContList Is Nothing AndAlso strContList.Trim.Length > 0 Then
                Dim strContArray As String() = strContList.Split(New Char() {","})
                Dim i As Integer
                For i = 0 To strContArray.Length - 1
                    If strContArray(i).Trim.Length > 0 Then
                        Dim strContXML As String = CConnection.PackageStringList("Lib_contacts.ContactToXML", strContArray(i))
                        strXML += strContXML
                    End If
                Next
            End If
            strXML += "</Contacts>"

            Dim strRTF As String = ResolveStyleSheet(strXML, "ContactsRTF.XSL", 0)
            Me.rtbContacts.Rtf = strRTF
        Catch ex As Exception
        End Try
    End Sub


    Private Sub GetSiteAddress(ByVal AddressID)
        If AddressID < 1 Then Exit Sub
        Dim cadr As MedscreenLib.Address.Caddress = _
    cSupport.SMAddresses.item(AddressID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
        cmdSiteAddress.Enabled = Not (cadr Is Nothing)
        If Not cadr Is Nothing Then     'got one display it 
            Me.Panel2.Show()
            Me.txtAddress.Text = cadr.BlockAddress(vbCrLf, True)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set form title up 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetTitle()
        Try
            If Not Me.myCollection Is Nothing Then
                With myCollection
                    Me.Text = "Collection Form - " & .ID & " - " & .VesselName & " - " & .SmClient.Info
                End With
            End If
        Catch ex As Exception
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 'Handle the label invoice double click event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lblInvoice_dblClick(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim frmHTML As New frmHTMLView()
        frmHTML.Document = Me.Collection.GetInvoiceInfo(myClient)
        frmHTML.ShowDialog()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle the exit menu event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle the view invoice menu event and context menu event 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuViewInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewInvoice.Click, mnuViewInvoice2.Click
        lblInvoice_dblClick(Nothing, Nothing)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle the browse button click event 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        AddDoc()

    End Sub
    Dim strDirectoryPath As String = ""
    Private tmpReference As String = ""
    Private Sub AddDoc()
        With OpenFileDialog1
            .InitialDirectory = strFileDir
            .FileName = Me.txtAddDoc.Text              ' Set the filename to the current one
            .Filter = "Word (*.doc)|*.doc|Excel|*.xls|Portable Document|*.pdf|All files|*.*"      ' Only allow selection of ".doc" files
            '.CancelErro = False                     ' Don't raise error if "cancel" clicked
            .Title = "Select a scanned document to attach"
            ' Flags:  Single file selection.  File must exist.  Can't change folder
            '         Use explorer style windows.  Path (folder) must exist
            '.Flag = cdlOFNAllowMultiselect + cdlOFNFileMustExist + cdlOFNNoChangeDir + _
            '         cdlOFNExplorer + cdlOFNPathMustExist
            If .ShowDialog() = DialogResult.OK Then ' Show the "Open file" dialog box
                Dim FilenameList As String() = .FileNames
                Dim i As Integer
                Dim stName As String = ""
                'Convert list of file names to comma delimited list

                strDirectoryPath = MedscreenLib.MedConnection.Instance.ServerPath & "\DM_DataFile\CollDocs\"
                If Me.Collection.HasID = 0 Then 'new collection create temporary directory store in .documents
                    If Not IO.Directory.Exists(strDirectoryPath & tmpReference) Then
                        IO.Directory.CreateDirectory(strDirectoryPath & tmpReference)
                    End If
                Else
                    'tmpReference = myCollection.ID
                    If Not IO.Directory.Exists(strDirectoryPath & myCollection.ID) Then
                        IO.Directory.CreateDirectory(strDirectoryPath & myCollection.ID)
                    End If
                End If
                For i = 0 To FilenameList.Length - 1
                    'If stName.Trim.Length > 0 Then stName += ","
                    stName = IO.Path.GetFileName(FilenameList(i))

                    'check the file name for ampersands
                    If InStr(stName, "&") * InStr(stName, "'") = 0 Then

                        Me.lvDocuments.Items.Add(New ListViewItem(stName))
                        'we need to put a copy of the document in a folder of serverpath\dm_datafile\CollDocs\cm number
                        IO.File.Copy(FilenameList(i), strDirectoryPath & tmpReference & "\" & stName)
                        Intranet.Support.AdditionalDocumentColl.CreateCollectionDocument(tmpReference, stName, "INSTRCTOIN")

                    Else
                        MsgBox("We can't support file names with an ampersand '&' or apostrophe ' in it please rename the file - " & stName, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                Next

            End If
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle the edit additional doc event 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuEdAddDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdAddDoc.Click, mnuEdAddDoc2.Click
        'Dim objWord As WordSupport.WordSupport
        Dim frmHTM As New frmHTMLView()
        Dim strHtml As String
        Dim strfilename As String = Me.txtAddDoc.Text
        If InStr(strfilename, ".") = 0 Then strfilename += ".doc"
        strHtml = "<a href=""" & Me.strFileDir & "\" & strfilename & _
         """>Scanned Document " & Me.txtAddDoc.Text & " </a><br>"
        frmHTM.Document = strHtml
        frmHTM.ShowDialog()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Try and disp[lay additonal documentation
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lbllkAddDoc_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lbllkAddDoc.LinkClicked
        ' Determine which link was clicked within the LinkLabel.
        Me.lbllkAddDoc.Links(lbllkAddDoc.Links.IndexOf(e.Link)).Visited = True
        ' Display the appropriate link based on the value of the LinkData property of the Link object.
        Try
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString())
        Catch ex As Exception
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Try and display invoice
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lblInvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Me.lblInvoice.Links(lblInvoice.Links.IndexOf(e.Link)).Visited = True
        lblInvoice_dblClick(Nothing, Nothing)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieve the conetents of form and validate 
    ''' </summary>
    ''' <returns>TRUE if no errors</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetForm() As Boolean
        Me.ErrorProvider1.SetError(Me.cbSampleTemplate, "")
        Dim blnRet As Boolean = True
        Try
            With myCollection

                'If we don't have client info (should never occur)
                'Addtional Costs 
                Try
                    'There may be a currency symbol in the first place, in a try block in case it goes do lalhi
                    If Me.txtAddCosts.Text.Chars(0) < "0" Or Me.txtAddCosts.Text.Chars(0) > "9" Then
                        .AgreedExpenses = Mid(Me.txtAddCosts.Text, 2)
                    ElseIf InStr(txtAddCosts.Text.ToUpper, "KR") > 0 Then
                        Dim amountArray As String() = txtAddCosts.Text.Trim.Split(New Char() {" "})
                        Dim strCosts As String = amountArray(0).Replace(".", "")
                        strCosts = strCosts.Replace(",", ".")

                        .AgreedExpenses = CDbl(strCosts)
                    Else
                        .AgreedExpenses = Me.txtAddCosts.Text
                    End If
                Catch ex As Exception
                    .AgreedExpenses = 0
                End Try

                'See if the currency selected is the same as the customer's if not convert
                If Me.cbCurrency.Text <> Me.CustomerCurrency Then
                    Dim strin As String
                    Dim strout As String

                    'No match set up for conversion, conversion use ISO three character names.
                    strin = Me.CustomerCurrency
                    'Deal with collecting officer currencies
                    strout = Me.cbCurrency.SelectedValue
                    'Use conversion routines and current conversion rate, 
                    'if any errors will almost certainly be caused by the conversion rate not being in the table
                    .AgreedExpenses = MedscreenLib.Glossary.Glossary.ExchangeRates.Convert(strout, strin, .AgreedExpenses)
                End If


                'Check for a customer site collection first 
                If Me.cbSiteVessel.SelectedIndex = 0 Then
                    .IsSiteVisit = True     'Ensure that it is listed as a site collection
                    '.vessel = Me.myClient.CompanyName 'Set vessel_name field to customer's company name
                    'TODO see if we have a customer site
                    'See if there is a customer site selected index > -1 
                    If Me.blnCustSiteSelected Then
                        'Ensure that it is a valid index
                        If Me.cbCustSites.SelectedIndex <= Me.cbCustSites.DataSource.count - 1 Then
                            'Get details of the customer site into a customer site object 
                            Dim custSite As Intranet.intranet.Address.CustomerSite = _
                                Me.cbCustSites.DataSource.item(Me.cbCustSites.SelectedIndex)
                            .vessel = custSite.VesselID         'Set Vessel_name field to site ID
                            .Addressid = custSite.AddressId     'Set addressID to site Address_id
                        End If
                    End If
                Else    'Maritime collection has a vessel
                    .IsVessel = True    'Set Site_vessel field to Vessel
                    Dim objVessel As Intranet.intranet.customerns.Vessel
                    If Me.cbVessel.SelectedIndex > -1 Then ' check we have selected something 
                        objVessel = Me.cbVessel.SelectedItem
                        If Not objVessel Is Nothing Then
                            myCollection.VesselObject = objVessel
                            If objVessel.VesselID.Trim.Length > 0 Then      ' Check is valid vessel not dummy one added to clear combo
                                .vessel = objVessel.VesselID                 'Set the vessel 
                            Else
                                'Set Vessel Error
                                Me.ErrorProvider1.SetError(Me.cbVessel, "You must select a vessel")
                                blnRet = False      'Indicate that there is an error
                            End If
                        End If
                    End If
                End If

                'Find out the location of the collection 
                Dim objPort As Intranet.intranet.jobs.Port
                If txtLocation.Text.Length > 0 Then
                    'We have a port selected this can be from the UK list or the Overseas list 
                    'dependent on the setting of the locality combo box
                    objPort = Me.myCollection.CollectionPort

                    'Check to see if we have a valid port 
                    If Not objPort Is Nothing Then
                        If objPort.PortId.Trim.Length = 0 Then  'Check it is not a dummy port
                            If (.CollectionType <> GCST_Type_External) Then
                                Me.ErrorProvider1.SetError(Me.txtLocation, "You must select a port")
                                blnRet = False
                            End If
                        Else            ' should be a valid port 
                            .Port = objPort.Identity    'Set Port Field
                            If Not (.CollectionType = GCST_Type_External Or .CollectionType = CMJob.GCST_Type_FixedSite Or .CollectionType = CMJob.GCST_Type_Referred) Then    'We are going to deal with External collections as well
                                If objPort.IsUK Then        'Check for UK or Overseas
                                    .CollectionType = CMJob.GCST_Type_UKRoutine   'Set collection type to UK or Overseas based on the Port location
                                    .UK = True
                                Else
                                    .CollectionType = CMJob.GCST_Type_OSRoutine
                                    .UK = False
                                End If
                            End If
                        End If
                    Else    'We have an error 
                        Me.ErrorProvider1.SetError(Me.txtLocation, "You must select a port")
                        blnRet = False
                    End If
                Else    'We have an error 
                    If (.CollectionType <> GCST_Type_External) Then 'We don't need a port for an external collection
                        Me.ErrorProvider1.SetError(Me.txtLocation, "You must select a port")
                        blnRet = False
                    End If
                End If
                'sET COLLECTION DATE and check its validity 
                .CollectionDate = Me.DTCollDate.Value       'Set the collection date 
                'Check collection dates for validity use compareto function 
                If Today.CompareTo(.CollectionDate) = 1 Then
                    Dim iRet As MsgBoxResult = MsgBox("This date is in the past, is this correct?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If iRet = MsgBoxResult.No Then
                        blnRet = False
                        Me.ErrorProvider1.SetError(Me.DTCollDate, "Date in the past")
                    End If
                    'use date parse to truncate date to day 
                ElseIf Today.CompareTo(Date.Parse(.CollectionDate.ToString("dd-MMM-yyyy")).AddDays(-2)) = 1 Then
                    Dim iRet As MsgBoxResult = MsgBox("There is a minimum of 48 hours notice for booking a collection, Advise the customer that there may be a surcharge.  Do they wish to proceed?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If iRet = MsgBoxResult.No Then
                        blnRet = False
                        Me.ErrorProvider1.SetError(Me.DTCollDate, "Collection for Today")
                    End If
                ElseIf Today.CompareTo(Date.Parse(.CollectionDate.ToString("dd-MMM-yyyy"))) = 0 Then
                    Dim iRet As MsgBoxResult = MsgBox("This collection is for today, is this correct?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If iRet = MsgBoxResult.No Then
                        blnRet = False
                        Me.ErrorProvider1.SetError(Me.DTCollDate, "Collection for Today")
                    End If
                End If
                'Set boring information in the collection 
                .Announced = (Me.cbAnnouced.SelectedIndex = 1)
                .SpecialProcedures = (Me.cbBlnProc.SelectedIndex = 0)

                'Do a check for donors being > 0 
                .NoOfDonors = Me.updnNoDonors.Value         'Set the expected number of Donors


                'This could alternatively be done using port.isuk 
                .UK = (Me.cBLocality.SelectedIndex = 0)


                .WhoToTest = Me.txtWhototest.Text            'Indcate Who to Test 
                .ConfirmationNumber = Me.txtConfAdddress.Text
                .ConfirmationContact = Me.TxtConfrecip.Text

                'Set which of the items is selected 
                Dim objPhrase As MedscreenLib.Glossary.Phrase
                'Confirmation type 
                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "")
                'If Me.cbConfType.SelectedIndex > -1 Then
                '    objPhrase = MedscreenLib.Glossary.Glossary.ConfTypeList.Item(Me.cbConfType.SelectedIndex)
                '    If Not objPhrase Is Nothing Then
                '        .ConfirmationType = objPhrase.PhraseText
                '        If .ConfirmationType = "" Then .ConfirmationType = "NONE"
                '    End If
                'End If
                'If confirmation type not = N we must have a confirmationaddress.
                'If .ConfirmationType <> "NONE" Then
                '    If .ConfirmationNumber.Trim.Length = 0 Then 'We can't do this
                '        Me.ErrorProvider1.SetError(Me.txtConfAdddress, "You must either provide confirmation details, or set to no confirmations")
                '        blnRet = False
                '    End If
                'End If

                'Type of collection 
                Me.ErrorProvider1.SetError(Me.cbRandom, "")
                If cbRandom.SelectedIndex > -1 Then
                    .CollType = cbRandom.SelectedValue
                    If .CollType.Trim.Length = 0 Then
                        Me.ErrorProvider1.SetError(Me.cbRandom, "A reason for test must be given")
                        blnRet = False
                        Application.DoEvents()
                    End If
                Else
                    Me.ErrorProvider1.SetError(Me.cbRandom, "A reason for test must be given")
                    blnRet = False
                End If

                'Confirmation Method 
                'Me.ErrorProvider1.SetError(Me.txtConfAdddress, "")
                'If cbConfMethod.SelectedIndex > -1 Then
                '    objPhrase = MedscreenLib.Glossary.Glossary.ConfMethodList.Item(cbConfMethod.SelectedIndex)
                '    If Not objPhrase Is Nothing Then
                '        .ConfirmationMethod = objPhrase.PhraseText
                '        If .ConfirmationMethod = "" Then .ConfirmationMethod = "N"
                '    End If
                'Else

                'End If
                'Try
                '    If .ConfirmationType <> "NONE" AndAlso .ConfirmationMethod.Length > 0 AndAlso .ConfirmationMethod.Chars(0) <> "N" Then
                '        'We need to establish that the confirmation address is valid for the type of the method
                '        If .ConfirmationMethod.Chars(0) = "E" Then
                '            'email()
                '            Dim eError As MedscreenLib.Constants.EmailErrors = ValidateEmail(.ConfirmationNumber)
                '            If eError = EmailErrors.NoAt Then
                '                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "No @ sign in email address")
                '                txtConfAdddress.Focus()
                '                blnRet = False

                '            ElseIf eError = EmailErrors.NoDomain Then
                '                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "No Domain in email address")
                '                blnRet = False
                '                txtConfAdddress.Focus()

                '            ElseIf eError = EmailErrors.NoAddress Then
                '                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Address is missing")
                '                blnRet = False
                '                txtConfAdddress.Focus()

                '            ElseIf eError <> EmailErrors.None Then
                '                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Invalid email address")
                '                blnRet = False
                '                txtConfAdddress.Focus()
                '            End If

                '        End If
                '        If .ConfirmationMethod.Chars(0) = "F" Then
                '            If Not IsNumber(.ConfirmationNumber) Or .ConfirmationNumber.Length = 0 Then
                '                Me.ErrorProvider1.SetError(Me.txtConfAdddress, "This is not a valid fax number")
                '                blnRet = False
                '                txtConfAdddress.Focus()
                '            End If
                '        End If
                '    End If
                'Catch ex As Exception
                '    LogError(ex, , "Getting confirmation method")
                'End Try
                'Test types 
                If cbTestType.SelectedIndex > -1 Then
                    objPhrase = MedscreenLib.Glossary.Glossary.TestTypeList.Item(cbTestType.SelectedIndex)
                    If Not objPhrase Is Nothing Then
                        .Tests = objPhrase.PhraseText
                    End If
                End If


                .PurchaseOrder = Me.txtPO.Text                       'Set purchase Order 
                'Check to see if a PO number is a user requirement 
                Me.ErrorProvider1.SetError(Me.txtPO, "")
                If Not Me.myClient.InvoiceInfo.CheckPONumber(.PurchaseOrder) Then
                    Me.ErrorProvider1.SetError(Me.txtPO, "Purchase Order doesn't match Customer requirements")
                    blnRet = False
                End If

                'Collection comments 
                .CollectionComments = Me.txtComments.Text          'Set Comments about collection 

                'Is it a programme managed customer 
                .ProgrammeManaged = (Me.cbPMAnaged.SelectedIndex = 0)
                If .ProgrammeManaged Then
                    'See if the user wants to change the programme manager for this customer, if so do it
                    .ProgrammeManager = Me.cbPManager.SelectedValue
                    If Not myClient Is Nothing AndAlso myClient.ProgrammeManager <> Me.cbPManager.SelectedValue Then
                        myClient.ProgrammeManager = Me.cbPManager.SelectedValue
                        myClient.ReportingInfo.Update()
                    End If
                End If
                'Agents info 
                .Agent = Me.txtAgentName.Text
                .AgentsPhone = Me.txtAgentphone.Text

                'Customers comments 
                '.CustomerComment = Me.txtPrvComment.Text
                .Document = Me.txtAddDoc.Text                    'Additional documents 



                .ActualNoOfDonors = Me.NoDonorsActual.Text.Chars(0)

                'Test panel, this needs more work when test panel handling is agreed.
                'Dim objTp As MedscreenLib.Glossary.Phrase = objSampleTemplates.Item(Me.cbSampleTemplate.SelectedIndex)
                'If Not objTp Is Nothing Then
                .TestScheduleID = Me.cbSampleTemplate.SelectedValue
                'End If
                'Check that we have a sample template 
                If .TestScheduleID.Trim.Length = 0 Then
                    Me.ErrorProvider1.SetError(Me.cbSampleTemplate, "A sample template must be provided")
                    blnRet = False
                End If

                'Deal with benzene testing flag 
                If Me.ckBenzene.Checked Then
                    .TestPanelID = "BENZENE"
                Else
                    .TestPanelID = ""
                End If
                .LaboratoryId = Me.cbAnalyticalLab.SelectedValue
                'check for site and overseas
                If Not .UK And .IsSiteVisit Then
                    MsgBox("This collection for a site overseas is not a normal situation, please check!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                End If
            End With
        Catch ex As Exception
            LogError(ex, , "Get form")
        End Try
        blnRet = blnRet And ValidateData() And ValidateDataExit()
        Return blnRet

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if there are any other errors in play 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/01/2006]</date><Action>Reason for test checks added to exit checks</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action>Vessel checks added to exit checks</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function ValidateData() As Boolean
        'check to see if the customer is maritime indicated by industry type that it is a vessel collection 
        Dim iRet As MsgBoxResult
        Dim blnret As Boolean = True


        If Me.myClient Is Nothing Then          'If no valid client bog off
            Return blnret
            Exit Function
        End If

        If Me.myCollection Is Nothing Then      'If no collection bog off 
            Return blnret
            Exit Function
        End If

        If Me.Collection.CollectionType = GCST_Type_External OrElse Me.Collection.CollectionType = Chr(Me.Collection.CollectionTypes.GCST_JobTypeFixed) Then       'External collections are less fastidious
            Return blnret
            Exit Function
        End If
        Me.ErrorProvider1.SetError(Me.cbRandom, "")
        Me.ErrorProvider1.SetError(Me.DTCollDate, "")

        'check for fixed sites 
        If Me.Collection.Port.Trim.Length = 0 Then
            blnret = False
            Me.ErrorProvider1.SetError(Me.txtLocation, "No port set")
        Else
            If Me.Collection.CollectionPort.IsFixedSite Then
                blnret = False
                Me.ErrorProvider1.SetError(Me.txtLocation, "Fixed site collections need to be booked using a different utility")
            Else
                Me.ErrorProvider1.SetError(Me.txtLocation, "")
            End If
        End If

        'validate number of donors 
        If myCollection.NoOfDonors = 0 Then
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "No of Donors must be greater than 1 ")
            blnret = False
        Else
            Me.ErrorProvider1.SetError(Me.updnNoDonors, "")
        End If

        'Check that the user has set the number of donors to actual or estimated
        If Me.NoDonorsActual.Text.Trim.Length = 1 Then
            Me.ErrorProvider1.SetError(Me.NoDonorsActual, "You must define whether the number of donors is known or an estimate")
            blnret = False
        Else
            Me.ErrorProvider1.SetError(Me.NoDonorsActual, "")
        End If

        'check for purchase order number
        If myClient.InvoiceInfo.PoRequired Then
            If myCollection.PurchaseOrder.Trim.Length = 0 Then
                Me.ErrorProvider1.SetError(Me.txtPO, "Purchase Order Required for this customer")
                blnret = False
            End If
        End If

        'check for stops on collection 
        With myCollection
            'Check to see if there is a VDI based stop on the collection check there is a 
            'Valid vessel object and it has an invoicing object and it is present
            'Then see if the stop collections flag is set
            If Not .VesselObject Is Nothing Then
                '    Not .VesselObject.Invoicing Is Nothing _
                '    AndAlso .VesselObject.Invoicing.Loaded Then
                If .VesselObject.StopCollections Then
                    Me.ErrorProvider1.SetError(Me.cbVessel, "There is a stop on this vessel this collection cannot proceed")
                    blnret = False
                End If
            End If

            'It may be there is a site stop 
            If Not .SiteObject Is Nothing Then
                '  Not .SiteObject.Invoicing Is Nothing _
                '  AndAlso .SiteObject.Invoicing.Loaded Then
                If .SiteObject.StopCollections Then
                    Me.ErrorProvider1.SetError(Me.cbCustSites, "There is a stop on this collection site this collection cannot proceed")
                    blnret = False
                End If
            End If


            'Check for a reason for test 
            If .ReasonForTest.Trim.Length = 0 Then
                '    Me.ErrorProvider1.SetError(Me.cbRandom, "A reason for test must be provided.")
                '    blnret = False
            End If
            Me.ErrorProvider1.SetError(Me.DTCollDate, "")
            '.CollectionDate = Me.DTCollDate.Value       'Set the collection date 
            'Check collection dates for validity use compareto function 
            If Collection.HasID = 0 Then
                If Today.CompareTo(Me.DTCollDate.Value) = 1 Then
                    blnret = True
                    Me.ErrorProvider1.SetError(Me.DTCollDate, "Date in the past")
                    'use date parse to truncate date to day 
                ElseIf Today.CompareTo(Date.Parse(Me.DTCollDate.Value.ToString("dd-MMM-yyyy")).AddDays(-2)) = 1 Then
                    blnret = True
                    Me.ErrorProvider1.SetError(Me.DTCollDate, "There is a minimum of 48 hours notice for booking a collection, Advise the customer that there may be a surcharge.")

                ElseIf Today.CompareTo(Date.Parse(Me.DTCollDate.Value.ToString("dd-MMM-yyyy"))) = 0 Then
                    blnret = True
                    Me.ErrorProvider1.SetError(Me.DTCollDate, "This collection is for today, is this correct")
                End If
            End If
        End With
        SetTitle()
        Return blnret
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exit only checks 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function ValidateDataExit() As Boolean
        Dim iRet As MsgBoxResult
        Dim blnret As Boolean = True

        If Me.Collection.CollectionType = GCST_Type_External Or Me.Collection.CollectionType = Chr(Me.Collection.CollectionTypes.GCST_JobTypeFixed) Then       'External collections are less fastidious
            Return blnret
            Exit Function
        End If

        Me.ErrorProvider1.SetError(Me.cbVessel, "")
        If Mid(Me.myClient.IndustryCode, 1, 3) = "MAR" Then
            If Not Me.Collection.IsVessel Then
                iRet = MsgBox("This collection appears to be for a maritime customer but is a site collection is this correct?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                blnret = (iRet = MsgBoxResult.Yes)
                Me.ErrorProvider1.SetError(Me.cbSiteVessel, "Maritime customer please check")
            End If

            If Me.Collection.VesselObject Is Nothing Then
                iRet = MsgBox("This collection appears to be for a maritime customer but no vessel has been selected, is this correct?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                blnret = (iRet = MsgBoxResult.Yes)
                Me.ErrorProvider1.SetError(Me.cbVessel, "Maritime customer please check")
            End If
        End If

        Return blnret

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update collection information and return 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        If Not GetForm() Then 'Check data is valid, if there are any issues raised 
            'Change the form return to none, this will make it return to displaying and the form will 
            'not close 
            Me.DialogResult = DialogResult.None
            Exit Sub
        End If

        Dim blnNew As Boolean = False


        If Collection.HasID = 0 Then        'Indicates a new collection
            blnNew = True

            'need to add checks for exotic collection types 
            Dim oColl As New Collection()
            oColl.Add(Collection.TestScheduleID)
            oColl.Add(Me.myClient.Identity)
            Dim strCheckReturn As String = CConnection.PackageStringList("Lib_cctool.NoLocalTests", oColl)
            If strCheckReturn = "T" Then 'No tests to be done locally 
                Collection.CollectionType = CMJob.GCST_Type_Referred
            Else                                                              'Need to see if it was referred
                If Collection.CollectionType = CMJob.GCST_Type_Referred Then  'If so we need switch back to UK or OS
                    If Collection.UK Then
                        Collection.CollectionType = CMJob.GCST_Type_UKRoutine
                    Else
                        Collection.CollectionType = CMJob.GCST_Type_OSRoutine
                    End If
                End If
            End If

            Dim Skip As Boolean = False
            Dim strCheckmessage As String = "Do you wish to create this collection for : " & Me.myClient.ClientID & " - " & _
                Me.myCollection.VesselName & " - at " & Me.myCollection.PortName & "?"
            If MsgBox(strCheckmessage, MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                CreateCollection()                                  'Actually create the collection
                'move (rename) collection documents folder
                cmdGenerateRandom.Enabled = True
                If strDirectoryPath.Trim.Length > 0 Then
                    IO.Directory.Move(Me.strDirectoryPath & Me.tmpReference, Me.strDirectoryPath & myCollection.ID)
                End If
                Intranet.Support.AdditionalDocumentColl.ChangeDocReference(tmpReference, myCollection.ID)
                'Change the document references.
                myCollection.Document = ""
                Dim strXML As String = myCollection.ToXML(Intranet.intranet.Support.cSupport.XMLCollectionOptions.Defaults)
                Dim strHTML1 As String
                If Not myCollection.SiteObject Is Nothing Then
                    strHTML1 = Medscreen.ResolveStyleSheet(strXML, "CollSiteRequest.xsl", 0)
                Else
                    strHTML1 = Medscreen.ResolveStyleSheet(strXML, "CollRequest.xsl", 0)
                End If
                Dim dlgRet As DialogResult
                dlgRet = Medscreen.ShowHtml(strHTML1, "I confirm that I have checked the displayed collection and it is complete and without errors", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                If dlgRet = DialogResult.Yes Then
                    Me.AddCheckedComment()                          'Add comment to say collection has been checked
                    If Not myClient.InvoiceInfo.CreditCard_pay Then 'Do we do officer assignment now 
                        'only if not a credit card customer
                        myCollection.Update()

                        dlgRet = MsgBox("Do you wish to assign this collection : " & myCollection.ID & " to a collecting officer now?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                        If dlgRet = DialogResult.Yes Then   'Yes we do want to assign now
                            'Call global assignment method
                            CommonForms.AssignOfficer(Me.myCollection)
                            'AssignOfficer()
                        End If
                    Else
                        'Client may be required to pay by credit card, if so generate an cc pay form

                        '    If myClient.InvoiceInfo.CreditCard_pay Then

                        Dim strHTML As String = Me.JobCostHTML(myClient.HasStandardPricing, myClient)
                        MsgBox("This client must pay by credit card, the necessary forms will be produced " & vbCrLf & _
                                    " and the collection put on hold until the form is returned" & vbCrLf & _
                                    "Please advise the customer of this, Costs are on the following screen", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Credit card stop")
                        If Medscreen.ShowHtml(strHTML, "Costs correct", MsgBoxStyle.OkOnly) = DialogResult.OK Then

                            myCollection.Status = Constants.GCST_JobStatusOnHold
                            myCollection.PaymentType = PaymentTypes.PaymentCard
                            Dim strContact As String = Medscreen.GetParameter(MyTypes.typString, "Contact for Card Payment", "Contact", myClient.Contact)

                            Dim strFax As String
                            'If Not myClient.MainAddress Is Nothing Then
                            '    strFax = Medscreen.GetParameter(MyTypes.typString, "Fax number for Contact", "Fax number", myClient.MainAddress.Fax)
                            'Else
                            strFax = Medscreen.GetParameter(MyTypes.typString, "Fax number for Contact", "Fax number", "")

                            'End If

                            myCollection.CreditCardConfirmationDocument(strContact, strFax, dblCost)
                            myCollection.Comments.CreateComment("Confirmation document sent to customer " & myCollection.ClientId, MedscreenLib.Constants.GCST_COMM_CollConfirm)
                            Me.Collection.FieldValues.Update(CConnection.DbConnection, , True)
                        End If
                    End If
                    DialogResult = DialogResult.OK
                Else
                    DialogResult = DialogResult.None
                    Skip = True
                End If

                If Not Skip Then
                    Me.Collection.WriteXML(Medscreen.GCSTCollectionXML & Me.Collection.ID & ".XML", 0)
                    'Check to see if another collection for this customer 
                    If MsgBox("Do you wish to create another collection for this customer (y/n) ?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Collection.FieldValues.InitialiseOldValues() 'Set old values back to defaults to try and get a clean write.
                        Collection.SetId = "0000000000"     'Reset collection variables
                        Collection.CollectionDate = Today
                        Collection.vessel = ""
                        Collection.Port = ""
                        Collection.Agent = ""
                        Collection.AgentsPhone = ""
                        '<Added code modified 06-Aug-2007 07:35 by taylor> to stop document info being carried over
                        Collection.Document = ""

                        '</Added code modified 06-Aug-2007 07:35 by taylor> 
                        Me.FillForm()
                        Me.DialogResult = DialogResult.None
                    End If
                End If
            End If
        ElseIf Collection.HasID <> 0 Then
            myCollection.Update()
            Dim strXML As String = myCollection.ToXML(Intranet.intranet.Support.cSupport.XMLCollectionOptions.Defaults)
            Dim strHTML1 As String
            If Not myCollection.SiteObject Is Nothing Then
                strHTML1 = Medscreen.ResolveStyleSheet(strXML, "CollSiteRequest.xsl", 0)
            Else
                strHTML1 = Medscreen.ResolveStyleSheet(strXML, "CollRequest.xsl", 0)
            End If
            Dim dlgRet As DialogResult
            dlgRet = Medscreen.ShowHtml(strHTML1, "I confirm that I have checked the displayed collection and it is complete and without errors", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If dlgRet = DialogResult.Yes Then                                   'Is okay
                Me.AddCheckedComment()                                          'Confirm collection has been checked
                If Me.Collection.FieldValues.RowID.Trim.Length = 0 Then
                    Me.Collection.Status = "P"
                    Me.Collection.FieldValues.Insert(CConnection.DbConnection)
                    Me.Collection.ConfirmCollection()
                Else
                    Me.Collection.FieldValues.Update(CConnection.DbConnection, , True)

                End If
                Me.Collection.WriteXML(Medscreen.GCSTCollectionXML & Me.Collection.ID & ".XML", 0)
                DialogResult = DialogResult.OK
            Else
                DialogResult = DialogResult.None
            End If
        Else

            Me.Collection.FieldValues.Update(CConnection.DbConnection, , True)
            Me.Collection.WriteXML(Medscreen.GCSTCollectionXML & Me.Collection.ID & ".XML", 0)

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if there is a checked comment
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/10/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub AddCheckedComment()
        'First check the number of comments 
        Dim oColl As New Collection()
        oColl.Add(Me.myCollection.ID)
        oColl.Add("UCON")
        Dim strCount As String = CConnection.PackageStringList("Lib_collection.CountCollectionComments", oColl)
        If CInt(strCount) = 0 Then 'We can add the comment
            Me.myCollection.Comments.CreateComment("Collection checked and has been confirmed as being accurate", "UCON", CollectionCommentList.CommentType.Collection)
        Else
            Dim ocmd As New OleDb.OleDbCommand("Select rowid,c.* from collection_comments c where cm_job_number = ? and comment_type = ?", CConnection.DbConnection)
            Dim oRead As OleDb.OleDbDataReader
            Dim oComm As New CollectionComment(CollectionCommentList.CommentType.Collection)

            Try
                ocmd.Parameters.Add(CConnection.StringParameter("CM", Me.myCollection.ID, 10))
                ocmd.Parameters.Add(CConnection.StringParameter("TYPE", "UCON", 10))
                If CConnection.ConnOpen Then
                    oRead = ocmd.ExecuteReader
                    If oRead.Read Then
                        oComm.Fields.readfields(oRead)
                    End If
                End If
                oRead.Close()
                oComm.CommentText += ", Edited on " & Now.ToString & " by " & MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                oComm.Update()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub CreateCollection()
        Try
            Dim OISID As Integer
            'Call collections collection method to create a new collection
            Dim objC As Intranet.intranet.jobs.CMJob = cSupport.ICollection(False).CreateCollection(Me.Collection, OISID)
            If Not objC Is Nothing Then 'Check we managed to create one
                With myCollection
                    myCollection.SetId = objC.ID    'Move the ID of the created collection into this one
                    objC.Load()
                    myCollection.FieldValues.RowID = objC.FieldValues.RowID
                    myCollection.SetStatus(MedscreenLib.Constants.GCST_JobStatusCreated)
                    'Set collection status to V
                    'Ensure that we have a valid user in SM terms
                    If Not CurrentSMUser() Is Nothing Then
                        'Set booker equal to SM user Id
                        myCollection.Booker = CurrentSMUser.Identity
                    End If
                    'wait 200 msecs for things to settle down
                    System.Threading.Thread.CurrentThread.Sleep(200)
                    'Let the user know the ID
                    MsgBox("Collection ID: " & objC.ID & " Created.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    myCollection.Status = MedscreenLib.Constants.GCST_JobStatusCreated
                    .DateReceived = Now     'Set date booked to now
                    'Add a comment into the collection comments table to record it was created 
                    .Comments.CreateComment("Collection created for customer " & .ClientId, MedscreenLib.Constants.GCST_COMM_NEWColl)
                    myCollection.FieldValues.Loaded = True
                    myCollection.Update()       'Ensure it gets written

                    'For a vessel or customer site collection ensure that the job vessel id field matches the collection one

                    'Vessel Chaneg collection method moves the current collection into the last collection field 
                    'and sets this collection into the next field
                    If .IsVessel Then                           ' a vessel collection needs the vessel info sorted
                        If Not .VesselObject Is Nothing Then
                            .VesselObject.ChangeCollection(.ID)
                            If Not .Job Is Nothing Then
                                If .Job.VesselName <> .vessel Then
                                    .Job.VesselName = .vessel
                                    .Job.Update()
                                End If
                            End If
                        End If
                    Else
                        If Not .SiteObject Is Nothing Then
                            .SiteObject.ChangeCollection(.ID)
                        End If
                    End If

                    'remove temporary collection and add this one.
                    Dim Index As Integer
                    Index = cSupport.ICollection(False).IndexOf(objC)
                    If Index > -1 Then cSupport.ICollection(False).RemoveAt(Index)
                    cSupport.ICollection(False).Add(myCollection)

                    'Does this customer need a temporary collecting officer assigned
                    Dim objOption As Intranet.intranet.customerns.SmClientOption
                    objOption = myClient.Options.Item("ASSIGNTEMP")
                    'check to see if we need to assign temp officer, temporary officer ID is in 
                    'the option value property
                    If (Not objOption Is Nothing) AndAlso (Not myClient.InvoiceInfo.CreditCard_pay) Then
                        myCollection.CollOfficer = objOption.Value
                        myCollection.Status = Constants.GCST_JobStatusAssigned
                        myCollection.Update()
                        myCollection.DoPrint("")
                        myCollection.Update()
                    End If
                End With
            Else
                'no we've failed 
                Me.DialogResult = DialogResult.Retry   'Say we've failed
                Exit Sub
            End If
            myCollection.Update()
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with changing client, will involve getting the new client and then redisplaying 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    'Dim id As Integer = -1


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Customer selected has changed, have to redispaly and change defaults
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If Not blnFormLoaded Then Exit Sub
        'If Me.cbCustomer.SelectedIndex = -1 Then Exit Sub

        SetWaitCursor()
        Dim myColl As Intranet.intranet.customerns.CustColl

        Me.ErrorProvider1.SetError(Me.txtCustomer, "")

        Dim myCust As Intranet.intranet.customerns.Client = Me.myClient
        If myCust Is Nothing Then Exit Sub

        If Not (myCust.HasService(myCust.Services.GCST_ROUTINE) Or myCust.HasService(myCust.Services.GCST_EXTERNAL) Or myCust.HasService(myCust.Services.GCST_FIXEDSITE)) Then
            Me.ErrorProvider1.SetError(Me.txtCustomer, "Customer Account is not suitable for collections")
            MsgBox("This customer is not set-up for collections " & vbCrLf & _
                "Please select a different account for this customer", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Customer Account Type Error")

        End If
        'TODO Get the customer collection defaults setup collection with them
        'TODO Findout if customer has vessels if not then force to site collection disable combo
        '   otherwise set to vessel and get vessels into vessel- combo

        myColl = myCust.CollectionInfo()
        'id = -1
        'Try
        '    id = CInt(Collection.ID)
        'Catch
        'End Try
        If Not myColl Is Nothing AndAlso myCollection.HasID = 0 Then
            myCollection.ClientId = myCust.Identity
            myClient = cSupport.CustList.Item(myCollection.ClientId)
            With myColl

                If .NoDonors > 0 Then
                    If myCollection.NoOfDonors = 0 Then myCollection.NoOfDonors = .NoDonors
                End If
                If myCollection.ConfirmationContact.Trim.Length = 0 Then myCollection.ConfirmationContact = .ConfirmationContactList
                If myCollection.ConfirmationNumber.Trim.Length = 0 Then myCollection.ConfirmationNumber = .ConfirmationSendAddress
                If myCollection.ConfirmationMethod.Trim.Length = 0 Then myCollection.ConfirmationMethod = .ConfirmationMethod
                'Don't default reason for test -  requirement changed by Nathalie. WO 1205
                'If myCollection.CollType.Trim.Length > 0 Then myCollection.CollType = .ReasonForTest
                If myCollection.ConfirmationType.Trim.Length > 0 Then myCollection.ConfirmationType = .ConfirmationType
                myCollection.SpecialProcedures = .Procedures
                If myCollection.WhoToTest.Trim.Length = 0 Then myCollection.WhoToTest = .WhoToTest
                'myCollection.Addressid = 0
                If myCollection.CollectionComments.Trim.Length = 0 Then myCollection.CollectionComments = .AccessToSite
                Debug.WriteLine(.TestType)

                Dim objPhrase As MedscreenLib.Glossary.Phrase
                Dim I As Integer

                If .ReasonForTest.Length > 0 Then
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position = 0
                    For I = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Count - 1
                        objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Current()
                        If objPhrase.PhraseID = .ReasonForTest Then
                            Exit For
                        End If
                        Me.BindingContext(MedscreenLib.Glossary.Glossary.ReasonForTestList).Position += 1
                    Next
                End If

                If .ConfirmationMethod.Length > 0 Then
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position = 0
                    For I = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Count - 1
                        objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Current()
                        If objPhrase.PhraseID = .ConfirmationMethod Then
                            Exit For
                        End If
                        Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfMethodList).Position += 1
                    Next
                End If

                'Confirmation type 
                If .ConfirmationType.Length > 0 Then
                    Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Position = 0
                    For I = 0 To Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Count - 1
                        objPhrase = Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Current()
                        If objPhrase.PhraseID = .ConfirmationType Then
                            Exit For
                        End If
                        Me.BindingContext(MedscreenLib.Glossary.Glossary.ConfTypeList).Position += 1
                    Next
                End If

                'TestType 
                If .TestType.Length > 0 Then
                    Me.BindingContext(Me.cbTestType.DataSource).Position = 0
                    For I = 0 To Me.BindingContext(cbTestType.DataSource).Count - 1
                        objPhrase = Me.BindingContext(cbTestType.DataSource).Current()
                        If objPhrase.PhraseID = .TestType Then
                            Exit For
                        End If
                        Me.BindingContext(cbTestType.DataSource).Position += 1
                        Me.cbTestType.SelectedIndex = BindingContext(Me.cbTestType.DataSource).Position
                    Next
                End If

                If .SampleTemplate.Length > 0 Then
                    Me.myCollection.TestScheduleID = .SampleTemplate
                End If


                'If Not myCollection Is Nothing Then
                '    myCollection.ConfirmationNumber = .ConfirmationAddress
                '    myCollection.ConfirmationContact = .ConfirmationContact
                'End If


            End With
            'SetUpClient()
            FillForm()
        End If
        setDefaultCursor()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set wait cursor
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub SetWaitCursor()
        Me.Cursor = Cursors.WaitCursor
        'Me.Invalidate()
        'Application.DoEvents()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return to default cursor 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub setDefaultCursor()
        Me.Cursor = Cursors.Default
        'Me.Invalidate()
        'Application.DoEvents()
    End Sub

    Private Sub RebindVessels()
        Me.cbVessel.DataSource = myClient.Vessels
        Me.cbVessel.BindingContext(Me.cbVessel.DataSource).SuspendBinding()
        Me.cbVessel.BindingContext(Me.cbVessel.DataSource).ResumeBinding()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Code to add a new vessel
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddVessel.Click
        Dim ovessel As Intranet.intranet.customerns.Vessel

        MsgBox("The most succesful way to manage a fleet of vessels is from the tree view!", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly)
        Me.cmdAddVessel.Enabled = False
        ovessel = Intranet.intranet.Support.cSupport.AddNewVessel(myClient.Identity)
        If Not ovessel Is Nothing Then
            SetWaitCursor()
            'myClient.Vessels.Add(ovessel)
            'myClient.Vessels.Refresh()
            Me.myClient.Vessels.Sort()
            RebindVessels()
            'Me.cbVessel.Refresh()
            Me.cbVessel.SelectedValue = ovessel
            Me.FlipSites()
            Dim intI As Integer = myClient.Vessels.IndexOf(ovessel)
            If intI > -1 Then
                Me.cbVessel.SelectedIndex = intI
            End If
            If Not Me.Collection Is Nothing Then
                Me.Collection.IsVessel = True
                Me.Collection.vessel = ovessel.VesselID
            End If
        Else
            MsgBox("Problem with adding vessel!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
        End If
        setDefaultCursor()
        Me.cmdAddVessel.Enabled = True
    End Sub

    Private Sub RebindSites()
        Me.cbCustSites.DataSource = myClient.Sites
        Me.cbCustSites.BindingContext(cbCustSites).SuspendBinding()
        Me.cbCustSites.BindingContext(cbCustSites).ResumeBinding()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' Add a new customer site 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNewSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewSite.Click
        'TODO : Add new collection site code needed
        If myClient Is Nothing Then Exit Sub

        Me.cmdNewSite.Enabled = False
        Dim oSite As Intranet.intranet.Address.CustomerSite
        SetWaitCursor()
        oSite = myClient.Sites.Create(myClient.Identity)

        If oSite Is Nothing Then
            MsgBox("Problem adding site!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        myClient.Sites.Sort()
        oSite.VesselStatus = "ACTIVE"
        oSite.VeselNextTest = Today.AddYears(1)

        Dim edFrm As New MedscreenCommonGui.frmVessel()
        edFrm.Vessel = oSite

        If edFrm.ShowDialog = DialogResult.OK Then
            oSite = edFrm.Vessel
            oSite.Owner = myClient.Identity
            oSite.Update()

            myClient.Sites.Refresh()
            myClient.Sites.Sort()
            RebindSites()
            'Me.cbCustSites.Refresh()
            FlipSites()
            Dim intI As Integer = myClient.Sites.IndexOf(oSite)
            If intI > -1 Then
                Me.cbCustSites.SelectedIndex = intI
                'Dim cadr As MedscreenLib.Address.Caddress = _
                '    cSupport.SMAddresses.item(oSite.AddressID)
                cmdSiteAddress.Enabled = Not (oSite.Address Is Nothing)
                'If Not cadr Is Nothing Then
                '    'Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
                'End If
                Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(oSite.VesselSubType)
                If Not oPort Is Nothing Then
                    Me.txtLocation.Text = oPort.Info2
                    'Me.cbLocation.SelectedValue = oPort.Identity
                End If
            End If
        End If
        setDefaultCursor()
        Me.cmdNewSite.Enabled = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Change site display 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
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

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Location has changed, form needs to redisplayed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try  ' Protecting
            'If Me.cbLocation.DataSource Is Nothing Then Exit Sub
            'If Me.cbLocation.SelectedIndex = -1 Then Exit Sub
            If Me.Collection Is Nothing Then Exit Sub
            If Not blnFormLoaded Then Exit Sub

            SetWaitCursor()

            If Me.Collection.UK Then
                'Check we are not going passed index
                'If Me.cbLocation.SelectedIndex >= cSupport.UKPorts.Count Then Exit Sub
                'objPort = cSupport.UKPorts.Item(Me.cbLocation.SelectedIndex)
            Else
                'Check we are not going passed index
                'If Me.cbLocation.SelectedIndex >= cSupport.OSPorts.Count Then Exit Sub
                'objPort = cSupport.OSPorts.Item(Me.cbLocation.SelectedIndex)
            End If
            If Not objPort Is Nothing Then
                'If objPort.IsUK Then
                '    Me.cBLocality.SelectedIndex = 0
                'Else
                '    Me.cBLocality.SelectedIndex = 1
                'End If
                Me.Collection.Port = objPort.Identity
                If objPort.IncursExtracharge Then
                    'Me.cbLocation.BackColor = Drawing.Color.Red
                    'Me.ToolTip1.SetToolTip(Me.cbLocation, "**** Additional Charges ****" & objPort.Info)
                    Me.cmdPortCosts.Show()
                    Me.NWInfo.Notify(objPort.CollectingOfficers.ToText(myClient.CurrencyCode))
                    'If Not myCollection Is Nothing AndAlso myCollection.AgreedExpenses = 0 AndAlso myCollection.ID = 0 Then
                    '    If Not objPort.CollectingOfficers Is Nothing Then
                    '        Me.txtAddCosts.Text = objPort.CollectingOfficers.MinimumCharge
                    '    End If
                    'End If
                Else
                    Me.cmdPortCosts.Hide()

                    If Not myCollection Is Nothing AndAlso myCollection.AgreedExpenses = 0 AndAlso myCollection.HasID = 0 Then
                        Me.txtAddCosts.Text = ""
                    End If
                    Me.txtLocation.BackColor = Drawing.SystemColors.Window
                    Me.ToolTip1.SetToolTip(Me.txtLocation, objPort.Info)
                End If
            End If
            'Me.cbLocation.SelectionStart = 1
            'Me.cbLocation.SelectionLength = 0
            Me.ValidateData()
            setDefaultCursor()
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbLocation_SelectedIndexChanged-3675")
        Finally
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Selected site has changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbCustSites_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCustSites.SelectedIndexChanged
        If Not Me.blnFormLoaded Then Exit Sub
        If cbCustSites.SelectedIndex = -1 Then Exit Sub
        Dim myPort As String = ""
        Me.ErrorProvider1.SetError(cbCustSites, "")
        Try  ' Protecting
            If myCollection Is Nothing Then Exit Sub
            If Me.cbSiteVessel.SelectedIndex = 1 Then Exit Sub
            If Me.myClient.Sites Is Nothing Then Exit Sub
            'we have valid data continue
            Me.Panel2.Hide()
            SetWaitCursor()
            Dim custSite As Intranet.intranet.Address.CustomerSite

            'Me.BindingContext(Me.cbCustSites.DataSource).Position = cbCustSites.SelectedIndex
            'If Me.BindingContext(Me.cbCustSites.DataSource).Position = -1 Then Exit Sub
            custSite = Me.cbCustSites.SelectedItem
            If Not custSite Is Nothing Then
                If custSite.SiteID = "-1" Then 'dummy site
                    Me.Collection.Site = ""
                    blnCustSiteSelected = False
                Else
                    blnCustSiteSelected = True
                    Me.ToolTip1.SetToolTip(Me.cbCustSites, custSite.Info)
                    If Not myCollection Is Nothing Then
                        cmdSiteAddress.Enabled = Not (custSite.Address Is Nothing)
                        If custSite.VesselSubType.Trim.Length > 0 Then
                            Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(custSite.VesselSubType)
                            If Not oPort Is Nothing Then
                                myCollection.Port = oPort.Identity
                                Me.txtLocation.Text = oPort.Info2
                            Else
                                MsgBox("Site location error - please fix by editing site", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                            End If
                        Else
                            Me.ErrorProvider1.SetError(Me.cbCustSites, "Customer site hasn't had location set")
                        End If

                    End If
                    Dim opcoll As New Collection()
                    opcoll.Add(myClient.Identity)
                    opcoll.Add(custSite.VesselID)
                    Dim strComment As String = CConnection.PackageStringList("LIB_CCTOOL.GETCUSTOMERCOLLCOMMENTS", opcoll)
                    If strComment.Length > 0 Then
                        strComment = Medscreen.ReplaceString(strComment, "|", vbCrLf)
                        Me.txtCustComment.Text = strComment.ToString
                        Me.txtCustComment.Show()
                    End If
                    'set the site id
                    Me.Collection.Site = custSite.SiteID
                End If
            End If
            setDefaultCursor()
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbCustSites_SelectedIndexChanged-3737")
        Finally
            Me.ValidateData()
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Changing between site and vessel or vice versa
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbSiteVessel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSiteVessel.SelectedIndexChanged
        If myClient Is Nothing Then Exit Sub
        If myCollection Is Nothing Then Exit Sub
        SetWaitCursor()
        If cbSiteVessel.SelectedIndex = 0 Then
            myCollection.IsVessel = False
            If Not Me.myClient.Sites Is Nothing Then

                RebindSites()
            End If
            Me.cbCustSites.DisplayMember = "Info"
            Me.cbCustSites.ValueMember = "SiteID"
            Me.lblVessel.Text = "Customer Sites"
            Me.cBLocality.SelectedIndex = 0

            Me.cbCustSites.SelectedIndex = -1
            Me.cbCustSites.Show()
            Me.cbVessel.SelectedIndex = -1
            Me.cbVessel.Hide()
            Me.cmdAddVessel.Hide()
            Me.cmdNewSite.Show()
            Me.cmdEditSite.Show()
        Else
            myCollection.IsVessel = True
            Me.lblVessel.Text = "Vessel Name"
            Me.cbVessel.SelectedIndex = -1
            Me.cbCustSites.SelectedIndex = -1
            Me.cbCustSites.Hide()
            Me.cbVessel.Show()
            Me.cmdAddVessel.Show()
            Me.cmdNewSite.Hide()
            Me.cmdEditSite.Hide()
            'Me.txtOther.Text = ""
            'myCollection.Port = ""
            Me.txtLocation.Text = ""
        End If
        Me.ValidateData()
        setDefaultCursor()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find next due vessel
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuVesselNextDue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find next vessel due 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNextDueVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNextDueVessel.Click
        If myClient Is Nothing Then Exit Sub
        If myClient.Vessels Is Nothing Then Exit Sub
        If myClient.Vessels.Count = 0 Then Exit Sub

        Dim i As Integer
        Dim oVessel As Intranet.intranet.customerns.Vessel
        Dim LastDate As Date = Now.AddYears(2)
        Dim IVessel As Integer = -1

        For i = 0 To myClient.Vessels.Count - 1

            oVessel = myClient.Vessels.Item(i)
            If oVessel.VeselNextTest > VesselFirstDate And oVessel.VeselNextTest < LastDate Then
                IVessel = i
                LastDate = oVessel.VeselNextTest
            End If

        Next


        If IVessel > -1 Then
            Me.cbVessel.SelectedIndex = IVessel
            oVessel = myClient.Vessels.Item(IVessel)
            VesselFirstDate = oVessel.VeselNextTest.AddDays(1)

            Me.cmdNextDueVessel.ImageIndex = 1
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculate cost of this job
    ''' </summary>
    ''' <param name="blnStandardPrice"></param>
    ''' <param name="smClient"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function JobCostHTML(ByVal blnStandardPrice As Boolean, _
ByVal smClient As Intranet.intranet.customerns.Client)
        Dim strHTML As String = Medscreen.HTMLHeader
        Dim cost As Double
        Dim MinCollCharge As Double
        Dim stdSampleCharge As Double
        Dim outOfHoursCharge As Double
        Dim outOfHoursCost As Double
        Dim inHoursCharge As Double
        Dim inHoursCost As Double
        Dim inMinutes As Integer
        Dim outMinutes As Integer
        Dim inHours As Double
        Dim outHours As Double

        Try
            Dim tsOutHours As TimeSpan = TimeSpan.Zero
            Dim tsInHours As TimeSpan = TimeSpan.Zero

            If blnStandardPrice Then
                MinCollCharge = cstStandardMinCollectionCharge
                stdSampleCharge = cstStandardSampleCost
            Else
                MinCollCharge = smClient.InvoiceInfo.MinimumCharge
                stdSampleCharge = smClient.InvoiceInfo.ChargePerSample
                outOfHoursCharge = smClient.InvoiceInfo.OutOfHoursCharge
                inHoursCharge = smClient.InvoiceInfo.inHoursCharge
            End If

            strHTML += "<BR>On site collection<br>"
            strHTML += "<table class=fixedwidnb><thead><th class=noborder width = 200>Item</th>" & _
            "<th class=noborder width = 100></th>Qty<th class=noborder width = 100>Price</th><th class=noborder width = 100>Total</th></thead>"
            cost = (myCollection.NoOfDonors * stdSampleCharge)
            If cost < MinCollCharge Then
                cost = MinCollCharge
                strHTML += "<tr><td class=nbal8>Minimum Charge applies</td>"
                strHTML += "<td class=nbal8>1</td>"
                strHTML += "<td class=nbar8>" & MinCollCharge.ToString("c") & "</td>"
                strHTML += "<td class=nbar8>" & cost.ToString("c") & "</td></tr>"
            Else
                strHTML += "<tr><td class=nbal8>Sample Collection and Analysis</td>"
                strHTML += "<td class=nbal8>" & myCollection.NoOfDonors & "</td>"
                strHTML += "<td class=nbar8>" & stdSampleCharge.ToString("c") & "</td>"
                strHTML += "<td class=nbar8>" & cost.ToString("c") & "</td></tr>"
            End If
            If Not blnStandardPrice Then


                inMinutes = myCollection.NoOfDonors * 20 + 60
                inHours = CInt(inMinutes / 20)  ' samples out of hours
                inHours = inHours / 3

                inHours = inHours - outHours

                strHTML += "<tr><td class=nbal8>In hours collection</td>"
                strHTML += "<td class=nbal8>" & inHours.ToString("0.00") & "</td>"
                strHTML += "<td class=nbar8>" & inHoursCharge.ToString("c") & "</td>"
                inHoursCost = (inHours) * inHoursCharge
                strHTML += "<td class=nbar8>" & inHoursCost.ToString("c") & "</td></tr>"
                cost += inHoursCost
                strHTML += "<tr><td class=nbal8>Total</td>"
                strHTML += "<td class=nbal8></td>"
                strHTML += "<td class=nbar8></td>"
                strHTML += "<td class=nbar8>" & cost.ToString("c") & "</td></tr>"
            End If
            'If smClient.InvoiceInfo.VATRate > 0 Then
            '    strHTML += "<tr><td class=nbal8>Vat</td>"
            '    strHTML += "<td class=nbal8></td>"
            '    strHTML += "<td class=nbar8>" & smClient.InvoiceInfo.VATRate & "%</td>"
            '    strHTML += "<td class=nbar8>" & CDbl(cost * (smClient.InvoiceInfo.VATRate / 100)).ToString("c") & "</td></tr>"
            '    cost = cost * (1 + (smClient.InvoiceInfo.VATRate / 100))
            '    strHTML += "<tr><td class=nbal8>Total in VAT</td>"
            '    strHTML += "<td class=nbal8></td>"
            '    strHTML += "<td class=nbar8></td>"
            '    strHTML += "<td class=nbar8>" & cost.ToString("c") & "</td></tr>"
            'End If

            strHTML += "</table><BR><BR>"
            If Not blnStandardPrice Then
                strHTML += "<div class = centred> <b>Extra charges</b> Extra charges will apply for out of hours working " & _
                    "and collecting officer travel costs.</div>"
            Else
                strHTML += "<div class = centred> <b>Cancellation charge</b>" & _
                    " of £150 will apply If less than one working days notice" & _
                    " is given or the testing cannot commence due to site or donor circumstances <br>" & _
                    " e.g. no appropriate Id (passport, new style driving licence, company photo " & _
                    " pass etc) or no manger on site etc. </div>"
                strHTML += "<div class = centred> <b>Additional time on site</b> - will be charged at £60 per hour pro rata when collectors are " & _
                "requested to stay on site longer than specified  </div>"
            End If
            strHTML += "</body></Html>"
            dblCost = cost
            Return strHTML
        Catch ex As Exception
            LogError(ex)
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit selected site
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdEditSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditSite.Click

        If Me.myClient.Sites Is Nothing Then Exit Sub

        Dim custSite As Intranet.intranet.Address.CustomerSite

        If Me.myClient.Sites.Count < cbCustSites.SelectedIndex Then Exit Sub
        If Me.cbCustSites.SelectedIndex = -1 Then Exit Sub

        custSite = Me.myClient.Sites.Item(Me.cbCustSites.SelectedIndex)
        If Not custSite Is Nothing Then
            Dim edFrm As New MedscreenCommonGui.frmVessel()
            edFrm.Vessel = custSite

            If edFrm.ShowDialog = DialogResult.OK Then
                custSite = edFrm.Vessel
                custSite.Update()
                RebindSites()
                Me.cbCustSites.DisplayMember = "Info"
                Me.cbCustSites.ValueMember = "SiteID"

                FlipSites()
                'Me.cbCustSites.SelectedIndex = myClient.SampleManagerClient.Sites.Count - 1
                'Dim cadr As MedscreenLib.Address.Caddress = _
                '    cSupport.SMAddresses.item(custSite.AddressID)
                cmdSiteAddress.Enabled = (Not custSite.Address Is Nothing)
                'If Not cadr Is Nothing Then
                '    ' Me.txtOther.Text = cadr.BlockAddress(vbCrLf)
                'End If
                Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(custSite.VesselSubType)
                If Not oPort Is Nothing Then
                    Dim intPort As Integer
                    If oPort.IsUK Then
                        myCollection.UK = True
                        intPort = cSupport.UKPorts.IndexOf(oPort)
                        'Me.cbLocation.DataSource = cSupport.UKPorts
                    Else
                        myCollection.UK = False
                        intPort = cSupport.OSPorts.IndexOf(oPort)
                        'Me.cbLocation.DataSource = cSupport.OSPorts
                    End If
                    If intPort > -1 Then
                        'Me.cbLocation.SelectedIndex = intPort
                    End If
                End If
                Dim intSitePos As Integer = myClient.Sites.IndexOf(custSite)
                If intSitePos > -1 Then
                    Me.cbCustSites.SelectedIndex = intSitePos
                End If
            End If

        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Editing comments on collection 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtComments_TextEdit(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.DoubleClick
        Me.txtComments.Text = Medscreen.TextEditor(Me.txtComments.Text)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit collection comment
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdCustomerComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomerComment.Click
        Me.Collection.CustomerComment = Medscreen.TextEditor(Me.Collection.CustomerComment)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Trying to get display of site address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSiteAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSiteAddress.Click
        If Me.cbCustSites.SelectedIndex = -1 Then Exit Sub

        Dim oCustSite As Intranet.intranet.Address.CustomerSite = myClient.Sites.Item(Me.cbCustSites.SelectedIndex)
        If oCustSite Is Nothing Then Exit Sub

        If Not oCustSite.Address Is Nothing Then
            'Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(myCollection.Addressid)
            'If Not cadr Is Nothing Then
            Medscreen.TextEditor(oCustSite.Address.BlockAddress(vbCrLf, True), True)
            'End If
            'Else
            'Dim objPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(myCollection.Port)       'Set the site 

            'If Not objPort Is Nothing Then
            '    Dim cadr As MedscreenLib.Address.Caddress = cSupport.SMAddresses.item(objPort.AddressId)
            '    If Not cadr Is Nothing Then
            '        Medscreen.TextEditor(cadr.BlockAddress(vbCrLf, True), True)
            '    End If
        End If

        'Vessel Addresses 

        If Not myCollection.VesselObject Is Nothing AndAlso myCollection.VesselObject.Invoiceable Then 'AndAlso _
            'Not myCollection.VesselObject.Invoicing Is Nothing AndAlso _
            'myCollection.VesselObject.Invoicing.Loaded Then

            Dim strOut As String = ""
            Dim cadr As MedscreenLib.Address.Caddress
            'Invoicing address
            strOut += "Invoicing Address " & vbCrLf
            If myCollection.VesselObject.InvoicingAddress > 0 Then
                cadr = cSupport.SMAddresses.item(myCollection.VesselObject.InvoicingAddress)
                If Not cadr Is Nothing Then
                    strOut += cadr.BlockAddress(vbCrLf, True) & vbCrLf
                End If
            End If

            'Mailing address
            strOut += "Mailing Address " & vbCrLf
            If Not myCollection.VesselObject Is Nothing AndAlso myCollection.VesselObject.MailingAddress > 0 Then
                cadr = cSupport.SMAddresses.item(myCollection.VesselObject.MailingAddress)
                If Not cadr Is Nothing Then
                    strOut += cadr.BlockAddress(vbCrLf, True) & vbCrLf
                End If
            End If


            'Shipping address
            strOut += "Shipping Address " & vbCrLf
            If Not myCollection.VesselObject Is Nothing AndAlso myCollection.VesselObject.ShippingAddress > 0 Then
                cadr = cSupport.SMAddresses.item(myCollection.VesselObject.ShippingAddress)
                If Not cadr Is Nothing Then
                    strOut += cadr.BlockAddress(vbCrLf, True) & vbCrLf
                End If
            End If

            Medscreen.TextEditor(strOut, True)
        End If
        ' End If
    End Sub

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
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get help 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelp.Click
        Help.ShowHelp(Me, CommonForms.ApplicationPath() & "\CCMaintain.chm", HelpNavigator.KeywordIndex, "Collection Form")

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit vessel details 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdEditVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditVessel.Click
        If Me.myClient.Sites Is Nothing Then Exit Sub

        Dim custSite As Intranet.intranet.customerns.Vessel

        If Me.myClient.Vessels.Count < Me.cbVessel.SelectedIndex Then Exit Sub
        If Me.cbVessel.SelectedIndex = -1 Then Exit Sub

        custSite = Me.myClient.Vessels.Item(Me.cbVessel.SelectedIndex)
        If Not custSite Is Nothing Then
            Dim edFrm As New MedscreenCommonGui.frmVessel()
            edFrm.Vessel = custSite

            If edFrm.ShowDialog = DialogResult.OK Then
                custSite = edFrm.Vessel
                custSite.Update()
                RebindVessels()

                FlipSites()
                'Me.cbCustSites.SelectedIndex = myClient.SampleManagerClient.Sites.Count - 1


            End If

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with port costs 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdPortCosts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPortCosts.Click
        If Me.Collection Is Nothing Then Exit Sub
        If Me.objPort Is Nothing OrElse Me.objPort.Identity <> Me.Collection.Port Then
            If Me.Collection Is Nothing Then Exit Sub
            Me.objPort = New Intranet.intranet.jobs.Port(Me.Collection.Port)
        End If
        If Not Me.objPort Is Nothing Then
            Medscreen.TextEditor(objPort.CollectingOfficers.ToText(myClient.CurrencyCode), True, 800)
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
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [25/11/2005]</date><Action>Added code to deal with active bookings, altered warning messages</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbVessel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVessel.SelectedIndexChanged
        If Me.cbVessel.DataSource Is Nothing Then Exit Sub
        If Me.cbVessel.SelectedIndex = -1 Then Exit Sub
        If blnFormLoaded = False Then Exit Sub

        Dim oVessel As Intranet.intranet.customerns.Vessel = _
           cbVessel.DataSource(cbVessel.SelectedIndex)

        Me.ErrorProvider1.SetError(Me.cbVessel, "")
        If oVessel Is Nothing Then
            Me.ErrorProvider1.SetError(Me.cbVessel, "Vessel Not selected properly")
            'MsgBox("Vessel Not selected properly", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
            Exit Sub
        End If

        'check for stop on vessel via VDI 
        If Not oVessel Is Nothing AndAlso oVessel.Invoiceable Then
            If oVessel.StopCollections Then
                Me.ErrorProvider1.SetError(Me.cbVessel, "There is a stop on this vessel this collection cannot proceed")
                MsgBox("There is a stop on this vessel this collection cannot proceed", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
            End If
            Dim oPColl As New Collection()
            oPColl.Add(myClient.Identity)
            oPColl.Add(oVessel.VesselID)
            Dim strComment As String = CConnection.PackageStringList("LIB_CCTOOL.GETCUSTOMERCOLLCOMMENTS", oPColl)
            If strComment.Length > 0 Then
                strComment = Medscreen.ReplaceString(strComment, "|", vbCrLf)
                Me.txtCustComment.Text = strComment.ToString
                Me.txtCustComment.Show()
            End If
            'Set vessel id
            Me.Collection.vessel = oVessel.VesselID
        End If

        'check for active booking 
        If Me.Collection Is Nothing Then Exit Sub
        Dim strId As String = oVessel.ActiveBooking(Me.Collection.ID)
        If strId.Trim.Length > 0 Then 'there is a booking active
            Dim collArray As String() = strId.Split(New Char() {","})
            Dim strCid As String
            For Each strCid In collArray

            Next
            MsgBox("Collection(s) : " & strId & " are active on this vessel!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Me.ErrorProvider1.SetError(Me.cbVessel, "Collection(s) : " & strId & " are active on this vessel!")
            Me.NWInfo.Notify("Collection(s) : " & strId & " are active on this vessel!")
        End If

        'deal with maritime sites
        If oVessel.VesselType = "MARTIMSITE" Then
            If oVessel.VesselSubType.Trim.Length > 0 Then
                Dim oPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(oVessel.VesselSubType)
                If Not oPort Is Nothing Then
                    myCollection.Port = oPort.Identity
                    Me.txtLocation.Text = oPort.Info2
                Else
                    MsgBox("Site location error - please fix by editing site", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                End If
            Else
                Me.ErrorProvider1.SetError(Me.cbVessel, "Customer site hasn't had location set")
            End If

        End If


        'Check for the vessel being a Maritime Site
        If Not oVessel Is Nothing AndAlso oVessel.VesselType = GCSTMaritimeSite Then
            'GetSiteAddress(oVessel.AddressID)
        End If
        Me.ValidateData()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Port has changed 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbLocation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If Me.cbLocation.DataSource Is Nothing Then Exit Sub
        If Not Me.blnFormLoaded Then Exit Sub


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fire up the customer defaults entry screen 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCustDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustDefaults.Click
        Dim frmCOP As New frmCollDefaults()

        If myClient.CollectionInfo Is Nothing Then
            Me.myClient.CollectionInfo = New Intranet.intranet.customerns.CustColl(myClient.Identity)
        End If

        frmCOP.CollectionInfo = myClient.CollectionInfo
        If frmCOP.ShowDialog = DialogResult.OK Then

            myClient.CollectionInfo = frmCOP.CollectionInfo
            myClient.CollectionInfo.Update()
            myClient.Refresh()
            Me.FillContacts()
            'Me.TreeView1.SelectedNode = Me.TreeView1.CustomerNode
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Form is activated 
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        Me.blnFormLoaded = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Programme Managed changed 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbPMAnaged_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.cbPMAnaged.SelectedIndex = -1 Then Exit Sub
        If Me.cbPMAnaged.SelectedIndex = 0 Then
            Me.cbPManager.Visible = True
        Else
            Me.cbPManager.Visible = False
        End If
        Me.Invalidate()
        Application.DoEvents()
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
    Private Sub DomainUpDown1_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cBLocality.Enabled = False
        If Me.cBLocality.SelectedIndex = 0 Then
            'Me.cbLocation.DataSource = cSupport.UKPorts
        Else
            'Me.cbLocation.DataSource = cSupport.OSPorts
        End If
        cBLocality.Enabled = True
    End Sub

    'Private Sub cbLocation_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbLocation.Click
    '    Me.cbLocation.SelectionLength = 0
    '    Me.cbLocation.SelectionStart = 1
    'End Sub

    'Private Sub cbLocation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbLocation.KeyDown
    '    Me.cbLocation.SelectionLength = 0
    '    Me.cbLocation.SelectionStart = 1
    'End Sub

    Private Sub LinkLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkLabel1.Click

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ' Determine which link was clicked within the LinkLabel.
        Me.LinkLabel1.Links(LinkLabel1.Links.IndexOf(e.Link)).Visited = True
        ' Display the appropriate link based on the value of the LinkData property of the Link object.
        Try
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString())
        Catch ex As Exception
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See additional documents
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lblAddDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblAddDoc.Click

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculate the cost of the job 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuJobCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuJobCost.Click
        If myCollection Is Nothing Then Exit Sub
        If myCollection.NoOfDonors <= 0 Then Exit Sub

        Try
            Dim Cost As Double
            Dim strHtml As String = Me.myCollection.PriceJobHTML(Cost)
            MedscreenLib.Medscreen.ShowHtml(strHtml)
        Catch ex As Exception
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Doisplay customer's pricing 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCustomerPricing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustomerPricing.Click
        If Me.myClient Is Nothing Then Exit Sub

        Try
            Dim Cost As Double
            Dim strHtml As String = Me.myClient.ToHTML("customer.xsl", MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\transforms\xsl\")
            MedscreenLib.Medscreen.ShowHtml(strHtml)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub updnNoDonors_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles updnNoDonors.Leave
        If mylvColumn Is Nothing Then Exit Sub
        If lvSubItem = 1 Then
            mylvColumn.Column.NumberOfDonors = udEditor.Value
        ElseIf lvSubItem = 2 Then
            mylvColumn.Column.From = udEditor.Value
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Number of donors has changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub updnNoDonors_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles updnNoDonors.ValueChanged
        If Me.myCollection Is Nothing Then Exit Sub
        If Me.updnNoDonors.Value > 0 Then Me.myCollection.NoOfDonors = Me.updnNoDonors.Value
        Me.updnNoDonors.Text = CStr(Me.updnNoDonors.Value)
        ValidateData()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Esitimated actual number of donors changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub NoDonorsActual_SelectedItemChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NoDonorsActual.SelectedItemChanged
        ValidateData()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Purchase order number has been changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtPO_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.myCollection Is Nothing Then Exit Sub
        If Me.myClient Is Nothing Then Exit Sub
        myCollection.PurchaseOrder = txtPO.Text
        ValidateData()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a new port 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuNewPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewPort.Click
        Dim aPort As Intranet.intranet.jobs.Port
        Dim aFrm As New MedscreenCommonGui.frmSelectPort()
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
                aPort = .Port
                cSupport.Ports.Add(aPort)
                If aPort.IsUK Then
                    cSupport.UKPorts.Add(aPort)
                    cSupport.UKPorts.RefreshChanged()
                    cSupport.UKPorts.SortByField("SITE_NAME")
                Else
                    cSupport.OSPorts.Add(aPort)
                    cSupport.OSPorts.SortByField("SITE_NAME")
                    cSupport.OSPorts.RefreshChanged()
                End If
            End If
        End With

        'Me.cbLocation.DataSource = Nothing
        If Me.cBLocality.SelectedIndex = 0 Then
            'cSupport.UKPorts.RefreshChanged()
            'Me.cbLocation.DataSource = cSupport.UKPorts
            'Me.cbLocation.DisplayMember = "Info"
            'Me.cbLocation.ValueMember = "PortId"

            'Me.cbLocation.Refresh()
            'Dim intInd As Integer = cbLocation.DataSource.IndexOf(aPort)
            'Me.cbLocation.SelectedIndex = intInd
        Else
            'cSupport.OSPorts.RefreshChanged()
            'Me.cbLocation.DataSource = cSupport.OSPorts
            'Me.cbLocation.DisplayMember = "Info"
            'Me.cbLocation.ValueMember = "PortId"

            'Me.cbLocation.Refresh()
            'Dim intInd As Integer = cbLocation.DataSource.IndexOf(aPort)
            'Me.cbLocation.SelectedIndex = intInd
        End If

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Select port
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdFindPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindPort.Click
        Dim afrm As frmSelectPort = New frmSelectPort()
        'Set form location to anything typed in already 
        afrm.txtPort.Text = Me.txtLocation.Text
        afrm.DoFindPort()
        If afrm.ShowDialog = DialogResult.OK Then
            Dim myPort As Intranet.intranet.jobs.Port = afrm.Port
            If Not myPort Is Nothing Then
                Me.txtLocation.Text = myPort.Info2
                Me.myCollection.Port = afrm.Port.Identity
                If afrm.Port.IsUK Then
                    Me.cBLocality.SelectedIndex = 0
                Else
                    Me.cBLocality.SelectedIndex = 1
                End If
                '<Added code modified 07-Aug-2007 06:20 by taylor>   Dealing with Tony's request for officer feedback              
                OfficerPortInfo()
                '</Added code modified 07-Aug-2007 06:20 by taylor> 

            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Routine to display information on port and officers
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	07/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub OfficerPortInfo()
        If Me.Collection Is Nothing Then Exit Sub
        If Me.Collection.Port.Trim.Length = 0 Then Exit Sub
        Me.Panel2.Show()
        Me.txtAddress.Show()
        Dim Coll As New Collection()
        Coll.Add(Me.Collection.Port)
        Coll.Add(Me.DTCollDate.Value.ToString("dd-MMM-yy"))

        Dim infoString As String = CConnection.PackageStringList("Lib_collection.GetPortOfficer", Coll)
        Me.txtAddress.Rtf = Medscreen.ReplaceString(infoString, Chr(10), vbCrLf)

        Me.AssignOfficer()
    End Sub


    Dim blnLocationFound As Boolean = False
    Dim LastFind As Date
    Private Sub txtLocation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLocation.TextChanged
        Dim strLoc As String = Me.txtLocation.Text
        Try
            If blnLocationFound Then
                If Now.Subtract(LastFind).TotalSeconds < 5 Then
                    Exit Sub
                End If
            End If
            If strLoc.Trim.Length > 0 Then
                Dim strReturn As String = CConnection.PackageStringList("LIB_CCTOOL.PORTSBYID", strLoc)
                If strReturn.Trim.Length > 0 Then 'Found at least one port 
                    Dim strPorts As String() = strReturn.Split(New Char() {","})
                    If strPorts.Length = 2 Then 'Found just one port 

                        '
                        '  disable txtLocation
                        '
                        blnLocationFound = True
                        LastFind = Now
                        Dim strPortId As String = strPorts.GetValue(0)
                        Dim myPort As New Intranet.intranet.jobs.Port(strPortId)
                        If Not myPort Is Nothing Then
                            Me.txtLocation.Text = myPort.Info2
                            Me.myCollection.Port = myPort.Identity
                            If myPort.IsUK Then
                                Me.cBLocality.SelectedIndex = 0
                            Else
                                Me.cBLocality.SelectedIndex = 1
                            End If
                            'Display info about ports and officers
                            Me.OfficerPortInfo()
                            'Me.AssignOfficer()
                            Me.ErrorProvider1.SetError(Me.txtLocation, "")
                        End If
                        Me.txtLocation.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection date has changed check to see what has happened about holidays
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	07/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DTCollDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTCollDate.ValueChanged
        Me.OfficerPortInfo()
        Me.ValidateData()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' WE nned to find the officer for the collection 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	09/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAssign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAssign.Click
        If Not GetForm() Then 'Check data is valid, if there are any issues raised 
            'Change the form return to none, this will make it return to displaying and the form will 
            'not close 
            Me.DialogResult = DialogResult.None
            MsgBox("There are errors in the collection data, please fix")
            Me.TabControl1.SelectedIndex = 0

            Exit Sub
        End If
        If myClient.InvoiceInfo.CreditCard_pay Then
            MsgBox("This client is set for credit card payment We can't assign an officer at this stage")
            Exit Sub
        End If


        If Me.myCollection.HasID = 0 Then         'Indicates a new collection
            'blnNew = True
            Dim strCheckmessage As String = "Do you wish to create this collection for : " & Me.myClient.ClientID & " - " & _
                Me.myCollection.VesselName & " - at " & Me.myCollection.Port & "?"
            If MsgBox(strCheckmessage, MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                CreateCollection()
            End If
        End If
        'Stage one get a list of officers for port
        If Not Me.Collection Is Nothing AndAlso Me.Collection.ID <> 0 AndAlso Me.strCollId.Trim.Length > 0 Then
            Me.Collection.CollOfficer = Me.strCollId
            Try
                Me.Collection.SetStatus(Constants.GCST_JobStatusAssigned)  'Set status to assigned
                Me.Collection.Update()                                     'Update collection 
            Catch cex As MedscreenExceptions.CanNotChangeCollectionStatus
                MsgBox(cex.ToString)
                Exit Sub
            End Try

            'Now see about sending
            If MsgBox("Do you wish to send this collection (" & myCollection.ID & _
                ") to the collecting officer (" & _
                myCollection.CollOfficer & ") now?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Try
                    myCollection.Status = MedscreenLib.Constants.GCST_JobStatusSent
                    myCollection.Update()
                    'SendToCollector(False, True) send and print collection
                    myCollection.SendToCollector("", "", False, True, "Send")
                    myCollection.Comments.CreateComment("Collection sent to Officer " & myCollection.CollOfficer, Constants.GCST_COMM_CollSendToOfficer)
                    myCollection.Refresh()         'Get modified data
                    myCollection.LogAction(Intranet.intranet.jobs.CMJob.Actions.SendToCollOfficer)

                Catch cex As MedscreenExceptions.CanNotChangeCollectionStatus
                    MsgBox(cex.ToString)
                    Exit Sub
                End Try
            End If

        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the officer list box 
    ''' </summary>
    ''' <param name="OfficerArray"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	09/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub FillOfficerList(ByVal OfficerArray As String())
        Me.ListView1.Items.Clear()
        If OfficerArray.Length = 0 Then Exit Sub

        Dim lvItem As MedscreenCommonGui.ListViewItems.LVCofficerItem
        Dim COff As Intranet.intranet.jobs.CollectingOfficer
        Dim k As Integer

        For k = 0 To OfficerArray.Length - 1
            COff = New Intranet.intranet.jobs.CollectingOfficer(OfficerArray.GetValue(k))
            COff.Load()
            lvItem = New LVCofficerItem(COff, Me.Collection.Port, Me.DTCollDate.Value) 'Put info into specialist element
            If lvItem.OfficerId = Me.Collection.CollOfficer Then lvItem.Selected = True 'If we match the Officer make it selected
            lvItem = Me.ListView1.Items.Add(lvItem)
            lvItem.BackColor = System.Drawing.SystemColors.Info
            lvItem.UseItemStyleForSubItems = False
        Next
    End Sub

    Private strCollId As String
    Dim lvItem As LVCofficerItem
    Dim intEmailChecked As Integer

    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click, ListView1.ItemActivate
        lvItem = Nothing
        Dim strId As String
        Dim strHTML As String
        Dim dQuote As String = Chr(34)

        Me.strCollId = ""
        If Me.ListView1.SelectedItems.Count <> 1 Then Exit Sub

        lvItem = Me.ListView1.SelectedItems.Item(0)
        strId = lvItem.Text
        Me.strCollId = strId
        'Me.txtOfficer.Text = strId
        'Me.GetPorts(strId)

        Dim CollOff As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(strId)
        If Not CollOff Is Nothing AndAlso Not CollOff.Ports Is Nothing Then
            Dim cPort As Intranet.intranet.jobs.CollPort
            cPort = CollOff.Ports.Item(Me.Collection.Port, 0)
            If Not cPort Is Nothing Then
                strHTML = cPort.ToHtml(Me.DTCollDate.Value)
            Else

                strHTML = CollOff.ToHtml(Me.DTCollDate.Value)
            End If
            Me.LVRedirect.Items.Clear()
            Me.intEmailChecked = -1
            'Default destination
            If CollOff.SendDestination.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem("Default Send")
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, ""))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.SendDestination))
                lvItemR.ImageIndex = 4
                lvItemR.Checked = True
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            'Deal with multisend
            'If Not CollOff.MultiSendAddresses Is Nothing Then
            '    Dim aSend As MultiSendColl

            '    For Each aSend In CollOff.MultiSendAddresses

            '        Dim lvItemR As ListViewItem = New ListViewItem(aSend.SourceOfficer)
            '        lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, aSend.DestOfficer))
            '        lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, aSend.EmailFaxNumber))
            '        lvitemr.ImageIndex = 5
            '        Me.LVRedirect.Items.Add(lvItemR)
            '    Next
            'End If
            If CollOff.HomePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Home Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.HomePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.OfficePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Office Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.OfficePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.MobilePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Mobile Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.MobilePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.HomeFax.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Home Fax"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.HomeFax))
                lvItemR.ImageIndex = 2
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.Email.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Email"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.Email))
                lvItemR.ImageIndex = 1
                Me.LVRedirect.Items.Add(lvItemR)
            End If

        Else
            strHTML = "<HTML><head><LINK REL=STYLESHEET TYPE=" & _
            dQuote & "text/css" & dQuote & " HREF = " & dQuote & _
                Medscreen.StyleSheet & dQuote & " Title=" & dQuote & _
                "Style" & dQuote & "><body>"

            strHTML += "Collecting Officer " & lvItem.FirstName & " " & lvItem.LastName & "<BR>"
            strHTML += "Office Phone " & lvItem.OfficePhone & "<BR>"
            strHTML += "Home Phone " & lvItem.HomePhone & "<BR>"
            strHTML += "Mobile Phone " & lvItem.Mobile & "<BR>"
            strHTML += "Pager " & lvItem.Beeper & "<BR>"
            strHTML += "Costs agreed " & lvItem.Agreed & "<BR>"
            strHTML += "Extra Costs " & lvItem.Cost.ToString & "<BR>"


            strHTML += "</body></HTML>"
        End If

        Try
            Me.HtmlEditor1.DocumentText = (strHTML)
        Catch ex As Exception
        End Try
        Me.strCollId = strId
        Me.cmdAssign.Enabled = True
    End Sub

    Private Sub AssignOfficer()
        'Stage one get a list of officers for port
        Dim strOfficers As String = CConnection.PackageStringList("Lib_collection.GetPortOfficerList", Me.Collection.Port)
        If strOfficers.Length > 0 Then
            Dim strOfficerArray As String() = strOfficers.Split(New Char() {","})
            FillOfficerList(strOfficerArray)
        End If
        'Me.cmdAssign.Enabled = True
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If Me.lvItem Is Nothing Then Exit Sub

        If Me.lvItem.Officer Is Nothing Then Exit Sub
        Dim frmedColl As New frmCollOfficerEd()

        frmedColl.CollectingOfficer = Me.lvItem.Officer
        If frmedColl.ShowDialog = DialogResult.OK Then
            Me.lvItem.Officer = frmedColl.CollectingOfficer
        End If

    End Sub

    Private Sub mnuAddOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddOfficer.Click
        Dim myPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(myCollection.Port)

        If myPort Is Nothing Then Exit Sub

        Dim objCollPort As Intranet.intranet.jobs.CollPort = New Intranet.intranet.jobs.CollPort()
        objCollPort.PortId = myPort.PortId

        Dim frmAO As New MedscreenCommonGui.frmAssignOfficer()
        Try
            frmAO.CollPort = objCollPort
            If frmAO.ShowDialog = DialogResult.OK Then
                objCollPort = frmAO.CollPort
                objCollPort.Update()
                myPort.CollectingOfficers.Add(objCollPort)

                Dim lvItem As LVCofficerItem = New LVCofficerItem(objCollPort.Officer, myCollection.Port, Me.DTCollDate.Value)
                lvItem = Me.ListView1.Items.Add(lvItem)
                lvItem.BackColor = System.Drawing.SystemColors.Info
                lvItem.UseItemStyleForSubItems = False
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "AddOfficer - frmcollofficer")
        End Try
    End Sub

    Private Sub mnuRemoveOfficer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuRemoveOfficer.Click
        If Me.ListView1.SelectedItems.Count <> 1 Then Exit Sub

        Dim lvItem As ListViewItem = ListView1.SelectedItems.Item(0)

        If TypeOf lvItem Is LVCofficerItem Then
            Dim lvCitem As LVCofficerItem = lvItem
            If Not lvCitem Is Nothing Then
                If Not lvCitem.CollectionPort Is Nothing Then
                    lvCitem.CollectionPort.Removed = True
                    lvCitem.CollectionPort.Update()
                    ListView1.Items.RemoveAt(lvItem.Index)
                End If
            End If
        End If

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Send email to officer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSendEmail.Click
        If Me.intEmailChecked = -1 Then Exit Sub
        Dim CollOff As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(strCollId)

        Dim strEmail As String = ""             'Default send destination to null string 
        If Me.intEmailChecked = 0 Then          'Default send method 
            If Not CollOff.CanEmailorFax Then
                MsgBox("Can't send to this type of address", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If CollOff.SendBy = Constants.CollOfficerSend.HomeEmail Or CollOff.SendBy = Constants.CollOfficerSend.WorkEmail Or _
            CollOff.SendBy = Constants.CollOfficerSend.HomeFax Or CollOff.SendBy = Constants.CollOfficerSend.WorkFax Then
                strEmail = CollOff.SendDestination
            End If
        Else 'from multisend 
            Dim lvitem As ListViewItem = Me.LVRedirect.Items(Me.intEmailChecked)
            If Not lvitem Is Nothing Then
                Dim sItem As ListViewItem.ListViewSubItem = lvitem.SubItems(1)
                Dim mailType As String = sItem.Text.ToUpper
                sItem = lvitem.SubItems(2)
                strEmail = Medscreen.RemoveChars(sItem.Text, " ")
                If InStr(mailType, "PHONE") Then
                    MsgBox("Can't send to a phone")
                    Exit Sub
                End If
                If Medscreen.IsNumber(strEmail) Then
                    strEmail += MedscreenLib.Constants.GCST_EMAIL_To_Fax_Send
                End If
            End If
        End If


        If strEmail.Trim.Length > 0 Then
            Dim objOutlook As outlook.ApplicationClass = New outlook.ApplicationClass()
            Dim objMessage As outlook.MailItemClass
            Try
                objMessage = objOutlook.CreateItem(outlook.OlItemType.olMailItem)
                objMessage.Subject = "Message from " & Intranet.intranet.Support.cSupport.User.Description
                objMessage.Recipients.Add(strEmail)
                objMessage.Display(True)
            Catch ex As Exception
            Finally
                objOutlook = Nothing
            End Try
        End If
    End Sub

    Private Sub LVRedirect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LVRedirect.Click
        Dim lvItem As ListViewItem
        If LVRedirect.SelectedItems.Count > 0 Then
            lvItem = LVRedirect.SelectedItems.Item(0)
            Me.intEmailChecked = lvItem.Index
        End If
    End Sub

    Private Function GetRandomFilename() As String
        Dim strFilename As String
        Dim strPath As String = Intranet.Support.AdditionalDocumentColl.CreateCollectionDocumentPath(Me.Collection.ID)
        strFilename = strPath & Me.Collection.ID & "-RandomNumbers.doc"
        Return strFilename
    End Function
    Private RandomNumberTemplate As RandomNumbers = Nothing
    Private Sub cmdGenerateRandom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateRandom.Click
        'Generate a set of random numbers 
        Dim intMax As Integer = 0
        Dim intNNum As Integer = 0

        '************************** Deal with Random number template ************************************'
        'Template should be loaded on selecting tab
        'If myClient IsNot Nothing AndAlso myClient.CollectionInfo IsNot Nothing AndAlso myClient.CollectionInfo.RandomTemplate.Trim.Length > 0 Then
        '    RandomNumberTemplate = New RandomNumbers(myClient.CollectionInfo.RandomTemplate)
        '    RandomNumberTemplate.ReadFromDB()
        'End If
        Me.lvRandom.Rows.Clear()
        Dim SubCollPosition As Integer = 0
        Dim TotalCols As Integer = 0
        Dim MaxRow As Integer = 0
        For Each col As RandomNumberColumn In RandomNumberTemplate.Columns
            If col.DataImportMethod = RandomNumberColumn.DataImport.Excel Then            'Need to get file
                OpenFileDialog1.FileName = ""
                Dim objFilename As String = ""
                If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    objFilename = OpenFileDialog1.FileName
                End If

                If objFilename Is Nothing OrElse objFilename.Trim.Length = 0 Then Exit Sub
                col.Filename = objFilename
                MedscreenLib.ExcelSupport.Open(col.Filename)
                For Each sb As SubColumn In col.SubColumns
                    'Get start columns
                    If sb.DataStart.Trim.Length > 0 Then
                        sb.DataStart = MedscreenLib.Medscreen.GetParameter(MyTypes.typString, "Data Position : " & sb.Header, , sb.DataStart)
                    End If
                    If sb.DataStart.Trim.Length > 0 Then
                        col.From = ExcelSupport.CountRows(sb.DataStart)
                        Exit For
                    End If
                Next
            End If
            If col.NumberOfDonors > MaxRow Then MaxRow = col.NumberOfDonors
        Next

        For row As Integer = 1 To MaxRow
            Dim lvItem As DataGridViewRow
            lvItem = New DataGridViewRow()
            For r As Integer = 1 To lvRandom.Columns.Count
                lvItem.Cells.Add(New DataGridViewTextBoxCell())
            Next
            Me.lvRandom.Rows.Add(lvItem)
        Next
        Dim blnHasCol1Index As Boolean = False
        For Each col As RandomNumberColumn In RandomNumberTemplate.Columns
            SubCollPosition += col.SubColumns.Count
            If col.From = 0 Or col.NumberOfDonors = 0 Then
                MsgBox("You need to enter the numbers")
                Exit Sub
            End If
            Dim Coll As New Collection()
            Dim i As Integer
            'Build list
            For i = 1 To col.From
                Coll.Add(i)
            Next


            Dim rndNum As Integer
            intNNum = col.NumberOfDonors
            intMax = col.From
            Dim outColl As New Hashtable
            Dim k As Integer
            Dim rndDbl As Double
            Try
                While intNNum > 0
                    rndDbl = Rnd()
                    rndNum = Int((intMax) * rndDbl + 1)
                    Dim nNUM As Integer = Coll.Item(rndNum)
                    If nNUM <> 0 Then
                        Dim key As String = CStr(nNUM).PadLeft(10)
                        key = key.Replace(" ", "0")
                        outColl.Add(key & "|" & CStr(rndDbl), key)
                        Coll.Remove(rndNum)
                        intNNum -= 1
                        intMax -= 1
                        Dim outstring As String = ""

                    End If
                End While
                k = 1

                For Each elm As DictionaryEntry In outColl
                    'Dim lvItem As DataGridViewRow



                    Dim blnFirst As Boolean = True
                    Dim strPos As String = elm.Key
                    Dim strNums As String() = strPos.Split(New Char() {"|"})

                    For Each subcol As SubColumn In col.SubColumns
                        Dim ColIndex As Integer = TotalCols + col.SubColumns.IndexOf(subcol)
                        If blnFirst AndAlso SubCollPosition = col.SubColumns.Count Then
                            If subcol.ColType = SubColumn.ColumnType.Index Then
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CStr(k)
                                blnHasCol1Index = True
                            ElseIf subcol.ColType = SubColumn.ColumnType.RandomNumber Then
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CStr(CInt(strNums.GetValue(0)))
                            ElseIf subcol.ColType = SubColumn.ColumnType.Name Or subcol.ColType = SubColumn.ColumnType.Dept Then
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = ExcelSupport.GetCell(subcol.DataStart, CInt(strNums.GetValue(0)), 0)
                            ElseIf subcol.ColType = SubColumn.ColumnType.Formula Then

                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CDbl(strNums.GetValue(1)).ToString("0.00000")

                            End If
                            blnFirst = False
                            'ElseIf blnHasCol1Index Then
                            'lvRandom.Rows(k - 1).Cells(0).Value = CStr(k)

                        Else
                            If blnHasCol1Index Then
                                lvRandom.Rows(k - 1).Cells(0).Value = CStr(k)
                                lvRandom.Rows(k - 1).Cells(0).Style.BackColor = Drawing.Color.Beige
                            End If
                            If col.Index > 1 Then
                                'For ik As Integer = lvItem.SubItems.Count To SubCollPosition - col.SubColumns.Count - 1
                                '    Dim lsunb As ListViewItem.ListViewSubItem
                                '    lsunb = New ListViewItem.ListViewSubItem(lvItem, "")
                                '    lvItem.SubItems.Add(lsunb)
                                'Next
                            End If
                            If subcol.ColType = SubColumn.ColumnType.Index AndAlso col.Index > 1 Then
                                
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CStr(k)
                            ElseIf subcol.ColType = SubColumn.ColumnType.RandomNumber Then
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CStr(CInt(strNums.GetValue(0)))
                            ElseIf subcol.ColType = SubColumn.ColumnType.Name Or subcol.ColType = SubColumn.ColumnType.Dept Then
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = ExcelSupport.GetCell(subcol.DataStart, CInt(strNums.GetValue(0)), 0)
                            ElseIf subcol.ColType = SubColumn.ColumnType.Formula Then
                                Dim lsun As ListViewItem.ListViewSubItem
                                'lsun = New ListViewItem.ListViewSubItem(lvItem, strNums.GetValue(1))
                                'lvItem.SubItems.Add(lsun)
                                lvRandom.Rows(k - 1).Cells(ColIndex).Value = CDbl(strNums.GetValue(1)).ToString("0.00000")
                            End If
                        End If
                    Next
                    k += 1
                Next
                TotalCols += col.SubColumns.Count
                If RandomNumberTemplate.ColumnGapSize > 0 Then TotalCols += 1
            Catch ex As Exception

            End Try
        Next



        'intNNum = intMax





    End Sub

    Private Sub frmCollectionForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ReBindResonForTest(ByVal objReason As Object)

        Me.cbRandom.DataSource = objReason
        Me.cbRandom.BindingContext(Me.cbRandom.DataSource).SuspendBinding()
        Me.cbRandom.BindingContext(Me.cbRandom.DataSource).ResumeBinding()
        Me.cbRandom.DisplayMember = "PhraseText"
        Me.cbRandom.ValueMember = "PhraseID"

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' On change of selected login template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	20/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cbSampleTemplate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSampleTemplate.SelectedIndexChanged
        'Need to get matrices and sample login template
        If Me.myClient Is Nothing Then Exit Sub
        Try  ' Protecting
            Dim strTempId As String = Me.cbSampleTemplate.SelectedValue

            Dim strReasonForTest As String = MedscreenCommonGUIConfig.CollectionQuery.Item("ReasonForTestRoutine")
            Dim objReason = New Glossary.PhraseCollection(strReasonForTest, Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , Me.myClient.Identity & "," & strTempId)
            ReBindResonForTest(objReason)
            Me.cbRandom.DisplayMember = "PhraseText"
            Me.cbRandom.ValueMember = "PhraseID"
            'need to make sure we update the combo 
            Try  ' Protecting

                If Not Me.myCollection Is Nothing Then
                    Me.cbRandom.SelectedValue = Me.myCollection.CollType
                End If
                Me.txtMatrix.Text = myClient.MatrixText(strTempId)

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5214")
            Finally

            End Try
            Dim strLabId As String = CConnection.PackageStringList("lib_customer2.GetLaboratory", strTempId)
            If Not strLabId Is Nothing AndAlso strLabId.Trim.Length > 0 Then Me.cbAnalyticalLab.SelectedValue = strLabId

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5202")
        Finally
        End Try
    End Sub

    Private Sub gbCollection_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbCollection.Enter

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

    End Sub


    Private Sub cbRandom_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRandom.Leave
        Me.ValidateData()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Display Lab Address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdLabAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLabAddress.Click
        Try
            Dim objLab As Intranet.INSTRUMENT.Instrument = New Intranet.INSTRUMENT.Instrument(Me.cbAnalyticalLab.SelectedValue)

            If Not objLab Is Nothing Then
                Dim cadr As MedscreenLib.Address.Caddress = objLab.Address
                If Not cadr Is Nothing Then
                    Medscreen.TextEditor(cadr.BlockAddress(vbCrLf, True), True)
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Sub cbAnalyticalLab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnalyticalLab.SelectedIndexChanged
        If Me.myCollection Is Nothing Then Exit Sub

        Dim strInstrument As String = Me.cbAnalyticalLab.SelectedValue
    End Sub

    Private Sub txtAddCosts_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddCosts.TextChanged

    End Sub

    Private Sub cbCustSites_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCustSites.SelectedValueChanged
        If Not blnFormLoaded Then Exit Sub
        If Not Me.cbCustSites.SelectedValue Is Nothing Then
            Debug.Write(Me.cbCustSites.SelectedValue)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert the currency to the customers 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	05/08/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConvert.Click
        With myCollection

            'If we don't have client info (should never occur)
            'Addtional Costs 
            Try
                'There may be a currency symbol in the first place, in a try block in case it goes do lalhi
                If Me.txtAddCosts.Text.Chars(0) < "0" Or Me.txtAddCosts.Text.Chars(0) > "9" Then
                    .AgreedExpenses = Mid(Me.txtAddCosts.Text, 2)
                Else
                    .AgreedExpenses = Me.txtAddCosts.Text
                End If
            Catch ex As Exception
                .AgreedExpenses = 0
            End Try

            'See if the currency selected is the same as the customer's if not convert
            If Me.cbCurrency.Text <> Me.CustomerCurrency Then
                Dim strin As String
                Dim strout As String

                'No match set up for conversion, conversion use ISO three character names.
                strin = Me.CustomerCurrency
                'Deal with collecting officer currencies
                strout = Me.cbCurrency.SelectedValue
                'Use conversion routines and current conversion rate, 
                'if any errors will almost certainly be caused by the conversion rate not being in the table
                .AgreedExpenses = MedscreenLib.Glossary.Glossary.ExchangeRates.Convert(strout, strin, .AgreedExpenses)
                Me.txtAddCosts.Text = .AgreedExpenses.ToString("c", cr)
                Me.cbCurrency.SelectedValue = Me.CustomerCurrency
            End If
        End With
    End Sub

    Private Sub DrawDocuments()
        Me.lvDocuments.Items.Clear()
        If Me.Collection Is Nothing Then Exit Sub
        If Not Me.Collection.ID Is Nothing AndAlso Me.Collection.ID.Trim.Length > 0 Then
            Dim strDocs As String = CConnection.PackageStringList("lib_collectionxml8.GetInstructionListIDs", Me.myCollection.ID)
            Dim docArray As String() = strDocs.Split(New Char() {","})
            Dim i As Integer
            For i = 0 To docArray.Length - 1
                Dim strDoc As String = docArray(i)
                If strDoc.Trim.Length > 0 Then
                    Dim lvAddDoc As New lvAdditionalDoc(New Intranet.Support.AdditionalDocument(strDoc))
                    Me.lvDocuments.Items.Add(lvAddDoc)
                End If

            Next
        End If
        If Me.Collection.SmClient Is Nothing Then Exit Sub
        Me.lvDocuments.Items.Add(New ListViewItem("Procedures"))
        Dim strProc As String = CConnection.PackageStringList("lib_collection.ProcedureFiles", myCollection.SmClient.Identity)
        Dim strFiles As String() = strProc.Split(New Char() {","})
        Dim k As Integer
        For k = 0 To strFiles.Length - 1
            Dim strFile As String = strFiles(k)
            If strFile.Trim.Length > 0 Then
                Dim strElements As String() = strFile.Split(New Char() {"|"})
                ' Path is in the second element, doc id in first 
                Dim fName As String = strElements(1)
                If IO.File.Exists(fName) Then

                    Dim lvAddDoc As New lvAdditionalDoc(New Intranet.Support.AdditionalDocument(strElements(0)))
                    Me.lvDocuments.Items.Add(lvAddDoc)
                End If
            End If
        Next

    End Sub

    Private Sub cbRandom_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cbRandom.Validating
        Me.ErrorProvider1.SetError(cbRandom, "")
        If cbRandom.SelectedValue = "" Then
            Me.ErrorProvider1.SetError(cbRandom, "A reason for test must be given")
            Application.DoEvents()
            e.Cancel() = True
        End If
    End Sub

    Private Sub cmdAddDoc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddDoc.Click
        AddDoc()
    End Sub

    Private CurrentDoc As lvAdditionalDoc = Nothing
    Private Sub lvDocuments_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDocuments.SelectedIndexChanged
        CurrentDoc = Nothing
        If lvDocuments.SelectedItems.Count = 1 Then
            CurrentDoc = lvDocuments.SelectedItems(0)
        End If
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        If CurrentDoc Is Nothing Then Exit Sub
        If Me.tmpReference.Trim.Length = 0 Then tmpReference = Me.myCollection.ID
        strDirectoryPath = MedscreenLib.MedConnection.Instance.ServerPath & "\DM_DataFile\CollDocs\"
        If Process1 Is Nothing Then Process1 = New System.Diagnostics.Process()
        If InStr(CurrentDoc.AdditionalDocument.Path, "\\") = 0 Then
            Process1.StartInfo.FileName = strDirectoryPath & tmpReference & "\" & CurrentDoc.AdditionalDocument.Path
        Else
            Process1.StartInfo.FileName = CurrentDoc.AdditionalDocument.Path

        End If

        Process1.Start()
        Process1.WaitForExit(3000)
    End Sub

    Private Sub cmdRemoveDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveDocument.Click
        If CurrentDoc Is Nothing Then Exit Sub
        If Me.tmpReference.Trim.Length = 0 Then tmpReference = Me.myCollection.ID
        strDirectoryPath = MedscreenLib.MedConnection.Instance.ServerPath & "\DM_DataFile\CollDocs\"
        Dim strFilename As String = strDirectoryPath & tmpReference & "\" & CurrentDoc.AdditionalDocument.Path
        If IO.File.Exists(strFilename) Then
            IO.File.Delete(strFilename)
        End If
        CurrentDoc.AdditionalDocument.Fields.Delete(CConnection.DbConnection)
    End Sub

    Private Sub lvDocuments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDocuments.DoubleClick
        Me.cmdEdit_Click(Nothing, Nothing)
    End Sub


    Private Sub cmdEditDoc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEditDoc.Click


        Dim strFilename As String = GetRandomFilename()
        If Not IO.File.Exists(strFilename) Then
            MsgBox("No Random File created")
            Exit Sub
        End If
        If Process1 Is Nothing Then Process1 = New System.Diagnostics.Process()

        Process1.StartInfo.FileName = strFilename


        Process1.Start()
        Process1.WaitForExit(3000)

    End Sub

    Private Sub DrawRandomNumbers()
        If Me.Collection.HasID Then
            If myClient IsNot Nothing AndAlso myClient.CollectionInfo IsNot Nothing AndAlso myClient.CollectionInfo.RandomTemplate.Trim.Length > 0 Then
                RandomNumberTemplate = New RandomNumbers(myClient.CollectionInfo.RandomTemplate)
                RandomNumberTemplate.ReadFromDB()
                Me.lblTemplate.Text = "Template : " & RandomNumberTemplate.Title
                Me.lvRandom.Columns.Clear()
                lvColumnGroups.Items.Clear()
                Try
                    For Each col As RandomNumberColumn In RandomNumberTemplate.Columns
                        Dim lvCol As New ListViewItems.lvRandomColumn(col, "Group " & lvColumnGroups.Items.Count + 1, lvColumnGroups.Items.Count + 1)
                        lvColumnGroups.Items.Add(lvCol)
                        For Each subcol As SubColumn In col.SubColumns
                            Dim colHead As New DataGridViewColumn()
                            colHead.CellTemplate = New DataGridViewTextBoxCell
                            colHead.DefaultCellStyle = lvRandom.DefaultCellStyle
                            colHead.HeaderText = subcol.Header
                            colHead.Width = 120
                            If subcol.ColType = SubColumn.ColumnType.Index OrElse subcol.ColType = SubColumn.ColumnType.Formula Then
                                colHead.ReadOnly = True
                                colHead.CellTemplate.Style.BackColor = Drawing.Color.AntiqueWhite
                            End If

                            lvRandom.Columns.Add(colHead)
                        Next
                        If RandomNumberTemplate.ColumnGapSize > 0 AndAlso col.Index <> RandomNumberTemplate.Columns.Count Then
                            Dim colHead As New DataGridViewColumn()
                            colHead.CellTemplate = New DataGridViewTextBoxCell
                            colHead.DefaultCellStyle = lvRandom.DefaultCellStyle
                            colHead.HeaderText = ""
                            colHead.Width = RandomNumberTemplate.ColumnGapSize
                            colHead.ReadOnly = True
                            colHead.CellTemplate.Style.BackColor = Drawing.Color.AntiqueWhite
                            lvRandom.Columns.Add(colHead)
                        End If
                    Next
                Catch ex As Exception
                End Try
            End If
            Dim strFilename As String = GetRandomFilename()

            Dim strDocScanned As String = CConnection.PackageStringList("lib_collection.DocumentScanned", strFilename)
            If strDocScanned IsNot Nothing AndAlso strDocScanned.Trim.Length > 0 Then cmdEditDoc.Enabled = True

        Else
            TabControl1.SelectedIndex = 0
        End If
    End Sub
    Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 2 Then
            DrawRandomNumbers()
        End If
    End Sub
    Private mylvColumn As ListViewItems.lvRandomColumn
    Private lvSubItem As Integer = -1
    Private Sub lvColumnGroups_SubItemClicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs) Handles lvColumnGroups.SubItemClicked
        Dim lvItem As ListViewItem = e.Item

        lvSubItem = e.SubItem
        mylvColumn = Nothing
        If TypeOf lvItem Is ListViewItems.lvRandomColumn Then
            mylvColumn = lvItem
        End If


        Dim Editors As Control() = New Control() {Me.udEditor}
        If lvSubItem = 1 Then
            sender.StartEditing(Editors(0), e.Item, e.SubItem)
        ElseIf lvSubItem = 2 Then
            sender.StartEditing(Editors(0), e.Item, e.SubItem)

        End If

    End Sub

    Private Sub udEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles udEditor.Leave
        If mylvColumn Is Nothing Then Exit Sub
        If lvSubItem = 1 Then
            mylvColumn.Column.NumberOfDonors = udEditor.Value
        ElseIf lvSubItem = 2 Then
            mylvColumn.Column.From = udEditor.Value

        End If
    End Sub

    Private Sub cmdCreateDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateDocument.Click

        If lvRandom.Rows.Count = 0 Then
            MsgBox("You need to generate and possibly edit the data in the grid.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            Exit Sub
        End If
        Dim strFilename As String
        Dim ioF As IO.StreamWriter
        Dim rndTemplate As String = Me.RandomNumberTemplate.WordTemplate
        Dim objWord As New MedscreenLib.WordSupport(rndTemplate, True)
        Dim TotalColWidth As Integer = 0
        
        Try


            'Generate Out File 

            strFilename = GetRandomFilename()
            objWord.SetBookmark("CompanyName", Me.Collection.SmClient.CompanyName & " " & " from a possible ")
            If Me.Collection.SmClient.ResultsAddress IsNot Nothing Then objWord.SetBookmark("CompanyAddress", Me.Collection.SmClient.ResultsAddress.BlockAddress(vbCrLf))
            Dim objTable As Word.Table = objWord.Document.Tables.Item(2)

            'insert header as in listview
            Dim colPos As Integer = 1
            For Each ch As DataGridViewColumn In lvRandom.Columns
                objWord.InsertCellContents(objTable, 1, colPos, ch.HeaderText)
                TotalColWidth += ch.Width
                colPos += 1
            Next

            Dim Rowpos As Integer = 1
            For Each lv As DataGridViewRow In lvRandom.Rows
                colPos = 1
                For Each sb As DataGridViewCell In lv.Cells
                    If sb.Value IsNot Nothing Then
                        objWord.InsertCellContents(objTable, Rowpos + 1, colPos, sb.Value.ToString)
                    Else
                        objWord.InsertCellContents(objTable, Rowpos + 1, colPos, "")
                    End If
                    colPos += 1
                Next

                Rowpos += 1
            Next
            Dim intTableWidth As Double = objWord.Document.PageSetup.PageWidth - objWord.Document.PageSetup.LeftMargin - objWord.Document.PageSetup.RightMargin

            Dim Ratio As Double = intTableWidth / TotalColWidth
            Dim iPos As Integer = 1
            For Each ch As DataGridViewColumn In lvRandom.Columns
                objTable.Columns.Item(iPos).Width = ch.Width * Ratio
                iPos += 1
            Next
        Catch ex As Exception
        Finally
            objWord.SaveAs(strFilename)
            objWord.Close()
        End Try
        Dim strDocScanned As String = CConnection.PackageStringList("lib_collection.DocumentScanned", strFilename)
        If strDocScanned IsNot Nothing AndAlso strDocScanned.Trim.Length > 0 Then Exit Sub
        Dim objAddDoc As Intranet.Support.AdditionalDocument = Intranet.Support.AdditionalDocumentColl.CreateCollectionDocument(Me.Collection.ID, strFilename, "INSTRCTOIN")

        'ioF.Flush()
        'ioF.Close()
        'Dim strNewNamestub As String = Mid(strFilename, 1, strFilename.Length - 5)
        'Dim strNewName As String = strNewNamestub & ".out"
        'Dim ik As Integer = Asc("1")
        'While IO.File.Exists(strNewName)
        '    strNewName = strNewNamestub & Chr(ik) & ".out"
        '    ik += 1
        'End While
        'IO.File.Move(strFilename, strNewName)
    End Sub
    Private lvRandomSubItemIndex As Integer = -1
    Private mylvGridColumn As ListViewItem
    Private Sub lvRandom_SubItemClicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs)
        Dim lvItem As ListViewItem = e.Item

        lvRandomSubItemIndex = e.SubItem
        mylvGridColumn = Nothing
        'If TypeOf lvItem Is ListViewItems.lvRandomColumn Then
        mylvGridColumn = lvItem
        'End If


        Dim Editors As Control() = New Control() {Me.txtEditor}

        sender.StartEditing(Editors(0), e.Item, e.SubItem)

    End Sub

    Private Sub txtEditor_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEditor.KeyDown
        Debug.Print(e.KeyValue)
        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Or e.KeyCode = Keys.Tab Then
            txtEditor_Leave(Nothing, Nothing)

        End If
        'lvRandom.SelectedItems.Clear()
        'lvRandom.Focus()

        'If e.KeyCode = Keys.Down Then
        '    Dim newRow As Integer = mylvGridColumn.Index + 1
        '    lvRandom.Items(newRow).Selected = True
        '    lvRandom.Items(newRow).EnsureVisible()

        '    Application.DoEvents()
        'End If
    End Sub



    Private Sub txtEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEditor.Leave
        If mylvGridColumn IsNot Nothing Then mylvGridColumn.SubItems(lvRandomSubItemIndex).Text = txtEditor.Text
    End Sub

    Private Sub cmdSendRandom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSendRandom.Click
        Dim strFilename As String = GetRandomFilename()
        If Not IO.File.Exists(strFilename) Then
            MsgBox("No Random File created")
            Exit Sub
        End If
        If myCollection.CollectingOfficer Is Nothing Then
            MsgBox("No collecting officer assigned can't continue", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Medscreen.BlatEmail("Random numbers for collection " & myCollection.ID, "Please find enclosed the random numbers for the above collection", myCollection.CollectingOfficer.SendDestination, , , , strFilename)
        Dim sLog As SendLogInfo = SendLogCollection.CreateSendLog()
        With sLog
            .Filename = strFilename
            .ReferenceID = myCollection.ID
            .ReferenceTo = "Random Numbers"
            .SendingProgram = "Collection Manager"
            .SendMethod = "MAIL"
            .SentOn = Today
            .ReceivingAddress = myCollection.CollectingOfficer.SendDestination
            .DoUpdate()
        End With
    End Sub

    
End Class
