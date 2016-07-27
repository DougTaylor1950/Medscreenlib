'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-28 17:21:42+00 $
'$Log: frmCollDefaults.vb,v $
'Revision 1.0  2005-11-28 17:21:42+00  taylor
'Commented
'
Imports Intranet.intranet
Imports MedscreenLib
Imports MedscreenLib.Glossary.Glossary

Imports System.Windows.Forms
Public Class frmCollDefaults
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'Set up reason for test 
        Me.cbRandom.DataSource = MedscreenLib.Glossary.Glossary.ReasonForTestList
        Me.cbRandom.DisplayMember = "PhraseText"
        Me.cbRandom.ValueMember = "PhraseID"

        'Set up test type 

        Me.cbTestType.DataSource = objSampleTemplates
        Me.cbTestType.DisplayMember = "PhraseText"
        Me.cbTestType.ValueMember = "PhraseID"

        'Set up Confirmation Method 
        Me.cbConfMethod.DataSource = MedscreenLib.Glossary.Glossary.ConfMethodList
        Me.cbConfMethod.DisplayMember = "PhraseText"
        Me.cbConfMethod.ValueMember = "PhraseID"

        'Set up confirmation Type
        Me.cbConfType.DataSource = MedscreenLib.Glossary.Glossary.ConfTypeList
        Me.cbConfType.DisplayMember = "PhraseText"
        Me.cbConfType.ValueMember = "PhraseID"

        Me.cbProgMan.DataSource = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.RoleHolders("PROGRAMMEMANAGER")
        Me.cbProgMan.DisplayMember = "User"
        Me.cbProgMan.ValueMember = "User"


        objContTypePhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_TYP")
        objContTypePhrase.Load("Phrase_text")
        objContMethodPhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_METH")
        objContMethodPhrase.Load("Phrase_text")
        objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("REP_OPTION")
        objRepOptionPhrase.Load("Phrase_text")

        objContStatusPhrase = New MedscreenLib.Glossary.PhraseCollection("CONT_STAT")
        objContStatusPhrase.Load("Phrase_text")

        '</Added code modified 05-Jul-2007 09:18 by LOGOS\Taylor> 

        Me.cbContactType.DataSource = objContTypePhrase
        Me.cbContactType.DisplayMember = "PhraseText"
        Me.cbContactType.ValueMember = "PhraseID"

        Me.cbContMethod.DataSource = objContMethodPhrase
        Me.cbContMethod.DisplayMember = "PhraseText"
        Me.cbContMethod.ValueMember = "PhraseID"

        Me.cbReportOptions.DataSource = objRepOptionPhrase
        Me.cbReportOptions.DisplayMember = "PhraseText"
        Me.cbReportOptions.ValueMember = "PhraseID"


        Me.cbContactStatus.DataSource = objContStatusPhrase
        Me.cbContactStatus.DisplayMember = "PhraseText"
        Me.cbContactStatus.ValueMember = "PhraseID"
        Dim strSquery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("LaboratoryQuery")
        '        Dim strSQuery As String = "select identity,description from instrument s " & _
        '"where s.removeflag='F' and insttype_id = 'LABORATORY'"
        Me.objAnalalyticalLabs = New MedscreenLib.Glossary.PhraseCollection(strSquery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)


        Me.cbAnalyticalLab.DataSource = objAnalalyticalLabs
        Me.cbAnalyticalLab.DisplayMember = "PhraseText"
        Me.cbAnalyticalLab.ValueMember = "PhraseID"



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
    Friend WithEvents CkProcedures As System.Windows.Forms.CheckBox
    Friend WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents nudNoOfDonors As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtWhoToTest As System.Windows.Forms.TextBox
    Friend WithEvents cbTestType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbRandom As System.Windows.Forms.ComboBox
    Friend WithEvents lblCollType As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbConfType As System.Windows.Forms.ComboBox
    Friend WithEvents cbConfMethod As System.Windows.Forms.ComboBox
    Friend WithEvents ckUK As System.Windows.Forms.CheckBox
    Friend WithEvents cbProgMan As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents pnlFixedSiteDetails As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtFixedSiteContactPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtFixedSiteContact As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents pnlContact As System.Windows.Forms.Panel
    Friend WithEvents cbContactStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label87 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdateContact As System.Windows.Forms.Button
    Friend WithEvents cbReportOptions As System.Windows.Forms.ComboBox
    Friend WithEvents Label85 As System.Windows.Forms.Label
    Friend WithEvents cbReportFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label86 As System.Windows.Forms.Label
    Friend WithEvents cbContMethod As System.Windows.Forms.ComboBox
    Friend WithEvents Label84 As System.Windows.Forms.Label
    Friend WithEvents cbContactType As System.Windows.Forms.ComboBox
    Friend WithEvents Label83 As System.Windows.Forms.Label
    Friend WithEvents AdrContact As MedscreenCommonGui.AddressPanel
    Friend WithEvents Label81 As System.Windows.Forms.Label
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents cmdFindContactAdr As System.Windows.Forms.Button
    Friend WithEvents cmdNewContactAdr As System.Windows.Forms.Button
    Friend WithEvents pnlContactList As System.Windows.Forms.Panel
    Friend WithEvents cmdDelContact As System.Windows.Forms.Button
    Friend WithEvents cmdAddContact As System.Windows.Forms.Button
    Friend WithEvents lvContacts As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddComment As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtMatrix As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents cbAnalyticalLab As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbCOSchedule As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbFSSchedule As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtCOMatrix As System.Windows.Forms.TextBox
    Friend WithEvents txtFSMatrix As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblCompanyname As System.Windows.Forms.Label
    Friend WithEvents Process1 As System.Diagnostics.Process
    Friend WithEvents tpProcedures As System.Windows.Forms.TabPage
    Friend WithEvents PnlProcControls As System.Windows.Forms.Panel
    Friend WithEvents pnlProcs As System.Windows.Forms.Panel
    Friend WithEvents lvProcedures As System.Windows.Forms.ListView
    Friend WithEvents cmdEditProcFile As System.Windows.Forms.Button
    Friend WithEvents cmdAddProc As System.Windows.Forms.Button
    Friend WithEvents tpRandomNumbers As System.Windows.Forms.TabPage
    Friend WithEvents txtRandomTitle As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtHeader As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lvSubColumns As ListViewEx.ListViewEx
    Friend WithEvents lvColumnGroups As System.Windows.Forms.ListView
    Friend WithEvents udColumnSets As System.Windows.Forms.NumericUpDown
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddGroup As System.Windows.Forms.Button
    Friend WithEvents cmdRandomSave As System.Windows.Forms.Button
    Friend WithEvents cmdAddRandomHeader As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveRandomHeader As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteProc As System.Windows.Forms.Button
    Friend WithEvents txtEditor As System.Windows.Forms.TextBox
    Friend WithEvents udFrom As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents udNoOfDonors As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents dgExample As System.Windows.Forms.DataGridView
    Friend WithEvents txtWordTemplate As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtGroupName As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents udColumnGap As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cbRandomType As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cbImportType As System.Windows.Forms.ComboBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents chDataStart As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbEditor As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCollDefaults))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.CkProcedures = New System.Windows.Forms.CheckBox
        Me.label1 = New System.Windows.Forms.Label
        Me.nudNoOfDonors = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtWhoToTest = New System.Windows.Forms.TextBox
        Me.cbTestType = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbRandom = New System.Windows.Forms.ComboBox
        Me.lblCollType = New System.Windows.Forms.Label
        Me.cbConfType = New System.Windows.Forms.ComboBox
        Me.cbConfMethod = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.ckUK = New System.Windows.Forms.CheckBox
        Me.cbProgMan = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtCustomerID = New System.Windows.Forms.TextBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.pnlFixedSiteDetails = New System.Windows.Forms.Panel
        Me.txtFixedSiteContactPhone = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtFixedSiteContact = New System.Windows.Forms.TextBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.lblCompanyname = New System.Windows.Forms.Label
        Me.txtFSMatrix = New System.Windows.Forms.TextBox
        Me.txtCOMatrix = New System.Windows.Forms.TextBox
        Me.cbFSSchedule = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cbCOSchedule = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbAnalyticalLab = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtMatrix = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.cmdAddComment = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.pnlContact = New System.Windows.Forms.Panel
        Me.cbContactStatus = New System.Windows.Forms.ComboBox
        Me.Label87 = New System.Windows.Forms.Label
        Me.cmdUpdateContact = New System.Windows.Forms.Button
        Me.cbReportOptions = New System.Windows.Forms.ComboBox
        Me.Label85 = New System.Windows.Forms.Label
        Me.cbReportFormat = New System.Windows.Forms.ComboBox
        Me.Label86 = New System.Windows.Forms.Label
        Me.cbContMethod = New System.Windows.Forms.ComboBox
        Me.Label84 = New System.Windows.Forms.Label
        Me.cbContactType = New System.Windows.Forms.ComboBox
        Me.Label83 = New System.Windows.Forms.Label
        Me.pnlContactList = New System.Windows.Forms.Panel
        Me.cmdDelContact = New System.Windows.Forms.Button
        Me.cmdAddContact = New System.Windows.Forms.Button
        Me.lvContacts = New System.Windows.Forms.ListView
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader
        Me.tpProcedures = New System.Windows.Forms.TabPage
        Me.pnlProcs = New System.Windows.Forms.Panel
        Me.lvProcedures = New System.Windows.Forms.ListView
        Me.PnlProcControls = New System.Windows.Forms.Panel
        Me.cmdDeleteProc = New System.Windows.Forms.Button
        Me.cmdAddProc = New System.Windows.Forms.Button
        Me.cmdEditProcFile = New System.Windows.Forms.Button
        Me.tpRandomNumbers = New System.Windows.Forms.TabPage
        Me.cbImportType = New System.Windows.Forms.ComboBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.cbRandomType = New System.Windows.Forms.ComboBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.udColumnGap = New System.Windows.Forms.NumericUpDown
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtGroupName = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtWordTemplate = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.dgExample = New System.Windows.Forms.DataGridView
        Me.udFrom = New System.Windows.Forms.NumericUpDown
        Me.Label17 = New System.Windows.Forms.Label
        Me.udNoOfDonors = New System.Windows.Forms.NumericUpDown
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtEditor = New System.Windows.Forms.TextBox
        Me.cmdRemoveRandomHeader = New System.Windows.Forms.Button
        Me.cmdAddRandomHeader = New System.Windows.Forms.Button
        Me.cmdRandomSave = New System.Windows.Forms.Button
        Me.cmdAddGroup = New System.Windows.Forms.Button
        Me.lvSubColumns = New ListViewEx.ListViewEx
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.lvColumnGroups = New System.Windows.Forms.ListView
        Me.udColumnSets = New System.Windows.Forms.NumericUpDown
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtHeader = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtRandomTitle = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cbEditor = New System.Windows.Forms.ComboBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Process1 = New System.Diagnostics.Process
        Me.AdrContact = New MedscreenCommonGui.AddressPanel
        Me.Label81 = New System.Windows.Forms.Label
        Me.Label82 = New System.Windows.Forms.Label
        Me.cmdFindContactAdr = New System.Windows.Forms.Button
        Me.cmdNewContactAdr = New System.Windows.Forms.Button
        Me.chDataStart = New System.Windows.Forms.ColumnHeader
        Me.Panel1.SuspendLayout()
        CType(Me.nudNoOfDonors, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFixedSiteDetails.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.pnlContact.SuspendLayout()
        Me.pnlContactList.SuspendLayout()
        Me.tpProcedures.SuspendLayout()
        Me.pnlProcs.SuspendLayout()
        Me.PnlProcControls.SuspendLayout()
        Me.tpRandomNumbers.SuspendLayout()
        CType(Me.udColumnGap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgExample, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udNoOfDonors, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udColumnSets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.AdrContact.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 661)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(896, 40)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(812, 8)
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
        Me.cmdOk.Location = New System.Drawing.Point(732, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 8
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CkProcedures
        '
        Me.CkProcedures.Location = New System.Drawing.Point(8, 24)
        Me.CkProcedures.Name = "CkProcedures"
        Me.CkProcedures.Size = New System.Drawing.Size(104, 24)
        Me.CkProcedures.TabIndex = 2
        Me.CkProcedures.Text = "Has Procedures"
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(232, 24)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(144, 23)
        Me.label1.TabIndex = 3
        Me.label1.Text = "Default number of samples"
        Me.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'nudNoOfDonors
        '
        Me.nudNoOfDonors.Location = New System.Drawing.Point(376, 24)
        Me.nudNoOfDonors.Name = "nudNoOfDonors"
        Me.nudNoOfDonors.Size = New System.Drawing.Size(120, 20)
        Me.nudNoOfDonors.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Who to test"
        '
        'txtWhoToTest
        '
        Me.txtWhoToTest.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWhoToTest.Location = New System.Drawing.Point(72, 56)
        Me.txtWhoToTest.Multiline = True
        Me.txtWhoToTest.Name = "txtWhoToTest"
        Me.txtWhoToTest.Size = New System.Drawing.Size(816, 48)
        Me.txtWhoToTest.TabIndex = 6
        '
        'cbTestType
        '
        Me.cbTestType.Location = New System.Drawing.Point(112, 152)
        Me.cbTestType.Name = "cbTestType"
        Me.cbTestType.Size = New System.Drawing.Size(312, 21)
        Me.cbTestType.TabIndex = 71
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 73
        Me.Label3.Text = "Routine Test Schedule"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbRandom
        '
        Me.cbRandom.Items.AddRange(New Object() {"", "NEW CUSTOMER", "OTHER", "PRE-EMPLOYMENT", "RANDOM", "ROUTINE"})
        Me.cbRandom.Location = New System.Drawing.Point(112, 120)
        Me.cbRandom.Name = "cbRandom"
        Me.cbRandom.Size = New System.Drawing.Size(136, 21)
        Me.cbRandom.TabIndex = 70
        '
        'lblCollType
        '
        Me.lblCollType.Location = New System.Drawing.Point(8, 120)
        Me.lblCollType.Name = "lblCollType"
        Me.lblCollType.Size = New System.Drawing.Size(100, 23)
        Me.lblCollType.TabIndex = 72
        Me.lblCollType.Text = "Reason for test"
        Me.lblCollType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbConfType
        '
        Me.cbConfType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConfType.Location = New System.Drawing.Point(488, 104)
        Me.cbConfType.Name = "cbConfType"
        Me.cbConfType.Size = New System.Drawing.Size(400, 21)
        Me.cbConfType.TabIndex = 74
        '
        'cbConfMethod
        '
        Me.cbConfMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConfMethod.Location = New System.Drawing.Point(488, 128)
        Me.cbConfMethod.Name = "cbConfMethod"
        Me.cbConfMethod.Size = New System.Drawing.Size(400, 21)
        Me.cbConfMethod.TabIndex = 75
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(357, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 23)
        Me.Label6.TabIndex = 78
        Me.Label6.Text = "Confirmation Type"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(359, 136)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 23)
        Me.Label7.TabIndex = 79
        Me.Label7.Text = "Confirmation Method"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ckUK
        '
        Me.ckUK.Location = New System.Drawing.Point(528, 24)
        Me.ckUK.Name = "ckUK"
        Me.ckUK.Size = New System.Drawing.Size(152, 24)
        Me.ckUK.TabIndex = 84
        Me.ckUK.Text = "Primarily UK collections "
        '
        'cbProgMan
        '
        Me.cbProgMan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbProgMan.Location = New System.Drawing.Point(600, 243)
        Me.cbProgMan.Name = "cbProgMan"
        Me.cbProgMan.Size = New System.Drawing.Size(280, 21)
        Me.cbProgMan.TabIndex = 86
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(480, 243)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 23)
        Me.Label13.TabIndex = 85
        Me.Label13.Text = "Programme Manager"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(40, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 16)
        Me.Label8.TabIndex = 87
        Me.Label8.Text = "Customer ID"
        '
        'txtCustomerID
        '
        Me.txtCustomerID.Enabled = False
        Me.txtCustomerID.Location = New System.Drawing.Point(144, 0)
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.Size = New System.Drawing.Size(128, 20)
        Me.txtCustomerID.TabIndex = 88
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        '
        'pnlFixedSiteDetails
        '
        Me.pnlFixedSiteDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlFixedSiteDetails.Controls.Add(Me.txtFixedSiteContactPhone)
        Me.pnlFixedSiteDetails.Controls.Add(Me.Label9)
        Me.pnlFixedSiteDetails.Controls.Add(Me.Label10)
        Me.pnlFixedSiteDetails.Controls.Add(Me.txtFixedSiteContact)
        Me.pnlFixedSiteDetails.Location = New System.Drawing.Point(16, 301)
        Me.pnlFixedSiteDetails.Name = "pnlFixedSiteDetails"
        Me.pnlFixedSiteDetails.Size = New System.Drawing.Size(368, 96)
        Me.pnlFixedSiteDetails.TabIndex = 89
        Me.pnlFixedSiteDetails.Visible = False
        '
        'txtFixedSiteContactPhone
        '
        Me.txtFixedSiteContactPhone.Location = New System.Drawing.Point(112, 64)
        Me.txtFixedSiteContactPhone.Name = "txtFixedSiteContactPhone"
        Me.txtFixedSiteContactPhone.Size = New System.Drawing.Size(232, 20)
        Me.txtFixedSiteContactPhone.TabIndex = 83
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(24, 32)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 24)
        Me.Label9.TabIndex = 84
        Me.Label9.Text = "Fixed Site Contact"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 23)
        Me.Label10.TabIndex = 85
        Me.Label10.Text = "Contact number"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFixedSiteContact
        '
        Me.txtFixedSiteContact.Location = New System.Drawing.Point(112, 32)
        Me.txtFixedSiteContact.Name = "txtFixedSiteContact"
        Me.txtFixedSiteContact.Size = New System.Drawing.Size(232, 20)
        Me.txtFixedSiteContact.TabIndex = 82
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.tpProcedures)
        Me.TabControl1.Controls.Add(Me.tpRandomNumbers)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(896, 661)
        Me.TabControl1.TabIndex = 90
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lblCompanyname)
        Me.TabPage1.Controls.Add(Me.txtFSMatrix)
        Me.TabPage1.Controls.Add(Me.txtCOMatrix)
        Me.TabPage1.Controls.Add(Me.cbFSSchedule)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.cbCOSchedule)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.cbAnalyticalLab)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Panel2)
        Me.TabPage1.Controls.Add(Me.txtMatrix)
        Me.TabPage1.Controls.Add(Me.pnlFixedSiteDetails)
        Me.TabPage1.Controls.Add(Me.txtCustomerID)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.cbProgMan)
        Me.TabPage1.Controls.Add(Me.Label13)
        Me.TabPage1.Controls.Add(Me.ckUK)
        Me.TabPage1.Controls.Add(Me.cbConfType)
        Me.TabPage1.Controls.Add(Me.cbConfMethod)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.cbTestType)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.cbRandom)
        Me.TabPage1.Controls.Add(Me.lblCollType)
        Me.TabPage1.Controls.Add(Me.txtWhoToTest)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.nudNoOfDonors)
        Me.TabPage1.Controls.Add(Me.label1)
        Me.TabPage1.Controls.Add(Me.CkProcedures)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(888, 635)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "General"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lblCompanyname
        '
        Me.lblCompanyname.AutoSize = True
        Me.lblCompanyname.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblCompanyname.Location = New System.Drawing.Point(304, 8)
        Me.lblCompanyname.Name = "lblCompanyname"
        Me.lblCompanyname.Size = New System.Drawing.Size(0, 13)
        Me.lblCompanyname.TabIndex = 101
        '
        'txtFSMatrix
        '
        Me.txtFSMatrix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFSMatrix.Location = New System.Drawing.Point(432, 216)
        Me.txtFSMatrix.Multiline = True
        Me.txtFSMatrix.Name = "txtFSMatrix"
        Me.txtFSMatrix.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtFSMatrix.Size = New System.Drawing.Size(456, 24)
        Me.txtFSMatrix.TabIndex = 99
        '
        'txtCOMatrix
        '
        Me.txtCOMatrix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCOMatrix.Location = New System.Drawing.Point(432, 184)
        Me.txtCOMatrix.Multiline = True
        Me.txtCOMatrix.Name = "txtCOMatrix"
        Me.txtCOMatrix.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtCOMatrix.Size = New System.Drawing.Size(456, 24)
        Me.txtCOMatrix.TabIndex = 98
        '
        'cbFSSchedule
        '
        Me.cbFSSchedule.Location = New System.Drawing.Point(112, 216)
        Me.cbFSSchedule.Name = "cbFSSchedule"
        Me.cbFSSchedule.Size = New System.Drawing.Size(312, 21)
        Me.cbFSSchedule.TabIndex = 96
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 216)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 23)
        Me.Label11.TabIndex = 97
        Me.Label11.Text = "Fixed Site Test Schedule"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCOSchedule
        '
        Me.cbCOSchedule.Location = New System.Drawing.Point(112, 184)
        Me.cbCOSchedule.Name = "cbCOSchedule"
        Me.cbCOSchedule.Size = New System.Drawing.Size(312, 21)
        Me.cbCOSchedule.TabIndex = 94
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 184)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 95
        Me.Label5.Text = "Call Out Test Schedule"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbAnalyticalLab
        '
        Me.cbAnalyticalLab.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAnalyticalLab.Location = New System.Drawing.Point(136, 270)
        Me.cbAnalyticalLab.Name = "cbAnalyticalLab"
        Me.cbAnalyticalLab.Size = New System.Drawing.Size(744, 21)
        Me.cbAnalyticalLab.TabIndex = 93
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 270)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 92
        Me.Label4.Text = "Analytical Lab"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.txtComments)
        Me.Panel2.Controls.Add(Me.lblComments)
        Me.Panel2.Location = New System.Drawing.Point(8, 408)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(880, 216)
        Me.Panel2.TabIndex = 91
        '
        'txtComments
        '
        Me.txtComments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtComments.Location = New System.Drawing.Point(128, 16)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComments.Size = New System.Drawing.Size(736, 184)
        Me.txtComments.TabIndex = 84
        '
        'lblComments
        '
        Me.lblComments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblComments.Location = New System.Drawing.Point(16, 16)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(120, 23)
        Me.lblComments.TabIndex = 85
        Me.lblComments.Text = "Associated Comments"
        '
        'txtMatrix
        '
        Me.txtMatrix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMatrix.Location = New System.Drawing.Point(432, 152)
        Me.txtMatrix.Multiline = True
        Me.txtMatrix.Name = "txtMatrix"
        Me.txtMatrix.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMatrix.Size = New System.Drawing.Size(456, 24)
        Me.txtMatrix.TabIndex = 90
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.cmdAddComment)
        Me.TabPage2.Controls.Add(Me.ListView1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(888, 635)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Comments"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'cmdAddComment
        '
        Me.cmdAddComment.Location = New System.Drawing.Point(64, 264)
        Me.cmdAddComment.Name = "cmdAddComment"
        Me.cmdAddComment.Size = New System.Drawing.Size(120, 23)
        Me.cmdAddComment.TabIndex = 3
        Me.cmdAddComment.Text = "Add Comment"
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(888, 248)
        Me.ListView1.TabIndex = 2
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Created"
        Me.ColumnHeader1.Width = 80
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Width = 1
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Comment"
        Me.ColumnHeader3.Width = 600
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.pnlContact)
        Me.TabPage3.Controls.Add(Me.pnlContactList)
        Me.TabPage3.Controls.Add(Me.AdrContact)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(888, 635)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Contacts"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'pnlContact
        '
        Me.pnlContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
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
        Me.pnlContact.Location = New System.Drawing.Point(576, 256)
        Me.pnlContact.Name = "pnlContact"
        Me.pnlContact.Size = New System.Drawing.Size(312, 379)
        Me.pnlContact.TabIndex = 243
        '
        'cbContactStatus
        '
        Me.cbContactStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContactStatus.Location = New System.Drawing.Point(44, 264)
        Me.cbContactStatus.Name = "cbContactStatus"
        Me.cbContactStatus.Size = New System.Drawing.Size(257, 21)
        Me.cbContactStatus.TabIndex = 235
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
        Me.cmdUpdateContact.Image = CType(resources.GetObject("cmdUpdateContact.Image"), System.Drawing.Image)
        Me.cmdUpdateContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdateContact.Location = New System.Drawing.Point(40, 296)
        Me.cmdUpdateContact.Name = "cmdUpdateContact"
        Me.cmdUpdateContact.Size = New System.Drawing.Size(112, 23)
        Me.cmdUpdateContact.TabIndex = 233
        Me.cmdUpdateContact.Text = "Save"
        Me.cmdUpdateContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbReportOptions
        '
        Me.cbReportOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReportOptions.Location = New System.Drawing.Point(40, 212)
        Me.cbReportOptions.Name = "cbReportOptions"
        Me.cbReportOptions.Size = New System.Drawing.Size(257, 21)
        Me.cbReportOptions.TabIndex = 7
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
        Me.cbReportFormat.Size = New System.Drawing.Size(257, 21)
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
        Me.cbContMethod.Location = New System.Drawing.Point(40, 96)
        Me.cbContMethod.Name = "cbContMethod"
        Me.cbContMethod.Size = New System.Drawing.Size(257, 21)
        Me.cbContMethod.TabIndex = 3
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
        Me.cbContactType.Location = New System.Drawing.Point(40, 40)
        Me.cbContactType.Name = "cbContactType"
        Me.cbContactType.Size = New System.Drawing.Size(257, 21)
        Me.cbContactType.TabIndex = 1
        '
        'Label83
        '
        Me.Label83.Location = New System.Drawing.Point(40, 16)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(100, 23)
        Me.Label83.TabIndex = 0
        Me.Label83.Text = "Contact Type"
        '
        'pnlContactList
        '
        Me.pnlContactList.Controls.Add(Me.cmdDelContact)
        Me.pnlContactList.Controls.Add(Me.cmdAddContact)
        Me.pnlContactList.Controls.Add(Me.lvContacts)
        Me.pnlContactList.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactList.Location = New System.Drawing.Point(576, 0)
        Me.pnlContactList.Name = "pnlContactList"
        Me.pnlContactList.Size = New System.Drawing.Size(312, 256)
        Me.pnlContactList.TabIndex = 241
        '
        'cmdDelContact
        '
        Me.cmdDelContact.Image = CType(resources.GetObject("cmdDelContact.Image"), System.Drawing.Image)
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
        Me.cmdAddContact.Image = CType(resources.GetObject("cmdAddContact.Image"), System.Drawing.Image)
        Me.cmdAddContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddContact.Location = New System.Drawing.Point(40, 216)
        Me.cmdAddContact.Name = "cmdAddContact"
        Me.cmdAddContact.Size = New System.Drawing.Size(96, 23)
        Me.cmdAddContact.TabIndex = 1
        Me.cmdAddContact.Text = "Add Contact"
        Me.cmdAddContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvContacts
        '
        Me.lvContacts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader23, Me.ColumnHeader24, Me.ColumnHeader28, Me.ColumnHeader25, Me.ColumnHeader26, Me.ColumnHeader27})
        Me.lvContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvContacts.FullRowSelect = True
        Me.lvContacts.Location = New System.Drawing.Point(0, 0)
        Me.lvContacts.Name = "lvContacts"
        Me.lvContacts.Size = New System.Drawing.Size(312, 200)
        Me.lvContacts.TabIndex = 0
        Me.lvContacts.UseCompatibleStateImageBehavior = False
        Me.lvContacts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "Contact"
        Me.ColumnHeader23.Width = 180
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
        'tpProcedures
        '
        Me.tpProcedures.Controls.Add(Me.pnlProcs)
        Me.tpProcedures.Controls.Add(Me.PnlProcControls)
        Me.tpProcedures.Location = New System.Drawing.Point(4, 22)
        Me.tpProcedures.Name = "tpProcedures"
        Me.tpProcedures.Size = New System.Drawing.Size(888, 635)
        Me.tpProcedures.TabIndex = 3
        Me.tpProcedures.Text = "Procedures"
        Me.tpProcedures.UseVisualStyleBackColor = True
        '
        'pnlProcs
        '
        Me.pnlProcs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlProcs.Controls.Add(Me.lvProcedures)
        Me.pnlProcs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlProcs.Location = New System.Drawing.Point(0, 0)
        Me.pnlProcs.Name = "pnlProcs"
        Me.pnlProcs.Padding = New System.Windows.Forms.Padding(3)
        Me.pnlProcs.Size = New System.Drawing.Size(688, 635)
        Me.pnlProcs.TabIndex = 1
        '
        'lvProcedures
        '
        Me.lvProcedures.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvProcedures.FullRowSelect = True
        Me.lvProcedures.Location = New System.Drawing.Point(3, 3)
        Me.lvProcedures.MultiSelect = False
        Me.lvProcedures.Name = "lvProcedures"
        Me.lvProcedures.Size = New System.Drawing.Size(678, 625)
        Me.lvProcedures.TabIndex = 0
        Me.lvProcedures.UseCompatibleStateImageBehavior = False
        Me.lvProcedures.View = System.Windows.Forms.View.List
        '
        'PnlProcControls
        '
        Me.PnlProcControls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PnlProcControls.Controls.Add(Me.cmdDeleteProc)
        Me.PnlProcControls.Controls.Add(Me.cmdAddProc)
        Me.PnlProcControls.Controls.Add(Me.cmdEditProcFile)
        Me.PnlProcControls.Dock = System.Windows.Forms.DockStyle.Right
        Me.PnlProcControls.Location = New System.Drawing.Point(688, 0)
        Me.PnlProcControls.Name = "PnlProcControls"
        Me.PnlProcControls.Size = New System.Drawing.Size(200, 635)
        Me.PnlProcControls.TabIndex = 0
        '
        'cmdDeleteProc
        '
        Me.cmdDeleteProc.Enabled = False
        Me.cmdDeleteProc.Image = CType(resources.GetObject("cmdDeleteProc.Image"), System.Drawing.Image)
        Me.cmdDeleteProc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDeleteProc.Location = New System.Drawing.Point(8, 96)
        Me.cmdDeleteProc.Name = "cmdDeleteProc"
        Me.cmdDeleteProc.Size = New System.Drawing.Size(176, 23)
        Me.cmdDeleteProc.TabIndex = 105
        Me.cmdDeleteProc.Text = "Delete File"
        Me.cmdDeleteProc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDeleteProc, "Delete procedure File")
        '
        'cmdAddProc
        '
        Me.cmdAddProc.Image = CType(resources.GetObject("cmdAddProc.Image"), System.Drawing.Image)
        Me.cmdAddProc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddProc.Location = New System.Drawing.Point(8, 56)
        Me.cmdAddProc.Name = "cmdAddProc"
        Me.cmdAddProc.Size = New System.Drawing.Size(176, 23)
        Me.cmdAddProc.TabIndex = 104
        Me.cmdAddProc.Text = "Add Proc File"
        Me.cmdAddProc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdAddProc, "Add procedure File")
        '
        'cmdEditProcFile
        '
        Me.cmdEditProcFile.Enabled = False
        Me.cmdEditProcFile.Image = CType(resources.GetObject("cmdEditProcFile.Image"), System.Drawing.Image)
        Me.cmdEditProcFile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEditProcFile.Location = New System.Drawing.Point(8, 16)
        Me.cmdEditProcFile.Name = "cmdEditProcFile"
        Me.cmdEditProcFile.Size = New System.Drawing.Size(176, 23)
        Me.cmdEditProcFile.TabIndex = 103
        Me.cmdEditProcFile.Text = "Edit File"
        Me.cmdEditProcFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEditProcFile, "Edit procedure File")
        '
        'tpRandomNumbers
        '
        Me.tpRandomNumbers.Controls.Add(Me.cbImportType)
        Me.tpRandomNumbers.Controls.Add(Me.Label22)
        Me.tpRandomNumbers.Controls.Add(Me.cbRandomType)
        Me.tpRandomNumbers.Controls.Add(Me.Label21)
        Me.tpRandomNumbers.Controls.Add(Me.udColumnGap)
        Me.tpRandomNumbers.Controls.Add(Me.Label20)
        Me.tpRandomNumbers.Controls.Add(Me.txtGroupName)
        Me.tpRandomNumbers.Controls.Add(Me.Label19)
        Me.tpRandomNumbers.Controls.Add(Me.txtWordTemplate)
        Me.tpRandomNumbers.Controls.Add(Me.Label18)
        Me.tpRandomNumbers.Controls.Add(Me.dgExample)
        Me.tpRandomNumbers.Controls.Add(Me.udFrom)
        Me.tpRandomNumbers.Controls.Add(Me.Label17)
        Me.tpRandomNumbers.Controls.Add(Me.udNoOfDonors)
        Me.tpRandomNumbers.Controls.Add(Me.Label16)
        Me.tpRandomNumbers.Controls.Add(Me.txtEditor)
        Me.tpRandomNumbers.Controls.Add(Me.cmdRemoveRandomHeader)
        Me.tpRandomNumbers.Controls.Add(Me.cmdAddRandomHeader)
        Me.tpRandomNumbers.Controls.Add(Me.cmdRandomSave)
        Me.tpRandomNumbers.Controls.Add(Me.cmdAddGroup)
        Me.tpRandomNumbers.Controls.Add(Me.lvSubColumns)
        Me.tpRandomNumbers.Controls.Add(Me.lvColumnGroups)
        Me.tpRandomNumbers.Controls.Add(Me.udColumnSets)
        Me.tpRandomNumbers.Controls.Add(Me.Label15)
        Me.tpRandomNumbers.Controls.Add(Me.txtHeader)
        Me.tpRandomNumbers.Controls.Add(Me.Label14)
        Me.tpRandomNumbers.Controls.Add(Me.txtRandomTitle)
        Me.tpRandomNumbers.Controls.Add(Me.Label12)
        Me.tpRandomNumbers.Controls.Add(Me.cbEditor)
        Me.tpRandomNumbers.Location = New System.Drawing.Point(4, 22)
        Me.tpRandomNumbers.Name = "tpRandomNumbers"
        Me.tpRandomNumbers.Size = New System.Drawing.Size(888, 635)
        Me.tpRandomNumbers.TabIndex = 4
        Me.tpRandomNumbers.Text = "Random Numbers"
        Me.tpRandomNumbers.UseVisualStyleBackColor = True
        '
        'cbImportType
        '
        Me.cbImportType.FormattingEnabled = True
        Me.cbImportType.Location = New System.Drawing.Point(644, 129)
        Me.cbImportType.Name = "cbImportType"
        Me.cbImportType.Size = New System.Drawing.Size(121, 21)
        Me.cbImportType.TabIndex = 27
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(589, 128)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(37, 13)
        Me.Label22.TabIndex = 26
        Me.Label22.Text = "Type :"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbRandomType
        '
        Me.cbRandomType.FormattingEnabled = True
        Me.cbRandomType.Location = New System.Drawing.Point(644, 34)
        Me.cbRandomType.Name = "cbRandomType"
        Me.cbRandomType.Size = New System.Drawing.Size(121, 21)
        Me.cbRandomType.TabIndex = 25
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(572, 43)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(37, 13)
        Me.Label21.TabIndex = 24
        Me.Label21.Text = "Type :"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'udColumnGap
        '
        Me.udColumnGap.Location = New System.Drawing.Point(644, 60)
        Me.udColumnGap.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.udColumnGap.Name = "udColumnGap"
        Me.udColumnGap.Size = New System.Drawing.Size(53, 20)
        Me.udColumnGap.TabIndex = 23
        Me.ToolTip1.SetToolTip(Me.udColumnGap, "Column Gap Zero indicates no Gap")
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(572, 62)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(68, 13)
        Me.Label20.TabIndex = 22
        Me.Label20.Text = "Column Gap:"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtGroupName
        '
        Me.txtGroupName.Location = New System.Drawing.Point(424, 129)
        Me.txtGroupName.Name = "txtGroupName"
        Me.txtGroupName.Size = New System.Drawing.Size(142, 20)
        Me.txtGroupName.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.txtGroupName, "Group Name")
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(346, 128)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(70, 13)
        Me.Label19.TabIndex = 20
        Me.Label19.Text = "Group Name:"
        '
        'txtWordTemplate
        '
        Me.txtWordTemplate.Location = New System.Drawing.Point(96, 59)
        Me.txtWordTemplate.Name = "txtWordTemplate"
        Me.txtWordTemplate.Size = New System.Drawing.Size(429, 20)
        Me.txtWordTemplate.TabIndex = 19
        Me.ToolTip1.SetToolTip(Me.txtWordTemplate, "Word Template for document")
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(42, 59)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(57, 13)
        Me.Label18.TabIndex = 18
        Me.Label18.Text = "Template :"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dgExample
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgExample.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgExample.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgExample.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgExample.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgExample.Location = New System.Drawing.Point(96, 262)
        Me.dgExample.Name = "dgExample"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Info
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgExample.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgExample.Size = New System.Drawing.Size(721, 58)
        Me.dgExample.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.dgExample, "Example of what table header will lokk like")
        '
        'udFrom
        '
        Me.udFrom.Location = New System.Drawing.Point(266, 126)
        Me.udFrom.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.udFrom.Name = "udFrom"
        Me.udFrom.Size = New System.Drawing.Size(53, 20)
        Me.udFrom.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.udFrom, "Maximum Random number")
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(194, 128)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(66, 13)
        Me.Label17.TabIndex = 15
        Me.Label17.Text = "Select From:"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'udNoOfDonors
        '
        Me.udNoOfDonors.Location = New System.Drawing.Point(96, 126)
        Me.udNoOfDonors.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.udNoOfDonors.Name = "udNoOfDonors"
        Me.udNoOfDonors.Size = New System.Drawing.Size(53, 20)
        Me.udNoOfDonors.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.udNoOfDonors, "Default number of donors to be used")
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 128)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(96, 13)
        Me.Label16.TabIndex = 13
        Me.Label16.Text = "Number of Donors:"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtEditor
        '
        Me.txtEditor.Location = New System.Drawing.Point(644, 8)
        Me.txtEditor.Multiline = True
        Me.txtEditor.Name = "txtEditor"
        Me.txtEditor.Size = New System.Drawing.Size(100, 20)
        Me.txtEditor.TabIndex = 12
        Me.txtEditor.Visible = False
        '
        'cmdRemoveRandomHeader
        '
        Me.cmdRemoveRandomHeader.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRemoveRandomHeader.Image = Global.MedscreenCommonGui.AdcResource.Cross16
        Me.cmdRemoveRandomHeader.Location = New System.Drawing.Point(834, 343)
        Me.cmdRemoveRandomHeader.Name = "cmdRemoveRandomHeader"
        Me.cmdRemoveRandomHeader.Size = New System.Drawing.Size(22, 23)
        Me.cmdRemoveRandomHeader.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.cmdRemoveRandomHeader, "Remove header")
        Me.cmdRemoveRandomHeader.UseVisualStyleBackColor = True
        '
        'cmdAddRandomHeader
        '
        Me.cmdAddRandomHeader.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAddRandomHeader.Image = Global.MedscreenCommonGui.AdcResource.Pharmacy
        Me.cmdAddRandomHeader.Location = New System.Drawing.Point(834, 314)
        Me.cmdAddRandomHeader.Name = "cmdAddRandomHeader"
        Me.cmdAddRandomHeader.Size = New System.Drawing.Size(22, 23)
        Me.cmdAddRandomHeader.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.cmdAddRandomHeader, "Add Header")
        Me.cmdAddRandomHeader.UseVisualStyleBackColor = True
        '
        'cmdRandomSave
        '
        Me.cmdRandomSave.Image = Global.MedscreenCommonGui.AdcResource.FLOPPY_2
        Me.cmdRandomSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRandomSave.Location = New System.Drawing.Point(96, 578)
        Me.cmdRandomSave.Name = "cmdRandomSave"
        Me.cmdRandomSave.Size = New System.Drawing.Size(111, 23)
        Me.cmdRandomSave.TabIndex = 9
        Me.cmdRandomSave.Text = "Save"
        Me.cmdRandomSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdRandomSave.UseVisualStyleBackColor = True
        '
        'cmdAddGroup
        '
        Me.cmdAddGroup.Image = Global.MedscreenCommonGui.AdcResource.Pharmacy
        Me.cmdAddGroup.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddGroup.Location = New System.Drawing.Point(156, 98)
        Me.cmdAddGroup.Name = "cmdAddGroup"
        Me.cmdAddGroup.Size = New System.Drawing.Size(111, 23)
        Me.cmdAddGroup.TabIndex = 8
        Me.cmdAddGroup.Text = "Add Group"
        Me.cmdAddGroup.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAddGroup.UseVisualStyleBackColor = True
        '
        'lvSubColumns
        '
        Me.lvSubColumns.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvSubColumns.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.chDataStart})
        Me.lvSubColumns.FullRowSelect = True
        Me.lvSubColumns.Location = New System.Drawing.Point(96, 326)
        Me.lvSubColumns.Name = "lvSubColumns"
        Me.lvSubColumns.Size = New System.Drawing.Size(721, 233)
        Me.lvSubColumns.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.lvSubColumns, "Columns within selected group")
        Me.lvSubColumns.UseCompatibleStateImageBehavior = False
        Me.lvSubColumns.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Header"
        Me.ColumnHeader4.Width = 287
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Type"
        Me.ColumnHeader5.Width = 220
        '
        'lvColumnGroups
        '
        Me.lvColumnGroups.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvColumnGroups.Location = New System.Drawing.Point(96, 180)
        Me.lvColumnGroups.Name = "lvColumnGroups"
        Me.lvColumnGroups.Size = New System.Drawing.Size(721, 76)
        Me.lvColumnGroups.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.lvColumnGroups, "List of groups of columns")
        Me.lvColumnGroups.UseCompatibleStateImageBehavior = False
        '
        'udColumnSets
        '
        Me.udColumnSets.Location = New System.Drawing.Point(96, 98)
        Me.udColumnSets.Name = "udColumnSets"
        Me.udColumnSets.Size = New System.Drawing.Size(53, 20)
        Me.udColumnSets.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.udColumnSets, "Number of groups of columns")
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(5, 106)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(85, 13)
        Me.Label15.TabIndex = 4
        Me.Label15.Text = "Column Groups :"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtHeader
        '
        Me.txtHeader.Location = New System.Drawing.Point(96, 36)
        Me.txtHeader.Name = "txtHeader"
        Me.txtHeader.Size = New System.Drawing.Size(429, 20)
        Me.txtHeader.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtHeader, "Header to be used in document")
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(42, 36)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 13)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "Header :"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtRandomTitle
        '
        Me.txtRandomTitle.Location = New System.Drawing.Point(96, 13)
        Me.txtRandomTitle.Name = "txtRandomTitle"
        Me.txtRandomTitle.Size = New System.Drawing.Size(429, 20)
        Me.txtRandomTitle.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtRandomTitle, "Description of template")
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(57, 13)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(33, 13)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Title :"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
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
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "pdf"
        Me.OpenFileDialog1.Filter = "PDF|*.pdf|Word files|*.doc|Excel|*.xls|All file|*.*"
        Me.OpenFileDialog1.Title = "Select Procedure File"
        '
        'Process1
        '
        Me.Process1.StartInfo.Domain = ""
        Me.Process1.StartInfo.LoadUserProfile = False
        Me.Process1.StartInfo.Password = Nothing
        Me.Process1.StartInfo.StandardErrorEncoding = Nothing
        Me.Process1.StartInfo.StandardOutputEncoding = Nothing
        Me.Process1.StartInfo.UserName = ""
        Me.Process1.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized
        Me.Process1.SynchronizingObject = Me
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
        Me.AdrContact.ShowFindCmd = False
        Me.AdrContact.ShowNewCmd = False
        Me.AdrContact.ShowPanel = False
        Me.AdrContact.ShowSaveCmd = False
        Me.AdrContact.ShowStatus = False
        Me.AdrContact.Size = New System.Drawing.Size(576, 635)
        Me.AdrContact.TabIndex = 242
        Me.AdrContact.Tooltip = Nothing
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
        Me.cmdFindContactAdr.Image = CType(resources.GetObject("cmdFindContactAdr.Image"), System.Drawing.Image)
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
        Me.cmdNewContactAdr.Image = CType(resources.GetObject("cmdNewContactAdr.Image"), System.Drawing.Image)
        Me.cmdNewContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewContactAdr.Location = New System.Drawing.Point(64, 256)
        Me.cmdNewContactAdr.Name = "cmdNewContactAdr"
        Me.cmdNewContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewContactAdr.TabIndex = 232
        Me.cmdNewContactAdr.Text = "New Address"
        Me.cmdNewContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chDataStart
        '
        Me.chDataStart.Text = "Data Start"
        '
        'frmCollDefaults
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(896, 701)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmCollDefaults"
        Me.Text = "Collection defaults for Customer"
        Me.Panel1.ResumeLayout(False)
        CType(Me.nudNoOfDonors, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFixedSiteDetails.ResumeLayout(False)
        Me.pnlFixedSiteDetails.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.pnlContact.ResumeLayout(False)
        Me.pnlContactList.ResumeLayout(False)
        Me.tpProcedures.ResumeLayout(False)
        Me.pnlProcs.ResumeLayout(False)
        Me.PnlProcControls.ResumeLayout(False)
        Me.tpRandomNumbers.ResumeLayout(False)
        Me.tpRandomNumbers.PerformLayout()
        CType(Me.udColumnGap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgExample, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udNoOfDonors, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udColumnSets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.AdrContact.ResumeLayout(False)
        Me.AdrContact.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myCollInfo As Intranet.intranet.customerns.CustColl

    Private myClient As Intranet.intranet.customerns.Client
    Private objContTypePhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContMethodPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objRepOptionPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContStatusPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objReason As MedscreenLib.Glossary.PhraseCollection
    Private objSampleTemplates As MedscreenLib.Glossary.PhraseCollection
    Private objCOSampleTemplates As MedscreenLib.Glossary.PhraseCollection
    Private objFSSampleTemplates As MedscreenLib.Glossary.PhraseCollection
    Private objContact As contacts.CustomerContact
    Private objFixedContact As contacts.CustomerContact
    Private objAnalalyticalLabs As MedscreenLib.Glossary.PhraseCollection        'Labs that can perform the analysis

    Private blnDrawing As Boolean = True
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection info transferred into form and returns data
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CollectionInfo() As Intranet.intranet.customerns.CustColl
        Get
            Return myCollInfo
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.CustColl)
            myCollInfo = Value
            If Not Value Is Nothing Then
                Me.blnDrawing = True
                myClient = Intranet.intranet.Support.cSupport.CustList.Item(myCollInfo.CustomerID)
                'see if we have a valid customer
                If Not myClient Is Nothing Then
                    objReason = New MedscreenLib.Glossary.PhraseCollection("lib_cctool.REASONFORTESTLISTROUTINE", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , Me.myClient.Identity & ",")
                    Me.cbRandom.DataSource = objReason
                    Me.cbRandom.DisplayMember = "PhraseText"
                    Me.cbRandom.ValueMember = "PhraseID"
                    'Dim strSQuery As String = "select identity,description from Samp_Tmpl_Header s,customer_logintemplate c " & _
                    '"where s.removeflag='F' and c.template_id = s.identity and c.customer_id = '" & myClient.Identity & "'"
                    'Me.objSampleTemplates = New MedscreenLib.Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                    Dim strSQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("SampleTemplate")
                    Me.objSampleTemplates = New MedscreenLib.Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                    Me.objCOSampleTemplates = New MedscreenLib.Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                    Me.objFSSampleTemplates = New MedscreenLib.Glossary.PhraseCollection(strSQuery, MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                    Me.cbTestType.DataSource = Me.objSampleTemplates
                    Me.cbTestType.DisplayMember = "PhraseText"
                    Me.cbTestType.ValueMember = "PhraseID"
                    Me.cbCOSchedule.DataSource = Me.objCOSampleTemplates
                    Me.cbCOSchedule.DisplayMember = "PhraseText"
                    Me.cbCOSchedule.ValueMember = "PhraseID"
                    Me.cbFSSchedule.DataSource = Me.objFSSampleTemplates
                    Me.cbFSSchedule.DisplayMember = "PhraseText"
                    Me.cbFSSchedule.ValueMember = "PhraseID"

                End If
            End If
            DrawForm()
        End Set
    End Property

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CkProcedures.CheckedChanged

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Build the form display
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawForm()
        If myCollInfo Is Nothing Then Exit Sub

        Me.CkProcedures.Checked = myCollInfo.Procedures
        Me.cmdEditProcFile.Enabled = myCollInfo.Procedures
        Me.nudNoOfDonors.Value = myCollInfo.NoDonors
        Me.txtWhoToTest.Text = myCollInfo.WhoToTest
        Me.txtComments.Text = myCollInfo.AccessToSite
        'Get primary Confirmation Contact
        objContact = myClient.CustomerContacts.Item(1, customerns.CustColl.GCST_Coll_contact_conf)
        If objContact Is Nothing Then                                               'Create a contact 
            objContact = Me.myClient.CustomerContacts.CreateCustomerContact()
            objContact.ContactType = customerns.CustColl.GCST_Coll_contact_conf        'Set to right type
            objContact.Contact() = myCollInfo.ConfirmationContact
            objContact.ContactStatus = "ACTIVE"

            If objContact.Address Is Nothing Then                                   'We will need an address
                objContact.Address = New MedscreenLib.Address.Caddress()
                objContact.Address.AddressId = MedscreenLib.Address.AddressCollection.NextAddressId
                objContact.Address.AddressType = "CONT"
                objContact.Address.Update()                                         'Update the address
                objContact.AddressId = objContact.Address.AddressId
            End If
            If Medscreen.IsNumber(myCollInfo.ConfirmationAddress) Then
                objContact.Address.Fax = myCollInfo.ConfirmationAddress
                objContact.ContactMethod = "FAX"
            Else
                objContact.Address.Email = myCollInfo.ConfirmationAddress
                objContact.ContactMethod = "EMAIL"
            End If
            objContact.Address.Contact = myCollInfo.ConfirmationContact
            objContact.Contact = myCollInfo.ConfirmationContact
            objContact.Address.Update()
            objContact.Update()
        Else
            'Check contact info present
            If Not objContact.Address Is Nothing Then
                If objContact.Address.Contact.Trim.Length = 0 Then
                    objContact.Address.Contact = myCollInfo.ConfirmationContact
                    objContact.Address.Update()
                End If
            End If
            If objContact.Contact.Trim.Length = 0 Then
                objContact.Contact = myCollInfo.ConfirmationContact
                objContact.Update()
            End If
        End If

        'Get Fixed Site Contact 
        objFixedContact = myClient.CustomerContacts.Item(1, customerns.CustColl.GCST_Coll_contact_Fixed)
        If Me.CollectionInfo.FixedSiteContact.Trim.Length > 0 Then 'We have fixed site info 
            If objFixedContact Is Nothing Then
                objFixedContact = Me.myClient.CustomerContacts.CreateCustomerContact()
                objFixedContact.ContactType = customerns.CustColl.GCST_Coll_contact_Fixed          'Set to right type
                objFixedContact.Contact = myCollInfo.FixedSiteContact
                objFixedContact.ContactStatus = "ACTIVE"
                objFixedContact.Update()
                objFixedContact.Address = New MedscreenLib.Address.Caddress()
                objFixedContact.Address.AddressId = MedscreenLib.Address.AddressCollection.NextAddressId
                objFixedContact.Address.AddressType = "CONT"
                objFixedContact.Address.Contact = myCollInfo.FixedSiteContact

                objFixedContact.Address.Update()                                         'Update the address
                If Medscreen.IsNumber(myCollInfo.FixedSiteAddress) Then
                    objFixedContact.Address.Fax = myCollInfo.FixedSiteAddress
                    objFixedContact.ContactMethod = "FAX"
                Else
                    objFixedContact.Address.Email = myCollInfo.FixedSiteAddress
                    objFixedContact.ContactMethod = "EMAIL"
                End If
                objFixedContact.Address.Update()
                objFixedContact.Update()
            End If
        End If

        'Me.txtConfAdddress.Text = myCollInfo.ConfirmationAddress
        Me.ckUK.Checked = myCollInfo.UKCollection
        Me.txtCustomerID.Text = myCollInfo.CustomerID
        'Me.txtFixedSiteContact.Text = myCollInfo.FixedSiteContact
        'Me.txtFixedSiteContactPhone.Text = myCollInfo.FixedSiteAddress

        'Dim intPos As Integer

        PositionCombo(Me.cbRandom, myCollInfo.ReasonForTest, Me.objReason)
        'PositionCombo(Me.cbTestType, myCollInfo.TestType, Me.objSampleTemplates)
        PositionCombo(Me.cbConfMethod, myCollInfo.ConfirmationMethod, MedscreenLib.Glossary.Glossary.ConfMethodList)
        PositionCombo(Me.cbConfType, myCollInfo.ConfirmationType, MedscreenLib.Glossary.Glossary.ConfTypeList)

        Me.cbAnalyticalLab.SelectedValue = myCollInfo.LaboratoryId
        Me.cbProgMan.SelectedIndex = -1
        Me.cbProgMan.Text = ""
        If Not myClient Is Nothing Then
            'Dim ik As Integer
            Me.cbProgMan.SelectedValue = myClient.ProgrammeManager
            DrawComments()
            Me.drawContactList()
            Me.lblCompanyname.Text = myClient.CompanyName
            'Sort out default sample template
            If myCollInfo.SampleTemplate.Trim.Length > 0 Then
                Me.cbTestType.SelectedValue = myCollInfo.SampleTemplate
                Me.txtMatrix.Text = myClient.MatrixText(myCollInfo.SampleTemplate)
            Else
                Me.cbTestType.SelectedValue = myClient.ReportingInfo.SampleTemplate
                Me.txtMatrix.Text = myClient.MatrixText(myClient.ReportingInfo.SampleTemplate)
            End If
            If myCollInfo.COSampleTemplate.Trim.Length > 0 Then
                Me.cbCOSchedule.SelectedValue = myCollInfo.COSampleTemplate
                Me.txtCOMatrix.Text = myClient.MatrixText(myCollInfo.COSampleTemplate)
            Else
                Me.cbCOSchedule.SelectedValue = myClient.ReportingInfo.SampleTemplate
                Me.txtCOMatrix.Text = myClient.MatrixText(myClient.ReportingInfo.SampleTemplate)
            End If
            If myCollInfo.FSSampleTemplate.Trim.Length > 0 Then
                Me.cbFSSchedule.SelectedValue = myCollInfo.FSSampleTemplate
                Me.txtFSMatrix.Text = myClient.MatrixText(myCollInfo.FSSampleTemplate)
            Else
                Me.cbFSSchedule.SelectedValue = myClient.ReportingInfo.SampleTemplate
                Me.txtFSMatrix.Text = myClient.MatrixText(myClient.ReportingInfo.SampleTemplate)
            End If
            DrawProcedureFiles()
        End If

    End Sub

    Private Sub DrawProcedureFiles()
        Me.lvProcedures.Items.Clear()
        Dim strProcFiles As String = CConnection.PackageStringList("lib_collection.ProcedureFiles", myClient.Identity)
        If Not strProcFiles Is Nothing AndAlso strProcFiles.Trim.Length > 0 Then
            Dim ProcArray As String() = strProcFiles.Split(New Char() {","})
            Dim i As Integer
            For i = 0 To ProcArray.Length - 1
                Dim procFile As String = ProcArray.GetValue(i)
                If procFile.Trim.Length > 0 Then
                    Dim infoArray As String() = procFile.Split(New Char() {"|"})
                    Dim AddDoc As New Intranet.Support.AdditionalDocument(infoArray(0))
                    Dim lvItem As New ListViewItems.lvAdditionalDoc(AddDoc)
                    Me.lvProcedures.Items.Add(lvItem)
                End If
            Next
        End If
    End Sub

    Private Sub PositionCombo(ByVal Cb As ComboBox, ByVal Phrase As String, ByVal List As MedscreenLib.Glossary.PhraseCollection)

        Cb.SelectedIndex = -1
        If Phrase.Trim.Length = 0 Then
            Cb.SelectedIndex = -1
            Exit Sub
        End If
        Dim intPos As Integer

        intPos = List.IndexOf(Phrase)
        If intPos <> -1 Then
            Cb.SelectedIndex = intPos
        Else
            Dim oPh As MedscreenLib.Glossary.Phrase
            intPos = 0

            For Each oPh In List
                If oPh.PhraseText.ToUpper = Phrase Then
                    Cb.SelectedIndex = intPos
                End If
                intPos += 1
            Next
        End If
    End Sub

    Private Sub GetForm()
        If myCollInfo Is Nothing Then Exit Sub

        myCollInfo.Procedures = Me.CkProcedures.Checked
        myCollInfo.NoDonors = Me.nudNoOfDonors.Value
        myCollInfo.WhoToTest = Me.txtWhoToTest.Text
        myCollInfo.AccessToSite = Me.txtComments.Text

        ' myCollInfo.ConfirmationAddress = Me.txtConfAdddress.Text
        ' myCollInfo.ConfirmationContact = Me.TxtConfRecip.Text
        myCollInfo.UKCollection = Me.ckUK.Checked
        ' <Added by: taylor at: 22/01/2007-12:06:17 on machine: ANDREW>
        'myCollInfo.FixedSiteContact = Me.txtFixedSiteContact.Text
        'myCollInfo.FixedSiteAddress = Me.txtFixedSiteContactPhone.Text

        ' </Added by: taylor at: 22/01/2007-12:06:17 on machine: ANDREW>

        Dim oPh As MedscreenLib.Glossary.Phrase
        If Me.cbRandom.SelectedIndex = -1 Then
            myCollInfo.ReasonForTest = ""
        Else
            myCollInfo.ReasonForTest = Me.cbRandom.SelectedValue
        End If

        If Me.cbTestType.SelectedIndex = -1 Then
            myCollInfo.SampleTemplate = ""
        Else
            myCollInfo.SampleTemplate = Me.cbTestType.SelectedValue
        End If

        myCollInfo.COSampleTemplate = Me.cbCOSchedule.SelectedValue
        myCollInfo.FSSampleTemplate = Me.cbFSSchedule.SelectedValue

        If Me.cbConfMethod.SelectedIndex = -1 Then
            myCollInfo.ConfirmationMethod = ""
        Else
            oPh = MedscreenLib.Glossary.Glossary.ConfMethodList.Item(Me.cbConfMethod.SelectedIndex)
            If Not oPh Is Nothing Then myCollInfo.ConfirmationMethod = oPh.PhraseID
        End If
        myCollInfo.LaboratoryId = Me.cbAnalyticalLab.SelectedValue

        If Me.cbConfType.SelectedIndex = -1 Then
            myCollInfo.ConfirmationType = ""
        Else
            oPh = MedscreenLib.Glossary.Glossary.ConfTypeList.Item(Me.cbConfType.SelectedIndex)
            If Not oPh Is Nothing Then myCollInfo.ConfirmationType = oPh.PhraseID
        End If

        myClient.ProgrammeManager = Me.cbProgMan.Text
        myClient.CollectionInfo = myCollInfo
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        '<EhHeader>
        Try
            '</EhHeader>
            GetForm()
            Me.AddressValidation()
            myCollInfo.Update()
            myClient.Update()
            Me.DialogResult = Windows.Forms.DialogResult.OK
            '<EhFooter>
        Catch oE As System.Exception
            ' Handle exception here
            MsgBox(oE.Message)
            Me.DialogResult = Nothing
        End Try '</EhFooter>
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Need to check that the address is either a valid fax number or a valid email address
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub txtConfAdddress_leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub AddressValidation()
        'Me.ErrorProvider1.SetError(Me.txtConfAdddress, "")
        'Check to see if we are dealing with fax or email
        Me.ErrorProvider1.SetError(Me.cbConfType, "")
        Me.ErrorProvider1.SetError(Me.cbConfMethod, "")
        If Me.cbConfType.SelectedIndex <> 1 AndAlso Me.cbConfMethod.SelectedIndex = 1 Then
            Me.ErrorProvider1.SetError(Me.cbConfType, "Not a valid setting")
            Me.ErrorProvider1.SetError(Me.cbConfMethod, "Not a valid setting")
            Throw New Exception(" Invalid settings for comboboxes")



        End If

        If Not (Me.cbConfMethod.SelectedIndex = 3 Or Me.cbConfMethod.SelectedIndex = 4) Then
            Exit Sub
        End If
        'Check faxes
        If Me.cbConfMethod.SelectedIndex = 3 Then ' FAX 
            'If Not MedscreenLib.Medscreen.IsNumber(Me.txtConfAdddress.Text) OrElse Me.txtConfAdddress.Text.Trim.Length = 0 Then
            'Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Not a valid fax number")
            'Else
            Exit Sub
            'End If
        End If
        'Check emails
        If Me.cbConfMethod.SelectedIndex = 4 Then ' Email
            ' Dim emailErr As MedscreenLib.Constants.EmailErrors = MedscreenLib.Medscreen.ValidateEmail(Me.txtConfAdddress.Text)
            'If emailErr = MedscreenLib.Constants.EmailErrors.None Then Exit Sub
            'If emailErr = MedscreenLib.Constants.EmailErrors.NoAt Then Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Email address is not valid no @")
            'If emailErr = MedscreenLib.Constants.EmailErrors.NoDomain Then Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Email address is not valid no domain")
            'If emailErr = MedscreenLib.Constants.EmailErrors.NoAddress Then Me.ErrorProvider1.SetError(Me.txtConfAdddress, "Email address is missing")
            Exit Sub
        End If
        'Throw New Exception(" address must be a valid e-mail address format." + _
        '       ControlChars.Cr + "For example 'someone@microsoft.com' or Fax number")

    End Sub

    Private Sub txtConfAdddress_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            AddressValidation()
        Catch ex As Exception
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            'txtConfAdddress.Select(0, txtConfAdddress.Text.Length)
        End Try
    End Sub

    Private Sub cmdAddComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddComment.Click
        If Me.CollectionInfo Is Nothing Then Exit Sub
        Dim objCust As New Intranet.intranet.customerns.Client(Me.CollectionInfo.CustomerID, True)
        'Dim objComment As jobs.CollectionComment = objCust.CustomerComments.CreateComment("", "COLL", jobs.CollectionCommentList.CommentType.Customer)
        Dim objRet As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", , "", False)

        If Not objRet Is Nothing Then
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, objCust.Identity, _
    "COLL", objRet, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            'objComment.CommentText = objRet
            'objComment.Update()
        End If
        DrawComments()
    End Sub

    Private Sub drawContactList()
        Me.lvContacts.Items.Clear()
        Dim objContact As contacts.CustomerContact
        Dim i As Integer
        For i = 0 To Me.myClient.CustomerContacts.Count - 1
            objContact = Me.myClient.CustomerContacts.Item(i)
            If objContact.ContactType = customerns.CustColl.GCST_Coll_contact_conf Or objContact.ContactType = Me.CollectionInfo.GCST_Coll_contact_Fixed Then
                Dim objLvCont As New MedscreenCommonGui.ListViewItems.LVCustomerContact(objContact)
                Me.lvContacts.Items.Add(objLvCont)
            End If
        Next
        If Me.lvContacts.Items.Count > 0 Then
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
        If Me.myClient Is Nothing Then Exit Sub

        Dim objContact As contacts.CustomerContact = Me.myClient.CustomerContacts.CreateCustomerContact()
        objContact.ContactType = customerns.CustColl.GCST_Coll_contact_conf
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
            Else
                Me.AdrContact.Address = Nothing
                Me.cbContactType.SelectedValue = ""
                Me.cbContMethod.SelectedValue = ""
                Me.cbReportOptions.SelectedValue = ""

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
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUpdateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateContact.Click
        'Update info 
        If myLvContact Is Nothing Then Exit Sub
        myLvContact.Contact.AddressId = Me.AdrContact.Address.AddressId
        myLvContact.Contact.Address = Me.AdrContact.Address
        myLvContact.Contact.Address.Update()
        myLvContact.Contact.ContactMethod = Me.cbContMethod.SelectedValue
        myLvContact.Contact.ContactType = Me.cbContactType.SelectedValue
        myLvContact.Contact.ReportOptions = Me.cbReportOptions.SelectedValue
        myLvContact.Contact.ContactStatus = Me.cbContactStatus.SelectedValue

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
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "", Me.myClient.Identity, MedscreenLib.Constants.AddressType.AddressMain)
        Me.AdrContact.Address = addr
        Me.myLvContact.Contact.Address = addr
        Me.myLvContact.Contact.AddressId = addr.AddressId


    End Sub

    Private Sub DrawComments()
        Dim objCommList As New Comment.Collection(Comment.CommentOn.Customer, Me.CollectionInfo.CustomerID, "COLL")
        objCommList.Load()

        'Dim strQuery As String = "Select * from MED_CUSTOMER_COMMENT where customer_id = ? and COMMENT_TYPE = ? order by COMMENT_NO "
        Me.ListView1.Items.Clear()
        'Dim oRead As OleDb.OleDbDataReader
        'Dim oCmd As New OleDb.OleDbCommand()
        'Dim oColl As New Collection()
        Try ' Protecting 
            'oCmd.Connection = CConnection.DbConnection
            'oCmd.CommandText = strQuery
            'oCmd.Parameters.Add(CConnection.StringParameter("CustId", Me.CollectionInfo.CustomerID, 10))
            'oCmd.Parameters.Add(CConnection.StringParameter("Type", "COLL", 10))
            'If CConnection.ConnOpen Then            'Attempt to open reader
            '    oRead = oCmd.ExecuteReader          'Get Date
            '    While oRead.Read                    'Loop through data
            Dim objPm As Comment
            For Each objPm In objCommList
                Dim lvitem As New ListViewItems.lvComment(objPm)
                Me.ListView1.Items.Add(lvitem)
                'End While
                'End If
            Next

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmProgrammeManage-DrawComments-588")
        Finally
            'CConnection.SetConnClosed()             'Close connection
            'If Not oRead Is Nothing Then            'Try and close reader
            '    If Not oRead.IsClosed Then oRead.Close()
            'End If
        End Try
    End Sub

    Dim Currentlvitem As ListViewItems.lvComment
    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Currentlvitem = Nothing
        If Me.ListView1.SelectedItems.Count = 1 Then
            Currentlvitem = Me.ListView1.SelectedItems.Item(0)
        End If
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If Currentlvitem Is Nothing Then Exit Sub
        If Currentlvitem.Comment Is Nothing Then Exit Sub
        Dim objRet As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", , Currentlvitem.Comment.CommentText, False)
        If Not objRet Is Nothing Then
            Currentlvitem.MedComment.Text = objRet
            Currentlvitem.MedComment.Update(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
        End If
        DrawComments()
    End Sub

    Private Sub cmdFindContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindContactAdr.Click
        If Me.myLvContact Is Nothing Then Exit Sub

        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = Windows.Forms.DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                'get address and fill control
                Dim addr As MedscreenLib.Address.Caddress
                If intAddrID > 0 Then
                    addr = MedscreenLib.Glossary.Glossary.SMAddresses.item(intAddrID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                    Me.AdrContact.Address = addr
                    Me.myLvContact.Contact.Address = addr
                    Me.myLvContact.Contact.AddressId = addr.AddressId

                    'MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
                    'TODO: Need to change address links
                End If
                'make controls invisible
                'Me.cmdNewAddress.Visible = False
                'Me.cmdFindAddress.Visible = False

            End If
        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find reason for test and matrix for sample template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	07/04/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cbSampleTemplate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTestType.SelectedIndexChanged
        'Need to get matrices and sample login template
        If Me.blnDrawing Then Exit Sub
        If Me.myClient Is Nothing Then Exit Sub
        Try  ' Protecting
            If TypeOf Me.cbTestType.SelectedValue Is String Then
                Dim strTempId As String = Me.cbTestType.SelectedValue

                Dim objReason = New MedscreenLib.Glossary.PhraseCollection("lib_cctool.REASONFORTESTLISTROUTINE", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction, , , Me.myClient.Identity & "," & strTempId)
                Me.cbRandom.DataSource = objReason
                Me.cbRandom.DisplayMember = "PhraseText"
                Me.cbRandom.ValueMember = "PhraseID"
                'need to make sure we update the combo 
                Try  ' Protecting

                    If Not Me.myCollInfo Is Nothing Then
                        Me.cbRandom.SelectedValue = Me.myCollInfo.ReasonForTest
                    End If
                Catch ex As Exception
                    MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5214")
                Finally

                End Try

                Me.txtMatrix.Text = myClient.MatrixText(strTempId)
                Dim strLabId As String = CConnection.PackageStringList("lib_customer2.GetLaboratory", strTempId)
                If Not strLabId Is Nothing AndAlso strLabId.Trim.Length > 0 Then Me.cbAnalyticalLab.SelectedValue = strLabId

            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbSampleTemplate_SelectedIndexChanged-5202")
        Finally
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find matrix for Call out template change
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	07/04/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cbCoSampleTemplate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCOSchedule.SelectedIndexChanged
        'Need to get matrices and sample login template
        If Me.blnDrawing Then Exit Sub
        If Me.myClient Is Nothing Then Exit Sub
        Try  ' Protecting
            If TypeOf Me.cbCOSchedule.SelectedValue Is String Then
                Dim strTempId As String = Me.cbCOSchedule.SelectedValue
                Me.txtCOMatrix.Text = myClient.MatrixText(strTempId)
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbcoSampleTemplate_SelectedIndexChanged-5202")
        Finally
        End Try
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find matrix for change in fixed site default template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	07/04/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cbFSSampleTemplate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbFSSchedule.SelectedIndexChanged
        'Need to get matrices and sample login template
        If Me.blnDrawing Then Exit Sub
        If Me.myClient Is Nothing Then Exit Sub
        Try  ' Protecting
            If TypeOf Me.cbFSSchedule.SelectedValue Is String Then
                Dim strTempId As String = Me.cbFSSchedule.SelectedValue
                Me.txtFSMatrix.Text = myClient.MatrixText(strTempId)
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmCollectionForm-cbFSSampleTemplate_SelectedIndexChanged-5202")
        Finally
        End Try
    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawing = False
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allow user to set the procedure file.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	06/08/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAddProc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddProc.Click
        Me.OpenFileDialog1.Multiselect = False
        If Me.OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim strFilename As String = Me.OpenFileDialog1.FileName
            Dim strNewFilename As String = MedscreenLib.MedConnection.Instance.ServerPath & "\procedures\" & Me.myClient.Identity & ".doc"
            Dim msgRes As MsgBoxResult
            'First need to check the extention of the found file
            Dim strExtn As String = IO.Path.GetExtension(strFilename).ToUpper
            If InStr(strExtn, "DOC") > 0 Then
                If IO.File.Exists(strNewFilename) Then
                    msgRes = MsgBox("Do you want to overwrite the existing procedures?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    If msgRes = MsgBoxResult.Yes Then
                        'IO.File.Delete(strNewFilename)
                        'IO.File.Copy(strFilename, strNewFilename)
                        Me.CkProcedures.Checked = True
                        Me.cmdEditProcFile.Enabled = True
                        Intranet.Support.AdditionalDocumentColl.CreateCustomerDocument(myClient.Identity, strNewFilename, "PROCEDURE")
                        Me.DrawProcedureFiles()
                        Exit Sub
                    End If
                End If
            End If


            Dim newFilename As Object
            If InStr(strExtn, "PDF") > 0 Then
                strNewFilename = Medscreen.ReplaceString(strFilename, "N:\", MedscreenLib.MedConnection.Instance.ServerPath & "\")
            Else
                newFilename = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Procedure file name max 40 characters," & vbCrLf & "Suggested name - PD-SMIDPROFILE-NNN")
                If Not newFilename Is Nothing Then
                    Dim strNewFile As String = newFilename
                    If strNewFile.Length > 80 Then
                        MsgBox("File name is too long - can't continue", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                    strNewFilename = MedscreenLib.MedConnection.Instance.ServerPath & "\procedures\" & strNewFile & IO.Path.GetExtension(strFilename)
                    ' Should now have the new name
                    ' check to see if they exist
                    If IO.File.Exists(strNewFilename) Then
                        MsgBox("Sorry can't continue that filename already exists in the procedure directory", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                    'now we can do the copy and add to additional document table
                    IO.File.Copy(strFilename, strNewFilename)
                    Me.CkProcedures.Checked = True
                    Me.cmdEditProcFile.Enabled = True

                End If
            End If

            Intranet.Support.AdditionalDocumentColl.CreateCustomerDocument(myClient.Identity, strNewFilename, "PROCEDURE")
            Me.CkProcedures.Checked = True
        End If
        Me.DrawProcedureFiles()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allow editing of procedure files
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	06/08/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdEditProcFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditProcFile.Click
        If myProcedure Is Nothing Then Exit Sub

        Dim strNewFilename As String = Me.myProcedure.AdditionalDocument.Path
        Try
            Process1.StartInfo.FileName = strNewFilename
            Process1.Start()
            'Dim psInfo As New _
            '             System.Diagnostics.ProcessStartInfo _
            '             (strNewFilename)
            'psInfo.WindowStyle = _
            '   System.Diagnostics.ProcessWindowStyle.Normal
            'Dim myProcess As Process = _
            '   System.Diagnostics.Process.Start(psInfo)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
        End Try
    End Sub

    Private myProcedure As ListViewItems.lvAdditionalDoc = Nothing
    Private Sub lvProcedures_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvProcedures.SelectedIndexChanged
        myProcedure = Nothing
        Me.cmdEditProcFile.Enabled = False
        cmdDeleteProc.Enabled = False

        If lvProcedures.SelectedItems.Count = 1 Then
            myProcedure = lvProcedures.SelectedItems(0)
            cmdEditProcFile.Enabled = True
            cmdDeleteProc.Enabled = True
        End If
    End Sub

    Private Sub SetReportOptions()
        Me.cbReportOptions.Text = ""
        Me.cbReportOptions.SelectedIndex = -1
        objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("select phrase_id,phrase_text from phrase where phrase_type = '" & cbContactType.SelectedValue & "'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        cbReportOptions.DataSource = objRepOptionPhrase

    End Sub

    Private Sub cbContactType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbContactType.SelectedIndexChanged
        If (blnDrawing) Then Exit Sub
        SetReportOptions()
    End Sub


    Private Sub cmdDeleteProc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteProc.Click
        If myProcedure Is Nothing Then Exit Sub
        myProcedure.AdditionalDocument.Fields.Delete(CConnection.DbConnection)
        'Dim strFilename As String = MedscreenLib.MedConnection.Instance.ServerPath & "\" & Me.myProcedure.AdditionalDocument.Path
        'IO.File.Delete(strFilename)
        Me.DrawProcedureFiles()
    End Sub
    Private RandomNumberTemplate As RandomNumbers
    Private Sub DisplayColumn(ByVal col As RandomNumberColumn)

        udNoOfDonors.Value = col.NumberOfDonors
        udFrom.Value = col.From
        txtGroupName.Text = col.GroupDescription
        cbImportType.SelectedIndex = col.DataImportMethod
        lvSubColumns.Items.Clear()
        If col.DataImportMethod = RandomNumberColumn.DataImport.Excel Then
            lvSubColumns.Columns.Item(2).Width = 60
        Else
            lvSubColumns.Columns.Item(2).Width = 0
        End If
        For Each sb As SubColumn In col.SubColumns
            Dim lsc As New ListViewItems.lvRandomSubColumn(sb)
            lvSubColumns.Items.Add(lsc)
        Next
    End Sub

    Private Sub DisplayColumns(ByVal Rc As RandomNumbers)
        Dim i As Integer = 0
        lvColumnGroups.Items.Clear()
        udColumnSets.Value = Rc.Columns.Count
        For Each col As RandomNumberColumn In Rc.Columns
            i += 1
            lvColumnGroups.Items.Add(New ListViewItems.lvRandomColumn(col, col.GroupDescription, i))
            If i = 1 Then           'Display first column
                DisplayColumn(col)
            End If
        Next
        DrawExampleHeader()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 4 Then
            If myClient IsNot Nothing AndAlso myClient.CollectionInfo IsNot Nothing AndAlso myClient.CollectionInfo.RandomTemplate.Trim.Length > 0 Then
                Dim S As String() = [Enum].GetNames(GetType(RandomNumbers.ListTypes))
                cbRandomType.Items.Clear()
                For Each ty As String In S
                    cbRandomType.Items.Add(ty)
                Next
                cbImportType.Items.Clear()
                cbImportType.Items.AddRange([Enum].GetNames(GetType(RandomNumberColumn.DataImport)))
                RandomNumberTemplate = New RandomNumbers(myClient.CollectionInfo.RandomTemplate)
                RandomNumberTemplate.ReadFromDB()
                With RandomNumberTemplate
                    txtRandomTitle.Text = .Title
                    txtHeader.Text = .Header
                    txtWordTemplate.Text = .WordTemplate
                    udColumnGap.Value = .ColumnGapSize
                    DisplayColumns(RandomNumberTemplate)
                    cbRandomType.SelectedIndex = .ListType
                End With
            End If
        End If
    End Sub

    Private Sub cmdAddGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddGroup.Click
        If RandomNumberTemplate Is Nothing Then
            Dim strBlobID As String = ""
            If myCollInfo.RandomTemplate.Trim.Length = 0 Then
                strBlobID = CConnection.NextSequence("SEQ_BLOB_VALUES")
                myCollInfo.RandomTemplate = strBlobID
                myCollInfo.Update()
            Else
                strBlobID = myCollInfo.RandomTemplate
            End If
            RandomNumberTemplate = New RandomNumbers(strBlobID)
        End If
        RandomNumberTemplate.Columns.Add(New RandomNumberColumn)
        DisplayColumns(RandomNumberTemplate)
    End Sub

    Private Sub cmdRandomSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRandomSave.Click
        If Not RandomNumberTemplate Is Nothing Then
            With RandomNumberTemplate
                .Header = txtHeader.Text
                .Title = txtRandomTitle.Text
                .WordTemplate = txtWordTemplate.Text
            End With
            RandomNumberTemplate.WriteToDb()
        End If
    End Sub

    Private mylvColumn As ListViewItems.lvRandomColumn = Nothing
    Private Sub lvColumnGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvColumnGroups.SelectedIndexChanged
        mylvColumn = Nothing
        If lvColumnGroups.SelectedItems.Count = 1 Then
            mylvColumn = lvColumnGroups.SelectedItems(0)
            DisplayColumn(mylvColumn.Column)
        End If
    End Sub

    Private Sub cmdAddRandomHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddRandomHeader.Click
        If mylvColumn Is Nothing Then Exit Sub
        Dim subCol As New SubColumn With {.header = "Newheader", .ColType = SubColumn.ColumnType.Name}
        mylvColumn.Column.SubColumns.Add(subCol)
        DisplayColumn(mylvColumn.Column)
        DrawExampleHeader()
    End Sub
    Private myLvSubColumn As ListViewItems.lvRandomSubColumn = Nothing
    Private Sub lvSubColumns_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvSubColumns.Click
        myLvSubColumn = Nothing
        If lvSubColumns.SelectedItems.Count = 1 Then
            myLvSubColumn = lvSubColumns.SelectedItems(0)
        End If
    End Sub
    Dim lvSubItem As Integer = -1
    Private Sub lvSubColumns_SubItemClicked(ByVal sender As Object, ByVal e As ListViewEx.SubItemClickEventArgs) Handles lvSubColumns.SubItemClicked
        Dim lvItem As ListViewItem = e.Item

        lvSubItem = e.SubItem
        myLvSubColumn = Nothing
        If TypeOf lvItem Is ListViewItems.lvRandomSubColumn Then
            myLvSubColumn = lvItem
        End If

        Me.cbEditor.Items.Clear()
        Dim Editors As Control() = New Control() {Me.cbEditor, Me.txtEditor}
        If lvSubItem = 1 Then
            Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.SubColumn.ColumnType))
            For Each it As String In S
                cbEditor.Items.Add(it)
            Next
            sender.StartEditing(Editors(0), e.Item, e.SubItem)
        ElseIf lvSubItem = 0 OrElse lvSubItem = 2 Then
            sender.StartEditing(Editors(1), e.Item, e.SubItem)
        End If
    End Sub

    Private Sub cbEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbEditor.Leave
        If myLvSubColumn Is Nothing Then Exit Sub
        Dim strText As String = Me.cbEditor.Text
        Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.SubColumn.ColumnType))
        Dim intPos As Integer = Array.IndexOf(S, strText)
        myLvSubColumn.SubColumn.ColType = intPos
        DrawExampleHeader()
    End Sub

    Private Sub txtEditor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEditor.Leave
        If myLvSubColumn Is Nothing Then Exit Sub
        If lvSubItem = 0 Then
            myLvSubColumn.SubColumn.Header = txtEditor.Text
            DrawExampleHeader()
        Else
            myLvSubColumn.SubColumn.DataStart = txtEditor.Text
        End If
    End Sub

    Private Sub udNoOfDonors_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles udNoOfDonors.Leave
        If mylvColumn IsNot Nothing Then
            mylvColumn.Column.NumberOfDonors = udNoOfDonors.Value
        End If
    End Sub

    Private Sub udFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles udFrom.Leave
        If mylvColumn IsNot Nothing Then
            mylvColumn.Column.From = udFrom.Value
        End If
    End Sub

    Private Sub cmdRemoveRandomHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveRandomHeader.Click
        If mylvColumn IsNot Nothing Then
            If myLvSubColumn IsNot Nothing Then
                mylvColumn.Column.SubColumns.Remove(myLvSubColumn.SubColumn)
                DisplayColumn(mylvColumn.Column)
            End If
        End If
        DrawExampleHeader()
    End Sub

    Private Sub DrawExampleHeader()
        Try
            dgExample.Columns.Clear()
            For Each col As RandomNumberColumn In RandomNumberTemplate.Columns
                For Each subcol As SubColumn In col.SubColumns
                    Dim colHead As New DataGridViewColumn()
                    colHead.CellTemplate = New DataGridViewTextBoxCell
                    colHead.DefaultCellStyle = dgExample.DefaultCellStyle
                    colHead.HeaderText = subcol.Header
                    colHead.Width = 120
                    If subcol.ColType = SubColumn.ColumnType.Index Then
                        colHead.ReadOnly = True
                        colHead.CellTemplate.Style.BackColor = Drawing.Color.AntiqueWhite
                    End If

                    dgExample.Columns.Add(colHead)
                Next
                If RandomNumberTemplate.ColumnGapSize > 0 Then
                    Dim colHead As New DataGridViewColumn()
                    colHead.CellTemplate = New DataGridViewTextBoxCell
                    colHead.DefaultCellStyle = dgExample.DefaultCellStyle
                    colHead.HeaderText = ""
                    colHead.Width = RandomNumberTemplate.ColumnGapSize
                    colHead.ReadOnly = True
                    colHead.CellTemplate.Style.BackColor = Drawing.Color.AntiqueWhite
                    dgExample.Columns.Add(colHead)
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtGroupName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupName.Leave
        If mylvColumn Is Nothing Then Exit Sub
        mylvColumn.Column.GroupDescription = txtGroupName.Text
        mylvColumn.Text = txtGroupName.Text
    End Sub

    Private Sub udColumnGap_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles udColumnGap.Leave
        If Not RandomNumberTemplate Is Nothing Then
            RandomNumberTemplate.ColumnGapSize = udColumnGap.Value
            DrawExampleHeader()
        End If
    End Sub

    Private Sub cbImportType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbImportType.SelectedIndexChanged
        If mylvColumn Is Nothing Then Exit Sub
        mylvColumn.Column.DataImportMethod = cbImportType.SelectedIndex
        DisplayColumn(mylvColumn.Column)
    End Sub
End Class
