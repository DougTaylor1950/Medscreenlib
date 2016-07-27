Imports System.Windows.Forms
Imports System.Drawing
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmReportSchedule
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for maintaing a report schedule 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action>
''' $Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-30 07:24:09+00 $ <para/>
'''$Log: frmReportSchedule.vb,v $
'''Revision 1.0  2005-11-30 07:24:09+00  taylor
'''Commented
''' <para/>
''' </Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
''' 
Imports MedscreenLib
Public Class frmReportSchedule
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Private SendOptions As New MedscreenLib.Glossary.PhraseCollection("SENDOPTION")
    Private myCrystalReports As MedscreenLib.Glossary.PhraseCollection
    'Dim TT2 As New MedscreenLib.CBalloonToolTip() '//Demo for mouse over tooltip
    Private Printers As New MedscreenLib.Glossary.PrinterColl()
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdAddRecipient As System.Windows.Forms.Button
    Friend WithEvents AddGroup As System.Windows.Forms.Button
    Friend WithEvents CkPhrase As System.Windows.Forms.CheckBox
    Friend WithEvents pnlPhrase As System.Windows.Forms.Panel
    Friend WithEvents cbPhraseList As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Friend WithEvents ctxReports As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EditReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloneReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtBCC As System.Windows.Forms.TextBox
    Private ValueTypeList As New MedscreenLib.Glossary.PhraseCollection("PARAMTYPE")
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()


        'Add any initialization after the InitializeComponent() call

        Me.cbParamType.DataSource = myReportParamTypes
        Me.cbParamType.DisplayMember = "PhraseText"
        Me.cbParamType.ValueMember = "PhraseID"

        ValueTypeList.Load()
        Me.cbValueTypes.DataSource = ValueTypeList
        Me.cbValueTypes.DisplayMember = "PhraseText"
        Me.cbValueTypes.ValueMember = "PhraseID"
        Me.cbPhraseList.DataSource = ValueTypeList
        Me.cbPhraseList.DisplayMember = "PhraseText"
        Me.cbPhraseList.ValueMember = "PhraseID"


        Me.cbParamField.DataSource = [Enum].GetNames(GetType(MedscreenLib.CRFormulaItem.FieldTypes))
        Me.cbParamName.DataSource = MedscreenLib.Glossary.FormulaTypes.Info
        Me.cbParamName.DisplayMember = "PhraseText"
        Me.cbParamName.ValueMember = "PhraseID"

        SendOptions.Load()
        Me.cbMethod.DataSource = SendOptions
        Me.cbMethod.DisplayMember = "PhraseText"
        Me.cbMethod.ValueMember = "PhraseID"

        myCrystalReports = New MedscreenLib.Glossary.PhraseCollection("select identity,menutext from crystal_reports", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        Dim NewRepPhrase As MedscreenLib.Glossary.Phrase = New MedscreenLib.Glossary.Phrase()
        NewRepPhrase.PhraseID = "NEW"
        NewRepPhrase.PhraseText = "Add a new report"
        myCrystalReports.Add(NewRepPhrase)
        Me.cbCrystalRep.DataSource = myCrystalReports
        Me.cbCrystalRep.DisplayMember = "PhraseText"
        Me.cbCrystalRep.ValueMember = "PhraseID"

        'TT2.Style = CBalloonToolTip.ttStyleEnum.TTBalloon
        'TT2.CreateToolTip(Me.txtReport.Handle.ToInt32, "Double click to edit report details")

        'TT2.Icon(Me.txtReport.Handle.ToInt32) = CBalloonToolTip.ttIconType.TTIconInfo
        'TT2.Title(Me.txtReport.Handle.ToInt32) = "Information"
        'TT2.PopupOnDemand = False
        'TT2.VisibleTime(Me.txtReport.Handle.ToInt32) = 6000 'After 6 Seconds tooltip will go away

        'Dim myFont As New Font("Tahoma", 10, FontStyle.Italic Or FontStyle.Underline)
        'TT2.TipFont(Me.txtReport.Handle.ToInt32) = myFont
        'TT2.CreateToolTip(Me.cbCrystalRep.Handle.ToInt32, "Select report to use or use add item at end")

        'TT2.Icon(Me.cbCrystalRep.Handle.ToInt32) = CBalloonToolTip.ttIconType.TTIconInfo
        'TT2.Title(Me.cbCrystalRep.Handle.ToInt32) = "Information"
        'TT2.PopupOnDemand = False
        'TT2.VisibleTime(Me.cbCrystalRep.Handle.ToInt32) = 6000 'After 6 Seconds tooltip will go away
        'TT2.TipFont(Me.cbCrystalRep.Handle.ToInt32) = myFont

        Printers.Load()
        Me.cbPrinters.DataSource = Me.Printers
        Me.cbPrinters.DisplayMember = "Description"
        Me.cbPrinters.ValueMember = "LogicalName"

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtReport As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DTPNextReport As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTPUntil As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtRecipients As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbMethod As System.Windows.Forms.ComboBox
    Friend WithEvents pnlParameters As System.Windows.Forms.Panel
    Friend WithEvents lvParameters As System.Windows.Forms.ListView
    Friend WithEvents chParameter As System.Windows.Forms.ColumnHeader
    Friend WithEvents chType As System.Windows.Forms.ColumnHeader
    Friend WithEvents chField As System.Windows.Forms.ColumnHeader
    Friend WithEvents chValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbParamType As System.Windows.Forms.ComboBox
    Friend WithEvents cbParamField As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtParamValue As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cbParamName As System.Windows.Forms.ComboBox
    Friend WithEvents pnlNormal As System.Windows.Forms.Panel
    Friend WithEvents CBUnits As System.Windows.Forms.ComboBox
    Friend WithEvents UDInterval As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckSpecial As System.Windows.Forms.CheckBox
    Friend WithEvents pnlSpecial As System.Windows.Forms.Panel
    Friend WithEvents cbSpecPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cbDayOfWeek As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cbRepType As System.Windows.Forms.ComboBox
    Friend WithEvents lblRunTime As System.Windows.Forms.Label
    Friend WithEvents txtRunTime As System.Windows.Forms.TextBox
    Friend WithEvents cbCrystalRep As System.Windows.Forms.ComboBox
    Friend WithEvents lblRecipients As System.Windows.Forms.Label
    Friend WithEvents cbPrinters As System.Windows.Forms.ComboBox
    Friend WithEvents ckRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents cbValueTypes As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReportSchedule))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtReport = New System.Windows.Forms.TextBox
        Me.ctxReports = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EditReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CloneReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label2 = New System.Windows.Forms.Label
        Me.DTPNextReport = New System.Windows.Forms.DateTimePicker
        Me.DTPUntil = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblRecipients = New System.Windows.Forms.Label
        Me.txtRecipients = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cbMethod = New System.Windows.Forms.ComboBox
        Me.pnlParameters = New System.Windows.Forms.Panel
        Me.cbValueTypes = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmdUpdate = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cbParamName = New System.Windows.Forms.ComboBox
        Me.txtParamValue = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cbParamField = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cbParamType = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.lvParameters = New System.Windows.Forms.ListView
        Me.chParameter = New System.Windows.Forms.ColumnHeader
        Me.chType = New System.Windows.Forms.ColumnHeader
        Me.chField = New System.Windows.Forms.ColumnHeader
        Me.chValue = New System.Windows.Forms.ColumnHeader
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.pnlNormal = New System.Windows.Forms.Panel
        Me.UDInterval = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.CBUnits = New System.Windows.Forms.ComboBox
        Me.ckSpecial = New System.Windows.Forms.CheckBox
        Me.pnlSpecial = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.cbDayOfWeek = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cbSpecPeriod = New System.Windows.Forms.ComboBox
        Me.pnlPhrase = New System.Windows.Forms.Panel
        Me.cbPhraseList = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.lblType = New System.Windows.Forms.Label
        Me.cbRepType = New System.Windows.Forms.ComboBox
        Me.lblRunTime = New System.Windows.Forms.Label
        Me.txtRunTime = New System.Windows.Forms.TextBox
        Me.cbCrystalRep = New System.Windows.Forms.ComboBox
        Me.cbPrinters = New System.Windows.Forms.ComboBox
        Me.ckRemoved = New System.Windows.Forms.CheckBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAddRecipient = New System.Windows.Forms.Button
        Me.AddGroup = New System.Windows.Forms.Button
        Me.txtCustomerID = New System.Windows.Forms.TextBox
        Me.CkPhrase = New System.Windows.Forms.CheckBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtBCC = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.ctxReports.SuspendLayout()
        Me.pnlParameters.SuspendLayout()
        Me.pnlNormal.SuspendLayout()
        CType(Me.UDInterval, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSpecial.SuspendLayout()
        Me.pnlPhrase.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 473)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(940, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(852, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 22
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(764, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 21
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Report"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtReport
        '
        Me.txtReport.BackColor = System.Drawing.SystemColors.Control
        Me.txtReport.ContextMenuStrip = Me.ctxReports
        Me.txtReport.Location = New System.Drawing.Point(112, 24)
        Me.txtReport.Name = "txtReport"
        Me.txtReport.Size = New System.Drawing.Size(136, 20)
        Me.txtReport.TabIndex = 2
        '
        'ctxReports
        '
        Me.ctxReports.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditReportToolStripMenuItem, Me.CloneReportToolStripMenuItem})
        Me.ctxReports.Name = "ctxReports"
        Me.ctxReports.Size = New System.Drawing.Size(146, 48)
        '
        'EditReportToolStripMenuItem
        '
        Me.EditReportToolStripMenuItem.Name = "EditReportToolStripMenuItem"
        Me.EditReportToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.EditReportToolStripMenuItem.Text = "&Edit report"
        '
        'CloneReportToolStripMenuItem
        '
        Me.CloneReportToolStripMenuItem.Name = "CloneReportToolStripMenuItem"
        Me.CloneReportToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.CloneReportToolStripMenuItem.Text = "&Clone report"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 23)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Next Report on"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTPNextReport
        '
        Me.DTPNextReport.CustomFormat = "dd MMM yyyy  HH:mm"
        Me.DTPNextReport.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPNextReport.Location = New System.Drawing.Point(112, 56)
        Me.DTPNextReport.Name = "DTPNextReport"
        Me.DTPNextReport.Size = New System.Drawing.Size(136, 20)
        Me.DTPNextReport.TabIndex = 4
        '
        'DTPUntil
        '
        Me.DTPUntil.Checked = False
        Me.DTPUntil.Location = New System.Drawing.Point(112, 88)
        Me.DTPUntil.Name = "DTPUntil"
        Me.DTPUntil.ShowCheckBox = True
        Me.DTPUntil.Size = New System.Drawing.Size(136, 20)
        Me.DTPUntil.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Until"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRecipients
        '
        Me.lblRecipients.Location = New System.Drawing.Point(8, 320)
        Me.lblRecipients.Name = "lblRecipients"
        Me.lblRecipients.Size = New System.Drawing.Size(80, 16)
        Me.lblRecipients.TabIndex = 10
        Me.lblRecipients.Text = "Recipients"
        Me.lblRecipients.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRecipients
        '
        Me.txtRecipients.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtRecipients.Location = New System.Drawing.Point(104, 312)
        Me.txtRecipients.Multiline = True
        Me.txtRecipients.Name = "txtRecipients"
        Me.txtRecipients.Size = New System.Drawing.Size(304, 125)
        Me.txtRecipients.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(250, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(21, 23)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "by"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbMethod
        '
        Me.cbMethod.Items.AddRange(New Object() {"Email", "PDF"})
        Me.cbMethod.Location = New System.Drawing.Point(277, 57)
        Me.cbMethod.Name = "cbMethod"
        Me.cbMethod.Size = New System.Drawing.Size(121, 21)
        Me.cbMethod.TabIndex = 13
        '
        'pnlParameters
        '
        Me.pnlParameters.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlParameters.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlParameters.Controls.Add(Me.cbValueTypes)
        Me.pnlParameters.Controls.Add(Me.Label5)
        Me.pnlParameters.Controls.Add(Me.cmdUpdate)
        Me.pnlParameters.Controls.Add(Me.cmdAdd)
        Me.pnlParameters.Controls.Add(Me.cbParamName)
        Me.pnlParameters.Controls.Add(Me.txtParamValue)
        Me.pnlParameters.Controls.Add(Me.Label10)
        Me.pnlParameters.Controls.Add(Me.cbParamField)
        Me.pnlParameters.Controls.Add(Me.Label9)
        Me.pnlParameters.Controls.Add(Me.cbParamType)
        Me.pnlParameters.Controls.Add(Me.Label8)
        Me.pnlParameters.Controls.Add(Me.Label7)
        Me.pnlParameters.Controls.Add(Me.lvParameters)
        Me.pnlParameters.Controls.Add(Me.Label13)
        Me.pnlParameters.Controls.Add(Me.Label14)
        Me.pnlParameters.Controls.Add(Me.ComboBox1)
        Me.pnlParameters.Controls.Add(Me.ComboBox2)
        Me.pnlParameters.Controls.Add(Me.Label15)
        Me.pnlParameters.Location = New System.Drawing.Point(523, 8)
        Me.pnlParameters.Name = "pnlParameters"
        Me.pnlParameters.Size = New System.Drawing.Size(417, 452)
        Me.pnlParameters.TabIndex = 14
        '
        'cbValueTypes
        '
        Me.cbValueTypes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbValueTypes.Location = New System.Drawing.Point(240, 428)
        Me.cbValueTypes.Name = "cbValueTypes"
        Me.cbValueTypes.Size = New System.Drawing.Size(136, 21)
        Me.cbValueTypes.TabIndex = 28
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.Location = New System.Drawing.Point(192, 428)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 23)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "Value Types"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(248, 340)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(75, 23)
        Me.cmdUpdate.TabIndex = 26
        Me.cmdUpdate.Text = "Update"
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.Location = New System.Drawing.Point(72, 340)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 25
        Me.cmdAdd.Text = "Add"
        '
        'cbParamName
        '
        Me.cbParamName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbParamName.Location = New System.Drawing.Point(64, 372)
        Me.cbParamName.Name = "cbParamName"
        Me.cbParamName.Size = New System.Drawing.Size(104, 21)
        Me.cbParamName.TabIndex = 24
        '
        'txtParamValue
        '
        Me.txtParamValue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtParamValue.Location = New System.Drawing.Point(240, 404)
        Me.txtParamValue.Name = "txtParamValue"
        Me.txtParamValue.Size = New System.Drawing.Size(136, 20)
        Me.txtParamValue.TabIndex = 23
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label10.Location = New System.Drawing.Point(184, 404)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 23)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Value"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParamField
        '
        Me.cbParamField.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbParamField.Location = New System.Drawing.Point(64, 404)
        Me.cbParamField.Name = "cbParamField"
        Me.cbParamField.Size = New System.Drawing.Size(104, 21)
        Me.cbParamField.TabIndex = 21
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label9.Location = New System.Drawing.Point(16, 404)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 23)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "field"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParamType
        '
        Me.cbParamType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbParamType.Location = New System.Drawing.Point(256, 372)
        Me.cbParamType.Name = "cbParamType"
        Me.cbParamType.Size = New System.Drawing.Size(104, 21)
        Me.cbParamType.TabIndex = 19
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.Location = New System.Drawing.Point(208, 372)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 23)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Type"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.Location = New System.Drawing.Point(8, 372)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 23)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Name"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvParameters
        '
        Me.lvParameters.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvParameters.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chParameter, Me.chType, Me.chField, Me.chValue})
        Me.lvParameters.Location = New System.Drawing.Point(8, 8)
        Me.lvParameters.Name = "lvParameters"
        Me.lvParameters.Size = New System.Drawing.Size(397, 320)
        Me.lvParameters.TabIndex = 15
        Me.lvParameters.UseCompatibleStateImageBehavior = False
        Me.lvParameters.View = System.Windows.Forms.View.Details
        '
        'chParameter
        '
        Me.chParameter.Text = "Name"
        Me.chParameter.Width = 100
        '
        'chType
        '
        Me.chType.Text = "Type"
        '
        'chField
        '
        Me.chField.Text = "Field"
        Me.chField.Width = 120
        '
        'chValue
        '
        Me.chValue.Text = "Value"
        Me.chValue.Width = 200
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(192, 256)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 23)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "Value"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(24, 224)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(40, 23)
        Me.Label14.TabIndex = 16
        Me.Label14.Text = "Name"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(72, 224)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(104, 21)
        Me.ComboBox1.TabIndex = 24
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(72, 256)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(104, 21)
        Me.ComboBox2.TabIndex = 21
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(24, 256)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(40, 23)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "field"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlNormal
        '
        Me.pnlNormal.Controls.Add(Me.UDInterval)
        Me.pnlNormal.Controls.Add(Me.Label4)
        Me.pnlNormal.Controls.Add(Me.CBUnits)
        Me.pnlNormal.Location = New System.Drawing.Point(24, 152)
        Me.pnlNormal.Name = "pnlNormal"
        Me.pnlNormal.Size = New System.Drawing.Size(384, 72)
        Me.pnlNormal.TabIndex = 15
        '
        'UDInterval
        '
        Me.UDInterval.Location = New System.Drawing.Point(144, 25)
        Me.UDInterval.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.UDInterval.Name = "UDInterval"
        Me.UDInterval.Size = New System.Drawing.Size(48, 20)
        Me.UDInterval.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(56, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 23)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Every"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CBUnits
        '
        Me.CBUnits.Items.AddRange(New Object() {"Custom", "Month", "Day", "Year", "Week", "Hour", "Minute"})
        Me.CBUnits.Location = New System.Drawing.Point(208, 25)
        Me.CBUnits.Name = "CBUnits"
        Me.CBUnits.Size = New System.Drawing.Size(121, 21)
        Me.CBUnits.TabIndex = 12
        '
        'ckSpecial
        '
        Me.ckSpecial.Location = New System.Drawing.Point(24, 120)
        Me.ckSpecial.Name = "ckSpecial"
        Me.ckSpecial.Size = New System.Drawing.Size(108, 24)
        Me.ckSpecial.TabIndex = 16
        Me.ckSpecial.Text = "Special handling"
        '
        'pnlSpecial
        '
        Me.pnlSpecial.Controls.Add(Me.Label12)
        Me.pnlSpecial.Controls.Add(Me.cbDayOfWeek)
        Me.pnlSpecial.Controls.Add(Me.Label11)
        Me.pnlSpecial.Controls.Add(Me.cbSpecPeriod)
        Me.pnlSpecial.Location = New System.Drawing.Point(24, 144)
        Me.pnlSpecial.Name = "pnlSpecial"
        Me.pnlSpecial.Size = New System.Drawing.Size(384, 88)
        Me.pnlSpecial.TabIndex = 17
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(32, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 23)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "of the month"
        '
        'cbDayOfWeek
        '
        Me.cbDayOfWeek.Items.AddRange(New Object() {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"})
        Me.cbDayOfWeek.Location = New System.Drawing.Point(240, 16)
        Me.cbDayOfWeek.Name = "cbDayOfWeek"
        Me.cbDayOfWeek.Size = New System.Drawing.Size(121, 21)
        Me.cbDayOfWeek.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 23)
        Me.Label11.TabIndex = 1
        Me.Label11.Text = "Next report on the "
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbSpecPeriod
        '
        Me.cbSpecPeriod.Items.AddRange(New Object() {"Next", "First", "Second", "Third", "Fourth"})
        Me.cbSpecPeriod.Location = New System.Drawing.Point(128, 16)
        Me.cbSpecPeriod.Name = "cbSpecPeriod"
        Me.cbSpecPeriod.Size = New System.Drawing.Size(88, 21)
        Me.cbSpecPeriod.TabIndex = 0
        '
        'pnlPhrase
        '
        Me.pnlPhrase.Controls.Add(Me.cbPhraseList)
        Me.pnlPhrase.Controls.Add(Me.Label16)
        Me.pnlPhrase.Location = New System.Drawing.Point(24, 144)
        Me.pnlPhrase.Name = "pnlPhrase"
        Me.pnlPhrase.Size = New System.Drawing.Size(384, 88)
        Me.pnlPhrase.TabIndex = 4
        '
        'cbPhraseList
        '
        Me.cbPhraseList.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbPhraseList.DisplayMember = "PhraseText"
        Me.cbPhraseList.Location = New System.Drawing.Point(148, 33)
        Me.cbPhraseList.Name = "cbPhraseList"
        Me.cbPhraseList.Size = New System.Drawing.Size(213, 21)
        Me.cbPhraseList.TabIndex = 30
        Me.cbPhraseList.ValueMember = "PhraseID"
        '
        'Label16
        '
        Me.Label16.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label16.Location = New System.Drawing.Point(16, 30)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(106, 23)
        Me.Label16.TabIndex = 29
        Me.Label16.Text = "Schedule pattern "
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(32, 240)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(100, 23)
        Me.lblType.TabIndex = 18
        Me.lblType.Text = "Schedule Type"
        '
        'cbRepType
        '
        Me.cbRepType.Items.AddRange(New Object() {"BACKGROUND", "CUSTOMER", "SOAP"})
        Me.cbRepType.Location = New System.Drawing.Point(168, 240)
        Me.cbRepType.Name = "cbRepType"
        Me.cbRepType.Size = New System.Drawing.Size(121, 21)
        Me.cbRepType.TabIndex = 19
        '
        'lblRunTime
        '
        Me.lblRunTime.Location = New System.Drawing.Point(40, 272)
        Me.lblRunTime.Name = "lblRunTime"
        Me.lblRunTime.Size = New System.Drawing.Size(100, 23)
        Me.lblRunTime.TabIndex = 20
        Me.lblRunTime.Text = "Run Time"
        '
        'txtRunTime
        '
        Me.txtRunTime.Location = New System.Drawing.Point(168, 272)
        Me.txtRunTime.Name = "txtRunTime"
        Me.txtRunTime.Size = New System.Drawing.Size(152, 20)
        Me.txtRunTime.TabIndex = 21
        '
        'cbCrystalRep
        '
        Me.cbCrystalRep.Location = New System.Drawing.Point(277, 24)
        Me.cbCrystalRep.Name = "cbCrystalRep"
        Me.cbCrystalRep.Size = New System.Drawing.Size(121, 21)
        Me.cbCrystalRep.TabIndex = 22
        '
        'cbPrinters
        '
        Me.cbPrinters.Location = New System.Drawing.Point(104, 320)
        Me.cbPrinters.Name = "cbPrinters"
        Me.cbPrinters.Size = New System.Drawing.Size(304, 21)
        Me.cbPrinters.TabIndex = 23
        '
        'ckRemoved
        '
        Me.ckRemoved.Location = New System.Drawing.Point(309, 120)
        Me.ckRemoved.Name = "ckRemoved"
        Me.ckRemoved.Size = New System.Drawing.Size(99, 24)
        Me.ckRemoved.TabIndex = 24
        Me.ckRemoved.Text = "Removed"
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'cmdAddRecipient
        '
        Me.cmdAddRecipient.Image = Global.MedscreenCommonGui.AdcResource.Pharmacy
        Me.cmdAddRecipient.Location = New System.Drawing.Point(407, 327)
        Me.cmdAddRecipient.Name = "cmdAddRecipient"
        Me.cmdAddRecipient.Size = New System.Drawing.Size(19, 23)
        Me.cmdAddRecipient.TabIndex = 25
        Me.ToolTip1.SetToolTip(Me.cmdAddRecipient, "Add Email Recipient from Active Directory")
        Me.cmdAddRecipient.UseVisualStyleBackColor = True
        '
        'AddGroup
        '
        Me.AddGroup.Image = Global.MedscreenCommonGui.AdcResource.Pharmacy
        Me.AddGroup.Location = New System.Drawing.Point(407, 362)
        Me.AddGroup.Name = "AddGroup"
        Me.AddGroup.Size = New System.Drawing.Size(19, 23)
        Me.AddGroup.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.AddGroup, "Add Email Group from Active Directory")
        Me.AddGroup.UseVisualStyleBackColor = True
        '
        'txtCustomerID
        '
        Me.txtCustomerID.Location = New System.Drawing.Point(277, 89)
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.Size = New System.Drawing.Size(240, 20)
        Me.txtCustomerID.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.txtCustomerID, "Id Of Customer")
        '
        'CkPhrase
        '
        Me.CkPhrase.Location = New System.Drawing.Point(137, 120)
        Me.CkPhrase.Name = "CkPhrase"
        Me.CkPhrase.Size = New System.Drawing.Size(111, 24)
        Me.CkPhrase.TabIndex = 27
        Me.CkPhrase.Text = "Use Phrase"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(16, 454)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(80, 16)
        Me.Label17.TabIndex = 29
        Me.Label17.Text = "BCC"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBCC
        '
        Me.txtBCC.Location = New System.Drawing.Point(102, 453)
        Me.txtBCC.Name = "txtBCC"
        Me.txtBCC.Size = New System.Drawing.Size(401, 20)
        Me.txtBCC.TabIndex = 30
        '
        'frmReportSchedule
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(940, 513)
        Me.Controls.Add(Me.txtBCC)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.txtCustomerID)
        Me.Controls.Add(Me.CkPhrase)
        Me.Controls.Add(Me.AddGroup)
        Me.Controls.Add(Me.cmdAddRecipient)
        Me.Controls.Add(Me.ckRemoved)
        Me.Controls.Add(Me.cbPrinters)
        Me.Controls.Add(Me.cbCrystalRep)
        Me.Controls.Add(Me.txtRunTime)
        Me.Controls.Add(Me.lblRunTime)
        Me.Controls.Add(Me.cbRepType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.ckSpecial)
        Me.Controls.Add(Me.pnlParameters)
        Me.Controls.Add(Me.cbMethod)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtRecipients)
        Me.Controls.Add(Me.lblRecipients)
        Me.Controls.Add(Me.DTPUntil)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DTPNextReport)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtReport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlSpecial)
        Me.Controls.Add(Me.pnlNormal)
        Me.Controls.Add(Me.pnlPhrase)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmReportSchedule"
        Me.Text = "Report Schedule"
        Me.Panel1.ResumeLayout(False)
        Me.ctxReports.ResumeLayout(False)
        Me.pnlParameters.ResumeLayout(False)
        Me.pnlParameters.PerformLayout()
        Me.pnlNormal.ResumeLayout(False)
        CType(Me.UDInterval, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSpecial.ResumeLayout(False)
        Me.pnlPhrase.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private mySched As MedscreenLib.ReportSchedule

    Private mylvReportParameter As MedscreenCommonGui.ListViewItems.lvReportParameter
    Private myReportParamTypes As MedscreenLib.Glossary.PhraseCollection = _
        MedscreenLib.Glossary.Glossary.PhraseList.Item("RPTPARTYPE")
    'Dim TT1 As New MedscreenLib.CBalloonToolTip()

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in reporting schedule
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Schedule() As MedscreenLib.ReportSchedule
        Get
            Return mySched
        End Get
        Set(ByVal Value As MedscreenLib.ReportSchedule)
            mySched = Value
            If Not Value Is Nothing Then
                mySched.Fields.Load(CConnection.DbConnection, "rowid = '" & mySched.Fields.RowID & "'")
                DrawForm()
            End If
        End Set
    End Property

    Private Sub DrawCustomer(ByVal CustomerId As String)
        'Deal with Customer
        If CustomerId.Trim.Length > 0 Then
            Dim strProfile As String = CConnection.PackageStringList("Lib_Customer.GetSMIDProfile", CustomerId)
            If strProfile IsNot Nothing AndAlso strProfile.Trim.Length > 0 Then
                txtCustomerID.Text = strProfile
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Display the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawForm()
        With mySched
            If Not .CrystalReport Is Nothing Then
                Me.txtReport.Text = .CrystalReport.MenuText
                Me.cbCrystalRep.SelectedValue = .ReportID
            End If
            Me.DTPNextReport.Value = .NextReport
            If .UntilDate = MedscreenLib.DateField.ZeroDate Then
                Me.DTPUntil.Checked = False
            Else
                Me.DTPUntil.Value = .UntilDate
                Me.DTPUntil.Checked = True
            End If
            Me.UDInterval.Value = .RepeatInterval
            Dim strInt As String = "MDYWHI"
            If .RepeatUnits.Trim.Length = 1 Then
                Dim intUnit As Integer = InStr(strInt, .RepeatUnits.Trim)
                Me.CBUnits.SelectedIndex = intUnit
                Me.ckSpecial.Checked = False
                Me.CkPhrase.Checked = False
                Me.pnlNormal.Show()
                Me.pnlSpecial.Hide()
                Me.pnlPhrase.Hide()
                Me.pnlNormal.BringToFront()
            ElseIf .RepeatUnits.Trim.Length > 3 Then
                Me.CkPhrase.Checked = True
                Me.ckSpecial.Checked = False
                Me.pnlNormal.Hide()
                pnlPhrase.Show()
                Me.pnlSpecial.Hide()
                Me.pnlPhrase.BringToFront()
                Me.cbPhraseList.SelectedValue = .RepeatUnits
            Else
                Me.CkPhrase.Checked = False
                Me.ckSpecial.Checked = True
                Me.pnlNormal.Hide()
                pnlPhrase.Hide()
                Me.pnlSpecial.Show()
                Me.pnlSpecial.BringToFront()
                Me.CBUnits.SelectedIndex = 0
                If .RepeatUnits.Length = 0 OrElse .RepeatUnits.Chars(2) = " " Then
                    Me.cbSpecPeriod.SelectedIndex = 0
                Else
                    Me.cbSpecPeriod.SelectedIndex = Val(.RepeatUnits.Chars(2))
                End If
                If .RepeatUnits.Length >= 3 Then Me.cbDayOfWeek.SelectedIndex = Val(.RepeatUnits.Chars(3))
            End If
            Me.txtRecipients.Text = .RecipientsRaw
            txtBCC.Text = .BCC

            'Sort out sending methods 
            Me.cbMethod.SelectedValue = .SendMethod
            Me.txtRecipients.Show()
            Me.cbPrinters.Hide()


            DrawCustomer(.CustomerID)

            If .SendMethod = "PRINT" Then
                Me.cbPrinters.Show()
                Dim objPrinter As MedscreenLib.Glossary.Printer
                Dim i As Integer
                For i = 0 To Printers.Count - 1
                    objPrinter = Printers.Item(i)
                    If objPrinter.LogicalName.ToUpper = .Recipients.ToUpper Then
                        Me.cbPrinters.SelectedIndex = i
                        Exit For
                    End If
                Next
                Me.txtRecipients.Hide()
                Me.txtBCC.Hide()
            Else
                Me.txtRecipients.Show()
                Me.txtBCC.Show()
            End If
            'If .SendMethod = "PDF" Then
            '    Me.cbMethod.SelectedIndex = 1
            'Else
            '    Me.cbMethod.SelectedIndex = 0
            'End If

            Me.txtRunTime.Text = .RunPeriod
            If .ScheduleType = MedscreenLib.ReportSchedule.GCST_BACKGROUND_REPORT Then
                Me.cbRepType.SelectedIndex = 0
            ElseIf .ScheduleType = MedscreenLib.ReportSchedule.GCST_CUSTOMER_REPORT Then
                Me.cbRepType.SelectedIndex = 1
            Else
                Me.cbRepType.SelectedIndex = 2
            End If

            'Show remove flag
            Me.ckRemoved.Checked = .RemoveFlag
            FillList()

        End With
    End Sub

    Private Sub FillList()
        Dim objParam As MedscreenLib.CRFormulaItem
        Me.lvParameters.Items.Clear()
        'if we have a crystal report show its formula
        If Not mySched.CrystalReport Is Nothing Then
            For Each objParam In mySched.CrystalReport.Formulae
                Me.lvParameters.Items.Add(New MedscreenCommonGui.ListViewItems.lvReportParameter(objParam))

            Next
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the data for the schedule from the controls
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetForm() As Boolean
        Dim blnRet As Boolean = True
        If mySched Is Nothing Then Exit Function

        With mySched

            .NextReport = Me.DTPNextReport.Value
            If Me.DTPUntil.Checked Then
                .UntilDate = Me.DTPUntil.Value
            Else
                .SetUntilDateNull()
            End If
            .RepeatInterval = Me.UDInterval.Value
            Dim strInt As String = " MDYWHI"
            If Me.ckSpecial.Checked Then
                Dim strRepeat As String = "XW"
                strRepeat += CStr(Me.cbSpecPeriod.SelectedIndex).Trim
                strRepeat += CStr(Me.cbDayOfWeek.SelectedIndex).Trim
                .RepeatUnits = strRepeat
            ElseIf Me.CkPhrase.Checked Then
                .RepeatUnits = cbPhraseList.SelectedValue
            Else
                .RepeatUnits = strInt.Chars(Me.CBUnits.SelectedIndex)
            End If
            .SendMethod = Me.cbMethod.SelectedValue
            If .SendMethod = "PRINT" Then
                .RecipientsRaw = Me.cbPrinters.SelectedValue
            Else
                .RecipientsRaw = Me.txtRecipients.Text
                .BCC = Me.txtBCC.Text
            End If
            .RunPeriod = Me.txtRunTime.Text

            'Remove flag
            .RemoveFlag = Me.ckRemoved.Checked
            .ScheduleType = Me.cbRepType.Text
        End With
        Return blnRet
    End Function

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        GetForm()
    End Sub

    Private Sub lvParameters_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvParameters.SelectedIndexChanged
        mylvReportParameter = Nothing
        'Me.txtParamName.Text = ""
        If Me.lvParameters.SelectedItems.Count <> 1 Then Exit Sub
        mylvReportParameter = Me.lvParameters.SelectedItems(0)
        If Not mylvReportParameter.ReportParameter Is Nothing Then
        End If
        EditPartFill()
    End Sub

    Private Sub EditPartFill()
        Try
            If mylvReportParameter.ReportParameter.Formula.Trim.Length > 0 AndAlso mylvReportParameter.ReportParameter.Formula <> "<New>" Then
                Me.cbParamName.SelectedValue = mylvReportParameter.ReportParameter.Formula.ToUpper
            Else
                Me.cbParamName.SelectedIndex = 0
            End If
            Me.cbParamType.SelectedValue = mylvReportParameter.ReportParameter.ParamType
            If mylvReportParameter.ReportParameter.FieldName.Trim.Length > 0 Then
                Me.cbParamField.SelectedIndex = [Enum].Parse(GetType(MedscreenLib.CRFormulaItem.FieldTypes), mylvReportParameter.ReportParameter.FieldName.ToUpper)
            Else
                Me.cbParamField.SelectedIndex = 0
            End If
            Me.txtParamValue.Text = mylvReportParameter.ReportParameter.Value
        Catch
        End Try
    End Sub

    Private Sub pnlParameters_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlParameters.Paint

    End Sub

    Private Sub ckSpecial_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckSpecial.CheckedChanged, CkPhrase.CheckedChanged
        If Not TypeOf sender Is CheckBox Then
            Exit Sub
        End If
        Dim ChkBox As CheckBox = sender

        If Not ChkBox.Checked Then
            Me.pnlNormal.Show()
            Me.pnlSpecial.Hide()
            Me.pnlPhrase.Hide()
            Me.pnlNormal.BringToFront()
        Else
            If ChkBox.Name = "ckSpecial" Then
                Me.pnlNormal.Hide()
                Me.pnlPhrase.Hide()
                Me.pnlSpecial.Show()
                Me.pnlSpecial.BringToFront()
            Else
                Me.pnlNormal.Hide()
                Me.pnlSpecial.Hide()
                Me.pnlPhrase.Show()
                Me.pnlPhrase.BringToFront()
                Me.cbPhraseList.Show()
            End If
        End If

    End Sub


    Private Sub CBUnits_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CBUnits.SelectedIndexChanged
        If Me.CBUnits.SelectedIndex = 5 OrElse CBUnits.SelectedIndex = 4 Then
            Me.DTPNextReport.Format = DateTimePickerFormat.Custom
            Me.DTPNextReport.CustomFormat = "dd-MMM-yyyy HH:mm"
        ElseIf Me.CBUnits.SelectedIndex = 6 Then
            Me.DTPNextReport.Format = DateTimePickerFormat.Custom
            Me.DTPNextReport.CustomFormat = "dd-MMM-yyyy HH:mm"
        Else
            Me.DTPNextReport.Format = DateTimePickerFormat.Custom
            Me.DTPNextReport.CustomFormat = "dd-MMM-yyyy"

        End If
    End Sub

    Private Sub txtReport_DblClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReport.DoubleClick, cbCrystalRep.DoubleClick
        Dim crForm As New MedscreenCommonGui.frmCrystalReport()
        If Me.Schedule Is Nothing Then Exit Sub
        If Me.Schedule.CrystalReport Is Nothing Then Exit Sub
        crForm.Report = Me.Schedule.CrystalReport
        If crForm.ShowDialog = DialogResult.OK Then
            crForm.Report.Fields.Update(MedscreenLib.CConnection.DbConnection)
        End If

    End Sub

    Private Sub cbCrystalRep_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCrystalRep.SelectedValueChanged
        If Me.mySched Is Nothing Then Exit Sub
        If Me.cbCrystalRep.SelectedValue = "NEW" Then
            Dim objReport As MedscreenLib.CRMenuItem = New MedscreenLib.CRMenuItem()
            objReport.MenuIdentity = MedscreenLib.CConnection.NextSequence("seq_crystal_reports")
            Me.Schedule.ReportID = objReport.MenuIdentity
            Me.Schedule.CrystalReport = objReport
            Me.txtReport.Text = objReport.MenuIdentity
            objReport.MenuText = "[New Report]"
            objReport.MenuType = "RPT"
            objReport.Fields.Insert(MedscreenLib.CConnection.DbConnection)
            Me.myCrystalReports.Clear()
            myCrystalReports = New MedscreenLib.Glossary.PhraseCollection("select identity,menutext from crystal_reports", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
            Dim NewRepPhrase As MedscreenLib.Glossary.Phrase = New MedscreenLib.Glossary.Phrase()
            NewRepPhrase.PhraseID = "NEW"
            NewRepPhrase.PhraseText = "Add a new report"
            myCrystalReports.Add(NewRepPhrase)

        Else
            Me.mySched.CrystalReport = Nothing
            Me.mySched.ReportID = Me.cbCrystalRep.SelectedValue
        End If

    End Sub

    Private Sub cbMethod_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbMethod.SelectedIndexChanged
        If Me.mySched Is Nothing Then Exit Sub

        Me.txtRecipients.Show()
        Me.cbPrinters.Hide()
        If cbMethod.SelectedValue = "PRINT" Then
            Me.cbPrinters.Show()
            Me.txtRecipients.Hide()
        Else
            Me.txtRecipients.Show()
            Me.cbPrinters.Hide()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a parameter based on the values in controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Taylor]</Author><date> [23/07/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim newParameter As MedscreenLib.CRFormulaItem = New MedscreenLib.CRFormulaItem()
        newParameter.Formula = "<New>"
        mylvReportParameter = New MedscreenCommonGui.ListViewItems.lvReportParameter(newParameter)
        mylvReportParameter.ReportParameter.Identity = mySched.ReportID
        mySched.CrystalReport.Formulae.Add(mylvReportParameter.ReportParameter)
        Me.FillList()
        'EditPartFill()
    End Sub


    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        If Not mylvReportParameter.ReportParameter Is Nothing Then
            mylvReportParameter.ReportParameter.Formula = Me.cbParamName.SelectedValue ' = [Enum].Parse(GetType(MedscreenLib.CRFormulaItem.FormulaTypes), mylvReportParameter.ReportParameter.Formula.ToUpper)
            mylvReportParameter.ReportParameter.ParamType = Me.cbParamType.SelectedValue
            If Me.cbParamField.SelectedIndex > -1 Then
                mylvReportParameter.ReportParameter.FieldName = Me.cbParamField.SelectedValue
            Else
                Me.cbParamField.SelectedIndex = -1
            End If
            mylvReportParameter.ReportParameter.Value = Me.txtParamValue.Text
            mylvReportParameter.ReportParameter.Update()
            mylvReportParameter.Refresh()
        End If

    End Sub

    Private Sub txtRecipients_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRecipients.DoubleClick

    End Sub


    Private Sub cmdAddRecipient_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddRecipient.Click
        Dim frm As New ActiveDirectoryWrapper.frmFindUser
        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
            If frm.User IsNot Nothing Then
                If txtRecipients.Text.Trim.Length > 0 Then txtRecipients.Text += ","

                txtRecipients.Text += ActiveDirectoryWrapper.PC.ADWrapper.GetProperty(frm.User, "mail")
            End If
        End If
    End Sub

    Private Sub txtReport_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReport.TextChanged

    End Sub

    Private Sub AddGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddGroup.Click
        Dim Ds As DataSet = ActiveDirectoryWrapper.PC.ADWrapper.GetAllGroups
        Dim ps As New MedscreenLib.Glossary.PhraseCollection("MailGroup")
        For Each rw As DataRow In Ds.Tables(0).Rows
            Dim ph As New MedscreenLib.Glossary.Phrase()
            'ph.PhraseType = "GROUP"
            ph.PhraseID = rw.Item(0)
            ph.PhraseText = rw.Item(0)
            ps.Add(ph)
        Next
        ps.Sort()
        'Medscreen.GetParameter(Ds.Tables(0), 0, 0)
        Dim objGroup As Object = Medscreen.GetParameter(ps, "Mail Groups")
        If objGroup Is Nothing Then Exit Sub
        If txtRecipients.Text.Trim.Length > 0 Then txtRecipients.Text += ","
        txtRecipients.Text += "[" & objGroup & "]"

    End Sub

    Private Sub pnlSpecial_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlSpecial.Paint

    End Sub
    ''' <developer></developer>
    ''' <summary>
    ''' Enter Customer details
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Private Sub txtCustomerID_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerID.DoubleClick
        Dim intRet As Windows.Forms.DialogResult

        'see if we include deleted 
        'ModTpanel.FormCustSelection.IncludeDeleted = (Me.TreeView1.CurrentViewState = Treenodes.AlphaNodeList.CustomerViewState.ViewAll)

        Dim frmCust As MedscreenCommonGui.frmCustomerSelection2 = New MedscreenCommonGui.frmCustomerSelection2()
        intRet = frmCust.ShowDialog
        If intRet = Windows.Forms.DialogResult.OK Then
            If Not frmCust.Client Is Nothing Then
                mySched.CustomerID = frmCust.Client.Identity
                DrawCustomer(mySched.CustomerID)
            End If
        End If
    End Sub


    Private Sub EditReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditReportToolStripMenuItem.Click
        txtReport_DblClick(Nothing, Nothing)
    End Sub

    Private Sub CloneReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloneReportToolStripMenuItem.Click
        'Create a clone of existing report, create a new crystal report
        If Me.Schedule Is Nothing Then Exit Sub
        If Me.Schedule.CrystalReport Is Nothing Then Exit Sub
        Dim tmpReport As CRMenuItem = Me.Schedule.CrystalReport.CloneReport
        Me.Schedule.CrystalReport = tmpReport
    End Sub

   
End Class
