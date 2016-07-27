Imports Intranet.intranet
Imports Intranet.intranet.customerns
Imports MedscreenLib
Public Class frmTestPanel
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.cbStdPanels.DataSource = objStdPanels
        Me.cbStdPanels.DisplayMember = "PhraseText"
        Me.cbStdPanels.ValueMember = "PhraseID"
    End Sub

    Dim objStdPanels As New MedscreenLib.Glossary.PhraseCollection("lib_cctool.StdTestSchedList", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction, ",")
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
    Friend WithEvents pnlButtons As System.Windows.Forms.Panel
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents cmdAddTest As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteTest As System.Windows.Forms.Button
    Friend WithEvents lvTestSchedules As System.Windows.Forms.ListView
    Friend WithEvents chAnalysis As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTSDescription As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTSConf As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTSStdTest As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents TxtBreath As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cbStdPanels As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CmdCreateTestPanels As System.Windows.Forms.Button
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents lvTestPanels As System.Windows.Forms.ListView
    Friend WithEvents chTestpanelID As System.Windows.Forms.ColumnHeader
    Friend WithEvents chTPDescription As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHDefaultPanel As System.Windows.Forms.ColumnHeader
    Friend WithEvents chMatrix As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAlcoholCutoff As System.Windows.Forms.ColumnHeader
    Friend WithEvents ckAddTests As System.Windows.Forms.CheckBox
    Friend WithEvents ckInvalidPanel As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblNoBreath As System.Windows.Forms.Label
    Friend WithEvents lblMatrix As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTestPanel))
        Me.pnlButtons = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.pnlControls = New System.Windows.Forms.Panel
        Me.lblMatrix = New System.Windows.Forms.Label
        Me.lblNoBreath = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtBreath = New System.Windows.Forms.TextBox
        Me.cmdDeleteTest = New System.Windows.Forms.Button
        Me.cmdAddTest = New System.Windows.Forms.Button
        Me.lvTestSchedules = New System.Windows.Forms.ListView
        Me.chAnalysis = New System.Windows.Forms.ColumnHeader
        Me.chTSDescription = New System.Windows.Forms.ColumnHeader
        Me.chTSConf = New System.Windows.Forms.ColumnHeader
        Me.chTSStdTest = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cbStdPanels = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.CmdCreateTestPanels = New System.Windows.Forms.Button
        Me.Label43 = New System.Windows.Forms.Label
        Me.lvTestPanels = New System.Windows.Forms.ListView
        Me.chTestpanelID = New System.Windows.Forms.ColumnHeader
        Me.chTPDescription = New System.Windows.Forms.ColumnHeader
        Me.CHDefaultPanel = New System.Windows.Forms.ColumnHeader
        Me.chMatrix = New System.Windows.Forms.ColumnHeader
        Me.chAlcoholCutoff = New System.Windows.Forms.ColumnHeader
        Me.ckAddTests = New System.Windows.Forms.CheckBox
        Me.ckInvalidPanel = New System.Windows.Forms.CheckBox
        Me.pnlButtons.SuspendLayout()
        Me.pnlControls.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlButtons
        '
        Me.pnlButtons.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlButtons.Controls.Add(Me.cmdCancel)
        Me.pnlButtons.Controls.Add(Me.cmdOk)
        Me.pnlButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButtons.Location = New System.Drawing.Point(0, 705)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(968, 48)
        Me.pnlButtons.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(860, 12)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 24
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(772, 12)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 23
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlControls
        '
        Me.pnlControls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlControls.Controls.Add(Me.lblMatrix)
        Me.pnlControls.Controls.Add(Me.lblNoBreath)
        Me.pnlControls.Controls.Add(Me.Label3)
        Me.pnlControls.Controls.Add(Me.Label1)
        Me.pnlControls.Controls.Add(Me.TxtBreath)
        Me.pnlControls.Controls.Add(Me.cmdDeleteTest)
        Me.pnlControls.Controls.Add(Me.cmdAddTest)
        Me.pnlControls.Controls.Add(Me.lvTestSchedules)
        Me.pnlControls.Controls.Add(Me.Label4)
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 321)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(968, 384)
        Me.pnlControls.TabIndex = 1
        '
        'lblMatrix
        '
        Me.lblMatrix.Location = New System.Drawing.Point(8, 248)
        Me.lblMatrix.Name = "lblMatrix"
        Me.lblMatrix.Size = New System.Drawing.Size(904, 56)
        Me.lblMatrix.TabIndex = 18
        '
        'lblNoBreath
        '
        Me.lblNoBreath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNoBreath.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblNoBreath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoBreath.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblNoBreath.Location = New System.Drawing.Point(456, 328)
        Me.lblNoBreath.Name = "lblNoBreath"
        Me.lblNoBreath.Size = New System.Drawing.Size(392, 23)
        Me.lblNoBreath.TabIndex = 17
        Me.lblNoBreath.Text = "No Breath Alcohol Test"
        Me.lblNoBreath.Visible = False
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(456, 352)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(480, 23)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "µg/100ml BrAc   - Breath Alcohol Cut-off is only relevant if a Breath Alcohol tes" & _
            "t is in the panel"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(224, 348)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Breath Cutoff"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtBreath
        '
        Me.TxtBreath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtBreath.Location = New System.Drawing.Point(352, 348)
        Me.TxtBreath.Name = "TxtBreath"
        Me.TxtBreath.Size = New System.Drawing.Size(100, 20)
        Me.TxtBreath.TabIndex = 13
        '
        'cmdDeleteTest
        '
        Me.cmdDeleteTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDeleteTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDeleteTest.Location = New System.Drawing.Point(104, 348)
        Me.cmdDeleteTest.Name = "cmdDeleteTest"
        Me.cmdDeleteTest.Size = New System.Drawing.Size(88, 23)
        Me.cmdDeleteTest.TabIndex = 12
        Me.cmdDeleteTest.Text = "Delete Test"
        Me.cmdDeleteTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDeleteTest.Visible = False
        '
        'cmdAddTest
        '
        Me.cmdAddTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAddTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddTest.Location = New System.Drawing.Point(8, 348)
        Me.cmdAddTest.Name = "cmdAddTest"
        Me.cmdAddTest.Size = New System.Drawing.Size(88, 23)
        Me.cmdAddTest.TabIndex = 11
        Me.cmdAddTest.Text = "Add Test"
        Me.cmdAddTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAddTest.Visible = False
        '
        'lvTestSchedules
        '
        Me.lvTestSchedules.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTestSchedules.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chAnalysis, Me.chTSDescription, Me.chTSConf, Me.chTSStdTest, Me.ColumnHeader1})
        Me.lvTestSchedules.FullRowSelect = True
        Me.lvTestSchedules.Location = New System.Drawing.Point(0, 8)
        Me.lvTestSchedules.Name = "lvTestSchedules"
        Me.lvTestSchedules.Size = New System.Drawing.Size(1724, 216)
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
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(8, 320)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(360, 23)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Double Click on test to edit details"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cbStdPanels)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(968, 64)
        Me.Panel1.TabIndex = 2
        '
        'cbStdPanels
        '
        Me.cbStdPanels.Location = New System.Drawing.Point(112, 16)
        Me.cbStdPanels.Name = "cbStdPanels"
        Me.cbStdPanels.Size = New System.Drawing.Size(616, 21)
        Me.cbStdPanels.TabIndex = 18
        Me.cbStdPanels.Text = "ComboBox1"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Std Panels"
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Splitter1.Location = New System.Drawing.Point(0, 318)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(968, 3)
        Me.Splitter1.TabIndex = 3
        Me.Splitter1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.CmdCreateTestPanels)
        Me.Panel2.Controls.Add(Me.Label43)
        Me.Panel2.Controls.Add(Me.lvTestPanels)
        Me.Panel2.Controls.Add(Me.ckAddTests)
        Me.Panel2.Controls.Add(Me.ckInvalidPanel)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 64)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(968, 254)
        Me.Panel2.TabIndex = 4
        '
        'CmdCreateTestPanels
        '
        Me.CmdCreateTestPanels.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdCreateTestPanels.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.CmdCreateTestPanels.Location = New System.Drawing.Point(160, 216)
        Me.CmdCreateTestPanels.Name = "CmdCreateTestPanels"
        Me.CmdCreateTestPanels.Size = New System.Drawing.Size(75, 23)
        Me.CmdCreateTestPanels.TabIndex = 13
        Me.CmdCreateTestPanels.Text = "&Create"
        Me.CmdCreateTestPanels.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CmdCreateTestPanels.Visible = False
        '
        'Label43
        '
        Me.Label43.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label43.Location = New System.Drawing.Point(8, 216)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(144, 23)
        Me.Label43.TabIndex = 12
        Me.Label43.Text = "Test Schedule Entries "
        '
        'lvTestPanels
        '
        Me.lvTestPanels.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTestPanels.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chTestpanelID, Me.chTPDescription, Me.CHDefaultPanel, Me.chMatrix, Me.chAlcoholCutoff})
        Me.lvTestPanels.Location = New System.Drawing.Point(8, 8)
        Me.lvTestPanels.Name = "lvTestPanels"
        Me.lvTestPanels.Size = New System.Drawing.Size(916, 184)
        Me.lvTestPanels.TabIndex = 11
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
        Me.CHDefaultPanel.Text = "Test Type"
        '
        'chMatrix
        '
        Me.chMatrix.Text = "Standard"
        '
        'chAlcoholCutoff
        '
        Me.chAlcoholCutoff.Text = "Replicate"
        Me.chAlcoholCutoff.Width = 80
        '
        'ckAddTests
        '
        Me.ckAddTests.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ckAddTests.Location = New System.Drawing.Point(264, 216)
        Me.ckAddTests.Name = "ckAddTests"
        Me.ckAddTests.Size = New System.Drawing.Size(150, 24)
        Me.ckAddTests.TabIndex = 15
        Me.ckAddTests.Text = "&Allow additional tests"
        '
        'ckInvalidPanel
        '
        Me.ckInvalidPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ckInvalidPanel.Location = New System.Drawing.Point(440, 216)
        Me.ckInvalidPanel.Name = "ckInvalidPanel"
        Me.ckInvalidPanel.Size = New System.Drawing.Size(150, 24)
        Me.ckInvalidPanel.TabIndex = 14
        Me.ckInvalidPanel.Text = "&Allow invalid panels"
        '
        'frmTestPanel
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 753)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlControls)
        Me.Controls.Add(Me.pnlButtons)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTestPanel"
        Me.Text = "Test Panel Maintenance"
        Me.pnlButtons.ResumeLayout(False)
        Me.pnlControls.ResumeLayout(False)
        Me.pnlControls.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myClient As Intranet.intranet.customerns.Client
    'Private myCustomerTestPanel As Intranet.intranet.TestPanels.CustomerTestPanel
    Dim objTS As Intranet.intranet.TestResult.TestScheduleEntry
    Private myTestSchedule As Intranet.intranet.TestResult.TestScheduleEntry
    Private blnDrawn As Boolean = False

    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            myClient = Value
            If Not Value Is Nothing Then
                Me.TxtBreath.Text = Client.ReportingInfo.BreathCutoff
                Me.ckAddTests.Checked = Client.InvoiceInfo.AddTests
                Me.ckInvalidPanel.Checked = Client.ReportingInfo.AllowInvalidPanel
                Me.Text = "Test Panel Maintainence - " & myClient.SMIDProfile
                Me.DrawTestPanels()
                Me.DrawTestSchedules()
            End If
        End Set
    End Property

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

    Private Sub CmdCreateTestPanels_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCreateTestPanels.Click
        If Me.Client Is Nothing Then
            Exit Sub
        End If
        'If Me.myCustomerTestPanel Is Nothing Then Exit Sub
        'Me.myCustomerTestPanel.CreateTestSchedule()
        If Me.Client.TestSchedule.Count > 0 AndAlso MsgBox("This will replace any existing test panel for this customer, do you want to continue?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.No Then
            Exit Sub
        End If

        ' remove existing tests 
        Dim i As Integer
        For i = Me.Client.TestSchedule.Count - 1 To 0 Step -1
            Me.Client.TestSchedule(i).Fields.Delete(CConnection.DbConnection)
        Next

        Me.Client.TestSchedule.Clear()
        For i = 0 To Me.objStdPanel.Count - 1
            Dim objTest As New Intranet.intranet.TestResult.TestScheduleEntry()
            objTest.AnalysisID = CStr(objStdPanel(i).AnalysisID)
            objTest.StandardTest = True
            objTest.IsAnalysis = True
            objTest.OrderNumber = Me.Client.TestSchedule.Count + 1
            objTest.ReplicateCount = 1
            objTest.TestPanelID = Me.Client.TestSchedule.TestScheduleId
            objTest.Fields.Insert(CConnection.DbConnection)
        Next
        Me.Client.TestSchedule.Clear()
        Me.Client.TestSchedule.Load()
        Me.Client.TestSchedule.Renumber()
        Me.DrawTestPanels()
        Me.DrawTestSchedules()



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
        If Me.objStdPanel Is Nothing Then Exit Sub
        Me.CmdCreateTestPanels.Show()

        For Each objTS In Me.objStdPanel
            Dim objLvTSched As New MedscreenCommonGui.ListViewItems.lvTestSchedule(objTS)
            Me.lvTestPanels.Items.Add(objLvTSched)
        Next

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


        Dim noAlcoholBreath As Boolean = True
        'Me.Client.sc()
        'No test schedules then show create button
        'If Not myCustomerTestPanel Is Nothing AndAlso myCustomerTestPanel.TestSchedules.Count = 0 Then
        '    Me.CmdCreateTestPanels.Show()
        'End If

        Me.lvTestSchedules.Show()
        Me.lvTestSchedules.Items.Clear()

        'For each test schedule add it to the list view 
        For Each objTS In Me.Client.TestSchedule
            Dim objLvTSched As New MedscreenCommonGui.ListViewItems.lvTestSchedule(objTS)
            Me.lvTestSchedules.Items.Add(objLvTSched)
            If objTS.AnalysisID = "IA6990" Then
                noAlcoholBreath = False
            End If
        Next
        Me.lblNoBreath.Visible = noAlcoholBreath
        If Not Client.CollectionInfo Is Nothing Then
            Dim sTemplate As String = Client.CollectionInfo.SampleTemplate
            If sTemplate.Trim.Length = 0 Then sTemplate = "WRK_STD"
            Me.lblMatrix.Text = Me.Client.MatrixText(sTemplate)
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


    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.Client Is Nothing Then Exit Sub
        Me.Client.ReportingInfo.BreathCutoff = Me.TxtBreath.Text
        Client.InvoiceInfo.AddTests = Me.ckAddTests.Checked
        Client.ReportingInfo.AllowInvalidPanel = Me.ckInvalidPanel.Checked

        Me.Client.ReportingInfo.Update()
        Me.Client.InvoiceInfo.Update()
    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        blnDrawn = True
    End Sub

    Dim objStdPanel As Intranet.intranet.TestResult.TestScheduleEntryList
    Private Sub cbStdPanels_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStdPanels.SelectedValueChanged
        If Not blnDrawn Then Exit Sub
        objStdPanel = New Intranet.intranet.TestResult.TestScheduleEntryList(cbStdPanels.SelectedValue)
        objStdPanel.Load()
        Me.DrawTestPanels()
    End Sub

    Private Sub cbStdPanels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStdPanels.SelectedIndexChanged

    End Sub
End Class
