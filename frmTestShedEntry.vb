Public Class frmTestShedEntry
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call


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
    Friend WithEvents pnlCommands As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents chkStandardTest As System.Windows.Forms.CheckBox
    Friend WithEvents nudReplicates As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtAnalysis As System.Windows.Forms.TextBox
    Friend WithEvents txtAnalysisDesc As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTestShedEntry))
        Me.pnlCommands = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.chkStandardTest = New System.Windows.Forms.CheckBox()
        Me.nudReplicates = New System.Windows.Forms.NumericUpDown()
        Me.txtAnalysis = New System.Windows.Forms.TextBox()
        Me.txtAnalysisDesc = New System.Windows.Forms.TextBox()
        Me.pnlCommands.SuspendLayout()
        CType(Me.nudReplicates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlCommands
        '
        Me.pnlCommands.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.pnlCommands.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCommands.Location = New System.Drawing.Point(0, 237)
        Me.pnlCommands.Name = "pnlCommands"
        Me.pnlCommands.Size = New System.Drawing.Size(384, 56)
        Me.pnlCommands.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(280, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(168, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkStandardTest
        '
        Me.chkStandardTest.Location = New System.Drawing.Point(24, 40)
        Me.chkStandardTest.Name = "chkStandardTest"
        Me.chkStandardTest.TabIndex = 1
        Me.chkStandardTest.Text = "Standard Test"
        '
        'nudReplicates
        '
        Me.nudReplicates.Location = New System.Drawing.Point(24, 72)
        Me.nudReplicates.Name = "nudReplicates"
        Me.nudReplicates.TabIndex = 2
        '
        'txtAnalysis
        '
        Me.txtAnalysis.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtAnalysis.Enabled = False
        Me.txtAnalysis.Location = New System.Drawing.Point(24, 8)
        Me.txtAnalysis.Name = "txtAnalysis"
        Me.txtAnalysis.TabIndex = 3
        Me.txtAnalysis.Text = ""
        '
        'txtAnalysisDesc
        '
        Me.txtAnalysisDesc.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtAnalysisDesc.Enabled = False
        Me.txtAnalysisDesc.Location = New System.Drawing.Point(144, 8)
        Me.txtAnalysisDesc.Name = "txtAnalysisDesc"
        Me.txtAnalysisDesc.Size = New System.Drawing.Size(200, 13)
        Me.txtAnalysisDesc.TabIndex = 4
        Me.txtAnalysisDesc.Text = ""
        '
        'frmTestShedEntry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(384, 293)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtAnalysisDesc, Me.txtAnalysis, Me.nudReplicates, Me.chkStandardTest, Me.pnlCommands})
        Me.Name = "frmTestShedEntry"
        Me.Text = "Test Schedule Entry"
        Me.pnlCommands.ResumeLayout(False)
        CType(Me.nudReplicates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myTestScheduleEntry As Intranet.intranet.TestResult.TestScheduleEntry

    Public Property TestScheduleEntry() As Intranet.intranet.TestResult.TestScheduleEntry
        Get
            Return myTestScheduleEntry
        End Get
        Set(ByVal Value As Intranet.intranet.TestResult.TestScheduleEntry)
            myTestScheduleEntry = Value
            DrawForm()
        End Set
    End Property

    Private Sub DrawForm()
        If myTestScheduleEntry Is Nothing Then Exit Sub
        Me.txtAnalysis.Text = myTestScheduleEntry.AnalysisID
        Me.txtAnalysisDesc.Text = myTestScheduleEntry.AnalysisDecription
        Me.nudReplicates.Value = myTestScheduleEntry.ReplicateCount
        Me.chkStandardTest.Checked = myTestScheduleEntry.StandardTest
    End Sub

    Private Sub GetForm()
        If myTestScheduleEntry Is Nothing Then Exit Sub
        myTestScheduleEntry.ReplicateCount = Me.nudReplicates.Value
        myTestScheduleEntry.StandardTest = Me.chkStandardTest.Checked
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        GetForm()
    End Sub
End Class
