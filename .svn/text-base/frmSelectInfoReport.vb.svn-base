Public Class frmSelectInfoReport
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        myReportlist.load()
        Dim objInfo As MedscreenLib.Glossary.InfomakerLink
        For Each objInfo In myreportlist
            Me.lvReports.Items.Add(New MedscreenCommonGui.ListViewItems.InfoReportItem(objInfo))
        Next

    End Sub

    Public ReadOnly Property Report() As MedscreenLib.Glossary.InfomakerLink
        Get
            Return Me.myReport.Report
        End Get

    End Property

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
    Friend WithEvents lvReports As System.Windows.Forms.ListView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents lvParameter As System.Windows.Forms.ListView
    Friend WithEvents chReport As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDesc As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDefault As System.Windows.Forms.ColumnHeader
    Friend WithEvents ChType As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHTable As System.Windows.Forms.ColumnHeader
    Friend WithEvents ChField As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSelectInfoReport))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lvParameter = New System.Windows.Forms.ListView()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.lvReports = New System.Windows.Forms.ListView()
        Me.chReport = New System.Windows.Forms.ColumnHeader()
        Me.chDesc = New System.Windows.Forms.ColumnHeader()
        Me.chDefault = New System.Windows.Forms.ColumnHeader()
        Me.ChType = New System.Windows.Forms.ColumnHeader()
        Me.CHTable = New System.Windows.Forms.ColumnHeader()
        Me.ChField = New System.Windows.Forms.ColumnHeader()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 469)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(776, 40)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(692, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(612, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 8
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lvParameter, Me.Splitter1, Me.lvReports})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(776, 469)
        Me.Panel2.TabIndex = 2
        '
        'lvParameter
        '
        Me.lvParameter.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chDesc, Me.chDefault, Me.ChType, Me.CHTable, Me.ChField})
        Me.lvParameter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvParameter.Location = New System.Drawing.Point(0, 235)
        Me.lvParameter.Name = "lvParameter"
        Me.lvParameter.Size = New System.Drawing.Size(776, 234)
        Me.lvParameter.TabIndex = 2
        Me.lvParameter.View = System.Windows.Forms.View.Details
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 232)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(776, 3)
        Me.Splitter1.TabIndex = 1
        Me.Splitter1.TabStop = False
        '
        'lvReports
        '
        Me.lvReports.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chReport})
        Me.lvReports.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvReports.Name = "lvReports"
        Me.lvReports.Size = New System.Drawing.Size(776, 232)
        Me.lvReports.TabIndex = 0
        Me.lvReports.View = System.Windows.Forms.View.Details
        '
        'chReport
        '
        Me.chReport.Text = "Report Description"
        Me.chReport.Width = 738
        '
        'chDesc
        '
        Me.chDesc.Text = "Description"
        Me.chDesc.Width = 200
        '
        'chDefault
        '
        Me.chDefault.Text = "Default Value"
        Me.chDefault.Width = 100
        '
        'ChType
        '
        Me.ChType.Text = "Type"
        Me.ChType.Width = 100
        '
        'CHTable
        '
        Me.CHTable.Text = "Table"
        Me.CHTable.Width = 100
        '
        'ChField
        '
        Me.ChField.Text = "Field"
        Me.ChField.Width = 100
        '
        'frmSelectInfoReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(776, 509)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1})
        Me.Name = "frmSelectInfoReport"
        Me.Text = "Select Infomaker Report"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myReportlist As New MedscreenLib.Glossary.InfomakerLinkList()
    Private myReport As MedscreenCommonGui.ListViewItems.InfoReportItem


    Private Sub lvReports_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvReports.Click
        myReport = Nothing
        Me.lvParameter.Items.Clear()

        If Me.lvReports.SelectedItems.Count = 1 Then
            myReport = Me.lvReports.SelectedItems.Item(0)
            FillParameters()

        End If
    End Sub

    Private Sub FillParameters()
        Dim objParam As MedscreenLib.Glossary.InfomakerParameter
        If myReport Is Nothing Then Exit Sub
        If myReport.Report Is Nothing Then Exit Sub
        For Each objParam In myReport.Report.Parameters
            Me.lvParameter.Items.Add(New MedscreenCommonGui.ListViewItems.InfoParameterItem(objParam))
        Next
    End Sub


End Class
