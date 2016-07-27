Public Class frmCrystalReport
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        myCrystalReports = New MedscreenLib.Glossary.PhraseCollection("select identity,menutext from crystal_reports", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        Me.cbCrystalRep.DataSource = myCrystalReports
        Me.cbCrystalRep.DisplayMember = "PhraseText"
        Me.cbCrystalRep.ValueMember = "PhraseID"

        Me.cbParamField.DataSource = [Enum].GetNames(GetType(MedscreenLib.CRFormulaItem.FieldTypes))
        Me.cbParamName.DataSource = MedscreenLib.Glossary.FormulaTypes.Info

        MyReportTypes = [Enum].GetNames(GetType(MedscreenLib.CRMenuItem.ReportTypes))

        Me.cbParamType.DataSource = myReportParamTypes
        Me.cbParamType.DisplayMember = "PhraseText"
        Me.cbParamType.ValueMember = "PhraseID"

        Me.cbParamName.DisplayMember = "PhraseText"
        Me.cbParamName.ValueMember = "PhraseID"

        Me.cbRepType.DataSource = myCrReportTypes
        Me.cbRepType.DisplayMember = "PhraseText"
        Me.cbRepType.ValueMember = "PhraseID"

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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents txtReportID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbRepType As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdFileFind As System.Windows.Forms.Button
    Friend WithEvents cbCrystalRep As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents txtParent As System.Windows.Forms.TextBox
    Friend WithEvents lvParameters As System.Windows.Forms.ListView
    Friend WithEvents chParameter As System.Windows.Forms.ColumnHeader
    Friend WithEvents chType As System.Windows.Forms.ColumnHeader
    Friend WithEvents chField As System.Windows.Forms.ColumnHeader
    Friend WithEvents chValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbParamName As System.Windows.Forms.ComboBox
    Friend WithEvents txtParamValue As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cbParamField As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cbParamType As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtServerInstance As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmdMinus As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCrystalReport))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtServerInstance = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtParent = New System.Windows.Forms.TextBox
        Me.cbCrystalRep = New System.Windows.Forms.ComboBox
        Me.cmdFileFind = New System.Windows.Forms.Button
        Me.txtPath = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbRepType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtReportID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.cmdMinus = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdUpdate = New System.Windows.Forms.Button
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 341)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(712, 40)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(624, 8)
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
        Me.cmdOk.Location = New System.Drawing.Point(536, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 21
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(16, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(688, 320)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtPassword)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.txtUserName)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.txtServerInstance)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.txtParent)
        Me.TabPage1.Controls.Add(Me.cbCrystalRep)
        Me.TabPage1.Controls.Add(Me.cmdFileFind)
        Me.TabPage1.Controls.Add(Me.txtPath)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.cbRepType)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.txtDescription)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.txtReportID)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(680, 294)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Report"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(392, 178)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(120, 20)
        Me.txtPassword.TabIndex = 30
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(280, 178)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 23)
        Me.Label12.TabIndex = 29
        Me.Label12.Text = "Password"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(128, 178)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(120, 20)
        Me.txtUserName.TabIndex = 28
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 178)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 23)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "User Name"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtServerInstance
        '
        Me.txtServerInstance.Location = New System.Drawing.Point(128, 132)
        Me.txtServerInstance.Name = "txtServerInstance"
        Me.txtServerInstance.Size = New System.Drawing.Size(120, 20)
        Me.txtServerInstance.TabIndex = 26
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 23)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Server"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtParent
        '
        Me.txtParent.Location = New System.Drawing.Point(128, 48)
        Me.txtParent.Name = "txtParent"
        Me.txtParent.Size = New System.Drawing.Size(120, 20)
        Me.txtParent.TabIndex = 24
        '
        'cbCrystalRep
        '
        Me.cbCrystalRep.Location = New System.Drawing.Point(128, 48)
        Me.cbCrystalRep.Name = "cbCrystalRep"
        Me.cbCrystalRep.Size = New System.Drawing.Size(121, 21)
        Me.cbCrystalRep.TabIndex = 23
        '
        'cmdFileFind
        '
        Me.cmdFileFind.Location = New System.Drawing.Point(520, 88)
        Me.cmdFileFind.Name = "cmdFileFind"
        Me.cmdFileFind.Size = New System.Drawing.Size(24, 23)
        Me.cmdFileFind.TabIndex = 14
        Me.cmdFileFind.Text = "..."
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(128, 88)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(384, 20)
        Me.txtPath.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Report"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbRepType
        '
        Me.cbRepType.Items.AddRange(New Object() {"BAR", "HTML", "MNU", "PROC", "RPT", "RPTX"})
        Me.cbRepType.Location = New System.Drawing.Point(384, 48)
        Me.cbRepType.Name = "cbRepType"
        Me.cbRepType.Size = New System.Drawing.Size(136, 21)
        Me.cbRepType.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(280, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Report/Menu type"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(384, 16)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(200, 20)
        Me.txtDescription.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(272, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Description"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Parent Report"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtReportID
        '
        Me.txtReportID.Enabled = False
        Me.txtReportID.Location = New System.Drawing.Point(128, 16)
        Me.txtReportID.Name = "txtReportID"
        Me.txtReportID.Size = New System.Drawing.Size(100, 20)
        Me.txtReportID.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Report ID"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.cmdMinus)
        Me.TabPage2.Controls.Add(Me.cmdAdd)
        Me.TabPage2.Controls.Add(Me.cmdUpdate)
        Me.TabPage2.Controls.Add(Me.cbParamName)
        Me.TabPage2.Controls.Add(Me.txtParamValue)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.cbParamField)
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.cbParamType)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.lvParameters)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(680, 294)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Parameters"
        '
        'cmdMinus
        '
        Me.cmdMinus.Image = CType(resources.GetObject("cmdMinus.Image"), System.Drawing.Image)
        Me.cmdMinus.Location = New System.Drawing.Point(8, 88)
        Me.cmdMinus.Name = "cmdMinus"
        Me.cmdMinus.Size = New System.Drawing.Size(24, 23)
        Me.cmdMinus.TabIndex = 35
        '
        'cmdAdd
        '
        Me.cmdAdd.Image = CType(resources.GetObject("cmdAdd.Image"), System.Drawing.Image)
        Me.cmdAdd.Location = New System.Drawing.Point(8, 48)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(24, 23)
        Me.cmdAdd.TabIndex = 34
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Image = CType(resources.GetObject("cmdUpdate.Image"), System.Drawing.Image)
        Me.cmdUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdate.Location = New System.Drawing.Point(464, 200)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(75, 23)
        Me.cmdUpdate.TabIndex = 33
        Me.cmdUpdate.Text = "Update"
        Me.cmdUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParamName
        '
        Me.cbParamName.Location = New System.Drawing.Point(464, 40)
        Me.cbParamName.Name = "cbParamName"
        Me.cbParamName.Size = New System.Drawing.Size(104, 21)
        Me.cbParamName.TabIndex = 32
        '
        'txtParamValue
        '
        Me.txtParamValue.Location = New System.Drawing.Point(464, 166)
        Me.txtParamValue.Name = "txtParamValue"
        Me.txtParamValue.Size = New System.Drawing.Size(136, 20)
        Me.txtParamValue.TabIndex = 31
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(408, 166)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 23)
        Me.Label10.TabIndex = 30
        Me.Label10.Text = "Value"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParamField
        '
        Me.cbParamField.Location = New System.Drawing.Point(464, 124)
        Me.cbParamField.Name = "cbParamField"
        Me.cbParamField.Size = New System.Drawing.Size(104, 21)
        Me.cbParamField.TabIndex = 29
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(408, 124)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 23)
        Me.Label9.TabIndex = 28
        Me.Label9.Text = "field"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbParamType
        '
        Me.cbParamType.Location = New System.Drawing.Point(464, 82)
        Me.cbParamType.Name = "cbParamType"
        Me.cbParamType.Size = New System.Drawing.Size(104, 21)
        Me.cbParamType.TabIndex = 27
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(408, 82)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 23)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Type"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(408, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 23)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Name"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvParameters
        '
        Me.lvParameters.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvParameters.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chParameter, Me.chType, Me.chField, Me.chValue})
        Me.lvParameters.Location = New System.Drawing.Point(40, 8)
        Me.lvParameters.Name = "lvParameters"
        Me.lvParameters.Size = New System.Drawing.Size(364, 268)
        Me.lvParameters.TabIndex = 16
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
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "RPT"
        Me.OpenFileDialog1.Filter = "Crystal Reports|*.rpt|All Files|*.*"
        Me.OpenFileDialog1.InitialDirectory = Global.MedscreenCommonGui.My.MySettings.Default.ReportDirectory
        Me.OpenFileDialog1.RestoreDirectory = True
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'frmCrystalReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(712, 381)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCrystalReport"
        Me.Text = "Reports"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myReport As MedscreenLib.CRMenuItem
    Private myCrystalReports As MedscreenLib.Glossary.PhraseCollection
    Private myReportType As MedscreenLib.CRMenuItem.ReportTypes
    Private MyReportTypes As String()
    Private myReportParamTypes As MedscreenLib.Glossary.PhraseCollection = _
         MedscreenLib.Glossary.Glossary.PhraseList.Item("RPTPARTYPE")
    Private myCrReportTypes As MedscreenLib.Glossary.PhraseCollection = MedscreenLib.Glossary.Glossary.PhraseList.Item("CRREPTYPE")
    'Dim TT1 As New MedscreenLib.CBalloonToolTip()

    Public Property Report() As MedscreenLib.CRMenuItem
        Get
            Return myReport
        End Get
        Set(ByVal Value As MedscreenLib.CRMenuItem)
            myReport = Value
            myReportType = myReport.ReportType
            DrawForm()
        End Set
    End Property

    Private Sub DrawForm()
        Me.txtReportID.Text = myReport.MenuIdentity
        Me.txtDescription.Text = myReport.MenuText
        Me.cbCrystalRep.SelectedValue = myReport.MenuParent


        Me.txtParent.Text = myReport.MenuParent
        Me.txtServerInstance.Text = myReport.Instance
        Me.txtUserName.Text = myReport.Username
        Me.txtPassword.Text = myReport.Password
        Dim i As Integer
        Me.cbRepType.SelectedValue = myReport.MenuType
        DrawByType()
        Me.txtPath.Text = myReport.MenuPath
        drawFormulaItems()
    End Sub

    Private Sub drawFormulaItems()
        Dim objParam As MedscreenLib.CRFormulaItem
        Me.lvParameters.Items.Clear()
        'if we have a crystal report show its formula
        If Not myReport Is Nothing Then
            For Each objParam In myReport.Formulae
                Me.lvParameters.Items.Add(New MedscreenCommonGui.ListViewItems.lvReportParameter(objParam))
            Next
        End If

    End Sub

    Private Sub DrawByType()
        If myReportType = MedscreenLib.CRMenuItem.ReportTypes.RPT Or myReport.ReportType = MedscreenLib.CRMenuItem.ReportTypes.RPTX Then
            Me.cmdFileFind.Show()
        Else
            Me.cmdFileFind.Hide()
        End If

        'sort out parents
        If myReportType = MedscreenLib.CRMenuItem.ReportTypes.BAR Or _
            myReportType = MedscreenLib.CRMenuItem.ReportTypes.MNU Or _
            myReportType = MedscreenLib.CRMenuItem.ReportTypes.RPT Then
            Me.cbCrystalRep.Show()
            Me.txtParent.Hide()
        Else
            Me.cbCrystalRep.Hide()
            Me.txtParent.Show()
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get contents of the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub getform()
        myReport.MenuPath = Me.txtPath.Text
        myReport.MenuIdentity = Me.txtReportID.Text
        myReport.MenuText = Me.txtDescription.Text
        myReport.Instance = Me.txtServerInstance.Text
        myReport.Username = Me.txtUserName.Text
        myReport.Password = Me.txtPassword.Text

        If Me.cbRepType.SelectedValue = "RPT" Or Me.cbRepType.Text = "MNU" Then
            myReport.MenuParent = Me.cbCrystalRep.SelectedValue
        Else
            myReport.MenuParent = Me.txtParent.Text
        End If
        myReport.MenuType = Me.cbRepType.SelectedValue

    End Sub

    Private Sub cmdFileFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFileFind.Click
        If Mid(Me.myReport.MenuPath, 1, 8).ToUpper = "WORKTEMP" Then
            Me.OpenFileDialog1.FileName = Mid(Me.myReport.MenuPath(), 10)
        End If
        If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim strFilename = Me.OpenFileDialog1.FileName
            Dim iPos As Integer = InStr(strFilename.toupper, "WORKTEMP") + InStr(strFilename.toupper, "CARDIFF")
            If iPos > 0 Then
                Me.txtPath.Text = Mid(strFilename, iPos)
            End If
        End If
    End Sub

    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        getform()
    End Sub

    Private Function SearchForType(ByVal _type As String) As Integer
        Dim intRet As Integer = 0
        For Each elem As String In MyReportTypes
            If elem.ToUpper = _type.ToUpper Then
                Exit For
            End If
            intRet += 1
        Next
        Return intRet
    End Function

    Private Sub cbRepType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbRepType.SelectedIndexChanged
        If myReport Is Nothing Then Exit Sub
        If cbRepType.SelectedValue Is Nothing Then Exit Sub
        Me.myReportType = SearchForType(cbRepType.SelectedValue)
        Me.DrawByType()
    End Sub

    Private mylvReportParameter As MedscreenCommonGui.ListViewItems.lvReportParameter

    Private Sub lvParameters_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvParameters.SelectedIndexChanged
        mylvReportParameter = Nothing
        'Me.txtParamName.Text = ""
        If Me.lvParameters.SelectedItems.Count <> 1 Then Exit Sub
        mylvReportParameter = Me.lvParameters.SelectedItems(0)
        If Not mylvReportParameter.ReportParameter Is Nothing Then

            Me.cbParamName.SelectedValue = mylvReportParameter.ReportParameter.Formula.ToUpper
            Me.cbParamType.SelectedValue = mylvReportParameter.ReportParameter.ParamType
            Try
                If mylvReportParameter.ReportParameter.FieldName.Trim.Length = 0 Then
                    Me.cbParamField.SelectedIndex = -1
                Else
                    Me.cbParamField.SelectedIndex = [Enum].Parse(GetType(MedscreenLib.CRFormulaItem.FieldTypes), mylvReportParameter.ReportParameter.FieldName.ToUpper)
                End If
            Catch
                Me.cbParamField.SelectedIndex = -1
            End Try
            Me.txtParamValue.Text = mylvReportParameter.ReportParameter.Value
        End If

    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        If mylvReportParameter Is Nothing Then Exit Sub
        If mylvReportParameter.ReportParameter Is Nothing Then Exit Sub
        With mylvReportParameter.ReportParameter
            .Formula = Me.cbParamName.Text
            .ParamType = Me.cbParamType.SelectedValue
            .FieldName = Me.cbParamField.Text
            .Value = Me.txtParamValue.Text
            .Update()
        End With
        mylvReportParameter.Refresh()
    End Sub

    Private Sub cmdMinus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMinus.Click
        If mylvReportParameter Is Nothing Then Exit Sub
        If mylvReportParameter.ReportParameter Is Nothing Then Exit Sub
        mylvReportParameter.ReportParameter.Delete()
        Dim intPos As Integer = myReport.Formulae.IndexOf(mylvReportParameter.ReportParameter)
        myReport.Formulae.RemoveAt(intPos)
        Me.drawFormulaItems()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim objCr As New MedscreenLib.CRFormulaItem()
        With objCr
            .Identity = myReport.MenuIdentity
            .Formula = "SENDTYPE"
            .FieldName = "SENDTYPE"
            .ParamType = "VALUE"
            .Value = "HTML"
            Dim intPos As Integer = myReport.Formulae.Add(objCr)

            objCr.Insert()
        End With
        drawFormulaItems()

    End Sub

   
End Class
