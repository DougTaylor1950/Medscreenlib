Imports system.windows.forms
Public Class frmIndustryDefaults
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.cbIndustry.DataSource = MedscreenLib.Glossary.Glossary.CustomerDefaults.IndustryDefaults
        Me.cbIndustry.DisplayMember = "Industry"
        Me.cbIndustry.ValueMember = "Industry"

        'Me.CBPanel.DataSource = Intranet.intranet.Support.cSupport.TestPanels
        Me.CBPanel.DisplayMember = "TestPanelDescription"
        Me.CBPanel.ValueMember = "TestPanelID"

        Me.ComboBox2.DataSource = Me.objOptions
        Me.ComboBox2.DisplayMember = "OptionDescription"
        Me.ComboBox2.ValueMember = "OptionID"

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
    Friend WithEvents cbIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents Industry As System.Windows.Forms.Label
    Friend WithEvents lvOptions As System.Windows.Forms.ListView
    Friend WithEvents chOption As System.Windows.Forms.ColumnHeader
    Friend WithEvents chDefault As System.Windows.Forms.ColumnHeader
    Friend WithEvents LvTestPanel As System.Windows.Forms.ListView
    Friend WithEvents CBPanel As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents chAnalysis As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUAlcohol As System.Windows.Forms.TextBox
    Friend WithEvents txtCannabis As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtBAlcohol As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chOptDesc As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAuthority As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmIndustryDefaults))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cbIndustry = New System.Windows.Forms.ComboBox()
        Me.Industry = New System.Windows.Forms.Label()
        Me.lvOptions = New System.Windows.Forms.ListView()
        Me.chOption = New System.Windows.Forms.ColumnHeader()
        Me.chDefault = New System.Windows.Forms.ColumnHeader()
        Me.LvTestPanel = New System.Windows.Forms.ListView()
        Me.CBPanel = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.chAnalysis = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtUAlcohol = New System.Windows.Forms.TextBox()
        Me.txtCannabis = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtBAlcohol = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.chOptDesc = New System.Windows.Forms.ColumnHeader()
        Me.chAuthority = New System.Windows.Forms.ColumnHeader()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 373)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(656, 40)
        Me.Panel1.TabIndex = 3
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(564, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(476, 5)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 12
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbIndustry
        '
        Me.cbIndustry.Location = New System.Drawing.Point(80, 16)
        Me.cbIndustry.Name = "cbIndustry"
        Me.cbIndustry.Size = New System.Drawing.Size(121, 21)
        Me.cbIndustry.TabIndex = 4
        Me.cbIndustry.Text = "cbIndustry "
        '
        'Industry
        '
        Me.Industry.Location = New System.Drawing.Point(24, 16)
        Me.Industry.Name = "Industry"
        Me.Industry.Size = New System.Drawing.Size(48, 23)
        Me.Industry.TabIndex = 5
        Me.Industry.Text = "Industry"
        '
        'lvOptions
        '
        Me.lvOptions.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lvOptions.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chOption, Me.chDefault, Me.chOptDesc, Me.chAuthority})
        Me.lvOptions.Location = New System.Drawing.Point(32, 56)
        Me.lvOptions.Name = "lvOptions"
        Me.lvOptions.Size = New System.Drawing.Size(600, 120)
        Me.lvOptions.TabIndex = 6
        Me.lvOptions.View = System.Windows.Forms.View.Details
        '
        'chOption
        '
        Me.chOption.Text = "Option"
        Me.chOption.Width = 100
        '
        'chDefault
        '
        Me.chDefault.Text = "Default"
        Me.chDefault.Width = 200
        '
        'LvTestPanel
        '
        Me.LvTestPanel.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.LvTestPanel.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chAnalysis, Me.ColumnHeader1})
        Me.LvTestPanel.Location = New System.Drawing.Point(32, 240)
        Me.LvTestPanel.Name = "LvTestPanel"
        Me.LvTestPanel.Size = New System.Drawing.Size(392, 121)
        Me.LvTestPanel.TabIndex = 7
        Me.LvTestPanel.View = System.Windows.Forms.View.Details
        '
        'CBPanel
        '
        Me.CBPanel.Location = New System.Drawing.Point(64, 216)
        Me.CBPanel.Name = "CBPanel"
        Me.CBPanel.Size = New System.Drawing.Size(360, 21)
        Me.CBPanel.TabIndex = 41
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 216)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "Test Panel"
        '
        'chAnalysis
        '
        Me.chAnalysis.Text = "Analysis "
        Me.chAnalysis.Width = 200
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Description"
        Me.ColumnHeader1.Width = 200
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(64, 184)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(360, 21)
        Me.ComboBox2.TabIndex = 43
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(440, 184)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 42
        Me.Button1.Text = "Add Option"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtBAlcohol, Me.Label4, Me.txtCannabis, Me.Label3, Me.txtUAlcohol, Me.Label2, Me.Label1})
        Me.Panel2.Location = New System.Drawing.Point(448, 240)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(200, 128)
        Me.Panel2.TabIndex = 45
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(64, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 45
        Me.Label1.Text = "cutoffs"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 23)
        Me.Label2.TabIndex = 46
        Me.Label2.Text = "Urine Alcohol"
        '
        'txtUAlcohol
        '
        Me.txtUAlcohol.Location = New System.Drawing.Point(112, 39)
        Me.txtUAlcohol.Name = "txtUAlcohol"
        Me.txtUAlcohol.Size = New System.Drawing.Size(32, 20)
        Me.txtUAlcohol.TabIndex = 47
        Me.txtUAlcohol.Text = ""
        '
        'txtCannabis
        '
        Me.txtCannabis.Location = New System.Drawing.Point(112, 66)
        Me.txtCannabis.Name = "txtCannabis"
        Me.txtCannabis.Size = New System.Drawing.Size(32, 20)
        Me.txtCannabis.TabIndex = 49
        Me.txtCannabis.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "Cannabis"
        '
        'txtBAlcohol
        '
        Me.txtBAlcohol.Location = New System.Drawing.Point(112, 93)
        Me.txtBAlcohol.Name = "txtBAlcohol"
        Me.txtBAlcohol.Size = New System.Drawing.Size(32, 20)
        Me.txtBAlcohol.TabIndex = 51
        Me.txtBAlcohol.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 23)
        Me.Label4.TabIndex = 50
        Me.Label4.Text = "Breath Alcohol"
        '
        'chOptDesc
        '
        Me.chOptDesc.Text = "Description"
        Me.chOptDesc.Width = 200
        '
        'chAuthority
        '
        Me.chAuthority.Text = "Authority"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 184)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Options"
        '
        'frmIndustryDefaults
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label5, Me.Panel2, Me.ComboBox2, Me.Button1, Me.CBPanel, Me.Label10, Me.LvTestPanel, Me.lvOptions, Me.Industry, Me.cbIndustry, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmIndustryDefaults"
        Me.Text = "Industry Defaults"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private objIndDef As MedscreenLib.Defaults.IndustryDefault
    'Private objTP As Intranet.intranet.TestPanels.TestPanel
    Private objOptions As New MedscreenLib.Glossary.OptionCollection()

    Private Sub DrawOptions()
        Me.lvOptions.Items.Clear()
        If objIndDef Is Nothing Then Exit Sub

        Dim objOptDef As MedscreenLib.Defaults.OptionDefault
        For Each objOptDef In objIndDef.OptionDefaults
            Dim objOption As MedscreenLib.Glossary.Options = MedscreenLib.Glossary.Glossary.Options.Item(objOptDef.OptionId)
            Dim objLvitem As New Windows.Forms.ListViewItem(objOptDef.OptionId)
            objLvitem.SubItems.Add(New Windows.Forms.ListViewItem.ListViewSubItem(objLvitem, objOptDef.OptionValue))
            Me.lvOptions.Items.Add(objLvitem)
            If Not objOption Is Nothing Then
                objLvitem.SubItems.Add(New Windows.Forms.ListViewItem.ListViewSubItem(objLvitem, objOption.OptionDescription))
                objLvitem.SubItems.Add(New Windows.Forms.ListViewItem.ListViewSubItem(objLvitem, objOption.OptionAuthority))
            End If
        Next

    End Sub

    Private Sub cbIndustry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIndustry.SelectedIndexChanged
        'objTP = Nothing
        If Me.cbIndustry.SelectedIndex > -1 Then
            objIndDef = MedscreenLib.Glossary.Glossary.CustomerDefaults.IndustryDefaults.Item(Me.cbIndustry.SelectedIndex)
            DrawOptions()
            Dim strPanel As String = objIndDef.TestPanel

            If strPanel.Trim.Length = 0 Then strPanel = "NONE"
            'objTP = Intranet.intranet.Support.cSupport.TestPanels.Item(strPanel)
            Dim intPos As Integer
            If Me.CBPanel.DataSource Is Nothing Then Exit Sub
            'If Not objTP Is Nothing Then
            '    ' intPos = Intranet.intranet.Support.cSupport.TestPanels.IndexOf(objTP)
            '    Me.CBPanel.SelectedIndex = intPos
            'End If
            DrawPanel()
            If objIndDef Is Nothing Then Exit Sub
            'do cutoffs

            Me.txtBAlcohol.Text = objIndDef.BreathAlcoholCutoff
            Me.txtCannabis.Text = objIndDef.CannabisCutoff
            Me.txtUAlcohol.Text = objIndDef.AlcoholCutoff
            Dim i As Integer

            'Me.ComboBox2.Items.Clear()                  'Clear option 
            Dim objToption As MedscreenLib.Glossary.Options

            Me.objOptions.Clear()           'Get rid of options
            For i = 0 To MedscreenLib.Glossary.Glossary.Options.Count - 1
                objToption = MedscreenLib.Glossary.Glossary.Options.Item(i)  'Get option and if not already used add it to combobox 
                If objIndDef.OptionDefaults.Item(objToption.OptionID) Is Nothing Then
                    'Me.ComboBox2.Items.Add(objToption.OptionDescription)
                    Me.objOptions.Add(objToption)

                End If
            Next
            Me.ComboBox2.DataSource = Me.objOptions
        End If
    End Sub

    Private Sub CBPanel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CBPanel.SelectedIndexChanged
        ' objTP = Nothing
        If CBPanel.SelectedIndex > -1 Then
            'Me.objTP = Intranet.intranet.Support.cSupport.TestPanels.Item(Me.CBPanel.SelectedIndex)
        End If
        DrawPanel()
    End Sub

    Private Sub DrawPanel()
        Me.LvTestPanel.Items.Clear()
        'If objTP Is Nothing Then Exit Sub

        Dim objTPI As Intranet.intranet.TestPanels.TestPanelTest
        'For Each objTPI In objTP.Tests
        '    Dim objli As New Windows.Forms.ListViewItem(objTPI.Test)
        '    Me.LvTestPanel.Items.Add(objli)
        '    If Not objTPI.ComponentView Is Nothing Then
        '        objli.SubItems.Add(New Windows.Forms.ListViewItem.ListViewSubItem(objli, objTPI.ComponentView.Name))
        '    End If
        'Next

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.ComboBox2.SelectedIndex = -1 Then Exit Sub

    End Sub
End Class
