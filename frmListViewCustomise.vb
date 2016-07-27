Imports MedscreenCommonGui
Imports System.Windows.Forms
Imports System.Drawing
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : frmListViewCustomise
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for customising List Views
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action>
''' $Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-30 07:12:31+00 $ <para/>
'''$Log: frmListViewCustomise.vb,v $
'''Revision 1.0  2005-11-30 07:12:31+00  taylor
'''Commented
''' <para/>
'''</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmListViewCustomise
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Draw()

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chWidth As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdUp As System.Windows.Forms.Button
    Friend WithEvents cmdDown As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtColumnHeader As System.Windows.Forms.TextBox
    Friend WithEvents ckShow As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NUWidth As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnChange As System.Windows.Forms.Button
    Friend WithEvents ckFixed As System.Windows.Forms.CheckBox
    Friend WithEvents cmdDefaults As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmListViewCustomise))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdDefaults = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnChange = New System.Windows.Forms.Button
        Me.NUWidth = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.ckShow = New System.Windows.Forms.CheckBox
        Me.txtColumnHeader = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdDown = New System.Windows.Forms.Button
        Me.cmdUp = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.chWidth = New System.Windows.Forms.ColumnHeader
        Me.ckFixed = New System.Windows.Forms.CheckBox
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.NUWidth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdDefaults)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 225)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(600, 48)
        Me.Panel1.TabIndex = 1
        '
        'cmdDefaults
        '
        Me.cmdDefaults.Image = CType(resources.GetObject("cmdDefaults.Image"), System.Drawing.Image)
        Me.cmdDefaults.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDefaults.Location = New System.Drawing.Point(48, 16)
        Me.cmdDefaults.Name = "cmdDefaults"
        Me.cmdDefaults.Size = New System.Drawing.Size(75, 23)
        Me.cmdDefaults.TabIndex = 8
        Me.cmdDefaults.Text = "Defaults"
        Me.cmdDefaults.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(424, 13)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "&Ok"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(504, 13)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.ckFixed)
        Me.Panel2.Controls.Add(Me.btnChange)
        Me.Panel2.Controls.Add(Me.NUWidth)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.ckShow)
        Me.Panel2.Controls.Add(Me.txtColumnHeader)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.cmdDown)
        Me.Panel2.Controls.Add(Me.cmdUp)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(272, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(328, 225)
        Me.Panel2.TabIndex = 2
        '
        'btnChange
        '
        Me.btnChange.Image = CType(resources.GetObject("btnChange.Image"), System.Drawing.Image)
        Me.btnChange.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnChange.Location = New System.Drawing.Point(136, 144)
        Me.btnChange.Name = "btnChange"
        Me.btnChange.Size = New System.Drawing.Size(75, 23)
        Me.btnChange.TabIndex = 7
        Me.btnChange.Text = "Change"
        Me.btnChange.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'NUWidth
        '
        Me.NUWidth.Location = New System.Drawing.Point(136, 88)
        Me.NUWidth.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NUWidth.Minimum = New Decimal(New Integer() {2, 0, 0, -2147483648})
        Me.NUWidth.Name = "NUWidth"
        Me.NUWidth.Size = New System.Drawing.Size(120, 20)
        Me.NUWidth.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(50, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Column Width"
        '
        'ckShow
        '
        Me.ckShow.Location = New System.Drawing.Point(93, 56)
        Me.ckShow.Name = "ckShow"
        Me.ckShow.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ckShow.Size = New System.Drawing.Size(56, 24)
        Me.ckShow.TabIndex = 4
        Me.ckShow.Text = "Show"
        '
        'txtColumnHeader
        '
        Me.txtColumnHeader.Location = New System.Drawing.Point(136, 21)
        Me.txtColumnHeader.Name = "txtColumnHeader"
        Me.txtColumnHeader.Size = New System.Drawing.Size(184, 20)
        Me.txtColumnHeader.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Column Header"
        '
        'cmdDown
        '
        Me.cmdDown.Image = CType(resources.GetObject("cmdDown.Image"), System.Drawing.Image)
        Me.cmdDown.Location = New System.Drawing.Point(5, 74)
        Me.cmdDown.Name = "cmdDown"
        Me.cmdDown.Size = New System.Drawing.Size(24, 23)
        Me.cmdDown.TabIndex = 1
        '
        'cmdUp
        '
        Me.cmdUp.Image = CType(resources.GetObject("cmdUp.Image"), System.Drawing.Image)
        Me.cmdUp.Location = New System.Drawing.Point(5, 48)
        Me.cmdUp.Name = "cmdUp"
        Me.cmdUp.Size = New System.Drawing.Size(24, 23)
        Me.cmdUp.TabIndex = 0
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader3, Me.ColumnHeader2, Me.ColumnHeader1, Me.chWidth})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(272, 225)
        Me.ListView1.TabIndex = 3
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = ""
        Me.ColumnHeader3.Width = 25
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Header"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Show"
        '
        'chWidth
        '
        Me.chWidth.Text = "Width"
        '
        'ckFixed
        '
        Me.ckFixed.AutoSize = True
        Me.ckFixed.Location = New System.Drawing.Point(263, 88)
        Me.ckFixed.Name = "ckFixed"
        Me.ckFixed.Size = New System.Drawing.Size(51, 17)
        Me.ckFixed.TabIndex = 8
        Me.ckFixed.Text = "Fixed"
        Me.ckFixed.UseVisualStyleBackColor = True
        '
        'frmListViewCustomise
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 273)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmListViewCustomise"
        Me.Text = "frmList"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.NUWidth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CurrentItem As ColumnListViewItem = Nothing
    Private SelectedIndex As Integer = -1

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : frmListViewCustomise.ColumnListViewItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Class that represents a list view column as a list view item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Class ColumnListViewItem
        Inherits ListViewItem
        Private myColInfo As UserDefaults.ColumnInfo
        Private myLv As ListView

        Public Sub New()
            MyBase.new()
        End Sub

        Public Sub New(ByVal lv As ListView, ByVal ColInfo As UserDefaults.ColumnInfo)
            MyBase.New()
            myLv = lv
            myColInfo = ColInfo
            myLv.Items.Add(Me)
            Me.Text = ColInfo.Index
            Me.BackColor = SystemColors.Control
            Draw()
        End Sub

        Public Property CollumnInfo() As UserDefaults.ColumnInfo
            Get
                Return Me.myColInfo
            End Get
            Set(ByVal Value As UserDefaults.ColumnInfo)
                myColInfo = Value
            End Set
        End Property

        Public Sub Draw()
            Me.UseItemStyleForSubItems = False
            Me.SubItems.Clear()
            Dim aSub As ListViewItem.ListViewSubItem
            Me.Text = myColInfo.Index.ToString
            aSub = New ListViewItem.ListViewSubItem(Me, myColInfo.HeaderText)
            Me.SubItems.Add(aSub)
            aSub.BackColor = Color.Wheat
            If myColInfo.Show Then
                aSub = New ListViewItem.ListViewSubItem(Me, "Yes")
            Else
                aSub = New ListViewItem.ListViewSubItem(Me, "No")
            End If
            Me.SubItems.Add(aSub)
            aSub.BackColor = Color.Wheat
            If myColInfo.ColumnWidth = -1 Then
                aSub = New ListViewItem.ListViewSubItem(Me, "Auto Size")
            Else
                aSub = New ListViewItem.ListViewSubItem(Me, myColInfo.ColumnWidth)
            End If
            Me.SubItems.Add(aSub)
            aSub.BackColor = Color.Wheat


        End Sub



    End Class

    Private Sub cmdUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUp.Click
        If Me.CurrentItem Is Nothing Then Exit Sub

        If Me.CurrentItem.CollumnInfo.Index = 0 Then Exit Sub 'Can't go up from the top 

        Dim intIndex As Integer = Me.CurrentItem.CollumnInfo.Index - 1
        'Dim prevItem As ColumnListViewItem = Me.ListView1.Items(intIndex)
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Item(Me.CurrentItem.CollumnInfo.Index, UserDefaults.ColumnInfoCollection.CICIndexOn.Order).Index -= 1
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Item(intIndex, UserDefaults.ColumnInfoCollection.CICIndexOn.Order).Index += 1
        SelectedIndex = intIndex

        Draw()
    End Sub

    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Me.CurrentItem = Nothing
        Me.SelectedIndex = -1
        If Me.ListView1.SelectedItems.Count <> 1 Then Exit Sub
        Me.CurrentItem = Me.ListView1.SelectedItems.Item(0)
        If Me.CurrentItem Is Nothing Then Exit Sub
        Me.SelectedIndex = Me.CurrentItem.CollumnInfo.Index
        With Me.CurrentItem.CollumnInfo
            Me.txtColumnHeader.Text = .HeaderText
            Me.ckShow.Checked = .Show
            Me.NUWidth.Value = .ColumnWidth
            Me.NUWidth.Enabled = Not .FixedWidth
            Me.ckFixed.Checked = .FixedWidth
        End With
    End Sub

    Private Sub Draw()
        Dim aColl As UserDefaults.ColumnInfo
        Me.ListView1.BeginUpdate()
        Me.ListView1.Items.Clear()
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Sort()
        For Each aColl In MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns
            Dim aLvItem As New ColumnListViewItem(Me.ListView1, aColl)
            If SelectedIndex > -1 AndAlso aLvItem.Index = SelectedIndex Then
                aLvItem.Selected = True
            End If
        Next
        Me.ListView1.EndUpdate()
        Me.ListView1.Sorting = SortOrder.None
        Me.ListView1_Click(Nothing, Nothing)
    End Sub

    Private Sub cmdDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDown.Click
        If Me.CurrentItem Is Nothing Then Exit Sub

        If Me.CurrentItem.CollumnInfo.Index = MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Count - 1 Then Exit Sub 'Can't go up from the top 

        Dim intIndex As Integer = Me.CurrentItem.CollumnInfo.Index + 1
        SelectedIndex = intIndex
        'Dim prevItem As ColumnListViewItem = Me.ListView1.Items(intIndex)
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Item(intIndex, UserDefaults.ColumnInfoCollection.CICIndexOn.Order).Index -= 1
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Item(Me.CurrentItem.CollumnInfo.FieldName, UserDefaults.ColumnInfoCollection.CICIndexOn.FieldName).Index += 1

        Draw()
    End Sub

    Private Sub cmdDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDefaults.Click
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns = MedscreenCommonGui.CommonForms.ListViewColumns.LVDiaryList
    End Sub

    Private Sub btnChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChange.Click
        With Me.CurrentItem.CollumnInfo
            .HeaderText = Me.txtColumnHeader.Text
            .Show = Me.ckShow.Checked
            .ColumnWidth = Me.NUWidth.Value
            .FixedWidth = Me.ckFixed.Checked
            CurrentItem.Draw()
        End With
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class

