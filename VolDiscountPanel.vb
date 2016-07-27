Imports System.Windows.Forms
Imports System.Drawing
Imports MedscreenLib
Imports System.Reflection
Imports System.Resources

Namespace Panels

    Public Class VolDiscountPanel
        Inherits Panel
#Region "Declarations"
        Private myVolDiscount As ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader
        Private myVolDiscountEntry As ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry
        Private myRoutineDiscount As ListViewItems.DiscountSchemes.lvCustDiscountScheme
        Friend WithEvents myMenu As ContextMenuStrip
        Friend WithEvents mnuEditVolDisc As ToolStripMenuItem



#Region "Events"
        Public Event VolumeDiscountSchemeAdded(ByVal SchemeId As Integer)
        Public Event VolumeSchemeDeleted(ByVal SchemeId As Integer)
        Public Event RoutineSchemeDeleted(ByVal SchemeId As Integer)
        Public Event VolumeSchemeChanged(ByVal SchemeId As Integer)
        Public Event RoutineSchemeChanged(ByVal SchemeId As Integer)
#End Region

#Region "Controls"
        Friend WithEvents dcvVolumeDiscounts As System.Windows.Forms.DataGridView
        Friend WithEvents dchSchemeEntry As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchType As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchAppliesTo As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchMinSample As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchMaxSample As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchMinCharge As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchSampleDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents dchExample As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents lblDiscScheme As System.Windows.Forms.Label
        Friend WithEvents lblVolDiscScheme As System.Windows.Forms.Label
        Friend WithEvents lvRoutineDiscount As System.Windows.Forms.ListView
        Friend WithEvents chNVDID As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDType As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDApplyto As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDDiscount As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDPrice As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDDisPrice As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDStart As System.Windows.Forms.ColumnHeader
        Friend WithEvents chNVDEnd As System.Windows.Forms.ColumnHeader
        Friend WithEvents cmdAddScheme As System.Windows.Forms.Button
        Friend WithEvents cmdDelDiscount As System.Windows.Forms.Button
        Friend WithEvents Label80 As System.Windows.Forms.Label
        Friend WithEvents cbVolScheme As System.Windows.Forms.ComboBox
        Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
        Friend WithEvents RoutinePanel As System.Windows.Forms.Panel
        Friend WithEvents VolumePanel As System.Windows.Forms.Panel
        Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
#End Region
#End Region

#Region "Properties"

        Private myCustomerID As String
        Private Client As Intranet.intranet.customerns.Client
        Public Property CustomerID() As String
            Get
                Return myCustomerID
            End Get
            Set(ByVal value As String)
                myCustomerID = value
                If value IsNot Nothing AndAlso value.Trim.Length > 0 Then
                    Client = New Intranet.intranet.customerns.Client(value)
                    DrawDiscounts()
                    DrawVolumeDiscounts()
                    SetupCombo()
                End If
            End Set
        End Property



#End Region

#Region "Methods"

        Private Sub lvRoutineDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRoutineDiscount.Click
            Me.myVolDiscount = Nothing
            Me.myRoutineDiscount = Nothing
            If Me.lvRoutineDiscount.SelectedItems.Count = 1 Then
                Dim objLvItem As ListViewItem
                objLvItem = Me.lvRoutineDiscount.SelectedItems(0)
                While TypeOf objLvItem Is ListViewItems.DiscountSchemes.lvDiscountScheme
                    objLvItem = Me.lvRoutineDiscount.Items(objLvItem.Index - 1)

                End While
                Me.myRoutineDiscount = objLvItem
            End If


        End Sub

        Private Sub lvRoutineDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRoutineDiscount.DoubleClick
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
                    RaiseEvent RoutineSchemeChanged(myRoutineDiscount.DiscountScheme.DiscountSchemeId)
                End If
            End With
            Me.DrawDiscounts()
        End Sub

        Private Sub lvVolumeDiscounts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dcvVolumeDiscounts.Click
            Me.myVolDiscountEntry = Nothing
            myVolDiscount = Nothing
            Me.myRoutineDiscount = Nothing
            If Me.dcvVolumeDiscounts.SelectedRows.Count = 1 Then
                Dim objLvItem As DataGridViewRow
                objLvItem = Me.dcvVolumeDiscounts.SelectedRows(0)
                If TypeOf objLvItem Is ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader Then
                    myVolDiscount = objLvItem
                ElseIf TypeOf objLvItem Is ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry Then
                    myVolDiscountEntry = objLvItem
                End If
                'Me.myVolDiscountEntry = objLvItem
            End If
        End Sub
        ''' <developer>CONCATENO\Taylor</developer>
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>19-Dec-2011 14:04</created><Author>CONCATENO\Taylor</Author></revision>
        ''' <revision><modified>19-Dec-2011 14:04 menu handler added</modified><Author>CONCATENO\Taylor</Author></revision>
        ''' </revisionHistory>
        Private Sub lvVolumeDiscounts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dcvVolumeDiscounts.DoubleClick, mnuEditVolDisc.Click
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
                    RaiseEvent VolumeSchemeChanged(myVolDiscount.Scheme.DiscountSchemeId)
                End If
            End With
            Me.DrawVolumeDiscounts()
        End Sub


        Public Sub DrawDiscounts()
            Me.lvRoutineDiscount.Items.Clear()
            Dim objCustDisnt As Intranet.intranet.customerns.CustomerDiscount
            If Client Is Nothing Then Exit Sub
            'Go through each discount scheme having refreshed them first
            Me.Client.InvoiceInfo.DiscountSchemes.Clear()
            Me.Client.InvoiceInfo.DiscountSchemes.Load()
            For Each objCustDisnt In Me.Client.InvoiceInfo.DiscountSchemes
                Dim objlvCitem As MedscreenCommonGui.ListViewItems.DiscountSchemes.lvCustDiscountScheme
                objlvCitem = New ListViewItems.DiscountSchemes.lvCustDiscountScheme(objCustDisnt)
                Me.lvRoutineDiscount.Items.Add(objlvCitem)
                Dim objEntry As MedscreenLib.Glossary.DiscountSchemeEntry

                'Go through each entry
                For Each objEntry In objCustDisnt.DiscountScheme.Entries
                    'Create line item 
                    Dim objLvItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.lvDiscountScheme(objEntry, Me.Client)
                    Me.lvRoutineDiscount.Items.Add(objLvItem)
                Next
            Next
        End Sub

        Public Sub DrawVolumeDiscounts()
            'Me.lvVolumeDiscounts.Items.Clear()
            Dim objCustDisnt As Intranet.intranet.customerns.CustomerVolumeDiscount
            dcvVolumeDiscounts.Rows.Clear()
            If Me.Client Is Nothing Then Exit Sub
            'Ensure data is refreshed before being drawn
            Me.Client.InvoiceInfo.VolumeDiscountSchemes.Clear()
            Me.Client.InvoiceInfo.VolumeDiscountSchemes.Load()
            'Go through all customer volume discounts
            For Each objCustDisnt In Me.Client.InvoiceInfo.VolumeDiscountSchemes
                Dim objlvCitem As MedscreenCommonGui.ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader
                objlvCitem = New ListViewItems.DiscountSchemes.dgvVolumeDiscountHeader(objCustDisnt)

                'Add header to list view
                Me.dcvVolumeDiscounts.Rows.Add(objlvCitem)
                Dim objEntry As MedscreenLib.Glossary.VolumeDiscountEntry

                'Go through each entry and add it
                For Each objEntry In objCustDisnt.DiscountScheme.Entries
                    'Dim objLvItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.lvVolumeDiscountEntry(objEntry, Me.Client)
                    'Me.lvVolumeDiscounts.Items.Add(objLvItem)
                    Dim objDCItem As New MedscreenCommonGui.ListViewItems.DiscountSchemes.dgvVolumeDiscountEntry(objEntry, Client)
                    Me.dcvVolumeDiscounts.Rows.Add(objDCItem)
                Next
            Next
        End Sub

        Private Sub cmdAddScheme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddScheme.Click
            Dim objSchemeId As Integer
            If Not TypeOf cbVolScheme.SelectedItem Is MedscreenLib.Glossary.VolumeDiscountHeader Then Exit Sub
            objSchemeId = CType(Me.cbVolScheme.SelectedItem, MedscreenLib.Glossary.VolumeDiscountHeader).DiscountSchemeId
            Dim objScheme As New Intranet.intranet.customerns.CustomerVolumeDiscount()
            objScheme.CustomerID = Me.Client.Identity
            objScheme.DiscountSchemeId = objSchemeId
            objScheme.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            objScheme.LinkField = "Customer.Identity"
            objScheme.Insert()
            Me.Client.InvoiceInfo.VolumeDiscountSchemes.Add(objScheme)
            Me.DrawVolumeDiscounts()
            RaiseEvent VolumeDiscountSchemeAdded(objSchemeId)

        End Sub

        Private Sub cmdDelDiscount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelDiscount.Click
            If Not myVolDiscount Is Nothing Then
                myVolDiscount.Scheme.EndDate = Today.AddDays(-1)
                myVolDiscount.Scheme.Update()
                Me.DrawVolumeDiscounts()
                RaiseEvent VolumeSchemeDeleted(myVolDiscount.Scheme.DiscountSchemeId)
            End If
            If Not Me.myRoutineDiscount Is Nothing Then
                myRoutineDiscount.DiscountScheme.EndDate = Today.AddDays(-1)
                myRoutineDiscount.DiscountScheme.Update()
                RaiseEvent RoutineSchemeDeleted(myRoutineDiscount.DiscountScheme.DiscountSchemeId)
                Me.DrawDiscounts()
            End If
        End Sub

#End Region

        Private Sub CreateControls()
            lblDiscScheme = New System.Windows.Forms.Label
            lblVolDiscScheme = New System.Windows.Forms.Label
            Me.dcvVolumeDiscounts = New System.Windows.Forms.DataGridView
            Me.dchSchemeEntry = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchType = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchAppliesTo = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchMinSample = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchMaxSample = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchMinCharge = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchSampleDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
            Me.dchExample = New System.Windows.Forms.DataGridViewTextBoxColumn
            lvRoutineDiscount = New System.Windows.Forms.ListView
            Me.chNVDID = New System.Windows.Forms.ColumnHeader
            Me.chNVDType = New System.Windows.Forms.ColumnHeader
            Me.chNVDApplyto = New System.Windows.Forms.ColumnHeader
            Me.chNVDDiscount = New System.Windows.Forms.ColumnHeader
            Me.chNVDPrice = New System.Windows.Forms.ColumnHeader
            Me.chNVDDisPrice = New System.Windows.Forms.ColumnHeader
            Me.chNVDStart = New System.Windows.Forms.ColumnHeader
            Me.chNVDEnd = New System.Windows.Forms.ColumnHeader
            Me.cmdDelDiscount = New System.Windows.Forms.Button
            Me.Label80 = New System.Windows.Forms.Label
            Me.cbVolScheme = New System.Windows.Forms.ComboBox
            Me.cmdAddScheme = New System.Windows.Forms.Button
            Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
            Me.WebBrowser1 = New System.Windows.Forms.WebBrowser
            myMenu = New System.Windows.Forms.ContextMenuStrip
            mnuEditVolDisc = New ToolStripMenuItem
            ' Me.RoutinePanel = New System.Windows.Forms.Panel
            'Me.VolumePanel = New System.Windows.Forms.Panel
        End Sub

        Private Sub InitialiseControls()
            'Me.RoutinePanel.SuspendLayout()
            'Me.VolumePanel.SuspendLayout()
            'Me.RoutinePanel.Location = New System.Drawing.Point(1, 1)
            'Me.RoutinePanel.Dock = DockStyle.Top
            'Me.RoutinePanel.Size = New System.Drawing.Size(760, 240)
            Me.SplitContainer1.Panel1.Controls.Add(lblDiscScheme)
            Me.SplitContainer1.Panel1.Controls.Add(lvRoutineDiscount)
            'Me.RoutinePanel.BorderStyle = Windows.Forms.BorderStyle.Fixed3D

            'Me.VolumePanel.Location = New System.Drawing.Point(1, 1)
            'Me.VolumePanel.Dock = DockStyle.Fill
            'Me.VolumePanel.Size = New System.Drawing.Size(760, 200)
            'Me.VolumePanel.Controls.Add(lblDiscScheme)
            'Me.VolumePanel.Controls.Add(lvRoutineDiscount)
            Me.SplitContainer1.Panel2.Controls.Add(lblVolDiscScheme)
            Me.SplitContainer1.Panel2.Controls.Add(dcvVolumeDiscounts)
            Me.SplitContainer1.Panel2.Controls.Add(Label80)
            Me.SplitContainer1.Panel2.Controls.Add(cbVolScheme)
            Me.SplitContainer1.Panel2.Controls.Add(cmdAddScheme)
            Me.SplitContainer1.Panel2.Controls.Add(cmdDelDiscount)
            Me.SplitContainer1.Panel2.Controls.Add(WebBrowser1)

            Me.SplitContainer1.Panel1.SuspendLayout()
            Me.SplitContainer1.Panel2.SuspendLayout()
            Me.SplitContainer1.SuspendLayout()
            CType(Me.dcvVolumeDiscounts, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'SplitContainer1
            '
            Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
            Me.SplitContainer1.Name = "SplitContainer1"
            Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
            '
            'SplitContainer1.Panel1
            '
            Me.SplitContainer1.Panel1.Controls.Add(Me.lblDiscScheme)
            Me.SplitContainer1.Panel1.Controls.Add(Me.lvRoutineDiscount)
            '
            'SplitContainer1.Panel2
            '
            Me.SplitContainer1.Panel2.Controls.Add(Me.lblVolDiscScheme)
            Me.SplitContainer1.Panel2.Controls.Add(Me.dcvVolumeDiscounts)
            Me.SplitContainer1.Panel2.Controls.Add(Me.Label80)
            Me.SplitContainer1.Panel2.Controls.Add(Me.cbVolScheme)
            Me.SplitContainer1.Panel2.Controls.Add(Me.cmdAddScheme)
            Me.SplitContainer1.Panel2.Controls.Add(Me.cmdDelDiscount)
            Me.SplitContainer1.Panel2.Controls.Add(Me.WebBrowser1)
            Me.SplitContainer1.Size = New System.Drawing.Size(770, 774)
            Me.SplitContainer1.SplitterDistance = 240
            Me.SplitContainer1.SplitterWidth = 5
            'Me.SplitContainer1.Panel2.BackColor = Color.LightGray
            Me.SplitContainer1.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
            Me.SplitContainer1.TabIndex = 0
            '
            'lblDiscScheme
            '
            Me.lblDiscScheme.AutoSize = True
            Me.lblDiscScheme.Location = New System.Drawing.Point(8, 8)
            Me.lblDiscScheme.Name = "lblDiscScheme"
            Me.lblDiscScheme.Size = New System.Drawing.Size(157, 13)
            Me.lblDiscScheme.TabIndex = 1
            Me.lblDiscScheme.Text = "Non Volume Discount Schemes"
            '
            'lvRoutineDiscount
            '
            Me.lvRoutineDiscount.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chNVDID, Me.chNVDType, Me.chNVDApplyto, Me.chNVDDiscount, Me.chNVDPrice, Me.chNVDDisPrice, Me.chNVDStart, Me.chNVDEnd})
            Me.lvRoutineDiscount.Location = New System.Drawing.Point(8, 32)
            Me.lvRoutineDiscount.Name = "lvRoutineDiscount"
            Me.lvRoutineDiscount.Size = New System.Drawing.Size(760, 200)
            Me.lvRoutineDiscount.TabIndex = 0
            Me.lvRoutineDiscount.UseCompatibleStateImageBehavior = False
            Me.lvRoutineDiscount.View = System.Windows.Forms.View.Details
            Me.lvRoutineDiscount.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top Or Bottom
            '
            'chNVDID
            '
            Me.chNVDID.Text = "Scheme-Entry"
            Me.chNVDID.Width = 150
            '
            'chNVDType
            '
            Me.chNVDType.Text = "Type"
            Me.chNVDType.Width = 200
            '
            'chNVDApplyto
            '
            Me.chNVDApplyto.Text = "Applys To (Line item)"
            Me.chNVDApplyto.Width = 300
            '
            'chNVDDiscount
            '
            Me.chNVDDiscount.Text = "Discount"
            Me.chNVDDiscount.Width = 120
            '
            'chNVDPrice
            '
            Me.chNVDPrice.Text = "Price"
            '
            'chNVDDisPrice
            '
            Me.chNVDDisPrice.Text = "Discount Price"
            '
            'chNVDStart
            '
            Me.chNVDStart.Text = "Start Date"
            Me.chNVDStart.Width = 120
            '
            'chNVDEnd
            '
            Me.chNVDEnd.Text = "End Date"
            Me.chNVDEnd.Width = 120
            '
            'lblVolDiscScheme
            '
            Me.lblVolDiscScheme.AutoSize = True
            Me.lblVolDiscScheme.Location = New System.Drawing.Point(8, 5)
            Me.lblVolDiscScheme.Name = "lblVolDiscScheme"
            Me.lblVolDiscScheme.Size = New System.Drawing.Size(134, 13)
            Me.lblVolDiscScheme.TabIndex = 1
            Me.lblVolDiscScheme.Text = "Volume Discount Schemes"
            '
            'dcvVolumeDiscounts
            '
            Me.dcvVolumeDiscounts.AllowUserToAddRows = False
            Me.dcvVolumeDiscounts.AllowUserToDeleteRows = False
            Me.dcvVolumeDiscounts.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            Me.dcvVolumeDiscounts.CausesValidation = False
            Me.dcvVolumeDiscounts.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.RaisedHorizontal
            Me.dcvVolumeDiscounts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dcvVolumeDiscounts.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dchSchemeEntry, Me.dchType, Me.dchAppliesTo, Me.dchMinSample, Me.dchMaxSample, Me.dchMinCharge, Me.dchSampleDiscount, Me.dchExample})
            Me.dcvVolumeDiscounts.GridColor = System.Drawing.SystemColors.Control
            Me.dcvVolumeDiscounts.Location = New System.Drawing.Point(1, 20)
            Me.dcvVolumeDiscounts.Name = "dcvVolumeDiscounts"
            Me.dcvVolumeDiscounts.Size = New System.Drawing.Size(760, 313)
            Me.dcvVolumeDiscounts.TabIndex = 1
            Me.dcvVolumeDiscounts.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            Me.dcvVolumeDiscounts.ContextMenuStrip = myMenu
            '
            'dchSchemeEntry
            '
            Me.dchSchemeEntry.HeaderText = "Scheme Entry"
            Me.dchSchemeEntry.Name = "dchSchemeEntry"
            Me.dchSchemeEntry.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchSchemeEntry.Width = 71
            '
            'dchType
            '
            Me.dchType.HeaderText = "Type"
            Me.dchType.Name = "dchType"
            Me.dchType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchType.Width = 37
            '
            'dchAppliesTo
            '
            Me.dchAppliesTo.HeaderText = "Change Comment / Applies To"
            Me.dchAppliesTo.Name = "dchAppliesTo"
            Me.dchAppliesTo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchAppliesTo.Width = 97
            '
            'dchMinSample
            '
            Me.dchMinSample.HeaderText = "Start Date / Min Sample "
            Me.dchMinSample.Name = "dchMinSample"
            Me.dchMinSample.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchMinSample.Width = 83
            '
            'dchMaxSample
            '
            Me.dchMaxSample.HeaderText = "End Date / Max Sample"
            Me.dchMaxSample.Name = "dchMaxSample"
            Me.dchMaxSample.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchMaxSample.Width = 83
            '
            'dchMinCharge
            '
            Me.dchMinCharge.HeaderText = "Min Charge"
            Me.dchMinCharge.Name = "dchMinCharge"
            Me.dchMinCharge.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchMinCharge.Width = 60
            '
            'dchSampleDiscount
            '
            Me.dchSampleDiscount.HeaderText = "Sample Discount"
            Me.dchSampleDiscount.Name = "dchSampleDiscount"
            Me.dchSampleDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchSampleDiscount.Width = 84
            '
            'dchExample
            '
            Me.dchExample.HeaderText = "Example"
            Me.dchExample.Name = "dchExample"
            Me.dchExample.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Me.dchExample.Width = 53
            '
            'Label80
            '
            Me.Label80.Location = New System.Drawing.Point(182, 403)
            Me.Label80.Name = "Label80"
            Me.Label80.Size = New System.Drawing.Size(424, 23)
            Me.Label80.TabIndex = 2
            Me.Label80.Text = "To add a volume discount scheme select scheme in combobox then hit add scheme"
            Me.Label80.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom Or AnchorStyles.Right
            '
            'cbVolScheme
            '
            Me.cbVolScheme.Location = New System.Drawing.Point(136, 365)
            Me.cbVolScheme.Name = "cbVolScheme"
            Me.cbVolScheme.Size = New System.Drawing.Size(576, 21)
            Me.cbVolScheme.Sorted = True
            Me.cbVolScheme.TabIndex = 1
            Me.cbVolScheme.Text = "ComboBox1"
            Me.cbVolScheme.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            '
            'cmdAddScheme
            '
            Me.cmdAddScheme.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.cmdAddScheme.Location = New System.Drawing.Point(24, 363)
            Me.cmdAddScheme.Name = "cmdAddScheme"
            Me.cmdAddScheme.Size = New System.Drawing.Size(96, 24)
            Me.cmdAddScheme.TabIndex = 0
            Me.cmdAddScheme.Text = "Add Scheme"
            Me.cmdAddScheme.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.cmdAddScheme.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom
            '
            'cmdDelDiscount
            '
            Me.cmdDelDiscount.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.cmdDelDiscount.Location = New System.Drawing.Point(24, 413)
            Me.cmdDelDiscount.Name = "cmdDelDiscount"
            Me.cmdDelDiscount.Size = New System.Drawing.Size(96, 24)
            Me.cmdDelDiscount.TabIndex = 3
            Me.cmdDelDiscount.Text = "Del Scheme"
            Me.cmdDelDiscount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.cmdDelDiscount.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

            Me.WebBrowser1.Location = New System.Drawing.Point(136, 423)
            Me.WebBrowser1.Name = "WebBrowser1"
            Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
            Me.WebBrowser1.Size = New System.Drawing.Size(576, 100)
            Me.WebBrowser1.TabIndex = 0
            Me.WebBrowser1.DocumentText = "<HTML></HTML>"
            Me.WebBrowser1.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            '
            'Frmtest
            '
            Me.SplitContainer1.Panel1.ResumeLayout(False)
            Me.SplitContainer1.Panel1.PerformLayout()
            Me.SplitContainer1.Panel2.ResumeLayout(False)
            Me.SplitContainer1.Panel2.PerformLayout()
            Me.SplitContainer1.ResumeLayout(False)
            CType(Me.dcvVolumeDiscounts, System.ComponentModel.ISupportInitialize).EndInit()

            Me.mnuEditVolDisc.Text = "Set Volume Discount Details"


            myMenu.Items.Add(mnuEditVolDisc)
            Me.ResumeLayout(False)
        End Sub
        Private Built As Boolean = False
        Private Sub SetupCombo()
            Dim myDataSource As New MedscreenLib.Glossary.CustomerVolumeDiscountHeaders(CustomerID)
            Me.cbVolScheme.DataSource = myDataSource
            If Me.cbVolScheme.BindingContext IsNot Nothing Then
                Me.cbVolScheme.BindingContext(Me.cbVolScheme.DataSource).SuspendBinding()
                Me.cbVolScheme.BindingContext(Me.cbVolScheme.DataSource).ResumeBinding()
            End If
            Me.cbVolScheme.DisplayMember = "FullDescription"
            Me.cbVolScheme.ValueMember = "DiscountSchemeId"
            Built = True
            Me.cbVolScheme.SelectedItem = MedscreenLib.Glossary.Glossary.VolumeDiscountSchemes(0).DiscountSchemeId
            cbVolScheme_SelectedIndexChanged(Nothing, Nothing)
        End Sub
        Public Sub New()
            MyBase.New()
            Me.Dock = System.Windows.Forms.DockStyle.Fill
            Me.Size = New System.Drawing.Size(763, 500)

            CreateControls()
            InitialiseControls()
            'Me.Controls.Add(RoutinePanel)
            Me.Controls.Add(Me.SplitContainer1)

            'Me.Controls.Add(VolumePanel)
            '<Added code modified 23-Apr-2007 06:33 by taylor> 
        End Sub


        Private Sub InitializeComponent()
            Me.SuspendLayout()
            Me.ResumeLayout(False)

        End Sub

        Private Sub cbVolScheme_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbVolScheme.SelectedIndexChanged
            Dim lResources As ResourceManager = New ResourceManager("MedscreenCommonGui.CommnonForms", [Assembly].GetExecutingAssembly())
            If Built Then
                Try
                    If Me.cbVolScheme.SelectedItem IsNot Nothing Then
                        Dim myObj As Object = Me.cbVolScheme.SelectedItem
                        Dim oColl As New Collection
                        oColl.Add(Me.myCustomerID)
                        oColl.Add(myObj.discountschemeid)
                        Dim strXML As String = CConnection.PackageStringList("Lib_customer.XMLExample", oColl)
                        If strXML IsNot Nothing Then
                            Dim strhtml As String = Medscreen.ResolveStyleSheet(strXML, "VolDiscExample.xslt", 0)
                            Me.WebBrowser1.DocumentText = strhtml
                        End If
                    End If
                Catch ex As Exception
                    Me.WebBrowser1.DocumentText = "<HTML><B>Details can't be displayed probably due to complexity</B></HTML>"
                End Try
            End If
        End Sub
    End Class
End Namespace
