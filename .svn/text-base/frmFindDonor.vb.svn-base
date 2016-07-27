Imports MedscreenLib

Public Class frmFindDonor
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim i As Integer
        For I = 0 To 10
            ColumnDirection(I) = 1
        Next

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
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents txtSurname As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFirstname As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lvDonors As System.Windows.Forms.ListView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDOB As System.Windows.Forms.TextBox
    Friend WithEvents cmdFind As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents NWInfo As VbPowerPack.NotificationWindow
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFindDonor))
        Me.pnlControls = New System.Windows.Forms.Panel
        Me.cmdFind = New System.Windows.Forms.Button
        Me.txtDOB = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFirstname = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSurname = New System.Windows.Forms.TextBox
        Me.cmdOk = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lvDonors = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.NWInfo = New VbPowerPack.NotificationWindow(Me.components)
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.pnlControls.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlControls
        '
        Me.pnlControls.Controls.Add(Me.cmdFind)
        Me.pnlControls.Controls.Add(Me.txtDOB)
        Me.pnlControls.Controls.Add(Me.Label3)
        Me.pnlControls.Controls.Add(Me.Label2)
        Me.pnlControls.Controls.Add(Me.txtFirstname)
        Me.pnlControls.Controls.Add(Me.Label1)
        Me.pnlControls.Controls.Add(Me.txtSurname)
        Me.pnlControls.Controls.Add(Me.cmdOk)
        Me.pnlControls.Controls.Add(Me.cmdCancel)
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 597)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(872, 56)
        Me.pnlControls.TabIndex = 0
        '
        'cmdFind
        '
        Me.cmdFind.Image = CType(resources.GetObject("cmdFind.Image"), System.Drawing.Image)
        Me.cmdFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFind.Location = New System.Drawing.Point(392, 32)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.Size = New System.Drawing.Size(75, 23)
        Me.cmdFind.TabIndex = 16
        Me.cmdFind.Text = "Search"
        Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDOB
        '
        Me.txtDOB.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDOB.Location = New System.Drawing.Point(272, 32)
        Me.txtDOB.Name = "txtDOB"
        Me.txtDOB.Size = New System.Drawing.Size(100, 20)
        Me.txtDOB.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtDOB, "Date can be entered in any valid format, however having the month as text eg DEC " & _
                "is the most relaible.")
        Me.txtDOB.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(272, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Date of Birth"
        Me.Label3.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(155, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "First Name"
        '
        'txtFirstname
        '
        Me.txtFirstname.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFirstname.Location = New System.Drawing.Point(152, 32)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstname.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtFirstname, "incomplete names are allowed.")
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(34, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Surname"
        '
        'txtSurname
        '
        Me.txtSurname.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSurname.Location = New System.Drawing.Point(32, 32)
        Me.txtSurname.Name = "txtSurname"
        Me.txtSurname.Size = New System.Drawing.Size(100, 20)
        Me.txtSurname.TabIndex = 10
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(680, 32)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 23)
        Me.cmdOk.TabIndex = 9
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(760, 32)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 23)
        Me.cmdCancel.TabIndex = 8
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lvDonors)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(872, 597)
        Me.Panel1.TabIndex = 1
        '
        'lvDonors
        '
        Me.lvDonors.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvDonors.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.lvDonors.FullRowSelect = True
        Me.lvDonors.Location = New System.Drawing.Point(16, 24)
        Me.lvDonors.Name = "lvDonors"
        Me.lvDonors.Size = New System.Drawing.Size(816, 552)
        Me.lvDonors.TabIndex = 0
        Me.lvDonors.UseCompatibleStateImageBehavior = False
        Me.lvDonors.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Case Number"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Surname"
        Me.ColumnHeader2.Width = 150
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "First Name"
        Me.ColumnHeader3.Width = 150
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Gender"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "DOB"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Customer Ref"
        Me.ColumnHeader6.Width = 100
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Invoice"
        Me.ColumnHeader7.Width = 100
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Customer"
        Me.ColumnHeader8.Width = 90
        '
        'NWInfo
        '
        Me.NWInfo.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.SystemColors.InactiveCaption, System.Drawing.SystemColors.Window)
        Me.NWInfo.DefaultText = "This may take some time to show your request please be patient"
        Me.NWInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NWInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.NWInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NWInfo.ShowStyle = VbPowerPack.NotificationShowStyle.Immediately
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Job"
        '
        'frmFindDonor
        '
        Me.AcceptButton = Me.cmdFind
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(872, 653)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlControls)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFindDonor"
        Me.Text = "Find by Donor"
        Me.pnlControls.ResumeLayout(False)
        Me.pnlControls.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click

        Try
            Me.cmdFind.Enabled = False
            Dim intLen As Integer = Me.txtDOB.Text.Trim.Length + Me.txtFirstname.Text.Trim.Length + Me.txtSurname.Text.Trim.Length
            If intLen = 0 Then                  'No query will take for ever.
                MsgBox("You must enter something!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If


            Dim strDate As String = ""
            If Me.txtDOB.Text.Trim.Length > 0 Then
                Dim tmpDate As Date = Date.Parse(Me.txtDOB.Text)
                strDate = tmpDate.ToString("dd-MMM-yyyy")
            End If

           
            Try ' Protecting 
                NWInfo.DefaultText = "This may take some time please be patient"
                NWInfo.Notify("It may take some time to retrieve your data, please be patient")
                Dim oColl As New Collection
                If Me.txtSurname.Text.Trim.Length > 0 Then
                    oColl.Add("%" & txtSurname.Text.ToUpper.Trim & "%")
                Else
                    oColl.Add("%")
                End If
                If Me.txtFirstname.Text.Trim.Length > 0 Then
                    oColl.Add("%" & txtFirstname.Text.ToUpper.Trim & "%")
                Else
                    oColl.Add("%")
                End If
                System.Windows.Forms.Application.DoEvents()
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                Dim strReturn As String = CConnection.PackageStringList("Lib_CCTool.FindDonorRecordsF", oColl)
                If strReturn IsNot Nothing AndAlso strReturn.Trim.Length > 0 Then
                    Dim strBarcodes As String = strReturn
                    Me.lvDonors.Items.Clear()
                    If Not strBarcodes Is Nothing Then
                        Dim BarcodeArray As String() = strBarcodes.Split(New Char() {","})
                        Dim Barcode As String
                        For Each Barcode In BarcodeArray
                            Dim strDonorInfo As String = CConnection.PackageStringList("Lib_sample8.GetDonorInfo", Barcode)
                            If Not strDonorInfo Is Nothing Then
                                Dim lvItem As New MedscreenCommonGui.ListViewItems.lvDonor(strDonorInfo)
                                Me.lvDonors.Items.Add(lvItem)
                            End If
                        Next
                    End If
                Else
                    MsgBox("No donors found that match your criteria in the time from scanned", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "frmFindDonor-cmdFind_Click-181")
            Finally
            End Try
        Catch
        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.cmdFind.Enabled = True
        End Try

    End Sub

    Dim myCurrentDonor As ListViewItems.lvDonor
    Private Sub lvDonors_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDonors.Click
        myCurrentDonor = Nothing
        If Me.lvDonors.SelectedItems.Count = 1 Then
            myCurrentDonor = Me.lvDonors.SelectedItems.Item(0)
        End If
    End Sub

    Public ReadOnly Property Donor() As Intranet.intranet.customerns.DONOR
        Get
            If Me.myCurrentDonor Is Nothing Then
                Return Nothing
            Else
                Return Me.myCurrentDonor.Donor
            End If
        End Get
    End Property

    Public ReadOnly Property CaseNumber() As String
        Get
            If Me.myCurrentDonor Is Nothing Then
                Return ""
            Else
                Return Me.myCurrentDonor.Donor.CaseNumber
            End If

        End Get
    End Property

    Private Sub lvDonors_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDonors.DoubleClick
        If Not myCurrentDonor Is Nothing Then
            Me.cmdOk_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private ColumnDirection(10) As Integer

    Private Sub lvDonors_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvDonors.ColumnClick
        Dim intdirection As Integer

        intdirection = Me.ColumnDirection(e.Column)     'Find out the currently chossen direction
        Me.ColumnDirection(e.Column) *= -1              'Invert direction and store

        Me.lvDonors.ListViewItemSorter = New Comparers.ListViewItemComparer(e.Column, intdirection)

    End Sub
End Class
