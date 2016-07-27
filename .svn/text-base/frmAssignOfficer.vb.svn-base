'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-27 10:58:53+00 $
'$Log: frmAssignOfficer.vb,v $
'Revision 1.0  2005-11-27 10:58:53+00  taylor
'Added comments
'
'Revision 1.0  2005-11-22 06:30:28+00  taylor
Imports Intranet
Imports Intranet.intranet.Support
Imports MedscreenLib
Imports MedscreenCommonGui.ListViewItems

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmAssignOfficer
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Assign officer to a port
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-11-27 10:58:53+00 $</date><Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmAssignOfficer
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Me.CBOfficers.DataSource = cSupport.CollectingOfficers
        'Me.CBOfficers.DisplayMember = "Info2"
        'Me.CBOfficers.ValueMember = "OfficerID"

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
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCosts As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ckAgreed As System.Windows.Forms.CheckBox
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOfficer As System.Windows.Forms.TextBox
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Country As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdFind As System.Windows.Forms.Button
    Friend WithEvents chDivision As System.Windows.Forms.ColumnHeader
    Friend WithEvents chSubDivision As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAssignOfficer))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCosts = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ckAgreed = New System.Windows.Forms.CheckBox()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtOfficer = New System.Windows.Forms.TextBox()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.Country = New System.Windows.Forms.ColumnHeader()
        Me.cmdFind = New System.Windows.Forms.Button()
        Me.chDivision = New System.Windows.Forms.ColumnHeader()
        Me.chSubDivision = New System.Windows.Forms.ColumnHeader()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 349)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(720, 56)
        Me.Panel1.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(632, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(552, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(25, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Port"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPort
        '
        Me.txtPort.Enabled = False
        Me.txtPort.Location = New System.Drawing.Point(80, 16)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.TabIndex = 7
        Me.txtPort.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 184)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Officer"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCosts
        '
        Me.txtCosts.Location = New System.Drawing.Point(80, 216)
        Me.txtCosts.Name = "txtCosts"
        Me.txtCosts.TabIndex = 11
        Me.txtCosts.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(32, 216)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Costs"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckAgreed
        '
        Me.ckAgreed.Location = New System.Drawing.Point(200, 216)
        Me.ckAgreed.Name = "ckAgreed"
        Me.ckAgreed.Size = New System.Drawing.Size(176, 24)
        Me.ckAgreed.TabIndex = 12
        Me.ckAgreed.Text = "Contractually Agreed Costs"
        '
        'txtComments
        '
        Me.txtComments.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtComments.Location = New System.Drawing.Point(32, 264)
        Me.txtComments.MaxLength = 200
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(648, 80)
        Me.txtComments.TabIndex = 14
        Me.txtComments.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(32, 248)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Comments"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOfficer
        '
        Me.txtOfficer.Location = New System.Drawing.Point(80, 184)
        Me.txtOfficer.Name = "txtOfficer"
        Me.txtOfficer.Size = New System.Drawing.Size(376, 20)
        Me.txtOfficer.TabIndex = 15
        Me.txtOfficer.Text = ""
        '
        'ListView1
        '
        Me.ListView1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.Country, Me.chDivision, Me.chSubDivision})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(32, 56)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(648, 110)
        Me.ListView1.TabIndex = 16
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Surname"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Forename"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Region"
        Me.ColumnHeader3.Width = 100
        '
        'Country
        '
        Me.Country.Text = "Country"
        Me.Country.Width = 100
        '
        'cmdFind
        '
        Me.cmdFind.Image = CType(resources.GetObject("cmdFind.Image"), System.Drawing.Bitmap)
        Me.cmdFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFind.Location = New System.Drawing.Point(496, 184)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.TabIndex = 24
        Me.cmdFind.Text = "Find"
        Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chDivision
        '
        Me.chDivision.Text = "Division"
        Me.chDivision.Width = 100
        '
        'chSubDivision
        '
        Me.chSubDivision.Text = "Sub Division"
        Me.chSubDivision.Width = 80
        '
        'frmAssignOfficer
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(720, 405)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdFind, Me.ListView1, Me.txtOfficer, Me.txtComments, Me.Label4, Me.ckAgreed, Me.txtCosts, Me.Label3, Me.Label2, Me.txtPort, Me.Label1, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAssignOfficer"
        Me.Text = "Assign Officer to Port"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myCollPort As Intranet.intranet.jobs.CollPort
    Private myCollOfficer As Intranet.intranet.jobs.CollectingOfficer

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the collection port entity officer is tied to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CollPort() As Intranet.intranet.jobs.CollPort
        Get
            Return myCollPort
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CollPort)
            myCollPort = Value
            If Value Is Nothing Then Exit Property

            With myCollPort
                Me.txtComments.Text = .Comments
                Me.txtPort.Text = .PortId
                Me.txtCosts.Text = .Costs
                Me.ckAgreed.Checked = .Agreed

                If .OfficerId.Trim.Length > 0 Then          'If we have a valid officer
                    Me.myCollOfficer = cSupport.CollectingOfficers.Item(.OfficerId)                                   'Get officer info
                    If Not myCollOfficer Is Nothing Then
                        Me.txtOfficer.Text = myCollOfficer.Info2
                    End If
                End If
            End With
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accept info
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        With myCollPort
            .Comments = Me.txtComments.Text         'Get all info for this officer
            .PortId = Me.txtPort.Text
            .Costs = Me.txtCosts.Text
            .Agreed = Me.ckAgreed.Checked
            If Not Me.myCollOfficer Is Nothing Then
                .OfficerId = myCollOfficer.OfficerID
            End If
        End With
    End Sub

    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        If Me.ListView1.SelectedItems.Count = 1 Then
            Dim lvItem As LVCofficerItem
            lvItem = Me.ListView1.SelectedItems.Item(0)
            Me.myCollOfficer = lvItem.Officer
        End If
    End Sub


    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
        Me.ListView1.Items.Clear()
        If Me.txtOfficer.Text.Trim.Length > 0 Then
            Dim strOfficers As String = CConnection.PackageStringList("LIB_CCTOOL.FINDOFFICER", Me.txtOfficer.Text)
            If strOfficers.Trim.Length > 0 Then
                Dim strOffArray As String() = strOfficers.Split(New Char() {","})
                Dim i As Integer
                For i = 0 To strOffArray.Length - 1
                    Dim strOffID As String = strOffArray.GetValue(i)
                    If strOffID.Trim.Length > 0 Then 'Valid officer id 
                        Dim objCollOff As New Intranet.intranet.jobs.CollectingOfficer(strOffID)
                        objCollOff.Load()

                        Dim lvItem As New LVCofficerItem(objCollOff, "", Today)
                        lvItem.DisplayOption = ListViewItems.LVCofficerItem.DisplayOptions.search
                        Me.ListView1.Items.Add(lvItem)
                    End If
                Next

            End If
        End If
    End Sub
End Class
