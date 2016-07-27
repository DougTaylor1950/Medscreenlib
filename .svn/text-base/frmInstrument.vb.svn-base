
Imports Intranet
Imports MedscreenLib
Public Class frmInstrument
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Dim objModel As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_MODEL")
        objModel.Load("phrase_text")
        Me.cbModel.DataSource = objModel
        Me.cbModel.DisplayMember = "PhraseText"
        Me.cbModel.ValueMember = "PhraseID"
        Dim objStatus As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_STAT")
        objStatus.Load("phrase_text")
        Me.cbStatus.DataSource = objStatus
        Me.cbStatus.DisplayMember = "PhraseText"
        Me.cbStatus.ValueMember = "PhraseID"
        Dim objInstLoc As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_LOC")
        objInstLoc.Load("phrase_text")
        Me.cbLocation.DataSource = objInstLoc
        Me.cbLocation.DisplayMember = "PhraseText"
        Me.cbLocation.ValueMember = "PhraseID"
        Dim objInstType As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_TYPE")
        objInstType.Load("phrase_text")
        Me.cbType.DataSource = objInstType
        Me.cbType.DisplayMember = "PhraseText"
        Me.cbType.ValueMember = "PhraseID"

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
    Friend WithEvents tabDetails As System.Windows.Forms.TabPage
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSerialNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbModel As System.Windows.Forms.ComboBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtLastService As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbLocation As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtLastCalib As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbType As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblAssignedTo As System.Windows.Forms.Label
    Friend WithEvents tabComments As System.Windows.Forms.TabPage
    Friend WithEvents lvComments As System.Windows.Forms.ListView
    Friend WithEvents chCommDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommType As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHComment As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCommBy As System.Windows.Forms.ColumnHeader
    Friend WithEvents ChCommAction As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdAddComment As System.Windows.Forms.Button
    Friend WithEvents cmdAssign As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmInstrument))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabDetails = New System.Windows.Forms.TabPage()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.dtLastCalib = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtLastService = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cbType = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cbStatus = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbLocation = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cbModel = New System.Windows.Forms.ComboBox()
        Me.txtSerialNo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblAssignedTo = New System.Windows.Forms.Label()
        Me.tabComments = New System.Windows.Forms.TabPage()
        Me.lvComments = New System.Windows.Forms.ListView()
        Me.chCommDate = New System.Windows.Forms.ColumnHeader()
        Me.chCommType = New System.Windows.Forms.ColumnHeader()
        Me.CHComment = New System.Windows.Forms.ColumnHeader()
        Me.chCommBy = New System.Windows.Forms.ColumnHeader()
        Me.ChCommAction = New System.Windows.Forms.ColumnHeader()
        Me.cmdAddComment = New System.Windows.Forms.Button()
        Me.cmdAssign = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.tabDetails.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.tabComments.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 557)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 32)
        Me.Panel1.TabIndex = 2
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(700, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(620, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabDetails, Me.tabComments})
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(784, 557)
        Me.TabControl1.TabIndex = 3
        '
        'tabDetails
        '
        Me.tabDetails.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel3, Me.Panel2})
        Me.tabDetails.Location = New System.Drawing.Point(4, 22)
        Me.tabDetails.Name = "tabDetails"
        Me.tabDetails.Size = New System.Drawing.Size(776, 531)
        Me.tabDetails.TabIndex = 0
        Me.tabDetails.Tag = "1"
        Me.tabDetails.Text = "Details"
        '
        'Panel3
        '
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAssign, Me.dtLastCalib, Me.Label6, Me.dtLastService, Me.Label4})
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 112)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(776, 100)
        Me.Panel3.TabIndex = 1
        '
        'dtLastCalib
        '
        Me.dtLastCalib.Location = New System.Drawing.Point(456, 16)
        Me.dtLastCalib.Name = "dtLastCalib"
        Me.dtLastCalib.Size = New System.Drawing.Size(136, 20)
        Me.dtLastCalib.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(328, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Last Calibration :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtLastService
        '
        Me.dtLastService.Location = New System.Drawing.Point(152, 16)
        Me.dtLastService.Name = "dtLastService"
        Me.dtLastService.Size = New System.Drawing.Size(136, 20)
        Me.dtLastService.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Last Service :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblAssignedTo, Me.cbType, Me.Label8, Me.cbStatus, Me.Label7, Me.cbLocation, Me.Label5, Me.cbModel, Me.txtSerialNo, Me.Label3, Me.Label2, Me.txtDescription, Me.Label1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(776, 112)
        Me.Panel2.TabIndex = 0
        '
        'cbType
        '
        Me.cbType.Location = New System.Drawing.Point(144, 48)
        Me.cbType.Name = "cbType"
        Me.cbType.Size = New System.Drawing.Size(121, 21)
        Me.cbType.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Type :"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbStatus
        '
        Me.cbStatus.Location = New System.Drawing.Point(144, 80)
        Me.cbStatus.Name = "cbStatus"
        Me.cbStatus.Size = New System.Drawing.Size(121, 21)
        Me.cbStatus.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Status :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbLocation
        '
        Me.cbLocation.Location = New System.Drawing.Point(608, 48)
        Me.cbLocation.Name = "cbLocation"
        Me.cbLocation.Size = New System.Drawing.Size(121, 21)
        Me.cbLocation.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(488, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Location :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbModel
        '
        Me.cbModel.Location = New System.Drawing.Point(376, 80)
        Me.cbModel.Name = "cbModel"
        Me.cbModel.Size = New System.Drawing.Size(121, 21)
        Me.cbModel.TabIndex = 6
        '
        'txtSerialNo
        '
        Me.txtSerialNo.Location = New System.Drawing.Point(376, 48)
        Me.txtSerialNo.Name = "txtSerialNo"
        Me.txtSerialNo.TabIndex = 5
        Me.txtSerialNo.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(272, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Serial No :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(312, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Model :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(144, 16)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(576, 20)
        Me.txtDescription.TabIndex = 1
        Me.txtDescription.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Description :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAssignedTo
        '
        Me.lblAssignedTo.Location = New System.Drawing.Point(544, 88)
        Me.lblAssignedTo.Name = "lblAssignedTo"
        Me.lblAssignedTo.Size = New System.Drawing.Size(208, 23)
        Me.lblAssignedTo.TabIndex = 13
        Me.lblAssignedTo.Text = "."
        '
        'tabComments
        '
        Me.tabComments.Controls.AddRange(New System.Windows.Forms.Control() {Me.lvComments, Me.cmdAddComment})
        Me.tabComments.Location = New System.Drawing.Point(4, 22)
        Me.tabComments.Name = "tabComments"
        Me.tabComments.Size = New System.Drawing.Size(776, 531)
        Me.tabComments.TabIndex = 1
        Me.tabComments.Text = "Comments"
        '
        'lvComments
        '
        Me.lvComments.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lvComments.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chCommDate, Me.chCommType, Me.CHComment, Me.chCommBy, Me.ChCommAction})
        Me.lvComments.Location = New System.Drawing.Point(8, 8)
        Me.lvComments.Name = "lvComments"
        Me.lvComments.Size = New System.Drawing.Size(752, 193)
        Me.lvComments.TabIndex = 2
        Me.lvComments.View = System.Windows.Forms.View.Details
        '
        'chCommDate
        '
        Me.chCommDate.Text = "Date"
        Me.chCommDate.Width = 100
        '
        'chCommType
        '
        Me.chCommType.Text = "Type"
        Me.chCommType.Width = 80
        '
        'CHComment
        '
        Me.CHComment.Text = "Comment"
        Me.CHComment.Width = 200
        '
        'chCommBy
        '
        Me.chCommBy.Text = "Created By"
        Me.chCommBy.Width = 100
        '
        'ChCommAction
        '
        Me.ChCommAction.Text = "Action By"
        '
        'cmdAddComment
        '
        Me.cmdAddComment.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.cmdAddComment.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddComment.Location = New System.Drawing.Point(8, 216)
        Me.cmdAddComment.Name = "cmdAddComment"
        Me.cmdAddComment.Size = New System.Drawing.Size(80, 23)
        Me.cmdAddComment.TabIndex = 3
        Me.cmdAddComment.Text = "Comment"
        Me.cmdAddComment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAssign
        '
        Me.cmdAssign.Image = CType(resources.GetObject("cmdAssign.Image"), System.Drawing.Bitmap)
        Me.cmdAssign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAssign.Location = New System.Drawing.Point(56, 56)
        Me.cmdAssign.Name = "cmdAssign"
        Me.cmdAssign.TabIndex = 7
        Me.cmdAssign.Text = "Assign"
        Me.cmdAssign.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmInstrument
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 589)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInstrument"
        Me.Text = "Edit Instrument"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.tabDetails.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.tabComments.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myInstrument As Intranet.INSTRUMENT.Instrument

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose instrument property 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	14/07/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Instrument() As Intranet.INSTRUMENT.Instrument
        Get
            Return myInstrument
        End Get
        Set(ByVal Value As Intranet.INSTRUMENT.Instrument)
            myInstrument = Value
            objCommentList = New MedscreenLib.Comment.Collection(Comment.CommentOn.instrument, Me.myInstrument.Identity)

            FillForm()
        End Set
    End Property

    Private Sub FillForm()
        If myInstrument Is Nothing Then Exit Sub
        With myInstrument
            Me.txtDescription.Text = .Description
            Me.cbModel.SelectedValue = .Model
            Me.txtSerialNo.Text = .SerialNo
            Me.dtLastService.Value = .ServiceDate
            Me.cbLocation.SelectedValue = .LocationId
            Me.dtLastCalib.Value = .CalibDate
            Me.cbStatus.SelectedValue = .Status
            Me.cbType.SelectedValue = .InsttypeId
            Me.lblAssignedTo.Text = .AssignedTo
        End With
        Me.DrawComments()
    End Sub

    Private Function GetForm() As Boolean
        Dim blnRet As Boolean = True
        With myInstrument
            .Description = Me.txtDescription.Text
            .Model = Me.cbModel.SelectedValue
            .SerialNo = Me.txtSerialNo.Text
            .ServiceDate = Me.dtLastService.Value
            .LocationId = Me.cbLocation.SelectedValue
            .CalibDate = Me.dtLastCalib.Value
            .Status = Me.cbStatus.SelectedValue
            .InsttypeId = Me.cbType.SelectedValue
            'Checks 
            If .LocationId <> Intranet.INSTRUMENT.Instrument.GCST_LocationField Then
                If .AssignedTo.Trim.Length > 0 Then 'Assigned to an officer
                    Dim dlgRet As MsgBoxResult = MsgBox("This instrument is assigned to " & .AssignedTo & _
                      " setting the location away from with officer will mean this assignment will be removed. Do you want to do this?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
                    Dim intRet As Integer
                    If dlgRet = MsgBoxResult.Yes Then

                        Dim oCmd As New OleDb.OleDbCommand()
                        Try ' Protecting 
                            oCmd.CommandType = CommandType.StoredProcedure
                            oCmd.CommandText = "lib_instrument.RemoveOfficerInstrument"
                            oCmd.Parameters.Add(CConnection.StringParameter("InstrumentID", .Identity, 10))
                            oCmd.Parameters.Add(CConnection.StringParameter("OfficerId", .Assignee, 10))
                            oCmd.Connection = CConnection.DbConnection
                            If CConnection.ConnOpen Then            'Attempt to open reader
                                intRet = oCmd.ExecuteNonQuery
                            End If

                        Catch ex As Exception
                            MedscreenLib.Medscreen.LogError(ex, , "frmInstrument-GetForm-387")
                        Finally
                            CConnection.SetConnClosed()             'Close connection
                        End Try
                    Else
                        blnRet = False
                    End If
                End If
            End If
        End With
        Return blnRet

    End Function

    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Not GetForm() Then
            Me.DialogResult = Windows.Forms.DialogResult.None
        End If
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub frmInstrument_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private objCommentList As MedscreenLib.Comment.Collection

    Private Sub DrawComments()
        Dim objComm As MedscreenLib.Comment
        Me.lvComments.Items.Clear()
        objCommentList.Clear()
        'objCommentList.Load()

        objCommentList.Load(Comment.CommentOn.instrument, Me.myInstrument.Identity)
        For Each objComm In objCommentList
            Me.lvComments.Items.Add(New MedscreenCommonGui.ListViewItems.lvComment(objComm))
        Next
    End Sub

    Private Sub cmdAddComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddComment.Click
        Dim afrm As New FrmCollComment()
        afrm.FormType = FrmCollComment.FCCFormTypes.Instrument
        afrm.CommentType = MedscreenLib.Constants.GCST_COMM_NEWInstrument
        afrm.CMNumber = Me.myInstrument.Identity
        If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Dim aComm As Intranet.intranet.jobs.CollectionComment = _
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.instrument, Me.myInstrument.Identity, afrm.CommentType, afrm.Comment, _
            MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            'If Not aComm Is Nothing Then
            '    aComm.ActionBy = afrm.ActionBy
            '    aComm.Update()
            Me.DrawComments()
            'End If
        End If

    End Sub

    Private Sub cmdAssign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAssign.Click
        Dim objCS As New frmSelectCollector()

        If objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim objColl As Intranet.intranet.jobs.CollectingOfficer = objCS.Officer
            If Me.myInstrument.AssignedTo.Trim.Length > 0 Then 'Already assigned

            End If
            Dim objCinst As Intranet.INSTRUMENT.OfficerInstrument = New Intranet.INSTRUMENT.OfficerInstrument()
            objCinst.InstrumentId = myInstrument.Identity
            objCinst.OfficerId = objColl.OfficerID
            objCinst.CreatedOn = Now
            objCinst.ModifiedOn = Now
            objCinst.CreatedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            objCinst.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            objCinst.Removed = False
            objCinst.DoUpdate()
            myInstrument.LocationId = Intranet.INSTRUMENT.Instrument.GCST_LocationField
            myInstrument.DoUpdate()
            objColl.Instruments.Add(objCinst)
            Me.FillForm()
        End If

    End Sub
End Class
