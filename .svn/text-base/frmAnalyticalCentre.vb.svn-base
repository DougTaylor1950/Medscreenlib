Imports Intranet
Imports MedscreenLib
Public Class frmAnalyticalCentre

#Region "Declarations"
    Private myInstrument As Intranet.INSTRUMENT.Instrument

#End Region

#Region "Properties"
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
#End Region

#Region "Instance Code"
    Private Sub FillForm()
        If myInstrument Is Nothing Then Exit Sub
        With myInstrument
            Me.txtDescription.Text = .Description
            Me.cbStatus.SelectedValue = .Status
            Me.AddressPanel1.Address = .Address
        End With
        Me.DrawComments()
    End Sub
#End Region

#Region "Comments"
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
#End Region

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim objStatus As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INST_STAT")
        objStatus.Load("phrase_text")
        Me.cbStatus.DataSource = objStatus
        Me.cbStatus.DisplayMember = "PhraseText"
        Me.cbStatus.ValueMember = "PhraseID"

    End Sub

    Private Sub AddressPanel1_FindAddress(ByVal AddressID As Integer) Handles AddressPanel1.FindAddress, AddressPanel1.NewAddress
        If myInstrument IsNot Nothing Then
            myInstrument.AddressID = AddressID
        End If
    End Sub

    Private Function GetForm() As Boolean
        Dim blnRet As Boolean = True
        With myInstrument
            .Description = Me.txtDescription.Text
            .Status = Me.cbStatus.SelectedValue
            'Checks 
        End With
        Return blnRet

    End Function

    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Not GetForm() Then
            Me.DialogResult = Windows.Forms.DialogResult.None
        End If
    End Sub
End Class