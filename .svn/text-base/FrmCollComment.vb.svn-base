'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-28 12:35:51+00 $
'$Log: FrmCollComment.vb,v $
'Revision 1.0  2005-11-28 12:35:51+00  taylor
'Commented
'
Imports MedscreenLib.Medscreen
Imports System.Windows.Forms
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : FrmCollComment
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A form for getting a colelction comment 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class FrmCollComment
    Inherits System.Windows.Forms.Form

    Private strCommentType As String
    Private DeptListSales As MedscreenLib.personnel.PersonnelList
    Private DeptList As MedscreenLib.personnel.PersonnelList
    Private DeptList2 As MedscreenLib.personnel.PersonnelList
    Private strPM As String = ""


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible modes for form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum FCCFormTypes
        ''' <summary>Define a comment</summary>
        def
        ''' <summary>Collectors Information </summary>
        CollectorInfo
        ''' <summary>Customer Comment </summary>
        customerComment
        ''' <summary>Customer Issue </summary>
        CustomerIssue
        CollectorComment
        Instrument
    End Enum

    Private MyType As FCCFormTypes = FCCFormTypes.def
    Private myPort As Intranet.intranet.jobs.Port
    Private ProgManager As String
    Private myCommentTypes As MedscreenLib.Glossary.PhraseCollection


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        DeptList = MedscreenLib.Glossary.Glossary.Personnel.DepartmentList("CUSTCENTRE")
        DeptListSales = MedscreenLib.Glossary.Glossary.Personnel.DepartmentList("SALES")
        DeptList2 = MedscreenLib.Glossary.Glossary.Personnel.DepartmentList("CUSTCENTRE")
        Me.cbPersonnel.DataSource = DeptList
        Me.cbPersonnel.DisplayMember = "Identity"
        Me.cbPersonnel.ValueMember = "Identity"

        Me.cbPersonnel2.DataSource = DeptList2
        Me.cbPersonnel2.DisplayMember = "Identity"
        Me.cbPersonnel2.ValueMember = "Identity"

        Me.myCommentTypes = New MedscreenLib.Glossary.PhraseCollection("Select phrase_id,phrase_text from phrase where phrase_type = 'COLLCOMTYP' and ICON < 'A'", _
            MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        Me.cbCommentType.DataSource = MedscreenLib.Glossary.Glossary.CollectionCommentsList
        Me.cbCommentType.DisplayMember = "PhraseText"
        Me.cbCommentType.ValueMember = "PhraseID"

        Me.cbOfficer.Hide()
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
    Friend WithEvents cbCommentType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents cbPersonnel As System.Windows.Forms.ComboBox
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents lblDestination As System.Windows.Forms.Label
    Friend WithEvents txtDestination As System.Windows.Forms.TextBox
    Friend WithEvents cbOfficer As System.Windows.Forms.ComboBox
    Friend WithEvents ckCopy As System.Windows.Forms.CheckBox
    Friend WithEvents cbPersonnel2 As System.Windows.Forms.ComboBox
    Friend WithEvents txtCopyDest As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCollComment))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.cbCommentType = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtComment = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblTo = New System.Windows.Forms.Label
        Me.cbPersonnel = New System.Windows.Forms.ComboBox
        Me.lblDestination = New System.Windows.Forms.Label
        Me.txtDestination = New System.Windows.Forms.TextBox
        Me.cbOfficer = New System.Windows.Forms.ComboBox
        Me.ckCopy = New System.Windows.Forms.CheckBox
        Me.cbPersonnel2 = New System.Windows.Forms.ComboBox
        Me.txtCopyDest = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 309)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(656, 56)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(568, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(488, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Add"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCommentType
        '
        Me.cbCommentType.Location = New System.Drawing.Point(128, 6)
        Me.cbCommentType.Name = "cbCommentType"
        Me.cbCommentType.Size = New System.Drawing.Size(344, 21)
        Me.cbCommentType.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Comment Type"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.txtComment)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(0, 48)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(656, 168)
        Me.Panel2.TabIndex = 4
        '
        'txtComment
        '
        Me.txtComment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtComment.Location = New System.Drawing.Point(8, 24)
        Me.txtComment.MaxLength = 400
        Me.txtComment.Multiline = True
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(632, 137)
        Me.txtComment.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Comment"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(24, 232)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(56, 23)
        Me.lblTo.TabIndex = 5
        Me.lblTo.Text = "Assign to"
        '
        'cbPersonnel
        '
        Me.cbPersonnel.Location = New System.Drawing.Point(104, 232)
        Me.cbPersonnel.Name = "cbPersonnel"
        Me.cbPersonnel.Size = New System.Drawing.Size(121, 21)
        Me.cbPersonnel.TabIndex = 6
        '
        'lblDestination
        '
        Me.lblDestination.Location = New System.Drawing.Point(240, 232)
        Me.lblDestination.Name = "lblDestination"
        Me.lblDestination.Size = New System.Drawing.Size(100, 23)
        Me.lblDestination.TabIndex = 7
        Me.lblDestination.Text = "Destination"
        '
        'txtDestination
        '
        Me.txtDestination.Location = New System.Drawing.Point(304, 232)
        Me.txtDestination.Name = "txtDestination"
        Me.txtDestination.Size = New System.Drawing.Size(336, 20)
        Me.txtDestination.TabIndex = 8
        '
        'cbOfficer
        '
        Me.cbOfficer.Location = New System.Drawing.Point(104, 232)
        Me.cbOfficer.Name = "cbOfficer"
        Me.cbOfficer.Size = New System.Drawing.Size(121, 21)
        Me.cbOfficer.TabIndex = 9
        '
        'ckCopy
        '
        Me.ckCopy.Checked = True
        Me.ckCopy.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckCopy.Location = New System.Drawing.Point(16, 272)
        Me.ckCopy.Name = "ckCopy"
        Me.ckCopy.Size = New System.Drawing.Size(80, 24)
        Me.ckCopy.TabIndex = 10
        Me.ckCopy.Text = "Send Copy"
        '
        'cbPersonnel2
        '
        Me.cbPersonnel2.Location = New System.Drawing.Point(104, 272)
        Me.cbPersonnel2.Name = "cbPersonnel2"
        Me.cbPersonnel2.Size = New System.Drawing.Size(121, 21)
        Me.cbPersonnel2.TabIndex = 11
        '
        'txtCopyDest
        '
        Me.txtCopyDest.Location = New System.Drawing.Point(304, 272)
        Me.txtCopyDest.Name = "txtCopyDest"
        Me.txtCopyDest.Size = New System.Drawing.Size(336, 20)
        Me.txtCopyDest.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(240, 272)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Destination"
        '
        'FrmCollComment
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 365)
        Me.Controls.Add(Me.txtCopyDest)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ckCopy)
        Me.Controls.Add(Me.txtDestination)
        Me.Controls.Add(Me.lblDestination)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbCommentType)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cbOfficer)
        Me.Controls.Add(Me.lblTo)
        Me.Controls.Add(Me.cbPersonnel2)
        Me.Controls.Add(Me.cbPersonnel)
        Me.Name = "FrmCollComment"
        Me.Text = "Collection Comments"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region "Exposed Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Mode form is to be run in 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormType() As FCCFormTypes
        Get
            Return MyType
        End Get
        Set(ByVal Value As FCCFormTypes)
            MyType = Value
            Me.Label3.Show()
            Me.txtCopyDest.Show()
            Me.cbPersonnel2.Show()
            Me.ckCopy.Show()
            Me.cbOfficer.Show()
            Me.Panel2.Height = 168
            Me.Text = "Collection Comment for CM " & myCMNumber
            If Value = FCCFormTypes.def Then        'If def mode a comment that will be assigned to another HO user
                Me.txtDestination.Hide()
                Me.lblDestination.Hide()
                Me.lblTo.Text = "&Assign to"
                Me.cbPersonnel.Show()
                Me.cbOfficer.Hide()
            ElseIf Value = FCCFormTypes.CollectorInfo Then
                Me.txtDestination.Show()
                Me.lblDestination.Show()
                Me.lblTo.Text = "&Send to"
                Me.cbPersonnel.Hide()
                Me.cbOfficer.Show()

            ElseIf Value = FCCFormTypes.CustomerIssue Then        'If def mode a comment that will be assigned to another HO user
                Me.txtDestination.Hide()
                Me.lblDestination.Hide()
                Me.lblTo.Text = "&Assign to"
                Me.cbPersonnel.DataSource = Me.DeptListSales
                Me.cbPersonnel.Show()
                Me.cbOfficer.Hide()
                Me.Label3.Hide()
                Me.Text = "Customer Comment for " & myCMNumber
            ElseIf Value = FCCFormTypes.customerComment Then
                Me.cbPersonnel.DataSource = Me.DeptListSales
                Me.txtDestination.Hide()
                Me.lblDestination.Hide()
                Me.lblTo.Text = ""
                Me.Text = "Customer Comment for " & myCMNumber
                Me.cbPersonnel.Hide()
                Me.cbOfficer.Show()
                Me.Label3.Hide()
                Me.txtCopyDest.Hide()
                Me.cbPersonnel2.Hide()
                Me.ckCopy.Hide()
                Me.cbOfficer.Hide()
                Me.Panel2.Height = 255
                Me.myCommentTypes = New MedscreenLib.Glossary.PhraseCollection("Select phrase_id,phrase_text from phrase where phrase_type = 'COLLCOMTYP' and ICON = 'CUSTOMER'", _
    MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                Me.cbCommentType.DataSource = myCommentTypes
                Me.cbCommentType.DisplayMember = "PhraseText"
                Me.cbCommentType.ValueMember = "PhraseID"
            ElseIf Value = FCCFormTypes.CollectorComment Then
                Me.cbPersonnel.DataSource = Me.DeptListSales
                Me.txtDestination.Hide()
                Me.lblDestination.Hide()
                Me.lblTo.Text = "&Send to"
                Me.cbPersonnel.Hide()
                Me.cbOfficer.Hide()
                Me.myCommentTypes = New MedscreenLib.Glossary.PhraseCollection("Select phrase_id,phrase_text from phrase where phrase_type = 'COLLCOMTYP' and ICON = 'OFFICER'", _
    MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
                Me.cbCommentType.DataSource = myCommentTypes
                Me.cbCommentType.DisplayMember = "PhraseText"
                Me.cbCommentType.ValueMember = "PhraseID"
            ElseIf Value = FCCFormTypes.Instrument Then
                Me.txtDestination.Hide()
                Me.lblDestination.Hide()
                Me.Text = "Comment for Instrument/Laboratory " & myCMNumber
                Me.lblTo.Hide()
                Me.cbPersonnel.Hide()
                Me.cbOfficer.Hide()
                Me.myCommentTypes = New MedscreenLib.Glossary.PhraseCollection("Select phrase_id,phrase_text from phrase where phrase_type = 'COLLCOMTYP' and ICON = 'INSTRUMENT'", _
    MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)

                Me.cbCommentType.DataSource = myCommentTypes
                Me.cbCommentType.DisplayMember = "PhraseText"
                Me.cbCommentType.ValueMember = "PhraseID"

            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CM Number being commented on 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myCMNumber As String = ""
    Public WriteOnly Property CMNumber() As String
        Set(ByVal Value As String)
            Me.Text += " " & Value
            myCMNumber = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Person program managing this collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ProgamManager() As String
        Get
            Return strPM
        End Get
        Set(ByVal Value As String)
            strPM = Value
            If Value.Trim.Length > 0 Then
                If MyType = FCCFormTypes.def Then
                    Me.cbPersonnel.Show()
                    Me.cbPersonnel.BringToFront()
                    Dim i As Integer
                    For i = 0 To Me.cbPersonnel.Items.Count - 1
                        If CType(Me.cbPersonnel.Items.Item(i), MedscreenLib.personnel.UserPerson).Identity = Value Then
                            Me.cbPersonnel.SelectedIndex = i
                            Me.txtDestination.Text = CType(Me.cbPersonnel.Items.Item(i), MedscreenLib.personnel.UserPerson).Email
                            Exit For
                        End If
                    Next
                Else
                    Dim i As Integer
                    For i = 0 To Me.cbPersonnel2.Items.Count - 1
                        If CType(Me.cbPersonnel2.Items.Item(i), MedscreenLib.personnel.UserPerson).Identity = Value Then
                            Me.cbPersonnel2.SelectedIndex = i
                            Me.txtCopyDest.Text = CType(Me.cbPersonnel2.Items.Item(i), MedscreenLib.personnel.UserPerson).Email
                            Exit For
                        End If
                    Next
                End If

            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Port for collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Port() As Intranet.intranet.jobs.Port
        Set(ByVal Value As Intranet.intranet.jobs.Port)
            myPort = Value
            If Not myPort Is Nothing Then
                Me.cbOfficer.DataSource = myPort.CollectingOfficers
                Me.cbOfficer.DisplayMember = "OfficerId"
                Me.cbOfficer.ValueMember = "OfficerId"
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collecting Officer doing collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CollOfficer() As String
        Get
            Dim oOff As Intranet.intranet.jobs.CollPort
            oOff = myPort.CollectingOfficers.Item(Me.cbOfficer.SelectedIndex)
            If Not oOff Is Nothing Then
                Return oOff.OfficerId
            End If
        End Get
        Set(ByVal Value As String)
            Dim oOff As Intranet.intranet.jobs.CollPort
            Dim i As Integer
            For i = 0 To myPort.CollectingOfficers.Count - 1
                oOff = myPort.CollectingOfficers.Item(i)
                If oOff.OfficerId.ToUpper = Value.ToUpper Then
                    Me.cbOfficer.SelectedIndex = i
                    'If Me.ProgamManager.trim.length = 0 Then
                    Me.txtDestination.Text = oOff.Officer.SendDestination
                    Me.cbPersonnel.Hide()
                    Me.cbOfficer.BringToFront()
                    Exit For
                    '    'Me.cbOfficer.Top = 232
                    '    Me.cbPersonnel2.Show()
                    'Else
                    '    Me.txtCopyDest.Text = oOff.Officer.SendDestination
                    '    'Me.cbOfficer.Top = 272
                    '    
                    'End If

                End If
            Next
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of comment
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CommentType() As String
        Get
            Return strCommentType
        End Get
        Set(ByVal Value As String)
            strCommentType = Value
            'Dim oPos As Intranet.intranet.glossary.Phrase = ModTpanel.CollectionCommentsList.Item(Value)
            'If Not oPos Is Nothing Then
            Dim intPos As Integer = Me.myCommentTypes.IndexOf(Value)
            If intPos > -1 Then
                Me.cbCommentType.SelectedIndex = intPos
            End If
            'End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Comment to be used 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Comment() As String
        Get
            Return Me.txtComment.Text
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Who will carry out the action required by this comment
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private strActionBy As String = Nothing
    Public Property ActionBy() As String
        Get
            If strActionBy Is Nothing Then
                Dim objP As MedscreenLib.personnel.UserPerson = DeptList.Item(Me.cbPersonnel.SelectedIndex)
                If Not objP Is Nothing Then
                    Return objP.Identity
                End If
            Else
                Return strActionBy
            End If
        End Get
        Set(ByVal Value As String)
            strActionBy = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Who to be sent info
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SendTo() As String
        Get
            Dim strText As String = Me.txtDestination.Text
            strText = MedscreenLib.Medscreen.ReplaceString(strText, " ", "")
            Dim outAddress As String() = strText.Split(New Char() {",", ";"})
            strText = ""
            Dim I As Integer
            For I = 0 To outAddress.Length - 1
                If strText.Length > 0 Then strText += ","
                Dim strTemp As String = outAddress.GetValue(I)
                If IsNumber(strTemp) Then strTemp += MedscreenLib.Constants.GCST_EMAIL_To_Fax_Send
                strText += strTemp
            Next
            Return strText
        End Get
        Set(ByVal Value As String)
            Me.txtDestination.Text = Value
        End Set
    End Property

    '<Added code modified 11-Apr-2007 14:32 by LOGOS\taylor> 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Destination to copy to 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CopyDestTo() As String
        Get
            Dim strText As String = Me.txtCopyDest.Text
            strText = MedscreenLib.Medscreen.ReplaceString(strText, " ", "")
            Dim outAddress As String() = strText.Split(New Char() {",", ";"})
            strText = ""
            Dim I As Integer
            For I = 0 To outAddress.Length - 1
                If strText.Length > 0 Then strText += ","
                Dim strTemp As String = outAddress.GetValue(I)
                If IsNumber(strTemp) Then strTemp += MedscreenLib.Constants.GCST_EMAIL_To_Fax_Send
                strText += strTemp
            Next
            Return strText
        End Get
    End Property
    '</Added code modified 11-Apr-2007 14:32 by LOGOS\taylor> 
    '<Added code modified 11-Apr-2007 14:38 by LOGOS\taylor> 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Does the user want a copy sent
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Copywanted() As Boolean
        Get
            Return Me.ckCopy.Checked
        End Get
    End Property
    '</Added code modified 11-Apr-2007 14:38 by LOGOS\taylor> 

#End Region

    Private Sub cbPersonnel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPersonnel.SelectedIndexChanged
        If MyType = FCCFormTypes.def Then
            Me.txtDestination.Text = CType(Me.cbPersonnel.Items.Item(cbPersonnel.SelectedIndex), MedscreenLib.personnel.UserPerson).Email
        End If
    End Sub

    Private Sub cbPersonnel2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPersonnel2.SelectedIndexChanged
        Me.txtCopyDest.Text = CType(Me.cbPersonnel2.Items.Item(cbPersonnel2.SelectedIndex), _
        MedscreenLib.personnel.UserPerson).Email
    End Sub

    Private Sub cbOfficer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOfficer.SelectedIndexChanged
        If MyType = FCCFormTypes.CollectorInfo Then
            Dim CollPort As Intranet.intranet.jobs.CollPort

            CollPort = CType(Me.cbOfficer.Items.Item(cbOfficer.SelectedIndex), Intranet.intranet.jobs.CollPort)
            If Not CollPort Is Nothing Then
                If Not CollPort.Officer Is Nothing Then
                    If CollPort.Officer.SendOption.Trim.Length = 0 Then   ' if the data hasn't been read 
                        CollPort.Officer.Load()

                    End If
                    Me.txtDestination.Text = CollPort.Officer.SendDestination
                End If
            End If
        End If
    End Sub

    Private Sub cbCommentType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCommentType.SelectedIndexChanged
        If Me.cbCommentType.SelectedIndex = -1 Then Exit Sub
        Dim objPhrase As MedscreenLib.Glossary.Phrase = Me.cbCommentType.SelectedItem
        Me.strCommentType = objPhrase.PhraseID
        Me.txtCopyDest.Show()
        Me.cbOfficer.Show()
        Me.cbPersonnel.Show()
        Me.cbPersonnel2.Show()
        Me.txtDestination.Show()
        Me.ckCopy.Show()
        Me.Label3.Show()
        Me.lblTo.Show()
        If Me.strCommentType = "COLL" Then
            Me.txtCopyDest.Hide()
            Me.txtDestination.Hide()
            Me.ckCopy.Hide()
            Me.Label3.Hide()
            Me.lblTo.Hide()
            Me.ckCopy.Checked = False
            Me.cbOfficer.Hide()
            Me.cbPersonnel.Hide()
            Me.cbPersonnel2.Hide()
            Me.ActionBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
        End If
    End Sub

    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    End Sub
End Class
