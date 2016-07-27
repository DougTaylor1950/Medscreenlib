Public Class frmOptions
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim myPhraseList As MedscreenLib.Glossary.PhraseCollection = MedscreenLib.Glossary.Glossary.PhraseList.Item("OPTTYPE")
        Me.cbOptionType.DataSource = myPhraseList
        Me.cbOptionType.ValueMember = "PhraseID"
        Me.cbOptionType.DisplayMember = "PhraseText"
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
    Friend WithEvents txtIdentity As System.Windows.Forms.TextBox
    Friend WithEvents lblDestination As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHelpComment As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbOptionType As System.Windows.Forms.ComboBox
    Friend WithEvents txtFunctionalGroup As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents udAuthorityLevel As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtLibrary As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtRoutine As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents udBoolean As System.Windows.Forms.DomainUpDown
    Friend WithEvents txtDefault As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmOptions))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.txtIdentity = New System.Windows.Forms.TextBox()
        Me.lblDestination = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtHelpComment = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbOptionType = New System.Windows.Forms.ComboBox()
        Me.txtFunctionalGroup = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.udAuthorityLevel = New System.Windows.Forms.NumericUpDown()
        Me.txtLibrary = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtRoutine = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.udBoolean = New System.Windows.Forms.DomainUpDown()
        Me.txtDefault = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        CType(Me.udAuthorityLevel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 325)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(616, 56)
        Me.Panel1.TabIndex = 9
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(528, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(448, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtIdentity
        '
        Me.txtIdentity.Location = New System.Drawing.Point(136, 8)
        Me.txtIdentity.Name = "txtIdentity"
        Me.txtIdentity.Size = New System.Drawing.Size(120, 20)
        Me.txtIdentity.TabIndex = 0
        Me.txtIdentity.Text = ""
        '
        'lblDestination
        '
        Me.lblDestination.Location = New System.Drawing.Point(8, 8)
        Me.lblDestination.Name = "lblDestination"
        Me.lblDestination.TabIndex = 9
        Me.lblDestination.Text = "Option Id"
        Me.lblDestination.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(136, 32)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(440, 20)
        Me.txtDescription.TabIndex = 2
        Me.txtDescription.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Description"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHelpComment
        '
        Me.txtHelpComment.Location = New System.Drawing.Point(136, 64)
        Me.txtHelpComment.Multiline = True
        Me.txtHelpComment.Name = "txtHelpComment"
        Me.txtHelpComment.Size = New System.Drawing.Size(440, 64)
        Me.txtHelpComment.TabIndex = 3
        Me.txtHelpComment.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Long Description"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(288, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Option Type"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbOptionType
        '
        Me.cbOptionType.Location = New System.Drawing.Point(432, 8)
        Me.cbOptionType.Name = "cbOptionType"
        Me.cbOptionType.Size = New System.Drawing.Size(121, 21)
        Me.cbOptionType.TabIndex = 1
        '
        'txtFunctionalGroup
        '
        Me.txtFunctionalGroup.Location = New System.Drawing.Point(136, 136)
        Me.txtFunctionalGroup.Name = "txtFunctionalGroup"
        Me.txtFunctionalGroup.Size = New System.Drawing.Size(120, 20)
        Me.txtFunctionalGroup.TabIndex = 4
        Me.txtFunctionalGroup.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Functional Group"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 169)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Authority Level"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udAuthorityLevel
        '
        Me.udAuthorityLevel.Location = New System.Drawing.Point(136, 169)
        Me.udAuthorityLevel.Name = "udAuthorityLevel"
        Me.udAuthorityLevel.TabIndex = 6
        '
        'txtLibrary
        '
        Me.txtLibrary.Location = New System.Drawing.Point(136, 208)
        Me.txtLibrary.Name = "txtLibrary"
        Me.txtLibrary.Size = New System.Drawing.Size(120, 20)
        Me.txtLibrary.TabIndex = 7
        Me.txtLibrary.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Library"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRoutine
        '
        Me.txtRoutine.Location = New System.Drawing.Point(408, 208)
        Me.txtRoutine.Name = "txtRoutine"
        Me.txtRoutine.Size = New System.Drawing.Size(120, 20)
        Me.txtRoutine.TabIndex = 8
        Me.txtRoutine.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(272, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Routine"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(408, 136)
        Me.txtReference.Name = "txtReference"
        Me.txtReference.Size = New System.Drawing.Size(120, 20)
        Me.txtReference.TabIndex = 5
        Me.txtReference.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(272, 136)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Reference"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(272, 168)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Default"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udBoolean
        '
        Me.udBoolean.Items.Add("False")
        Me.udBoolean.Items.Add("True")
        Me.udBoolean.Location = New System.Drawing.Point(408, 168)
        Me.udBoolean.Name = "udBoolean"
        Me.udBoolean.TabIndex = 27
        '
        'txtDefault
        '
        Me.txtDefault.Location = New System.Drawing.Point(408, 168)
        Me.txtDefault.Name = "txtDefault"
        Me.txtDefault.Size = New System.Drawing.Size(120, 20)
        Me.txtDefault.TabIndex = 28
        Me.txtDefault.Text = ""
        '
        'frmOptions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 381)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtDefault, Me.udBoolean, Me.Label9, Me.txtReference, Me.Label8, Me.txtRoutine, Me.Label7, Me.txtLibrary, Me.Label6, Me.udAuthorityLevel, Me.Label5, Me.txtFunctionalGroup, Me.Label4, Me.cbOptionType, Me.Label3, Me.txtHelpComment, Me.Label2, Me.txtDescription, Me.Label1, Me.txtIdentity, Me.lblDestination, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOptions"
        Me.Text = "Options Edit"
        Me.Panel1.ResumeLayout(False)
        CType(Me.udAuthorityLevel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declarations"
    Private myOption As MedscreenLib.Glossary.Options
#End Region

#Region "Public Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose option 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Property _Option() As MedscreenLib.Glossary.Options
        Get
            Return myOption
        End Get
        Set(ByVal Value As MedscreenLib.Glossary.Options)
            myOption = Value
            FillForm()
        End Set
    End Property


#End Region

#Region "Local methods"
    Private Sub FillForm()
        If myOption Is Nothing Then Exit Sub

        With myOption
            Me.txtIdentity.Text = .OptionID
            Me.txtDescription.Text = .OptionDescription
            Me.txtHelpComment.Text = .HelpComment
            Me.cbOptionType.SelectedValue = .OptionType
            Me.udAuthorityLevel.Value = .OptionAuthority
            Me.txtFunctionalGroup.Text = .FunctionalGroup
            Me.txtLibrary.Text = .Library
            Me.txtRoutine.Text = .Routine
            Me.txtReference.Text = .Reference
            Me.udBoolean.Hide()
            Me.txtDefault.Hide()
            If .OptionType = "BOOL" Then
                Me.udBoolean.Show()
                Me.udBoolean.BringToFront()
                If .DefaultValue = "T" Then
                    Me.udBoolean.SelectedIndex = 1
                Else
                    Me.udBoolean.SelectedIndex = 0
                End If
            Else
                Me.txtDefault.Show()
                Me.txtDefault.BringToFront()
                Me.txtDefault.Text = .DefaultValue
            End If
        End With
    End Sub

    Private Sub GetForm()
        If myOption Is Nothing Then Exit Sub

        With myOption
            .OptionID = Me.txtIdentity.Text
            .OptionDescription = Me.txtDescription.Text
            .HelpComment = Me.txtHelpComment.Text
            .OptionType = Me.cbOptionType.SelectedValue
            .OptionAuthority = Me.udAuthorityLevel.Value
            .FunctionalGroup = Me.txtFunctionalGroup.Text
            .Library = Me.txtLibrary.Text
            .Routine = Me.txtRoutine.Text
            .Reference = Me.txtReference.Text
            If .OptionType = "BOOL" Then
                If Me.udBoolean.SelectedIndex = 0 Then
                    .DefaultValue = "F"
                Else
                    .DefaultValue = "T"
                End If
            Else
                .DefaultValue = Me.txtDefault.Text
            End If

        End With
    End Sub
#End Region

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        GetForm()
    End Sub

    Private Sub cbOptionType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOptionType.SelectedIndexChanged

    End Sub
End Class
