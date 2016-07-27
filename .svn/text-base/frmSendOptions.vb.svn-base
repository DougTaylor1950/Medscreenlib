'$Revision$
'$Author$
'$Date$
'$Log$
'Revision 1.2  2005-06-17 07:05:25+01  taylor
'Milestone check in
'
'Revision 1.1  2004-05-06 15:00:34+01  taylor
'<>
'
'<>
Imports System.Windows.Forms

Public Class frmSendOptions
    Inherits System.Windows.Forms.Form
    Private objContMethodPhrase As MedscreenLib.Glossary.PhraseCollection

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        objContMethodPhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_METH")
        objContMethodPhrase.Load()

        Me.cbDestType.DataSource = objContMethodPhrase
        Me.cbDestType.DisplayMember = "PhraseText"
        Me.cbDestType.ValueMember = "PhraseID"

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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCMNumber As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ckModify As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbDestType As System.Windows.Forms.ComboBox
    Friend WithEvents txtDestAddr As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSendOptions))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtDestAddr = New System.Windows.Forms.TextBox()
        Me.cbDestType = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ckModify = New System.Windows.Forms.CheckBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.txtCMNumber = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 189)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(352, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(213, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(127, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 8
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox1, Me.RadioButton2, Me.RadioButton1, Me.txtCMNumber, Me.Label1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(352, 189)
        Me.Panel2.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtDestAddr, Me.cbDestType, Me.Label3, Me.Label2, Me.ckModify})
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(24, 81)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(296, 95)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'txtDestAddr
        '
        Me.txtDestAddr.Location = New System.Drawing.Point(142, 48)
        Me.txtDestAddr.Name = "txtDestAddr"
        Me.txtDestAddr.Size = New System.Drawing.Size(139, 20)
        Me.txtDestAddr.TabIndex = 4
        Me.txtDestAddr.Text = ""
        '
        'cbDestType
        '
        Me.cbDestType.Location = New System.Drawing.Point(142, 19)
        Me.cbDestType.Name = "cbDestType"
        Me.cbDestType.Size = New System.Drawing.Size(144, 21)
        Me.cbDestType.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(23, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(106, 23)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Destination &Address"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(29, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Destination &type"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ckModify
        '
        Me.ckModify.Location = New System.Drawing.Point(18, -5)
        Me.ckModify.Name = "ckModify"
        Me.ckModify.Size = New System.Drawing.Size(124, 24)
        Me.ckModify.TabIndex = 0
        Me.ckModify.Text = "&Modify Destination"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(192, 47)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(134, 24)
        Me.RadioButton2.TabIndex = 3
        Me.RadioButton2.Text = "&Print Collection Only"
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(24, 48)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(149, 24)
        Me.RadioButton1.TabIndex = 2
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Send by &default method"
        '
        'txtCMNumber
        '
        Me.txtCMNumber.Enabled = False
        Me.txtCMNumber.Location = New System.Drawing.Point(152, 13)
        Me.txtCMNumber.Name = "txtCMNumber"
        Me.txtCMNumber.TabIndex = 1
        Me.txtCMNumber.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Collection Number"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSendOptions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 229)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSendOptions"
        Me.Text = "Send Collection"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public WriteOnly Property CMNumber() As String
        Set(ByVal Value As String)
            Me.txtCMNumber.Text = Value
        End Set
    End Property

    Public ReadOnly Property Modify() As Boolean
        Get
            Return Me.ckModify.Checked
        End Get
    End Property

    Public ReadOnly Property PrintOnly() As Boolean
        Get
            Return Me.RadioButton2.Checked
        End Get
    End Property

    Public Property DestAddress() As String
        Get
            Return Me.txtDestAddr.Text
        End Get
        Set(ByVal value As String)
            Me.txtDestAddr.Text = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set destination type
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	22/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DestType() As String
        Set(ByVal Value As String)
            Me.cbDestType.SelectedValue = Value

        End Set
        Get
            Dim sdestType As String

            'If cbDestType.Text = "Contact By HOME FAX" Then
            '    sdestType = "FAX"
            'ElseIf cbDestType.Text = "Contact By WORK FAX" Then
            '    sdestType = "FAX"
            'ElseIf cbDestType.Text = "Contact By HOME E-MAIL" Then
            '    sdestType = "EMAIL"
            'ElseIf cbDestType.Text = "Contact By WORK E-MAIL" Then
            '    sdestType = "EMAIL"
            'End If
            sdestType = Me.cbDestType.SelectedValue
            Return sdestType
        End Get
    End Property

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub cbDestType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDestType.SelectedIndexChanged
        Dim oColl As New Collection()
        If Me.txtCMNumber.Text.Trim.Length = 0 Then Exit Sub
        oColl.Add(Me.txtCMNumber.Text.PadLeft(10).Replace(" ", "0").Replace("C", "0").Replace("M", "0"))
        If Me.DestType.Trim.Length = 0 Then Exit Sub
        oColl.Add(Me.DestType)
        Dim strDestination As String = MedscreenLib.CConnection.PackageStringList("lib_collection.OfficerSendDestination", oColl)
        Me.txtDestAddr.Text = strDestination
    End Sub
End Class
