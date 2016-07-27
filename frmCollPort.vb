'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 08:44:22+00 $
'$Log: frmCollPort.vb,v $
'Revision 1.0  2005-11-29 08:44:22+00  taylor
'Commented
'
Imports Intranet
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCollPort
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to get/display data for a collecting officer and port 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCollPort
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents CKAgreedCosts As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCosts As System.Windows.Forms.TextBox
    Friend WithEvents LblCurrency As System.Windows.Forms.Label
    Friend WithEvents CKRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCollPort))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.CKAgreedCosts = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCosts = New System.Windows.Forms.TextBox()
        Me.LblCurrency = New System.Windows.Forms.Label()
        Me.CKRemoved = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 229)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(512, 32)
        Me.Panel1.TabIndex = 2
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(428, 4)
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
        Me.cmdOk.Location = New System.Drawing.Point(348, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CKAgreedCosts
        '
        Me.CKAgreedCosts.Location = New System.Drawing.Point(24, 24)
        Me.CKAgreedCosts.Name = "CKAgreedCosts"
        Me.CKAgreedCosts.Size = New System.Drawing.Size(288, 32)
        Me.CKAgreedCosts.TabIndex = 3
        Me.CKAgreedCosts.Text = "Costs contractually agreed with Collecting Officer"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Costs"
        '
        'txtCosts
        '
        Me.txtCosts.Location = New System.Drawing.Point(72, 61)
        Me.txtCosts.Name = "txtCosts"
        Me.txtCosts.TabIndex = 5
        Me.txtCosts.Text = ""
        '
        'LblCurrency
        '
        Me.LblCurrency.Location = New System.Drawing.Point(168, 62)
        Me.LblCurrency.Name = "LblCurrency"
        Me.LblCurrency.TabIndex = 6
        '
        'CKRemoved
        '
        Me.CKRemoved.Location = New System.Drawing.Point(24, 104)
        Me.CKRemoved.Name = "CKRemoved"
        Me.CKRemoved.TabIndex = 7
        Me.CKRemoved.Text = "Removed"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 136)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Comments"
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(24, 160)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(456, 56)
        Me.txtComments.TabIndex = 9
        Me.txtComments.Text = ""
        '
        'frmCollPort
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(512, 261)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtComments, Me.Label2, Me.CKRemoved, Me.LblCurrency, Me.txtCosts, Me.Label1, Me.CKAgreedCosts, Me.Panel1})
        Me.Name = "frmCollPort"
        Me.Text = "Collection Officer Port Details"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicate whether costs are contractually agreed 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AgreedCosts() As Boolean
        Get
            Return Me.CKAgreedCosts.Checked
        End Get
        Set(ByVal Value As Boolean)
            Me.CKAgreedCosts.Checked = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Comments added about this officer and port 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Comments() As String
        Get
            Return Me.txtComments.Text
        End Get
        Set(ByVal Value As String)
            Me.txtComments.Text = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removed this officer no longer works this port 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Removed() As Boolean
        Get
            Return Me.CKRemoved.Checked
        End Get
        Set(ByVal Value As Boolean)
            Me.CKRemoved.Checked = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cost of using this officer, may be estimate or accurate agreed amount
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Cost() As Double
        Get
            Return Me.txtCosts.Text
        End Get
        Set(ByVal Value As Double)
            Me.txtCosts.Text = Value.ToString("0.00")
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Currency for cost 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Currency() As String
        Set(ByVal Value As String)
            Me.LblCurrency.Text = Value
        End Set
    End Property

End Class
