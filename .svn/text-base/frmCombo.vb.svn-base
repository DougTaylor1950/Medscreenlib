Public Class frmCombo
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
  Friend WithEvents cmbOptions As System.Windows.Forms.ComboBox
  Friend WithEvents lblOptions As System.Windows.Forms.Label
  Friend WithEvents btnCancel As System.Windows.Forms.Button
  Friend WithEvents btnOK As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCombo))
    Me.cmbOptions = New System.Windows.Forms.ComboBox()
    Me.lblOptions = New System.Windows.Forms.Label()
    Me.btnCancel = New System.Windows.Forms.Button()
    Me.btnOK = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'cmbOptions
    '
    Me.cmbOptions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cmbOptions.Location = New System.Drawing.Point(52, 12)
    Me.cmbOptions.Name = "cmbOptions"
    Me.cmbOptions.Size = New System.Drawing.Size(176, 21)
    Me.cmbOptions.TabIndex = 1
    '
    'lblOptions
    '
    Me.lblOptions.Location = New System.Drawing.Point(8, 16)
    Me.lblOptions.Name = "lblOptions"
    Me.lblOptions.Size = New System.Drawing.Size(44, 16)
    Me.lblOptions.TabIndex = 0
    Me.lblOptions.Text = "&Options"
    '
    'btnCancel
    '
    Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Top
    Me.btnCancel.BackColor = System.Drawing.SystemColors.Window
    Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Bitmap)
    Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.btnCancel.Location = New System.Drawing.Point(132, 52)
    Me.btnCancel.Name = "btnCancel"
    Me.btnCancel.Size = New System.Drawing.Size(68, 28)
    Me.btnCancel.TabIndex = 3
    Me.btnCancel.Text = "&Cancel"
    Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'btnOK
    '
    Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Top
    Me.btnOK.BackColor = System.Drawing.SystemColors.Window
    Me.btnOK.Image = CType(resources.GetObject("btnOK.Image"), System.Drawing.Bitmap)
    Me.btnOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.btnOK.Location = New System.Drawing.Point(48, 52)
    Me.btnOK.Name = "btnOK"
    Me.btnOK.Size = New System.Drawing.Size(52, 28)
    Me.btnOK.TabIndex = 2
    Me.btnOK.Text = "&OK"
    Me.btnOK.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'frmCombo
    '
    Me.AcceptButton = Me.btnOK
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.CancelButton = Me.btnCancel
    Me.ClientSize = New System.Drawing.Size(236, 97)
    Me.ControlBox = False
    Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOK, Me.btnCancel, Me.lblOptions, Me.cmbOptions})
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmCombo"
    Me.ShowInTaskbar = False
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Select Item"
    Me.ResumeLayout(False)

  End Sub

#End Region
    Public CBText As String = ""

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub

    Private Sub cmbOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOptions.SelectedIndexChanged
        btnOK.Enabled = cmbOptions.SelectedIndex >= 0
        CBText = cmbOptions.Text
    End Sub

    Public WriteOnly Property List() As String()
        Set(ByVal Value As String())
            Dim i As Integer
            Me.cmbOptions.Items.Clear()
            For i = 0 To Value.Length - 1
                Me.cmbOptions.Items.Add(Value.GetValue(i))
            Next
        End Set
    End Property

End Class
