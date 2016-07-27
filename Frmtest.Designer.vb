<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frmtest
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.VolDiscountPanel1 = New MedscreenCommonGui.Panels.VolDiscountPanel
        Me.SuspendLayout()
        '
        'VolDiscountPanel1
        '
        Me.VolDiscountPanel1.CustomerID = Nothing
        Me.VolDiscountPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.VolDiscountPanel1.Location = New System.Drawing.Point(0, 0)
        Me.VolDiscountPanel1.Name = "VolDiscountPanel1"
        Me.VolDiscountPanel1.Size = New System.Drawing.Size(1089, 840)
        Me.VolDiscountPanel1.TabIndex = 0
        '
        'Frmtest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1089, 840)
        Me.Controls.Add(Me.VolDiscountPanel1)
        Me.Name = "Frmtest"
        Me.Text = "Frmtest"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RoutinePanel As System.Windows.Forms.Panel
    Friend WithEvents VolumePanel As System.Windows.Forms.Panel
    Friend WithEvents VolDiscountPanel1 As MedscreenCommonGui.Panels.VolDiscountPanel

End Class
