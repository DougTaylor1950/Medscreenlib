<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRandomNumbers
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRandomNumbers))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblCollection = New System.Windows.Forms.Label
        Me.lblFilename = New System.Windows.Forms.Label
        Me.cmdSendRandom = New System.Windows.Forms.Button
        Me.lvRandom = New System.Windows.Forms.DataGridView
        Me.ctxCopyReplace = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.txtEditor = New System.Windows.Forms.TextBox
        Me.cmdCreateDocument = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.udEditor = New System.Windows.Forms.NumericUpDown
        Me.lvColumnGroups = New ListViewEx.ListViewEx
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.lblTemplate = New System.Windows.Forms.Label
        Me.cmdEditDoc = New System.Windows.Forms.Button
        Me.cmdGenerateRandom = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.CustomerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CollectionOptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.lvRandom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ctxCopyReplace.SuspendLayout()
        CType(Me.udEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(905, 583)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblCollection)
        Me.Panel1.Controls.Add(Me.lblFilename)
        Me.Panel1.Controls.Add(Me.cmdSendRandom)
        Me.Panel1.Controls.Add(Me.lvRandom)
        Me.Panel1.Controls.Add(Me.txtEditor)
        Me.Panel1.Controls.Add(Me.cmdCreateDocument)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.udEditor)
        Me.Panel1.Controls.Add(Me.lvColumnGroups)
        Me.Panel1.Controls.Add(Me.lblTemplate)
        Me.Panel1.Controls.Add(Me.cmdEditDoc)
        Me.Panel1.Controls.Add(Me.cmdGenerateRandom)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1063, 577)
        Me.Panel1.TabIndex = 1
        '
        'lblCollection
        '
        Me.lblCollection.AutoSize = True
        Me.lblCollection.Location = New System.Drawing.Point(212, 4)
        Me.lblCollection.Name = "lblCollection"
        Me.lblCollection.Size = New System.Drawing.Size(0, 13)
        Me.lblCollection.TabIndex = 28
        '
        'lblFilename
        '
        Me.lblFilename.AutoSize = True
        Me.lblFilename.Location = New System.Drawing.Point(163, 217)
        Me.lblFilename.Name = "lblFilename"
        Me.lblFilename.Size = New System.Drawing.Size(0, 13)
        Me.lblFilename.TabIndex = 27
        '
        'cmdSendRandom
        '
        Me.cmdSendRandom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSendRandom.Image = Global.MedscreenCommonGui.AdcResource.Mail1
        Me.cmdSendRandom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSendRandom.Location = New System.Drawing.Point(758, 315)
        Me.cmdSendRandom.Name = "cmdSendRandom"
        Me.cmdSendRandom.Size = New System.Drawing.Size(130, 23)
        Me.cmdSendRandom.TabIndex = 25
        Me.cmdSendRandom.Text = "Send Document"
        Me.cmdSendRandom.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdSendRandom.UseVisualStyleBackColor = True
        '
        'lvRandom
        '
        Me.lvRandom.AllowUserToOrderColumns = True
        Me.lvRandom.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvRandom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.lvRandom.ContextMenuStrip = Me.ctxCopyReplace
        Me.lvRandom.Location = New System.Drawing.Point(163, 250)
        Me.lvRandom.Name = "lvRandom"
        Me.lvRandom.Size = New System.Drawing.Size(576, 320)
        Me.lvRandom.TabIndex = 24
        '
        'ctxCopyReplace
        '
        Me.ctxCopyReplace.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem, Me.CutToolStripMenuItem})
        Me.ctxCopyReplace.Name = "ctxCopyReplace"
        Me.ctxCopyReplace.Size = New System.Drawing.Size(153, 70)
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.CopyToolStripMenuItem.Text = "Copy "
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.PasteToolStripMenuItem.Text = "Paste"
        '
        'CutToolStripMenuItem
        '
        Me.CutToolStripMenuItem.Name = "CutToolStripMenuItem"
        Me.CutToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.CutToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.CutToolStripMenuItem.Text = "Cut"
        '
        'txtEditor
        '
        Me.txtEditor.Location = New System.Drawing.Point(643, 34)
        Me.txtEditor.Name = "txtEditor"
        Me.txtEditor.Size = New System.Drawing.Size(100, 20)
        Me.txtEditor.TabIndex = 23
        Me.txtEditor.Visible = False
        '
        'cmdCreateDocument
        '
        Me.cmdCreateDocument.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCreateDocument.Image = Global.MedscreenCommonGui.AdcResource.NewDoc
        Me.cmdCreateDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCreateDocument.Location = New System.Drawing.Point(758, 236)
        Me.cmdCreateDocument.Name = "cmdCreateDocument"
        Me.cmdCreateDocument.Size = New System.Drawing.Size(130, 23)
        Me.cmdCreateDocument.TabIndex = 22
        Me.cmdCreateDocument.Text = "Create Document"
        Me.cmdCreateDocument.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdCreateDocument.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(160, 62)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Column Groups"
        '
        'udEditor
        '
        Me.udEditor.Location = New System.Drawing.Point(493, 34)
        Me.udEditor.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.udEditor.Name = "udEditor"
        Me.udEditor.Size = New System.Drawing.Size(120, 20)
        Me.udEditor.TabIndex = 20
        Me.udEditor.Visible = False
        '
        'lvColumnGroups
        '
        Me.lvColumnGroups.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvColumnGroups.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.ColumnHeader1})
        Me.lvColumnGroups.FullRowSelect = True
        Me.lvColumnGroups.Location = New System.Drawing.Point(163, 83)
        Me.lvColumnGroups.Name = "lvColumnGroups"
        Me.lvColumnGroups.Size = New System.Drawing.Size(576, 118)
        Me.lvColumnGroups.TabIndex = 19
        Me.lvColumnGroups.UseCompatibleStateImageBehavior = False
        Me.lvColumnGroups.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Column Group"
        Me.ColumnHeader10.Width = 129
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "No of Donors"
        Me.ColumnHeader11.Width = 137
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "From"
        Me.ColumnHeader12.Width = 169
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Import Type"
        Me.ColumnHeader1.Width = 100
        '
        'lblTemplate
        '
        Me.lblTemplate.AutoSize = True
        Me.lblTemplate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemplate.Location = New System.Drawing.Point(212, 30)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.Size = New System.Drawing.Size(59, 15)
        Me.lblTemplate.TabIndex = 18
        Me.lblTemplate.Text = "Template :"
        '
        'cmdEditDoc
        '
        Me.cmdEditDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEditDoc.Image = Global.MedscreenCommonGui.AdcResource.NOTES_4
        Me.cmdEditDoc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEditDoc.Location = New System.Drawing.Point(758, 276)
        Me.cmdEditDoc.Name = "cmdEditDoc"
        Me.cmdEditDoc.Size = New System.Drawing.Size(130, 23)
        Me.cmdEditDoc.TabIndex = 17
        Me.cmdEditDoc.Text = "Edit Document"
        Me.cmdEditDoc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdEditDoc.UseVisualStyleBackColor = True
        '
        'cmdGenerateRandom
        '
        Me.cmdGenerateRandom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGenerateRandom.Image = Global.MedscreenCommonGui.AdcResource.CHECKMRK
        Me.cmdGenerateRandom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdGenerateRandom.Location = New System.Drawing.Point(824, 62)
        Me.cmdGenerateRandom.Name = "cmdGenerateRandom"
        Me.cmdGenerateRandom.Size = New System.Drawing.Size(75, 23)
        Me.cmdGenerateRandom.TabIndex = 16
        Me.cmdGenerateRandom.Text = "Generate"
        Me.cmdGenerateRandom.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Multiselect = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.AlwaysBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomerToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.HelpProvider1.SetShowHelp(Me.MenuStrip1, True)
        Me.MenuStrip1.Size = New System.Drawing.Size(150, 24)
        Me.MenuStrip1.TabIndex = 26
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'CustomerToolStripMenuItem
        '
        Me.CustomerToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CollectionOptionsToolStripMenuItem})
        Me.CustomerToolStripMenuItem.Name = "CustomerToolStripMenuItem"
        Me.CustomerToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.CustomerToolStripMenuItem.Text = "Customer"
        '
        'CollectionOptionsToolStripMenuItem
        '
        Me.CollectionOptionsToolStripMenuItem.Name = "CollectionOptionsToolStripMenuItem"
        Me.CollectionOptionsToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.CollectionOptionsToolStripMenuItem.Text = "Collection Options"
        Me.CollectionOptionsToolStripMenuItem.ToolTipText = "You can edit Defaults here"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 100
        Me.ToolTip1.AutoPopDelay = 10000
        Me.ToolTip1.InitialDelay = 20
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 20
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(150, 126)
        Me.ToolStripContainer1.Location = New System.Drawing.Point(8, 8)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(150, 175)
        Me.ToolStripContainer1.TabIndex = 2
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.MenuStrip1)
        '
        'frmRandomNumbers
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1063, 624)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmRandomNumbers"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Random Numbers Management"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.lvRandom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ctxCopyReplace.ResumeLayout(False)
        CType(Me.udEditor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdSendRandom As System.Windows.Forms.Button
    Friend WithEvents lvRandom As System.Windows.Forms.DataGridView
    Friend WithEvents txtEditor As System.Windows.Forms.TextBox
    Friend WithEvents cmdCreateDocument As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents udEditor As System.Windows.Forms.NumericUpDown
    Friend WithEvents lvColumnGroups As ListViewEx.ListViewEx
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblTemplate As System.Windows.Forms.Label
    Friend WithEvents cmdEditDoc As System.Windows.Forms.Button
    Friend WithEvents cmdGenerateRandom As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents CustomerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CollectionOptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblFilename As System.Windows.Forms.Label
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents lblCollection As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ctxCopyReplace As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider

End Class
