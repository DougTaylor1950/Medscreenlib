'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-09 10:08:58+00 $
'$Log: frmMessages.vb,v $
'Revision 1.0  2005-11-09 10:08:58+00  taylor
'Checked in after initial development
'
Imports MedscreenLib.Messaging.Messaging

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : frmMessages
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A form that allows users to see internal messages 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-11-09 10:08:58+00 $</date><Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmMessages
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
    Friend WithEvents lvMessages As System.Windows.Forms.ListView
    Friend WithEvents CHId As System.Windows.Forms.ColumnHeader
    Friend WithEvents chFrom As System.Windows.Forms.ColumnHeader
    Friend WithEvents chRead As System.Windows.Forms.ColumnHeader
    Friend WithEvents chExpires As System.Windows.Forms.ColumnHeader
    Friend WithEvents chSent As System.Windows.Forms.ColumnHeader
    Friend WithEvents chSize As System.Windows.Forms.ColumnHeader
    Friend WithEvents nwMessage As VbPowerPack.NotificationWindow
    Friend WithEvents ctxPopupMessages As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuMarkRead As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRead As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMessages))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.lvMessages = New System.Windows.Forms.ListView()
        Me.CHId = New System.Windows.Forms.ColumnHeader()
        Me.chFrom = New System.Windows.Forms.ColumnHeader()
        Me.chRead = New System.Windows.Forms.ColumnHeader()
        Me.chExpires = New System.Windows.Forms.ColumnHeader()
        Me.chSent = New System.Windows.Forms.ColumnHeader()
        Me.chSize = New System.Windows.Forms.ColumnHeader()
        Me.nwMessage = New VbPowerPack.NotificationWindow(Me.components)
        Me.ctxPopupMessages = New System.Windows.Forms.ContextMenu()
        Me.mnuMarkRead = New System.Windows.Forms.MenuItem()
        Me.mnuRead = New System.Windows.Forms.MenuItem()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 241)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(520, 32)
        Me.Panel1.TabIndex = 2
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(440, 4)
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
        Me.cmdOk.Location = New System.Drawing.Point(360, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvMessages
        '
        Me.lvMessages.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CHId, Me.chFrom, Me.chRead, Me.chExpires, Me.chSent, Me.chSize})
        Me.lvMessages.ContextMenu = Me.ctxPopupMessages
        Me.lvMessages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvMessages.Name = "lvMessages"
        Me.lvMessages.Size = New System.Drawing.Size(520, 241)
        Me.lvMessages.TabIndex = 3
        Me.lvMessages.View = System.Windows.Forms.View.Details
        '
        'CHId
        '
        Me.CHId.Text = "Id"
        '
        'chFrom
        '
        Me.chFrom.Text = "From"
        Me.chFrom.Width = 100
        '
        'chRead
        '
        Me.chRead.Text = "Read"
        '
        'chExpires
        '
        Me.chExpires.Text = "Expires"
        Me.chExpires.Width = 100
        '
        'chSent
        '
        Me.chSent.Text = "Sent On"
        Me.chSent.Width = 80
        '
        'chSize
        '
        Me.chSize.Text = "Size"
        '
        'nwMessage
        '
        Me.nwMessage.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.SystemColors.InactiveCaption, System.Drawing.SystemColors.Window)
        Me.nwMessage.DefaultText = Nothing
        Me.nwMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nwMessage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.nwMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.nwMessage.ShowStyle = VbPowerPack.NotificationShowStyle.Slide
        '
        'ctxPopupMessages
        '
        Me.ctxPopupMessages.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuMarkRead, Me.mnuRead})
        '
        'mnuMarkRead
        '
        Me.mnuMarkRead.Index = 0
        Me.mnuMarkRead.Text = "&Mark Read"
        '
        'mnuRead
        '
        Me.mnuRead.Index = 1
        Me.mnuRead.Text = "&Read"
        '
        'frmMessages
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lvMessages, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMessages"
        Me.Text = "Messages"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myMessages As MedscreenLib.Messaging.MessageList
    Private lvItem As ListViewItems.lvMessageItem

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes the messages passed to and from this class 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>Set method will populate the display 
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Messages() As MedscreenLib.Messaging.MessageList
        Get
            Return myMessages
        End Get
        Set(ByVal Value As MedscreenLib.Messaging.MessageList)
            myMessages = Value
            Me.lvMessages.Items.Clear()
            Dim objMessage As MedscreenLib.Messaging.Messaging
            For Each objMessage In myMessages
                Me.lvMessages.Items.Add(New ListViewItems.lvMessageItem(objMessage))
            Next
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Select the item in the list view 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lvMessages_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMessages.Click
        If Me.lvMessages.SelectedItems.Count = 1 Then
            lvItem = Me.lvMessages.SelectedItems.Item(0)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Action the message, either by using a notify box or html viewer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lvMessages_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMessages.DoubleClick, mnuRead.Click
        If Not lvItem Is Nothing Then
            'Deal with visible message types 'Attachments may be visible
            If lvItem.Message.MessageType = MedscreenLib.Messaging.Messaging.GCST_MessageTypeText Then
                Me.nwMessage.Notify(lvItem.Message.MessageText)
            ElseIf lvItem.Message.MessageType = MedscreenLib.Messaging.Messaging.GCST_MessageTypeHTML Then
                MedscreenLib.Medscreen.ShowHtml(lvItem.Message.MessageText)
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the message read 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuMarkRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMarkRead.Click
        If Me.lvItem Is Nothing Then Exit Sub

        Me.lvItem.Message.Read = True
        Me.lvItem.Message.Update()
    End Sub


End Class

