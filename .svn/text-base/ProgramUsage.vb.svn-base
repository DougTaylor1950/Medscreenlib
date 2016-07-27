'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-10 08:11:49+00 $
'$Log: ProgramUsage.vb,v $
'Revision 1.0  2005-11-10 08:11:49+00  taylor
'Finished documenting
'
'Revision 1.0  2005-11-09 10:30:03+00  taylor
'Checked in after minor editing
'
'Revision 1.1  2005-11-09 10:06:56+00  taylor
'Loging header added
'
Imports MedscreenLib

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : ProgramUsage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Shows users what programs they have running according to the program usage table
''' and allows them to log themseleves out
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author>
''' <date> $Date: 2005-11-10 08:11:49+00 $</date><Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class ProgramUsage
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
    Friend WithEvents lvProgramUsage As System.Windows.Forms.ListView
    Friend WithEvents chProgram As System.Windows.Forms.ColumnHeader
    Friend WithEvents chServer As System.Windows.Forms.ColumnHeader
    Friend WithEvents chLastLogin As System.Windows.Forms.ColumnHeader
    Friend WithEvents chStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents ctxProgram As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuItemLogout As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ProgramUsage))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.lvProgramUsage = New System.Windows.Forms.ListView()
        Me.chProgram = New System.Windows.Forms.ColumnHeader()
        Me.chServer = New System.Windows.Forms.ColumnHeader()
        Me.chLastLogin = New System.Windows.Forms.ColumnHeader()
        Me.chStatus = New System.Windows.Forms.ColumnHeader()
        Me.ctxProgram = New System.Windows.Forms.ContextMenu()
        Me.mnuItemLogout = New System.Windows.Forms.MenuItem()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 241)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(512, 32)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(432, 4)
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
        Me.cmdOk.Location = New System.Drawing.Point(352, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvProgramUsage
        '
        Me.lvProgramUsage.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chProgram, Me.chServer, Me.chLastLogin, Me.chStatus})
        Me.lvProgramUsage.ContextMenu = Me.ctxProgram
        Me.lvProgramUsage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvProgramUsage.FullRowSelect = True
        Me.lvProgramUsage.Name = "lvProgramUsage"
        Me.lvProgramUsage.Size = New System.Drawing.Size(512, 241)
        Me.lvProgramUsage.TabIndex = 2
        Me.lvProgramUsage.View = System.Windows.Forms.View.Details
        '
        'chProgram
        '
        Me.chProgram.Text = "Program"
        Me.chProgram.Width = 120
        '
        'chServer
        '
        Me.chServer.Text = "Server"
        Me.chServer.Width = 120
        '
        'chLastLogin
        '
        Me.chLastLogin.Text = "Last Logged in"
        Me.chLastLogin.Width = 180
        '
        'chStatus
        '
        Me.chStatus.Text = "Status"
        Me.chStatus.Width = 80
        '
        'ctxProgram
        '
        Me.ctxProgram.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemLogout})
        '
        'mnuItemLogout
        '
        Me.mnuItemLogout.Index = 0
        Me.mnuItemLogout.Text = "&Logout "
        '
        'ProgramUsage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(512, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lvProgramUsage, Me.Panel1})
        Me.Name = "ProgramUsage"
        Me.Text = "ProgramUsage"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myUser As MedscreenLib.User
    Private myProgramUsage As MedscreenLib.ProgramUsageList
    Private myProgram As MedscreenLib.ProgramUsage
    Private aLvItem As ListViewItems.lvProgUsageItem

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Supply user these information are for.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property User() As MedscreenLib.User
        Set(ByVal Value As MedscreenLib.User)
            If Not Value Is Nothing Then
                myUser = Value
                Me.Text = "Programs in use for : " & myUser.FullName
                myProgramUsage = New MedscreenLib.ProgramUsageList()
                If myProgramUsage.LoadByUser(myUser.UserID) Then
                    Dim oPU As MedscreenLib.ProgramUsage
                    For Each oPU In myProgramUsage
                        Dim oLVI As ListViewItems.lvProgUsageItem = New ListViewItems.lvProgUsageItem(oPU)
                        Me.lvProgramUsage.Items.Add(oLVI)
                    Next
                End If


                'myUser.UserLoginInfo.GetUsersInProgram()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Will fire when a list view item row is clicked 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub lvProgramUsage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvProgramUsage.Click
        Me.myProgram = Nothing
        Me.aLvItem = Nothing
        If Me.lvProgramUsage.SelectedItems.Count = 1 Then
            aLvItem = Me.lvProgramUsage.SelectedItems.Item(0)
            myProgram = aLvItem.ProgramUsage
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Will fire when the logout menu is clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuItemLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemLogout.Click
        If myProgram Is Nothing Then Exit Sub
        If myProgram.LoggedIn Then
            myProgram.LoggedIn = False
            myProgram.Update()
            aLvItem.Refresh()
        End If

    End Sub
End Class
