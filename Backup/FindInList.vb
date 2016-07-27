'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-21 14:51:03+00 $
'$Log: FindInList.vb,v $
'Revision 1.0  2005-11-21 14:51:03+00  taylor
'Commented
'
Imports System.Windows.Forms
#Region "---"

#End Region
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : FindInList
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to help with finding an item in a list view
''' </summary>
''' <remarks>Can be used to trigger an external search if FirstColumnCMNumber is set
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-11-21 14:51:03+00 $</date><Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class FindInList
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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CBColumns As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FindInList))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CBColumns = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 125)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(440, 56)
        Me.Panel1.TabIndex = 4
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(352, 16)
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
        Me.cmdOk.Location = New System.Drawing.Point(272, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Find"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(64, 64)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "TextBox1"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(64, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Value to Find"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(232, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Column"
        '
        'CBColumns
        '
        Me.CBColumns.Location = New System.Drawing.Point(231, 64)
        Me.CBColumns.Name = "CBColumns"
        Me.CBColumns.Size = New System.Drawing.Size(121, 21)
        Me.CBColumns.TabIndex = 3
        '
        'FindInList
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 181)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CBColumns, Me.Label2, Me.Label1, Me.TextBox1, Me.Panel1})
        Me.Name = "FindInList"
        Me.Text = "Find Item In List View"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim LV As ListView
    Dim blnFirstCM As Boolean = False
    Dim tFind As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose find string
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FindString() As String
        Get
            Return tFind
        End Get
        Set(ByVal Value As String)
            tFind = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is first column primary key
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FirstColumnCMNumber() As Boolean
        Get
            Return Me.blnFirstCM
        End Get
        Set(ByVal Value As Boolean)
            blnFirstCM = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in list view
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property View() As ListView
        Get
            Return LV
        End Get
        Set(ByVal Value As ListView)
            LV = Value                              'List view to use
            Dim I As Integer
            Me.CBColumns.Items.Clear()              'Clear combo
            If Not LV Is Nothing Then               'If we have a valid list view
                For I = 0 To LV.Columns.Count - 1   'Go through all collumns in list view
                    Me.CBColumns.Items.Add(LV.Columns.Item(I).Text)     'Add column description to combo
                Next
                Me.CBColumns.SelectedIndex = 0      'Set first item selected
                Me.TextBox1.Text = ""
            End If
        End Set

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find item in list box
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim lvi As ListViewItem
        Dim lvis As ListViewItem.ListViewSubItem
        tFind = Me.TextBox1.Text.ToUpper
        Dim intI As Integer = -1

        For Each lvi In LV.Items
            If Me.CBColumns.SelectedIndex = 0 Then
                If InStr(lvi.Text.ToUpper, tFind) <> 0 Then
                    lvi.Selected = True
                    lvi.EnsureVisible()
                    intI = lvi.Index
                    Exit For
                End If
            Else
                lvis = lvi.SubItems(Me.CBColumns.SelectedIndex)
                If InStr(lvis.Text.ToUpper, tFind) <> 0 Then
                    lvi.Selected = True
                    lvi.EnsureVisible()
                    intI = lvi.Index
                    Exit For
                End If
            End If
        Next

        'if not found and the boolean find external set bog off
        If intI = -1 AndAlso Me.blnFirstCM Then
            Me.DialogResult = DialogResult.Retry
            Exit Sub
        End If

        If intI = -1 Then
            MsgBox("Failed to find the item you were after!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
        End If

    End Sub
End Class
