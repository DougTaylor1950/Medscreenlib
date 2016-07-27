'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-09 10:30:03+00 $
'$Log: frmSelectPort.vb,v $
'Revision 1.0  2005-11-09 10:30:03+00  taylor
'Checked in after minor editing
'
'Revision 1.1  2005-11-09 10:06:56+00  taylor
'Loging header added
'

Imports System.Windows.Forms
Imports MedscreenLib

'''<summary>
''' Form to provide a selection facility for a port
''' </summary>
''' <remarks>
''' This form is designed to provide a consistent approach to selecting ports.  
''' It is currently only used by the CCTool, but is designed to be useable from any application.
''' It may be used through the <see cref="frmSelectPort.Port"/> property, which exposes the port selected by the user.
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-11-09 10:30:03+00 $</date>
''' <Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmSelectPort
    Inherits System.Windows.Forms.Form
    Private RootNode As TreeNode
    Private UK As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode
    Private Overseas As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode
    Private ColumnDirection(10) As Integer
    Private blnDrawn As Boolean = False
    Private blnDrawing As Boolean = False
    Private oPort As Intranet.intranet.jobs.Port
    Private isoPortItem As ListViewItems.lvIsoportItem
    Private intCountryID As Integer = -1

    Public Enum PortSelectionType
        AllPorts
        PortSitesOnly
        NoFixedSite
    End Enum
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        Dim i As Integer
        For i = 0 To 10
            ColumnDirection(i) = 1
        Next
        Me.cbCountry1.DataSource = MedscreenLib.Glossary.Glossary.Countries
        Me.cbCountry1.DisplayMember = "CountryName"
        Me.cbCountry1.ValueMember = "CountryId"
        Me.cbCountry1.Hide()

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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents PortId As System.Windows.Forms.ColumnHeader
    Friend WithEvents PRegion As System.Windows.Forms.ColumnHeader
    Friend WithEvents Costs As System.Windows.Forms.ColumnHeader
    Friend WithEvents Desc As System.Windows.Forms.ColumnHeader
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cbCountry1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdRestore As System.Windows.Forms.Button
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents chUNLocode As System.Windows.Forms.ColumnHeader
    Friend WithEvents chIsoRegion As System.Windows.Forms.ColumnHeader
    Friend WithEvents ctxPortEdit As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuEditPort As System.Windows.Forms.MenuItem
    Friend WithEvents chCoordinates As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectPort))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.cmdRestore = New System.Windows.Forms.Button
        Me.cbCountry1 = New System.Windows.Forms.ComboBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.txtPort = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.cbCountry = New System.Windows.Forms.ComboBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.PortId = New System.Windows.Forms.ColumnHeader
        Me.PRegion = New System.Windows.Forms.ColumnHeader
        Me.Costs = New System.Windows.Forms.ColumnHeader
        Me.Desc = New System.Windows.Forms.ColumnHeader
        Me.chUNLocode = New System.Windows.Forms.ColumnHeader
        Me.chIsoRegion = New System.Windows.Forms.ColumnHeader
        Me.chCoordinates = New System.Windows.Forms.ColumnHeader
        Me.ctxPortEdit = New System.Windows.Forms.ContextMenu
        Me.mnuEditPort = New System.Windows.Forms.MenuItem
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdRemove)
        Me.Panel1.Controls.Add(Me.cmdRestore)
        Me.Panel1.Controls.Add(Me.cbCountry1)
        Me.Panel1.Controls.Add(Me.cmdAdd)
        Me.Panel1.Controls.Add(Me.cmdSearch)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.txtPort)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 381)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(624, 48)
        Me.Panel1.TabIndex = 0
        '
        'cmdRemove
        '
        Me.cmdRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRemove.Image = CType(resources.GetObject("cmdRemove.Image"), System.Drawing.Image)
        Me.cmdRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRemove.Location = New System.Drawing.Point(360, 16)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(72, 23)
        Me.cmdRemove.TabIndex = 12
        Me.cmdRemove.Text = "&Remove"
        Me.cmdRemove.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdRemove.Visible = False
        '
        'cmdRestore
        '
        Me.cmdRestore.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRestore.Image = CType(resources.GetObject("cmdRestore.Image"), System.Drawing.Image)
        Me.cmdRestore.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRestore.Location = New System.Drawing.Point(360, 16)
        Me.cmdRestore.Name = "cmdRestore"
        Me.cmdRestore.Size = New System.Drawing.Size(72, 23)
        Me.cmdRestore.TabIndex = 11
        Me.cmdRestore.Text = "&Restore"
        Me.cmdRestore.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdRestore.Visible = False
        '
        'cbCountry1
        '
        Me.cbCountry1.Location = New System.Drawing.Point(240, 16)
        Me.cbCountry1.Name = "cbCountry1"
        Me.cbCountry1.Size = New System.Drawing.Size(121, 21)
        Me.cbCountry1.TabIndex = 10
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.Image = CType(resources.GetObject("cmdAdd.Image"), System.Drawing.Image)
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(360, 16)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(72, 23)
        Me.cmdAdd.TabIndex = 9
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAdd.Visible = False
        '
        'cmdSearch
        '
        Me.cmdSearch.Image = CType(resources.GetObject("cmdSearch.Image"), System.Drawing.Image)
        Me.cmdSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSearch.Location = New System.Drawing.Point(156, 16)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(75, 23)
        Me.cmdSearch.TabIndex = 8
        Me.cmdSearch.Text = "Search"
        Me.cmdSearch.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(448, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "&Ok"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(528, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "C&ancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(44, 16)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(100, 20)
        Me.txtPort.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Port"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        Me.ImageList1.Images.SetKeyName(5, "")
        Me.ImageList1.Images.SetKeyName(6, "")
        Me.ImageList1.Images.SetKeyName(7, "")
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.CheckBox1)
        Me.Panel2.Controls.Add(Me.cbCountry)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(624, 40)
        Me.Panel2.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Country"
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(336, 8)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(104, 24)
        Me.CheckBox1.TabIndex = 19
        Me.CheckBox1.Text = "List View"
        '
        'cbCountry
        '
        Me.cbCountry.Items.AddRange(New Object() {"ENGLAND", "SCOTLAND", "IRELAND", "WALES"})
        Me.cbCountry.Location = New System.Drawing.Point(96, 8)
        Me.cbCountry.Name = "cbCountry"
        Me.cbCountry.Size = New System.Drawing.Size(176, 21)
        Me.cbCountry.TabIndex = 18
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.ListView1)
        Me.Panel3.Controls.Add(Me.TreeView1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 40)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(624, 208)
        Me.Panel3.TabIndex = 21
        '
        'ListView1
        '
        Me.ListView1.AllowColumnReorder = True
        Me.ListView1.BackColor = System.Drawing.Color.AliceBlue
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.PortId, Me.PRegion, Me.Costs, Me.Desc, Me.chUNLocode, Me.chIsoRegion, Me.chCoordinates})
        Me.ListView1.ContextMenu = Me.ctxPortEdit
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.LargeImageList = Me.ImageList1
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(624, 208)
        Me.ListView1.SmallImageList = Me.ImageList1
        Me.ListView1.TabIndex = 24
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'PortId
        '
        Me.PortId.Text = "Port Id"
        Me.PortId.Width = 100
        '
        'PRegion
        '
        Me.PRegion.Text = "Region"
        Me.PRegion.Width = 100
        '
        'Costs
        '
        Me.Costs.Text = "Costs"
        '
        'Desc
        '
        Me.Desc.Text = "Description"
        Me.Desc.Width = 150
        '
        'chUNLocode
        '
        Me.chUNLocode.Text = "UNLocode"
        Me.chUNLocode.Width = 80
        '
        'chIsoRegion
        '
        Me.chIsoRegion.Text = "ISO Region"
        '
        'chCoordinates
        '
        Me.chCoordinates.Text = "Coordinates"
        Me.chCoordinates.Width = 120
        '
        'ctxPortEdit
        '
        Me.ctxPortEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditPort})
        '
        'mnuEditPort
        '
        Me.mnuEditPort.Index = 0
        Me.mnuEditPort.Text = "&Edit Port"
        '
        'TreeView1
        '
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.HideSelection = False
        Me.TreeView1.ImageIndex = 1
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 0
        Me.TreeView1.Size = New System.Drawing.Size(624, 208)
        Me.TreeView1.Sorted = True
        Me.TreeView1.TabIndex = 23
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 248)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(624, 3)
        Me.Splitter1.TabIndex = 22
        Me.Splitter1.TabStop = False
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 251)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(624, 130)
        Me.HtmlEditor1.TabIndex = 23
        '
        'frmSelectPort
        '
        Me.AcceptButton = Me.cmdSearch
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(624, 429)
        Me.Controls.Add(Me.HtmlEditor1)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectPort"
        Me.Text = "Port Selection"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myFormMode As PortSelectionType = PortSelectionType.AllPorts

    Public Property FormMode() As PortSelectionType
        Get
            Return myFormMode
        End Get
        Set(ByVal Value As PortSelectionType)
            myFormMode = Value
        End Set
    End Property

    Private Sub SetupForm()
        'Add any initialization after the InitializeComponent() call
        blnDrawn = True
        Me.cbCountry.DataSource = Intranet.intranet.Support.cSupport.Ports.Countries
        Me.cbCountry.DisplayMember = "Info"
        Me.cbCountry.ValueMember = "PortId"

        Me.ListView1.Show()
        Me.TreeView1.Hide()
        RootNode = Me.TreeView1.Nodes.Add("Ports")
        Dim HeaderNode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode
        UK = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("United Kingdom")
        Me.RootNode.Nodes.Add(UK)
        HeaderNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("Overseas")
        Overseas = HeaderNode
        Me.RootNode.Nodes.Add(HeaderNode)

        HeaderNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("England")
        Dim EngNode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode = HeaderNode
        UK.Nodes.Add(HeaderNode)
        HeaderNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("Wales")
        Dim WalesNode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode = HeaderNode
        UK.Nodes.Add(HeaderNode)
        HeaderNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("Scotland")
        Dim ScotNode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode = HeaderNode
        UK.Nodes.Add(HeaderNode)
        HeaderNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode("Northern Ireland")
        Dim IreNode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode = HeaderNode
        UK.Nodes.Add(HeaderNode)

        'Me.TreeView1.BeginUpdate()
        Me.ListView1.BeginUpdate()
        Dim objPort As Intranet.intranet.jobs.Port = Intranet.intranet.Support.cSupport.Ports.Item("ENGLAND")
        'Me.AddChildren(EngNode, objPort)
        'EngNode.EnsureVisible()
        'Me.TreeView1.EndUpdate()

        Application.DoEvents()
        System.Threading.Thread.CurrentThread.Sleep(200)
        'Me.TreeView1.BeginUpdate()
        'objPort = Intranet.intranet.support.cSupport.Ports.Item("WALES")
        'Me.AddChildren(WalesNode, objPort)
        'WalesNode.EnsureVisible()

        'Me.TreeView1.EndUpdate()
        'Application.DoEvents()
        'System.Threading.Thread.CurrentThread.Sleep(200)
        'Me.TreeView1.BeginUpdate()
        'objPort = Intranet.intranet.support.cSupport.Ports.Item("SCOTLAND")
        'Me.TreeView1.EndUpdate()
        'ScotNode.EnsureVisible()
        'Application.DoEvents()
        'System.Threading.Thread.CurrentThread.Sleep(200)
        'Me.TreeView1.BeginUpdate()
        'Me.AddChildren(ScotNode, objPort)
        'objPort = Intranet.intranet.support.cSupport.Ports.Item("NORTHERN_IRELAND")
        'Me.TreeView1.EndUpdate()
        'IreNode.EnsureVisible()
        'Application.DoEvents()
        'System.Threading.Thread.CurrentThread.Sleep(200)
        'Me.TreeView1.BeginUpdate()
        'Me.AddChildren(IreNode, objPort)

        Dim i As Integer
        Dim oCnt As Intranet.intranet.jobs.Port
        Dim TT As TreeNode
        'For i = 0 To Intranet.intranet.support.cSupport.Ports.Continents.Count - 1
        '    oCnt = Intranet.intranet.support.cSupport.Ports.Continents.Item(i)
        '    If InStr("ENGLAND,WALES,SCOTLAND,NORTHEN_IRELAND", oCnt.PortId) = 0 Then
        '        TT = New TreeNode(oCnt.Info)
        '        TT.Tag = oCnt
        '        Overseas.Nodes.Add(TT)
        '        Me.TreeView1.EndUpdate()
        '        TT.EnsureVisible()
        '        Application.DoEvents()
        '        System.Threading.Thread.CurrentThread.Sleep(200)
        '        Me.TreeView1.BeginUpdate()
        '        Me.AddChildren(TT, oCnt)
        '    End If
        'Next
        'Me.TreeView1.EndUpdate()
        Me.ListView1.EndUpdate()

    End Sub

    'Private Sub AddChildren(ByVal aNode As TreeNode, ByVal aPort As Intranet.intranet.jobs.Port)
    '    Dim cCol As Intranet.intranet.jobs.CentreCollection

    '    If aPort Is Nothing Then Exit Sub

    '    cCol = Intranet.intranet.support.cSupport.Ports.Children(aPort)
    '    Dim i As Integer
    '    For i = 0 To cCol.Count - 1
    '        Dim objPort As Intranet.intranet.jobs.Port
    '        objPort = cCol.Item(i)
    '        Dim rNode As New PortTreeNode(objPort)
    '        'rNode.Tag = objPort
    '        If objPort.PortType = "P" Or objPort.PortType = "S" Then
    '            Dim lvItem As lvPortItem = New lvPortItem(objPort)
    '            lvItem.ImageIndex = 0
    '            Me.ListView1.Items.Add(lvItem)
    '        End If
    '        aNode.Nodes.Add(rNode)
    '        'If objPort.PortId = myVessel.VesselSubType Then
    '        '    Me.TreeView1.SelectedNode = rNode
    '        '    rNode.EnsureVisible()
    '        'End If
    '        AddChildren(rNode, objPort)
    '    Next
    'End Sub


    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Me.ListView1.Show()
            Me.TreeView1.Hide()
            Me.cbCountry.Hide()
            'Me.txtPort.Show()
        Else
            Me.TreeView1.Show()
            Me.ListView1.Hide()
            Me.cbCountry.Show()
            'Me.txtPort.Hide()
        End If
    End Sub

    Public Property CountryID() As Integer
        Get
            Return Me.intCountryID
        End Get
        Set(ByVal Value As Integer)
            intCountryID = Value
            If intCountryID > 0 Then

                Dim objCountry As MedscreenLib.Address.Country = MedscreenLib.Glossary.Glossary.Countries.Item(Value, 0)
                If Not objCountry Is Nothing Then
                    Me.cbCountry1.SelectedIndex = MedscreenLib.Glossary.Glossary.Countries.IndexOf(objCountry)
                End If
            Else
                Me.cbCountry1.SelectedItem = -1
            End If

        End Set
    End Property

    Private Sub cbCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCountry.SelectedIndexChanged
        Dim objPort As Intranet.intranet.jobs.Port

        If Not blnDrawn Then Exit Sub
        If Not blnDrawing Then Exit Sub
        objPort = Intranet.intranet.Support.cSupport.Ports.Item(cbCountry.SelectedValue)

        If objPort Is Nothing Then Exit Sub
        If Not Overseas.HasNode(objPort.PortId) Then
            Dim ptNode As MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode = New MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode(objPort)
            Me.Overseas.Nodes.Add(ptNode)
        End If
        'FindPort(Overseas, objPort)

    End Sub

    'Private Sub FindPort(ByVal aNode As TreeNode, ByVal aPort As Intranet.intranet.jobs.Port)
    '    Dim trNode As TreeNode
    '    For Each trNode In aNode.Nodes
    '        If TypeOf trNode Is PortTreeNode Then
    '            Dim ptNode As PortTreeNode
    '            If CType(trNode.Tag, Intranet.intranet.jobs.Port).PortId = aPort.PortId Then
    '                Me.TreeView1.SelectedNode = trNode
    '                trNode.EnsureVisible()
    '                Exit For
    '            Else
    '                FindPort(trNode, aPort)
    '            End If
    '        Else
    '            FindPort(trNode, aPort)
    '        End If
    '    Next
    'End Sub


    Private Sub ListView1_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        Dim intdirection As Integer

        intdirection = Me.ColumnDirection(e.Column)
        Me.ColumnDirection(e.Column) *= -1

        Me.ListView1.ListViewItemSorter = New Comparers.ListViewItemComparer(e.Column, intdirection)

    End Sub

    'Private Sub txtPort_TextChanged()
    '    Dim j As ListViewItem
    '    For Each j In Me.ListView1.Items
    '        If InStr(j.Text, Me.txtPort.Text.ToUpper) = 1 Then
    '            j.Selected = True
    '            j.EnsureVisible()
    '            Me.ListView1.EnsureVisible(j.Index)
    '            Me.ListView1.Invalidate()
    '            'DataGrid1_Click(Nothing, Nothing)
    '            Application.DoEvents()
    '            Exit For
    '        End If
    '    Next
    '    Me.TreeView1.CollapseAll()
    '    Me.RootNode.Expand()
    '    FindPort(RootNode, j.Tag)

    'End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim tNode As TreeNode = Me.TreeView1.SelectedNode
        If Not TypeOf tNode Is MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode Then Exit Sub

        'Dim oPort As Intranet.intranet.jobs.Port
        Dim tpNode As MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode
        tpNode = tNode
        oPort = tpNode.Port
        Me.HtmlEditor1.DocumentText = (oPort.ToHtml("port1.xsl"))
        If tpNode.Nodes.Count = 0 And oPort.Children.Count > 0 Then
            Dim oPort1 As Intranet.intranet.jobs.Port
            For Each oPort1 In oPort.Children
                Dim tpNode1 As MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode
                tpNode1 = New MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode(oPort1)
                tpNode.Nodes.Add(tpNode1)
            Next
        End If
    End Sub

    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        isoPortItem = Nothing
        Me.cbCountry1.Hide()
        Me.cmdRestore.Hide()
        Me.cmdRemove.Hide()
        Me.ListView1.ContextMenu = Nothing
        If Me.ListView1.SelectedItems.Count = 1 Then
            If TypeOf Me.ListView1.SelectedItems.Item(0) Is ListViewItems.lvPortItem Then
                Dim lvItem As ListViewItems.lvPortItem = Me.ListView1.SelectedItems.Item(0)
                If Not lvItem.Port Is Nothing Then
                    'Dim oPort As Intranet.intranet.jobs.Port
                    oPort = lvItem.Port
                    Me.HtmlEditor1.DocumentText = (oPort.ToHtml("port1.xsl"))
                    If oPort.Removed Then
                        cmdRestore.Show() 'Enable users to restore a port
                    Else
                        cmdRemove.Show()
                    End If
                End If
                Me.ListView1.ContextMenu = Me.ctxPortEdit
                Me.cmdAdd.Hide()
            Else    ' not a port item, will we allow a create
                isoPortItem = Me.ListView1.SelectedItems.Item(0)
                Me.cbCountry1.Visible = (isoPortItem.ID = "-----")
                Me.cmdAdd.Show()
                Me.HtmlEditor1.DocumentText = ("<HTML><BODY><DIV>You can create this port by clicking the add button</DIV></BODY></HTML>")
            End If

        End If
    End Sub

    Private Sub frmSelectPort_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If Not blnDrawn Then
            Me.SetupForm()
            blnDrawn = True
            blnDrawing = True
        End If
    End Sub

    Public Sub DoFindPort()
        Me.cmdSearch_Click(Nothing, Nothing)
    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        If Me.txtPort.Text.Trim.Length = 0 Then Exit Sub 'Check we have something to look for 

        '   Dim strQuery As String = "Select identity from port where upper(identity) like '%" & Me.txtPort.Text.Trim.ToUpper & _
        '       "%' or upper(description) like  '%" & Me.txtPort.Text.Trim.ToUpper & _
        '      "%' or iso_id = '" & Me.txtPort.Text.Trim.ToUpper & "' or upper(site_name) like '%" & _
        '      Me.txtPort.Text.Trim.ToUpper & "%'"

        Dim oCmd As New OleDb.OleDbCommand("", MedscreenLib.CConnection.DbConnection)
        Me.cmdRemove.Hide()
        Me.cmdAdd.Show()
        Me.cmdAdd.BringToFront()
        Me.cmdRestore.Hide()
        'Dim oRead As OleDb.OleDbDataReader
        Dim aColl As String()
        Dim pColl As String()
        Me.ListView1.Items.Clear()
        If Not Me.Overseas Is Nothing Then Me.Overseas.Nodes.Clear()

        Try
            'If MedscreenLib.CConnection.ConnOpen Then
            '    oRead = oCmd.ExecuteReader
            '    While oRead.Read
            '        aColl.Add(oRead.GetValue(0))
            '    End While
            'End If

            ''Switch to looking at Iso Ports
            'If Not oRead Is Nothing Then
            '    If Not oRead.IsClosed Then oRead.Close()
            'End If
            Dim Parameters As New Collection()
            Parameters.Add(Me.txtPort.Text)
            If Me.myFormMode = PortSelectionType.PortSitesOnly Or Me.myFormMode = PortSelectionType.NoFixedSite Then
                Parameters.Add("T")
            End If
            Dim PortIDString As String = MedscreenLib.CConnection.PackageStringList("lib_cctool.PortsById", Parameters)
            aColl = PortIDString.Split(New Char() {","})
            Dim IsoPortIDString As String = MedscreenLib.CConnection.PackageStringList("lib_cctool.ISOPortsById", Me.txtPort.Text)
            pColl = IsoPortIDString.Split(New Char() {","})


            'oCmd.CommandText = "Select * from isoports where  upper(port_name) like upper('%" & Me.txtPort.Text.Trim.ToUpper & "%')"
            'oRead = oCmd.ExecuteReader
            'While oRead.Read
            '    pColl.Add(oRead.GetValue(0) & ";" & oRead.GetValue(1))
            'End While

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        Finally
            MedscreenLib.CConnection.SetConnClosed()
            'If Not oRead Is Nothing Then
            '    If Not oRead.IsClosed Then oRead.Close()
            'End If
        End Try

        Try
            Intranet.intranet.Support.cSupport.Ports.RefreshChanged()
            Dim i As Integer
            For i = 0 To aColl.Length - 1
                Dim strPortId As String = aColl.GetValue(i)
                'check we got one 
                If strPortId.Trim.Length > 0 Then

                    Dim aPort As Intranet.intranet.jobs.Port = Intranet.intranet.Support.cSupport.Ports.Item(strPortId)
                    If Not aPort Is Nothing Then
                        Try
                            Dim anItem As ListViewItems.lvPortItem = New ListViewItems.lvPortItem(aPort)
                            Me.ListView1.Items.Add(anItem)
                        Catch ex As Exception
                            Medscreen.LogError(ex, , "Selecting port")
                        End Try
                    End If
                End If
            Next

            For i = 0 To pColl.Length - 1
                Dim tmpstr As String = pColl.GetValue(i)
                'See if we have a valid string 
                If tmpstr.Trim.Length > 0 Then
                    Dim items As String() = tmpstr.Split(New Char() {";"})
                    'Split again to get description from string 
                    Dim aCoPort As Intranet.intranet.jobs.Port = Intranet.intranet.Support.cSupport.Ports.Item(items(0), Intranet.intranet.jobs.CentreCollection.GetPortBy.ISOID)
                    If aCoPort Is Nothing Then
                        Dim anitem As ListViewItems.lvIsoportItem = New ListViewItems.lvIsoportItem(items(0), items(1))
                        Me.ListView1.Items.Add(anItem)
                    End If
                End If
            Next
            Me.ListView1.Items.Add(New ListViewItems.lvIsoportItem("-----", "Add Site by name"))
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "Frm Select")
        End Try

    End Sub

    Private Function AddSpecialTreeNode(ByVal port As String, ByVal aPnode As TreeNode) As TreeNode
        Dim aCoPort As Intranet.intranet.jobs.Port = Intranet.intranet.Support.cSupport.Ports.Item(port)
        If Not aCoPort Is Nothing Then
            If TypeOf aPnode Is MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode Then

                Dim axnode As MedscreenCommonGui.Treenodes.LocationNodes.PortHeaderNode = aPnode
                aPnode = Nothing
                If axnode.HasNode(port) Then
                    aPnode = axnode.FindNode(port)
                Else
                    aPnode = New MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode(aCoPort)
                    axnode.nodes.add(aPnode)
                End If
            ElseIf TypeOf aPnode Is MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode Then
                Dim axnode As MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode = aPnode
                aPnode = Nothing
                If axnode.HasNode(port) Then
                    aPnode = axnode.FindNode(port)
                Else
                    aPnode = New MedscreenCommonGui.Treenodes.LocationNodes.PortTreeNode(aCoPort)
                    axnode.nodes.add(aPnode)
                End If
            End If

        End If
        Return aPnode
    End Function

    '''<summary>
    ''' Port which the user has selected <see cref="Intranet.intranet.jobs.Port" />
    ''' </summary>
    Public Property Port() As Intranet.intranet.jobs.Port
        Get
            Return oPort
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.Port)
            oPort = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the list view to other users
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Listview() As Boolean
        Set(ByVal Value As Boolean)
            Me.CheckBox1.Checked = Value
            'CheckBox1_CheckedChanged(Nothing, Nothing)
            If CheckBox1.Checked Then
                Me.ListView1.Show()
                Me.TreeView1.Hide()
                Me.cbCountry.Hide()
                'Me.txtPort.Show()
            Else
                Me.TreeView1.Show()
                Me.ListView1.Hide()
                Me.cbCountry.Show()
                'Me.txtPort.Hide()
            End If

        End Set
    End Property




    Private Function GetIsoCountryCode(ByVal ISoCode As String) As Integer
        Dim intReturn As Integer = -1
        Dim objReturn As Object

        Dim oCmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("select portfunction from isoports where iso_id ='" & ISoCode & "'", CConnection.DbConnection)
        Try
            If CConnection.ConnOpen Then
                objReturn = oCmd.ExecuteScalar

            End If
            If Not objReturn Is Nothing Then intReturn = CInt(objReturn)
        Catch ex As Exception

        Finally
            CConnection.SetConnClosed()

        End Try
        Return intReturn
    End Function


    Private Function GetIsoContinent(ByVal ISoCode As String) As String
        Dim strReturn As String = """"
        Dim objReturn As Object

        Dim oCmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("select COORDINATES from isoports where iso_id ='" & ISoCode & "'", CConnection.DbConnection)
        Try
            If CConnection.ConnOpen Then
                objReturn = oCmd.ExecuteScalar

            End If
            If Not objReturn Is Nothing Then strReturn = CStr(objReturn)
        Catch ex As Exception

        Finally
            CConnection.SetConnClosed()

        End Try
        Return strReturn
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Dim aPort As Intranet.intranet.jobs.Port
        'create a new port
        Me.Cursor = Cursors.WaitCursor
        MedscreenLib.Glossary.Glossary.Countries.Clear()
        MedscreenLib.Glossary.Glossary.Countries.Load()
        Me.cbCountry.DataSource = MedscreenLib.Glossary.Glossary.Countries
        If isoPortItem Is Nothing Then Exit Sub
        Me.cmdAdd.Enabled = False
        If isoPortItem.ID = "-----" Then            'Not found in the ISO Ports list
            'Okay get new ID
            If Me.txtPort.Text.Trim.Length = 0 Then
                MsgBox("Must have a name for the new port!", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                Exit Sub
            End If
            'Check is not a number
            If Mid(Me.txtPort.Text, 1, 5) = "00000" Then
                MsgBox("Must have a alphabetical name for the new port!", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                Exit Sub

            End If
            If Me.cbCountry1.SelectedValue = -1 Then    'Check we have a country
                MsgBox("please select a country")       'Force them to select a country
                Me.cmdAdd.Enabled = True                'Reenable add
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            aPort = New Intranet.intranet.jobs.Port(MedscreenLib.CConnection.NextSequence("SEQPORTID"), Intranet.intranet.Support.cSupport.Ports)
            aPort.PortType = "S"        'Set to be a site
            aPort.CountryId = Me.cbCountry1.SelectedValue
            aPort.SiteName = Me.txtPort.Text        'Set name 
            aPort.PortDescription = Me.txtPort.Text
            aPort.IncursExtracharge = False
            aPort.Update()
        Else
            'We have an ISO port get new ID
            aPort = New Intranet.intranet.jobs.Port(isoPortItem.ID, Intranet.intranet.Support.cSupport.Ports)
            aPort.UNLocode = isoPortItem.ID     'Add Loccode
            If isoPortItem.CountryId = -1 Then
                'We need to create a country 
                Dim strCountry As String = GetCountry(isoPortItem.ID)
                Dim objCountryId As Object = Nothing
                Dim strContinent As String = ""
                'Ask the user for the international dialling code 
                If GetIsoCountryCode(Mid$(isoPortItem.ID, 1, 2)) = -1 Then
                    MsgBox("We need to find the international dialling code for the country in question, (" & _
                    strCountry & ") please use CountryCode Tool.", MsgBoxStyle.Question)
                    objCountryId = Medscreen.GetParameter(Medscreen.MyTypes.typeInteger, "Country's Dialling Code", "Please enter Countries dialling code")
                Else
                    objCountryId = GetIsoCountryCode(Mid$(isoPortItem.ID, 1, 2))
                    strContinent = GetIsoContinent(Mid$(isoPortItem.ID, 1, 2))
                End If
                If Not objCountryId Is Nothing Then
                    Dim objCountry As New MedscreenLib.Address.Country()
                    objCountry.CountryId = CInt(objCountryId)
                    objCountry.CountryName = strCountry
                    objCountry.CountryCode = Mid(isoPortItem.ID, 1, 2)
                    objCountry.ContinentalArea = strContinent
                    If Not objCountry.Insert() Then
                        Throw New System.Exception("Can not create Country")
                        Exit Sub
                    End If
                    MedscreenLib.Glossary.Glossary.Countries.Add(objCountry)
                    Me.cbCountry.DataSource = MedscreenLib.Glossary.Glossary.Countries
                    'We also need to create the port entity if id is longer than 2 charcaters
                    If isoPortItem.ID.Trim.Length > 2 Then
                        Dim aCntPort As Intranet.intranet.jobs.Port = New Intranet.intranet.jobs.Port(objCountry.CountryCode, Intranet.intranet.Support.cSupport.Ports)
                        aCntPort.CountryId = objCountry.CountryId
                        aCntPort.SiteName = strCountry
                        aCntPort.PortDescription = strCountry
                        aCntPort.PortType = "C"
                        aCntPort.UNLocode = objCountry.CountryCode
                        aCntPort.PortRegion = strContinent
                        aCntPort.Update()
                        aPort.PortRegion = aCntPort.Identity
                    End If
                    isoPortItem.Refresh()
                End If
            End If
            aPort.CountryId = isoPortItem.CountryId 'Add country
            If isoPortItem.ID.Trim.Length = 2 Then
                aPort.PortType = "C"
            Else
                aPort.PortType = "P"                'Is a port
            End If
            aPort.SiteName = isoPortItem.Name   'Set name to ISO name
            aPort.PortDescription = aPort.SiteName
        End If
        If aPort.CountryId = 44 Then aPort.IsUK = True
        aPort.Update()
        Dim oCMd As New OleDb.OleDbCommand("SElect identity from port where type = 'C' and country_id = " & _
            aPort.CountryId, MedscreenLib.CConnection.DbConnection)
        If aPort.PortType = "C" Then
            oCMd.CommandText = "SElect CONTINENTAL_AREA from country C where country_id =  " & aPort.CountryId
        End If
        Dim strRegion As String
        Try
            If MedscreenLib.CConnection.ConnOpen Then
                strRegion = oCMd.ExecuteScalar
            End If
        Catch ex As Exception
        Finally
            MedscreenLib.CConnection.SetConnClosed()
        End Try
        aPort.PortRegion = strRegion
        aPort.IsUK = (aPort.CountryId = 44)
        aPort.Created = Now
        aPort.IncursExtracharge = True
        aPort.Update()
        Intranet.intranet.Support.cSupport.Ports.Add(aPort)
        oPort = aPort
        Me.cmdAdd.Enabled = True
        Me.cmdSearch_Click(Nothing, Nothing)
        Me.Cursor = Cursors.Default


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a country name from the port code, countries are stored in the ISOPORTs table as two characater codes
    ''' </summary>
    ''' <param name="PortCode">5 character port code</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function GetCountry(ByVal PortCode As String) As String
        Dim ocmd As New OleDb.OleDbCommand("Select port_name from isoports where iso_id = ?", CConnection.DbConnection)
        Dim objReturn As Object
        Dim strReturn As String = ""
        Try
            If CConnection.ConnOpen Then            'Attempt to open reader
                ocmd.Parameters.Add(CConnection.StringParameter("PortID", Mid(PortCode, 1, 2), 5))
                objReturn = ocmd.ExecuteScalar
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , )
        Finally
            CConnection.SetConnClosed()             'Close connection

        End Try
        If Not objReturn Is Nothing AndAlso Not objReturn Is DBNull.Value Then
            strReturn = Mid(CStr(objReturn), 2)
            strReturn = Medscreen.Capitalise(strReturn)
        End If
        Return strReturn

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Restore a port and collecting officers
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdRestore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRestore.Click
        If oPort.RestorePort Then
            Dim lvItem As ListViewItems.lvPortItem = Me.ListView1.SelectedItems.Item(0)
            lvItem.Refresh()
            Me.cmdRestore.Hide()
        End If
    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        If oPort.RemovePort Then
            Dim lvItem As ListViewItems.lvPortItem = Me.ListView1.SelectedItems.Item(0)
            lvItem.Refresh()
            Me.cmdRemove.Hide()
        End If

    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        Dim li As ListViewItem
        If Me.ListView1.SelectedItems.Count = 1 Then
            li = ListView1.SelectedItems.Item(0)
            If TypeOf li Is ListViewItems.lvPortItem Then
                Dim lip As ListViewItems.lvPortItem = li
                EditPort(lip.Port)
                lip.Refresh()
                Me.ListView1_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub EditPort(ByVal aPort As Intranet.intranet.jobs.Port)
        Dim afrm As New frmPort()
        afrm.Port = aPort
        If afrm.ShowDialog = DialogResult.OK Then
            afrm.Port.Update()
        End If
    End Sub

    Private Sub mnuEditPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditPort.Click
        Dim li As ListViewItem
        If Me.ListView1.SelectedItems.Count = 1 Then
            li = ListView1.SelectedItems.Item(0)
            If TypeOf li Is ListViewItems.lvPortItem Then
                Dim lip As ListViewItems.lvPortItem = li
                EditPort(lip.Port)
                lip.Refresh()
                Me.ListView1_Click(Nothing, Nothing)
            End If
        End If

    End Sub

    'Private Sub txtLocation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim strLoc As String = Me.txtLocation.Text
    '    Try
    '        If blnLocationFound Then
    '            If Now.Subtract(LastFind).TotalSeconds < 5 Then
    '                Exit Sub
    '            End If
    '        End If
    '        If strLoc.Trim.Length > 0 Then
    '            Dim strReturn As String = CConnection.PackageStringList("LIB_CCTOOL.PORTSBYID", strLoc)
    '            If strReturn.Trim.Length > 0 Then 'Found at least one port 
    '                Dim strPorts As String() = strReturn.Split(New Char() {","})
    '                If strPorts.Length = 2 Then 'Found just one port 

    '                    '
    '                    '  disable txtLocation
    '                    '
    '                    blnLocationFound = True
    '                    LastFind = Now
    '                    Dim strPortId As String = strPorts.GetValue(0)
    '                    Dim myPort As New Intranet.intranet.jobs.Port(strPortId)
    '                    If Not myPort Is Nothing Then
    '                        Me.txtLocation.Text = myPort.Info2
    '                        Me.strportid = myPort.Identity
    '                        'Display info about ports and officers
    '                        'Me.OfficerPortInfo()
    '                        'Me.AssignOfficer()
    '                        Me.ErrorProvider1.SetError(Me.txtLocation, "")
    '                    End If
    '                    Me.txtLocation.Enabled = True
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub


End Class



