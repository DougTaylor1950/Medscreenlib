'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 07:31:33+00 $
'$Log: FrmCollOfficer.vb,v $
'Revision 1.0  2005-11-29 07:31:33+00  taylor
'Commented
'
'Revision 1.2  2005-06-16 17:38:56+01  taylor
'Milestone Check in
'
'Revision 1.1  2004-05-04 10:04:54+01  taylor
'Header addition check in
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports Intranet.intranet
Imports Intranet.intranet.jobs
Imports Intranet.intranet.Support
Imports MedscreenLib
Imports MedscreenCommonGui.ListViewItems
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : FrmCollOfficer
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for the selection of a collecting Officer
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> $Date: 2005-11-29 07:31:33+00 $</date>
''' <Action>$Revision: 1.0 $</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class FrmCollOfficer
    Inherits System.Windows.Forms.Form
    Private ColumnDirection(10) As Integer
    'Private RootNode As TreeNode = New TreeNode("All")
    'Private UKNode As TreeNode = New TreeNode("UK")
    'Private OverseasNode As TreeNode = New TreeNode("Overseas")
    Private myCollectionDate As Date
    Private intEmailChecked As Integer = -1
    Private myCollection As Intranet.intranet.jobs.CMJob
    Friend WithEvents chDistanceFromBase As System.Windows.Forms.ColumnHeader
    Friend WithEvents chAvgCostperJob As System.Windows.Forms.ColumnHeader


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible values for the collecting officer to take on 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum CollOffFrmType
        ''' <summary>Send Collection</summary>
        Send
        ''' <summary>Confirmation Mode</summary>
        Confirm
    End Enum

    Public Property collection() As Intranet.intranet.jobs.CMJob
        Get
            Return myCollection
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CMJob)
            myCollection = Value
            If Not Value Is Nothing Then
                Me.lblCollection.Text = myCollection.CMID
                Me.LblCustomer.Text = "Customer : " & myCollection.ClientId
                Me.lblPort.Text = "Port : " & myCollection.Port

            End If
        End Set
    End Property


    Private FrmType As CollOffFrmType = CollOffFrmType.Send

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the form type property
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormType() As CollOffFrmType
        Get
            Return FrmType
        End Get
        Set(ByVal Value As CollOffFrmType)
            FrmType = Value
        End Set
    End Property

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        Dim I As Integer

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        ' CConnection.DatabaseInUse = CConnection.DatabaseInUse
        'Add any initialization after the InitializeComponent() call
        For I = 0 To 10
            ColumnDirection(I) = 1
        Next

        Try
            Dim objPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item("ENGLAND")
            'GetCountries()
            'Me.TreeView1.Nodes.Add(RootNode)
            'RootNode.Nodes.Add(UKNode)
            'RootNode.Nodes.Add(OverseasNode)

            'Dim aNode As SiteNode
            'aNode = New SiteNode(objPort)
            'UKNode.Nodes.Add(aNode)
            'AddChildNode(aNode, objPort)

            'objPort = cSupport.Ports.Item("WALES")
            'aNode = New SiteNode(objPort)
            'UKNode.Nodes.Add(aNode)
            'AddChildNode(aNode, objPort)
            'objPort = cSupport.Ports.Item("SCOTLAND")
            'aNode = New SiteNode(objPort)
            'UKNode.Nodes.Add(aNode)
            'AddChildNode(aNode, objPort)
            'Find continental ports
            Dim strSelectPortByType As String = MedscreenCommonGUIConfig.CollectionQuery("SelectPortByType")
            strSelectPortByType = Medscreen.ExpandQuery(strSelectPortByType, "D")
            Dim oCmd As New OleDb.OleDbCommand(strSelectPortByType)
            Dim oRead As OleDb.OleDbDataReader
            Try
                oCmd.Connection = CConnection.DbConnection
                CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                Dim oColl As New Collection()
                Dim strPort As String
                While oRead.Read
                    strPort = oRead.GetString(0)
                    oColl.Add(strPort)
                End While
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                Dim ij As Integer
                For ij = 1 To oColl.Count
                    strPort = oColl.Item(ij)
                    objPort = cSupport.Ports.Item(strPort)
                    If Not objPort Is Nothing Then
                        'aNode = New SiteNode(objPort)
                        'OverseasNode.Nodes.Add(aNode)
                        'AddChildNode(aNode, objPort)
                    End If
                Next

            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
        End Try
        Catch ex As Exception
        End Try


    End Sub

    Private Sub AddChildNodex(ByVal Anode As SiteNode, ByVal objPort As Intranet.intranet.jobs.Port)
        Dim bNode As SiteNode
        Dim aPort As Intranet.intranet.jobs.Port
        For Each aPort In objPort.Children
            bNode = New SiteNode(aPort)
            Anode.Nodes.Add(bNode)
            'AddChildNode(bNode, aPort)
        Next
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents lvOfficerPorts As System.Windows.Forms.ListView
    Friend WithEvents colID As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColForeName As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSurname As System.Windows.Forms.ColumnHeader
    Friend WithEvents Splitter3 As System.Windows.Forms.Splitter
    Friend WithEvents TreePorts As System.Windows.Forms.TreeView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOfficer As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    'Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents lblPort As System.Windows.Forms.Label
    Friend WithEvents chOfficerCost As System.Windows.Forms.ColumnHeader
    Friend WithEvents chCostType As System.Windows.Forms.ColumnHeader
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Friend WithEvents Splitter4 As System.Windows.Forms.Splitter
    Friend WithEvents LVRedirect As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents multisendImg As System.Windows.Forms.ImageList
    Friend WithEvents LblCustomer As System.Windows.Forms.Label
    Friend WithEvents lblCollection As System.Windows.Forms.Label
    Friend WithEvents ctxOfficer As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents ctxOfficerEmail As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuSendEmail As System.Windows.Forms.MenuItem
    Friend WithEvents cmdClosest As System.Windows.Forms.Button
    Friend WithEvents chNoOfJobs As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCollOfficer))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdClosest = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lblCollection = New System.Windows.Forms.Label
        Me.LblCustomer = New System.Windows.Forms.Label
        Me.lblPort = New System.Windows.Forms.Label
        Me.txtOfficer = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.lvOfficerPorts = New System.Windows.Forms.ListView
        Me.colID = New System.Windows.Forms.ColumnHeader
        Me.ColForeName = New System.Windows.Forms.ColumnHeader
        Me.colSurname = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.chOfficerCost = New System.Windows.Forms.ColumnHeader
        Me.chCostType = New System.Windows.Forms.ColumnHeader
        Me.chNoOfJobs = New System.Windows.Forms.ColumnHeader
        Me.ctxOfficer = New System.Windows.Forms.ContextMenu
        Me.mnuAddOfficer = New System.Windows.Forms.MenuItem
        Me.mnuRemoveOfficer = New System.Windows.Forms.MenuItem
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter3 = New System.Windows.Forms.Splitter
        Me.TreePorts = New System.Windows.Forms.TreeView
        Me.Splitter4 = New System.Windows.Forms.Splitter
        Me.LVRedirect = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ctxOfficerEmail = New System.Windows.Forms.ContextMenu
        Me.mnuSendEmail = New System.Windows.Forms.MenuItem
        Me.multisendImg = New System.Windows.Forms.ImageList(Me.components)
        Me.chDistanceFromBase = New System.Windows.Forms.ColumnHeader
        Me.chAvgCostperJob = New System.Windows.Forms.ColumnHeader
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdClosest)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 541)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(800, 32)
        Me.Panel1.TabIndex = 0
        '
        'cmdClosest
        '
        Me.cmdClosest.Location = New System.Drawing.Point(48, 8)
        Me.cmdClosest.Name = "cmdClosest"
        Me.cmdClosest.Size = New System.Drawing.Size(112, 23)
        Me.cmdClosest.TabIndex = 8
        Me.cmdClosest.Text = "Closest Officer"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(717, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(631, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lblCollection)
        Me.Panel2.Controls.Add(Me.LblCustomer)
        Me.Panel2.Controls.Add(Me.lblPort)
        Me.Panel2.Controls.Add(Me.txtOfficer)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(800, 32)
        Me.Panel2.TabIndex = 1
        '
        'lblCollection
        '
        Me.lblCollection.Location = New System.Drawing.Point(664, 8)
        Me.lblCollection.Name = "lblCollection"
        Me.lblCollection.Size = New System.Drawing.Size(100, 23)
        Me.lblCollection.TabIndex = 4
        '
        'LblCustomer
        '
        Me.LblCustomer.Location = New System.Drawing.Point(480, 8)
        Me.LblCustomer.Name = "LblCustomer"
        Me.LblCustomer.Size = New System.Drawing.Size(176, 23)
        Me.LblCustomer.TabIndex = 3
        Me.LblCustomer.Text = "Customer : "
        '
        'lblPort
        '
        Me.lblPort.Location = New System.Drawing.Point(264, 8)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(208, 16)
        Me.lblPort.TabIndex = 2
        Me.lblPort.Text = "Port : "
        '
        'txtOfficer
        '
        Me.txtOfficer.Enabled = False
        Me.txtOfficer.Location = New System.Drawing.Point(120, 7)
        Me.txtOfficer.Name = "txtOfficer"
        Me.txtOfficer.Size = New System.Drawing.Size(100, 20)
        Me.txtOfficer.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Current Officer"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 32)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 509)
        Me.Splitter1.TabIndex = 3
        Me.Splitter1.TabStop = False
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.HtmlEditor1.Location = New System.Drawing.Point(3, 229)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HtmlEditor1.Size = New System.Drawing.Size(797, 312)
        Me.HtmlEditor1.TabIndex = 6
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Splitter2.Location = New System.Drawing.Point(3, 226)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(797, 3)
        Me.Splitter2.TabIndex = 7
        Me.Splitter2.TabStop = False
        '
        'lvOfficerPorts
        '
        Me.lvOfficerPorts.AllowColumnReorder = True
        Me.lvOfficerPorts.BackColor = System.Drawing.SystemColors.Info
        Me.lvOfficerPorts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colID, Me.ColForeName, Me.colSurname, Me.ColumnHeader1, Me.chOfficerCost, Me.chCostType, Me.chAvgCostperJob, Me.chNoOfJobs, Me.chDistanceFromBase})
        Me.lvOfficerPorts.ContextMenu = Me.ctxOfficer
        Me.lvOfficerPorts.Dock = System.Windows.Forms.DockStyle.Left
        Me.lvOfficerPorts.FullRowSelect = True
        Me.lvOfficerPorts.GridLines = True
        Me.lvOfficerPorts.HideSelection = False
        Me.lvOfficerPorts.LargeImageList = Me.ImageList2
        Me.lvOfficerPorts.Location = New System.Drawing.Point(3, 32)
        Me.lvOfficerPorts.Name = "lvOfficerPorts"
        Me.lvOfficerPorts.Size = New System.Drawing.Size(485, 194)
        Me.lvOfficerPorts.SmallImageList = Me.ImageList1
        Me.lvOfficerPorts.TabIndex = 8
        Me.lvOfficerPorts.UseCompatibleStateImageBehavior = False
        Me.lvOfficerPorts.View = System.Windows.Forms.View.Details
        '
        'colID
        '
        Me.colID.Text = "Identity"
        Me.colID.Width = 80
        '
        'ColForeName
        '
        Me.ColForeName.Text = "Fore Name"
        Me.ColForeName.Width = 80
        '
        'colSurname
        '
        Me.colSurname.Text = "Surname"
        Me.colSurname.Width = 80
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Pin Person"
        '
        'chOfficerCost
        '
        Me.chOfficerCost.Text = "Cost"
        '
        'chCostType
        '
        Me.chCostType.Text = "Cost type"
        '
        'chNoOfJobs
        '
        Me.chNoOfJobs.Text = "No of Jobs"
        '
        'ctxOfficer
        '
        Me.ctxOfficer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddOfficer, Me.mnuRemoveOfficer})
        '
        'mnuAddOfficer
        '
        Me.mnuAddOfficer.Index = 0
        Me.mnuAddOfficer.Text = "&Add Officer"
        '
        'mnuRemoveOfficer
        '
        Me.mnuRemoveOfficer.Index = 1
        Me.mnuRemoveOfficer.Text = "&Remove Officer"
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "")
        Me.ImageList2.Images.SetKeyName(1, "")
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        '
        'Splitter3
        '
        Me.Splitter3.Location = New System.Drawing.Point(488, 32)
        Me.Splitter3.Name = "Splitter3"
        Me.Splitter3.Size = New System.Drawing.Size(3, 194)
        Me.Splitter3.TabIndex = 9
        Me.Splitter3.TabStop = False
        '
        'TreePorts
        '
        Me.TreePorts.BackColor = System.Drawing.SystemColors.Info
        Me.TreePorts.Dock = System.Windows.Forms.DockStyle.Top
        Me.TreePorts.Location = New System.Drawing.Point(491, 32)
        Me.TreePorts.Name = "TreePorts"
        Me.TreePorts.Size = New System.Drawing.Size(309, 120)
        Me.TreePorts.TabIndex = 10
        '
        'Splitter4
        '
        Me.Splitter4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter4.Location = New System.Drawing.Point(491, 152)
        Me.Splitter4.Name = "Splitter4"
        Me.Splitter4.Size = New System.Drawing.Size(309, 3)
        Me.Splitter4.TabIndex = 11
        Me.Splitter4.TabStop = False
        '
        'LVRedirect
        '
        Me.LVRedirect.BackColor = System.Drawing.SystemColors.Info
        Me.LVRedirect.CheckBoxes = True
        Me.LVRedirect.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.LVRedirect.ContextMenu = Me.ctxOfficerEmail
        Me.LVRedirect.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LVRedirect.FullRowSelect = True
        Me.LVRedirect.Location = New System.Drawing.Point(491, 155)
        Me.LVRedirect.Name = "LVRedirect"
        Me.LVRedirect.Size = New System.Drawing.Size(309, 71)
        Me.LVRedirect.SmallImageList = Me.multisendImg
        Me.LVRedirect.TabIndex = 12
        Me.LVRedirect.UseCompatibleStateImageBehavior = False
        Me.LVRedirect.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "From"
        Me.ColumnHeader4.Width = 80
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "To"
        Me.ColumnHeader5.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Address"
        Me.ColumnHeader6.Width = 200
        '
        'ctxOfficerEmail
        '
        Me.ctxOfficerEmail.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSendEmail})
        '
        'mnuSendEmail
        '
        Me.mnuSendEmail.Index = 0
        Me.mnuSendEmail.Text = "&Send Email"
        '
        'multisendImg
        '
        Me.multisendImg.ImageStream = CType(resources.GetObject("multisendImg.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.multisendImg.TransparentColor = System.Drawing.Color.Transparent
        Me.multisendImg.Images.SetKeyName(0, "")
        Me.multisendImg.Images.SetKeyName(1, "")
        Me.multisendImg.Images.SetKeyName(2, "")
        Me.multisendImg.Images.SetKeyName(3, "")
        Me.multisendImg.Images.SetKeyName(4, "")
        Me.multisendImg.Images.SetKeyName(5, "")
        '
        'chDistanceFromBase
        '
        Me.chDistanceFromBase.Text = "Distance"
        '
        'chAvgCostperJob
        '
        Me.chAvgCostperJob.Text = "avg cost job"
        '
        'FrmCollOfficer
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(800, 573)
        Me.Controls.Add(Me.LVRedirect)
        Me.Controls.Add(Me.Splitter4)
        Me.Controls.Add(Me.TreePorts)
        Me.Controls.Add(Me.Splitter3)
        Me.Controls.Add(Me.lvOfficerPorts)
        Me.Controls.Add(Me.Splitter2)
        Me.Controls.Add(Me.HtmlEditor1)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCollOfficer"
        Me.Text = "Collection Officer Assignment"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Private oConn As New OleDb.OleDbConnection()
    'Private oCmd As New OracleClient.OracleCommand()
    Private oCmd As New OleDb.OleDbCommand()
    Private strCountry As String = " in ('ENGLAND','WALES','SCOTLAND','NORTHERN_IRELAND')"
    Private strActualCountry As String = "ENGLAND"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible values for Region 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum RegionsToBeSelected
        UK = 1
        OS = 2
        All = 3
        byCountry = 4
        Misc = 5
    End Enum

    Private SelectedRegion As RegionsToBeSelected = RegionsToBeSelected.UK
    Private strCollId As String = ""
    Private strPort As String
    Private myPort As Intranet.intranet.jobs.Port

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Port passed into the form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Port() As String
        Get
            Return strPort
        End Get
        Set(ByVal Value As String)
            Dim strQuery As String
            'Dim oRead As OracleClient.OracleDataReader
            Dim strlCountry As String
            Dim strRegion As String
            Dim rNode As TreeNode
            Dim aNode As SiteNode
            Dim bNode As SiteNode
            Dim blnOverseas As Boolean = False
            'clear local port variable 
            myPort = Nothing
            'Dim aPort As Intranet.intranet.jobs.Port

            strPort = Value
            'Check data integrity 
            If strPort Is Nothing Then Exit Property
            If strPort.Length = 0 Then Exit Property

            'if no ports load the data set 
            If cSupport.Ports.Count = 0 Then cSupport.Ports.Load()
            myPort = cSupport.Ports.Item(strPort)

            If myPort Is Nothing Then Exit Property 'Check validity of port
            Try
                'rNode = RootNode                    ' Get root node

                If Not myPort.IsUK Then              'Not UK
                    'rNode = OverseasNode            ' Overseas
                    blnOverseas = True
                Else
                    'rNode = UKNode ' UK
                End If

                'Me.TreeView1.SelectedNode = Nothing  'Remove selection 
                Dim strReturn As String = CConnection.PackageStringList("LIB_CCTOOL.PortPathToRoot", myPort.Identity)
                If strReturn.Trim.Length > 0 Then
                    'We should now have a list of ports to the root 
                    Dim PortArray As String() = strReturn.Split(New Char() {"|"})
                    Dim i As Integer = PortArray.Length - 1

                    While i >= 0
                        Dim strPortId As String = PortArray.GetValue(i)
                        Dim blnFail As Boolean = True
                        For Each bNode In rNode.Nodes
                            If bNode.Port.Identity.ToUpper = strPortId.ToUpper Then
                                rNode = bNode
                                blnFail = False
                                Exit For
                            End If
                        Next
                        If blnFail Then 'Can't find port add node 
                            Dim objPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(strPortId.ToUpper)
                            bNode = New SiteNode(objPort)
                            rNode.Nodes.Add(bNode)
                            rNode = bNode
                        End If
                        i -= 1
                    End While

                End If

                'Find the node for this port 
                'bNode = FindPort(aPort.Country, aNode)
                If bNode Is Nothing Then Exit Property


                'bNode = FindPort(aPort.Identity, bNode)

                'If Not bNode Is Nothing Then
                '    'Me.GetPortSites(aNode)
                'Else
                '    Dim oPort As Intranet.intranet.jobs.Port
                '    oPort = aPort.ParentPort
                '    While Not oPort Is Nothing AndAlso bNode Is Nothing
                '        bNode = FindPort(oPort.Identity, aNode)
                '        oPort = oPort.ParentPort
                '    End While

                'End If
                If Not bNode Is Nothing Then
                    'Me.TreeView1.SelectedNode = bNode
                    Me.lblPort.Text = "Port : " & bNode.Text
                End If

            Catch ex As Exception
            Finally
            End Try
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find a port returning the node of the port
    ''' </summary>
    ''' <param name="PortId">Port to find</param>
    ''' <param name="aNode">Current node</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function FindPort(ByVal PortId As String, ByVal aNode As TreeNode) As TreeNode
        Dim bNode As TreeNode
        Dim rNode As TreeNode
        'If aNode.Nodes.Count = 0 Then Me.GetPortSites(aNode)
        For Each bNode In aNode.Nodes                   'Check all nodes
            If bNode.Text.ToUpper = PortId.ToUpper Then     'If a match found
                rNode = bNode                           'return the found node
                Exit For
            Else
                rNode = FindPort(PortId, bNode)         'Do  a recursive search on the node 
                If Not rNode Is Nothing Then
                    Exit For
                End If
            End If
        Next
        Return rNode
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get collecting officer details 
    ''' </summary>
    ''' <param name="CID">Optional officer ID passed to Fill method</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub GetCollectors(Optional ByVal CID As String = "")
        'Dim oRead As OracleClient.OracleDataReader
        Dim oRead As OleDb.OleDbDataReader
        'Dim lvItem As CofficerItem

        Try
            'oConn.ConnectionString = ModTpanel.ConnectString(2)
            oCmd.Connection = CConnection.DbConnection
            Dim strOfficerInfo As String = MedscreenCommonGUIConfig.CollectionQuery("OfficerInfo")

            'oCmd.CommandText = "Select IDENTITY,SURNAME,FIRST_NAME," & _
            '    "HOME_PHONE,WORK_PHONE,MOBILE_PHONE,BLEEP_NUM,COUNTRY,area_coord From Coll_officer "
            If SelectedRegion = RegionsToBeSelected.UK Then
                strOfficerInfo = Medscreen.ExpandQuery(strOfficerInfo, " Where COUNTRY " & strCountry)
                'oCmd.CommandText += " Where COUNTRY " & strCountry
            ElseIf SelectedRegion = RegionsToBeSelected.OS Then
                strOfficerInfo = Medscreen.ExpandQuery(strOfficerInfo, " Where COUNTRY NOT " & strCountry)
                'oCmd.CommandText += " Where COUNTRY NOT " & strCountry
            ElseIf SelectedRegion = RegionsToBeSelected.byCountry Then
                strOfficerInfo = Medscreen.ExpandQuery(strOfficerInfo, " Where Country = '" & Me.strActualCountry.Trim & "'")
                'oCmd.CommandText += " Where Country = '" & Me.strActualCountry.Trim & "'"
            ElseIf SelectedRegion = RegionsToBeSelected.Misc Then
                strOfficerInfo = Medscreen.ExpandQuery(strOfficerInfo, " Where COUNTRY < 'A'")
                'oCmd.CommandText += " Where Country <= 'A'"
            Else
                strOfficerInfo = Medscreen.ExpandQuery(strOfficerInfo, " Where COUNTRY= country")
                'oCmd.CommandText += " Where Country = Country "
            End If
            'oCmd.CommandText += " and removeflag = 'F'"
            'oCmd.CommandText += " order by identity"
            oCmd.CommandText = strOfficerInfo
            CConnection.SetConnOpen()
            oRead = oCmd.ExecuteReader                      'Get collecting officer details 
            FillCollectors(oRead, CID)                      'Fill control 
        Catch ex As Exception
        Finally
            'oConn.Close()
            Me.lvOfficerPorts.EndUpdate()
            Me.lvOfficerPorts.Invalidate()
            CConnection.SetConnClosed()
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the contents of the collectors list view 
    ''' </summary>
    ''' <param name="Oread">Dataset</param>
    ''' <param name="CID">Optional Officer ID</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillCollectors(ByVal Oread As OleDb.OleDbDataReader, Optional ByVal CID As String = "")
        Me.lvOfficerPorts.Items.Clear()                  'Clear the list view and get ready to fill
        Me.lvOfficerPorts.SelectedItems.Clear()
        Me.lvOfficerPorts.BeginUpdate()

        Dim lvItem As LVCofficerItem
        Dim COff As Intranet.intranet.jobs.CollectingOfficer        'Holds collecting Officer

        If strPort Is Nothing Then Exit Sub ''                       No port bog off
        While Oread.Read
            COff = cSupport.CollectingOfficers.Item(Oread.GetValue(0))     'Get collecting officer
            If COff Is Nothing Then                                 'No officer !!!
            Else
                lvItem = New LVCofficerItem(COff, strPort, Me.myCollectionDate) 'Put info into specialist element
                If lvItem.OfficerId = CID Then lvItem.Selected = True 'If we match the Officer make it selected
                lvItem = Me.lvOfficerPorts.Items.Add(lvItem)
                lvItem.BackColor = System.Drawing.SystemColors.Info
                lvItem.UseItemStyleForSubItems = False
            End If
        End While
        Oread.Close()                                               'Close reader
        'Set up sorter
        Me.lvOfficerPorts.ListViewItemSorter = New MedscreenCommonGui.Comparers.ListViewItemComparer(3, -1)


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in the collection date (for holiday checks)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CollectionDate() As Date
        Get
            Return myCollectionDate
        End Get
        Set(ByVal Value As Date)
            myCollectionDate = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collecting Officer 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Collector() As String
        Get
            Return Me.strCollId
        End Get
        Set(ByVal Value As String)
            strCollId = Value
            If strCollId Is Nothing Then
                strCollId = ""
            End If
            Me.txtOfficer.Text = strCollId                      'Set display 

            If strCollId.Trim.Length = 0 Then                   'We have a collecting Officer 
                If strPort.Trim.Length = 0 Then
                    GetCollectors()                             'Get collecting officers that match
                    Exit Property
                Else
                    Me.GetPortCollOfficers1(strPort)             'Get just those for this port 
                    Exit Property
                End If
            Else
                If strPort.Trim.Length = 0 Then                 'See if we have a port
                    GetCollectors(Value)                    'Get Officers passing collecting officer
                Else
                    Me.GetPortCollOfficers1(strPort)
                End If
            End If
            strCollId = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Resort List view on this column
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ListView1_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs)
        Dim intdirection As Integer

        intdirection = Me.ColumnDirection(e.Column)
        Me.ColumnDirection(e.Column) *= -1

        Me.lvOfficerPorts.ListViewItemSorter = New MedscreenCommonGui.Comparers.ListViewItemComparer(e.Column, intdirection)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check UK only 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub rbUKOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectedRegion = RegionsToBeSelected.UK
        Me.GetCollectors()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ccheck Overseas only 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub rbOSOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectedRegion = RegionsToBeSelected.OS
        Me.GetCollectors()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check all locations
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub rbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectedRegion = RegionsToBeSelected.All
        Me.GetCollectors()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Treview has been selected
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    ''Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
    ''    Dim strText As String
    ''    Dim Anode As TreeNode

    ''    Anode = Me.TreeView1.SelectedNode
    ''    Me.HtmlEditor1.DocumentText =("<HTML></HTML>")
    ''    If TypeOf Anode Is SiteNode Then
    ''        If TypeOf Anode Is SiteRegionNode Then
    ''            If Anode.Nodes.Count = 0 Then
    ''                GetPortSites(Anode)
    ''            End If
    ''            Me.GetRegionCollOfficers(CType(Anode, SiteNode).Port.Identity)
    ''        Else
    ''            GetPortCollOfficers(CType(Anode, SiteNode).Port.Identity)
    ''        End If
    ''    Else
    ''        strText = Anode.Text
    ''        Me.TreePorts.Nodes.Clear()
    ''        Me.ListView1.Items.Clear()

    ''        If strText.ToUpper = "UK" Then
    ''            Me.SelectedRegion = RegionsToBeSelected.UK
    ''        ElseIf strText.ToUpper = "OVERSEAS" Then
    ''            Me.SelectedRegion = RegionsToBeSelected.OS
    ''        ElseIf strText.ToUpper = "ALL" Then
    ''            Me.SelectedRegion = RegionsToBeSelected.All
    ''        ElseIf strText.Trim.Length = 0 Then
    ''            Me.SelectedRegion = RegionsToBeSelected.Misc
    ''        Else
    ''            Me.SelectedRegion = RegionsToBeSelected.byCountry
    ''            Me.strActualCountry = strText.ToUpper.Trim

    ''        End If
    ''        Me.GetCollectors()
    ''        If Anode.Nodes.Count = 0 Then
    ''            GetPortSites(Anode)
    ''        End If
    ''    End If
    ''End Sub

    Private Sub GetPortCollOfficers1(ByVal Port As String)
        If myPort Is Nothing Then
            GetPortCollOfficers(Port)
            Exit Sub
        ElseIf myPort IsNot Nothing AndAlso myPort.UNLocode.Trim.Length = 0 Then
            GetPortCollOfficers(Port)
            Exit Sub
        End If
        chCostType.Text = "Avg Cost per Sample"
        chOfficerCost.Text = "Avg Expenses"
        Dim PortColl As New BOSampleManager.CollectorPorts(myPort.UNLocode, BOSampleManager.CollectorPorts.LoadBy.Ports)
        PortColl.SortByDistance()
        Try
            Me.lvOfficerPorts.Items.Clear()
            Me.lvOfficerPorts.BeginUpdate()

            Dim blnSelected As Boolean = False
            For Each pcoll As BOSampleManager.CollectorPort In PortColl
                Dim lvItem As New LVCofficerItem(pcoll, Me.myCollectionDate)
                Me.lvOfficerPorts.Items.Add(lvItem)
                If lvItem.OfficerId = Me.Collector Then
                    lvItem.Selected = True

                    blnSelected = True
                End If
            Next
            lvOfficerPorts.EndUpdate()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub GetPortCollOfficers(ByVal strPort As String)
        Dim oRead As OleDb.OleDbDataReader
        Dim lvItem As LVCofficerItem

        Try
            chCostType.Text = "Cost Type"
            chOfficerCost.Text = "Cost"

            'oConn.ConnectionString = ModTpanel.ConnectString(2)
            Me.strPort = strPort
            oCmd.Connection = CConnection.DbConnection
            'Find the query then run it
            Dim strOfficerPortsByPort As String = MedscreenCommonGUIConfig.CollectionQuery("OfficerPortsByPort")
            strOfficerPortsByPort = Medscreen.ExpandQuery(strOfficerPortsByPort, strPort)
            oCmd.CommandText = strOfficerPortsByPort
            CConnection.SetConnOpen()
            oRead = oCmd.ExecuteReader
            Me.lvOfficerPorts.Items.Clear()
            Me.lvOfficerPorts.BeginUpdate()
            Dim oCollList As New Collection()
            While oRead.Read
                oCollList.Add(oRead.GetValue(0))
            End While
            Dim intI As Integer

            oRead.Close()
            Dim blnSelected As Boolean = False
            For intI = 1 To oCollList.Count
                Dim strOfficer As String = oCollList.Item(intI)
                lvItem = Me.SetCollOfficerValues(strOfficer)
                Me.lvOfficerPorts.Items.Add(lvItem)
                If lvItem.OfficerId = Me.Collector Then
                    lvItem.Selected = True

                    blnSelected = True
                End If
            Next
            If lvOfficerPorts.Items.Count > 0 Then
                If Not blnSelected Then
                    Me.lvOfficerPorts.Items(0).Selected = True
                End If
            End If
            Me.ListView1_Click(Nothing, Nothing)

        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            'oConn.Close()
            Me.lvOfficerPorts.EndUpdate()
            Me.lvOfficerPorts.Invalidate()
            CConnection.SetConnClosed()
        End Try


    End Sub

    Private Function SetCollOfficerValues(ByVal Officer As String) As LVCofficerItem
        Dim lvItem As LVCofficerItem
        Dim COff As Intranet.intranet.jobs.CollectingOfficer
        COff = cSupport.CollectingOfficers.Item(Officer)

        If strPort Is Nothing Then Exit Function
        If Not COff Is Nothing Then

            lvItem = New LVCofficerItem(COff, strPort, Me.myCollectionDate)

            Return lvItem
        End If

    End Function

    Private Sub GetRegionCollOfficers(ByVal strRegion As String)
        Dim oRead As OleDb.OleDbDataReader
        'Dim lvItem As CofficerItem

        Try
            'oConn.ConnectionString = ModTpanel.ConnectString(2)
            oCmd.Connection = CConnection.DbConnection
            Dim strOfficerByRegion As String = MedscreenCommonGUIConfig.CollectionQuery("OfficerByRegion")
            strOfficerByRegion = Medscreen.ExpandQuery(strOfficerByRegion, strRegion)
            oCmd.CommandText = strOfficerByRegion
            'oCmd.CommandText = "Select distinct o.IDENTITY " & _
            '    "From Coll_officer o, Coll_ports pc, port p  " & _
            '    "where o.identity = pc.coll_id and p.identity = pc.port_id" & _
            '    " and p.region = '" & strRegion & "'"

            'oCmd.CommandText += " and o.removeflag = 'F' and p.removeflag = 'F'"
            'oCmd.CommandText += " order by o.identity"
            'oConn.Open()
            CConnection.SetConnOpen()
            oRead = oCmd.ExecuteReader
            Me.FillCollectors(oRead)


        Catch ex As Exception
        Finally
            'oConn.Close()
            Me.lvOfficerPorts.EndUpdate()
            Me.lvOfficerPorts.Invalidate()
            CConnection.SetConnClosed()
        End Try


    End Sub


    Private Sub GetPortSites(ByVal aNode As TreeNode)
        'Dim Oread As OracleClient.OracleDataReader
        Dim strType As String
        Dim sNode As SiteNode

        If aNode.Nodes.Count > 0 Then Exit Sub
        Try
            Dim oPorti As Intranet.intranet.jobs.Port = cSupport.Ports.Item(aNode.Text)

            If oPorti Is Nothing Then Exit Sub

            'im Ports As Intranet.intranet.jobs.CentreCollection = ModTpanel.Ports.Children(oPort)

            If oPorti.Children Is Nothing Then Exit Sub

            Dim oPort As Intranet.intranet.jobs.Port
            For Each oPort In oPorti.Children
                strType = oPort.PortType
                If strType = "R" Then
                    sNode = New SiteRegionNode(oPort)
                    sNode.Text = oPort.Identity
                    aNode.Nodes.Add(sNode)
                ElseIf strType = "P" Then
                    sNode = New SiteSiteNode(oPort)
                    sNode.Text = oPort.Identity
                    aNode.Nodes.Add(sNode)
                ElseIf strType = "S" Then
                    sNode = New SiteSiteNode(oPort)
                    sNode.Text = oPort.Identity
                    aNode.Nodes.Add(sNode)
                Else
                    sNode = New SiteNode(oPort)
                    sNode.Text = oPort.Identity
                    aNode.Nodes.Add(sNode)
                End If
                GetPortSites(sNode)
            Next
        Catch ex As Exception
        Finally
            'oConn.Close()
            'aNode.Expand()
            'Me.TreeView1.EndUpdate()
        End Try

    End Sub

    Private Sub AddChildren(ByVal aNode As TreeNode, ByVal aPort As Intranet.intranet.jobs.Port)
        Dim cCol As Intranet.intranet.jobs.CentreCollection

        If aPort Is Nothing Then Exit Sub

        cCol = cSupport.Ports.Children(aPort)
        Dim i As Integer
        For i = 0 To cCol.Count - 1
            Dim objPort As Intranet.intranet.jobs.Port
            objPort = cCol.Item(i)
            Dim rNode As New TreeNode(objPort.Info)
            aNode.Nodes.Add(rNode)
            AddChildren(rNode, objPort)
        Next
    End Sub



    Private Sub GetPorts(ByVal COfficer As String)
        Dim oRead As OleDb.OleDbDataReader
        Dim onode As TreeNode

        Try
            Me.TreePorts.Nodes.Clear()
            'oConn.ConnectionString = ModTpanel.ConnectString(2)
            oCmd.Connection = CConnection.DbConnection
            Dim strPortByOfficer As String = MedscreenCommonGUIConfig.CollectionQuery("PortByOfficer")
            strPortByOfficer = Medscreen.ExpandQuery(strPortByOfficer, COfficer)
            oCmd.CommandText = strPortByOfficer '"Select distinct PORT_ID,COSTS_AGREED," & _
            '"CONT_COSTS from coll_ports where COLL_ID = '" & COfficer & _
            '"' and removeflag = 'F' order by port_id"
            CConnection.SetConnOpen()
            oRead = oCmd.ExecuteReader
            While oRead.Read
                Dim portId As String = oRead.GetString(oRead.GetOrdinal("PORTID"))
                'Dim PortName As String = CConnection.PackageStringList("Lib_collection.getport", portid)
                Me.TreePorts.Nodes.Add(portId)
            End While
            oRead.Close()
        Catch ex As Exception

        Finally
            If Not oRead.IsClosed Then oRead.Close()
            CConnection.SetConnClosed()
        End Try
    End Sub


    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvOfficerPorts.Click, lvOfficerPorts.ItemActivate
        Dim lvItem As LVCofficerItem
        Dim strId As String
        Dim strHTML As String
        Dim dQuote As String = Chr(34)

        Me.strCollId = ""
        If CurrentOfficer Is Nothing AndAlso Me.lvOfficerPorts.SelectedItems.Count <> 1 Then Exit Sub
        If Not CurrentOfficer Is Nothing Then
            lvItem = CurrentOfficer
            CurrentOfficer = Nothing
        Else
            lvItem = Me.lvOfficerPorts.SelectedItems.Item(0)
        End If
        strId = lvItem.Text
        Me.strCollId = strId
        Me.txtOfficer.Text = strId
        Me.GetPorts(strId)

        Dim CollOff As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(strId)
        If Not CollOff Is Nothing AndAlso Not CollOff.Ports Is Nothing Then
            Dim strXMLOff As String = CollOff.ToXml(, , , , True, True, True)
            'Dim strHTML2 = CollOff.ToHtml(Me.CollectionDate)

            Dim cPort As Intranet.intranet.jobs.CollPort
            cPort = CollOff.Ports.Item(Me.strPort, 0)
            If Not cPort Is Nothing Then
                Dim strXML As String = "<Poff>" & strXMLOff
                'Dim strXMLcPort As String = cPort.ToXML()
                'strHTML = cPort.ToHtml(Me.CollectionDate)
                'Dim oColl As New Collection()
                'oColl.Add(Me.myPort.Identity)
                'oColl.Add(Me.myPort.CountryId)
                'Dim strXML As String = CConnection.PackageStringList("lib_officer.ClosestOfficerstoPortXML", oColl)
                'If strXML Is Nothing OrElse strXML.Length < 20 Then
                'MsgBox("This port doesn't have any geographical details Longitude or Latitude, so this information can't be calculated.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                'Exit Sub
                'End If
                'Dim strHTML1 As String = Medscreen.ResolveStyleSheet(strXML, "Nearest.xslt", 0)
                'Dim strHead As String = Mid(strHTML, 1, strHTML.Length - 14)
                'Dim strTail As String = Mid(strHTML, strHTML.Length - 13)
                'strHTML = strHead & strhtml1 & strTail
                Dim strXMLPort As String
                If Not myPort Is Nothing Then
                    strXMLPort = Me.myPort.toXML(CentreCollection.XMLOptions.base)
                    strXML += strXMLPort
                End If
                strXML += "</Poff>"
                strHTML = Medscreen.ResolveStyleSheet(strXML, "CollOfficerEd.xsl", 0)
            Else
                Dim oColl As New Collection()
                oColl.Add(Me.myPort.Identity)
                oColl.Add(Me.myPort.CountryId)
                Dim strXML As String = CConnection.PackageStringList("lib_officer.ClosestOfficerstoPortXML", oColl)
                If strXML Is Nothing OrElse strXML.Length < 20 Then
                    MsgBox("This port doesn't have any geographical details Longitude or Latitude, so this information can't be calculated.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    'Exit Sub
                End If
                Dim strHTML1 As String = Medscreen.ResolveStyleSheet(strXML, "Nearest.xslt", 0)

                strHTML = CollOff.ToHtml(Me.CollectionDate)
            End If
            Me.LVRedirect.Items.Clear()
            Me.intEmailChecked = -1
            'Default destination
            If CollOff.SendDestination.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem("Default Send")
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, ""))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.SendDestination))
                lvItemR.ImageIndex = 4
                lvItemR.Checked = True
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            'Deal with multisend
            If Not CollOff.MultiSendAddresses Is Nothing Then
                Dim aSend As MultiSendColl

                For Each aSend In CollOff.MultiSendAddresses

                    Dim lvItemR As ListViewItem = New ListViewItem(aSend.SourceOfficer)
                    lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, aSend.DestOfficer))
                    lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, aSend.EmailFaxNumber))
                    lvItemR.ImageIndex = 5
                    Me.LVRedirect.Items.Add(lvItemR)
                Next
            End If
            If CollOff.HomePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Home Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.HomePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.OfficePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Office Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.OfficePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.MobilePhone.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Mobile Phone"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.MobilePhone))
                lvItemR.ImageIndex = 3
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.HomeFax.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Home Fax"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.HomeFax))
                lvItemR.ImageIndex = 2
                Me.LVRedirect.Items.Add(lvItemR)
            End If
            If CollOff.Email.Trim.Length > 0 Then
                Dim lvItemR As ListViewItem = New ListViewItem(CollOff.Name)
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, "Email"))
                lvItemR.SubItems.Add(New ListViewItem.ListViewSubItem(lvItemR, CollOff.Email))
                lvItemR.ImageIndex = 1
                Me.LVRedirect.Items.Add(lvItemR)
            End If

        Else
            strHTML = "<HTML><head><LINK REL=STYLESHEET TYPE=" & _
            dQuote & "text/css" & dQuote & " HREF = " & dQuote & _
                Medscreen.StyleSheet & dQuote & " Title=" & dQuote & _
                "Style" & dQuote & "><body>"

            strHTML += "Collecting Officer " & lvItem.FirstName & " " & lvItem.LastName & "<BR>"
            strHTML += "Office Phone " & lvItem.OfficePhone & "<BR>"
            strHTML += "Home Phone " & lvItem.HomePhone & "<BR>"
            strHTML += "Mobile Phone " & lvItem.Mobile & "<BR>"
            strHTML += "Pager " & lvItem.Beeper & "<BR>"
            strHTML += "Costs agreed " & lvItem.Agreed & "<BR>"
            strHTML += "Extra Costs " & lvItem.Cost.ToString & "<BR>"


            strHTML += "</body></HTML>"
        End If

        Try
            Me.HtmlEditor1.DocumentText = (strHTML)
        Catch ex As Exception
        End Try
        Me.strCollId = strId

    End Sub

#Region "Site Nodes"


    Private Class SiteNode
        Inherits TreeNode
        Private oPort As Intranet.intranet.jobs.Port

        Public Sub New(ByVal Port As Intranet.intranet.jobs.Port)
            MyBase.New()
            Me.BackColor = System.Drawing.SystemColors.Info
            Me.ForeColor = System.Drawing.SystemColors.InfoText
            oPort = Port
            If oPort Is Nothing Then Exit Sub
            If oPort.UNLocode.Trim.Length > 0 Then
                Me.Text = oPort.SiteName & " (" & oPort.UNLocode & ")"
                If oPort.SubDivision.Trim.Length > 0 Then
                    Me.Text += " -" & oPort.SubDivision
                Else
                    'oPort.GetISOInfo()
                    'If oPort.SubDivision.Trim.Length > 0 Then
                    'Me.Text += " -" & oPort.SubDivision
                    'End If
                End If
            Else
                Me.Text = oPort.SiteName
            End If
        End Sub

        Public Property Port() As Intranet.intranet.jobs.Port
            Get
                Return oPort
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.Port)
                oPort = Value

            End Set
        End Property
    End Class

    'Private Class SiteSubItem
    '    Inherits ListViewItem.ListViewSubItem


    '    Public Sub New()
    '        MyBase.New()
    '        Me.BackColor = System.Drawing.SystemColors.Info
    '        Me.ForeColor = System.Drawing.SystemColors.InfoText

    '    End Sub
    'End Class

    Private Class SiteRegionNode
        Inherits SiteNode

        Public Sub New(ByVal Port As Intranet.intranet.jobs.Port)
            MyBase.New(Port)
            Me.BackColor = System.Drawing.SystemColors.Info
            Me.ForeColor = System.Drawing.SystemColors.InfoText
        End Sub
    End Class

    Private Class SiteSiteNode
        Inherits SiteNode

        Public Sub New(ByVal Port As Intranet.intranet.jobs.Port)
            MyBase.New(Port)
        End Sub
    End Class

#End Region


    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim CoFF As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(Me.Collector)

        If CoFF.Holidays.HolidayIn(Me.myCollectionDate) Then
            Dim hol As Intranet.intranet.jobs.CollectorHoliday
            hol = CoFF.Holidays.Item(Me.myCollectionDate)
            If Not hol Is Nothing Then
                MsgBox("Can't assign collecting officer due to holidays (" & _
                    hol.HolidayStart.ToString("dd-MMM-yyyy") & " - " & _
                    hol.HolidayEnd.ToString("dd-MMM-yyyy") & ")", MsgBoxStyle.OKOnly, "Holiday Error")
            Else
                MsgBox("Can't assign collecting officer due to holidays ", MsgBoxStyle.OKOnly, "Holiday Error")
            End If
            Me.DialogResult = DialogResult.None
            Exit Sub
        End If


        Dim dRet As MsgBoxResult
        'Get collection costs string from PL/SQL
        Dim strCost As String = CConnection.PackageStringList("Lib_collection.GetOfficerCostsString", Me.Port)

        Dim strSendTo As String = CConnection.PackageStringList("Lib_collection.GetOfficerName", Me.Collector)
        If FrmType = CollOffFrmType.Send Then
            dRet = MsgBox(strCost & vbCrLf & "Do you wish to Send to : " & strSendTo, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
        Else
            dRet = MsgBox("Do you wish to Confirm to : " & strSendTo, MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
        End If
        If dRet = MsgBoxResult.No Then
            Me.DialogResult = DialogResult.None
        Else
            Me.DialogResult = DialogResult.OK
        End If

    End Sub

    Private Sub mnuAddOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddOfficer.Click
        Dim myPort As Intranet.intranet.jobs.Port = cSupport.Ports.Item(strPort)

        If myPort Is Nothing Then Exit Sub

        Dim objCollPort As Intranet.intranet.jobs.CollPort = New Intranet.intranet.jobs.CollPort()
        objCollPort.PortId = myPort.PortId

        Dim frmAO As New MedscreenCommonGui.frmAssignOfficer()
        Try
            frmAO.CollPort = objCollPort
            If frmAO.ShowDialog = DialogResult.OK Then
                objCollPort = frmAO.CollPort
                objCollPort.Update()
                myPort.CollectingOfficers.Add(objCollPort)

                Dim lvItem As LVCofficerItem = New LVCofficerItem(objCollPort.Officer, strPort, Me.myCollectionDate)
                lvItem = Me.lvOfficerPorts.Items.Add(lvItem)
                lvItem.BackColor = System.Drawing.SystemColors.Info
                lvItem.UseItemStyleForSubItems = False
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "AddOfficer - frmcollofficer")
        End Try
    End Sub

    Private Sub mnuRemoveOfficer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuRemoveOfficer.Click
        If Me.lvOfficerPorts.SelectedItems.Count <> 1 Then Exit Sub

        Dim lvItem As ListViewItem = lvOfficerPorts.SelectedItems.Item(0)

        If TypeOf lvItem Is LVCofficerItem Then
            Dim lvCitem As LVCofficerItem = lvItem
            If Not lvCitem Is Nothing Then
                If Not lvCitem.CollectionPort Is Nothing Then
                    lvCitem.CollectionPort.Removed = True
                    lvCitem.CollectionPort.Update()
                    lvOfficerPorts.Items.RemoveAt(lvItem.Index)
                End If
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Send email to officer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSendEmail.Click
        If Me.intEmailChecked = -1 Then Exit Sub
        Dim CollOff As Intranet.intranet.jobs.CollectingOfficer = cSupport.CollectingOfficers.Item(strCollId)

        Dim strEmail As String = ""             'Default send destination to null string 
        If Me.intEmailChecked = 0 Then          'Default send method 
            If Not CollOff.CanEmailorFax Then
                MsgBox("Can't send to this type of address", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
                Exit Sub
            End If
            If CollOff.SendBy = Constants.CollOfficerSend.HomeEmail Or CollOff.SendBy = Constants.CollOfficerSend.WorkEmail Or _
            CollOff.SendBy = Constants.CollOfficerSend.HomeFax Or CollOff.SendBy = Constants.CollOfficerSend.WorkFax Then
                strEmail = CollOff.SendDestination
            End If
        Else 'from multisend 
            Dim lvitem As ListViewItem = Me.LVRedirect.Items(Me.intEmailChecked)
            If Not lvitem Is Nothing Then
                Dim sItem As ListViewItem.ListViewSubItem = lvitem.SubItems(1)
                Dim mailType As String = sItem.Text.ToUpper
                sItem = lvitem.SubItems(2)
                strEmail = Medscreen.RemoveChars(sItem.Text, " ")
                If InStr(mailType, "PHONE") Then
                    MsgBox("Can't send to a phone")
                    Exit Sub
                End If
                If Medscreen.IsNumber(strEmail) Then
                    strEmail += MedscreenLib.Constants.GCST_EMAIL_To_Fax_Send
                End If
            End If
        End If


        If strEmail.Trim.Length > 0 Then
            Dim objOutlook As Outlook.ApplicationClass = New Outlook.ApplicationClass()
            Dim objMessage As outlook.MailItemClass
            Try
                objMessage = objOutlook.CreateItem(outlook.OlItemType.olMailItem)
                objMessage.Subject = "Message from " & Intranet.intranet.Support.cSupport.User.Description
                objMessage.Recipients.Add(strEmail)
                objMessage.Display(True)
            Catch ex As Exception
            Finally
                objOutlook = Nothing
            End Try
        End If
    End Sub

    Private Sub LVRedirect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LVRedirect.Click
        Dim lvItem As ListViewItem
        If LVRedirect.SelectedItems.Count > 0 Then
            lvItem = LVRedirect.SelectedItems.Item(0)
            Me.intEmailChecked = lvItem.Index
        End If
    End Sub

    Dim CurrentOfficer As LVCofficerItem
    Dim bActivated As Boolean = False
    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        If strCollId.Trim.Length > 0 AndAlso Not bActivated Then
            bActivated = True
            Dim lvItem As LVCofficerItem
            Dim i As Integer
            lvOfficerPorts.SelectedItems.Clear()
            For i = 0 To lvOfficerPorts.Items.Count - 1
                lvItem = lvOfficerPorts.Items.Item(i)
                If lvItem.OfficerId = strCollId Then
                    lvOfficerPorts.Items(i).Selected = True
                    CurrentOfficer = lvItem
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(20)
                    lvItem.Selected = True
                    'ListView1.Select()
                    Exit For
                End If
            Next
            Application.DoEvents()
            Debug.WriteLine(Me.lvOfficerPorts.SelectedItems.Count)
            Me.ListView1_Click(Nothing, Nothing)

        End If
    End Sub

    Private Sub cmdClosest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClosest.Click
        If Me.myCollection Is Nothing Then Exit Sub
        MedscreenCommonGui.menus.CollectionMenu.FindNearestOfficer(myCollection)
        'Dim oColl As New Collection()
        'oColl.Add(Me.myPort.Identity)
        'oColl.Add(Me.myPort.CountryId)
        'Dim strXML As String = CConnection.PackageStringList("lib_officer.ClosestOfficerstoPortXML", oColl)
        'If strXML Is Nothing OrElse strXML.Length < 20 Then
        '    MsgBox("This port doesn't have any geographical details Longitude or Latitude, so this information can't be calculated.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
        '    Exit Sub
        'End If
        'Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "Nearest.xslt", 0)
        'Medscreen.ShowHtml(strHTML)

    End Sub
End Class

