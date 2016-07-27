'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 09:58:22+00 $
'$Log: frmDiary.vb,v $
'Revision 1.0  2005-11-29 09:58:22+00  taylor
'Staging check in
'
'Revision 1.2  2005-06-16 17:36:21+01  taylor
'Milestone check in
'
'Revision 1.1  2004-05-04 10:09:07+01  taylor
'Added header check in
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports MedscreenLib.Constants
Imports MedscreenLib
Imports Intranet.intranet
Imports Intranet.intranet.Support
Imports MedscreenCommonGui
Imports MedscreenCommonGui.CommonForms
Imports System.Windows.Forms
Imports System.Windows
Imports System.Reflection
Imports System.Drawing
Imports Microsoft.Office.Interop


''' -----------------------------------------------------------------------------
''' Project  : CCTool
''' Class  : frmDiary
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form designed to manage collections, displaying timing etc.
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action>
''' $Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-29 09:58:22+00 $ <para/>
'''$Log: frmDiary.vb,v $
'''Revision 1.0  2005-11-29 09:58:22+00  taylor
'''Staging check in
''' <para/>
'''</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmDiary


    Inherits System.Windows.Forms.Form
#Region "Declarations"
    Private WithEvents CurrentDocument As HtmlDocument
    Dim MousePoint As System.Drawing.Point
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintPreviewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Dim Ele As HtmlElement
    Friend WithEvents EmailTextToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents lblLink As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNavigate As System.Windows.Forms.TextBox
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ListSortBox As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents FromCardiffTicketToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    'Friend WithEvents TestToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    Private myCollMenu As menus.CollectionMenu
#End Region


#Region "Public Enumerations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible Destinations, for collections to be filtered by
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum Destination
        ''' <summary>All destinations</summary>
        All
        ''' <summary>UK collections only</summary>
        UK
        ''' <summary>Overseas collections only</summary>
        Overseas
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of person using application 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum UserType
        ''' <summary>Normal User</summary>
        Normal
        ''' <summary></summary>
        ToBeAssigned
        ''' <summary>Program Manager</summary>
        ProgrammeManged
    End Enum
#End Region

#Region "Class CollInfo"

    Private Class CollInfo
        'Used to store collections found in query 
        Dim strId As String
        Dim datMod As Date

        Public Sub New(ByVal CMId As String, ByVal ModDate As Date)
            strId = CMId
            datMod = ModDate
        End Sub

        Public ReadOnly Property CMID() As String
            Get
                Return strId
            End Get
        End Property

        Public ReadOnly Property ModifictionDate() As Date
            Get
                Return datMod
            End Get
        End Property
    End Class

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Try

            Dim I As Integer
            Medscreen.LogTiming("Diary form Open")
            'This call is required by the Windows Form Designer.
            InitializeComponent()
            'CConnection.DatabaseInUse = CConnection.DatabaseInUse

            If MedscreenCommonGui.CommonForms.UserOptions Is Nothing Then
                Try
                    MedscreenCommonGui.CommonForms.UserOptions = New MedscreenCommonGui.UserDefaults.UserOptions()
                    'Get user saved defaults
                    MedscreenCommonGui.CommonForms.UserOptions.Load()
                Catch ex As Exception
                    MsgBox("Getting USer Info" & ex.ToString)
                End Try
            End If
            Try
                Medscreen.LogTiming("Diary form Finish initialise")
                'Add any initialization after the InitializeComponent() call
                LastQueryStamp = Now
                MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.CreateColumns(Me.lvDiaryList, Me.imgColHeadList)
                For I = 0 To 12
                    ColumnDirection(I) = 1
                Next

                Me.strGStatus = MedscreenCommonGui.CommonForms.UserOptions.DiaryStatusFilter
                Me.strGType = MedscreenCommonGui.CommonForms.UserOptions.DiaryTypeFilter
                If strGStatus.Trim <> "" Then
                    Me.mnuStatus.Text = "Status - " & strGStatus
                    Me.mnuStatus.Checked = True
                Else
                    Me.strGStatus = "V"
                    Me.mnuStatus.Checked = False
                    Me.mnuStatus.Text = "Status"
                End If
            Catch ex As Exception
                MsgBox("after load " & ex.ToString)
            End Try

            Me.myDestination = MedscreenCommonGui.CommonForms.UserOptions.CollDestination
            ColumnDirection(MedscreenCommonGui.CommonForms.UserOptions.DiaryColumn) = MedscreenCommonGui.CommonForms.UserOptions.DiarySortDirection

            If MedscreenCommonGui.CommonForms.UserOptions.RefreshTime > 0 Then Me.Timer1.Interval = MedscreenCommonGui.CommonForms.UserOptions.RefreshTime

            Medscreen.LogTiming("Diary form finish open")
            'see if user can cancel collections 
            Me.blnCanCancel = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_COLL_CANCEL")
            Me.myCache = New WebCache()
            Try
                Dim MyAssembly As [Assembly]
                MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
                Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
                myCollMenu = menus.CollectionMenu.GetCollectionMenu
                Dim myMenu As New MedscreenLib.DynamicContextMenu(directory & "\CollContextMenu.xml", myCollMenu)
                Me.ctxMenu = myMenu.LoadDynamicMenu()
            Catch ex As Exception
            End Try

            ListSortBox.Items.Clear()
            For Each Item As String In MedscreenCommonGui.My.Settings.MenuListCombo
                ListSortBox.Items.Add(Item)
            Next


            If MedscreenCommonGui.My.Settings.OfficerStyleSheet = "Collection2.xsl" Then
                ListSortBox.SelectedIndex = 0
            Else
                ListSortBox.SelectedIndex = 1
            End If

        Catch ex As Exception
            MsgBox("Terminate " & ex.ToString)
        End Try
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
    Friend WithEvents lvDiaryList As MedscreenCommonGui.DiaryListView
    Friend WithEvents ColCentre As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColCustomer As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColID As System.Windows.Forms.ColumnHeader
    Friend WithEvents colStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColDueDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColVessel As System.Windows.Forms.ColumnHeader
    Friend WithEvents colType As System.Windows.Forms.ColumnHeader
    'Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFilter As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFilterOn As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCustomer As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPortSite As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuClearAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuVessel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuStatus As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAmendColl As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAssignActive As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAssignComplete As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAssignPayDetails As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSend As System.Windows.Forms.MenuItem
    Friend WithEvents mnuConfirm As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCancel As System.Windows.Forms.MenuItem
    Friend WithEvents colJStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColJName As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuShowCancelled As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents colOfficer As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColPurchOrder As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuDisplay As System.Windows.Forms.MenuItem
    Friend WithEvents MnuSmallIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLargeIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDetails As System.Windows.Forms.MenuItem
    Friend WithEvents imgListSmall As System.Windows.Forms.ImageList
    Friend WithEvents imgListLrge As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAvailable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuBar1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuCollRegion As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowDestall As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowDestUK As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowDestOverseas As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCollection As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCreditCard As System.Windows.Forms.MenuItem
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents mnuUser As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuColors As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCustomise As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResend As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAmend2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOfficer2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAssigntoActive2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAssignComplete2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAssignPayDetails2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSend2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuConfirm2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCancel2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResend2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuJobsOnADay As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCreate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewCollection As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCallOut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCalloutInstructions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuResendColl As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResendConf As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResendColl2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBar2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuResendConf2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBar3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuResendCOConfirmation As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCollType As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuTypeRoutine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTypeCallout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTypeFixedSite As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTypeAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTypeWorldWide As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents cmdFind As System.Windows.Forms.Button
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents mnuHuntCollection As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuProblem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddComment As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRaiseProblem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddComment2 As System.Windows.Forms.MenuItem
    Friend WithEvents ctxEditor As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuPrintHTML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEmailHTML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents imgListStat As System.Windows.Forms.ImageList
    Friend WithEvents mnuCOComment As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFiltIssue As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuResolveProblem As System.Windows.Forms.MenuItem
    Friend WithEvents timRefresh As System.Windows.Forms.Timer
    Friend WithEvents ColProgMan As System.Windows.Forms.ColumnHeader
    Friend WithEvents mnuRecover As System.Windows.Forms.MenuItem
    Friend WithEvents mnuColumns As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents imgColHeadList As System.Windows.Forms.ImageList
    Friend WithEvents nwInfo As VbPowerPack.NotificationWindow
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents mnuDeferCollection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCollDocReceived As System.Windows.Forms.MenuItem
    Friend WithEvents cmdConfirmPaper As System.Windows.Forms.Button
    Friend WithEvents MnuViewOfficerDetails As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOfficerHead As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCollectorMaintain As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuVessels As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHuntVessel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPortsSites As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditVessel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewPort As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuView1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNormalView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOutstandingView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDoneView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuReceived As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuUnconfirmCollection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAlterDestination As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPutOnHold As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRefreshTime As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuJobCost As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewCustomerPricing As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCustomers As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCustReports As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCustCollReports As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCollectorCollList As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPortCollectionList As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEmmailOfficer As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMailOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSpecialPay As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuViewCustomerCollections As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuViewSMIDSCollections As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPastDays As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSetCollDate As System.Windows.Forms.MenuItem
    Friend WithEvents MnuInvoicing As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRecordCollPayment As System.Windows.Forms.MenuItem
    Friend WithEvents MnuEmailInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdOfficer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLockPay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCreateExternal As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuUtilities As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHuntbybarcode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHuntbyDonor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRecordCollectionPayment As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCollToSelf As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewCollRequest As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChangeCustomer As System.Windows.Forms.MenuItem
    Friend WithEvents cmdBrowserBack As System.Windows.Forms.Button
    Friend WithEvents CmdBrowserForward As System.Windows.Forms.Button
    Friend WithEvents mnuCompleteCollection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCollAgreedCosts As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScannedDocs As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDiary))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.lblLink = New System.Windows.Forms.Label
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.CmdBrowserForward = New System.Windows.Forms.Button
        Me.cmdBrowserBack = New System.Windows.Forms.Button
        Me.cmdConfirmPaper = New System.Windows.Forms.Button
        Me.cmdFind = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.mnuFilter = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFilterOn = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCustomer = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuMine = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuVessel = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPortSite = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDate = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAvailable = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuStatus = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuShowCancelled = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCollRegion = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuShowDestall = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuShowDestUK = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuShowDestOverseas = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCollType = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTypeRoutine = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTypeWorldWide = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTypeCallout = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTypeFixedSite = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTypeAll = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFiltIssue = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuBar1 = New System.Windows.Forms.ToolStripSeparator
        Me.mnuClearAll = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRefresh = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCollection = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCreate = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNewCollection = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCallOut = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCreateExternal = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAmend2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOfficer2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAssigntoActive2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAssignComplete2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAssignPayDetails2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuSend2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuConfirm2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCancel2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResend2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResendColl = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResendConf = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHuntCollection = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuProblem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddComment = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOfficerHead = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCollectorMaintain = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCollectorCollList = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEmmailOfficer = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCustomers = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCustReports = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCustCollReports = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuVessels = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHuntVessel = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEditVessel = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPortsSites = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNewPort = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPortCollectionList = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuUser = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuView1 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNormalView = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOutstandingView = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDoneView = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuReceived = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuJobsOnADay = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuSpecialPay = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuViewCustomerCollections = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuViewSMIDSCollections = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuUtilities = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHuntbybarcode = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHuntbyDonor = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuColors = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCustomise = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuColumns = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRefreshTime = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPastDays = New System.Windows.Forms.ToolStripMenuItem
        Me.ListSortBox = New System.Windows.Forms.ToolStripComboBox
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.imgListLrge = New System.Windows.Forms.ImageList(Me.components)
        Me.imgListSmall = New System.Windows.Forms.ImageList(Me.components)
        Me.imgListStat = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MnuViewOfficerDetails = New System.Windows.Forms.MenuItem
        Me.ctxMenu = New System.Windows.Forms.ContextMenu
        Me.mnuDisplay = New System.Windows.Forms.MenuItem
        Me.MnuSmallIcons = New System.Windows.Forms.MenuItem
        Me.mnuLargeIcons = New System.Windows.Forms.MenuItem
        Me.mnuList = New System.Windows.Forms.MenuItem
        Me.mnuDetails = New System.Windows.Forms.MenuItem
        Me.mnuBar2 = New System.Windows.Forms.ToolStripSeparator
        Me.mnuAmendColl = New System.Windows.Forms.MenuItem
        Me.mnuRecover = New System.Windows.Forms.MenuItem
        Me.mnuOfficer = New System.Windows.Forms.MenuItem
        Me.mnuAssignActive = New System.Windows.Forms.MenuItem
        Me.mnuAssignComplete = New System.Windows.Forms.MenuItem
        Me.mnuPay = New System.Windows.Forms.MenuItem
        Me.mnuAssignPayDetails = New System.Windows.Forms.MenuItem
        Me.mnuLockPay = New System.Windows.Forms.MenuItem
        Me.mnuMailOfficer = New System.Windows.Forms.MenuItem
        Me.mnuEdOfficer = New System.Windows.Forms.MenuItem
        Me.mnuSend = New System.Windows.Forms.MenuItem
        Me.mnuCalloutInstructions = New System.Windows.Forms.MenuItem
        Me.mnuConfirm = New System.Windows.Forms.MenuItem
        Me.mnuUnconfirmCollection = New System.Windows.Forms.MenuItem
        Me.mnuCancel = New System.Windows.Forms.MenuItem
        Me.mnuResend = New System.Windows.Forms.MenuItem
        Me.mnuResendColl2 = New System.Windows.Forms.MenuItem
        Me.mnuResendCOConfirmation = New System.Windows.Forms.MenuItem
        Me.mnuResendConf2 = New System.Windows.Forms.MenuItem
        Me.mnuCollToSelf = New System.Windows.Forms.MenuItem
        Me.mnuViewCollRequest = New System.Windows.Forms.MenuItem
        Me.mnuCollDocReceived = New System.Windows.Forms.MenuItem
        Me.mnuDeferCollection = New System.Windows.Forms.MenuItem
        Me.mnuSetCollDate = New System.Windows.Forms.MenuItem
        Me.mnuChangeCustomer = New System.Windows.Forms.MenuItem
        Me.mnuAlterDestination = New System.Windows.Forms.MenuItem
        Me.mnuRecordCollectionPayment = New System.Windows.Forms.MenuItem
        Me.mnuCompleteCollection = New System.Windows.Forms.MenuItem
        Me.mnuPutOnHold = New System.Windows.Forms.MenuItem
        Me.mnuCollAgreedCosts = New System.Windows.Forms.MenuItem
        Me.mnuBar3 = New System.Windows.Forms.ToolStripSeparator
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuJobCost = New System.Windows.Forms.MenuItem
        Me.MnuInvoicing = New System.Windows.Forms.MenuItem
        Me.mnuViewCustomerPricing = New System.Windows.Forms.MenuItem
        Me.mnuCreditCard = New System.Windows.Forms.MenuItem
        Me.mnuRecordCollPayment = New System.Windows.Forms.MenuItem
        Me.MnuEmailInvoice = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuRaiseProblem2 = New System.Windows.Forms.MenuItem
        Me.mnuResolveProblem = New System.Windows.Forms.MenuItem
        Me.mnuAddComment2 = New System.Windows.Forms.MenuItem
        Me.mnuCOComment = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.ctxEditor = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EmailTextToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintPreviewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPrintHTML = New System.Windows.Forms.MenuItem
        Me.mnuEmailHTML = New System.Windows.Forms.MenuItem
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.timRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.imgColHeadList = New System.Windows.Forms.ImageList(Me.components)
        Me.nwInfo = New VbPowerPack.NotificationWindow(Me.components)
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.mnuScannedDocs = New System.Windows.Forms.MenuItem
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtNavigate = New System.Windows.Forms.TextBox
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem
        Me.lvDiaryList = New MedscreenCommonGui.DiaryListView
        Me.ColID = New System.Windows.Forms.ColumnHeader
        Me.ColCustomer = New System.Windows.Forms.ColumnHeader
        Me.ColCentre = New System.Windows.Forms.ColumnHeader
        Me.ColVessel = New System.Windows.Forms.ColumnHeader
        Me.colStatus = New System.Windows.Forms.ColumnHeader
        Me.ColDueDate = New System.Windows.Forms.ColumnHeader
        Me.colType = New System.Windows.Forms.ColumnHeader
        Me.colJStatus = New System.Windows.Forms.ColumnHeader
        Me.ColJName = New System.Windows.Forms.ColumnHeader
        Me.colOfficer = New System.Windows.Forms.ColumnHeader
        Me.ColPurchOrder = New System.Windows.Forms.ColumnHeader
        Me.ColProgMan = New System.Windows.Forms.ColumnHeader
        Me.FromCardiffTicketToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.ctxEditor.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.lblLink)
        Me.Panel1.Controls.Add(Me.cmdRefresh)
        Me.Panel1.Controls.Add(Me.CmdBrowserForward)
        Me.Panel1.Controls.Add(Me.cmdBrowserBack)
        Me.Panel1.Controls.Add(Me.cmdConfirmPaper)
        Me.Panel1.Controls.Add(Me.cmdFind)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 713)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1113, 40)
        Me.Panel1.TabIndex = 0
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(348, 9)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(100, 23)
        Me.ProgressBar1.TabIndex = 19
        '
        'lblLink
        '
        Me.lblLink.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLink.AutoEllipsis = True
        Me.lblLink.Location = New System.Drawing.Point(454, 9)
        Me.lblLink.Name = "lblLink"
        Me.lblLink.Size = New System.Drawing.Size(465, 23)
        Me.lblLink.TabIndex = 18
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Image = CType(resources.GetObject("cmdRefresh.Image"), System.Drawing.Image)
        Me.cmdRefresh.Location = New System.Drawing.Point(309, 9)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(32, 23)
        Me.cmdRefresh.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.cmdRefresh, "Refresh")
        '
        'CmdBrowserForward
        '
        Me.CmdBrowserForward.Image = CType(resources.GetObject("CmdBrowserForward.Image"), System.Drawing.Image)
        Me.CmdBrowserForward.Location = New System.Drawing.Point(264, 8)
        Me.CmdBrowserForward.Name = "CmdBrowserForward"
        Me.CmdBrowserForward.Size = New System.Drawing.Size(32, 23)
        Me.CmdBrowserForward.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.CmdBrowserForward, "Browser Forward")
        '
        'cmdBrowserBack
        '
        Me.cmdBrowserBack.Image = CType(resources.GetObject("cmdBrowserBack.Image"), System.Drawing.Image)
        Me.cmdBrowserBack.Location = New System.Drawing.Point(216, 8)
        Me.cmdBrowserBack.Name = "cmdBrowserBack"
        Me.cmdBrowserBack.Size = New System.Drawing.Size(32, 23)
        Me.cmdBrowserBack.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.cmdBrowserBack, "Browser Back ")
        '
        'cmdConfirmPaper
        '
        Me.cmdConfirmPaper.Image = CType(resources.GetObject("cmdConfirmPaper.Image"), System.Drawing.Image)
        Me.cmdConfirmPaper.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdConfirmPaper.Location = New System.Drawing.Point(104, 8)
        Me.cmdConfirmPaper.Name = "cmdConfirmPaper"
        Me.cmdConfirmPaper.Size = New System.Drawing.Size(75, 23)
        Me.cmdConfirmPaper.TabIndex = 14
        Me.cmdConfirmPaper.Text = "Paperwork"
        Me.cmdConfirmPaper.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdConfirmPaper.Visible = False
        '
        'cmdFind
        '
        Me.cmdFind.Image = CType(resources.GetObject("cmdFind.Image"), System.Drawing.Image)
        Me.cmdFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFind.Location = New System.Drawing.Point(16, 8)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.Size = New System.Drawing.Size(75, 23)
        Me.cmdFind.TabIndex = 13
        Me.cmdFind.Text = "Find"
        Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(1028, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(942, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.StatusBar1)
        Me.Panel2.Controls.Add(Me.MenuStrip1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1113, 24)
        Me.Panel2.TabIndex = 1
        '
        'StatusBar1
        '
        Me.StatusBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusBar1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar1.Location = New System.Drawing.Point(0, 24)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2, Me.StatusBarPanel3})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(1113, 35)
        Me.StatusBar1.TabIndex = 5
        Me.StatusBar1.Text = "StatusBar1"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 190
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel2.Name = "StatusBarPanel2"
        Me.StatusBarPanel2.Width = 857
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.Name = "StatusBarPanel3"
        Me.StatusBarPanel3.Width = 50
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFilter, Me.mnuCollection, Me.mnuOfficerHead, Me.mnuCustomers, Me.mnuVessels, Me.mnuPortsSites, Me.mnuUser, Me.mnuView1, Me.mnuUtilities, Me.mnuOptions, Me.mnuHelp})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1113, 24)
        Me.MenuStrip1.TabIndex = 6
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'mnuFilter
        '
        Me.mnuFilter.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFilterOn})
        Me.mnuFilter.Name = "mnuFilter"
        Me.mnuFilter.Size = New System.Drawing.Size(43, 20)
        Me.mnuFilter.Text = "&Filter"
        '
        'mnuFilterOn
        '
        Me.mnuFilterOn.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCustomer, Me.mnuMine, Me.mnuVessel, Me.mnuPortSite, Me.mnuDate, Me.mnuAvailable, Me.mnuStatus, Me.mnuShowCancelled, Me.mnuCollRegion, Me.mnuCollType, Me.mnuFiltIssue, Me.mnuBar1, Me.mnuClearAll, Me.mnuRefresh})
        Me.mnuFilterOn.Name = "mnuFilterOn"
        Me.mnuFilterOn.Size = New System.Drawing.Size(126, 22)
        Me.mnuFilterOn.Text = "Filter On"
        '
        'mnuCustomer
        '
        Me.mnuCustomer.Name = "mnuCustomer"
        Me.mnuCustomer.Size = New System.Drawing.Size(182, 22)
        Me.mnuCustomer.Text = "&Customer"
        '
        'mnuMine
        '
        Me.mnuMine.Name = "mnuMine"
        Me.mnuMine.Size = New System.Drawing.Size(182, 22)
        Me.mnuMine.Text = "&My Customers"
        '
        'mnuVessel
        '
        Me.mnuVessel.Name = "mnuVessel"
        Me.mnuVessel.Size = New System.Drawing.Size(182, 22)
        Me.mnuVessel.Text = "&Vessel"
        '
        'mnuPortSite
        '
        Me.mnuPortSite.Name = "mnuPortSite"
        Me.mnuPortSite.Size = New System.Drawing.Size(182, 22)
        Me.mnuPortSite.Text = "&Port/Site"
        '
        'mnuDate
        '
        Me.mnuDate.Name = "mnuDate"
        Me.mnuDate.Size = New System.Drawing.Size(182, 22)
        Me.mnuDate.Text = "&Date "
        '
        'mnuAvailable
        '
        Me.mnuAvailable.Name = "mnuAvailable"
        Me.mnuAvailable.Size = New System.Drawing.Size(182, 22)
        Me.mnuAvailable.Text = "&Available Collections"
        '
        'mnuStatus
        '
        Me.mnuStatus.Name = "mnuStatus"
        Me.mnuStatus.Size = New System.Drawing.Size(182, 22)
        Me.mnuStatus.Text = "&Status"
        '
        'mnuShowCancelled
        '
        Me.mnuShowCancelled.Name = "mnuShowCancelled"
        Me.mnuShowCancelled.Size = New System.Drawing.Size(182, 22)
        Me.mnuShowCancelled.Text = "Show Ca&ncelled"
        '
        'mnuCollRegion
        '
        Me.mnuCollRegion.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuShowDestall, Me.mnuShowDestUK, Me.mnuShowDestOverseas})
        Me.mnuCollRegion.Name = "mnuCollRegion"
        Me.mnuCollRegion.Size = New System.Drawing.Size(182, 22)
        Me.mnuCollRegion.Text = "Collection Region"
        '
        'mnuShowDestall
        '
        Me.mnuShowDestall.Checked = True
        Me.mnuShowDestall.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuShowDestall.Image = CType(resources.GetObject("mnuShowDestall.Image"), System.Drawing.Image)
        Me.mnuShowDestall.Name = "mnuShowDestall"
        Me.mnuShowDestall.Size = New System.Drawing.Size(185, 22)
        Me.mnuShowDestall.Text = "Show &All"
        '
        'mnuShowDestUK
        '
        Me.mnuShowDestUK.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.mnuShowDestUK.Name = "mnuShowDestUK"
        Me.mnuShowDestUK.Size = New System.Drawing.Size(185, 22)
        Me.mnuShowDestUK.Text = "Show &UK Only"
        '
        'mnuShowDestOverseas
        '
        Me.mnuShowDestOverseas.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.mnuShowDestOverseas.Name = "mnuShowDestOverseas"
        Me.mnuShowDestOverseas.Size = New System.Drawing.Size(185, 22)
        Me.mnuShowDestOverseas.Text = "Show &Overseas Only"
        '
        'mnuCollType
        '
        Me.mnuCollType.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuTypeRoutine, Me.mnuTypeWorldWide, Me.mnuTypeCallout, Me.mnuTypeFixedSite, Me.mnuTypeAll})
        Me.mnuCollType.Name = "mnuCollType"
        Me.mnuCollType.Size = New System.Drawing.Size(182, 22)
        Me.mnuCollType.Text = "Collection &Type"
        '
        'MnuTypeRoutine
        '
        Me.MnuTypeRoutine.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.MnuTypeRoutine.Name = "MnuTypeRoutine"
        Me.MnuTypeRoutine.Size = New System.Drawing.Size(177, 22)
        Me.MnuTypeRoutine.Text = "&Routine UK"
        '
        'mnuTypeWorldWide
        '
        Me.mnuTypeWorldWide.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.mnuTypeWorldWide.Name = "mnuTypeWorldWide"
        Me.mnuTypeWorldWide.Size = New System.Drawing.Size(177, 22)
        Me.mnuTypeWorldWide.Text = "Routine &WorldWide"
        '
        'mnuTypeCallout
        '
        Me.mnuTypeCallout.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.mnuTypeCallout.Name = "mnuTypeCallout"
        Me.mnuTypeCallout.Size = New System.Drawing.Size(177, 22)
        Me.mnuTypeCallout.Text = "&Call out "
        '
        'mnuTypeFixedSite
        '
        Me.mnuTypeFixedSite.Image = Global.MedscreenCommonGui.AdcResource.radiooff
        Me.mnuTypeFixedSite.Name = "mnuTypeFixedSite"
        Me.mnuTypeFixedSite.Size = New System.Drawing.Size(177, 22)
        Me.mnuTypeFixedSite.Text = "&Fixed Site"
        '
        'mnuTypeAll
        '
        Me.mnuTypeAll.Checked = True
        Me.mnuTypeAll.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuTypeAll.Image = Global.MedscreenCommonGui.AdcResource.radioon
        Me.mnuTypeAll.Name = "mnuTypeAll"
        Me.mnuTypeAll.Size = New System.Drawing.Size(177, 22)
        Me.mnuTypeAll.Text = "&All"
        '
        'mnuFiltIssue
        '
        Me.mnuFiltIssue.Name = "mnuFiltIssue"
        Me.mnuFiltIssue.Size = New System.Drawing.Size(182, 22)
        Me.mnuFiltIssue.Text = "With &Issues"
        '
        'mnuBar1
        '
        Me.mnuBar1.Name = "mnuBar1"
        Me.mnuBar1.Size = New System.Drawing.Size(179, 6)
        '
        'mnuClearAll
        '
        Me.mnuClearAll.Name = "mnuClearAll"
        Me.mnuClearAll.Size = New System.Drawing.Size(182, 22)
        Me.mnuClearAll.Text = "C&lear All Filters"
        '
        'mnuRefresh
        '
        Me.mnuRefresh.Name = "mnuRefresh"
        Me.mnuRefresh.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuRefresh.Size = New System.Drawing.Size(182, 22)
        Me.mnuRefresh.Text = "&Refresh"
        '
        'mnuCollection
        '
        Me.mnuCollection.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCreate, Me.mnuAmend2, Me.mnuOfficer2, Me.mnuSend2, Me.mnuConfirm2, Me.mnuCancel2, Me.mnuResend2, Me.mnuHuntCollection, Me.mnuProblem, Me.mnuAddComment})
        Me.mnuCollection.Name = "mnuCollection"
        Me.mnuCollection.Size = New System.Drawing.Size(65, 20)
        Me.mnuCollection.Text = "&Collection"
        '
        'mnuCreate
        '
        Me.mnuCreate.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewCollection, Me.mnuCallOut, Me.mnuCreateExternal, Me.FromCardiffTicketToolStripMenuItem})
        Me.mnuCreate.Name = "mnuCreate"
        Me.mnuCreate.Size = New System.Drawing.Size(187, 22)
        Me.mnuCreate.Text = "&Create"
        '
        'mnuNewCollection
        '
        Me.mnuNewCollection.Name = "mnuNewCollection"
        Me.mnuNewCollection.Size = New System.Drawing.Size(176, 22)
        Me.mnuNewCollection.Text = "&Routine"
        '
        'mnuCallOut
        '
        Me.mnuCallOut.Name = "mnuCallOut"
        Me.mnuCallOut.Size = New System.Drawing.Size(176, 22)
        Me.mnuCallOut.Text = "&Call out"
        '
        'mnuCreateExternal
        '
        Me.mnuCreateExternal.Name = "mnuCreateExternal"
        Me.mnuCreateExternal.Size = New System.Drawing.Size(176, 22)
        Me.mnuCreateExternal.Text = "E&xternal "
        '
        'mnuAmend2
        '
        Me.mnuAmend2.Name = "mnuAmend2"
        Me.mnuAmend2.Size = New System.Drawing.Size(187, 22)
        Me.mnuAmend2.Text = "&Amend"
        '
        'mnuOfficer2
        '
        Me.mnuOfficer2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAssigntoActive2, Me.mnuAssignComplete2, Me.mnuAssignPayDetails2})
        Me.mnuOfficer2.Name = "mnuOfficer2"
        Me.mnuOfficer2.Size = New System.Drawing.Size(187, 22)
        Me.mnuOfficer2.Text = "&Officer"
        '
        'mnuAssigntoActive2
        '
        Me.mnuAssigntoActive2.Name = "mnuAssigntoActive2"
        Me.mnuAssigntoActive2.Size = New System.Drawing.Size(226, 22)
        Me.mnuAssigntoActive2.Text = "&Assign to Active"
        '
        'mnuAssignComplete2
        '
        Me.mnuAssignComplete2.Name = "mnuAssignComplete2"
        Me.mnuAssignComplete2.Size = New System.Drawing.Size(226, 22)
        Me.mnuAssignComplete2.Text = "Assign to &Complete"
        '
        'mnuAssignPayDetails2
        '
        Me.mnuAssignPayDetails2.Name = "mnuAssignPayDetails2"
        Me.mnuAssignPayDetails2.Size = New System.Drawing.Size(226, 22)
        Me.mnuAssignPayDetails2.Text = "&Enter/view Pay Details (ctrl I)"
        '
        'mnuSend2
        '
        Me.mnuSend2.Name = "mnuSend2"
        Me.mnuSend2.Size = New System.Drawing.Size(187, 22)
        Me.mnuSend2.Text = "&Send"
        '
        'mnuConfirm2
        '
        Me.mnuConfirm2.Name = "mnuConfirm2"
        Me.mnuConfirm2.Size = New System.Drawing.Size(187, 22)
        Me.mnuConfirm2.Text = "&Confirm Coll Officer"
        '
        'mnuCancel2
        '
        Me.mnuCancel2.Name = "mnuCancel2"
        Me.mnuCancel2.Size = New System.Drawing.Size(187, 22)
        Me.mnuCancel2.Text = "Cance&l"
        '
        'mnuResend2
        '
        Me.mnuResend2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuResendColl, Me.mnuResendConf})
        Me.mnuResend2.Name = "mnuResend2"
        Me.mnuResend2.Size = New System.Drawing.Size(187, 22)
        Me.mnuResend2.Text = "&Resend/Print"
        '
        'mnuResendColl
        '
        Me.mnuResendColl.Name = "mnuResendColl"
        Me.mnuResendColl.Size = New System.Drawing.Size(206, 22)
        Me.mnuResendColl.Text = "&Collection"
        '
        'mnuResendConf
        '
        Me.mnuResendConf.Name = "mnuResendConf"
        Me.mnuResendConf.Size = New System.Drawing.Size(206, 22)
        Me.mnuResendConf.Text = "&Confirmation to customer"
        '
        'mnuHuntCollection
        '
        Me.mnuHuntCollection.Name = "mnuHuntCollection"
        Me.mnuHuntCollection.Size = New System.Drawing.Size(187, 22)
        Me.mnuHuntCollection.Text = "&Hunt Collection"
        '
        'mnuProblem
        '
        Me.mnuProblem.Name = "mnuProblem"
        Me.mnuProblem.Size = New System.Drawing.Size(187, 22)
        Me.mnuProblem.Text = "Raise &Problem Locally"
        '
        'mnuAddComment
        '
        Me.mnuAddComment.Name = "mnuAddComment"
        Me.mnuAddComment.Size = New System.Drawing.Size(187, 22)
        Me.mnuAddComment.Text = "Add Co&mment"
        '
        'mnuOfficerHead
        '
        Me.mnuOfficerHead.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCollectorMaintain, Me.mnuCollectorCollList, Me.mnuEmmailOfficer})
        Me.mnuOfficerHead.Name = "mnuOfficerHead"
        Me.mnuOfficerHead.Size = New System.Drawing.Size(57, 20)
        Me.mnuOfficerHead.Text = "&Officers"
        '
        'mnuCollectorMaintain
        '
        Me.mnuCollectorMaintain.Name = "mnuCollectorMaintain"
        Me.mnuCollectorMaintain.Size = New System.Drawing.Size(150, 22)
        Me.mnuCollectorMaintain.Text = "&Maintain"
        '
        'mnuCollectorCollList
        '
        Me.mnuCollectorCollList.Name = "mnuCollectorCollList"
        Me.mnuCollectorCollList.Size = New System.Drawing.Size(150, 22)
        Me.mnuCollectorCollList.Text = "&Collection List"
        '
        'mnuEmmailOfficer
        '
        Me.mnuEmmailOfficer.Name = "mnuEmmailOfficer"
        Me.mnuEmmailOfficer.Size = New System.Drawing.Size(150, 22)
        Me.mnuEmmailOfficer.Text = "&Email Officer"
        '
        'mnuCustomers
        '
        Me.mnuCustomers.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCustReports})
        Me.mnuCustomers.Name = "mnuCustomers"
        Me.mnuCustomers.Size = New System.Drawing.Size(70, 20)
        Me.mnuCustomers.Text = "Cu&stomers"
        '
        'mnuCustReports
        '
        Me.mnuCustReports.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCustCollReports})
        Me.mnuCustReports.Name = "mnuCustReports"
        Me.mnuCustReports.Size = New System.Drawing.Size(123, 22)
        Me.mnuCustReports.Text = "&Reports"
        '
        'mnuCustCollReports
        '
        Me.mnuCustCollReports.Name = "mnuCustCollReports"
        Me.mnuCustCollReports.Size = New System.Drawing.Size(172, 22)
        Me.mnuCustCollReports.Text = "&Collection Reports"
        '
        'mnuVessels
        '
        Me.mnuVessels.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHuntVessel, Me.mnuEditVessel})
        Me.mnuVessels.Name = "mnuVessels"
        Me.mnuVessels.Size = New System.Drawing.Size(54, 20)
        Me.mnuVessels.Text = "&Vessels"
        '
        'mnuHuntVessel
        '
        Me.mnuHuntVessel.Name = "mnuHuntVessel"
        Me.mnuHuntVessel.Size = New System.Drawing.Size(141, 22)
        Me.mnuHuntVessel.Text = "&Hunt Vessel"
        '
        'mnuEditVessel
        '
        Me.mnuEditVessel.Name = "mnuEditVessel"
        Me.mnuEditVessel.Size = New System.Drawing.Size(141, 22)
        Me.mnuEditVessel.Text = "&Edit Vessel"
        '
        'mnuPortsSites
        '
        Me.mnuPortsSites.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewPort, Me.mnuPortCollectionList})
        Me.mnuPortsSites.Name = "mnuPortsSites"
        Me.mnuPortsSites.Size = New System.Drawing.Size(71, 20)
        Me.mnuPortsSites.Text = "&Ports/Sites"
        '
        'mnuNewPort
        '
        Me.mnuNewPort.Name = "mnuNewPort"
        Me.mnuNewPort.Size = New System.Drawing.Size(153, 22)
        Me.mnuNewPort.Text = "&New Port"
        '
        'mnuPortCollectionList
        '
        Me.mnuPortCollectionList.Name = "mnuPortCollectionList"
        Me.mnuPortCollectionList.Size = New System.Drawing.Size(153, 22)
        Me.mnuPortCollectionList.Text = "&Collection List "
        '
        'mnuUser
        '
        Me.mnuUser.Name = "mnuUser"
        Me.mnuUser.Size = New System.Drawing.Size(41, 20)
        Me.mnuUser.Text = "&User"
        '
        'mnuView1
        '
        Me.mnuView1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNormalView, Me.mnuOutstandingView, Me.mnuDoneView, Me.mnuReceived, Me.mnuJobsOnADay, Me.mnuSpecialPay, Me.mnuViewCustomerCollections, Me.mnuViewSMIDSCollections})
        Me.mnuView1.Name = "mnuView1"
        Me.mnuView1.Size = New System.Drawing.Size(41, 20)
        Me.mnuView1.Text = "V&iew"
        '
        'mnuNormalView
        '
        Me.mnuNormalView.Checked = True
        Me.mnuNormalView.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuNormalView.Name = "mnuNormalView"
        Me.mnuNormalView.Size = New System.Drawing.Size(192, 22)
        Me.mnuNormalView.Text = "&Normal"
        '
        'mnuOutstandingView
        '
        Me.mnuOutstandingView.Name = "mnuOutstandingView"
        Me.mnuOutstandingView.Size = New System.Drawing.Size(192, 22)
        Me.mnuOutstandingView.Text = "&Outstanding"
        '
        'mnuDoneView
        '
        Me.mnuDoneView.Name = "mnuDoneView"
        Me.mnuDoneView.Size = New System.Drawing.Size(192, 22)
        Me.mnuDoneView.Text = "&Done"
        '
        'mnuReceived
        '
        Me.mnuReceived.Name = "mnuReceived"
        Me.mnuReceived.Size = New System.Drawing.Size(192, 22)
        Me.mnuReceived.Text = "&Received"
        '
        'mnuJobsOnADay
        '
        Me.mnuJobsOnADay.Name = "mnuJobsOnADay"
        Me.mnuJobsOnADay.Size = New System.Drawing.Size(192, 22)
        Me.mnuJobsOnADay.Text = "&Jobs on a Day"
        '
        'mnuSpecialPay
        '
        Me.mnuSpecialPay.Name = "mnuSpecialPay"
        Me.mnuSpecialPay.Size = New System.Drawing.Size(192, 22)
        Me.mnuSpecialPay.Text = "&Special payments"
        '
        'mnuViewCustomerCollections
        '
        Me.mnuViewCustomerCollections.Name = "mnuViewCustomerCollections"
        Me.mnuViewCustomerCollections.Size = New System.Drawing.Size(192, 22)
        Me.mnuViewCustomerCollections.Text = "&Customer's Collections"
        '
        'mnuViewSMIDSCollections
        '
        Me.mnuViewSMIDSCollections.Name = "mnuViewSMIDSCollections"
        Me.mnuViewSMIDSCollections.Size = New System.Drawing.Size(192, 22)
        Me.mnuViewSMIDSCollections.Text = "Smid's Collections"
        '
        'mnuUtilities
        '
        Me.mnuUtilities.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHuntbybarcode, Me.mnuHuntbyDonor})
        Me.mnuUtilities.Name = "mnuUtilities"
        Me.mnuUtilities.Size = New System.Drawing.Size(53, 20)
        Me.mnuUtilities.Text = "&Utilities"
        '
        'mnuHuntbybarcode
        '
        Me.mnuHuntbybarcode.Name = "mnuHuntbybarcode"
        Me.mnuHuntbybarcode.Size = New System.Drawing.Size(165, 22)
        Me.mnuHuntbybarcode.Text = "&Hunt by barcode"
        '
        'mnuHuntbyDonor
        '
        Me.mnuHuntbyDonor.Name = "mnuHuntbyDonor"
        Me.mnuHuntbyDonor.Size = New System.Drawing.Size(165, 22)
        Me.mnuHuntbyDonor.Text = "Hunt by&Donor"
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuColors, Me.mnuColumns, Me.mnuRefreshTime, Me.mnuPastDays, Me.ListSortBox})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(56, 20)
        Me.mnuOptions.Text = "&Options"
        '
        'mnuColors
        '
        Me.mnuColors.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCustomise})
        Me.mnuColors.Name = "mnuColors"
        Me.mnuColors.Size = New System.Drawing.Size(187, 22)
        Me.mnuColors.Text = "&Colours"
        '
        'mnuCustomise
        '
        Me.mnuCustomise.Name = "mnuCustomise"
        Me.mnuCustomise.Size = New System.Drawing.Size(134, 22)
        Me.mnuCustomise.Text = "&Customise"
        '
        'mnuColumns
        '
        Me.mnuColumns.Name = "mnuColumns"
        Me.mnuColumns.Size = New System.Drawing.Size(187, 22)
        Me.mnuColumns.Text = "Columns"
        '
        'mnuRefreshTime
        '
        Me.mnuRefreshTime.Name = "mnuRefreshTime"
        Me.mnuRefreshTime.Size = New System.Drawing.Size(187, 22)
        Me.mnuRefreshTime.Text = "&Refresh Time"
        '
        'mnuPastDays
        '
        Me.mnuPastDays.Name = "mnuPastDays"
        Me.mnuPastDays.Size = New System.Drawing.Size(187, 22)
        Me.mnuPastDays.Text = "&Query Days into Past"
        '
        'ListSortBox
        '
        Me.ListSortBox.DropDownWidth = 200
        Me.ListSortBox.Items.AddRange(New Object() {"Menu List by Customer", "Menu List by Collection Date"})
        Me.ListSortBox.Name = "ListSortBox"
        Me.ListSortBox.Size = New System.Drawing.Size(121, 21)
        '
        'mnuHelp
        '
        Me.mnuHelp.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(40, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'imgListLrge
        '
        Me.imgListLrge.ImageStream = CType(resources.GetObject("imgListLrge.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgListLrge.TransparentColor = System.Drawing.Color.Transparent
        Me.imgListLrge.Images.SetKeyName(0, "PeePot0.ico")
        Me.imgListLrge.Images.SetKeyName(1, "PeePot1.ico")
        Me.imgListLrge.Images.SetKeyName(2, "PeePot2.ico")
        Me.imgListLrge.Images.SetKeyName(3, "PeePot3.ico")
        Me.imgListLrge.Images.SetKeyName(4, "PeePot4.ico")
        Me.imgListLrge.Images.SetKeyName(5, "PeePot5.ico")
        Me.imgListLrge.Images.SetKeyName(6, "PeePot.ico")
        '
        'imgListSmall
        '
        Me.imgListSmall.ImageStream = CType(resources.GetObject("imgListSmall.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgListSmall.TransparentColor = System.Drawing.Color.Transparent
        Me.imgListSmall.Images.SetKeyName(0, "PeePot0.ico")
        Me.imgListSmall.Images.SetKeyName(1, "PeePot1.ico")
        Me.imgListSmall.Images.SetKeyName(2, "PeePot2.ico")
        Me.imgListSmall.Images.SetKeyName(3, "PeePot3.ico")
        Me.imgListSmall.Images.SetKeyName(4, "PeePot4.ico")
        Me.imgListSmall.Images.SetKeyName(5, "PeePot5.ico")
        Me.imgListSmall.Images.SetKeyName(6, "PeePot.ico")
        '
        'imgListStat
        '
        Me.imgListStat.ImageStream = CType(resources.GetObject("imgListStat.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgListStat.TransparentColor = System.Drawing.Color.Transparent
        Me.imgListStat.Images.SetKeyName(0, "Tick16.ico")
        Me.imgListStat.Images.SetKeyName(1, "Cross16.ico")
        Me.imgListStat.Images.SetKeyName(2, "radioon.bmp")
        Me.imgListStat.Images.SetKeyName(3, "radiooff.bmp")
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = -1
        Me.MenuItem1.Text = ""
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = -1
        Me.MenuItem2.Text = ""
        '
        'MnuViewOfficerDetails
        '
        Me.MnuViewOfficerDetails.Index = 3
        Me.MnuViewOfficerDetails.Text = "&View Contact Details"
        '
        'ctxMenu
        '
        '
        'mnuDisplay
        '
        Me.mnuDisplay.Index = -1
        Me.mnuDisplay.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuSmallIcons, Me.mnuLargeIcons, Me.mnuList, Me.mnuDetails})
        Me.mnuDisplay.Text = "Display "
        '
        'MnuSmallIcons
        '
        Me.MnuSmallIcons.Index = 0
        Me.MnuSmallIcons.Text = "&Small Icons"
        '
        'mnuLargeIcons
        '
        Me.mnuLargeIcons.Index = 1
        Me.mnuLargeIcons.Text = "&Large Icons"
        '
        'mnuList
        '
        Me.mnuList.Index = 2
        Me.mnuList.Text = "L&ist"
        '
        'mnuDetails
        '
        Me.mnuDetails.Index = 3
        Me.mnuDetails.Text = "&Details"
        '
        'mnuBar2
        '
        Me.mnuBar2.Name = "mnuBar2"
        Me.mnuBar2.Size = New System.Drawing.Size(6, 6)
        '
        'mnuAmendColl
        '
        Me.mnuAmendColl.Index = -1
        Me.mnuAmendColl.Text = "&Amend"
        '
        'mnuRecover
        '
        Me.mnuRecover.Index = -1
        Me.mnuRecover.Text = "&Recover Collection"
        Me.mnuRecover.Visible = False
        '
        'mnuOfficer
        '
        Me.mnuOfficer.Index = -1
        Me.mnuOfficer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAssignActive, Me.mnuAssignComplete, Me.mnuPay, Me.MnuViewOfficerDetails, Me.mnuMailOfficer, Me.mnuEdOfficer})
        Me.mnuOfficer.Text = "&Officer"
        '
        'mnuAssignActive
        '
        Me.mnuAssignActive.Index = 0
        Me.mnuAssignActive.Text = "Assign to &Active"
        '
        'mnuAssignComplete
        '
        Me.mnuAssignComplete.Index = 1
        Me.mnuAssignComplete.Text = "Assign to &Complete"
        '
        'mnuPay
        '
        Me.mnuPay.Index = 2
        Me.mnuPay.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAssignPayDetails, Me.mnuLockPay})
        Me.mnuPay.Text = "&Pay "
        '
        'mnuAssignPayDetails
        '
        Me.mnuAssignPayDetails.Index = 0
        Me.mnuAssignPayDetails.Text = "&Enter/view Pay Details (ctrl I)"
        '
        'mnuLockPay
        '
        Me.mnuLockPay.Index = 1
        Me.mnuLockPay.Text = "&Lock Pay"
        '
        'mnuMailOfficer
        '
        Me.mnuMailOfficer.Index = 4
        Me.mnuMailOfficer.Text = "Email &Officer re collection"
        Me.mnuMailOfficer.Visible = False
        '
        'mnuEdOfficer
        '
        Me.mnuEdOfficer.Index = 5
        Me.mnuEdOfficer.Text = "Edit Collecting Officer &Details"
        '
        'mnuSend
        '
        Me.mnuSend.Index = -1
        Me.mnuSend.Text = "&Send"
        '
        'mnuCalloutInstructions
        '
        Me.mnuCalloutInstructions.Index = -1
        Me.mnuCalloutInstructions.Text = "&Send Call out Instructions "
        Me.mnuCalloutInstructions.Visible = False
        '
        'mnuConfirm
        '
        Me.mnuConfirm.Index = -1
        Me.mnuConfirm.Text = "&Confirm Coll Officer"
        '
        'mnuUnconfirmCollection
        '
        Me.mnuUnconfirmCollection.Index = -1
        Me.mnuUnconfirmCollection.Text = "&Unconfirm Collection"
        '
        'mnuCancel
        '
        Me.mnuCancel.Index = -1
        Me.mnuCancel.Text = "Cance&l"
        '
        'mnuResend
        '
        Me.mnuResend.Index = -1
        Me.mnuResend.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuResendColl2, Me.mnuResendCOConfirmation, Me.mnuResendConf2, Me.mnuCollToSelf, Me.mnuViewCollRequest})
        Me.mnuResend.Text = "&Resend/Print"
        '
        'mnuResendColl2
        '
        Me.mnuResendColl2.Index = 0
        Me.mnuResendColl2.Text = "&Collection"
        '
        'mnuResendCOConfirmation
        '
        Me.mnuResendCOConfirmation.Index = 1
        Me.mnuResendCOConfirmation.Text = "&Officer Confirmation"
        '
        'mnuResendConf2
        '
        Me.mnuResendConf2.Index = 2
        Me.mnuResendConf2.Text = "C&ustomer Confirmation"
        '
        'mnuCollToSelf
        '
        Me.mnuCollToSelf.Index = 3
        Me.mnuCollToSelf.Text = "Collection Request (To &Self)"
        '
        'mnuViewCollRequest
        '
        Me.mnuViewCollRequest.Index = 4
        Me.mnuViewCollRequest.Text = "&View Collection Request"
        '
        'mnuCollDocReceived
        '
        Me.mnuCollDocReceived.Index = -1
        Me.mnuCollDocReceived.Text = "Collec&tion Documentation received"
        '
        'mnuDeferCollection
        '
        Me.mnuDeferCollection.Index = -1
        Me.mnuDeferCollection.Text = "&Defer Collection"
        '
        'mnuSetCollDate
        '
        Me.mnuSetCollDate.Index = -1
        Me.mnuSetCollDate.Text = "Set Collection Date"
        '
        'mnuChangeCustomer
        '
        Me.mnuChangeCustomer.Index = -1
        Me.mnuChangeCustomer.Text = "Change Customer"
        '
        'mnuAlterDestination
        '
        Me.mnuAlterDestination.Index = -1
        Me.mnuAlterDestination.Text = "C&hange destination"
        '
        'mnuRecordCollectionPayment
        '
        Me.mnuRecordCollectionPayment.Index = -1
        Me.mnuRecordCollectionPayment.Text = "Record Collection Pa&yment"
        '
        'mnuCompleteCollection
        '
        Me.mnuCompleteCollection.Index = -1
        Me.mnuCompleteCollection.Text = "Complete Collection"
        '
        'mnuPutOnHold
        '
        Me.mnuPutOnHold.Index = -1
        Me.mnuPutOnHold.Text = "Put Collection on Hol&d"
        '
        'mnuCollAgreedCosts
        '
        Me.mnuCollAgreedCosts.Index = -1
        Me.mnuCollAgreedCosts.Text = "Set a&greed costs"
        '
        'mnuBar3
        '
        Me.mnuBar3.Name = "mnuBar3"
        Me.mnuBar3.Size = New System.Drawing.Size(6, 6)
        '
        'mnuView
        '
        Me.mnuView.Index = -1
        Me.mnuView.Text = "&View Workflow"
        '
        'mnuJobCost
        '
        Me.mnuJobCost.Index = -1
        Me.mnuJobCost.Text = "View Jo&b Cost"
        '
        'MnuInvoicing
        '
        Me.MnuInvoicing.Index = -1
        Me.MnuInvoicing.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewCustomerPricing, Me.mnuCreditCard, Me.mnuRecordCollPayment, Me.MnuEmailInvoice})
        Me.MnuInvoicing.Text = "&Invoicing And Pricing"
        '
        'mnuViewCustomerPricing
        '
        Me.mnuViewCustomerPricing.Index = 0
        Me.mnuViewCustomerPricing.Text = "View Customer &Pricing"
        '
        'mnuCreditCard
        '
        Me.mnuCreditCard.Index = 1
        Me.mnuCreditCard.Text = "Credit Card &Ticket Copy"
        '
        'mnuRecordCollPayment
        '
        Me.mnuRecordCollPayment.Index = 2
        Me.mnuRecordCollPayment.Text = "&Record Collection Payment"
        Me.mnuRecordCollPayment.Visible = False
        '
        'MnuEmailInvoice
        '
        Me.MnuEmailInvoice.Index = 3
        Me.MnuEmailInvoice.Text = "&Email Invoice"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = -1
        Me.MenuItem3.Text = "-"
        '
        'mnuRaiseProblem2
        '
        Me.mnuRaiseProblem2.Index = -1
        Me.mnuRaiseProblem2.Text = "Raise &Problem Locally"
        '
        'mnuResolveProblem
        '
        Me.mnuResolveProblem.Index = -1
        Me.mnuResolveProblem.Text = "Reso&lve Problem "
        '
        'mnuAddComment2
        '
        Me.mnuAddComment2.Index = -1
        Me.mnuAddComment2.Text = "Add Com&ment"
        '
        'mnuCOComment
        '
        Me.mnuCOComment.Index = -1
        Me.mnuCOComment.Text = "Raise Problem With C/&O"
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 24)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(1113, 3)
        Me.Splitter1.TabIndex = 3
        Me.Splitter1.TabStop = False
        '
        'ctxEditor
        '
        Me.ctxEditor.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EmailTextToolStripMenuItem, Me.PrintToolStripMenuItem, Me.PrintPreviewToolStripMenuItem})
        Me.ctxEditor.Name = "ctxEditor"
        Me.ctxEditor.Size = New System.Drawing.Size(149, 70)
        '
        'EmailTextToolStripMenuItem
        '
        Me.EmailTextToolStripMenuItem.Name = "EmailTextToolStripMenuItem"
        Me.EmailTextToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.EmailTextToolStripMenuItem.Text = "Email Text"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.PrintToolStripMenuItem.Text = "Print "
        '
        'PrintPreviewToolStripMenuItem
        '
        Me.PrintPreviewToolStripMenuItem.Name = "PrintPreviewToolStripMenuItem"
        Me.PrintPreviewToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.PrintPreviewToolStripMenuItem.Text = "Print Preview"
        '
        'mnuPrintHTML
        '
        Me.mnuPrintHTML.Index = -1
        Me.mnuPrintHTML.Text = "&Print Info"
        '
        'mnuEmailHTML
        '
        Me.mnuEmailHTML.Index = -1
        Me.mnuEmailHTML.Text = "&Email Info"
        '
        'Timer1
        '
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.ContextMenuStrip = Me.ctxEditor
        Me.HtmlEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.HelpProvider1.SetHelpKeyword(Me.HtmlEditor1, "Information Panel")
        Me.HelpProvider1.SetHelpNavigator(Me.HtmlEditor1, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.HtmlEditor1.IsWebBrowserContextMenuEnabled = False
        Me.HtmlEditor1.Location = New System.Drawing.Point(0, 227)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        Me.HelpProvider1.SetShowHelp(Me.HtmlEditor1, True)
        Me.HtmlEditor1.Size = New System.Drawing.Size(1113, 486)
        Me.HtmlEditor1.TabIndex = 7
        '
        'timRefresh
        '
        Me.timRefresh.Enabled = True
        Me.timRefresh.Interval = 3000000
        '
        'imgColHeadList
        '
        Me.imgColHeadList.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgColHeadList.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgColHeadList.TransparentColor = System.Drawing.Color.Transparent
        '
        'nwInfo
        '
        Me.nwInfo.Blend = New VbPowerPack.BlendFill(VbPowerPack.BlendStyle.Vertical, System.Drawing.SystemColors.InactiveCaption, System.Drawing.SystemColors.Window)
        Me.nwInfo.DefaultText = Nothing
        Me.nwInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nwInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.nwInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.nwInfo.ShowStyle = VbPowerPack.NotificationShowStyle.Slide
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter2.Location = New System.Drawing.Point(0, 171)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(1113, 3)
        Me.Splitter2.TabIndex = 4
        Me.Splitter2.TabStop = False
        '
        'mnuScannedDocs
        '
        Me.mnuScannedDocs.Index = -1
        Me.mnuScannedDocs.Text = "&Look for Scanned Documents"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.StatusStrip1)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.txtNavigate)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 174)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1113, 53)
        Me.Panel3.TabIndex = 6
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(9, 24)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(239, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(111, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BorderStyle = System.Windows.Forms.Border3DStyle.RaisedOuter
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(111, 17)
        Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "address"
        '
        'txtNavigate
        '
        Me.txtNavigate.Location = New System.Drawing.Point(119, 1)
        Me.txtNavigate.Name = "txtNavigate"
        Me.txtNavigate.Size = New System.Drawing.Size(562, 20)
        Me.txtNavigate.TabIndex = 0
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(152, 22)
        Me.ToolStripMenuItem2.Text = "1"
        '
        'lvDiaryList
        '
        Me.lvDiaryList.AllowColumnReorder = True
        Me.lvDiaryList.BackColor = System.Drawing.SystemColors.Window
        Me.lvDiaryList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvDiaryList.CheckBoxes = True
        Me.lvDiaryList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColID, Me.ColCustomer, Me.ColCentre, Me.ColVessel, Me.colStatus, Me.ColDueDate, Me.colType, Me.colJStatus, Me.ColJName, Me.colOfficer, Me.ColPurchOrder, Me.ColProgMan})
        Me.lvDiaryList.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvDiaryList.FullRowSelect = True
        Me.lvDiaryList.HideSelection = False
        Me.lvDiaryList.LargeImageList = Me.imgListLrge
        Me.lvDiaryList.Location = New System.Drawing.Point(0, 27)
        Me.lvDiaryList.Name = "lvDiaryList"
        Me.lvDiaryList.Size = New System.Drawing.Size(1113, 144)
        Me.lvDiaryList.SmallImageList = Me.imgListSmall
        Me.lvDiaryList.StateImageList = Me.imgListStat
        Me.lvDiaryList.TabIndex = 2
        Me.lvDiaryList.UseCompatibleStateImageBehavior = False
        Me.lvDiaryList.View = System.Windows.Forms.View.Details
        '
        'ColID
        '
        Me.ColID.Text = "CM Job Number"
        Me.ColID.Width = 100
        '
        'ColCustomer
        '
        Me.ColCustomer.Text = "Customer"
        Me.ColCustomer.Width = 120
        '
        'ColCentre
        '
        Me.ColCentre.Text = "Site/Port"
        Me.ColCentre.Width = 120
        '
        'ColVessel
        '
        Me.ColVessel.Text = "Vessel"
        Me.ColVessel.Width = 120
        '
        'colStatus
        '
        Me.colStatus.Text = "Status"
        Me.colStatus.Width = 120
        '
        'ColDueDate
        '
        Me.ColDueDate.Text = "DueDate"
        Me.ColDueDate.Width = 120
        '
        'colType
        '
        Me.colType.Text = "Collection Type"
        '
        'colJStatus
        '
        Me.colJStatus.Text = "Job Status"
        Me.colJStatus.Width = 100
        '
        'ColJName
        '
        Me.ColJName.Text = "Job Name"
        '
        'colOfficer
        '
        Me.colOfficer.Text = "Officer"
        Me.colOfficer.Width = 100
        '
        'ColPurchOrder
        '
        Me.ColPurchOrder.Text = "PO Number"
        Me.ColPurchOrder.Width = 100
        '
        'ColProgMan
        '
        Me.ColProgMan.Text = "Prog Manger"
        Me.ColProgMan.Width = 100
        '
        'FromCardiffTicketToolStripMenuItem
        '
        Me.FromCardiffTicketToolStripMenuItem.Name = "FromCardiffTicketToolStripMenuItem"
        Me.FromCardiffTicketToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.FromCardiffTicketToolStripMenuItem.Text = "From Cardiff &Ticket"
        '
        'frmDiary
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1113, 753)
        Me.Controls.Add(Me.HtmlEditor1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Splitter2)
        Me.Controls.Add(Me.lvDiaryList)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.HelpProvider1.SetHelpKeyword(Me, "Collection Diary")
        Me.HelpProvider1.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmDiary"
        Me.HelpProvider1.SetShowHelp(Me, True)
        Me.Text = "Diary"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ctxEditor.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declarations"
    Private ColumnDirection(12) As Integer
    Private MyUserType As UserType

    Private ListButton As MouseButtons
    Private LastQueryStamp As Date

    Private blnStopRefresh As Boolean = False
    Dim myCentre As Intranet.FixedSiteBookings.cCentre
    Private strCustomer As String = ""
    Private strGPort As String = ""
    Private strGStatus As String = ""
    Private strGVessel As String = ""
    Private strGType As String = ""
    Private blnIssue As Boolean = False
    Private blnShowCancelled As Boolean = False
    Private blnDataLoaded As Boolean = False
    Private myDestination As Destination = Destination.All
    Private dtLoginTimestamp As Date = Now.AddMinutes(-30)

    Private datFrom As Date
    Private datTo As Date
    Friend lvItemCurrent As ListViewItems.ListViewCMDiaryItem
    'Private oleConn As New OleDb.OleDbConnection()
    Private strCollector As String = ""
    Private myCollection As Intranet.intranet.jobs.CMJob
    Private myUser As MedscreenLib.personnel.UserPerson      'User definition class
    Private myOfficer As Intranet.intranet.jobs.CollectingOfficer
    Private strGUser As String = ""

    Private DS As Intranet.DSCollection
    Private dr As Intranet.DSCollection.CollectionRow
    Private myColls As Intranet.intranet.jobs.CMCollection

    Private blnEditorOther As Boolean = False
    Private myCache As WebCache

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the title bar of the form
    ''' </summary>
    ''' <remarks>The title bar shows the status values filtered on and any destination filtering
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetTitleBar()
        Me.Text = "Diary"
        If Me.strGStatus.Trim.Length > 0 Then Me.Text += " - Filter by status " & strGStatus

        If myDestination <> Destination.All Then
            Me.Text += " - " & myDestination.ToString()
        End If
        If Me.strGUser.Trim.Length > 0 Then
            Me.Text += " - " & strGUser & " customers"
        End If
    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current login time stamp, this is part of the ensure current user is logged in code
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LoginTimeStamp() As Date
        Get
            Return Me.dtLoginTimestamp
        End Get
        Set(ByVal Value As Date)
            dtLoginTimestamp = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This form can be used to show a collecting officer's diary, if this is set then it will be in this mode
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Property CollOfficer() As String
        Get
            Return strCollector
        End Get
        Set(ByVal Value As String)
            strCollector = Value
            If Not strCollector Is Nothing Then
                If strCollector.Trim.Length > 0 Then
                    myOfficer = cSupport.CollectingOfficers.Item(strCollector.Trim)
                End If
            End If
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current user of the application
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property User() As MedscreenLib.personnel.UserPerson
        Set(ByVal Value As MedscreenLib.personnel.UserPerson)
            myUser = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Diary for a particular centre
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Property Centre() As Intranet.FixedSiteBookings.cCentre
        Get
            Return myCentre
        End Get
        Set(ByVal Value As Intranet.FixedSiteBookings.cCentre)
            myCentre = Value
            'If Not Value Is Nothing Then
            'End If
        End Set
    End Property

    'Get Collection details

#End Region
#End Region

#Region "Private Instance"
#Region "Functions"
    'Build an item for the list view, most of the work goes on in the constructor of the 
    'ListViewDiaryItem which is an Enhanced ListViewItem, See details on this private class
    Private Overloads Function BuildItem(ByVal objColl As Intranet.intranet.jobs.CMJob) As ListViewItems.ListViewDiaryItem
        Dim Item As ListViewItems.ListViewDiaryItem = Nothing
        If objColl Is Nothing Then
            Return Nothing
            Exit Function
        End If
        Try
            With objColl
                Item = New ListViewItems.ListViewDiaryItem(objColl)
                Item.UseItemStyleForSubItems = False
            End With
        Catch ex As Exception
        End Try
        Return Item

    End Function

    Private Overloads Function BuildItem(ByVal cmJobNumber As String) As ListViewItems.ListViewCMDiaryItem
        Dim Item As ListViewItems.ListViewCMDiaryItem = Nothing
        If cmJobNumber.Trim.Length = 0 Then
            Return Nothing
            Exit Function
        End If
        Try

            Item = New ListViewItems.ListViewCMDiaryItem(cmJobNumber)
            Item.UseItemStyleForSubItems = False

        Catch ex As Exception
        End Try
        Return Item

    End Function


    'Overload for dealing with collecting Officer's holidays
    Private Overloads Function BuildItem(ByVal objHol As Intranet.intranet.jobs.CollectorHoliday, Optional ByVal blnStart As Boolean = True) As ListViewItems.ListViewDiaryItem
        Dim Item As ListViewItems.ListViewDiaryItem
        If objHol Is Nothing Then
            Return Nothing
            Exit Function
        End If
        With objHol
            Item = New ListViewItems.ListViewDiaryItem("Holiday", _
                            "H", " - ", " - ", .CollectorId)
            Item.SubItems.Add(" .")
            Item.SubItems.Add(" .")
            Item.Collection = Nothing

            Item.Port = " ."
            Item.CollDate = .HolidayStart

            Item.SubItems.Add("")
            Item.ForeColor = System.Drawing.SystemColors.WindowText
            Dim csItem As ListViewItems.CollStatusSubItem = New ListViewItems.CollStatusSubItem("H", "", MedscreenCommonGui.CommonForms.UserOptions.DiaryColours)
            Item.SubItems.Add(csItem)
            Item.ForeColor = csItem.ForeColour
            Item.BackColor = csItem.BackColour
            If blnStart Then
                Dim datItem As ListViewItems.SubItems.DateItem = New ListViewItems.SubItems.DateItem(.HolidayStart)

                Item.SubItems.Add(datItem)
                Item.SubItems.Add(.HolidayEnd.ToString("dd-MMM-yyyy"))
            Else
                Dim datItem As ListViewItems.SubItems.DateItem = New ListViewItems.SubItems.DateItem(.HolidayEnd)

                Item.SubItems.Add(datItem)
                Item.SubItems.Add("")
            End If


            Dim jsItem As ListViewItems.JobStatusSubItem = New ListViewItems.JobStatusSubItem(0, "")
            Item.SubItems.Add(jsItem)
            Item.SubItems.Add(" ")
            Item.SubItems.Add(.CollectorId)
            Item.SubItems.Add(" ")

        End With
        Return Item
    End Function

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get details of the collections
    ''' </summary>
    ''' <remarks>
    ''' Builds Query, then calls DoQuery
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub GetCollectionDetails()

        Dim strQuery As String

        Try
            Me.StatusBarPanel2.Text = "Loading data"
            'Set up query create basic query first then add filtering step by step

            strQuery = "Select c.identity,rowid " & _
                "from collection c" & _
                " where COLL_DATE >= to_date('" & _
                datFrom.ToString("dd-MMM-yyyy") & "','dd-Mon-yyyy') and Coll_date < to_date('" & _
                datTo.ToString("dd-MMM-yyyy") & "','dd-Mon-yyyy') "

            If strCustomer.Length > 0 Then      'Are we filtering on the customer?
                strQuery += " and c.Customer_id = '" & strCustomer.Trim.ToUpper & "' "
            End If
            If strGPort.Length > 0 Then         'Are we filtering on the port?
                strQuery += " and upper(c.Port) = '" & strGPort.Trim.ToUpper & "' "
            End If
            If strGType.Length > 0 Then         'Are we filtering on collection type
                strQuery += " and collectiontype = '" & strGType.Trim.ToUpper & "' "
            End If

            If Me.strGStatus.Length = 1 Then    'IS there a status filter and only a single status
                Me.strGStatus = "('" & Me.strGStatus & "')"
            End If
            If Me.strGStatus.Length > 0 Then

                strQuery += " and c.Status in " & strGStatus.Trim.ToUpper & " "
            Else
                If Not Me.blnShowCancelled Then 'Are we showing cancelled collections
                    strQuery += " and c.Status not in('X','S') "
                Else
                    strQuery += " and c.Status not in('S') "
                End If
            End If
            If Me.strGVessel.Length > 0 Then    'Is there a vessel filter
                strQuery += " and c.Vessel_name = '" & strGVessel.Trim.ToUpper & "' "
            End If
            If Me.strCollector.Length > 0 Then
                strQuery += " and c.Coll_officer = '" & strCollector.Trim.ToUpper & "' "
            End If
            'Deal with Programme managers
            If Me.strGUser.Length > 0 Then
                strQuery += " and c.Cc_pman = '" & strGUser.Trim.ToUpper & "' "
            End If

            'Deal with UK or Overseas
            If myDestination = Destination.Overseas Then
                strQuery += " and UK_OR_OS = 'F' "
            End If
            If myDestination = Destination.UK Then
                strQuery += " and UK_OR_OS = 'T' "
            End If

            If blnIssue Then
                strQuery += " and cancellation_status = 'P'"
            End If

            strQuery += " order by coll_date desc"
            'Finally finished 
            DoQuery(strQuery)   'Do actual query and fill list view 
            LastQueryStamp = Now

            Dim Item As ListViewItems.ListViewDiaryItem
            If Me.CollOfficer.Trim.Length > 0 Then      'Is this a collecting officer's diary 
                '                                       'If so show holidays
                If Not Me.myOfficer Is Nothing Then
                    Dim objHol As Intranet.intranet.jobs.CollectorHoliday

                    For Each objHol In myOfficer.Holidays
                        Item = BuildItem(objHol)

                        If Not Item Is Nothing Then Me.lvDiaryList.Items.Add(Item)
                        Item = BuildItem(objHol, False)
                        If Not Item Is Nothing Then Me.lvDiaryList.Items.Add(Item)
                    Next

                End If
            End If

        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            'Me.lvDiaryList.EndUpdate()
            Me.FinishLoad()
            Me.lvDiaryList.Invalidate()
            Me.lvDiaryList.ListViewItemSorter = New Comparers.ListViewItemComparer(MedscreenCommonGui.CommonForms.UserOptions.DiaryColumn, _
                MedscreenCommonGui.CommonForms.UserOptions.DiarySortDirection)
            Me.StatusBarPanel2.Text = "Collection Data Loaded"
            SetTitleBar()                       'Show what we have filtered on 
            Application.DoEvents()
        End Try

    End Sub

    Private MyCollections As New Intranet.intranet.jobs.CMCollection(False)
    Private myCollIDS As Collection
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fills collection from query then calls fill list
    ''' </summary>
    ''' <param name="strQuery">Query to use</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DoQuery(ByVal strQuery As String)
        'Does the actual querying of the database using query constructed
        Medscreen.LogTiming("Diary form Do Query start - " & strQuery)
        Me.Cursor = Cursors.WaitCursor              'Let the world know this might take some time
        'MyCollections.Load("Lib_collection.GetDiaryRecords", strQuery)
        Dim oCMD As New OleDb.OleDbCommand(strQuery, CConnection.DbConnection)
        Dim oRead As OleDb.OleDbDataReader
        myCollIDS = New Collection()
        Try
            If CConnection.ConnOpen Then
                oRead = oCMD.ExecuteReader
                While oRead.Read
                    myCollIDS.Add(oRead.GetValue(0))
                End While
            End If
        Catch
        Finally
            If Not oRead.IsClosed Then oRead.Close()
            CConnection.SetConnClosed()

        End Try
        Me.FillList() 'oColls)                         'Fill list view with values found
        Me.Cursor = Cursors.Default                 'Set cursor back to normal
        Medscreen.LogTiming("Diary form Do Query End")
    End Sub

    'Private Sub DoQuery(ByVal strQuery As String)
    '    'Does the actual querying of the database using query constructed
    '    Medscreen.LogTiming("Diary form Do Query start - " & strQuery)
    '    Me.Cursor = Cursors.WaitCursor              'Let the world know this might take some time
    '    MyCollections.Load("Lib_collection.GetDiaryRecords", strQuery)
    '    Me.FillList() 'oColls)                         'Fill list view with values found
    '    Me.Cursor = Cursors.Default                 'Set cursor back to normal
    '    Medscreen.LogTiming("Diary form Do Query End")
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill List viewer
    ''' </summary>
    ''' <param name="oColls">Collection created in DoQuery</param>
    ''' <remarks>
    '''     ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillList() 'ByVal oColls As ArrayList)
        'Routine to fill list view with the contents of the the collection of collections
        Dim J As Integer
        'Dim objColl As Intranet.intranet.jobs.CMJob
        Dim objColl As String
        Dim Item As ListViewItems.ListViewCMDiaryItem
#If DEBUG Then
        Medscreen.Logging = True
#End If
        Medscreen.LogTiming("Diary form Fill List Start")
        'Me.MyCollections.Purge(Now.AddMinutes(-15))        'Get rid of old colelctions

        'System.GC.Collect()                                             'Kick off garbage collection

        Me.lvDiaryList.Items.Clear()                'Clear out list view 

        Dim intI As Integer = 0         'Set a counter

        Dim startTime As Date = Now
        For Each objColl In Me.myCollIDS   'Me.MyCollections                  'The collections are 1 based
            Medscreen.LogTiming("  Diary form get item " & J)
            Medscreen.LogTiming("  Diary form about to build " & J)

            Item = Me.BuildItem(objColl)            'Create a list view item 
            '                               note that this is an inherited ListViewDiaryItem item 
            '                               that handles the building and drawing of itself
            Medscreen.LogTiming("  Diary form about to built " & J)
            If Not Item Is Nothing Then
                Me.lvDiaryList.Items.Add(Item)
                intI += 1                           'Increment counter
                If intI Mod 20 = 0 Then             'If 20th etc item update display
                    Me.lvDiaryList.EndUpdate()
                    Me.lvDiaryList.Invalidate()
                    Me.StatusBarPanel2.Text = "Loading Data ... " & intI & " Loaded"
                    Dim LoadString As String = MedscreenCommonGUIConfig.Diary.Item("FillCounter")
                    LoadString = Medscreen.ReplaceString(LoadString, "!*", "<")
                    LoadString = Medscreen.ReplaceString(LoadString, "*!", ">")
                    Dim TotalSecs As Double = Now.Subtract(startTime).TotalSeconds
                    Dim Average As Double = TotalSecs / intI
                    Dim expected As Double = (myCollIDS.Count - intI) * Average
                    Dim argString As String = CStr(intI) & "," & myCollIDS.Count & "," & CInt(expected) & "," & CInt(CDbl(myCollIDS.Count * Average))
                    Dim strArgs As String() = argString.Split(New Char() {","})
                    Dim args As Integer()
                    'System.Array.Copy(strArgs, args, 4)
                    Try
                        LoadString = String.Format(LoadString, strArgs)
                    Catch
                    End Try
                    Me.HtmlEditor1.DocumentText = (LoadString)
                    Me.ToolStripStatusLabel1.Text = CStr(intI)
                    'Me.StatusBar1.Invalidate()
                    Application.DoEvents()
                    Me.lvDiaryList.BeginUpdate()
                End If
            End If
            'Me.ToolStripStatusLabel1.Text = CStr(intI)

        Next
        Me.ToolStripStatusLabel1.Text = lvDiaryList.ContentsString
        Me.StatusBar1.Invalidate()
        Me.lvDiaryList.EndUpdate()
        Application.DoEvents()
        Medscreen.LogTiming("Diary form Fill List Finish")
    End Sub

    ''The user wants to edit a collection this uses the collection form 
    'Private Sub EditCollection(Optional ByVal FormState As frmCollectionForm.FormModes = frmCollectionForm.FormModes.Edit)
    '    'First check is to see if it is a call out  (callout form) or routine 
    '    If (Me.myCollection.CollectionType = Constants.GCST_JobTypeCallout) Or _
    '    (Me.myCollection.CollectionType = Constants.GCST_JobTypeCalloutOverseas) Then
    '        Dim FormCallOut As New frmCallOut()
    '        With FormCallOut          'Use call out form 
    '            .Collection = lvItemCurrent.Collection
    '            Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
    '            Me.Invalidate()
    '            .Collection.Refresh()
    '            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
    '                lvItemCurrent.Collection = .Collection
    '                If .Collection.CollectionType = Constants.GCST_JobTypeCallout Then
    '                    'If .Collection.Status = Constants.GCST_JobStatusCreated Then
    '                    '    AssignCallOutToCollector(.Collection)
    '                    'End If
    '                End If
    '            End If
    '            Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            Me.Invalidate()
    '        End With
    '    Else            'Routine collection 
    '        With CommonForms.FormCollection
    '            .FormMode = FormState           'Set the form into a particular state 
    '            .Collection = lvItemCurrent.Collection      'Set the colelction property of the form to list view item collection 
    '            Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
    '            Me.Invalidate()
    '            Try
    '                .ShowDialog()       'Show form, no need to get a return as the DB will be updated if the user wants 
    '            Catch ex As Exception

    '            End Try
    '            Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            Me.Invalidate()
    '        End With
    '    End If
    '    'Ensure that we refresh the display 
    '    If Not Me.lvItemCurrent Is Nothing Then
    '        Me.lvItemCurrent.Refresh()
    '        Me.lvDiaryList_Click(Nothing, Nothing)
    '    End If
    'End Sub

    'Send notification message about the cancellation 


    Private Sub frmDiary_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MedscreenCommonGui.CommonForms.UserOptions.CollDestination = myDestination
        MedscreenCommonGui.CommonForms.UserOptions.DiaryStatusFilter = Me.strGStatus
        MedscreenCommonGui.CommonForms.UserOptions.DiaryTypeFilter = Me.strGType
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.GetColumns(Me.lvDiaryList)
    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.StatusBarPanel1.Text = Mid(Me.LoginTimeStamp.AddMinutes(120).Subtract(Now).ToString, 1, 10)
        Me.StatusBar1.Invalidate()
        Application.DoEvents()
        Me.Timer1.Enabled = True
    End Sub

    Protected Overrides Sub Finalize()
        myCollection = Nothing
        myUser = Nothing      'User definition class
        myOfficer = Nothing
        If Not myColls Is Nothing Then
            myColls.Clear()
            myColls = Nothing
        End If
        Me.myCache.Clear()
        MyBase.Finalize()
    End Sub

    Private Sub AssignCallOutToCollector(ByVal objCM As Intranet.intranet.jobs.CMJob)
        'TODO add to list offer to assign to collector
        If MsgBox("Do you want to assign the collecting officer now?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            objCM.Status = Constants.GCST_JobStatusSent
            objCM.Update()
            CommonForms.ConfirmCollection(objCM)
            objCM.Refresh()                   'Ensure any changes have been updated
            objCM.CallOutConfirmation()
        End If

    End Sub

    Private Sub GetCmJobNumber(ByVal strBarcode As String)
        If strBarcode.Length() > 0 Then
            'See if collection is valid

            Forms.Cursor.Current = Cursors.WaitCursor
            'Dim objColl As Intranet.intranet.jobs.CMJob
            Dim id As Integer = -1
            Try
                id = strBarcode
            Catch
            End Try
            If id > -1 Then                     'Numeric ID
                strBarcode = strBarcode.PadLeft(10)
                strBarcode = strBarcode.Replace(" ", "0")
            End If
            Dim strCheck As String = CConnection.PackageStringList("Lib_collection.isactiveCollCust", strBarcode)
            If strCheck.Trim.Length > 0 AndAlso (strCheck.Chars(0) <> "Z" OrElse strBarcode.Chars(0) = "T") Then
                'objColl = MyCollections.Item(strBarcode)
                If strBarcode.Chars(0) = "T" AndAlso strCheck.Chars(0) = "Z" Then
                    Dim Exists As String = CConnection.PackageStringList("lib_collectionxml8.CollectionExists", strBarcode)
                    If Exists IsNot Nothing AndAlso Exists = "F" Then
                        CreateCollFromCardiffTicket(strBarcode)
                    End If
                End If
                Me.lvDiaryList.SelectedItems.Clear()
                'If Not objColl Is Nothing Then
                'see if in current list
                Dim Items As ListViewItem = Me.lvDiaryList.Search(strBarcode)

                If Items Is Nothing Then
                    Dim Item As ListViewItems.ListViewCMDiaryItem = Me.BuildItem(strBarcode)
                    If Not Item Is Nothing Then
                        Me.lvDiaryList.Items.Add(Item)
                        Item.Selected = True
                        Item.EnsureVisible()
                        Me.lvItemCurrent = Item
                        Me.lvDiaryList_Click(Nothing, Nothing)
                    End If
                Else
                    Items.Selected = True
                    Items.EnsureVisible()
                    lvItemCurrent = Items
                    Me.lvDiaryList_Click(Nothing, Nothing)


                End If

            Else ' Issue a waring 
                MsgBox("That collection {" & strBarcode & ") doesn't exist or it has an invalid customer", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            End If
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub FindCm(ByVal TFind As String)
        Dim lvi As ListViewItem
        'Dim lvis As ListViewItem.ListViewSubItem

        Dim intI As Integer = -1

        Me.lvDiaryList.SelectedItems.Clear()
        For Each lvi In Me.lvDiaryList.Items
            If InStr(lvi.Text.ToUpper, TFind) <> 0 Then
                lvi.Selected = True
                lvi.EnsureVisible()
                intI = lvi.Index
                Exit For
            End If
        Next

    End Sub

    Private Sub UnCheckViewMenus()
        Me.mnuNormalView.Checked = False
        Me.mnuDoneView.Checked = False
        Me.mnuReceived.Checked = False
        Me.mnuOutstandingView.Checked = False
        Me.mnuJobsOnADay.Checked = False
        Me.mnuSpecialPay.Checked = False
        Me.mnuViewCustomerCollections.Checked = False
        mnuViewSMIDSCollections.Checked = False
    End Sub

#End Region

#Region "Menu Handlers"

    'Filter on a port or site
    Private Sub MnuPortSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPortSite.Click
        strGPort = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Enter Port ID")
        'Find out from the user what port / site
        If strGPort.Length > 0 Then         'If one chosen provide feedback
            Me.mnuPortSite.Checked = True
            Me.mnuPortSite.Text = "&Port/Site - " & strGPort.ToUpper
        Else                                'or clear menu 
            Me.mnuPortSite.Checked = False
            Me.mnuPortSite.Text = "&Port/Site"
        End If
        Me.GetCollectionDetails()           'Refresh list view 
    End Sub

    'Filter on a customer 
    Private Sub mnuCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustomer.Click
        Dim objF As New frmCustomerSelection2()
        'use the customer selection form 
        objF.CallOutOnly = False        'Not restricted to call out customers
        If objF.ShowDialog = Windows.Forms.DialogResult.OK Then       'Ok pressed
            If objF.Client Is Nothing Then Exit Sub 'Somethings gone wrong bog off
            strCustomer = objF.Client.Identity      'Set customer chosen 
        Else
            Exit Sub
        End If

        'If we have a customer then provide feedback or clear filter 
        If strCustomer.Length > 0 Then
            Me.mnuCustomer.Checked = True
            Me.mnuCustomer.Text = "&Customer - " & strCustomer.ToUpper
        Else
            Me.mnuCustomer.Checked = False
            Me.mnuCustomer.Text = "&Customer"
        End If
        Me.GetCollectionDetails()           'Refresh contents of list 
    End Sub

    'Filter by date 
    Private Sub mnuDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDate.Click
        Dim strInput As String = ""
        strInput = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Enter From Date ")
        'Get a date using a call to get parameter
        If strInput.Length > 0 Then     'A date's been provided give feedback 
            Me.mnuDate.Checked = True
            Me.datFrom = Date.Parse(strInput)
            Me.mnuDate.Text = "&From Date - " & datFrom.ToString("dd-MMM-yyyy")
            Me.datTo = Now.AddMonths(3)
        Else
            Me.datFrom = Now.AddMonths(-3)      'default to three months in the past 
            Me.mnuDate.Checked = False
            Me.mnuDate.Text = "&From Date"
            Me.datTo = Now.AddMonths(3)         'And three months in the future 
        End If
        Me.GetCollectionDetails()               'rebuild list 
    End Sub

    'Filter on statuses 
    Private Sub mnuStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStatus.Click
        Dim frmStat As frmStatusSelect

        frmStat = New frmStatusSelect()         'uses status form to identify status needed
        Me.mnuAvailable.Checked = False
        frmStat.Status = Me.strGStatus          'Initialise the form with current status values 
        If frmStat.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.strGStatus = frmStat.Status
            Me.mnuStatus.Text = "Status in  " & frmStat.Status  'get statuses from form use set operator
            Me.mnuStatus.Checked = True
            Me.mnuAvailable.Checked = (strGStatus = "V")    'aVailable is one of the statuses
        Else                                    'Leave as is 
        End If
        'If there are no statuses chosen default it to Sent, Available and waiting 
        If strGStatus.Trim.Length = 0 Then strGStatus = "('P','V','W')"
        MedscreenCommonGui.CommonForms.UserOptions.DiaryStatusFilter = strGStatus  'Save it in the user options 
        Me.GetCollectionDetails()               'Refill list view 
    End Sub

    'Clear out all filters 
    Private Sub mnuClearAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuClearAll.Click
        Me.strGPort = ""
        Me.strGVessel = ""
        Me.strCustomer = ""
        Me.strGStatus = ""
        Me.strGType = ""
        Me.mnuCustomer.Text = "&Customer"
        Me.mnuCustomer.Checked = False
        Me.mnuPortSite.Text = "&Port/Site"
        Me.mnuPortSite.Checked = False
        Me.mnuDate.Checked = False
        Me.mnuDate.Text = "&Date From "
        Dim tmpDate As Date = Now.AddMonths(-1)
        Me.datFrom = DateSerial(tmpDate.Year, tmpDate.Month, 1)
        Me.datTo = Now.AddMonths(3)
        Me.mnuStatus.Text = "&Status"
        Me.mnuStatus.Checked = False
        Me.mnuAvailable.Checked = False

        Me.GetCollectionDetails()           'And redisplay default view 
    End Sub

    'Filter on a vessel 
    Private Sub mnuVessel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuVessel.Click
        'TODO: change to use vessel selection form so we get a vessel id 
        Me.strGVessel = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Enter Vessel")
        If strGVessel.Length > 0 Then
            Me.mnuVessel.Checked = True
            Me.mnuVessel.Text = "&Vessel - " & strGVessel.ToUpper
        Else
            Me.mnuVessel.Checked = False
            Me.mnuVessel.Text = "&Vessel"
        End If
        Me.GetCollectionDetails()
    End Sub


    'Filter allow cancelled collections to show
    Private Sub mnuShowCancelled_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowCancelled.Click
        blnShowCancelled = Not blnShowCancelled
        Me.mnuShowCancelled.Checked = blnShowCancelled
        Me.GetCollectionDetails()
    End Sub

    'Assign a collecting officer to a collection 
    Private Sub mnuAssignActive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAssignActive.Click

        Dim objColl As Intranet.intranet.jobs.CMJob = Me.lvItemCurrent.Collection

        If objColl Is Nothing Then Exit Sub 'If we don't have a valid collection bog off 

        If objColl.Status = Constants.GCST_JobStatusCollected Then
            MsgBox("Can't (re)assign collection it has already been done", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        If objColl.Status = Constants.GCST_JobStatusReceived Then
            MsgBox("Can't (re)assign collection it has already been done", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        If objColl.Status = Constants.GCST_JobStatusCommitted Then
            MsgBox("Can't (re)assign collection it is finished", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        If objColl.Status = Constants.GCST_JobStatusConfirmed Then
            Dim dlgret As MsgBoxResult = MsgBox("Do you want to unconfirm the collection?, so you can reassign it", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
            If dlgret = MsgBoxResult.No Then Exit Sub
            Try
                objColl.CollRedirectReport()
                objColl.Comments.CreateComment("Collection Unconfirmed", Constants.GCST_COMM_CCFix)

                'MsgBox("This function is not yet available")
                objColl.SetStatus(Constants.GCST_JobStatusCreated)
                objColl.Update()
            Catch exm As MedscreenExceptions.CanNotChangeCollectionStatus
                MsgBox(exm.Message)
            End Try
        End If



        CommonForms.AssignOfficer(objColl)    'Make a call to global routine to assign officer 
        Forms.Cursor.Current = Cursors.Default
        Me.Cursor = Cursors.Default
        lvItemCurrent.Refresh()         'Update display 

        'End If
    End Sub


    'Not yet active, will allow users to enter pay details 
    Private Sub mnuAssignPayDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssignPayDetails.Click
        Dim objColl As Intranet.intranet.jobs.CMJob
        Dim objPort As Intranet.intranet.jobs.Port
        Try
            If Not Me.lvItemCurrent Is Nothing Then
                If InStr("AVPWSC", Me.lvItemCurrent.Collection.Status) <> 0 Then
                    MsgBox("Pay statements can only be entered for jobs that have been done or have samples", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    Exit Sub
                End If
                Dim CollPayForm As New MedscreenCommonGui.frmCollPayEntry()
                With CollPayForm
                    objColl = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)
                    If Not objColl Is Nothing Then
                        objPort = cSupport.Ports.Item(objColl.Port)
                        If Not objPort Is Nothing Then
                            'If objPort.IsFixedSite Then
                            'MsgBox("No collector pay at fixed site ports", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)
                            'Else '
                            If Date.Compare(objColl.CollectionDate, Now) > 0 Then
                                MsgBox("Collection not yet done!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            Else
                                Dim objCollPay As Intranet.collectorpay.CollectorPayCollection = New Intranet.collectorpay.CollectorPayCollection()

                                objCollPay.Load("CM_JOB_NUMBER = '" & objColl.ID & "'")

                                Dim objCollPayItem As Intranet.collectorpay.CollectorPay

                                If objCollPay.Count > 0 Then
                                    objCollPayItem = objCollPay.Item(objColl.ID, Intranet.collectorpay.CollectorPayCollection.IndexPayOn.CMNumber)
                                    If objCollPayItem.PayType.Trim.Length = 0 Then objCollPayItem.PayType = "C"
                                    If objCollPayItem.CollOfficer.Trim.Length = 0 Then objCollPayItem.CollOfficer = objColl.CollOfficer
                                    .CollPay = objCollPayItem
                                Else
                                    objCollPayItem = objCollPay.Create
                                    objCollPayItem.PayType = "C"
                                    objCollPayItem.CMidentity = objColl.ID
                                    objCollPayItem.CollOfficer = objColl.CollOfficer
                                    objCollPayItem.PaymentMonth = DateSerial(Today.Year, Today.Month, 1)
                                    If Not objColl.CollectingOfficer Is Nothing Then
                                        If objColl.CollectionType = "O" Or objColl.CollectionType = "C" Then
                                            objCollPayItem.PayMethod = "H"
                                        Else
                                            objCollPayItem.PayMethod = objColl.CollectingOfficer.PayMethod
                                        End If
                                    End If
                                    objCollPayItem.DoUpdate()
                                    'Get the entries 
                                    objCollPayItem.CreateEntries()
                                    .CollPay = objCollPayItem
                                End If
                                ' = objColl
                                'If .PayDetails Then

                                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                                    .CollPay.DoUpdate()
                                End If
                                'Else
                                '   .ShowDialog()
                                'End If
                            End If
                        End If
                        'End If
                    End If
                End With
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, True)
        End Try
    End Sub

    Dim blnCanCancel As Boolean = False

    'We are cancelling the collection need to do loads of checks 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' handle cancelling a collection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2006]</date><Action>Added diary refresh</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancel.Click
        Dim objColl As Intranet.intranet.jobs.CMJob
        Dim nt As System.Security.Principal.WindowsIdentity
        Dim p As System.Security.Principal.WindowsPrincipal
        'Dim lvs As ListViewItem.ListViewSubItem
        'Dim cs As ListViewItems.CollStatusSubItem

        nt = System.Security.Principal.WindowsIdentity.GetCurrent
        p = New System.Security.Principal.WindowsPrincipal(nt)

        'We need to check that the user has sufficient rights to do this they need to be 
        'a memeber of CM_aadministrator 
        Debug.WriteLine(nt.Name)
        If p.IsInRole("LOGOS\CM_Administrator") Or blnCanCancel Then
            'If they haven't used the application for twenty minutes check their password
            If Now.Subtract(Me.LoginTimeStamp).TotalMinutes > 20 Then  'Validate user 
                If Medscreen.CheckPassword() Then
                    Me.LoginTimeStamp = Now
                Else
                    MsgBox("Invalid User Name or Password")
                    Me.nwInfo.Notify("Invalid User Name or Password")
                    Exit Sub
                End If
            End If

            'Stop refresshing the list, this will confuse matters
            Me.blnStopRefresh = True
            'Get collection object
            objColl = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)
            If Not objColl Is Nothing Then          'If we have a valid collection
                If objColl.Status = Constants.GCST_JobJobStatusCancelled Then   'Has it already been cancelled
                    'MsgBox("This collection has already been cancelled", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                    Try
                        objColl.UnCancel()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    Finally
                        objColl.Refresh()
                        Me.lvItemCurrent.Refresh()
                    End Try
                    Exit Sub
                End If
                CommonForms.CancelCollection(objColl)         'Go to common cancel collection routine
                objColl.Refresh()                           'Get changed details
                Me.lvItemCurrent.Refresh()
            End If

            Me.lvItemCurrent.Refresh()                      'Refresh the line in the display
            Me.blnStopRefresh = False
            SendEmailToManager(objColl)                     'Send an email to let people know of the cancellation 
            'Me.GetCollectionDetails()                       
        Else
            MsgBox("You have insufficent rights to perform this operation, please contact IT")
            nwInfo.Notify("You have insufficent rights to perform this operation, please contact IT")
        End If
    End Sub

    '''<summary>
    '''Confirm the collection i.e. Collecting officer has said they can do the collection
    '''Result is a SM job will be created and status will be C
    ''' </summary>
    Private Sub mnuConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConfirm.Click

        Try
            blnStopRefresh = True           'Stop list refresh
            Dim objColl As Intranet.intranet.jobs.CMJob = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)

            If Not objColl Is Nothing Then      'If we don't have a valid collection bog off 
                If objColl.Status = Constants.GCST_JobStatusConfirmed Then
                    MsgBox("This collection has already been confirmed", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                    nwInfo.Notify("This collection has already been confirmed")
                    Exit Sub
                End If
                CommonForms.ConfirmCollection(objColl) 'Use common confirm routine

                objColl.Refresh()                   'Ensure any changes have been updated

                Me.lvItemCurrent.Refresh()          'Refresh list item
                Me.lvItemCurrent.Selected = True
                objColl.Refresh()
                If objColl.JobName.Trim.Length > 0 Then     ' If the external process has completed 
                    '               Display the job details
                    '               If it hasn't no issue as it will display when it completes but no message
                    MsgBox("Job Name is : " & objColl.JobName)
                    nwInfo.Notify("Job Name is : " & objColl.JobName)
                End If
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "mnu confirm")
        Finally
            blnStopRefresh = False
        End Try
    End Sub

    'Send a collection request to collecting officers
    Private Sub mnuSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSend.Click
        Try
            blnStopRefresh = True
            If Me.lvItemCurrent Is Nothing Then
                MsgBox("No collection selected ")
                nwInfo.Notify("No collection selected ")
                Exit Sub
            End If
            Dim objColl As Intranet.intranet.jobs.CMJob = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)

            If Not objColl Is Nothing Then      'If we don't have a valid collection bog off 
                If InStr("CDRM", objColl.Status) > 0 Or objColl.JobName.Trim.Length > 0 Then
                    MsgBox("This collection has already been Confirmed", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                    nwInfo.Notify("This collection has already been Confirmed " & objColl.JobName)
                    Exit Sub
                End If
                Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", objColl.ID)
                SendToCollector(objColl, False, True, "Sending", sendMethod)
                Me.lvItemCurrent.Refresh()
            End If

        Catch ex As MedscreenExceptions.WordOutputException
            MsgBox(ex.ToString)
        Catch ex As Exception
            Medscreen.LogError(ex, , "mnu send diary")
        Finally
            blnStopRefresh = False
        End Try
    End Sub

    Private Sub mnuAmendColl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAmendColl.Click
        ''
        If Not Me.lvItemCurrent Is Nothing Then
            EditCollection(lvItemCurrent.Collection, frmCollectionForm.FormModes.Edit)
        End If
    End Sub

    Private Sub mnuDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDetails.Click
        Me.lvDiaryList.View = View.Details
        Me.lvDiaryList.Invalidate()
    End Sub

    Private Sub mnuList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuList.Click
        Me.lvDiaryList.View = View.List
        Me.lvDiaryList.Invalidate()
    End Sub

    Private Sub mnuLargeIcons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuLargeIcons.Click
        Me.lvDiaryList.View = View.LargeIcon
        Me.lvDiaryList.Invalidate()
    End Sub

    Private Sub MnuSmallIcons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuSmallIcons.Click
        Me.lvDiaryList.View = View.SmallIcon
        Me.lvDiaryList.Invalidate()
    End Sub

    'Through up workflow form 
    Private Sub mnuView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuView.Click
        If (myCollection.HasID > 0) Then
            Dim tf As New frmWorkflow()

            tf.DisplayType = 0
            tf.Collection = myCollection
            tf.User = myUser
            tf.Action = Intranet.intranet.jobs.CMJob.Actions.Display
            tf.ShowDialog()

        End If

    End Sub

    Private Sub mnuAvailable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAvailable.Click
        If mnuAvailable.Checked Then
            Me.strGStatus = ""
            mnuAvailable.Checked = False
        Else
            mnuAvailable.Checked = True
            Me.strGStatus = "('" & MedscreenLib.Constants.GCST_JobJobStatusAvailable & "')"
            Me.GetCollectionDetails()
        End If
    End Sub

    Private Sub mnuMine_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMine.Click
        Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item( _
             System.Security.Principal.WindowsIdentity.GetCurrent.Name, MedscreenLib.personnel.PersonnelList.FindIdBy.WindowsIdentity)

        If Me.mnuMine.Checked Then
            Me.mnuMine.Checked = False
            Me.strGUser = ""
        Else
            If Not up Is Nothing Then
                If up.Identity = "TAYLOR" Then
                    Me.strGUser = "MACKIN"
                Else
                    Me.strGUser = up.Identity
                End If

            Else
                Me.strGUser = ""

            End If
            Me.mnuMine.Checked = True
        End If
        SetTitleBar()
        GetCollectionDetails()
    End Sub

    Private Sub mnuRadioChange(ByVal Mnu As ToolStripMenuItem, ByVal state As Boolean)
        Mnu.Checked = state
        If state Then
            Const INT_on As Integer = 3
            Mnu.Image = imgListStat.Images(INT_on)
        Else
            Const INT_off As Integer = 2
            Mnu.Image = imgListStat.Images(INT_off)
        End If
    End Sub
    Private Sub mnuShowDestall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowDestall.Click
        Me.myDestination = Destination.All
        mnuRadioChange(Me.mnuShowDestOverseas, False)
        mnuRadioChange(mnuShowDestall, True)
        mnuRadioChange(mnuShowDestUK, False)
        SetTitleBar()
        GetCollectionDetails()

    End Sub



    Private Sub mnuShowDestUK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowDestUK.Click
        myDestination = Destination.UK
        mnuRadioChange(Me.mnuShowDestOverseas, False)
        mnuRadioChange(mnuShowDestall, False)

        mnuRadioChange(Me.mnuShowDestUK, True)
        SetTitleBar()
        GetCollectionDetails()
    End Sub

    Private Sub mnuShowDestOverseas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowDestOverseas.Click
        myDestination = Destination.Overseas
        mnuRadioChange(Me.mnuShowDestOverseas, True)
        mnuRadioChange(Me.mnuShowDestall, False)
        mnuRadioChange(Me.mnuShowDestUK, False)
        SetTitleBar()
        Me.GetCollectionDetails()
    End Sub

    Private Sub mnuNewCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewCollection.Click
        'Dim intRet As DialogResult
        'Dim blnClient As Boolean
        'Dim oVessel As Intranet.intranet.customerns.Vessel

        Me.LoginTimeStamp = Now
        If Now.Subtract(Me.LoginTimeStamp).TotalMinutes > 120 Then  'Validate user 
            If Medscreen.CheckPassword() Then
                Me.LoginTimeStamp = Now
            Else
                MsgBox("Invalid User Name or Password")
                nwInfo.Notify("Invalid User Name or Password")
                Exit Sub
            End If
        Else

        End If
        Forms.Cursor.Current = Cursors.WaitCursor
        Dim blnRet As Boolean
        Dim objCm As Intranet.intranet.jobs.CMJob = CommonForms.BookRoutine(blnRet)
        Try
            If Not objCm Is Nothing Then        'Check we have got a valid collection
                Dim Item As ListViewItems.ListViewCMDiaryItem
                If blnRet Then                  'Check again 
                    Item = Me.BuildItem(objCm.ID)
                    Me.lvDiaryList.Items.Add(Item)          'Create a new row 
                    Me.lvDiaryList.SelectedItems.Clear()    'Make it the selected item
                    Item.Selected = True
                    Item.EnsureVisible()
                    Me.lvDiaryList_Click(Nothing, Nothing)  'Try and get everything displayed
                    lvItemCurrent.Collection.LogAction(jobs.CMJob.Actions.CreateCollection)


                End If
            End If
            'Me.mnuRefresh_Click(Nothing, Nothing)
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try
    End Sub

    Private Sub MnuCreditCard_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreditCard.Click
        Me.PrintDialog1.PrinterSettings = New System.Drawing.Printing.PrinterSettings()
        Me.PrintDialog1.ShowDialog()
        myCollection.CreditCardTicket(PrintDialog1.PrinterSettings)
    End Sub

    Private Sub mnuCustomise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustomise.Click
        'Dim frmCust As New FrmCustomise()

        'frmCust.ShowDialog()
        'Me.GetCollectionDetails()
    End Sub

    Private Sub mnuAmend2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAmend2.Click
        Me.mnuAmendColl_Click(sender, e)
    End Sub

    Private Sub mnuAssignComplete2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssignComplete2.Click
        Me.mnuAssignComplete_Click(sender, e)
    End Sub

    Private Sub mnuAssigntoActive2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssigntoActive2.Click
        Me.mnuAssignActive_Click(sender, e)
    End Sub

    Private Sub mnuAssignComplete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssignComplete.Click
        Me.mnuAssignActive_Click(sender, e)
    End Sub

    Private Sub mnuAssignPayDetails2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssignPayDetails2.Click
        Me.mnuAssignPayDetails_Click(sender, e)
    End Sub

    Private Sub mnuSend2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSend2.Click
        Me.mnuSend_Click(sender, e)
    End Sub

    Private Sub mnuConfirm2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConfirm2.Click
        Me.mnuConfirm_Click(sender, e)
    End Sub

    Private Sub mnuCancel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancel2.Click
        Me.mnuCancel_Click(sender, e)
    End Sub

    Private Sub mnuResend2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuResendColl.Click
        Me.mnuResend_Click(sender, e)
    End Sub

    Private Sub mnuResend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuResendColl2.Click
        '
        'Me.mnuSend_Click(Nothing, Nothing)  'Temporay redirection to test code 

        Dim strSendType As String
        Dim objColl As Intranet.intranet.jobs.CMJob = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)

        If objColl Is Nothing Then Exit Sub ' No valid collection exit 

        If (InStr("WV", objColl.Status) <> 0) Or (objColl.CollOfficer.Trim.Length = 0) Then 'Waiting or not assigned
            MessageBox.Show("Can not resend : collection has not been assigned, you will be taken to assignment", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.mnuAssignActive_Click(Nothing, Nothing)
            Exit Sub
        End If


        'Deal with callouts they have different documents
        If objColl.CollClassType = jobs.CMJob.CollectionTypes.GCST_JobTypeCallout Or _
            objColl.CollClassType = jobs.CMJob.CollectionTypes.GCST_JobTypeCalloutOverseas Then    'Deal with callouts 
            objColl.CallOutConfirmation()
            Exit Sub
        End If

        Dim blnPrintOnly As Boolean = False
        Select Case MsgBox("Would you like to Re-send collection " & objColl.ID & "?" & vbCrLf & _
                           "to " & objColl.CollOfficer & vbCrLf & _
                           "(No = Print collection only)", vbYesNoCancel + vbQuestion)
            Case vbYes
                strSendType = "re-sent"

            Case vbNo
                strSendType = "printed"
                blnPrintOnly = True
            Case vbCancel
                Exit Sub
        End Select
        Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", objColl.ID)
        CommonForms.SendToCollector(objColl, blnPrintOnly, , "Resend", sendMethod)
        'WriteResendCollectionDetails(strCollRef, strSendType)

    End Sub

    Private Sub mnuJobsOnADay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuJobsOnADay.Click
        '
        Try
            Dim strInput As Object
            Dim strDate As String = ""
            strInput = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Date to filter on")
            If Not strInput Is Nothing Then
                strDate = strInput
                Me.blnStopRefresh = True
                UnCheckViewMenus()
                Me.mnuJobsOnADay.Checked = True
                Dim tmpDate As Date = Date.Parse(strDate)

                Dim strSquery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("JobsOnADay")
                strSquery = Medscreen.ExpandQuery(strSquery, tmpDate.ToString("dd-MMM-yy"), tmpDate.AddDays(1).ToString("dd-MMM-yy"))
                Me.DoQuery(strSquery)
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuCallOut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCallOut.Click
        'Dim intRet As DialogResult
        'Dim blnClient As Boolean
        'Dim oVessel As Intranet.intranet.customerns.Vessel

        If Now.Subtract(Me.LoginTimeStamp).TotalMinutes > 20 Then  'Validate user 
            If Medscreen.CheckPassword() Then
                Me.LoginTimeStamp = Now
            Else
                MsgBox("Invalid User Name or Password")
                nwInfo.Notify("Invalid User Name or Password")
                Exit Sub
            End If
        Else

        End If

        Dim objCm As Intranet.intranet.jobs.CMJob = Nothing

        If CommonForms.BookCallOut(objCm) Then
            'AssignCallOutToCollector(objCm)


        End If
        If objCm Is Nothing Then Exit Sub
        Dim Item As ListViewItems.ListViewCMDiaryItem
        Item = Me.BuildItem(objCm.ID)
        Me.lvDiaryList.Items.Add(Item)
        Me.lvDiaryList.SelectedItems.Clear()
        Item.Selected = True
        Item.EnsureVisible()
        Me.lvDiaryList_Click(Nothing, Nothing)


    End Sub

    Private Sub mnuCalloutInstructions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCalloutInstructions.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        Me.lvItemCurrent.Collection.CallOutConfirmation()
    End Sub

    Private Sub mnuResendConf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuResendConf.Click
        With Me.lvItemCurrent.Collection
            If .ConfirmationType.Trim = "COMPLETED ONLY" OrElse .ConfirmationType.Trim = "BOTH ARRANGED AND COMPLETED" Then
                If Me.lvItemCurrent.Collection.CollectionDate > Now Then
                    Me.lvItemCurrent.Collection.CustomerConfirmation(ConfirmationType.Arrange)
                Else
                    Me.lvItemCurrent.Collection.CustomerConfirmation(ConfirmationType.Confirm)
                End If
            End If
        End With
    End Sub

    Private Sub mnuResendConf2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuResendConf2.Click
        Me.mnuResendConf_Click(sender, e)
    End Sub

    Private Sub mnuResendCOConfirmation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuResendCOConfirmation.Click
        'Send out callout request to cofficer.
        If Me.lvItemCurrent.Collection.CollectionType = Constants.GCST_JobTypeCallout Then
            Me.lvItemCurrent.Collection.CallOutConfirmation()
        Else
            Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", lvItemCurrent.Collection.ID)
            CommonForms.SendToCollector(lvItemCurrent.Collection, False, , "Resend Confirmation", sendMethod)
        End If
    End Sub

    Private Sub mnuTypeFixedSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTypeFixedSite.Click
        Me.strGType = Constants.GCST_JobTypeFixed
        mnuRadioChange(Me.mnuTypeAll, False)
        mnuRadioChange(Me.mnuTypeCallout, False)
        mnuRadioChange(mnuTypeFixedSite, True)
        Me.mnuRadioChange(MnuTypeRoutine, False)
        mnuRadioChange(mnuTypeWorldWide, False)
        Me.GetCollectionDetails()
    End Sub

    Private Sub mnuTypeAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTypeAll.Click
        Me.strGType = ""
        mnuRadioChange(mnuTypeAll, True)
        mnuRadioChange(mnuTypeCallout, False)
        mnuRadioChange(mnuTypeFixedSite, False)
        Me.mnuRadioChange(MnuTypeRoutine, False)
        mnuRadioChange(mnuTypeWorldWide, False)
        Me.GetCollectionDetails()
    End Sub

    Private Sub MnuTypeRoutine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTypeRoutine.Click
        Me.strGType = Constants.GCST_JobTypeWorkplace
        mnuRadioChange(mnuTypeAll, False)
        mnuRadioChange(mnuTypeCallout, False)
        mnuRadioChange(mnuTypeFixedSite, False)
        Me.mnuRadioChange(MnuTypeRoutine, True)
        Me.GetCollectionDetails()

    End Sub

    Private Sub mnuTypeCallout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTypeCallout.Click
        Me.strGType = Constants.GCST_JobTypeCallout
        mnuRadioChange(mnuTypeAll, False)
        mnuRadioChange(mnuTypeCallout, True)
        mnuRadioChange(mnuTypeFixedSite, False)
        Me.mnuRadioChange(MnuTypeRoutine, False)
        mnuRadioChange(mnuTypeWorldWide, False)
        Me.GetCollectionDetails()

    End Sub

    Private Sub mnuTypeWorldWide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTypeWorldWide.Click
        Me.strGType = Constants.GCST_JobTypeOverseas
        mnuRadioChange(mnuTypeAll, False)
        mnuRadioChange(mnuTypeCallout, False)
        mnuRadioChange(mnuTypeFixedSite, False)
        Me.mnuRadioChange(MnuTypeRoutine, False)
        mnuRadioChange(mnuTypeWorldWide, True)
        Me.GetCollectionDetails()

    End Sub

    Private Sub mnuHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelp.Click
        Help.ShowHelp(Me, CommonForms.ApplicationPath & "\CCMaintain.chm", HelpNavigator.KeywordIndex, "Collection Diary")

    End Sub

    Private Sub mnuCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollection.Click

    End Sub

    Private Sub mnuHuntCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHuntCollection.Click
        Dim strBarcode As String
        strBarcode = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Enter Barcode", "Barcode Entry")
        Me.GetCmJobNumber(strBarcode)
    End Sub

    Private Sub mnuProblem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuProblem.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusPay Then
            MsgBox("This action is not relevant to Pay collections")
            nwInfo.Notify("This action is not relevant to Pay collections")
        End If
        If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusCommitted Then
            MsgBox("This action is not relevant to completed collections")
            nwInfo.Notify("This action is not relevant to completed collections")
        End If

        Dim afrm As New FrmCollComment()
        afrm.CommentType = "PROB"
        afrm.CMNumber = Me.lvItemCurrent.Collection.ID
        If Me.lvItemCurrent.Collection.ProgrammeManager.Trim.Length > 0 Then
            afrm.ProgamManager = Me.lvItemCurrent.Collection.ProgrammeManager
        End If
        If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                Me.lvItemCurrent.Collection.Comments.CreateComment(afrm.Comment, afrm.CommentType)
            If Not aComm Is Nothing Then
                aComm.ActionBy = afrm.ActionBy
                aComm.Update()
                lvItemCurrent.Collection.CancellationStatus = Constants.GCST_CANC_Problem
                lvItemCurrent.Collection.Update()
                lvItemCurrent.Collection.Refresh(False)
                Dim UP As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item(afrm.ActionBy, _
                    MedscreenLib.personnel.PersonnelList.FindIdBy.identity)
                Dim strSubject As String = "Issue with Collection CM" & lvItemCurrent.Collection.ID
                If Not lvItemCurrent.Collection.VesselObject Is Nothing Then
                    strSubject += " vessel " & lvItemCurrent.Collection.VesselName
                ElseIf Not lvItemCurrent.Collection.SiteObject Is Nothing Then
                    strSubject += " site " & lvItemCurrent.Collection.VesselName
                Else
                    strSubject += lvItemCurrent.Collection.vessel
                End If

                strSubject += " Port " & lvItemCurrent.Collection.Port
                strSubject += " Officer " & lvItemCurrent.Collection.CollOfficer
                If Not UP Is Nothing Then
                    Medscreen.BlatEmail(strSubject, _
                    afrm.Comment, afrm.txtDestination.Text, cSupport.User.Email)
                End If
                If afrm.ckCopy.Checked Then
                    Medscreen.BlatEmail(strSubject, _
                    afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)
                End If
                lvItemCurrent.Collection.LogAction(jobs.CMJob.Actions.RaiseIssue, afrm.txtCopyDest.Text)

            End If
        End If
        lvItemCurrent.Refresh()

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddComment.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        Dim afrm As New FrmCollComment()
        afrm.CommentType = "MISC"
        afrm.FormType = FrmCollComment.FCCFormTypes.def
        afrm.ProgamManager = Me.lvItemCurrent.Collection.ProgrammeManager
        'afrm.CollOfficer = Me.lvItemCurrent.Collection.CollOfficer

        If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                Me.lvItemCurrent.Collection.Comments.CreateComment(afrm.Comment, afrm.CommentType)
            aComm.ActionBy = afrm.ActionBy
            aComm.Update()
            lvItemCurrent.Collection.Update()
            lvItemCurrent.Collection.Refresh(False)
            If afrm.ckCopy.Checked Then
                Dim strSubject As String = "Comment has been made on collection : " & lvItemCurrent.Collection.ID
                Medscreen.BlatEmail(strSubject, _
                afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)

            End If
        End If

    End Sub

    Private Sub mnuRaiseProblem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRaiseProblem2.Click
        Me.mnuProblem_Click(Nothing, Nothing)
    End Sub

    Private Sub mnuAddComment2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddComment2.Click
        Me.MenuItem3_Click(Nothing, Nothing)
    End Sub

    Private Sub mnuPrintHTML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPrintHTML.Click
        Me.HtmlEditor1.Print()
    End Sub

    Private Sub mnuEmailHTML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEmailHTML.Click
        Dim elements As mshtml.IHTMLElementCollection = Me.HtmlEditor1.Document.All

        Dim strBody = "<HTML>" & elements.item(0).innerhtml & "</HTML>"
        Dim stremail As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail

        If blnEditorOther Then
            Medscreen.BlatEmail("Report for User ", strBody, stremail, , True)
        Else
            Medscreen.BlatEmail("Report for collection " & Me.lvItemCurrent.Collection.ID, strBody, stremail, , True)
        End If

    End Sub

    Private Sub mnuRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefresh.Click
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.GetColumns(Me.lvDiaryList)
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.CreateColumns(Me.lvDiaryList, Me.imgColHeadList)
        Me.GetCollectionDetails()
    End Sub

    Private Sub mnuCOComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCOComment.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        Try
            If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusPay Then
                MsgBox("This action is not relevant to Pay collections")
                nwInfo.Notify("This action is not relevant to Pay collections")
            End If
            If Me.lvItemCurrent.Collection.CollOfficer.Trim.Length = 0 Then
                nwInfo.Notify("This action is not relevant to collections to which an officer has not been assigned")
                MsgBox("This action is not relevant to collections to which an officer has not been assigned")
            End If


            Dim afrm As New FrmCollComment()
            afrm.CommentType = "PROB"
            afrm.FormType = FrmCollComment.FCCFormTypes.CollectorInfo
            afrm.ProgamManager = lvItemCurrent.Collection.SmClient.ProgrammeManager
            afrm.Port = lvItemCurrent.Collection.CollectionPort
            afrm.CollOfficer = lvItemCurrent.Collection.CollOfficer
            afrm.CMNumber = Me.lvItemCurrent.Collection.ID

            If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                    Me.lvItemCurrent.Collection.Comments.CreateComment(afrm.Comment, afrm.CommentType)
                aComm.ActionBy = afrm.CollOfficer
                aComm.Update()
                lvItemCurrent.Collection.CancellationStatus = Constants.GCST_CANC_Problem
                lvItemCurrent.Collection.Update()
                lvItemCurrent.Collection.Refresh(False)
                'TODO: add code to produce an email / fax the document

                Dim myColl As Intranet.intranet.jobs.CMJob = lvItemCurrent.Collection
                Dim myOff As Intranet.intranet.jobs.CollectingOfficer = myColl.CollectingOfficer

                If afrm.ckCopy.Checked Then
                    If Medscreen.IsNumber(afrm.txtCopyDest.Text) Then
                        Medscreen.QuckEmail("Issue with Collection CM" & lvItemCurrent.Collection.ID, _
                        afrm.Comment, afrm.txtCopyDest.Text & Constants.GCST_EMAIL_To_Fax_Send, cSupport.User.Email)
                    Else
                        Medscreen.QuckEmail("Issue with Collection CM" & lvItemCurrent.Collection.ID, _
                        afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)
                    End If

                End If
                myColl.SendOfficerProblem(afrm.txtCopyDest.Text, afrm.Comment, afrm.SendTo)
                lvItemCurrent.Refresh()

            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "MnuCOComment")
        End Try
    End Sub

    Private Sub mnuFiltIssue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFiltIssue.Click
        Me.blnIssue = Not blnIssue
        Me.mnuFiltIssue.Checked = blnIssue

        Me.GetCollectionDetails()
    End Sub

    Private Sub mnuResolveProblem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuResolveProblem.Click
        Me.lvItemCurrent.Collection.CancellationStatus = "N"
        Me.lvItemCurrent.Collection.Update()
        Me.lvItemCurrent.Refresh()
    End Sub

    Private Sub mnuRecover_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRecover.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        With Me.lvItemCurrent.Collection
            .Status = Constants.GCST_JobStatusCreated
            .Update()
            Dim myClient As customerns.Client = cSupport.CustList.Item(.ClientId)
            If myClient Is Nothing Then Exit Sub
            If myClient.CollectionInfo Is Nothing Then Exit Sub
            .WhoToTest = myClient.CollectionInfo.WhoToTest
            .NoOfDonors = myClient.CollectionInfo.NoDonors
            .ConfirmationContact = myClient.CollectionInfo.ConfirmationContactList
            .ConfirmationNumber = myClient.CollectionInfo.ConfirmationSendAddress
            .ConfirmationMethod = myClient.CollectionInfo.ConfirmationMethod
            .CollType = myClient.CollectionInfo.ReasonForTest
            .ConfirmationType = myClient.CollectionInfo.ConfirmationType
            .SpecialProcedures = myClient.CollectionInfo.Procedures
            .CollectionComments = myClient.CollectionInfo.AccessToSite
            .Tests = MedscreenLib.Glossary.Glossary.TestTypeList.Item(myClient.CollectionInfo.TestType).PhraseText
            .UK = myClient.CollectionInfo.UKCollection
            .Status = Constants.GCST_JobStatusCreated
            .Update()
        End With
    End Sub

    Private Sub mnuColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColumns.Click
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.GetColumns(Me.lvDiaryList)
        Dim afrm As New frmListViewCustomise()
        afrm.ShowDialog()
        Me.mnuRefresh_Click(Nothing, Nothing)
        ' myUserOptions.DiaryColumns.CreateColumns(Me.lvDiaryList)
    End Sub

    Private Sub mnuDeferCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeferCollection.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        With Me.lvItemCurrent.Collection
            Dim newDate As Date = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                "New Collection Date", "Defer Collection " & .CMID & " - " & .ClientId & _
                " " & .VesselName, .CollectionDate)
            If newDate.Day * newDate.Month * newDate.Year <> 1 Then
                If MsgBox("Change collection date to " & newDate.ToString("dd-MMM-yyyy"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    .CollectionDate = newDate
                    If .Status = Constants.GCST_JobStatusConfirmed Then
                        If Not lvItemCurrent.Collection.Job Is Nothing Then
                            .Job.CollectionDate = newDate
                            .Job.Update()
                        End If
                    End If
                    .Comments.CreateComment("Collection Deffered to " & newDate.ToString("dd-MMM-yyyy"), Constants.GCST_COMM_CCFix)
                    .Update()
                    .Refresh()
                    If .Status = Constants.GCST_JobStatusSent Or .Status = Constants.GCST_JobStatusConfirmed Then
                        If MessageBox.Show("Send a note to the collecting officer to tell them the collection has been delayed?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                            Me.mnuCOComment_Click(Nothing, Nothing)
                        End If

                    End If
                End If
            End If
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Record that the collection documents have been obtained    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action>Date shipped is now updated</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCollDocReceived_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollDocReceived.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        If Not (Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusConfirmed Or _
            Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusReceived) Then
            MessageBox.Show("Only confirmed or received collections can be set as Collected", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If
        With Me.lvItemCurrent
            If MessageBox.Show("Do you wish to record that the documents indicating that the collection has been done, have been received for " _
                & .Collection.CMID & "?", "Collection Done", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
                Dim objRet As Object
                Try
                    objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                        "Collection Date", "Collection Date", .Collection.CollectionDate)
                    If objRet Is Nothing Then Exit Sub
                Catch ex As Exception
                    objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                        "Collection Date", "Collection Date", Now.AddDays(3))
                    If objRet Is Nothing Then Exit Sub
                End Try
                Dim dateRec As Date = objRet
                objRet = Medscreen.GetParameter(Medscreen.MyTypes.typeInteger, _
                    "No of Donors", "No of Donors", .Collection.NoOfDonors)
                If objRet Is Nothing Then Exit Sub
                Dim NDonors As Integer = objRet
                If Not .Collection.VesselObject Is Nothing Then ' we have a vessel
                    If Mid(.Collection.VesselObject.VesselID, 1, 4) = "VESS" Or _
                        Mid(.Collection.VesselObject.VesselID, 1, 3) = "IMO" Then 'We have no ID
                        objRet = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Please enter any vessel id given on crew list", "", "")
                        .Collection.CreditCardNumber = objRet
                    End If
                End If
                If .Collection.Status = Constants.GCST_JobStatusConfirmed Then .Collection.Status = Constants.GCST_JobStatusCollected
                .Collection.Comments.CreateComment("Documents received from collecting officer", "CONF")
                .Collection.RecordCollectionAction("DOCREC", dateRec, NDonors)
                .Collection.Medscreenref = NDonors
                .Collection.DateShipped = dateRec
                .Collection.CollectionDate = dateRec
                objRet = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Airway Bill Number", , "")

                If objRet Is Nothing Then Exit Sub
                Dim strAirway As String = objRet
                strAirway = Medscreen.ReplaceString(strAirway, " ", "")
                .Collection.AirwayBillNo = strAirway
                .Collection.Update()
                .Refresh()
            End If
        End With

    End Sub

    Private Sub MnuViewOfficerDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuViewOfficerDetails.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        Try
            If Not Me.lvItemCurrent.Collection.CollectingOfficer Is Nothing Then
                Medscreen.TextEditor(Me.lvItemCurrent.Collection.CollectingOfficer.ContactDetails, True, , "Contact details")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub mnuCollectorMaintain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCollectorMaintain.Click
        'TODO edit select collector code 
        Dim objCS As New frmSelectCollector()

        If objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim objColl As Intranet.intranet.jobs.CollectingOfficer = objCS.Officer

            Dim frmedColl As New frmCollOfficerEd()

            frmedColl.CollectingOfficer = objColl
            If frmedColl.ShowDialog = Windows.Forms.DialogResult.OK Then
                objColl = frmedColl.CollectingOfficer
                objColl.Update()

            End If
        End If
    End Sub

    Private Sub mnuHuntVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHuntVessel.Click
        Dim oVessel As Intranet.intranet.customerns.Vessel = HuntVessel()
        If Not oVessel Is Nothing Then
            Dim objCm As Intranet.intranet.jobs.CMJob
            Dim lvItem As ListViewItems.ListViewCMDiaryItem = Nothing
            For Each objCm In oVessel.Collections
                For Each lvItem In Me.lvDiaryList.Items
                    If lvItem.Collection.ID = objCm.ID Then
                        lvItem.Checked = True
                        lvItem.Selected = True
                        lvItem.EnsureVisible()
                        Exit For
                    End If
                    lvItem = Nothing
                Next
                If lvItem Is Nothing Then 'Add a new item
                    lvItem = New ListViewItems.ListViewCMDiaryItem(objCm.ID)
                    lvItem.Checked = True
                    lvItem.Selected = True
                    Me.lvDiaryList.Items.Add(lvItem)
                    lvItem.EnsureVisible()

                End If
            Next
            Me.lvDiaryList_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub mnuEditVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditVessel.Click

        Dim oVessel As Intranet.intranet.customerns.Vessel
        oVessel = HuntVessel()

        If oVessel Is Nothing Then Exit Sub
        Dim edFrm As New frmVessel()
        edFrm.Vessel = oVessel

        If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            oVessel = edFrm.Vessel
            oVessel.Update()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a new port to the collection of ports
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuNewPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewPort.Click
        'Dim PortId As String = Medscreen.GetParameter(Medscreen.MyTypes.typString, "New Port ID", "New Port ID", "")
        'If PortId.Trim.Length > 20 Then
        '    MsgBox("Id is too long 20 characters max")
        '    Exit Sub

        'End If

        'PortId = PortId.ToUpper

        Dim aPort As Intranet.intranet.jobs.Port = Nothing
        Dim aFrm As New frmSelectPort()
        With aFrm
            .CountryID = 44
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                aPort = .Port
            End If
        End With
        If aPort Is Nothing Then Exit Sub
        'With TreeView1.CurrentPort
        '    .Children.Add(aPort)
        '    aPort.PortRegion = .PortId
        '    aPort.CountryId = .CountryId
        '    aPort.IsUK = .IsUK
        '    aPort.PortType = "P"
        'End With
        'aPort.PortType = "P"
        'aPort.CountryId = 44
        aPort.Update()
        'Me.TreeView1.SelectedNode.Nodes.Add(New SitePortNode(aPort))
        Try
            Dim frm As New frmPort()
            With frm
                .Port = aPort
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    aPort = .Port
                    aPort.Update()

                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuNormalView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNormalView.Click
        Me.blnStopRefresh = True
        Me.mnuNormalView.Checked = True
        Me.GetCollectionDetails()
        Me.blnStopRefresh = False
    End Sub

    Private Sub mnuDoneView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDoneView.Click
        UnCheckViewMenus()
        Me.blnStopRefresh = True
        Me.mnuDoneView.Checked = True
        Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsDone")
        Me.DoQuery(strQuery)
        'Me.DoQuery("Select c.*,rowid from collection c where status = 'D'")
    End Sub

    Private Sub mnuReceived_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReceived.Click
        Me.blnStopRefresh = True
        UnCheckViewMenus()
        Me.mnuReceived.Checked = True
        Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsReceived")
        Me.DoQuery(strQuery)

    End Sub

    Private Sub mnuOutstanding_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOutstandingView.Click
        Me.blnStopRefresh = True
        UnCheckViewMenus()
        Me.mnuOutstandingView.Checked = True
        Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsOutstanding")
        Me.DoQuery(strQuery)
        '        Me.DoQuery("Select c.*,rowid from collection c where status in ('C','D') and coll_date < sysdate - 5")

    End Sub

    Private Sub mnuUnconfirmCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUnconfirmCollection.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusReceived Or _
        Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusCollected Or _
        Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusCommitted Then
            MsgBox("This collection can not be unconfirmed now", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If Not Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusConfirmed Then
            MsgBox("This collection is not confirmed", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Try
            If MsgBox("This option will change the status of a collection from confirmed to available, " & vbCrLf & _
            "you will then be offered the chance of changing the details of the collection." & _
            vbCrLf & "Do you want to proceed (CM" & Me.lvItemCurrent.Collection.ID & ") ?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
            Try
                Me.lvItemCurrent.Collection.CollRedirectReport()
                Me.lvItemCurrent.Collection.Comments.CreateComment("Collection Unconfirmed", Constants.GCST_COMM_CCFix)

                'MsgBox("This function is not yet available")
                Me.lvItemCurrent.Collection.SetStatus(Constants.GCST_JobStatusCreated)
                Me.lvItemCurrent.Collection.Update()
            Catch exm As MedscreenExceptions.CanNotChangeCollectionStatus
                MsgBox(exm.Message)
            End Try

            'See if we are coming here from rearranging the collection
            If Not TypeOf e Is Intranet.intranet.jobs.JobEvent Then

                EditCollection(lvItemCurrent.Collection, frmCollectionForm.FormModes.ChangePort)

                If MsgBox("Do you want to reassign this collection", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    Me.lvItemCurrent.Collection.Update()
                    Me.mnuAssignActive_Click(Nothing, Nothing)
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a new port 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ChangeDestination()
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        Dim dlgRet As Microsoft.VisualBasic.MsgBoxResult
        'See if it is a banned status
        If InStr("DRM", Me.lvItemCurrent.Collection.Status) <> 0 Then
            MsgBox("this collection is too far advanced, it is recorded as being done!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        'See about Confirmed collections 
        If Me.lvItemCurrent.Collection.Status = MedscreenLib.Constants.GCST_JobStatusConfirmed Then
            dlgRet = MsgBox("This collection has been confirmed, do you want to unconfirm it then alter the destination?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
            If dlgRet = MsgBoxResult.No Then Exit Sub

            Dim e As New Intranet.intranet.jobs.JobEvent(Me.lvItemCurrent.Collection.ID)

            Me.mnuUnconfirmCollection_Click(Me.lvItemCurrent.Collection, e)

        End If

        Dim afrm As New MedscreenCommonGui.frmSelectPort()
        Dim oPort As Intranet.intranet.jobs.Port
        If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
            oPort = afrm.Port
            If Not oPort Is Nothing Then
                Me.lvItemCurrent.Collection.Port = oPort.Identity
                Me.lvItemCurrent.Collection.UK = oPort.IsUK
                If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusConfirmed Then
                    Me.lvItemCurrent.Collection.Job.Location = oPort.Identity
                    Me.lvItemCurrent.Collection.Job.Update()
                End If
                Me.lvItemCurrent.Status = Constants.GCST_JobStatusCreated

                Me.lvItemCurrent.Collection.Update()
            End If

        End If
    End Sub

    Private Sub mnuAlterDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAlterDestination.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        Try
            'Run change destination code
            Me.ChangeDestination()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub mnuPutOnHold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPutOnHold.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        If Me.lvItemCurrent.Collection.Status = Constants.GCST_JobStatusCommitted Then Exit Sub

        If Me.lvItemCurrent.Collection.SetStatus(Constants.GCST_JobStatusOnHold) Then
            Me.lvItemCurrent.Collection.Update()
        End If

    End Sub

    Private Sub mnuRefreshTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefreshTime.Click
        Dim intRefresh As Integer = MedscreenCommonGui.CommonForms.UserOptions.RefreshTime / 600000
        intRefresh = Medscreen.GetParameter(Medscreen.MyTypes.typeInteger, "Refresh time", , intRefresh)
        MedscreenCommonGui.CommonForms.UserOptions.RefreshTime = intRefresh * 600000
    End Sub

    Private Sub mnuJobCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuJobCost.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        Try
            Dim Cost As Double
            Dim strHtml As String = Me.lvItemCurrent.Collection.PriceJobHTML(Cost)
            MedscreenLib.Medscreen.ShowHtml(strHtml)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuViewCustomerPricing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewCustomerPricing.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub

        Try
            'Dim Cost As Double
            Dim strHtml As String = Me.lvItemCurrent.Collection.SmClient.ToHTML("customer.xsl", MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\transforms\xsl\")
            MedscreenLib.Medscreen.ShowHtml(strHtml)
        Catch ex As Exception
        End Try

    End Sub

#End Region

#Region "List Event Handlers"
    Private Sub lvDiaryList_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvDiaryList.ColumnClick
        ' Set the ListViewItemSorter property to a new ListViewItemComparer 
        ' object. Setting this property immediately sorts the 
        ' ListView using the ListViewItemComparer object.

        Dim intdirection As Integer

        intdirection = Me.ColumnDirection(e.Column)     'Find out the currently chossen direction
        MedscreenCommonGui.CommonForms.UserOptions.DiaryColumn = e.Column            'Record the column 
        MedscreenCommonGui.CommonForms.UserOptions.DiarySortDirection = intdirection 'and direction 
        Me.ColumnDirection(e.Column) *= -1              'Invert direction and store

        Me.lvDiaryList.ListViewItemSorter = New Comparers.ListViewItemComparer(e.Column, intdirection)
        'Create a new sorter with these options 
    End Sub


    'This event will fire when the user clicks a row in the list view 
    Private Sub lvDiaryList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDiaryList.Click
        Dim tmplv As ListViewItems.ListViewCMDiaryItem
        blnEditorOther = False
        Dim blnFutureColl As Boolean = False    'Is this a collection in the future 
        Me.LoginTimeStamp = Now
        tmplv = Me.lvItemCurrent
        Me.lvItemCurrent = Nothing              'Set the currently selected item to nothing 
        If Me.lvDiaryList.SelectedItems.Count > 0 Then          'We have an item selected 
            Me.lvItemCurrent = Me.lvDiaryList.SelectedItems.Item(0) 'Set current collection to first selected 
            If tmplv IsNot Nothing AndAlso tmplv.Collection.ID = Me.lvItemCurrent.Collection.ID AndAlso sender IsNot Nothing Then
                Exit Sub
            End If
            If Not Me.lvItemCurrent.Collection Is Nothing Then      'If we have a valid selection refresh it and redraw
                Me.lvItemCurrent.Refresh()  ' Refresh and redraw item
            End If
            menus.CollectionMenu.ContextMenuRedraw(lvItemCurrent.Collection, ctxMenu)
            Me.StatusBarPanel2.Text = lvItemCurrent.ToolTip         'Set header display to contents of row 
            'Me.nwInfo.Notify(lvItemCurrent.ToolTip)
            blnFutureColl = (Date.Compare(Me.lvItemCurrent.CollDate, Now) > 0)  'Is it in the future?

            Me.ContextMenu = Me.ctxMenu        'Set up context menu 
        Else
            Me.lvItemCurrent = Nothing
            Me.ContextMenu = Nothing
        End If

        If Not Me.lvItemCurrent Is Nothing Then
            myCollection = Me.lvItemCurrent.Collection
            'myCollection = ModTpanel.iCollections(False).Item(Me.lvItemCurrent.CMJobNumber)
            myCollection.Refresh(False)          'Make sure details are up to date 
            lvItemCurrent.Refresh(False)
            'Is this a collection done by credit card
            If myCollection.CreditCardNumber.Trim.Length > 0 Then
                Me.mnuCreditCard.Visible = True     'Set the credit card menu option 
            End If
            If myCollection.CollectionType = Constants.GCST_JobTypeCallout Then     'Is this a UK call out 
                Me.mnuCalloutInstructions.Visible = True
                Me.mnuSend.Visible = False
                Me.mnuSend2.Visible = False
            End If
            Try
                'Display collection as HTML
                Dim strHTML As String = myCollection.ToHTML(, , True, False)
                Me.HtmlEditor1.DocumentText = (strHTML)
                If Not myCache Is Nothing Then Me.txtNavigate.Text = Me.myCache.Add("http://andrew/WebGCMSLate/WebFormpositives.aspx?Action=collection&id=" & myCollection.ID & "&StyleSheet=test3.xsl", strHTML).HyperlinkXML
                txtNavigate.Refresh()
                Application.DoEvents()
            Catch ex As Exception
            End Try
            If myCollection.Invoice Is Nothing Then
                Me.MnuEmailInvoice.Visible = False
            Else
                Me.MnuEmailInvoice.Visible = True
            End If
            ToolStripStatusLabel2.Text = myCollection.ToolTip
        End If
        Me.lvDiaryList.Enabled = True
        Me.Panel2.Enabled = True
    End Sub



    'In case we change it programatically 
    Private Sub lvDiaryList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvDiaryList.SelectedIndexChanged
        If Me.lvDiaryList.SelectedItems.Count = 0 Then Exit Sub

        'Okay use value 
        'lvItemCurrent = Me.lvDiaryList.SelectedItems.Item(0)
        'Me.ToolTip1.SetToolTip(Me.lvDiaryList, lvItemCurrent.ToolTip)
    End Sub


    'Double click event indicates we want to edit the collection 
    Private Sub lvDiaryList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDiaryList.DoubleClick
        If Not Me.lvItemCurrent Is Nothing Then
            EditCollection(lvItemCurrent.Collection, frmCollectionForm.FormModes.Edit)         'Call edit method 
        End If
    End Sub

    Private Sub lvDiaryList_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDiaryList.MouseHover

    End Sub

    Private Sub lvDiaryList_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvDiaryList.MouseDown
        Me.ListButton = e.Button
    End Sub

    Private Sub lvDiaryList_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvDiaryList.ItemCheck
        If e.CurrentValue = CheckState.Checked Then
            'Me.lvDiaryList.Items.Item(e.Index).Checked = False
            e.NewValue = CheckState.Unchecked
        Else
            'Me.lvDiaryList.Items.Item(e.Index).Checked = True
            e.NewValue = CheckState.Checked
        End If
    End Sub

#End Region

#Region "Other Control Handlers"


    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
        Dim frm As New FindInList()
        Try
            frm.View = Me.lvDiaryList
            frm.FirstColumnCMNumber = True
            Dim dlgRes As DialogResult = frm.ShowDialog
            If dlgRes = Windows.Forms.DialogResult.OK Then
                Me.lvDiaryList = frm.View
                Me.GetCmJobNumber(frm.FindString.ToUpper)
            ElseIf dlgRes = Windows.Forms.DialogResult.Retry Then
                Me.GetCmJobNumber(frm.FindString.ToUpper)
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "CmdFind")
        End Try
    End Sub


    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Control Then
            If e.KeyCode = Keys.F Then
                Me.cmdFind_Click(Nothing, Nothing)
            ElseIf e.KeyCode = Keys.T Then
                Me.cmdConfirmPaper_Click(Nothing, Nothing)
            ElseIf e.KeyCode = Keys.I Then
                Me.mnuAssignPayDetails_Click(Nothing, Nothing)
            End If

        End If

    End Sub


    Private Sub FinishLoad()
        If Me.lvDiaryList.Items.Count > 0 Then
            Dim strFirst As String = MedscreenCommonGUIConfig.Diary("First")
            strFirst = Medscreen.ReplaceString(strFirst, "!*", "<")
            strFirst = Medscreen.ReplaceString(strFirst, "*!", ">")

            Me.HtmlEditor1.DocumentText = (strFirst)
            Application.DoEvents()
            Me.lvDiaryList.Items(0).Selected = True
            Me.lvDiaryList_Click(Nothing, Nothing)
        Else
            Dim strFirst As String = MedscreenCommonGUIConfig.Diary("NoCollections")
            strFirst = Medscreen.ReplaceString(strFirst, "!*", "<")
            strFirst = Medscreen.ReplaceString(strFirst, "*!", ">")

            Me.HtmlEditor1.DocumentText = (strFirst)
            Application.DoEvents()
        End If

    End Sub


    Private Sub timRefresh_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timRefresh.Tick
        Try
            timRefresh.Interval = 3000000
            If Not blnStopRefresh Then
                If Not Me.lvItemCurrent Is Nothing Then
                    Dim strCm As String = Me.lvItemCurrent.Collection.ID

                    Me.mnuRefresh_Click(Nothing, Nothing)
                    Me.FindCm(strCm)
                    Me.lvItemCurrent = Me.lvDiaryList.SelectedItems.Item(0)
                    Me.lvItemCurrent.Refresh()
                Else
                    Me.mnuRefresh_Click(Nothing, Nothing)
                    'FinishLoad()
                End If
            End If
        Catch ex As Exception
        Finally
            Me.timRefresh.Enabled = True
        End Try
    End Sub

    Private Sub cmdConfirmPaper_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConfirmPaper.Click
        Dim frm As New FindInList()

        frm.View = Me.lvDiaryList
        frm.FirstColumnCMNumber = True
        Dim dlgRes As DialogResult = frm.ShowDialog
        If dlgRes = Windows.Forms.DialogResult.OK Then
            Me.lvDiaryList = frm.View
            Me.mnuCollDocReceived_Click(Nothing, Nothing)
        ElseIf dlgRes = Windows.Forms.DialogResult.Retry Then
            Me.GetCmJobNumber(frm.FindString)
            Me.mnuCollDocReceived_Click(Nothing, Nothing)
        End If

    End Sub

    Private Sub HtmlEditor1_Navigated(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles HtmlEditor1.Navigated
        CurrentDocument = HtmlEditor1.Document
        If myCache IsNot Nothing Then
            txtNavigate.Text = myCache.GetElement.HyperLink
            txtNavigate.Refresh()
            Application.DoEvents()
        End If
        'Me.lblLink.Text = e.Url.ToString
    End Sub


    Private Sub HtmlEditor1_BeforeNavigate(ByVal s As Object, ByVal e As WebBrowserNavigatingEventArgs) Handles HtmlEditor1.Navigating

        Debug.WriteLine("before " & e.Url.OriginalString)
        If InStr(e.Url.OriginalString.ToUpper, "/WEBGCMSLATE/WEBFORMPOSITIVES.ASPX?ACTION") <> 0 Then
            e.Cancel = True
            Intranet.intranet.Support.cSupport.DoNavigate(e.Url.OriginalString, Me.HtmlEditor1, , myCache)

        ElseIf InStr(e.Url.OriginalString, "about:blank//Local") <> 0 Then
            e.Cancel = True
            DoLocalNavigate(e.Url.OriginalString)
        ElseIf InStr(e.Url.OriginalString.ToLower, "//local") <> 0 Then
            e.Cancel = True
            DoLocalNavigate(e.Url.OriginalString)
        ElseIf InStr(e.Url.OriginalString, "about:blank") <> 0 Then

        Else
            txtNavigate.Text = e.Url.OriginalString
            txtNavigate.Refresh()
            Application.DoEvents()
            If Not myCache Is Nothing Then
                myCache.Add(e.Url.OriginalString, "")
            End If

        End If

    End Sub

    Private Sub DoLocalNavigate(ByVal target As String)
        Dim strAction As String = QueryString(target, "Action")

        If strAction.ToUpper = "VESSELREPORT" Then
        ElseIf strAction.ToUpper = "VIEWEMAIL" Then
            Dim strEmailID As String = QueryString(target & ".", "id")
            Shell("C:\Program Files\Medscreen\OutlookSupport\osupport emailbyid=" & strEmailID, AppWinStyle.MinimizedFocus)
            MsgBox("Your document should be open in a custome viewer", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            'Me.HtmlEditor1.DocumentText =("<html><body>Your document should be open in Outlook</body></html>")
        ElseIf strAction.ToUpper = "OFFICEREDIT" Then
            Dim strOfficer As String = QueryString(target & ".", "officer")

            Dim edFrm As New frmCollOfficerEd()
            Dim objOfficer As Intranet.intranet.jobs.CollectingOfficer = _
                Intranet.intranet.Support.cSupport.CollectingOfficers.Item(strOfficer)
            If Not objOfficer Is Nothing Then
                edFrm.CollectingOfficer = objOfficer
                If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    edFrm.CollectingOfficer.Update()
                End If
            End If
            'Deal with the vessel editing action 
        ElseIf strAction.ToUpper = "VIEWEMAIL" Then
            Dim strEmailID As String = QueryString(target & ".", "id")
            Shell("C:\Program Files\Medscreen\OutlookSupport\osupport emailbyid=" & strEmailID, AppWinStyle.MinimizedFocus)
            MsgBox("Your document should be open in Outlook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            'Me.HtmlEditor1.DocumentText =("<html><body>Your document should be open in Outlook</body></html>")
        ElseIf strAction.ToUpper = "OFFICEREDIT" Then
            Dim strOfficer As String = QueryString(target & ".", "officer")

            Dim edFrm As New frmCollOfficerEd()
            Dim objOfficer As Intranet.intranet.jobs.CollectingOfficer = _
                Intranet.intranet.Support.cSupport.CollectingOfficers.Item(strOfficer)
            If Not objOfficer Is Nothing Then
                edFrm.CollectingOfficer = objOfficer
                If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    edFrm.CollectingOfficer.Update()
                End If
            End If
            'Deal with the vessel editing action 
            'Deal with the vessel editing action 
        ElseIf strAction.ToUpper = "VESSELEDIT" Then
            Dim strVesselID As String = QueryString(target & ".", "vessel")
            Dim edFrm As New frmVessel()
            Dim oVessel As New Intranet.intranet.customerns.Vessel() '= ModTpanel.Vessels.Item(strVesselID, customerns.VesselCollection.VesselIndex.ID, True)
            oVessel.VesselID = strVesselID
            oVessel.Load()
            If Not oVessel Is Nothing Then
                edFrm.Vessel = oVessel

                If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    oVessel = edFrm.Vessel
                    oVessel.Update()
                End If
            End If
            'deal with Moveto cuctomer action
        ElseIf strAction.ToUpper = "CUSTOMERPOSITION" Then
        ElseIf strAction.ToUpper = "VESSELBOOK" Then
            Dim strVesselID As String = QueryString(target & ".", "vessel")
            Dim strCustID As String = QueryString(target & ".", "customer")
            Dim blnRet As Boolean
            Dim objCm As Intranet.intranet.jobs.CMJob
            objCm = CommonForms.BookRoutine(blnRet, strCustID, strVesselID)
        ElseIf strAction.ToUpper = "INVOICE" Then
            Dim strInvoice As String = QueryString(target & ".", "id")
            strInvoice = Mid(strInvoice, 1, strInvoice.Length - 1)
            Dim ACCInt As AccountsInterface = New AccountsInterface()
            With ACCInt
                .Status = MedscreenLib.Constants.GCST_IFace_StatusRequest
                .Reference = strInvoice
                .UserId = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                .SendAddress = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                .SetSendDateNull()
                .CopyType = 2
                .InvoiceType = 4

                .RequestCode = "SEND"

                .Update()
                Dim strError As String = ""
                Dim ChRet As Char
                ChRet = .Wait(strError, 100)
                While ChRet <> Constants.GCST_IFace_StatusCreated And strError.Trim.Length = 0
                    Threading.Thread.Sleep(100)
                    ChRet = .Wait(strError, 1)
                End While

                'Me.lblEmailed.Show()
                If strError.Trim.Length = 0 Then
                    MsgBox("Invoice Emailed/Printed", MsgBoxStyle.OkOnly)
                Else
                    MsgBox(strError)
                End If
                ACCInt.Status = Constants.GCST_IFace_StatusDelete
                ACCInt.Delete()
            End With
        End If
    End Sub


    Private Function QueryString(ByVal target As String, ByVal Search As String) As String
        Dim intPos As Integer = InStr(target.ToUpper, Search.ToUpper & "=")
        Dim strRet As String = ""

        target += ";"
        If intPos = 0 Then
            Return strRet
        Else
            intPos = intPos + Search.Length
            While intPos <= target.Length - 1 And (System.Char.IsLetterOrDigit(target.Chars(intPos)) _
                Or (target.Chars(intPos) = ".") Or (target.Chars(intPos) = "-") _
                Or (target.Chars(intPos) = "@") Or (target.Chars(intPos) = Chr(34)) _
                Or (target.Chars(intPos) = ",") Or (target.Chars(intPos) = "_"))
                strRet += target.Chars(intPos)
                intPos += 1
            End While
            Return strRet
        End If

    End Function

#End Region

#Region "Properties"

#End Region

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a report of the collections done for a customer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCustCollReports_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustCollReports.Click

        Dim objF As MedscreenCommonGui.frmCustomerSelection2 = CommonForms.FormCustSelection
        Dim myClient As Intranet.intranet.customerns.Client = Nothing
        objF.CallOutOnly = False
        If objF.ShowDialog = Windows.Forms.DialogResult.OK Then       'Get the customer
            If objF.Client Is Nothing Then Exit Sub
            myClient = objF.Client                      'save customer in local variable
        End If
        Me.blnStopRefresh = True
        Me.Cursor = Cursors.WaitCursor
        objF = Nothing


        'get start date
        Dim objRet As Object                'Temporaryt variable
        Try
            'Get a date parameter, from start of year
            objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                "Start Date", "Start Date", DateSerial(Now.Year, 1, 1))
            If objRet Is Nothing Then Exit Sub 'Cancelled so ignore
        Catch ex As Exception
            Exit Sub
        End Try
        Dim dateStart As Date = objRet          'convert to date

        'build query
        Dim strQuery As String = "Customer_id = '" & myClient.Identity & "' and coll_date > '" & _
            dateStart.ToString("dd-MMM-yy") & "'"

        'create collection of Collections
        Dim iCollection As New Intranet.intranet.jobs.CMCollection(strQuery)

        'produce XML
        Dim strXML As String = iCollection.ToXML(Intranet.intranet.Support.cSupport.XMLCollectionOptions.Simple Or Intranet.intranet.Support.cSupport.XMLCollectionOptions.WithVessel)
        'produce HTML
        Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "Collection2.xsl", 0)
        Me.blnEditorOther = True
        Me.HtmlEditor1.DocumentText = (strHTML)
        Me.blnStopRefresh = False
        Me.Cursor = Cursors.Default
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show a report of collections by a collecting officer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCollectorCollList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollectorCollList.Click, mnuPortCollectionList.Click

        Dim objCS As New frmSelectCollector()
        Dim objFSP As New MedscreenCommonGui.frmSelectPort()
        Dim objColl As Intranet.intranet.jobs.CollectingOfficer = Nothing
        Dim objPort As Intranet.intranet.jobs.Port = Nothing

        'Select collecting officer 
        If sender Is Me.mnuCollectorCollList AndAlso objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
            objColl = objCS.Officer
            If objColl Is Nothing Then Exit Sub 'No valid Officer 
        ElseIf sender Is Me.mnuPortCollectionList AndAlso objFSP.ShowDialog = Windows.Forms.DialogResult.OK Then
            objPort = objFSP.Port
            If objPort Is Nothing Then Exit Sub 'No valid port 
        End If



        'get start date
        Dim objRet As Object                'Temporaryt variable
        Try
            'Get a date parameter, from start of year
            objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                "Start Date", "Start Date", DateSerial(Now.Year, 1, 1))
            If objRet Is Nothing Then Exit Sub 'Cancelled so ignore
        Catch ex As Exception
            Exit Sub
        End Try
        Dim dateStart As Date = objRet          'convert to date

        Cursor = Cursors.WaitCursor
        'build query
        Dim strQuery As String = ""
        If sender Is Me.mnuCollectorCollList Then
            strQuery = "COLL_OFFICER = '" & objColl.OfficerID & "' and coll_date >= '" & _
                dateStart.ToString("dd-MMM-yy") & "'  and status <> 'S'"
        ElseIf sender Is Me.mnuPortCollectionList Then
            strQuery = "PORT = '" & objPort.Identity & "' and coll_date >= '" & _
                dateStart.ToString("dd-MMM-yy") & "' and status <> 'S'"
        End If
        'create collection of Collections
        Dim iCollection As New Intranet.intranet.jobs.CMCollection(strQuery)

        'produce XML
        Dim strXML As String = iCollection.ToXML(Intranet.intranet.Support.cSupport.XMLCollectionOptions.Simple Or Intranet.intranet.Support.cSupport.XMLCollectionOptions.WithVessel)
        'produce HTML
        Dim strStylesheet As String = MedscreenCommonGui.My.Settings.OfficerStyleSheet
        Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strStylesheet, 0)
        Me.blnEditorOther = True
        Me.HtmlEditor1.DocumentText = (strHTML)
        Me.blnStopRefresh = False
        Me.Cursor = Cursors.Default


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempt to send an email to the officer selected 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuEmailOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEmmailOfficer.Click
        Dim objCS As New frmSelectCollector()

        If objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim objColl As Intranet.intranet.jobs.CollectingOfficer = objCS.Officer
            If objColl.Email.Trim.Length = 0 Then
                MsgBox("Officer doesn't have an email address", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            Dim objOutlook As Outlook.ApplicationClass

            Try
                objOutlook = New Outlook.ApplicationClass()
                Dim objMessage As Outlook.MailItemClass = objOutlook.CreateItem(Outlook.OlItemType.olMailItem)
                objMessage.Recipients.Add(objColl.Email)
                objMessage.Display(True)
                MedscreenLib.SendLogCollection.LogEmailToOfficerNC(objColl.OfficerID, objColl.Email)
                'Debug.WriteLine(objMessage.SentOn)
                'objMessage.Send()
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
            Finally
                objOutlook = Nothing
            End Try
        End If
    End Sub

    Private Sub mnuMailOfficer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMailOfficer.Click
        If Me.myCollection Is Nothing Then Exit Sub
        If Me.myCollection.CollectingOfficer.CanEmailorFax Then
            If myCollection.CollectingOfficer.SendDestination.Trim.Length = 0 Then
                MsgBox("Officer doesn't have an email address", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            Dim objOutlook As Outlook.ApplicationClass

            Try
                objOutlook = New Outlook.ApplicationClass()
                Dim objMessage As Outlook.MailItemClass = objOutlook.CreateItemFromTemplate(MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\dbreports\templates\collofficer.Oft")
                objMessage.Subject = "Re : Collection " & myCollection.ID & " at " & myCollection.Port & " for " & myCollection.VesselName
                objMessage.Recipients.Add(myCollection.CollectingOfficer.SendDestination)
                objMessage.Body += " "
                objMessage.Display(True)
                'Debug.WriteLine(objMessage.SentOn)
                'objMessage.Send()
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
            Finally
                objOutlook = Nothing
            End Try

        End If
    End Sub

    Private Sub mnuSpecialPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSpecialPay.Click
        Me.blnStopRefresh = True
        UnCheckViewMenus()
        Me.mnuSpecialPay.Checked = True
        Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsPay")
        strQuery = Medscreen.ExpandQuery(strQuery, MedscreenCommonGui.CommonForms.UserOptions.PastDays)
        Me.DoQuery(strQuery)
    End Sub

    Private Sub mnuViewCustomerCollections_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewCustomerCollections.Click
        Dim objClient As Intranet.intranet.customerns.Client = Nothing
        Dim frm As New frmCustomerSelection2()
        Me.blnStopRefresh = True
        UnCheckViewMenus()
        Me.mnuViewCustomerCollections.Checked = True
        With frm
            .AnalysisOnly = False
            .CallOutOnly = False
            If .ShowDialog Then
                objClient = frm.Client
            End If
        End With
        If Not objClient Is Nothing Then
            Me.lvDiaryList.Items.Clear()
            Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsByCustomer")
            strQuery = Medscreen.ExpandQuery(strQuery, objClient.Identity, MedscreenCommonGui.CommonForms.UserOptions.PastDays)
            Me.DoQuery(strQuery)
        End If
    End Sub

    Private Sub mnuViewSMIDSCollections_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewSMIDSCollections.Click
        Dim objClient As Intranet.intranet.customerns.Client = Nothing
        Dim frm As New frmCustomerSelection2()
        Me.blnStopRefresh = True
        UnCheckViewMenus()
        Me.mnuViewSMIDSCollections.Checked = True
        With frm
            .AnalysisOnly = False
            .CallOutOnly = False
            If .ShowDialog Then
                objClient = frm.Client
            End If
        End With
        If Not objClient Is Nothing Then
            Me.lvDiaryList.Items.Clear()
            Dim objSmid As Intranet.intranet.customerns.SMID = Intranet.intranet.Support.cSupport.SMIDS.Item(objClient.SMID)
            If Not objSmid Is Nothing Then
                Dim strQuery As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CollectionsBySMID")
                strQuery = Medscreen.ExpandQuery(strQuery, objSmid.IdentityList, MedscreenCommonGui.CommonForms.UserOptions.PastDays)
                Me.DoQuery(strQuery)
            End If
        End If
    End Sub

    Private Sub mnuPastDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPastDays.Click
        Dim intPastDays As Integer = MedscreenCommonGui.CommonForms.UserOptions.PastDays
        intPastDays = Medscreen.GetParameter(Medscreen.MyTypes.typeInteger, "No of days to go back in queries", , intPastDays)
        MedscreenCommonGui.CommonForms.UserOptions.PastDays = intPastDays

    End Sub

#Region "Set Collection Date"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the collection date for the collection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	10/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuSetCollDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetCollDate.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        Dim objDate As Object = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Collection Date", "Collection Date", Me.lvItemCurrent.Collection.CollectionDate)
        If Not objDate Is Nothing Then
            Me.lvItemCurrent.Collection.CollectionDate = CDate(objDate)
            Me.lvItemCurrent.Collection.Update()
        End If
    End Sub
#End Region
#Region "Invoicing"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Record that the payment has been made
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/05/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuRecordCollPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRecordCollPayment.Click
        MedscreenCommonGui.CommonForms.DoCardPay(Me.lvItemCurrent.Collection)
    End Sub


#End Region

    Private Sub MnuEmailInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuEmailInvoice.Click
        If Me.myCollection Is Nothing Then Exit Sub
        If Me.myCollection.Invoice Is Nothing Then Exit Sub
        Dim ACCInt As AccountsInterface = New AccountsInterface()
        ACCInt.PrintInvoice(Me.myCollection.Invoice.InvoiceID, Me.myCollection.Invoice.TransactionType)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit the collecting Officer details
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	16/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuEdOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdOfficer.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection.CollectingOfficer Is Nothing Then Exit Sub
        Dim frmedColl As New frmCollOfficerEd()

        frmedColl.CollectingOfficer = Me.lvItemCurrent.Collection.CollectingOfficer
        If frmedColl.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.lvItemCurrent.Collection.CollectingOfficer = frmedColl.CollectingOfficer
            Me.lvItemCurrent.Collection.CollectingOfficer.Update()

        End If


    End Sub

    Private Sub mnuLockPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLockPay.Click

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create an external collection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' Strategy is to 
    ''' 1) get the customer
    ''' 2) get the cm number
    ''' 3) create collection and use normal editing facility
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	11/10/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCreateExternal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreateExternal.Click
        Forms.Cursor.Current = Cursors.WaitCursor
        Dim blnRet As Boolean
        Dim objCm As Intranet.intranet.jobs.CMJob = CommonForms.BookExternal(blnRet)
        Try
            If Not objCm Is Nothing Then        'Check we have got a valid collection
                Dim Item As ListViewItems.ListViewCMDiaryItem
                If blnRet Then                  'Check again 
                    Item = Me.BuildItem(objCm.ID)
                    Me.lvDiaryList.Items.Add(Item)          'Create a new row 
                    Me.lvDiaryList.SelectedItems.Clear()    'Make it the selected item
                    Item.Selected = True
                    Item.EnsureVisible()
                    Me.lvDiaryList_Click(Nothing, Nothing)  'Try and get everything displayed
                    lvItemCurrent.Collection.LogAction(jobs.CMJob.Actions.CreateCollection)


                End If
            End If
            Me.mnuRefresh_Click(Nothing, Nothing)
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try

    End Sub

    Private Sub mnuHuntbybarcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHuntbybarcode.Click
        Dim objBarcode As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Barcode in collection", "Hunt by barcode")
        If objBarcode Is Nothing Then Exit Sub

        Dim strBarcode As String = CStr(objBarcode)
        Dim strId As String = CConnection.PackageStringList("Lib_collection.GetCollectionFromBarcode", strBarcode)
        If strId Is Nothing OrElse strId.Trim.Length = 0 Then
            MsgBox("No collection found for sample", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        GetCmJobNumber(strId)
    End Sub

    Private Sub mnuHuntbyDonor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHuntbyDonor.Click
        Dim frm As New MedscreenCommonGui.frmFindDonor()
        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim strBarcode As String = frm.CaseNumber
            Dim strId As String = CConnection.PackageStringList("Lib_collection.GetCollectionFromBarcode", strBarcode)
            If strId Is Nothing OrElse strId.Trim.Length = 0 Then
                MsgBox("No collection found for sample", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            GetCmJobNumber(strId)

        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Record details of the payment received
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuRecordCollectionPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRecordCollectionPayment.Click
        MedscreenCommonGui.CommonForms.DoCardPay(Me.lvItemCurrent.Collection)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Output a Collection sending it to yourself 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	22/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCollToSelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollToSelf.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        Dim strDestination As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
        Dim Destinations As String() = strDestination.Split(New Char() {"|"})
        Dim objTextDoc As Reporter.clsTextFile = Me.lvItemCurrent.Collection.CollectorOutFile(Destinations, "SEND")
        Dim objReporter As Reporter.Report = New Reporter.Report()
        Dim strPath As String = "\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates|\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates\Scanned Documents"


        Dim PathArray As String() = strPath.Split(New Char() {"|"})
        objReporter.ReportSettings.SetTemplatePaths(PathArray)
        Dim objCReport As Reporter.clsWordReport = objReporter.GetDatafile(objTextDoc.ExportStringArray)
        Dim myRecipient As Reporter.clsRecipient = New Reporter.clsRecipient()
        myRecipient.address = strDestination
        myRecipient.method = "EMAIL"
        objCReport.SendByEmail(myRecipient)
    End Sub

    Private Sub mnuViewCollRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewCollRequest.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub

        Dim strDestination As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
        Dim Destinations As String() = strDestination.Split(New Char() {"|"})
        Dim objTextDoc As Reporter.clsTextFile = Me.lvItemCurrent.Collection.CollectorOutFile(Destinations, "SEND")
        Dim objReporter As Reporter.Report = New Reporter.Report()
        Dim strPath As String = "\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates|\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates\Scanned Documents"


        Dim PathArray As String() = strPath.Split(New Char() {"|"})
        objReporter.ReportSettings.SetTemplatePaths(PathArray)
        Dim objCReport As Reporter.clsWordReport = objReporter.GetDatafile(objTextDoc.ExportStringArray)
        objCReport.SetWordVisible(True)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ChangeCustomer for a collection 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' R2004 version and beyond will need to ensure that the login template exists in for the new customer
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	22/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuChangeCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChangeCustomer.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        Dim intOk As Integer = InStr("VWP", lvItemCurrent.Collection.Status)
        If intOk = 0 Then
            MsgBox("The customer can only be changed for Collections that haven't yet been confirmed.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim objClient As Intranet.intranet.customerns.Client

        Dim frmCust As MedscreenCommonGui.frmCustomerSelection2 = New MedscreenCommonGui.frmCustomerSelection2()
        Dim intRet As DialogResult
        intRet = frmCust.ShowDialog
        If intRet = DialogResult.OK Then
            If Not frmCust.Client Is Nothing Then _
                objClient = frmCust.Client
        End If
        'See if we have a valid customer 
        If objClient Is Nothing Then
            MsgBox("No new customer selected exiting", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim strOldSMidProfile As String = Me.lvItemCurrent.Collection.SmClient.SMIDProfile
        Dim strNewSmidProfile As String = objClient.SMIDProfile

        'need to see if the client has the appropriate service added 
        Dim Service As String
        If lvItemCurrent.Collection.CollectionType = "U" Or lvItemCurrent.Collection.CollectionType = "W" Then
            Service = "ROUTINE"
        ElseIf lvItemCurrent.Collection.CollectionType = "C" Or lvItemCurrent.Collection.CollectionType = "O" Then
            Service = "CALLOUT"
        ElseIf lvItemCurrent.Collection.CollectionType = "F" Then
            Service = "FIXED"
        End If

        If Not SetupService(objClient, Service) Then Exit Sub


        Dim msgRes As MsgBoxResult = MsgBox("Do you want to replace " & strOldSMidProfile & " with " & strNewSmidProfile & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)

        If msgRes = MsgBoxResult.No Then Exit Sub

        'Need to deal with vessels and customer sites 
        If lvItemCurrent.Collection.VesselName <> "-" Then 'Must be a vessel or site
            MsgBox("This collection is for a vessel or a customer site - that will remain with the previous customer, this may lead to issues.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            lvItemCurrent.Collection.Comments.CreateComment("Vessel Customer site warning issued.", "FIX", jobs.CollectionCommentList.CommentType.Collection)

        End If
        lvItemCurrent.Collection.Comments.CreateComment("Customer changed from " & strOldSMidProfile & " to " & strNewSmidProfile, "FIX", jobs.CollectionCommentList.CommentType.Collection)

        lvItemCurrent.Collection.ClientId = objClient.Identity
        lvItemCurrent.Collection.SmClient = Nothing
        lvItemCurrent.Collection.Update()
        lvItemCurrent.Refresh()


    End Sub

    Private Sub cmdBrowserBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowserBack.Click
        If Not myCache Is Nothing Then
            If myCache.Position > 0 Then
                myCache.Position -= 1
                Me.HtmlEditor1.DocumentText = myCache.GetElement.Read
                txtNavigate.Text = myCache.GetElement.HyperLink
                txtNavigate.Refresh()
                Application.DoEvents()
            End If

        End If
        'Me.HtmlEditor1.GoBack()
    End Sub

    Private Sub CmdBrowserForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdBrowserForward.Click
        If Not myCache Is Nothing Then
            If myCache.Position < myCache.Count - 1 Then
                myCache.Position += 1
                Me.HtmlEditor1.DocumentText = myCache.GetElement.Read
                txtNavigate.Text = myCache.GetElement.HyperLink
                txtNavigate.Refresh()
                Application.DoEvents()
            End If

        End If
        'Me.HtmlEditor1.GoForward()
    End Sub

    Private Sub mnuCompleteCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompleteCollection.Click
        Try
            If Me.lvItemCurrent Is Nothing Then Exit Sub
            If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
            Me.lvItemCurrent.Collection.CompleteJob()
            'Now get a comment 
            Dim objString As Object = MedscreenLib.Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", "Comment on completion")
            If Not objString Is Nothing Then
                lvItemCurrent.Collection.Comments.CreateComment(objString.ToString, "INFO", jobs.CollectionCommentList.CommentType.Collection)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' set agreed costs without having to enter collection 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub mnuCollAgreedCosts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCollAgreedCosts.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        MedscreenCommonGui.menus.CollectionMenu.SetAgreedCosts(Me.lvItemCurrent.Collection)
        Me.lvItemCurrent.Refresh()
    End Sub

    Private Sub mnuScannedDocs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScannedDocs.Click
        If Me.lvItemCurrent Is Nothing Then Exit Sub
        If Me.lvItemCurrent.Collection Is Nothing Then Exit Sub
        Me.lvItemCurrent.Collection.GetScannedDocuments()
    End Sub

    Private Sub frmDiary_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.blnDataLoaded = (Me.lvDiaryList.Items.Count > 0)
        If Not Me.blnDataLoaded Then
            Me.Timer1.Interval = 6000
            Me.Timer1.Enabled = True
            Dim tmpDate As Date = Now.AddMonths(-1)
            Me.datFrom = DateSerial(tmpDate.Year, tmpDate.Month, 1)
            Me.datTo = Now.AddMonths(3)
            Me.mnuDate.Text = "From &Date - " & datFrom.ToString("dd-MMM-yyyy")
            Me.mnuDate.Checked = True
            Me.timRefresh.Interval = 1000
            'GetCollectionDetails()
            Me.blnDataLoaded = True

        End If
    End Sub

    Private Sub EmailTextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailTextToolStripMenuItem.Click
        Dim elements As Windows.Forms.HtmlElementCollection = Me.HtmlEditor1.Document.All

        Dim strBody = "<HTML>" & elements.Item(0).InnerHtml & "</HTML>"
        Dim stremail As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail

        If blnEditorOther Then
            Medscreen.BlatEmail("Report for User ", strBody, stremail, , True)
        Else
            Medscreen.BlatEmail("Report for collection " & Me.lvItemCurrent.Collection.ID, strBody, stremail, , True)
        End If

    End Sub


    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Me.HtmlEditor1.ShowPrintDialog()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_click(ByVal dender As Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Me.HtmlEditor1.ShowPrintPreviewDialog()

    End Sub

    Private Sub CurrentDocument_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles CurrentDocument.MouseMove
        Try
            MousePoint = New Point(e.MousePosition.X, e.MousePosition.Y)
            Dim CurrentElement As HtmlElement = CurrentDocument.GetElementFromPoint(MousePoint)
            If CurrentElement.TagName = "A" Then
                lblLink.Text = Medscreen.ReplaceString(CurrentElement.GetAttribute("href"), "&", "&&")
                lblLink.Refresh()
                Application.DoEvents()
            Else
                lblLink.Text = ""
                lblLink.Refresh()
                Application.DoEvents()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        If Not myCache Is Nothing Then
            If myCache.Position > 0 Then
                'myCache.Position -= 1
                Me.HtmlEditor1.Navigate(myCache.GetElement.HyperLink)
                Me.txtNavigate.Text = myCache.GetElement.HyperLink
                txtNavigate.Refresh()
                Application.DoEvents()
            End If

        End If

    End Sub

    Private Sub HtmlEditor1_NewWindow(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles HtmlEditor1.NewWindow

    End Sub

    Private Sub HtmlEditor1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles HtmlEditor1.PreviewKeyDown
        If e.KeyCode = Keys.F AndAlso e.Control Then
            Debug.Print("here")
            cmdFind_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub HtmlEditor1_ProgressChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserProgressChangedEventArgs) Handles HtmlEditor1.ProgressChanged
        Me.ProgressBar1.Maximum = Convert.ToInt32(e.MaximumProgress)
        Me.ProgressBar1.Value = Convert.ToInt32(e.CurrentProgress)
    End Sub

    Private Sub CurrentDocument_MouseOver(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles CurrentDocument.MouseOver

    End Sub

    Private Sub txtNavigate_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNavigate.DoubleClick
        HtmlEditor1.Navigate(txtNavigate.Text)
    End Sub

    Private Sub txtNavigate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNavigate.KeyPress
        If e.KeyChar = vbCr Then
            HtmlEditor1.Navigate(txtNavigate.Text)
        End If
    End Sub

    Private Sub frmDiary_ContextMenuChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ContextMenuChanged

    End Sub

    Private Sub ctxMenu_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctxMenu.Popup
        Debug.Print("CTX popped up " & Now.ToLongTimeString)
    End Sub

    Private Sub frmDiary_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyData = Keys.F Then
            Debug.Print("here")
        End If
    End Sub


    Private Sub ListSortBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListSortBox.SelectedIndexChanged
        If ListSortBox.SelectedIndex = 0 Then
            MedscreenCommonGui.My.Settings.OfficerStyleSheet = "Collection2.xsl"
        ElseIf ListSortBox.SelectedIndex = 1 Then
            MedscreenCommonGui.My.Settings.OfficerStyleSheet = "Collection2CollDate.xsl"
        End If
        MedscreenCommonGui.My.Settings.Save()
    End Sub

    Private Sub CreateCollFromCardiffTicket(ByVal objTicket As string)
        Dim CollInfo As String = BAOCardiff2.Collection.CollectionInfo(objTicket)
        If CollInfo IsNot Nothing AndAlso CollInfo.Length > 0 Then
            Dim DCCode As String = CConnection.PackageStringList("Lib_collectionxml8.getcollectioninfo", "TT" & objTicket.ToString.Trim)
            'check we don't have a collection
            Dim CollElems As String() = CollInfo.Split(New Char() {","})
            If DCCode.Length > 2 AndAlso DCCode.Chars(0) = "," AndAlso DCCode.Chars(1) = "," Then
                Dim cm As New Intranet.intranet.jobs.CMJob()
                cm.SetId = "TT" & objTicket.ToString.Trim

                If CollElems.Length > 3 Then
                    Dim strOFFID As String = CConnection.PackageStringList("GetOfficerByToastID", CollElems(4))
                    If strOFFID IsNot Nothing Then
                        cm.CollOfficer = strOFFID
                    End If
                    cm.ClientId = CollElems(2)
                    cm.Invoiceid = CollElems(0)

                End If
                If CollElems.Length > 16 Then
                    cm.JobName = CollElems(11)
                    cm.CollectionType = CollElems(15)
                    cm.Status = CollElems(16)
                End If
                cm.Update()
                Dim Item As ListViewItems.ListViewCMDiaryItem

                Item = Me.BuildItem(cm.ID)
                Me.lvDiaryList.Items.Add(Item)          'Create a new row
                Me.lvDiaryList.SelectedItems.Clear()    'Make it the selected item
                Item.Selected = True
                Item.EnsureVisible()
                Me.lvDiaryList_Click(Nothing, Nothing)  'Try and get everything displayed
                lvItemCurrent.Collection.LogAction(jobs.CMJob.Actions.CreateCollection)

            End If
        End If
    End Sub
    Private Sub FromCardiffTicketToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromCardiffTicketToolStripMenuItem.Click
        'Get ticket and then get Ticket data 
        Dim objTicket As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Cardiff Ticket", "Cardiff Ticket")
        If Not objTicket Is Nothing Then

            CreateCollFromCardiffTicket(objTicket)
        End If
    End Sub
End Class


