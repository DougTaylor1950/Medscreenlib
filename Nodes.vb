Imports System.Windows.Forms
Imports System.Reflection
Imports System.Drawing
Imports MedscreenLib
Imports MedscreenLib.Glossary
Imports Intranet.intranet
Imports Intranet.intranet.customerns

Namespace Treenodes


#Region "TreeNodes"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : MedTreeNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A derived node used in various GUI Elements
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class MedTreeNode
        Inherits TreeNode
        ''' <summary>Standard select icon</summary>
        Public Const cstIconSelect1 As Integer = 14
        Private Const cstIconSelect2 As Integer = 15
        Public Event NodeChanged()

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory></revisionHistory>
        Public Property User() As String
            Get
                If TypeOf Me.TreeView Is CustomTreeBase Then
                    Dim ct As CustomTreeBase = Me.TreeView
                    Return ct.User
                Else
                    Return ""
                End If

            End Get
            Set(ByVal value As String)
                If TypeOf Me.TreeView Is CustomTreeBase Then
                    Dim ct As CustomTreeBase = Me.TreeView
                    ct.User = value
                End If
            End Set
        End Property

        Public Property CurrentCustomerViewState() As Treenodes.AlphaNodeList.CustomerViewState
            Get
                If TypeOf Me.TreeView Is CustomTreeBase Then
                    Dim ct As CustomTreeBase = Me.TreeView
                    Return ct.AlphaNodes.CurrentViewState
                Else
                    Dim objTrenodes As TreeNode
                    For Each objTrenodes In Me.TreeView.Nodes
                        If TypeOf objTrenodes Is AlphaNodeList Then
                            Dim objAlpha As AlphaNodeList = objTrenodes
                            Return objAlpha.CurrentViewState
                        End If
                    Next
                End If
            End Get
            Set(ByVal Value As Treenodes.AlphaNodeList.CustomerViewState)

            End Set
        End Property

        Protected Sub OnNodeChanged()
            RaiseEvent NodeChanged()
        End Sub

        Public Overloads Sub DecorateNode(ByVal SourcePhrase As Phrase)
            If SourcePhrase Is Nothing Then Exit Sub
            If SourcePhrase.Formatter Is Nothing Then Exit Sub
            SourcePhrase.Formatter.Decorate(Me)
            'Me.ForeColor = SourcePhrase.ForeColour
            'Me.BackColor = SourcePhrase.backColour
            'If SourcePhrase.ImageIndex > -1 Then Me.ImageIndex = SourcePhrase.ImageIndex
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new tree node 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.SelectedImageIndex = cstIconSelect1

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new tree node with Text 
        ''' </summary>
        ''' <param name="strText">Supplied text</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal strText As String)
            MyBase.New(strText)
            Me.SelectedImageIndex = cstIconSelect1

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert node to html, place holder for overrides
        ''' </summary>
        ''' <returns>Node converted to HTML</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overridable Function ToHTML() As String
            Dim strXML As String = "<A></A>"
            Dim strMedtreeNodeStyle As String = MedscreenCommonGUIConfig.NodeStyleSheets.Item("MedtreeNodeStyle")
            Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strMedtreeNodeStyle, 0)
            Return strHTML

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Look for a child with a particular text string
        ''' </summary>
        ''' <param name="strText">Text string to look for</param>
        ''' <returns>Position of node</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function HasChild(ByVal strText As String) As Integer
            Dim aNode As TreeNode
            Dim blnRet As Boolean = False

            Dim index As Integer = 0
            For Each aNode In Me.Nodes
                If aNode.Text.ToUpper = strText.ToUpper Then
                    blnRet = True
                    Exit For
                End If
                index += 1
            Next
            If Not blnRet Then index = -1
            Return index
        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.AlphaNodeList
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection of AlphaNodes 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class AlphaNodeList
        Inherits MedTreeNode

#Region "Enumerations"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Possible view states for the application
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Enum CustomerViewState
            ''' <summary>Normal view, exclude Deleted accounts</summary>
            ViewNormal = 1
            ''' <summary>View all accounts</summary>
            ViewAll = 2
            ''' <summary>View New accounts only</summary>
            ViewNew = 3
            ''' <summary>View Active Jobs</summary>
            ViewActiveJobs = 4
            '''' <summary>View LUL Contractors Only</summary>
            'ViewLULOnly = 5
            '''' <summary>View Esso Contractors Only</summary>
            'ViewEssoOnly = 6
            '''' <summary>View Rail Track customers only</summary>
            'ViewRailTrackOnly = 7
            ''' <summary>View Persons accounts</summary>
            ViewMine = 8
            ViewIndustry = 9
            ViewBySage = 10
        End Enum

        Public Enum NodeListType
            Smid
            Sage
        End Enum
#End Region

        Private myListType As NodeListType = NodeListType.Smid

        Public Property ListType() As NodeListType
            Get
                Return myListType
            End Get
            Set(ByVal Value As NodeListType)
                myListType = Value
            End Set
        End Property


        Private myCurrentViewState As Treenodes.AlphaNodeList.CustomerViewState = Treenodes.AlphaNodeList.CustomerViewState.ViewNormal

        Public Property CurrentViewState() As Treenodes.AlphaNodeList.CustomerViewState
            Get
                Return myCurrentViewState
            End Get
            Set(ByVal Value As Treenodes.AlphaNodeList.CustomerViewState)
                myCurrentViewState = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new List 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [22/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal HeaderText As String)
            MyBase.New(HeaderText)
            'Dim i As Integer

            'For i = Asc("A") To Asc("Z")                'Create A - Z nodes 
            '    Me.Nodes.Add(New Treenodes.AlphaNode(Chr(i)))
            'Next
            Me.Refresh()

        End Sub

        Public Function Refresh() As Boolean
            Dim strList As String = ""
            If Me.ListType = NodeListType.Sage Then
                strList = CConnection.PackageStringList("lib_cctool.GetSageLetters", "")
            Else
                strList = CConnection.PackageStringList("lib_cctool.GetSmidLetters", "")
            End If
            Try
                Me.Nodes.Clear()
                Dim strLetterArray As String() = strList.Split(New Char() {","})

                Dim Letter As String
                For Each Letter In strLetterArray
                    If Letter.Trim.Length > 0 Then Me.Nodes.Add(New AlphaNode(Letter, Me.ListType))
                    Debug.WriteLine(Letter)
                Next
            Catch ex As Exception
            End Try
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Find a header node for a particular letter of the alphabet
        ''' </summary>
        ''' <param name="HeaderChar">Letter of alphabet</param>
        ''' <returns></returns>
        ''' <remarks>Used to get to right letter in alpha nodes
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function FindHeaderNode(ByVal HeaderChar As Char) As TreeNode
            'Dim i As Integer
            Dim tNode As TreeNode = Nothing
            For Each tNode In Me.Nodes
                If tNode.Text = HeaderChar Then
                    Exit For
                End If
                tNode = Nothing
            Next
            If Not tNode Is Nothing Then
                If TypeOf tNode Is Treenodes.AlphaNode Then
                    Return CType(tNode, Treenodes.AlphaNode)
                Else
                    Return tNode
                End If
            Else
                Return Nothing
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get number of active jobs for customer
        ''' </summary>
        ''' <param name="strCust">Customer</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Function GetJobCount(ByVal strCust As String) As Integer
            '    Dim aJob As ActiveCustomerJob
            '    aJob = ActiveJobs.Item(strCust)
            '    If aJob Is Nothing Then
            '        Return 0
            '    Else
            '        Return aJob.JobCount()
            '    End If
            Dim strActiveJobCount As String = MedscreenCommonGUIConfig.NodePLSQL.Item("ActiveJobCount")
            Return CConnection.PackageStringList(strActiveJobCount, strCust)
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add customer node by SMID
        ''' </summary>
        ''' <param name="myClient">Smid object</param>
        ''' <param name="HeaderChar">For header</param>
        ''' <param name="hNode">Header node</param>
        ''' <param name="pNode">Returned node</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [21/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Sub addCustomerNode(ByVal myClient As Intranet.intranet.customerns.SMID, _
         ByRef HeaderChar As Char, ByRef hNode As Treenodes.AlphaNode, ByVal pNode As Treenodes.CustomerNodes.CustNode)


            Try
                'Debug.WriteLine(myClient.SMID)
                'LogTiming("Adding " & myClient.SMID)
                Dim cNode As Treenodes.CustomerNodes.CustSMIDNode
                'If Not myClient.SMID = "TEMPLATES" Then

                cNode = New Treenodes.CustomerNodes.CustSMIDNode(myClient)
                'End If
                Dim HeaderChar2 As Char = Mid(myClient.SMID, 1, 1)
                hNode = Me.FindHeaderNode(HeaderChar2)

                If hNode Is Nothing Then            'We don't have a header node so add it 
                    hNode = New Treenodes.AlphaNode(HeaderChar2, Me.ListType)
                    Me.Nodes.Add(hNode)
                End If
                If Not cNode Is Nothing Then
                    hNode.Nodes.Add(cNode)
                End If


            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex)
            End Try
        End Sub


    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.AlphaNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Class that contains a set of Medscreen SMID Profiles
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class AlphaNode
        Inherits MedTreeNode
        Private myLetter As String
        Private myMenu As MedscreenLib.DynamicContextMenu
        Private myCollMenu As menus.AlphaMenu
        Private myListType As AlphaNodeList.NodeListType
        Friend WithEvents ctxMenu As System.Windows.Forms.ContextMenu

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new node
        ''' </summary>
        ''' <param name="AlphaLetter"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [22/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal AlphaLetter As String, ByVal ListType As AlphaNodeList.NodeListType)
            MyBase.new()
            myLetter = AlphaLetter
            Me.Text = myLetter
            myListType = ListType
            Try
                Dim MyAssembly As [Assembly]
                MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
                Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
                myCollMenu = menus.AlphaMenu.GetAlphaMenu
                myMenu = New MedscreenLib.DynamicContextMenu(directory & "\AlphaContextMenu.xml", myCollMenu)
                ctxMenu = myMenu.LoadDynamicMenu()
            Catch ex As Exception
            End Try

        End Sub

        'expose context menu 
        Public ReadOnly Property ContextMenu() As ContextMenu
            Get
                Return Me.ctxMenu
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Function to find SMID 
        ''' </summary>
        ''' <param name="SMID"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	29/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function FindSimdNode(ByVal SMID As String) As Treenodes.CustomerNodes.CustSMIDNode
            Dim aSMIDNode As Treenodes.CustomerNodes.CustSMIDNode = Nothing
            Dim aTreeNode As TreeNode

            If Me.Nodes.Count = 0 Then Me.Refresh()
            'go through all nodes looking for SMID nodes
            For Each aTreeNode In Me.Nodes
                If TypeOf aTreeNode Is Treenodes.CustomerNodes.CustSMIDNode Then
                    aSMIDNode = aTreeNode
                    If aSMIDNode.SMID.SMID = SMID Then
                        Exit For
                    End If
                    aSMIDNode = Nothing
                End If
            Next
            Return aSMIDNode
        End Function

        Public Function FindSageNode(ByVal sageName As String) As Treenodes.CustomerNodes.CustSMIDNode
            Dim aSMIDNode As Treenodes.CustomerNodes.CustSMIDNode = Nothing
            Dim aTreeNode As TreeNode
            If Me.Nodes.Count = 0 Then Me.Refresh()

            'go through all nodes looking for SMID nodes
            For Each aTreeNode In Me.Nodes
                If TypeOf aTreeNode Is Treenodes.CustomerNodes.CustSMIDNode Then
                    aSMIDNode = aTreeNode
                    If aSMIDNode.SMID.SMID = sageName Then
                        Exit For
                    End If
                    aSMIDNode = Nothing
                End If
            Next
            Return aSMIDNode
        End Function

        Private Function GetNodeList(ByVal Start As String) As String
            Dim strNodeList As String = ""
            Dim oColl As New Collection()
            oColl.Add(Me.myLetter)
            oColl.Add(Start)
            If myListType = AlphaNodeList.NodeListType.Smid Then
                Dim ViewState As AlphaNodeList.CustomerViewState = CType(Me.Parent, AlphaNodeList).CurrentViewState
                If ViewState = AlphaNodeList.CustomerViewState.ViewNormal Then
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetSmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewAll Then
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetFullSmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewNew Then
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetNewSmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewIndustry Then
                    oColl.Add(CType(Me.TreeView, CustomTreeBase).IndustryCode)
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetIndustrySmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewActiveJobs Then
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetActiveSmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewMine Then
                    Dim strUser As String = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    If strUser = "TAYLOR" Then strUser = "DAWN"
                    oColl.Add(strUser)
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetMySmidList", oColl)
                ElseIf ViewState = AlphaNodeList.CustomerViewState.ViewBySage Then
                    strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetSageList", oColl)

                End If
            Else
                strNodeList = CConnection.PackageStringList("lib_cctoolviews.GetSageList", oColl)
            End If
            Return strNodeList
        End Function

        Private Sub ProcessNodeArray(ByVal NodeArray As String(), ByVal EntityType As Intranet.intranet.customerns.SMID.SMIDEntityType)
            Dim Smid As String
            Dim LastSmid As String = ""
            For Each Smid In NodeArray
                If Smid <> "TEMPLATES" Then
                    If Smid <> "+" Then
                        Dim aSMIDNode As New Treenodes.CustomerNodes.CustSMIDNode(New Intranet.intranet.customerns.SMID(Smid, EntityType))
                        Me.Nodes.Add(aSMIDNode)
                        LastSmid = Smid
                    Else
                        Dim strNodeList As String = GetNodeList(LastSmid)
                        Dim NewNodeArray As String() = strNodeList.Split(New Char() {","})
                        ProcessNodeArray(NewNodeArray, EntityType)
                    End If
                End If
            Next

        End Sub

        Public Function Refresh() As Boolean

            Dim strNodeList As String = GetNodeList("")
            Dim NodeArray As String() = strNodeList.Split(New Char() {","})
            Dim entityType As Intranet.intranet.customerns.SMID.SMIDEntityType = SMID.SMIDEntityType.SMID
            Dim ViewState As AlphaNodeList.CustomerViewState = CType(Me.Parent, AlphaNodeList).CurrentViewState
            If ViewState = AlphaNodeList.CustomerViewState.ViewBySage Then entityType = SMID.SMIDEntityType.SageName
            Me.Nodes.Clear()

            ProcessNodeArray(NodeArray, entityType)
        End Function

        Public Function RefreshNodes(ByVal ShowComment As Boolean) As Boolean
            'Dim sNode As Treenodes.CustomerNodes.CustSMIDNode
            'For Each sNode In Me.Nodes
            '    sNode.ShowComment = ShowComment
            'sNode.SetNodeStatus()
            '    Debug.WriteLine(sNode.Text)
            'Next
            Return True
        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CrystalReports
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Crystal Reports
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CrystalReportsHeader
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode

        Public Sub New()
            MyBase.New("Crystal Reports")
        End Sub

        Public Function ToHTML() As String
            Dim strRet As String = ""
            Dim strXML As String = "<CrystalReports>"
            For Each anode As TreeNode In Me.Nodes
                If TypeOf anode Is CrystalReportsNode Then
                    Dim crnode As CrystalReportsNode = anode
                    strXML += crnode.Report.ToXML & vbCrLf
                ElseIf TypeOf anode Is CrystalReportsHeader Then
                    Dim crnode As CrystalReportsHeader = anode
                    strXML += "<CrystalReport><menutext>" & crnode.Text & "</menutext></CrystalReport>"

                End If
            Next
            strXML += "</CrystalReports>"
            strRet = Medscreen.ResolveStyleSheet(strXML, "CrystalReports.xsl", 0)
            Return strRet
        End Function
    End Class

    Public Class CrystalReportsNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        Private objCr As MedscreenLib.CRMenuItem

        Public Sub New(ByVal Cr As MedscreenLib.CRMenuItem)
            MyBase.New(Cr.MenuText)
            objCr = Cr
            Select Case objCr.MenuType
                Case "MNU"
                    Me.ImageIndex = 20
                Case "RPT", "RPTX", "RPTP"
                    Me.ImageIndex = 21
                Case Else
                    Me.ImageIndex = 2
            End Select
        End Sub

        Public Property Report() As MedscreenLib.CRMenuItem
            Get
                Return objCr
            End Get
            Set(ByVal Value As MedscreenLib.CRMenuItem)
                objCr = Value
            End Set
        End Property

        Public Function ToHTML() As String
            Dim strRet As String = ""
            Return strRet
        End Function

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CustomerReportNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Customer Report node
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CustomerReportNode
        Inherits Treenodes.ReportHeaderNode

        Private myReports As MedscreenLib.ReportSchedules

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create customer report node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/04/2007]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("Customer Reports")
            myReports = New MedscreenLib.ReportSchedules()
            myReports.Load()
            Dim objRepSched As MedscreenLib.ReportSchedule
            For Each objRepSched In myReports
                If objRepSched.ScheduleType = MedscreenLib.ReportSchedule.GCST_CUSTOMER_REPORT Then
                    Dim objRepNode As New Treenodes.ReportNode(objRepSched, True)
                    Me.Nodes.Add(objRepNode)
                End If
            Next
        End Sub
    End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.BackgroundReportNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Background Scheduled Reports
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class BackgroundReportNode
        Inherits Treenodes.ReportHeaderNode

        Private myReports As MedscreenLib.ReportSchedules

        Public Property Reports() As MedscreenLib.ReportSchedules
            Get
                Return myReports
            End Get
            Set(ByVal value As MedscreenLib.ReportSchedules)
                myReports = value
            End Set
        End Property

        Public Sub New()
            MyBase.New("Background Reports")
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Me.Nodes.Clear()
            myReports = New MedscreenLib.ReportSchedules()
            myReports.Load()
            Dim objRepSched As MedscreenLib.ReportSchedule
            For Each objRepSched In myReports
                If objRepSched.ScheduleType = MedscreenLib.ReportSchedule.GCST_BACKGROUND_REPORT Then
                    Dim objRepNode As New Treenodes.ReportNode(objRepSched)
                    Me.Nodes.Add(objRepNode)
                ElseIf objRepSched.ScheduleType = MedscreenLib.ReportSchedule.GCST_SOAP Then
                    Dim objRepNode As New Treenodes.ReportNode(objRepSched)
                    Me.Nodes.Add(objRepNode)
                End If
            Next
        End Function
    End Class

    Public Class ReportHeaderSMIDNode
        Inherits ReportHeaderNode
        Private mySMID As String

        Public Function ToHTML() As String
            Dim strXML As String = ToXML()

        End Function

        Public Function ToXML() As String
            Dim strRet As String = "<ReportScheules>"
            For Each aNode As TreeNode In Nodes
                If TypeOf aNode Is ReportNode Then
                    Dim rNode As ReportNode = aNode
                    strRet += rNode.ReportSchedule.ToXMLElement.ToString
                End If
            Next
            strRet += "</ReportScheules>"
            Return strRet

        End Function


        Public Sub New(ByVal SMID As String)
            MyBase.New()
            mySMID = SMID
            Me.Text = "Scheduled Reports for " & SMID
            Dim strQuery As String = " select schedule_id from report_schedule where customer_id in (select identity from customer where removeflag = 'F' and smid = ?)"
            Dim oCmd As New OleDb.OleDbCommand(strQuery, CConnection.DbConnection)
            Dim oRead As OleDb.OleDbDataReader
            Dim oColl As New Collection
            Try
                oCmd.Parameters.Add(CConnection.StringParameter("SMID", SMID, 10))
                If CConnection.ConnOpen Then
                    oRead = oCmd.ExecuteReader
                    While oRead.Read
                        oColl.Add(oRead.GetValue(0))
                    End While
                End If
                For Each schedid As String In oColl
                    Dim RepSched As New ReportSchedule(schedid)
                    Dim SchedNode As New ReportNode(RepSched, True)
                    Me.Nodes.Add(SchedNode)
                Next

            Catch ex As Exception
            Finally

            End Try


        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.ReportHeaderNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Header node for scheduled reports
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ReportHeaderNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new report header node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.ImageIndex = 4
            Me.Text = "Scheduled Reports"
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new report header node with descriptive text 
        ''' </summary>
        ''' <param name="text"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal text As String)
            MyBase.new(text)
            Me.ImageIndex = 4
        End Sub
    End Class

    Public Class OrdersHeaderNode
        Inherits InvoiceHeaderNode
        Public Sub New()
            MyBase.new()
            Me.Text = "Orders"
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.InvoiceHeaderNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Header node for invoices
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class InvoiceHeaderNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        Private myInvoices As Intranet.intranet.jobs.Invoices

        Public Property invoices() As Intranet.intranet.jobs.Invoices
            Get
                If myInvoices Is Nothing Then myInvoices = New Intranet.intranet.jobs.Invoices()
                Return myInvoices
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.Invoices)
                myInvoices = Value
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new invoice header node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.Text = "Invoices"
            Me.SelectedImageIndex = cstIconSelect1
            Me.ImageIndex = 26
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new invoice header node with text 
        ''' </summary>
        ''' <param name="text"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal text As String)
            MyBase.new(text)
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.SampleHeaderNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A sample header node 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/07/2009]</date><Action>Added invoice property</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class SampleHeaderNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        'Private AcSampList As ActiveSampleList
        Private CollSampList As Intranet.intranet.sample.oSampleCollection
        Private strJob As String
        Private strClient As String
        Private myContent As ContentType = ContentType.job
        Private strTable As String = "Sample"
        Private myWhereClause As String
        Private myInvoiceNumber As String = ""
        Private Enum ContentType
            none
            job
            samplelist
            aoClient
        End Enum


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Currently processed invoice number for customer (acts as a container)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property InvoiceNumber() As String
            Get
                Return Me.myInvoiceNumber
            End Get
            Set(ByVal Value As String)
                myInvoiceNumber = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Position Barcode to right node
        ''' </summary>
        ''' <param name="Barcode"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/03/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function PositionBarcode(ByVal Barcode As String) As SampleNode
            Dim iNode As SampleNode = Nothing
            For Each iNode In Me.Nodes
                If Not iNode.sample Is Nothing AndAlso iNode.sample.Barcode = Barcode Then
                    iNode.EnsureVisible()

                    Exit For
                End If
                iNode = Nothing
            Next
            If iNode Is Nothing Then
                Dim objSamp As New Intranet.intranet.sample.OSample(Barcode, Intranet.intranet.sample.OSample.CreateWith.Barcode)
                If Not objSamp Is Nothing AndAlso Not Me.CollSampList Is Nothing Then
                    Me.CollSampList.Add(objSamp)
                    iNode = New SampleNode(objSamp)
                    Me.Nodes.Add(iNode)
                End If
            End If
            Return iNode
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new sample header node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.ImageIndex = 33
            Me.SelectedImageIndex = cstIconSelect1
            Me.Text = "Samples"
            myContent = ContentType.none
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new sample header node
        ''' </summary>
        ''' <param name="Job"></param>
        ''' <param name="TableSet"></param>
        ''' <remarks>
        ''' If no tableset will be analysis only
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Job As String, ByVal TableSet As String)
            MyClass.New()
            Me.strJob = Job
            If TableSet Is Nothing Then
                Me.Text = "Analysis Only Samples "
                myContent = ContentType.aoClient
                Me.strTable = "SAMPLE"
            ElseIf TableSet.ToUpper.Trim = "SAMPLE" Then
                Me.strTable = TableSet
                Me.Text = "Samples - Active - " & strJob
                myContent = ContentType.job

            ElseIf TableSet.Trim.Length > 0 And TableSet <> "0" Then
                Me.strTable = "C_sample" & TableSet
                Me.Text = "Samples - " & strJob
                myContent = ContentType.job
            End If
        End Sub



        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new sample header node with a collection of samples
        ''' </summary>
        ''' <param name="ACSamples">Collection of samples <see cref="Intranet.intranet.sample.oSampleCollection"/></param>
        ''' <remarks>Will create the appropriate sample nodes
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal ACSamples As Intranet.intranet.sample.oSampleCollection)
            MyBase.new("Samples ")
            CollSampList = ACSamples
            Me.ImageIndex = 33
            Me.SelectedImageIndex = cstIconSelect1
            BuildList()
            Me.Text = "Samples " & CollSampList.Count
            If Not ACSamples Is Nothing Then Me.myWhereClause = ACSamples.WhereClause
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Build a list of nodes
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub BuildList()
            Dim j As Integer
            Dim acSamp As Intranet.intranet.sample.OSample
            Dim saNode As SampleNode
            Me.Nodes.Clear()
            myContent = ContentType.samplelist
            If CollSampList Is Nothing Then Exit Sub
            For j = 0 To CollSampList.Count - 1
                acSamp = CollSampList.Item(j)
                saNode = New SampleNode(acSamp)
                Me.Nodes.Add(saNode)

            Next

        End Sub
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the sample nodes
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Try
                Me.Nodes.Clear()
            Catch

            End Try
            Dim strQuery As String
            If Me.myContent = ContentType.samplelist Then
                If Me.CollSampList Is Nothing Then Exit Function
                Me.CollSampList.Refresh()           'Refresh the collection
                Me.BuildList()
                Me.Text = "Samples " & CollSampList.Count  'Update the text item
            ElseIf Me.myContent = ContentType.job Then
                strQuery = MedscreenCommonGUIConfig.NodeQuery.Item("SamplesByJob")
                strQuery = String.Format(strQuery, strTable, strJob)
                FillList(strQuery)
            ElseIf Me.myContent = ContentType.aoClient Then
                If strJob Is Nothing OrElse strJob.Trim.Length = 0 Then
                    If TypeOf Me.Parent Is CustomerNodes.CustNode Then
                        strJob = CType(Me.Parent, CustomerNodes.CustNode).client.Identity
                    End If
                End If

                strQuery = MedscreenCommonGUIConfig.NodeQuery.Item("AnalysisOnlySamples")
                strQuery = String.Format(strQuery, strJob, Today.AddMonths(-1).ToString("dd-MMM-yy"))
                FillList(strQuery)

            End If
            BuildList()
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Fill list of samples using a supplied query
        ''' </summary>
        ''' <param name="strQuery"></param>
        ''' <remarks>
        ''' Queries for content type are held in config file see Refresh function for call
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub FillList(ByVal strQuery As String)
            Dim AColl As New Collection()
            Dim bColl As New Collection()
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            Try
                oCmd.Connection = CConnection.DbConnection
                oCmd.CommandText = strQuery
                Me.CollSampList = New Intranet.intranet.sample.oSampleCollection()

                If CConnection.ConnOpen Then
                    oRead = oCmd.ExecuteReader
                    While oRead.Read
                        If Not oRead.IsDBNull(0) Then
                            AColl.Add(oRead.GetValue(0))
                            bColl.Add(oRead.GetValue(1))
                        End If
                    End While
                End If
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If

                Dim i As Integer
                For i = 1 To AColl.Count
                    Dim oSamp As New Intranet.intranet.sample.OSample(AColl.Item(i), CStr(bColl.Item(i)))
                    Me.CollSampList.Add(oSamp)
                Next
            Catch ex As Exception
                Medscreen.LogError(ex)
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()


            End Try

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.SampleNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A node relating to a sample
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class SampleNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        'Private objSample As ActiveSample
        Private oSample As Intranet.intranet.sample.OSample
        Private ctxMenu As ContextMenu
        Public Event OnCustomerChange(ByVal Barcode As String)
        Private LabwareSampleCase As BAOLabware.SampleCase = Nothing

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Text to be put on a tooltip
        ''' </summary>
        ''' <returns>Text to be put on a tooltip</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToolTipText() As String
            Dim strText As String
            If Mid(sample.TestSchedule, 1, 3) = "COZ" Then
                GetLabwareSample()
                If LabwareSampleCase IsNot Nothing Then
                    strText = Me.Text
                End If
            Else

                strText = Me.Text
                If Date.Compare(oSample.DateReceived, DateSerial(1970, 1, 1)) > 0 Then
                    strText += vbCrLf & "Received : " & oSample.DateReceived.ToString("dd-MMM-yyyy HH:mm")

                End If
                If Date.Compare(oSample.LoginDate, DateSerial(1970, 1, 1)) > 0 Then
                    strText += vbCrLf & "Logged in : " & oSample.LoginDate.ToString("dd-MMM-yyyy HH:mm") & _
                    " loggedin by " & oSample.LoggedInBy
                End If
                If Date.Compare(oSample.DateStarted, DateSerial(1970, 1, 1)) > 0 Then
                    strText += vbCrLf & "Started : " & oSample.DateStarted.ToString("dd-MMM-yyyy HH:mm") & _
                    " Started by " & oSample.StartedBy
                End If
                If oSample.InvoiceNumber.Trim.Length > 0 Then
                    strText += vbCrLf & "Invoice Number : " & oSample.InvoiceNumber
                End If
            End If
            Return strText
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh sample and node
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Dim strText As String = ""
            If Mid(sample.TestSchedule, 1, 3) = "COZ" Then
                GetLabwareSample()
                If LabwareSampleCase IsNot Nothing Then
                    strText = LabwareSampleCase.Barcode & " - " & _
                    LabwareSampleCase.Status & _
                    " DueOff - "
                    If LabwareSampleCase.DueOutDate.Year > 1900 Then
                        strText += LabwareSampleCase.DueOutDate.ToString("dd-MMM-yyyy")
                    Else
                        strText += "Not defined"

                    End If
                    Text = strText
                End If
            Else

                strText = oSample.Barcode & " - " & _
                    oSample.StatusDescription & _
                    " DueOff - "
                If oSample.DateResultsRequired > DateField.ZeroDate Then
                    strText += oSample.DateResultsRequired.ToString("dd-MMM-yyyy")
                Else
                    strText += "Not defined"
                End If
                Text = strText
                Select Case oSample.Status
                    Case "A", "P"
                        Me.ForeColor = System.Drawing.Color.Green
                    Case "C"
                        Me.ForeColor = System.Drawing.Color.Orange
                    Case "M", "X", "H"
                        Me.ForeColor = System.Drawing.Color.Yellow
                    Case "X"
                        Me.ForeColor = System.Drawing.Color.Red
                    Case Else
                        Me.ForeColor = System.Drawing.SystemColors.ControlText
                End Select
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new sample node 
        ''' </summary>
        ''' <param name="sample">Sample for node</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal sample As Intranet.intranet.sample.OSample)
            MyBase.New()
            oSample = sample
            Me.ImageIndex = 10
            Try
                Dim MyAssembly As [Assembly]
                MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
                Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
                Dim mySampleMenu As menus.SampleMenu = menus.SampleMenu.GetSampleMenu
                Dim myMenu As New MedscreenLib.DynamicContextMenu(directory & "\SampleContextMenu.xml", mySampleMenu)
                ctxMenu = myMenu.LoadDynamicMenu()
            Catch ex As Exception
            End Try

            Me.Refresh()
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose the context menu 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ContextMenu() As ContextMenu
            Get
                Return Me.ctxMenu
            End Get
        End Property

        Private Sub DrawMenuItem(ByVal SampleState As String, ByVal SendMethod As String, ByVal HasRole As Boolean, ByVal hasImportRole As Boolean, ByVal mnuItem As TaggedMenuItem)
            mnuItem.Visible = True
            mnuItem.Enabled = False

            Debug.Print(mnuItem.Text & "-" & mnuItem.Group)

            If mnuItem.Group = "IMPORT" Then
                If Not hasImportRole Then                       'nOT ALLOWED TO IMPORT
                    mnuItem.Enabled = False
                Else
                    If sample.Barcode.Chars(0) <> "H" AndAlso mnuItem.MenuItems.Count = 0 Then
                        mnuItem.Enabled = False
                    Else
                        mnuItem.Enabled = True
                    End If
                End If
            Else


                If mnuItem.Mnemonic = "A" AndAlso InStr("ARC", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "U" AndAlso InStr("ARC", SampleState) > 0 AndAlso Me.sample.Status = "X" AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "R" AndAlso InStr("RC", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "C" AndAlso InStr("C", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "I" AndAlso InStr("ARC", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "H" AndAlso InStr("RC", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "T" AndAlso InStr("C", SampleState) > 0 AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "L" Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "E" Then
                    mnuItem.Enabled = (sample.DateReported > DateField.ZeroDate AndAlso sample.Status <> "M" AndAlso sample.HasCertificate)
                ElseIf mnuItem.Mnemonic = "M" AndAlso HasRole Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Mnemonic = "O" Then
                    mnuItem.Enabled = (sample.Status = "M" AndAlso sample.HasMRORequest)
                ElseIf mnuItem.Mnemonic = "S" Then
                    mnuItem.Enabled = (sample.HasScreenResult)
                ElseIf mnuItem.Mnemonic = "F" AndAlso (SendMethod = "MANFAX" Or SendMethod = "FAX") Then
                    mnuItem.Enabled = True
                ElseIf Char.IsDigit(mnuItem.Mnemonic) Then
                    mnuItem.Enabled = True
                ElseIf mnuItem.Group = "BUPA" Then
                    'Check to see if BUPACENTRE
                    mnuItem.Enabled = False
                    mnuItem.Visible = False
                    Dim strRet As String = CConnection.PackageStringList("lib_sample8.IsBUPACENTRECustomer", Me.sample.CustomerID)
                    If strRet = "T" Then
                        If sample.DonorIdentifier.Trim.Length = 0 Then
                            mnuItem.Enabled = (sample.DateReported > DateField.ZeroDate AndAlso sample.Status <> "M" AndAlso sample.HasCertificate)
                            mnuItem.Visible = True
                        End If

                    End If
                Else
                    mnuItem.Enabled = False
                End If
            End If
        End Sub
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' redraw the contextmenu taking into accout any changes there may have been to the sample state 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub ContextmenuRedraw()
            Dim SampleState As String = CConnection.PackageStringList("lib_Uncommit.GetSampleStatus", Me.sample.Barcode)
            Dim SendMethod As String = CConnection.PackageStringList("Lib_sample8.ReportSentBy", Me.sample.Barcode)
            Dim HasRole As Boolean = (MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT"))
            Dim hasImportRole As Boolean = (MedscreenLib.Glossary.Glossary.UserHasRole("DATA_IMPORT") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT"))
            'If HasRole Then
            Dim mnuItem As TaggedMenuItem
            For Each mnuItem In ctxMenu.MenuItems
                DrawMenuItem(SampleState, SendMethod, HasRole, hasImportRole, mnuItem)
                If mnuItem.MenuItems.Count > 0 And mnuItem.Enabled Then
                    For Each mItem As TaggedMenuItem In mnuItem.MenuItems
                        DrawMenuItem(SampleState, SendMethod, HasRole, hasImportRole, mItem)
                    Next
                End If
            Next
            'Else
            'Dim mnuItem As MenuItem
            'For Each mnuItem In ctxMenu.MenuItems
            '    If mnuitem.Mnemonic = "E" Then
            '        mnuItem.Visible = True
            '        mnuitem.Enabled = (sample.DateReported > DateField.ZeroDate AndAlso sample.Status <> "M" AndAlso sample.HasCertificate)
            '        'ElseIf mnuitem.Mnemonic = "M" Then
            '        '    mnuitem.Enabled = True
            '        '    mnuItem.Visible = True
            '    ElseIf mnuitem.Mnemonic = "O" Then
            '        mnuitem.Enabled = (sample.Status = "M" AndAlso sample.HasMRORequest)
            '        mnuItem.Visible = True
            '    ElseIf mnuitem.Mnemonic = "S" Then
            '        mnuitem.Enabled = (sample.HasScreenResult)
            '        mnuItem.Visible = True
            '    Else
            '        mnuItem.Visible = False
            '    End If
            'Next
            'End If
            

        End Sub


        Public Sub NotifyChange()
            Me.Refresh()
            RaiseEvent OnCustomerChange(Me.sample.Barcode)

        End Sub

        Public Sub raiseChange()
            RaiseEvent OnCustomerChange(Me.sample.Barcode)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Sample associated with node
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property sample() As Intranet.intranet.sample.OSample
            Get
                Return oSample
            End Get

        End Property


        Private Sub GetLabwareSample()
            If LabwareSampleCase Is Nothing Then
                LabwareSampleCase = New BAOLabware.SampleCase(sample.Barcode)
            End If
        End Sub
        Public Overrides Function ToHTML() As String

            Return sample.ToHTML
        End Function
    End Class
    ''' <summary>
    ''' A node to handle individual reports for a SMID
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    ''' <author></author>
    Public Class ReportSmidNode
        Inherits ReportNode
    End Class
 
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.ReportNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A scheduled report node
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Class ReportNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        Private MyRepSched As MedscreenLib.ReportSchedule

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new vanilla scheduled report node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.ImageIndex = 4

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new node with a report schedule
        ''' </summary>
        ''' <param name="rep">Report Schedule</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal rep As MedscreenLib.ReportSchedule, Optional ByVal AddSMID As Boolean = False)
            MyBase.new()
            Me.ImageIndex = 4
            MyRepSched = rep
            Dim cr As MedscreenLib.CRMenuItem
            cr = MedscreenLib.Glossary.Glossary.Menus.Item(MyRepSched.ReportID)

            rep.CrystalReport = cr
            If Not cr Is Nothing Then
                Me.Text = cr.MenuText
            Else
                Me.Text = MyRepSched.ReportID
            End If
            If AddSMID Then
                Dim objCust As Intranet.intranet.customerns.Client = Intranet.intranet.Support.cSupport.CustList(Me.MyRepSched.CustomerID)
                If Not objCust Is Nothing Then
                    Me.Text += " " & objCust.SMIDProfile
                Else
                    Me.Text += " " & MyRepSched.CustomerID
                End If

            End If
            Me.Text += " Next Report " & MyRepSched.NextReport.ToString("dd-MMM-yyyy") & " - " & _
                MyRepSched.Recipients
            If MyRepSched.RemoveFlag Then
                Me.BackColor = Color.Red
                Me.ForeColor = Color.Yellow
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh Node
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            'MyRepSched.CrystalReport.Fields.Load(CConnection.DbConnection)
            If Not MyRepSched.CrystalReport Is Nothing Then
                Me.Text = MyRepSched.CrystalReport.MenuText
            Else
                Me.Text = MyRepSched.ReportID
            End If
            Me.Text += " Next Report " & MyRepSched.NextReport.ToString("dd-MMM-yyyy") & " - " & _
                MyRepSched.Recipients

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Report Schedule belonging to Node
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ReportSchedule() As MedscreenLib.ReportSchedule
            Get
                Return Me.MyRepSched
            End Get
            Set(ByVal Value As MedscreenLib.ReportSchedule)
                MyRepSched = Value
            End Set
        End Property
    End Class


#End Region
End Namespace
