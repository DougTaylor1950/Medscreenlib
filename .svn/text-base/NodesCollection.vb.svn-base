Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports MedscreenLib.Glossary
Imports System.Windows.Forms
Imports System.Reflection
Imports Intranet.intranet
Imports Intranet.intranet.Support

Namespace Treenodes.CollectionNodes

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CollectionNodes.CentreCollectionHeaderNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specialist node for Centres 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CentreCollectionHeaderNode
        Inherits CollectionHeaderNode
        Public Sub New(ByVal myColl As Intranet.intranet.jobs.CMCollection)
            MyBase.New(myColl, "")
            Dim objCm As Intranet.intranet.jobs.CMJob
            For Each objCm In myColl
                Me.Nodes.Add(New CollectionNode(objCm))
            Next
        End Sub
    End Class

    Public Class ActiveJobHeader
        Inherits CollectionHeaderNode

        Public Sub New()
            MyBase.new("Active Jobs")
            Me.ImageIndex = 11
            Me.SelectedImageIndex = cstIconSelect1


        End Sub

        Public Shadows Function Refresh() As Boolean
            Try
                Me.Nodes.Clear()
            Catch
            End Try
            Try
                If TypeOf Me.Parent Is CustomerNodes.CustNode Then
                    Dim cuNode As CustomerNodes.CustNode = Me.Parent
                    If Not cuNode.client Is Nothing Then

                        Me.Collections.Clear()
                        Dim strQuery As String = "Select cm_job_number from job_header where job_status not in ('X','A') and customer_id = '" & _
                        cuNode.client.Identity & "'"

                        Dim ocmd As New OleDb.OleDbCommand(strQuery, CConnection.DbConnection)
                        Dim oColl As New Collection()
                        Dim oRead As OleDb.OleDbDataReader
                        If CConnection.ConnOpen Then
                            oRead = ocmd.ExecuteReader
                            While oRead.Read
                                oColl.Add(oRead.GetValue(0))
                            End While
                            oRead.Close()
                            CConnection.SetConnClosed()
                            Dim i As Integer
                            For i = 1 To oColl.Count
                                Dim objCm As New Intranet.intranet.jobs.CMJob(oColl.Item(i), False)
                                Me.Collections.Add(objCm)
                            Next
                        End If
                    End If
                End If
                Dim cmJob As Intranet.intranet.jobs.CMJob
                For Each cmJob In Me.Collections
                    Dim coNode As Treenodes.CollectionNodes.CollectionNode
                    coNode = New Treenodes.CollectionNodes.CollectionNode(cmJob)

                    If Not coNode Is Nothing Then Me.Nodes.Add(coNode)
                Next

            Catch
            End Try
        End Function
    End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CollectionHeaderNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection header node
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CollectionHeaderNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        Private myCollection As Intranet.intranet.jobs.CMCollection
        Private myCustomerId As String
        Friend WithEvents ctxMenu As System.Windows.Forms.ContextMenu

        Private myHeaderType As HeaderType = HeaderType.Profile
        Public Enum HeaderType
            Profile
            Smid
            Sage
            Vessel
        End Enum

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection header
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(Optional ByVal _type As HeaderType = HeaderType.Profile)
            MyBase.New("Collections")
            Me.ImageIndex = 32
            Me.SelectedImageIndex = cstIconSelect1
            myHeaderType = _type
            Try
                Dim MyAssembly As [Assembly]
                MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
                Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
                Dim myCollMenu As menus.CollectionHeaderMenu = menus.CollectionHeaderMenu.GetCollectionHeaderMenu
                Dim myMenu As New MedscreenLib.DynamicContextMenu(directory & "\CollHeadContextMenu.xml", myCollMenu)
                ctxMenu = myMenu.LoadDynamicMenu()
            Catch ex As Exception
            End Try

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Find Node and return it 
        ''' </summary>
        ''' <param name="CMJobNumber"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/03/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function CheckNodeForCollection(ByVal CMJobNumber As String) As CollectionNode
            Dim iNode As Treenodes.CollectionNodes.CollectionNode
            Dim retNode As CollectionNode = Nothing

            For Each iNode In Me.Nodes
                If iNode.Collection.ID = CMJobNumber Then
                    retNode = iNode

                    Exit For
                End If
            Next
            Return retNode
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection header with a colllection of Collections
        ''' </summary>
        ''' <param name="myColl">Collection of Collections</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal myColl As Intranet.intranet.jobs.CMCollection, ByVal CustomerId As String, Optional ByVal _type As HeaderType = HeaderType.Profile)
            MyBase.New("Collections")
            Me.ImageIndex = 32
            Me.SelectedImageIndex = cstIconSelect1
            myCollection = myColl
            myCustomerId = CustomerId
            myHeaderType = _type

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection header node with text
        ''' </summary>
        ''' <param name="text">Text to use</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal text As String, Optional ByVal _type As HeaderType = HeaderType.Profile)
            MyBase.new(text)
            Me.ImageIndex = 32
            Me.SelectedImageIndex = cstIconSelect1
            Me.myHeaderType = _type

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Collection of Collections
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Collections() As Intranet.intranet.jobs.CMCollection
            Get
                Return myCollection
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CMCollection)
                myCollection = Value
                'Me.Refresh()
            End Set
        End Property

        'expose context menu 
        Public ReadOnly Property ContextMenu() As ContextMenu
            Get
                Return Me.ctxMenu
            End Get
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Customer ID for collection of collections 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CustomerID() As String
            Get
                Return Me.myCustomerId
            End Get
            Set(ByVal Value As String)
                myCustomerId = Value
                If Me.myCollection Is Nothing AndAlso Me.myHeaderType = HeaderType.Smid Then
                    myCollection = New Intranet.intranet.jobs.CMCollection(myCustomerId, jobs.CMCollection.collectionType.Smid, "")
                ElseIf Me.myCollection Is Nothing AndAlso Me.myHeaderType = HeaderType.Profile Then
                    myCollection = New Intranet.intranet.jobs.CMCollection(myCustomerId, jobs.CMCollection.collectionType.Profile, "")
                End If
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the contents of the node
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            If Me.TreeView Is Nothing Then Exit Function
            If Me.Parent Is Nothing Then Exit Function
            Try
                Me.TreeView.BeginUpdate()
                Me.Nodes.Clear()
            Catch
            Finally
                Me.TreeView.EndUpdate()
            End Try
            Try
                If TypeOf Me.Parent Is CustomerNodes.CustNode And TypeOf Me Is CollectionHeaderNode Then
                    Dim cuNode As CustomerNodes.CustNode = Me.Parent

                    If Not cuNode.client Is Nothing Then
                        If cuNode.client.Collections.Count = 0 OrElse Now.Subtract(cuNode.client.Collections.TimeStamp).TotalMinutes > 20 Then cuNode.client.Collections.Load("Lib_collection.GetDiaryRecords", "Select c.*,rowid from collection c where customer_id = '" & _
                           Me.CustomerID & "' and coll_date > '" & Today.AddMonths(-2).ToString("dd-MMM-yy") & "'")
                        Me.Collections = cuNode.client.Collections

                    End If
                ElseIf TypeOf Me.Parent Is VesselNodes.VesselNode And TypeOf Me Is CollectionHeaderNode Then
                    Dim cuNode As VesselNodes.VesselNode = Me.Parent
                    If Not cuNode.Vessel Is Nothing Then
                        cuNode.Vessel.Collections.RefreshChanged()
                        Me.Collections = cuNode.Vessel.Collections
                    End If
                    'deal with being of a SMID node
                ElseIf TypeOf Me.Parent Is Treenodes.CustomerNodes.CustSMIDNode And TypeOf Me Is CollectionHeaderNode Then
                    Dim cuNode As CustomerNodes.CustSMIDNode = Me.Parent
                    Dim ocoll = New Collection()
                    ocoll.add(CConnection.StringParameter("StatusX", cuNode.SMID.SMID, 10))
                    If CType(Me.TreeView, CustomTreeBase).AlphaNodes.CurrentViewState = AlphaNodeList.CustomerViewState.ViewBySage Then
                        Me.Collections = New jobs.CMCollection("customer_id in (select identity from customer where sage_name = ? and removeflag = 'F') and coll_date > sysdate -60 and status not in ('S','T')", ocoll)
                    Else
                        Me.Collections = New jobs.CMCollection("customer_id in (select identity from customer where smid = ? and removeflag = 'F') and coll_date > sysdate -60 and status not in ('S','T')", ocoll)

                    End If
                End If
                Dim cmJob As Intranet.intranet.jobs.CMJob
                Me.TreeView.BeginUpdate()
                For Each cmJob In Me.Collections
                    Dim coNode As Treenodes.CollectionNodes.CollectionNode
                    coNode = New Treenodes.CollectionNodes.CollectionNode(cmJob)

                    If Not coNode Is Nothing Then Me.Nodes.Add(coNode)
                Next

            Catch ex As Exception
            Finally
                Me.TreeView.EndUpdate()
            End Try
        End Function
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CollectionNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A node with a collection attached
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CollectionNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        Private myColl As Intranet.intranet.jobs.CMJob
        Private blnSamplesDisplayed As Boolean = False
        Private objSampleHeadernode As Treenodes.SampleHeaderNode
        Friend WithEvents ctxMenu As System.Windows.Forms.ContextMenu
        Friend WithEvents mnuResend As System.Windows.Forms.MenuItem
        Friend WithEvents mnuCollToSelf As System.Windows.Forms.MenuItem
        Friend WithEvents mnuSetCollDate As System.Windows.Forms.MenuItem
        Private blnOld As Boolean = False
        Private myStatusPhrase As Phrase

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection node
        ''' </summary>
        ''' <param name="objColl">Colelction to pass</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal objColl As Intranet.intranet.jobs.CMJob)
            MyBase.New()

            myColl = objColl
            If objColl Is Nothing Then Exit Sub
            ReDraw()
            blnOld = True
            Try
                Dim MyAssembly As [Assembly]
                MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
                Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
                Dim myCollMenu As menus.CollectionMenu = menus.CollectionMenu.GetCollectionMenu
                Dim myMenu As New MedscreenLib.DynamicContextMenu(directory & "\CollContextMenu.xml", myCollMenu)
                ctxMenu = myMenu.LoadDynamicMenu()
            Catch ex As Exception
            End Try

            Me.SelectedImageIndex = cstIconSelect1

        End Sub

        Public Sub CMNodeChanged()
            MyBase.OnNodeChanged()
        End Sub


        Public Function ReDraw() As Boolean
            Select Case myColl.Status
                Case "M"
                    Me.ImageIndex = 41
                Case "R"
                    Me.ImageIndex = 40
                Case "D"
                    Me.ImageIndex = 39
                Case "C"
                    Me.ImageIndex = 38
                Case "P"
                    Me.ImageIndex = 37
                Case "W"
                    Me.ImageIndex = 36
                Case Else
                    Me.ImageIndex = 35
            End Select
            If myColl.ID.Trim.Length > 0 Then Me.Text = Me.ToolTipText
            If myStatusPhrase Is Nothing Then
                myStatusPhrase = New Phrase("COLL_STAT", myColl.Status)
                MyBase.DecorateNode(myStatusPhrase)
            End If

            If Not blnOld Then MyBase.OnNodeChanged()
        End Function

        'expose context menu 
        Public ReadOnly Property ContextMenu() As ContextMenu
            Get
                If Not ctxMenu Is Nothing Then ContextMenuRedraw()
                Return Me.ctxMenu
            End Get
        End Property

        Private Sub SetMenuContextState(ByVal sampleState As String, ByVal Menuitems As Menu.MenuItemCollection)
            Dim mnuItem As TaggedMenuItem
            For Each mnuItem In Menuitems
                mnuItem.Visible = True
                If mnuItem.Group = "UNCOMMIT" Then
                    mnuItem.Enabled = False
                    If mnuItem.Mnemonic = "A" And (InStr(sampleState, "A") + InStr(sampleState, "C") + InStr(sampleState, "R")) > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "R" And (InStr(sampleState, "R") + InStr(sampleState, "C")) > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "C" And InStr(sampleState, "C") > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "H" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "T" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "I" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "P" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "D" Then mnuItem.Enabled = True

                End If
                If Not mnuItem.MenuItems Is Nothing Then
                    SetMenuContextState(sampleState, mnuItem.MenuItems)
                End If
            Next


        End Sub
        
        ''' <developer>CONCATENO\Taylor</developer>
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><modified>20-Dec-2011 13:24 Check for roles added</modified><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Public Sub ContextMenuRedraw()
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") OrElse MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                Dim SampleState As String = ""
                If Me.Collection.JobName.Trim.Length > 0 Then
                    SampleState = CConnection.PackageStringList("lib_Uncommit.GetJobStatus", Me.Collection.JobName)
                    SetMenuContextState(SampleState, ctxMenu.MenuItems)
                End If
            Else
                Dim mnuItem As MenuItem
                For Each mnuItem In ctxMenu.MenuItems
                    If mnuItem.Text.ToUpper = "JOBS" Then mnuItem.Visible = False
                Next
            End If
            DynamicContextMenu.SetMenuRoleState(ctxMenu.MenuItems)
            

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
        Private Sub mnuSetCollDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles mnuSetCollDate.Click
            If Me.myColl Is Nothing Then Exit Sub
            Dim objDate As Object = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Collection Date", "Collection Date", Me.myColl.CollectionDate)
            If Not objDate Is Nothing Then
                Me.myColl.CollectionDate = CDate(objDate)
                Me.myColl.Update()
            End If
        End Sub

        Private Sub mnuCompleteCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            If Me.Collection Is Nothing Then Exit Sub
            Try
                Me.Collection.CompleteJob()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub
#End Region

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
        Private Sub mnuCollToSelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles mnuCollToSelf.Click
            If Me.myColl Is Nothing Then Exit Sub

            Dim strDestination As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
            Dim Destinations As String() = strDestination.Split(New Char() {"|"})
            Dim objTextDoc As Reporter.clsTextFile = Me.myColl.CollectorOutFile(Destinations, "SEND")
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

        Private Sub mnuCOComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            If Me.Collection Is Nothing Then Exit Sub
            Try
                If Me.Collection.Status = Constants.GCST_JobStatusPay Then
                    MsgBox("This action is not relevant to Pay collections")
                    'nwInfo.Notify("This action is not relevant to Pay collections")
                End If
                If Me.Collection.CollOfficer.Trim.Length = 0 Then
                    'nwInfo.Notify("This action is not relevant to collections to which an officer has not been assigned")
                    MsgBox("This action is not relevant to collections to which an officer has not been assigned")
                End If


                Dim afrm As New FrmCollComment()
                afrm.CommentType = "PROB"
                afrm.FormType = FrmCollComment.FCCFormTypes.CollectorInfo
                afrm.ProgamManager = Collection.SmClient.ProgrammeManager
                afrm.Port = Collection.CollectionPort
                afrm.CollOfficer = Collection.CollOfficer
                afrm.CMNumber = Me.Collection.ID

                If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                        Me.Collection.Comments.CreateComment(afrm.Comment, afrm.CommentType)
                    aComm.ActionBy = afrm.CollOfficer
                    aComm.Update()
                    Collection.CancellationStatus = Constants.GCST_CANC_Problem
                    Collection.Update()
                    Collection.Refresh(False)
                    'TODO: add code to produce an email / fax the document

                    Dim myColl As Intranet.intranet.jobs.CMJob = Collection
                    Dim myOff As Intranet.intranet.jobs.CollectingOfficer = myColl.CollectingOfficer

                    If afrm.ckCopy.Checked Then
                        If Medscreen.IsNumber(afrm.txtCopyDest.Text) Then
                            Medscreen.QuckEmail("Issue with Collection CM" & Collection.ID, _
                            afrm.Comment, afrm.txtCopyDest.Text & Constants.GCST_EMAIL_To_Fax_Send, cSupport.User.Email)
                        Else
                            Medscreen.QuckEmail("Issue with Collection CM" & Collection.ID, _
                            afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)
                        End If

                    End If
                    Collection.LogAction(jobs.CMJob.Actions.RaiseIssue, afrm.txtCopyDest.Text)

                    Dim LitOutput As String
                    Dim strVessel As String = myColl.VesselName
                    'If Not myColl.VesselObject Is Nothing Then
                    '    strVessel = myColl.VesselObject.VesselName
                    'ElseIf Not myColl.SiteObject Is Nothing Then
                    '    strVessel = myColl.SiteObject.VesselName
                    'Else
                    '    strVessel = myColl.vessel
                    'End If
                    'Dummy declares to satisfy function 
                    Dim strBooker As String = ""
                    Dim strUK As String = ""
                    LitOutput = myColl.GetHeader(strVessel, strBooker, strUK, "INFO")

                    Dim strFilename As String
                    'If Intranet.intranet.Support.cSupport.DatabaseInUse = MedscreenLib.CConnection.useDatabase.LIVE Then
                    strFilename = Medscreen.GetNextFileName(MedscreenLib.MedConnection.Instance.ServerPath & "\WordOutput\Collection_Info-" & Mid(Collection.ID, 6, 5), "xxxx")
                    'Else'
                    'strFilename = Medscreen.GetNextFileName("\\fs01\common\xmltemp\WordOutput\Collection_Info-" & Mid(meCurrent.Collection.ID, 6, 5), "xxxx")
                    'End If
                    Dim iof As New IO.StreamWriter(strFilename)

                    iof.WriteLine("Output_type=EMAIL")
                    If cSupport.User Is Nothing Then
                        iof.WriteLine("Output_address=" & afrm.SendTo)
                    ElseIf cSupport.User.Email.Trim.Length = 0 Then
                        iof.WriteLine("Output_address=" & afrm.SendTo)
                    ElseIf afrm.SendTo.ToUpper <> cSupport.User.Email.ToUpper Then
                        iof.WriteLine("Output_address=" & afrm.SendTo & ";" & cSupport.User.Email)
                    Else
                        iof.WriteLine("Output_address=" & afrm.SendTo)
                    End If
                    'Set the from address in the email
                    iof.WriteLine(MedscreenLib.Constants.ReporterEmailFrom)
                    Dim RepTemp As String = MedscreenCommonGUIConfig.Templates("CollOffInfo")
                    iof.WriteLine("Report_template=" & RepTemp) 'CollOffInfo")
                    iof.WriteLine("txtCollOfficer=" & myOff.Name)
                    iof.WriteLine("txtCustomerID=" & myColl.SmClient.ClientID)
                    iof.WriteLine("txtCMNumber=" & myColl.ID)
                    iof.WriteLine("txtPort=" & myColl.Port)
                    iof.WriteLine("txtVesselName=" & strVessel)
                    iof.WriteLine("txtCollDate=" & myColl.CollectionDate.ToString("dd-MMM-yyyy"))
                    iof.WriteLine("txtSender=" & MedscreenLib.Glossary.Glossary.CurrentSMUser.Description)
                    iof.WriteLine("txtSubject=" & LitOutput)
                    iof.WriteLine("[FreeBodyText]=txtComments")

                    iof.WriteLine(afrm.Comment)
                    iof.WriteLine("[End]")
                    'DT added 4th March 2010 return address for comments
                    'iof.WriteLine("CollAdmin=" & MedscreenLib.Constants.ReporterEmailFrom)
                    iof.WriteLine("Subject=" & LitOutput)
                    iof.WriteLine("SEND_LOG=TRUE")
                    iof.WriteLine("SEND_LOG_REF=Officer Comment :" & myColl.ID)

                    iof.Flush()
                    iof.Close()

                    Dim strNewName As String = Mid(strFilename, 1, strFilename.Length - 4) & "out"
                    IO.File.Move(strFilename, strNewName)
                    'Refresh()

                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "MnuCOComment")
            End Try
        End Sub


        Private Sub mnuDeferCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            If Me.Collection Is Nothing Then Exit Sub
            With Me.Collection
                Dim newDate As Date = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                    "New Collection Date", "Defer Collection " & .CMID & " - " & .ClientId & _
                    " " & .VesselName, .CollectionDate)
                If newDate.Day * newDate.Month * newDate.Year <> 1 Then
                    If MsgBox("Change collection date to " & newDate.ToString("dd-MMM-yyyy"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        .CollectionDate = newDate
                        If .Status = Constants.GCST_JobStatusConfirmed Then
                            If Not Me.Collection.Job Is Nothing Then
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
        ''' Are the sample nodes displayed
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property SamplesDisplayed() As Boolean
            Get
                Return Me.blnSamplesDisplayed
            End Get
            Set(ByVal Value As Boolean)
                blnSamplesDisplayed = Value
            End Set

        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert to HTML
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	22/03/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Overrides Function ToHTML() As String
            If Not Me.Collection Is Nothing Then
                If Collection.IsLabware Then
                    Return Me.Collection.ToHTML("test302Labware.xsl", , , False)
                Else
                    Return Me.Collection.ToHTML(, , , False)
                End If

            Else
                Return ""
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Make sure samples are displayed
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub DisplaySamples()
            If Me.Collection Is Nothing Then Exit Sub ' No collectio bog off
            'If Not (Me.Collection.Status = "M" Or Me.Collection.Status = "R") Then Exit Sub
            If Me.SamplesDisplayed Then Exit Sub 'Already displayed bog off

            objSampleHeadernode = New Treenodes.SampleHeaderNode(Me.Collection.JobName, Me.Collection.TableSet)     'Create a header node
            Me.Nodes.Add(objSampleHeadernode)                           'Add that node
            'Dim objSamp As Intranet.intranet.sample.OSample             'Declare a sample variable

            'For Each objSamp In Me.Collection.CollectionSamples         'For each sample create a node
            '    Dim oSamNode As New Treenodes.SampleNode(objSamp)
            '    objSampleHeadernode.Nodes.Add(oSamNode)
            'Next
            SamplesDisplayed = True
            objSampleHeadernode.Refresh()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Collection for this node
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Collection() As Intranet.intranet.jobs.CMJob
            Get
                Return myColl
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CMJob)
                myColl = Value
                If Not Value Is Nothing Then Me.Text = ToolTipText()
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tooltip that will be displayed if tooltips are on
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToolTipText() As String
            Dim strText As String
            Dim strPort As String = CConnection.PackageStringList("Lib_collection.GetPort", myColl.Port)
            strText = Me.myColl.CMID & " " & Me.myColl.Status & " - " & CStr(myColl.NoOfDonors) & " - "
            If Not strPort Is Nothing Then
                strText += strPort
            End If
            strText += " - " & Me.myColl.VesselName
            strText += " - " & myColl.CollectionDate.ToString("dd-MMM-yyyy HH:mm")
            Return strText
        End Function

        Public Function PositionBarcode(ByVal Barcode As String) As SampleNode
            DisplaySamples()
            If Not Me.objSampleHeadernode Is Nothing Then
                Return Me.objSampleHeadernode.PositionBarcode(Barcode)
            End If

        End Function
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Treenodes.CollectionNodes.JobNode
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A job header node but internally represented as a collection
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class JobNode
        Inherits MedscreenCommonGui.Treenodes.MedTreeNode
        'Private objJob As CMCollection
        Private objcJob As Intranet.intranet.jobs.CMJob
        Private intJobs As Integer
        Private blnFirst As Boolean = True
        Private tP As ToolTip
        'Private f As Form

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new job node from a collection object
        ''' </summary>
        ''' <param name="job">Collection</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal job As Intranet.intranet.jobs.CMJob)
            MyBase.New()
            objcJob = job
            Me.ImageIndex = 11
            'f = fIn
            DoRefresh()

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tooltip for node
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToolTipText() As String
            Dim strText As String
            strText = objcJob.JobName & " - status " & objcJob.Status & _
    " Samples = " & objcJob.CollectionSamples.Count & _
    " (" & objcJob.CollectionSamples.UntestedSamples & ") "
            If objcJob.IsVessel Then
                If Not objcJob.VesselObject Is Nothing Then
                    strText += "Vessel " & objcJob.VesselObject.VesselName & " - Port " & objcJob.Port & vbCrLf
                Else
                    strText += "Vessel " & objcJob.vessel & " - Port " & objcJob.Port & vbCrLf
                End If
            Else
                strText += "Client " & objcJob.vessel & " - Site " & objcJob.Port & vbCrLf
            End If

            strText += " - CM" & objcJob.ID & " - invoice " & objcJob.Invoiceid & vbCrLf & _
                    " Collection Officer - " & objcJob.CollOfficer & _
                    " Arranged by - " & objcJob.Booker & vbCrLf
            If objcJob.Announced Then
                strText += "Announced "
            Else
                strText += "Unannounced"
            End If
            If objcJob.UK Then
                strText += " UK Collection "
            Else
                strText += " Overseas Collection"
            End If
            strText += " Collection Type - " & objcJob.CollectionType

            'End If
            Return strText

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' refresh node
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub DoRefresh()
            'tP = f.ToolTip1
            If blnFirst Then
                If Not objcJob.CollectionSamples Is Nothing Then
                    If objcJob.JobName.Trim.Length > 0 Then
                        Text = objcJob.JobName & " - status " & objcJob.Status & _
                            " Samples = " & objcJob.CollectionSamples.Count & _
                            " (" & objcJob.CollectionSamples.UntestedSamples & ") " & _
                            objcJob.vessel & " - " & objcJob.Port
                    Else
                        Text = objcJob.Invoiceid & " - status " & objcJob.Status & _
                                                " Samples = " & objcJob.CollectionSamples.Count & _
                                                " (" & objcJob.CollectionSamples.UntestedSamples & ") " & _
                                                objcJob.vessel & " - " & objcJob.Port
                    End If
                ElseIf objcJob.JobName.Trim.Length > 0 Then
                    Text = objcJob.JobName & " - status " & objcJob.Status
                Else
                    Text = objcJob.Invoiceid & " - status " & objcJob.Status
                End If
                blnFirst = False
            End If
            Select Case objcJob.Status
                Case "C"
                    Me.ForeColor = System.Drawing.Color.LightBlue
                Case "V"
                    Me.ForeColor = System.Drawing.Color.OrangeRed
                Case "W"
                    Me.ForeColor = System.Drawing.Color.DarkOrange
                Case "P"
                    Me.ForeColor = System.Drawing.Color.Green
                Case "A"
                    Me.ForeColor = System.Drawing.Color.DarkGreen
                Case Else
                    Me.ForeColor = System.Drawing.SystemColors.ControlText
            End Select

            'End If

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh node
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Dim objNode As TreeNode
            Dim sNode As Treenodes.SampleNode

            Me.objcJob.Refresh()
            objNode = Nodes.Item(0)
            For Each sNode In objNode.Nodes
                sNode.Refresh()
            Next
            DoRefresh()
        End Function



        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose collection for node
        ''' </summary>
        ''' <param name="blnJ"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads ReadOnly Property Job(ByVal blnJ As Boolean) As Intranet.intranet.jobs.CMJob
            Get
                Return objcJob
            End Get

        End Property
    End Class


End Namespace
