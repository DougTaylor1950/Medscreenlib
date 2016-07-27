Imports MedscreenLib
Imports Intranet.intranet.Support
Imports Intranet.intranet
Imports Reporter
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Namespace menus

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : menus.MenuBase
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection menus
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public MustInherit Class MenuBase

        Friend Shared Function GetSenderType(ByVal sender As Object) As Object
            'see if we have a collection passed 
            If TypeOf sender Is Intranet.intranet.jobs.CMJob Then
                Return sender
                Exit Function
            End If
            Dim ctxMenu As ContextMenu
            Dim MM As MainMenu
            If TypeOf sender Is ContextMenu Then
                ctxMenu = sender
                Dim a As MenuItem
            ElseIf TypeOf sender Is MenuItem Then
                'its parent should be a contextmenu

                While TypeOf sender.parent Is MenuItem
                    sender = sender.parent
                End While
                If TypeOf sender.parent Is ContextMenu Then
                    ctxMenu = sender.parent
                ElseIf TypeOf sender.parent Is MainMenu Then
                    MM = sender.parent
                    Dim ACont As System.ComponentModel.Container = MM.Container
                End If
            Else
                Return Nothing
                Exit Function
            End If
            If ctxMenu.SourceControl Is Nothing Then
                Exit Function
            End If
            Dim tv As TreeView
            Dim dr As frmDiary
            If TypeOf ctxMenu.SourceControl Is TreeView Then
                tv = ctxMenu.SourceControl
                Return tv.SelectedNode
            ElseIf TypeOf ctxMenu.SourceControl Is frmDiary Then
                dr = ctxMenu.SourceControl
                Return dr.lvItemCurrent
            End If
        End Function

        Friend Shared Sub UpdateSender(ByVal Sender As Object)
            If TypeOf Sender Is ListViewItems.ListViewDiaryItem Then
                CType(Sender, ListViewItems.ListViewDiaryItem).Refresh()
            ElseIf TypeOf Sender Is Treenodes.CollectionNodes.CollectionNode Then
                CType(Sender, Treenodes.CollectionNodes.CollectionNode).ReDraw()
                CType(Sender, Treenodes.CollectionNodes.CollectionNode).CMNodeChanged()
            ElseIf TypeOf Sender Is Treenodes.CollectionNodes.CollectionHeaderNode Then
                CType(Sender, Treenodes.CollectionNodes.CollectionHeaderNode).Refresh()
            ElseIf TypeOf Sender Is Treenodes.CustomerNodes.CustNode Then
                CType(Sender, Treenodes.CustomerNodes.CustNode).Refresh()
            End If
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : menus.CollectionMenu
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Menus for collection items
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	15/10/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class CollectionMenu
        Inherits MenuBase
        Private Shared myInstance As CollectionMenu
        Private Shared blnCanCancel As Boolean = False
        Private Sub New()

        End Sub

        Public Shared Function GetCollectionMenu() As CollectionMenu
            If myInstance Is Nothing Then
                myInstance = New CollectionMenu()
                blnCanCancel = Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MED_COLL_CANCEL")

            End If
            Return myInstance
        End Function


        Private Shared Function GetCollection(ByVal senderType As Object) As jobs.CMJob
            Dim cNode As Treenodes.CollectionNodes.CollectionNode
            Dim anItem As ListViewItems.ListViewDiaryItem
            Dim aCMItem As ListViewItems.ListViewCMDiaryItem
            Dim myColl As Intranet.intranet.jobs.CMJob
            If TypeOf senderType Is Treenodes.CollectionNodes.CollectionNode Then
                cNode = senderType
                myColl = cNode.Collection
            ElseIf TypeOf senderType Is Intranet.intranet.jobs.CMJob Then
                myColl = senderType
            ElseIf TypeOf senderType Is ListViewItems.ListViewCMDiaryItem Then
                aCMItem = senderType
                myColl = aCMItem.Collection
            ElseIf TypeOf senderType Is ListViewItems.ListViewDiaryItem Then
                anItem = senderType
                myColl = anItem.Collection
            ElseIf TypeOf senderType Is Treenodes.VesselNodes.VesselNode Then

            Else
                myColl = Nothing
            End If
            Return myColl
        End Function

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
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            'bog off if we have no collection
            If myColl Is Nothing Then Exit Sub

            'Ask the user for the date
            Dim objDate As Object = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Collection Date", "Collection Date", myColl.CollectionDate)
            If Not objDate Is Nothing Then
                'If we have been given a date update it (and any job date as well).
                myColl.CollectionDate = CDate(objDate)
                myColl.Update()
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuCompleteCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Try
                myColl.CompleteJob()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            'Now get a comment 
            Dim objString As Object = MedscreenLib.Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", "Comment on completion")
            If Not objString Is Nothing Then
                myColl.Comments.CreateComment(objString.ToString, "INFO", jobs.CollectionCommentList.CommentType.Collection)
            End If
            MyBase.UpdateSender(SenderType)

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
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Dim strDestination As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
            Dim Destinations As String() = strDestination.Split(New Char() {"|"})
            Dim objTextDoc As Reporter.clsTextFile = myColl.CollectorOutFile(Destinations, "SEND")
            Medscreen.LogAction("Creating reporter")
            Dim UKorOS As String = "U"
            If myColl.CollectionPort.CountryId <> 44 Then UKorOS = "W"
            Dim strFilename As String = Medscreen.GetNextFileName(MedscreenLib.MedConnection.Instance.ServerPath & "\WordOutput", "Collection_Request-" & Mid(myColl.ID, 5, 6) & "-" & UKorOS, "out")
            objTextDoc.Save(strFilename)

            'Dim objReporter As Reporter.Report = New Reporter.Report()
            'Medscreen.LogAction("scanning templates")

            'Dim strPath As String = "\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates|\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates\Scanned Documents"


            'Dim PathArray As String() = strPath.Split(New Char() {"|"})
            'objReporter.ReportSettings.SetTemplatePaths(PathArray)
            'Dim objCReport As Reporter.clsWordReport = objReporter.GetDatafile(objTextDoc.ExportStringArray)
            'Dim myRecipient As Reporter.clsRecipient = New Reporter.clsRecipient()
            'myRecipient.address = strDestination
            'myRecipient.method = "EMAIL"
            'Medscreen.LogAction("about to send")

            'objCReport.SendByEmail(myRecipient)
        End Sub

        Private Sub mnuCollInstrView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles mnuCollToSelf.Click
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Dim strXML As String = myColl.ToXML(Intranet.intranet.Support.cSupport.XMLCollectionOptions.Defaults)
            Dim strHTML As String
            If Not myColl.SiteObject Is Nothing Then
                strHTML = Medscreen.ResolveStyleSheet(strXML, "CollSiteRequest.xsl", 0)
            Else
                strHTML = Medscreen.ResolveStyleSheet(strXML, "CollRequest.xsl", 0)
            End If
            Medscreen.ShowHtml(strHTML)


        End Sub

        Private Sub mnuCOComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Try
                If myColl.Status = Constants.GCST_JobStatusPay Then
                    MsgBox("This action is not relevant to Pay collections")
                    Exit Sub
                    'nwInfo.Notify("This action is not relevant to Pay collections")
                End If
                If myColl.CollOfficer.Trim.Length = 0 Then
                    'nwInfo.Notify("This action is not relevant to collections to which an officer has not been assigned")
                    MsgBox("This action is not relevant to collections to which an officer has not been assigned")
                    Exit Sub
                End If


                Dim afrm As New FrmCollComment()
                afrm.CommentType = "PROB"
                afrm.FormType = FrmCollComment.FCCFormTypes.CollectorInfo
                afrm.ProgamManager = myColl.SmClient.ProgrammeManager
                afrm.Port = myColl.CollectionPort
                afrm.CollOfficer = myColl.CollOfficer
                afrm.CMNumber = myColl.ID

                If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                    myColl.Comments.CreateComment(afrm.Comment, afrm.CommentType)
                    aComm.ActionBy = afrm.CollOfficer
                    aComm.Update()
                    myColl.CancellationStatus = Constants.GCST_CANC_Problem
                    myColl.Update()
                    myColl.Refresh(False)
                    'TODO: add code to produce an email / fax the document

                    'Dim myColl As Intranet.intranet.jobs.CMJob = myColl
                    'Dim myOff As Intranet.intranet.jobs.CollectingOfficer = myColl.CollectingOfficer

                    If afrm.ckCopy.Checked Then
                        If Medscreen.IsNumber(afrm.txtCopyDest.Text) Then
                            Medscreen.QuckEmail("Issue with Collection CM" & myColl.ID, _
                     afrm.Comment, afrm.txtCopyDest.Text & Constants.GCST_EMAIL_To_Fax_Send, cSupport.User.Email)
                        Else
                            Medscreen.QuckEmail("Issue with Collection CM" & myColl.ID, _
                    afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)
                        End If

                    End If
                    '
                    'cNode.Refresh()
                    If afrm.SendTo.ToUpper = "@STARFAX.CO.UK" Then
                        MsgBox("No recipient specified : problem not sent ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    Else
                        myColl.SendOfficerProblem(afrm.txtCopyDest.Text, afrm.Comment, afrm.SendTo)
                    End If

                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "MnuCOComment")
            End Try
            MyBase.UpdateSender(SenderType)
        End Sub


        Private Sub mnuDeferCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            With myColl
                Dim newDate As Date = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                    "New Collection Date", "Defer Collection " & .CMID & " - " & .ClientId & _
                    " " & .VesselName, .CollectionDate)
                If newDate.Day * newDate.Month * newDate.Year <> 1 Then
                    If MsgBox("Change collection date to " & newDate.ToString("dd-MMM-yyyy"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        .CollectionDate = newDate
                        If .Status = Constants.GCST_JobStatusConfirmed Then
                            If Not .Job Is Nothing Then
                                .Job.CollectionDate = newDate
                                .Job.Update()
                            End If
                        End If
                        .Comments.CreateComment("Collection Deffered to " & newDate.ToString("dd-MMM-yyyy"), Constants.GCST_COMM_CCFix)
                        .Update()
                        .Refresh()
                        If .Status = Constants.GCST_JobStatusSent Or .Status = Constants.GCST_JobStatusConfirmed Then
                            If MessageBox.Show("Send a note to the collecting officer to tell them the collection has been delayed?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                                Me.mnuCOComment_Click(sender, e)
                            End If

                        End If
                    End If
                End If
            End With
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuChangeVessel_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            UpdateSender(SenderType)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Try
                Dim vessForm As New frmVesselSelect()
                vessForm.FormMode = frmVesselSelect.Modes.WithinSMID
                vessForm.SMID = myColl.SmClient.SMID
                If vessForm.ShowDialog = DialogResult.OK Then
                    myColl.vessel = vessForm.Vessel.VesselID
                    If Not myColl.Job Is Nothing Then
                        myColl.Job.VesselName = vessForm.Vessel.VesselID
                        myColl.Job.Update()
                    End If
                    Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                        myColl.Comments.CreateComment("Vessel or site for collection has been changed", "FIX")
                    If Not aComm Is Nothing Then
                        aComm.ActionBy = ""
                        aComm.Update()
                        myColl.CancellationStatus = Constants.GCST_CANC_Problem
                        myColl.Update()
                        myColl.Refresh(False)
                        If myColl.Status = "P" Or myColl.Status = "C" Then
                            'Dim UP As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item(afrm.ActionBy, _
                            '    MedscreenLib.personnel.PersonnelList.FindIdBy.identity)
                            Dim strSubject As String = "Issue with Collection CM" & myColl.ID
                            If Not myColl.VesselObject Is Nothing Then
                                strSubject += " vessel " & myColl.VesselName
                            ElseIf Not myColl.SiteObject Is Nothing Then
                                strSubject += " site " & myColl.VesselName
                            Else
                                strSubject += myColl.vessel
                            End If

                            strSubject += " Port " & myColl.Port
                            strSubject += " Officer " & myColl.CollOfficer
                            If Not myColl.CollectingOfficer Is Nothing AndAlso myColl.CollectingOfficer.SendDestination.Trim.Length > 0 Then
                                Medscreen.BlatEmail(strSubject, _
                                aComm.CommentText, myColl.CollectingOfficer.SendDestination, cSupport.User.Email)
                            End If
                        End If
                    End If
                    myColl.Update()
                End If
            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuAlterDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim cNode As Treenodes.CollectionNodes.CollectionNode
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Try
                'Run change destination code
                ChangeDestination(myColl)
            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

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
        Private Sub ChangeDestination(ByVal myColl As jobs.CMJob)
            Dim dlgRet As Microsoft.VisualBasic.MsgBoxResult
            'See if it is a banned status
           

            'See about Confirmed collections 
            'If myColl.Status = MedscreenLib.Constants.GCST_JobStatusConfirmed Then
            '    dlgRet = MsgBox("This collection has been confirmed, do you want to unconfirm it then alter the destination?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
            '    If dlgRet = MsgBoxResult.No Then Exit Sub

            '    Dim e As New Intranet.intranet.jobs.JobEvent(myColl.ID)

            '    mnuUnconfirmCollection_Click(myColl, e)

            'End If

            Dim afrm As New MedscreenCommonGui.frmSelectPort()
            Dim oPort As Intranet.intranet.jobs.Port
            If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                oPort = afrm.Port
                If Not oPort Is Nothing Then
                    myColl.Port = oPort.Identity
                    myColl.UK = oPort.IsUK
                    myColl.Update()
                    If InStr("CDRM", myColl.Status) > 0 Then           'We will have a job in existence
                        'Need to check the job name and update it if necessary
                        myColl.Job.Location = oPort.Identity
                        myColl.Job.Update()
                        If oPort.IsUK Then
                            If myColl.Job.JobName.Chars(0) = "W" Then                   'Change it to Overseas
                                ChangeJobName(myColl.JobName, "U" & Mid(myColl.Job.JobName, 2))
                                myColl.Job.JobName = "U" & Mid(myColl.Job.JobName, 2)
                                myColl.JobName = myColl.Job.JobName
                            ElseIf myColl.Job.JobName.Chars(0) = "U" Then               'Change it to UK
                                ChangeJobName(myColl.JobName, "W" & Mid(myColl.Job.JobName, 2))
                                myColl.Job.JobName = "W" & Mid(myColl.Job.JobName, 2)
                                myColl.JobName = myColl.Job.JobName

                            End If
                        End If
                    End If
                    myColl.Status = Constants.GCST_JobStatusCreated

                    myColl.Update()
                End If

            End If
            ' MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub ChangeJobName(ByVal oldName As String, ByVal newNAme As String)
            Dim ocmd As New OleDb.OleDbCommand("lib_sample8.ChangeJobName", CConnection.DbConnection)
            Try
                ocmd.CommandType = CommandType.StoredProcedure
                ocmd.Parameters.Add(CConnection.StringParameter("oldName", oldName, 30))
                ocmd.Parameters.Add(CConnection.StringParameter("newName", newNAme, 30))
                If CConnection.ConnOpen Then
                    Dim intret As Integer = ocmd.ExecuteNonQuery
                End If
            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()
            End Try
        End Sub

        Private Sub mnuUnconfirmCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            If myColl.Status = Constants.GCST_JobStatusReceived Or _
            myColl.Status = Constants.GCST_JobStatusCollected Or _
            myColl.Status = Constants.GCST_JobStatusCommitted Then
                MsgBox("This collection can not be unconfirmed now", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            If Not myColl.Status = Constants.GCST_JobStatusConfirmed Then
                MsgBox("This collection is not confirmed", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            Try
                If MsgBox("This option will change the status of a collection from confirmed to available, " & vbCrLf & _
                "you will then be offered the chance of changing the details of the collection." & _
                vbCrLf & "Do you want to proceed (CM" & myColl.ID & ") ?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
                Try
                    myColl.CollRedirectReport()
                    myColl.Comments.CreateComment("Collection Unconfirmed", Constants.GCST_COMM_CCFix)

                    'MsgBox("This function is not yet available")
                    myColl.SetStatus(Constants.GCST_JobStatusCreated)
                    myColl.Update()
                Catch exm As MedscreenExceptions.CanNotChangeCollectionStatus
                    MsgBox(exm.Message)
                End Try

                'See if we are coming here from rearranging the collection
                If Not TypeOf e Is Intranet.intranet.jobs.JobEvent Then

                    'Me.EditCollection(frmCollectionForm.FormModes.ChangePort)

                    'If MsgBox("Do you want to reassign this collection", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    '    myColl.Update()
                    '    Me.mnuAssignActive_Click(Nothing, Nothing)
                    'End If
                End If

            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuAssignComplete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.mnuAssignActive_Click(sender, e)
        End Sub

        Private Sub mnuAssignActive_Click(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            If myColl Is Nothing Then Exit Sub 'If we don't have a valid collection bog off 

            'If myColl.Status = Constants.GCST_JobStatusCollected Then
            '    MsgBox("Can't (re)assign collection it has already been done", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            '    Exit Sub
            'End If

            'If myColl.Status = Constants.GCST_JobStatusReceived Then
            '    MsgBox("Can't (re)assign collection it has already been done", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            '    Exit Sub
            'End If

            'If myColl.Status = Constants.GCST_JobStatusCommitted Then
            '    MsgBox("Can't (re)assign collection it is finished", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            '    Exit Sub
            'End If

            'If myColl.Status = Constants.GCST_JobStatusConfirmed Then
            '    'Dim dlgret As MsgBoxResult = MsgBox("Do you want to unconfirm the collection?, so you can reassign it", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
            '    'If dlgret = MsgBoxResult.No Then Exit Sub
            '    Try
            '        myColl.CollRedirectReport()
            '        myColl.Comments.CreateComment("Collection Unconfirmed", Constants.GCST_COMM_CCFix)

            '        'MsgBox("This function is not yet available")
            '        myColl.SetStatus(Constants.GCST_JobStatusCreated)
            '        myColl.Update()
            '    Catch exm As MedscreenExceptions.CanNotChangeCollectionStatus
            '        MsgBox(exm.Message)
            '    End Try
            'End If



            CommonForms.AssignOfficer(myColl)    'Make a call to global routine to assign officer 
            UpdateSender(SenderType)
            'Forms.Cursor.Current = Cursors.Default
            'Me.Cursor = Cursors.Default
            'lvItemCurrent.Refresh()         'Update display 

            'End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub MnuViewOfficerDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Try
                If Not myColl.CollectingOfficer Is Nothing Then
                    Medscreen.TextEditor(myColl.CollectingOfficer.ContactDetails, True, , "Contact details")
                End If
            Catch ex As Exception

            End Try
            MyBase.UpdateSender(SenderType)

        End Sub

        Private frmMoveBC As frmBarcodes

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' add barcodes or move barcodes into a collection 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [25/02/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuAddBarcodes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If objColl Is Nothing Then Exit Sub

            If frmMoveBC Is Nothing Then
                frmMoveBC = New frmBarcodes()
            End If

            Dim nSamplesProcessed As Integer = 0
            frmMoveBC.CMID = objColl.ID
            If frmMoveBC.ShowDialog = DialogResult.OK Then


            End If

        End Sub

        Private Sub mnuResendCerts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If objColl Is Nothing Then Exit Sub

            ' step 1 get list of certificates

            Dim StrRepList As String = CConnection.PackageStringList("lib_collectionxml8.GetCollectionReportList", objColl.ID)
            If Not StrRepList Is Nothing AndAlso StrRepList.Trim.Length > 0 Then 'We have reports
                Dim RepList As String() = StrRepList.Split(New Char() {","})
                Dim i As Integer
                If RepList.Length > 1 Then
                    MsgBox("There are " & RepList.Length & " Sets of certificates they will be sent separately", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
                Dim strRet As String = ""
                For i = 0 To RepList.Length - 1
                    Dim strep As String = RepList.GetValue(i)
                    Dim StrRepFileName As String = strep
                    If strep.Trim.Length > 0 Then                                'We have a report we need to strip off after the '.'
                        Dim intpos = InStr(strep, ".")
                        If intpos > 0 Then
                            strep = Mid(strep, intpos + 1)
                            Dim oColl As New Collection                             'Initalise the parameter collection
                            oColl.Add(strep)
                            oColl.Add(objColl.CollectionDate.ToString("dd-MMM-yy"))

                            'Dim strxml As String = CConnection.PackageStringList("lib_collectionxml8.GetCollectionReportsSent", oColl, True)
                            Dim strBarcodes As String = CConnection.PackageStringList("lib_collectionxml8.GetReportSampleList", strep)
                            'If Not strxml Is Nothing Then strRet += strxml
                            If Not strBarcodes Is Nothing Then
                                'strRet = Mid(strRet, 1, strRet.Length - 9)
                                strRet = "Barcodes in this report :- " & strBarcodes
                                strRet = Medscreen.ReplaceString(strRet, ",", ", ")
                                MsgBox(strRet, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                                SampleMenu.DoResendCertificate(StrRepFileName)
                            End If
                        End If
                    End If


                Next
            End If


        End Sub


        Private Sub mnuMailOfficer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If myColl.CollectingOfficer Is Nothing Then
                MsgBox("No collecting officer assigned!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If
            If myColl.CollectingOfficer.CanEmailorFax Then
                If myColl.CollectingOfficer.SendDestination.Trim.Length = 0 Then
                    MsgBox("Officer doesn't have an email address", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
                Dim objOutlook As Outlook.ApplicationClass

                Try
                    objOutlook = New Outlook.ApplicationClass()
                    Dim ns As Outlook.NameSpace
                    Dim fdMail As Outlook.MAPIFolder

                    ns = objOutlook.GetNamespace("MAPI")
                    ns.Logon(, , True, True)
                    fdMail = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox)


                    Dim objMessage As Outlook.MailItem = objOutlook.CreateItemFromTemplate(MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\dbreports\templates\collofficer.Oft")
                    objMessage.Subject = "Re : Collection " & myColl.ID & " at " & myColl.Port & " for " & myColl.VesselName
                    objMessage.To = myColl.CollectingOfficer.SendDestination
                    'objMessage.Body += " "
                    'fdMail.Items.Add(objMessage)
                    objMessage.Display(True)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly)
                    Debug.WriteLine(ex.ToString)
                Finally
                    objOutlook = Nothing
                End Try

            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuEdOfficer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If myColl.CollectingOfficer Is Nothing Then Exit Sub
            Dim frmedColl As New frmCollOfficerEd()

            frmedColl.CollectingOfficer = myColl.CollectingOfficer
            If frmedColl.ShowDialog = Windows.Forms.DialogResult.OK Then
                myColl.CollectingOfficer = frmedColl.CollectingOfficer
                myColl.CollectingOfficer.Update()

            End If

            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Try
                'blnStopRefresh = True

                If InStr("CDRM", myColl.Status) > 0 Or myColl.JobName.Trim.Length > 0 Then
                    MsgBox("This collection has already been Confirmed", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                    'nwInfo.Notify("This collection has already been Confirmed " & objColl.JobName)
                    Exit Sub
                End If
                Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", myColl.ID)
                CommonForms.SendToCollector(myColl, False, True, "Sending", sendMethod)

            Catch ex As MedscreenExceptions.WordOutputException
                MsgBox(ex.ToString)
            Catch ex As Exception
                Medscreen.LogError(ex, , "mnu send diary")
            Finally
                'blnStopRefresh = False
            End Try
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub CollectionContext_PopUp(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            ' now need to set enabled state of menu items Sender will be a contextmenu
            Dim myItem As MenuItem
            Dim ctx As ContextMenu
            ctx = sender
            For Each myItem In ctx.MenuItems
                If myColl IsNot Nothing Then SetItemStatus(myItem, myColl.Status)
            Next

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Recurse through the menu items setting whether they are enabled or not 
        ''' </summary>
        ''' <param name="MyItem"></param>
        ''' <param name="Status"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	17/10/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub SetItemStatus(ByVal MyItem As MenuItem, ByVal Status As String)
            'MyItem.Enabled = True
            If TypeOf MyItem Is TaggedMenuItem Then
                Dim aTag As TaggedMenuItem = MyItem
                If aTag.Group <> "UNCOMMIT" Then aTag.Enabled = True
                aTag.ResetText()
                If InStr(aTag.MenuTag, Status) > 0 Then MyItem.Enabled = False
                If InStr(aTag.ChangeStatus, Status) > 0 Then
                    aTag.ChangeText()
                End If
            End If
            Dim aItem As MenuItem
            For Each aItem In MyItem.MenuItems
                SetItemStatus(aItem, Status)
            Next
        End Sub



        '''<summary>
        '''Confirm the collection i.e. Collecting officer has said they can do the collection
        '''Result is a SM job will be created and status will be C
        ''' </summary>
        Private Sub mnuConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)


            Try
                '              blnStopRefresh = True           'Stop list refresh

                If Not myColl Is Nothing Then      'If we don't have a valid collection bog off 
                    If myColl.Status = Constants.GCST_JobStatusConfirmed Then
                        MsgBox("This collection has already been confirmed", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                        'nwInfo.Notify("This collection has already been confirmed")
                        Exit Sub
                    End If
                    'Check to see if we have a job if so, it will have been previously confirmed and just needs the status changing.
                    If myColl.JobName.Trim.Length > 0 Then
                        myColl.SetStatus(Constants.GCST_JobStatusConfirmed)
                        myColl.Update()                                           'Do an update to save the changes before the refresh
                    Else
                        CommonForms.ConfirmCollection(myColl) 'Use common confirm routine
                    End If


                    myColl.Refresh()                   'Ensure any changes have been updated

                    'Me.lvItemCurrent.Refresh()          'Refresh list item
                    'Me.lvItemCurrent.Selected = True
                    myColl.Refresh()
                    If myColl.JobName.Trim.Length > 0 Then     ' If the external process has completed 
                        '               Display the job details
                        '               If it hasn't no issue as it will display when it completes but no message
                        MsgBox("Job Name is : " & myColl.JobName)
                        'nwInfo.Notify("Job Name is : " & myColl.JobName)
                    End If
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "mnu confirm")
            Finally
                'blnStopRefresh = False
            End Try
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuAmendColl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            ''
            If Not myColl Is Nothing Then
                CommonForms.EditCollection(myColl, frmCollectionForm.FormModes.Edit)
            End If
            MyBase.UpdateSender(SenderType)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Put a collection to status on hold (I)
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' Collections can't be completed or on credit card hold or Cancelled, this will be handled in the menu config file
        ''' A comment needs to be added to indicate what has been done, the user needs to be able to add their own comment.
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [07/09/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuPutCollectionOnHold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            myColl.SetStatus(Constants.GCST_JobStatusInterrupted)
            Dim objComment As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment for why Processing Held", "Comment")
            If objComment Is Nothing Then objComment = "User supplied no Comment"
            myColl.Comments.CreateComment(objComment, "HOLD", jobs.CollectionCommentList.CommentType.Collection)
            myColl.Update()
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuResumeCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim strStatus As String = myColl.OldStatus
            If strStatus = Constants.GCST_JobStatusInterrupted Then
                MsgBox("Can't determine what the collection status was before it was interrupted, please let IT know, please provide CM number " & myColl.ID, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            myColl.SetStatus(strStatus)
            Dim objComment As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment for why Processing Resumed", "Comment")
            If objComment Is Nothing Then objComment = "User supplied no Comment"
            myColl.Comments.CreateComment(objComment, "RESUME", jobs.CollectionCommentList.CommentType.Collection)
            myColl.Update()
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuRandomNumbers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim frmRandom As New frmRandomNumbers
            frmRandom.Collection = myColl
            If frmRandom.ShowDialog = DialogResult.OK Then

            End If
        End Sub


        Private Sub mnuCardiffInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            'If myColl.TestScheduleID = "HAIR1" Then
            Dim myService As New MedscreenLib.ts011.CardiffService()
            Dim strXML As String = myService.GetReceivedSamplesByCM(myColl.ID, "")
            Dim strStylesheet As String = MedscreenCommonGUIConfig.NodeStyleSheets("CMCardiffView")
            Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strStylesheet, 0)
            Medscreen.ShowHtml(strHTML)
            'End If
        End Sub

        Private Sub mnuCardPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            MedscreenCommonGui.CommonForms.DoCardPay(myColl, True)
            MyBase.UpdateSender(SenderType)
        End Sub
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
        Private Sub mnuCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles mnuCancel.Click
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim nt As System.Security.Principal.WindowsIdentity
            Dim p As System.Security.Principal.WindowsPrincipal
            'Dim lvs As ListViewItem.ListViewSubItem
            'Dim cs As ListViewItems.CollStatusSubItem

            nt = System.Security.Principal.WindowsIdentity.GetCurrent
            p = New System.Security.Principal.WindowsPrincipal(nt)

            'We need to check that the user has sufficient rights to do this they need to be 
            'a memeber of CM_aadministrator 
            Debug.WriteLine(nt.Name)
            If p.IsInRole("CONCATENO\CM_Administrator") Or blnCanCancel Then
                'If they haven't used the application for twenty minutes check their password
                'If Now.Subtract(Me.LoginTimeStamp).TotalMinutes > 20 Then  'Validate user 
                'End If

                'Stop refresshing the list, this will confuse matters
                'Me.blnStopRefresh = True
                'Get collection object
                'myColl = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)
                If Not myColl Is Nothing Then          'If we have a valid collection
                    If myColl.Status = Constants.GCST_JobJobStatusCancelled Then   'Has it already been cancelled
                        'MsgBox("This collection has already been cancelled", MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
                        Try
                            myColl.UnCancel()
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        Finally
                            myColl.Refresh()
                        End Try
                        Exit Sub
                    End If
                    If myColl.Status = Constants.GCST_JobStatusPay Then 'Ir isn't a collection
                        MsgBox("This is not a collection but a pay statement it can't be cancelled!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                    CommonForms.CancelCollection(myColl)         'Go to common cancel collection routine
                    myColl.Refresh()                           'Get changed details
                End If

                'Me.blnStopRefresh = False
                CommonForms.SendEmailToManager(myColl)                     'Send an email to let people know of the cancellation 
                'Me.GetCollectionDetails()                       
            Else
                MsgBox("You have insufficent rights to perform this operation, please contact IT")
                'nwInfo.Notify("You have insufficent rights to perform this operation, please contact IT")
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuRescan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            Dim strSendType As String

            If myColl Is Nothing Then Exit Sub ' No valid collection exit 
            myColl.GetScannedDocuments()
        End Sub
        Private Sub mnuResend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            '
            'Me.mnuSend_Click(Nothing, Nothing)  'Temporay redirection to test code 
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            Dim strSendType As String

            If myColl Is Nothing Then Exit Sub ' No valid collection exit 

            If (InStr("WV", myColl.Status) <> 0) Or (myColl.CollOfficer.Trim.Length = 0) Then 'Waiting or not assigned
                MessageBox.Show("Can not resend : collection has not been assigned, you will be taken to assignment", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Me.mnuAssignActive_Click(sender, e)
                Exit Sub
            End If


            'Deal with callouts they have different documents
            If myColl.CollClassType = jobs.CMJob.CollectionTypes.GCST_JobTypeCallout Or _
                myColl.CollClassType = jobs.CMJob.CollectionTypes.GCST_JobTypeCalloutOverseas Then    'Deal with callouts 
                myColl.CallOutConfirmation()
                Exit Sub
            End If

            Dim blnPrintOnly As Boolean = False
            Select Case MsgBox("Would you like to Re-send collection " & myColl.ID & "?" & vbCrLf & _
                               "to " & myColl.CollOfficer & vbCrLf & _
                               "(No = Print collection only)", vbYesNoCancel + vbQuestion)
                Case vbYes
                    strSendType = "re-sent"

                Case vbNo
                    strSendType = "printed"
                    blnPrintOnly = True
                Case vbCancel
                    Exit Sub
            End Select
            Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", myColl.ID)
            CommonForms.SendToCollector(myColl, blnPrintOnly, , "Resend", sendMethod)
            'WriteResendCollectionDetails(strCollRef, strSendType)
            MyBase.UpdateSender(SenderType)

        End Sub


        Private Sub mnuResendCOConfirmation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Send out callout request to cofficer.
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl.CollectionType = Constants.GCST_JobTypeCallout Then
                myColl.CallOutConfirmation()
            Else
                Dim sendMethod As String = CConnection.PackageStringList("Lib_collection.OfficerSendMethod", myColl.ID)
                CommonForms.SendToCollector(myColl, False, , "Resend Confirmation", sendMethod)
            End If
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuResendConf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            With myColl
                If .ConfirmationType.Trim = "COMPLETED ONLY" OrElse .ConfirmationType.Trim = "BOTH ARRANGED AND COMPLETED" Then
                    If .CollectionDate > Now Then
                        .CustomerConfirmation(MedscreenLib.Constants.ConfirmationType.Arrange)
                    Else
                        .CustomerConfirmation(MedscreenLib.Constants.ConfirmationType.Confirm)
                    End If
                End If
            End With
            MyBase.UpdateSender(SenderType)
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
        Private Sub mnuCollDocReceived_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            If Not (myColl.Status = Constants.GCST_JobStatusConfirmed Or _
                myColl.Status = Constants.GCST_JobStatusReceived) Then
                MessageBox.Show("Only confirmed or received collections can be set as Collected", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If
            With myColl
                If MessageBox.Show("Do you wish to record that the documents indicating that the collection has been done, have been received for " _
                    & .CMID & "?", "Collection Done", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
                    Dim objRet As Object
                    Try
                        objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                            "Collection Date", "Collection Date", .CollectionDate)
                        If objRet Is Nothing Then Exit Sub
                    Catch ex As Exception
                        objRet = Medscreen.GetParameter(Medscreen.MyTypes.typDate, _
                            "Collection Date", "Collection Date", Now.AddDays(3))
                        If objRet Is Nothing Then Exit Sub
                    End Try
                    Dim dateRec As Date = objRet
                    objRet = Medscreen.GetParameter(Medscreen.MyTypes.typeInteger, _
                        "No of Donors", "No of Donors", .NoOfDonors)
                    If objRet Is Nothing Then Exit Sub
                    Dim NDonors As Integer = objRet
                    If Not .VesselObject Is Nothing Then ' we have a vessel
                        If Mid(.VesselObject.VesselID, 1, 4) = "VESS" Or _
                            Mid(.VesselObject.VesselID, 1, 3) = "IMO" Then 'We have no ID
                            objRet = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Please enter any vessel id given on crew list", "", "")
                            .CreditCardNumber = objRet
                        End If
                    End If
                    If .Status = Constants.GCST_JobStatusConfirmed Then .Status = Constants.GCST_JobStatusCollected
                    .Comments.CreateComment("Documents received from collecting officer", "CONF")
                    .RecordCollectionAction("DOCREC", dateRec, NDonors)
                    .Medscreenref = NDonors
                    .DateShipped = dateRec
                    .CollectionDate = dateRec
                    objRet = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Airway Bill Number", , "")

                    If objRet Is Nothing Then Exit Sub
                    Dim strAirway As String = objRet
                    strAirway = Medscreen.ReplaceString(strAirway, " ", "")
                    .AirwayBillNo = strAirway
                    .Update()
                    .Refresh()
                End If
            End With
            MyBase.UpdateSender(SenderType)

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Set the agreed costs for a collection
        ''' </summary>
        ''' <param name="aColl"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/11/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub SetAgreedCosts(ByVal aColl As Intranet.intranet.jobs.CMJob)
            If aColl Is Nothing Then Exit Sub
            If aColl.SmClient Is Nothing Then Exit Sub
            Dim strOutCurrency As String = aColl.SmClient.Currency
            Dim ctest As MedscreenCommonGui.CurrencyConvert = New CurrencyConvert()
            ctest.InCurrency = strOutCurrency
            ctest.OutCurrency = strOutCurrency
            ctest.LockCurrency = True
            ctest.Value = aColl.AgreedExpenses
            Dim newValue As Double
            If ctest.ShowDialog = DialogResult.OK Then
                newValue = ctest.Value
                aColl.AgreedExpenses = newValue
                aColl.Update()
                'check to see if it has been invoiced already 
                If aColl.Invoiceid.Trim.Length > 0 Then
                    Try
                        MsgBox("This collection has already been invoiced " & aColl.Invoiceid & "; an email will be sent to accounts to get it reinvoiced to include this charge!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                        Dim strEmail As String = MedscreenCommonGUIConfig.CollectionQuery("InvoiceNotification")
                        Dim strParams As String = "agreed extra costs," & aColl.CMID & "," & aColl.Invoiceid & "," & aColl.SmClient.SMIDProfile & " - " & aColl.SmClient.SageID
                        Dim parray As String() = strParams.Split(New Char() {","})

                        strEmail = String.Format(strEmail, parray)

                        'We need to email accounts.
                        Dim strEmailToRole As String = MedscreenCommonGUIConfig.CollectionQuery("InvoiceNotificationTo")
                        Dim strEmailTo As String = Medscreen.RoleEmailList(strEmailToRole)
                        Medscreen.BlatEmail("Invoice may need changing", strEmail, strEmailTo)
                    Catch
                    End Try
                End If
            End If

        End Sub

        Private Sub mnuCollAgreedCosts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            SetAgreedCosts(myColl)
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuProblem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            If myColl Is Nothing Then Exit Sub
            If myColl.Status = Constants.GCST_JobStatusPay Then
                MsgBox("This action is not relevant to Pay collections")
                'nwInfo.Notify("This action is not relevant to Pay collections")
            End If
            If myColl.Status = Constants.GCST_JobStatusCommitted Then
                MsgBox("This action is not relevant to completed collections")
                'nwInfo.Notify("This action is not relevant to completed collections")
            End If

            Dim afrm As New FrmCollComment()
            afrm.FormType = FrmCollComment.FCCFormTypes.CollectorComment
            afrm.CommentType = "PROB"
            afrm.CMNumber = myColl.ID
            If myColl.ProgrammeManager.Trim.Length > 0 Then
                afrm.ProgamManager = myColl.ProgrammeManager
            End If
            If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                    myColl.Comments.CreateComment(afrm.Comment, afrm.CommentType)
                If Not aComm Is Nothing Then
                    aComm.ActionBy = afrm.ActionBy
                    aComm.Update()
                    myColl.CancellationStatus = Constants.GCST_CANC_Problem
                    myColl.Update()
                    myColl.Refresh(False)
                    Dim UP As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.Personnel.Item(afrm.ActionBy, _
                        MedscreenLib.personnel.PersonnelList.FindIdBy.identity)
                    Dim strSubject As String = "Issue with Collection CM" & myColl.ID
                    If Not myColl.VesselObject Is Nothing Then
                        strSubject += " vessel " & myColl.VesselName
                    ElseIf Not myColl.SiteObject Is Nothing Then
                        strSubject += " site " & myColl.VesselName
                    Else
                        strSubject += myColl.vessel
                    End If

                    strSubject += " Port " & myColl.Port
                    strSubject += " Officer " & myColl.CollOfficer
                    If Not UP Is Nothing Then
                        Medscreen.BlatEmail(strSubject, _
                        afrm.Comment, afrm.txtDestination.Text, cSupport.User.Email)
                    End If
                    If afrm.ckCopy.Checked Then
                        Medscreen.BlatEmail(strSubject, _
                        afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)
                    End If
                    myColl.LogAction(jobs.CMJob.Actions.RaiseIssue, afrm.txtCopyDest.Text)

                End If
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuResolveProblem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            If myColl Is Nothing Then Exit Sub
            myColl.CancellationStatus = "N"
            myColl.Update()

            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuPoEntry_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Dim objPo As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Purchase Order", "Purchase Order No Entry", myColl.PurchaseOrder)
            If objPo Is Nothing Then Exit Sub
            myColl.PurchaseOrder = CStr(objPo)
            myColl.Update()
            If myColl.Invoiceid.Trim.Length > 0 Then
                Try
                    MsgBox("This collection has already been invoiced " & myColl.Invoiceid & "; an email will be sent to accounts to get it reinvoiced to include this charge!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                    Dim strEmail As String = MedscreenCommonGUIConfig.CollectionQuery("InvoiceNotification")
                    Dim strParams As String = "purchase order no," & myColl.CMID & "," & myColl.Invoiceid & "," & myColl.SmClient.SMIDProfile & " - " & myColl.SmClient.SageID
                    Dim parray As String() = strParams.Split(New Char() {","})

                    strEmail = String.Format(strEmail, parray)

                    'We need to email accounts.
                    Dim strEmailToRole As String = MedscreenCommonGUIConfig.CollectionQuery("InvoiceNotificationTo")
                    Dim strEmailTo As String = Medscreen.RoleEmailList(strEmailToRole)
                    Medscreen.BlatEmail("Invoice may need changing", strEmail, strEmailTo)
                Catch
                End Try
            End If

            MyBase.UpdateSender(SenderType)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send collection invoice copy to user
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/01/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuSendInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If myColl.Invoiceid.Trim.Length = 0 OrElse myColl.Invoice Is Nothing Then
                MsgBox("This collection doesn't appear to have been invoiced yet", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                Exit Sub
            End If

            Dim ACCInt As AccountsInterface = New AccountsInterface()
            ACCInt.PrintInvoice(myColl.Invoiceid, myColl.Invoice.TransactionType)

        End Sub
        ''' <developer>CONCATENO\Taylor</developer>
        ''' <summary>
        ''' Add an invoice to collection
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory>
        ''' <revision><modified>23-Dec-2011 07:44 Tidied exception message</modified><Author>CONCATENO\Taylor</Author></revision>
        ''' <revision><created>20-Dec-2011 11:08</created><Author>CONCATENO\Taylor</Author></revision>
        ''' </revisionHistory>
        Private Sub mnuAddInvoiceNumber_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If myColl.Invoiceid.Trim.Length > 0 OrElse myColl.Invoice IsNot Nothing Then
                MsgBox("This collection has already been invoiced", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Exit Sub
            End If

            Dim objInvoice As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Invoice Number", "Invoice Number for " & myColl.ID)
            If objInvoice Is Nothing Then '             user canceleld
                Exit Sub
            End If

            Dim ocmd As New OleDb.OleDbCommand With {.CommandText = "ASSIGNINVOICE", .CommandType = CommandType.StoredProcedure, .Connection = CConnection.DbConnection}
            Dim intRet As Integer
            Try
                ocmd.Parameters.Add(CConnection.StringParameter("cm", myColl.ID, 10))
                ocmd.Parameters.Add(CConnection.StringParameter("invNum", objInvoice, 10))
                If CConnection.ConnOpen Then
                    intRet = ocmd.ExecuteNonQuery
                End If
            Catch ex As OleDb.OleDbException
                Dim strMessage As String = CConnection.TidyOracleUserException(ex.Message)
                Medscreen.LogError(strMessage & " Invoice number : " & objInvoice, True)

            Catch ex As Exception
                Medscreen.LogError(ex, True, "Invoice number : " & objInvoice)
            Finally
                CConnection.SetConnClosed()
            End Try

            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuAgreedExtraCosts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Dim frmCurr As New frmCurrencyEntry()

            frmCurr.OutputCurrency = myColl.SmClient.Currency
            frmCurr.InputCurrency = myColl.SmClient.Currency
            frmCurr.InputAmount = myColl.AgreedExpenses
            If frmCurr.ShowDialog = DialogResult.OK Then

            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuAddComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Dim afrm As New FrmCollComment()
            afrm.CommentType = "MISC"
            afrm.FormType = FrmCollComment.FCCFormTypes.def
            afrm.ProgamManager = myColl.ProgrammeManager
            afrm.CMNumber = myColl.ID

            'afrm.CollOfficer = mycoll.CollOfficer

            If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                    myColl.Comments.CreateComment(afrm.Comment, afrm.CommentType)
                aComm.ActionBy = afrm.ActionBy
                aComm.Update()
                myColl.Update()
                myColl.Refresh(False)
                If afrm.ckCopy.Checked Then
                    Dim strSubject As String = "Comment has been made on collection : " & myColl.ID
                    Medscreen.BlatEmail(strSubject, _
                    afrm.Comment, afrm.txtCopyDest.Text, cSupport.User.Email)

                End If
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' View the collection as a workflow
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	01/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If (myColl.HasID > 0) Then
                Dim tf As New frmWorkflow()

                tf.DisplayType = 0
                tf.Collection = myColl
                tf.User = MedscreenLib.Glossary.Glossary.CurrentSMUser
                tf.Action = Intranet.intranet.jobs.CMJob.Actions.Display
                tf.ShowDialog()

            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Show job cost
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	01/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuJobCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Try
                Dim Cost As Double
                Dim strHtml As String = myColl.PriceJobHTML(Cost)
                MedscreenLib.Medscreen.ShowHtml(strHtml)
            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)
        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' Handles edit menu comments
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>08-Dec-2011 14:55</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Private Sub mnuEditComments_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Try
                Dim myComment As Object = myColl.CollectionComments
                myComment = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Collection Comment", , myComment, False)
                If Not myComment Is Nothing Then
                    myColl.CollectionComments = myComment
                    myColl.Update()
                End If
            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' Code to allow editing of agents details
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>08-Dec-2011 15:00</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Private Sub mnuEditAgents_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Dim blnUpdate As Boolean = False
            Try
                Dim myAgent As Object = myColl.Agent
                myAgent = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Agent", , myAgent)
                If Not myAgent Is Nothing Then
                    myColl.Agent = myAgent
                    blnUpdate = True
                End If
                Dim myAgentsContact As Object = myColl.AgentsPhone
                myAgentsContact = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Agents Contact details", , myAgentsContact)
                If Not myAgentsContact Is Nothing Then
                    myColl.AgentsPhone = myAgentsContact
                    blnUpdate = True
                End If
                If blnUpdate Then myColl.Update()

            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

        End Sub
        ''' <developer></developer>
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>12-Dec-2011 15:58</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Private Sub mnuEditVessel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            If myColl.VesselSiteObject Is Nothing Then
                MsgBox("No vessel or site to edit", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                Exit Sub
            End If

            Dim edFrm As New frmVessel()
            edFrm.Vessel = myColl.VesselSiteObject

            If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                'aVessNode.Vessel = edFrm.Vessel
                myColl.VesselSiteObject.Update()
            End If

            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuDocView_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub
            Dim frmView As New frmViewDocs
            frmView.CollectionId = myColl.ID

            frmView.ShowDialog()
        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>08-Dec-2011 15:31</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
        Private Sub mnuEditWTT_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim myColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If myColl Is Nothing Then Exit Sub

            Try
                Dim myComment As Object = myColl.WhoToTest
                myComment = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Who to Test", , myComment, False)
                If Not myComment Is Nothing Then
                    myColl.WhoToTest = myComment
                    myColl.Update()
                End If
            Catch ex As Exception
            End Try
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuEnterPayDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim objPort As Intranet.intranet.jobs.Port
            Try
                Dim objForm As New frmCollPay()
                Dim StrOfficers As String = Intranet.intranet.jobs.CollPay.OtherOfficers(objColl.ID, objColl.CollOfficer)
                If StrOfficers.Trim.Length > 0 Then
                    MsgBox("This collection has pay statements for these Officers :- " & StrOfficers, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If

                objForm.Collection = objColl
                If objForm.ShowDialog = DialogResult.OK Then

                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub mnuEnterPayDetailsAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim objPort As Intranet.intranet.jobs.Port
            Dim objCS As New frmSelectCollector()
            Dim objOfficer As jobs.CollectingOfficer
            Try
                If objCS.ShowDialog = Windows.Forms.DialogResult.OK Then
                    objOfficer = objCS.Officer

                End If
            Catch ex As Exception
                Medscreen.LogError(ex, True)
            End Try

            Try
                Dim objForm As New frmCollPay()
                'objForm.Collection = objColl
                Dim oPay As New jobs.CollPay
                oPay.CID = objColl.ID
                oPay.Officer = objOfficer.OfficerID
                oPay.Donors = objColl.NoOfDonors
                oPay.PayType = ""
                oPay.DoUpdate()
                objForm.PayItem = oPay
                If objForm.ShowDialog = DialogResult.OK Then
                    oPay = objForm.PayItem
                    oPay.DoUpdate()
                Else
                    oPay.Delete()
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub mnuAssignPayDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            Dim objPort As Intranet.intranet.jobs.Port
            MsgBox("This menu is not now supported", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            'Try
            '    If Not objColl Is Nothing Then
            '        If InStr("AVPWSC", objColl.Status) <> 0 Then
            '            MsgBox("Pay statements can only be entered for jobs that have been done or have samples", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
            '            Exit Sub
            '        End If
            '        Dim CollPayForm As New MedscreenCommonGui.frmCollPayEntry()
            '        With CollPayForm
            '            'objColl = Me.MyCollections.Item(Me.lvItemCurrent.CMJobNumber)
            '            If Not objColl Is Nothing Then
            '                objPort = cSupport.Ports.Item(objColl.Port)
            '                If Not objPort Is Nothing Then
            '                    'If objPort.IsFixedSite Then
            '                    'MsgBox("No collector pay at fixed site ports", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)
            '                    'Else '
            '                    If Date.Compare(objColl.CollectionDate, Now) > 0 Then
            '                        MsgBox("Collection not yet done!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            '                    Else
            '                        Dim objCollPay As Intranet.collectorpay.CollectorPayCollection = New Intranet.collectorpay.CollectorPayCollection()

            '                        objCollPay.Load("CM_JOB_NUMBER = '" & objColl.ID & "'")

            '                        Dim objCollPayItem As Intranet.collectorpay.CollectorPay

            '                        If objCollPay.Count > 0 Then
            '                            objCollPayItem = objCollPay.Item(objColl.ID, Intranet.collectorpay.CollectorPayCollection.IndexPayOn.CMNumber)
            '                            If objCollPayItem.PayType.Trim.Length = 0 Then objCollPayItem.PayType = "C"
            '                            If objCollPayItem.CollOfficer.Trim.Length = 0 Then objCollPayItem.CollOfficer = objColl.CollOfficer
            '                            .CollPay = objCollPayItem
            '                        Else
            '                            objCollPayItem = objCollPay.Create
            '                            objCollPayItem.PayType = "C"
            '                            objCollPayItem.CMidentity = objColl.ID
            '                            objCollPayItem.CollOfficer = objColl.CollOfficer
            '                            objCollPayItem.PaymentMonth = DateSerial(Today.Year, Today.Month, 1)
            '                            If Not objColl.CollectingOfficer Is Nothing Then
            '                                If objColl.CollectionType = "O" Or objColl.CollectionType = "C" Then
            '                                    objCollPayItem.PayMethod = "H"
            '                                Else
            '                                    objCollPayItem.PayMethod = objColl.CollectingOfficer.PayMethod
            '                                End If
            '                            End If
            '                            objCollPayItem.DoUpdate()
            '                            objCollPayItem.CreateEntries()
            '                            .CollPay = objCollPayItem
            '                        End If

            '                        If .ShowDialog() = Windows.Forms.DialogResult.OK Then
            '                            .CollPay.DoUpdate()
            '                        End If
            '                    End If
            '                End If
            '                'End If
            '            End If
            '        End With
            '    End If
            'Catch ex As Exception
            '    Medscreen.LogError(ex, True)
            'End Try
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuUnAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                Dim SenderType As Object = GetSenderType(sender)
                Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
                Dim oCmd As New OleDb.OleDbCommand("lib_Uncommit.UnAuthoriseJob", CConnection.DbConnection)
                Dim intRet As Integer
                Dim objNonc As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Non Conformance no", , "NONC")
                If objNonc Is Nothing Then Exit Sub
                Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                If objReason Is Nothing Then
                    objReason = ""
                End If
                Dim strReason As String = CStr(objReason)
                Try
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("jobName", objColl.JobName, 40))
                    oCmd.Parameters.Add(CConnection.StringParameter("noncRef", CStr(objNonc), 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery
                    End If
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "UnAuthorise Job")

                End Try
                MyBase.UpdateSender(SenderType)

            End If

        End Sub

        Private Sub mnuUnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                Dim SenderType As Object = GetSenderType(sender)
                Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
                Dim oCmd As New OleDb.OleDbCommand("lib_Uncommit.UnreportJob", CConnection.DbConnection)
                Dim intRet As Integer
                Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                If objReason Is Nothing Then
                    Exit Sub
                End If
                Dim strReason As String = CStr(objReason)
                Try
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("jobName", objColl.JobName, 40))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery
                    End If
                Catch ex As Exception
                End Try
                MyBase.UpdateSender(SenderType)

            End If

        End Sub

        Private Sub mnuUnCommit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                If objReason Is Nothing Then
                    Exit Sub
                End If
                Dim strReason As String = CStr(objReason)
                Dim SenderType As Object = GetSenderType(sender)
                Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
                Dim oCmd As New OleDb.OleDbCommand("lib_Uncommit.UncommitJob", CConnection.DbConnection)
                Dim intRet As Integer
                Try
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("jobName", objColl.JobName, 40))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery
                    End If
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "UnCommit Job")
                End Try
                MyBase.UpdateSender(SenderType)
            End If

        End Sub


        Private Sub ChangeJobCustomer(ByVal Retest As String, ByVal objcoll As jobs.CMJob)
            'First task get new customer 
            Dim objCust As Intranet.intranet.customerns.Client
            Dim objSelCust As New MedscreenCommonGui.frmCustomerSelection2()
            If objSelCust.ShowDialog = DialogResult.OK Then
                objCust = objSelCust.Client

                Dim objNonc As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Non Conformance no", , "NONC")
                If objNonc Is Nothing Then Exit Sub
                Dim oCmd As New OleDb.OleDbCommand("lib_uncommit.ChangeJobCustomer", CConnection.DbConnection)
                Dim intRet As Integer
                Dim strReason As String = Retest
                'if there are samples we need to test if the test panel is similar
                If objcoll.CollectionSamples.Count > 0 Then
                    Dim oColl As New Collection()
                    oColl.Add(objcoll.CollectionSamples(0).Barcode)
                    oColl.Add(objCust.Identity)
                    Dim strTestPanel As String = CConnection.PackageStringList("lib_uncommit.GetTestPanelStatus", oColl)
                    If Not strTestPanel Is Nothing AndAlso strTestPanel.Trim.Length > 0 AndAlso Retest <> "RETEST" Then
                        Dim strMsg = "Tests that may need to be performed are " & vbCrLf & strTestPanel & vbCrLf & vbCrLf & _
                        "If there are any Xs(Not Done) or Ts(Not on Panel)listed the sample should be retested" & vbCrLf & _
                        "You chose " & Retest & " Initially" & vbCrLf & _
                        "Do you want to reconsider your choice?"
                        Dim intRes As MsgBoxResult = MsgBox(strMsg, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question)
                        If intRes = MsgBoxResult.Yes Then
                            'code to change selection
                            Dim cColl As New Collection()
                            cColl.Add("RETEST")
                            cColl.Add("REPORT")
                            cColl.Add("INVOICE")
                            Dim objRet As Object
                            objRet = Medscreen.GetParameter(Medscreen.MyTypes.typItem, "Choice", "What to do with samples", "RETEST", , cColl)
                            If objRet Is Nothing Then
                                Exit Sub
                            End If
                            strReason = CStr(objRet)
                        ElseIf intRes = MsgBoxResult.Cancel Then    'User wants to bog off
                            Exit Sub
                        End If
                    End If
                End If

                Try
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("jobName", objcoll.JobName, 40))
                    oCmd.Parameters.Add(CConnection.StringParameter("newCustID", objCust.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("noncRef", CStr(objNonc), 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("action", strReason, 10))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery
                    End If
                    Dim CommentText As String = "Collection moved from - " & objcoll.SmClient.SMIDProfile & _
                       " to - " & objCust.SMIDProfile & " in response to Non Conformance: " & CStr(objNonc)
                    objcoll.Comments.CreateComment(CommentText, "FIX")
                    CollectionEmail(objcoll.SmClient.SMIDProfile, objCust.SMIDProfile, Retest, objcoll)
                    objcoll.SmClient = Nothing
                    objcoll.ClientId = objCust.Identity

                Catch ex As Exception
                    Medscreen.LogError(ex, True, "ChangeJob Customer")
                End Try


            End If
        End Sub

        Private Sub mnuFindNearestOfficer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)

            If objColl Is Nothing Then Exit Sub
            'we can approach this two ways, if the collection is at a customer's site we can use the customer's site address
            'Otherwise we have to rely on the port information.
            FindNearestOfficer(objColl)
        End Sub


        Public Shared Sub FindNearestOfficer(ByVal objColl As Intranet.intranet.jobs.CMJob)
            Dim strXML As String
            Dim StrXMLOuter As String = "<Nearest>"
            Dim strPLSQLFunction As String
            Dim oColl As New Collection()
            If Not objColl.CustomerSite Is Nothing AndAlso Not objColl.CustomerSite.Address Is Nothing AndAlso objColl.CustomerSite.Address.PostCode.Trim.Length > 0 Then
                Dim strPostCode As String = objColl.CustomerSite.Address.PostCode.Trim
                Dim strOutCode As String = ""
                Dim I As Integer = 0
                While I < strPostCode.Length AndAlso strPostCode.Chars(I) <> " "
                    strOutCode += strPostCode.Chars(I)
                    I += 1
                End While
                oColl.Add(strOutCode)
                oColl.Add(objColl.CustomerSite.Address.Country)
                strPLSQLFunction = MedscreenCommonGUIConfig.NodePLSQL.Item("NearestOfficerToOutcode")
                StrXMLOuter += "<SiteName>" & objColl.CustomerSite.SiteName & "</SiteName>"
                StrXMLOuter += objColl.CustomerSite.Address.ToXML

                StrXMLOuter += "<PostCode>" & objColl.CustomerSite.Address.PostCode.Trim & "</PostCode>"
                StrXMLOuter += "<Port>" & objColl.CollectionPort.PortDescription & "</Port>"
            Else
                oColl.Add(objColl.CollectionPort.Identity)
                oColl.Add(objColl.CollectionPort.CountryId)
                strPLSQLFunction = MedscreenCommonGUIConfig.NodePLSQL.Item("NearestOfficertoPort")
                strXML = CConnection.PackageStringList("lib_officer.ClosestOfficerstoPortXML", oColl)
                StrXMLOuter += "<Port>" & objColl.CollectionPort.PortDescription & "</Port>"
            End If
            strXML = CConnection.PackageStringList(strPLSQLFunction, oColl)

            If strXML Is Nothing OrElse strXML.Length < 20 Then
                MsgBox("This port doesn't have any geographical details Longitude or Latitude, so this information can't be calculated." & vbCrLf & "This information can be entered on the location tab for the port, you can often use Google searching on the port and latitude to find out where it is.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                Exit Sub
            End If
            strXML = StrXMLOuter & strXML & "</Nearest>"
            Dim strStyleSheet As String = MedscreenCommonGUIConfig.NodeStyleSheets.Item("NearestOfficer")
            Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strStyleSheet, 0)
            Medscreen.ShowHtml(strHTML)


        End Sub

        Private Sub mnuChangeCust_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeJobCustomer("REPORT", objColl)
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuChangeCustRetest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeJobCustomer("RETEST", objColl)
            End If
            MyBase.UpdateSender(SenderType)

        End Sub


        Private Sub mnuChangeCustInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeJobCustomer("INVOICE", objColl)
            End If
            MyBase.UpdateSender(SenderType)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' change the customer on a collection
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuChangeCustColl_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim objColl As Intranet.intranet.jobs.CMJob = GetCollection(SenderType)
            'code to change customer
            'Check status is not (XDRM)
            If InStr("XDRM", objColl.Status) <> 0 Then
                'put out bog off message
                MsgBox("I'm sorry but this can't be done with a collection of status : " & objColl.StatusString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            'Get new customer 
            Dim objCust As Intranet.intranet.customerns.Client
            Dim objSelCust As New MedscreenCommonGui.frmCustomerSelection2()
            If objSelCust.ShowDialog = DialogResult.OK Then
                objCust = objSelCust.Client
                'now all we have to do is change the customer and if confirmed the job customer as well.
                'then warn the user about the impact and add a comment
                objColl.Comments.CreateComment("Collection customer changed to" & objCust.SMIDProfile, MedscreenLib.Constants.GCST_COMM_CCFix)

                objColl.ClientId = objCust.Identity
                If objColl.JobName.Trim.Length > 0 Then
                    objColl.Job.Customer = objCust.Identity
                    objColl.Job.Update()
                End If
                'We can set the programme manager 
                objColl.ProgrammeManager = objCust.ProgrammeManager
                objColl.ProgrammeManaged = (objCust.ProgrammeManager.Trim.Length > 0)

                objColl.Update()
                Dim strChange As String = MedscreenCommonGUIConfig.CollectionQuery("CustomerChange")
                MsgBox(strChange, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            End If
            UpdateSender(SenderType)
        End Sub

        Private Sub CollectionEmail(ByVal oldCust As String, ByVal newCust As String, ByVal action As String, ByVal collection As jobs.CMJob)
            Dim strEmail As String = "This collection - " & collection.ID & " -has been moved from " & oldCust & " to " & newCust & "<BR/>"
            strEmail += "The collection consists of the following barcodes "
            Dim i As Integer
            Dim strBarcodes As String = ""
            For i = 0 To collection.CollectionSamples.Count - 1
                If strBarcodes.Length > 0 Then strBarcodes += ","
                strBarcodes += collection.CollectionSamples.Item(i).Barcode
            Next
            strEmail += strBarcodes & "<BR/>" & "<BR/>"
            If action = "INVOICE" Then
                strEmail += "Some or all of these samples will have been uncommitted"
            ElseIf action = "REPORT" Then
                strEmail += "Some or all of these samples will have been uncommitted and unreported"
            Else
                strEmail += "Some or all of these samples will have been uncommitted, unreported and unauthorised, " & _
                "as the samples require reanalysing they have been left on hold, please contact the relevant Lab Personnel to arrange for the analysis to be completed."
            End If
            collection.Refresh(False)
            strEmail += "<BR/>"
            If collection.Invoiceid.Trim.Length > 0 Then
                strEmail += "This collection id now invoiced against invoice number: " & collection.Invoiceid
            End If

            strEmail += "<BR/>" & "Please ensure that the Non confomance form is updated correctly."

            Medscreen.BlatEmail("Moving Collection to new customer", strEmail, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email, , , "Ben.boughton@concateno.com,doug.taylor@concateno.com,accounts@concateno.com")

        End Sub



        Private Shared Sub SetMenuContextState(ByVal sampleState As String, ByVal Menuitems As Menu.MenuItemCollection)
            Dim mnuItem As TaggedMenuItem
            For Each mnuItem In Menuitems
                mnuItem.Visible = True
                If mnuItem.Group = "UNCOMMIT" Then
                    If mnuItem.Mnemonic <> "P" Then mnuItem.Enabled = False
                    If mnuItem.Mnemonic = "A" And (InStr(sampleState, "A") + InStr(sampleState, "C") + InStr(sampleState, "R")) > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "R" And (InStr(sampleState, "R") + InStr(sampleState, "C")) > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "C" And InStr(sampleState, "C") > 0 Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "H" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "T" Then mnuItem.Enabled = True
                    If mnuItem.Mnemonic = "I" Then mnuItem.Enabled = True

                End If
                If Not mnuItem.MenuItems Is Nothing Then
                    SetMenuContextState(sampleState, mnuItem.MenuItems)
                End If
            Next


        End Sub

        Public Shared Sub ContextMenuRedraw(ByVal Collection As jobs.CMJob, ByVal ctxMenu As ContextMenu)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") OrElse MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                If Collection.JobName.Trim.Length > 0 Then
                    Dim SampleState As String = CConnection.PackageStringList("lib_Uncommit.GetJobStatus", Collection.JobName)
                    SetMenuContextState(SampleState, ctxMenu.MenuItems)
                End If
            Else
                Dim mnuItem As MenuItem
                For Each mnuItem In ctxMenu.MenuItems
                    If mnuItem.Text.ToUpper = "JOBS" Then mnuItem.Visible = False
                Next
            End If

        End Sub
    End Class


    Public Class CollectionHeaderMenu
        Inherits MenuBase

        Private Shared myInstance As CollectionHeaderMenu
        Private Sub New()

        End Sub

        Public Shared Function GetCollectionHeaderMenu() As CollectionHeaderMenu
            If myInstance Is Nothing Then
                myInstance = New CollectionHeaderMenu()

            End If
            Return myInstance
        End Function


        Private Sub MnuBookRoutine_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            'Dim intRet As DialogResult
            'Dim blnClient As Boolean
            'Dim oVessel As Intranet.intranet.customerns.vessel

            Dim SenderType As Object = GetSenderType(sender)

            Dim myClient As Intranet.intranet.customerns.Client = GetCustomer(SenderType)


            Dim blnRet As Boolean
            Dim objCm As Intranet.intranet.jobs.CMJob
            objCm = BookRoutine(blnRet, myClient)


            UpdateSender(SenderType)

        End Sub

 
        Public Function BookRoutine(ByRef blnRet As Boolean, ByVal oClient As Intranet.intranet.customerns.Client) As Intranet.intranet.jobs.CMJob
            Dim objCm As Intranet.intranet.jobs.CMJob = New Intranet.intranet.jobs.CMJob()
            objCm.FieldValues.Initialise()


            Cursor.Current = Cursors.WaitCursor
            Dim formCollection As New MedscreenCommonGui.frmCollectionForm()
            Try
                cSupport.Ports.RefreshChanged()    'See if the ports list needs updating 
                If oClient Is Nothing Then Exit Function
                If Not oClient.CollectionInfo Is Nothing Then
                    With oClient.CollectionInfo
                        objCm.ClientId = oClient.Identity
                        objCm.WhoToTest = .WhoToTest
                        objCm.NoOfDonors = .NoDonors
                        objCm.ConfirmationContact = .ConfirmationContactList
                        objCm.ConfirmationNumber = .ConfirmationSendAddress
                        objCm.ConfirmationMethod = .ConfirmationMethod
                        objCm.CollType = .ReasonForTest
                        objCm.ConfirmationType = .ConfirmationType
                        objCm.SpecialProcedures = .Procedures
                        objCm.CollectionComments = .AccessToSite
                        If .TestType.Trim.Length > 0 Then
                            Dim objPhrase As MedscreenLib.Glossary.Phrase = _
                                MedscreenLib.Glossary.Glossary.TestTypeList.Item(.TestType)
                            If Not objPhrase Is Nothing Then objCm.Tests = objPhrase.PhraseText
                        End If

                        If .UKCollection Then
                            objCm.UK = True
                        End If
                    End With
                End If
                Cursor.Current = Cursors.WaitCursor
                With formCollection

                    .Collection = objCm
                    If .ShowDialog = DialogResult.OK Then 'If there is a low level issue will be set to retry
                        'TODO add to list offer to assign to collector
                        blnRet = True
                    Else
                        blnRet = False
                    End If
                End With
                formCollection = Nothing
                Cursor.Current = Cursors.Default
            Catch ex As Exception
                Medscreen.LogError(ex)
            End Try
            Return objCm


        End Function

        Private Function GetCustomer(ByVal senderType As Object) As customerns.Client
            Dim cNode As Treenodes.CollectionNodes.CollectionHeaderNode
            Dim CustNode As Treenodes.CustomerNodes.CustNode
            Dim CustSMIDNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim myClient As Intranet.intranet.customerns.Client
            If TypeOf senderType Is Treenodes.CollectionNodes.CollectionHeaderNode Then
                cNode = senderType
                If TypeOf cNode.Parent Is Treenodes.CustomerNodes.CustNode Then
                    CustNode = cNode.Parent
                    myClient = CustNode.client
                End If
                If TypeOf cNode.Parent Is Treenodes.CustomerNodes.CustSMIDNode Then
                    CustSMIDNode = cNode.Parent
                    myClient = Nothing
                End If

            Else
                myClient = Nothing
            End If
            Return myClient
        End Function

        Private Sub MnuRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            'Dim intRet As DialogResult
            'Dim blnClient As Boolean
            'Dim oVessel As Intranet.intranet.customerns.vessel

            Dim SenderType As Object = GetSenderType(sender)
            UpdateSender(SenderType)
        End Sub


        End Class


End Namespace