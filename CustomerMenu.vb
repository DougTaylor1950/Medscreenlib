Imports System.Windows.Forms
Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports Intranet.intranet
Imports Intranet.intranet.Support.cSupport
Imports System.Reflection
Imports MsOutlook = Microsoft.Office.Interop.Outlook

Namespace menus

    Public Class CustomerMenuBase
        Inherits MenuBase


        Friend Overloads Shared Sub AddComment(ByVal objClient As Intranet.intranet.customerns.Client)
            If Not objClient Is Nothing Then
                Dim afrm As New FrmCollComment()
                afrm.FormType = FrmCollComment.FCCFormTypes.customerComment
                afrm.CommentType = MedscreenLib.Constants.GCST_COMM_CollCOOfficerInfo
                afrm.CMNumber = objClient.Identity
                If objClient.ProgrammeManager.Trim.Length > 0 Then
                    afrm.ProgamManager = objClient.SalesPerson
                End If
                If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, objClient.Identity, _
                     afrm.CommentType, afrm.Comment, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
                    
                End If
            End If
        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' Create a comment for a SMID
        ''' </summary>
        ''' <param name="SMID"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><Author>Taylor</Author><created>22-Mar-2011</created></revision></revisionHistory>
        Friend Overloads Shared Sub AddComment(ByVal SMID As String)
            Dim afrm As New FrmCollComment()
            afrm.FormType = FrmCollComment.FCCFormTypes.customerComment
            afrm.CommentType = MedscreenLib.Constants.GCST_COMM_CollCOOfficerInfo
            afrm.CMNumber = SMID
            afrm.ProgamManager = ""
            If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                MedscreenLib.Comment.CreateComment(Comment.CommentOn.SMID, SMID, _
                 afrm.CommentType, afrm.Comment, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            End If
        End Sub

        Public Shared Sub ContextMenuRedraw(ByVal ctxMenu As ContextMenu, Optional ByVal Profile As String = "", Optional ByVal SMID As String = "")
            Dim mnuItem As MenuItem
            For Each mnuItem In ctxMenu.MenuItems
                If TypeOf mnuItem Is TaggedMenuItem Then
                    Dim mnuTitem As TaggedMenuItem = mnuItem
                    Dim Roles As String = mnuTitem.Roles
                    If Roles.Trim.Length > 0 Then ' We have roles 
                        Dim blnInRole As Boolean = False
                        Dim RoleList As String() = Roles.Split(New Char() {","})
                        Dim aRole As String
                        For Each aRole In RoleList
                            blnInRole = MedscreenLib.Glossary.Glossary.UserHasRole(aRole)
                            If blnInRole Then Exit For
                        Next
                        mnuTitem.Visible = blnInRole
                    End If
                    mnuTitem.Enabled = True
                    MenuRedraw(mnuItem, Profile, SMID)
                Else
                    MenuRedraw(mnuItem)
                End If
            Next

        End Sub

        Private Shared Sub MenuRedraw(ByVal amnuItem As MenuItem, Optional ByVal profile As String = "", Optional ByVal smid As String = "")
            Dim mnuItem As MenuItem
            For Each mnuItem In amnuItem.MenuItems
                If TypeOf mnuItem Is TaggedMenuItem Then
                    Dim mnuTitem As TaggedMenuItem = mnuItem
                    Dim Roles As String = mnuTitem.Roles
                    If Roles.Trim.Length > 0 Then ' We have roles 
                        Dim blnInRole As Boolean = False
                        Dim RoleList As String() = Roles.Split(New Char() {","})
                        Dim aRole As String
                        For Each aRole In RoleList
                            blnInRole = MedscreenLib.Glossary.Glossary.UserHasRole(aRole)
                            If blnInRole Then Exit For
                        Next
                        mnuTitem.Visible = blnInRole
                    End If
                    Debug.Print(mnuTitem.Group & " - " & mnuTitem.Text & " - " & mnuTitem.Roles)
                    If mnuTitem.Group = "CHILD" Then
                        Try
                            Dim intCount As Integer = CConnection.PackageStringList("lib_customer2.CountBarcodeAccounts", smid)
                            If intCount <> 1 Then
                                mnuTitem.Enabled = False
                            Else
                                Dim strProfile As String = CConnection.PackageStringList("lib_customer2.FindBarcodeAccount", smid)
                                If strProfile = profile Then
                                    mnuTitem.Enabled = False
                                End If
                            End If
                        Catch ex As Exception

                        End Try
                    ElseIf mnuTitem.Group = "NOCHILD" Then
                        Dim strResp As String = CConnection.PackageStringList("lib_customer2.IsAChild", profile)
                        If strResp = "T" Then
                            mnuTitem.Enabled = True
                        Else
                            mnuTitem.Enabled = False
                        End If
                    End If

                    MenuRedraw(mnuItem)
                Else
                    MenuRedraw(mnuItem)
                End If
            Next

        End Sub

        Public Sub mnuNewCustomer(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim senderType As Object = MenuBase.GetSenderType(sender)
            Dim tmpClient As Intranet.intranet.customerns.Client
            Dim strMessage As String = ""
            Dim Anode As TreeNode

            ' Dim tpNode As Treenodes.CustomerNodes.CustPanelNode
            Dim strCommand As String
            Dim smidFrm As New MedscreenCommonGui.frmSmidProfile()
            Dim aMenuItem As MenuItem

            If TypeOf senderType Is TreeNode Then
                'aMenuItem = sender
                'If aMenuItem.Index = mnuNewCustomerC.Index Then
                Anode = senderType
                If TypeOf Anode Is MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode Then
                    smidFrm.SMID = CType(Anode, MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode).SMID.SMID
                    smidFrm.SmidProfiles = CType(Anode.TreeView, CustomTreeBase).TemplateNode.SMID.Profiles
                End If
                'End 'If
            End If
            If smidFrm.ShowDialog = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If

            Dim CustomerDetails As New frmCustomerDetails()
            'Check to see if reviving an old account
            If smidFrm.Client Is Nothing Then
                'Get next available customer ID number 
                strCommand = MedscreenLib.CConnection.NextSequence("SEQ_CUSTOMER_ID")
                CustomerDetails.NewClient = True
                tmpClient = New Intranet.intranet.customerns.Client(strCommand, False)
                'tmpClient.Identity = strCommand

                'tmpClient.ContractDate = Now.AddYears(1)
                tmpClient.NewCustomer = True
                tmpClient.AccStatus = Constants.GCST_Cust_Status_New
                tmpClient.SMID = smidFrm.SMID
                tmpClient.IndustryCode = smidFrm.IndustryCode
                tmpClient.CompanyName = smidFrm.CompanyProfileName

                'Sort out test panel selected details 
                'Dim objTP As Intranet.intranet.TestPanels.TestPanel
                'Dim strTP As String
                'objTP = Intranet.intranet.Support.cSupport.TestPanels.Item(smidFrm.TestPanel)
                'If objTP Is Nothing Then    'Can't match panel in the list 
                '    strTP = smidFrm.TestPanel
                'Else
                '    strTP = objTP.TestPanelDescription      'Use panel description 
                'End If

                Dim oIndDef As MedscreenLib.Defaults.IndustryDefault = MedscreenLib.Glossary.Glossary.CustomerDefaults.IndustryDefaults.Item(smidFrm.IndustryCode)
                If Not oIndDef Is Nothing Then
                    'If MsgBox("Add chosen test panel - " & strTP & " to customer?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    '    tmpClient.TestPanels.AddTestPanel(smidFrm.TestPanel)
                    '    tmpClient.AccountCollectionType = oIndDef.CollectionType
                    '    If tmpClient.TestPanels.Count = 1 Then
                    '        With tmpClient.TestPanels.Item(0)
                    '            .Alcohol = oIndDef.AlcoholCutoff
                    '            .BreathCutoff = oIndDef.BreathAlcoholCutoff
                    '            .Cannabis = oIndDef.CannabisCutoff
                    '            .CreateTestSchedule()
                    '        End With
                    '    End If
                    'End If

                    'deal with options 
                    Dim optDef As MedscreenLib.Defaults.OptionDefault
                    For Each optDef In oIndDef.OptionDefaults
                        'get option 
                        Dim objOpt As MedscreenLib.Glossary.Options = MedscreenLib.Glossary.Glossary.Options.Item(optDef.OptionId)
                        If Not objOpt Is Nothing Then
                            If objOpt.OptionType = "BOOL" Then
                                tmpClient.Options.AddOption(tmpClient.Identity, optDef.OptionId, (optDef.OptionValue = "TRUE"), "")
                            Else
                                tmpClient.Options.AddOption(tmpClient.Identity, optDef.OptionId, "", optDef.OptionValue)
                            End If
                        End If
                    Next
                Else    'No defaults 
                    'If MsgBox("Add chosen test panel - " & strTP & " to customer?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    '    tmpClient.TestPanels.AddTestPanel(smidFrm.TestPanel)
                    'End If
                End If
            Else
                tmpClient = smidFrm.Client
                tmpClient.AccStatus = Constants.GCST_Cust_Status_Active
                tmpClient.Removed = False

            End If
            tmpClient.Profile = smidFrm.Profile
            tmpClient.IndustryCode = smidFrm.IndustryCode
            CustomerDetails.cClient = tmpClient



            With CustomerDetails


                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    tmpClient = .cClient
                    tmpClient.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    tmpClient.ModifiedOn = Now
                    tmpClient.Update()
                    If tmpClient.AccStatus = Constants.GCST_Cust_Status_New Then Intranet.Intranet.Support.cSupport.CustList.Add(tmpClient)

                    If Len(strMessage) > 0 Then
                        MsgBox(strMessage)
                    End If
                    Dim cNode As Treenodes.CustomerNodes.CustSMIDNode = Nothing
                    'For Each Anode In Me.TreeView1.AlphaNodes.Nodes
                    '    If Anode.Text = tmpClient.Identity.Chars(0) Then
                    '        Exit For
                    '    End If
                    'Next

                    If CType(Anode.TreeView, CustomTreeBase).CheckNodeForSMID(tmpClient, cNode) Then
                        cNode.Refresh()
                        'cNode = New Treenodes.CustomerNodes.CustNode(tmpClient, 0)
                        'Anode.Nodes.Add(cNode)
                        'TreeView1.CustomerNode = cNode
                        'Me.TreeView1.SelectedNode = TreeView1.CustomerNode
                        'TreeView1.CustomerNode.EnsureVisible()
                        'Me.TreeView1.Invalidate()
                        'Me.TreeView1.Select()
                        Dim custNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode = Nothing
                        If CType(Anode.TreeView, CustomTreeBase).CheckNodeForCustomer(tmpClient, custNode) Then
                            Anode.TreeView.SelectedNode = custNode
                            Anode.TreeView.SelectedNode.EnsureVisible()
                            Anode.TreeView.Invalidate()
                            Anode.TreeView.Select()

                        End If

                    End If
                    'Me.mnuAssignTestPanel_Click(Nothing, Nothing)
                Else
                End If
                .NewClient = False
            End With



        End Sub

        Private Function TrimTo10(ByVal inString As String) As String
            If inString.Trim.Length > 10 Then
                Return Mid(inString, 1, 10)
            Else
                Return inString
            End If
        End Function

        Private Function CreateFromTemplate(ByVal newID As String, ByVal Template As String, _
 ByVal Smid As String, ByVal Profile As String, ByVal CompanyName As String) As Intranet.intranet.customerns.Client

            '<Added code modified 15-May-2007 06:16 by CONCATENO\taylor> 'Check on validity of profile
            If Profile.Trim.Length = 0 Then
                MsgBox("No profile entered will terminate!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
                Return Nothing
                Exit Function
            End If
            '</Added code modified 15-May-2007 06:16 by CONCATENO\taylor> 

            Dim ocmd As New OleDb.OleDbCommand()
            Dim intRet As Integer
            Dim retClient As Intranet.intranet.customerns.Client = Nothing
            Try
                ocmd.CommandType = CommandType.StoredProcedure
                ocmd.CommandText = "pckCustomer.CreateCustomerFromTemplate" 'use procedure Clone customer in package pckCustomer
                'Create parameter list 
                Dim idParam As New OleDb.OleDbParameter("id", newID)
                Dim smidParam As New OleDb.OleDbParameter("inSMID", TrimTo10(Smid))
                Dim profileParam As New OleDb.OleDbParameter("inProfile", TrimTo10(Profile))
                Dim cNameParam As New OleDb.OleDbParameter("inCompanyName", CompanyName)
                Dim TemplateParam As New OleDb.OleDbParameter("Template", Template)
                Dim oldidParam As New OleDb.OleDbParameter("ClonedID", "")
                Dim InOpParam As New OleDb.OleDbParameter("InOperator", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
                'Set stored procedure parameters
                ocmd.Parameters.Add(idParam)
                ocmd.Parameters.Add(smidParam)
                ocmd.Parameters.Add(profileParam)
                ocmd.Parameters.Add(cNameParam)
                ocmd.Parameters.Add(TemplateParam)
                ocmd.Parameters.Add(oldidParam)
                ocmd.Parameters.Add(InOpParam)

                ocmd.Connection = CConnection.DbConnection
                If CConnection.ConnOpen Then
                    intRet = ocmd.ExecuteNonQuery()
                End If
                retClient = New Intranet.intranet.customerns.Client(newID, True)
                MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, newID, "NEW", _
                "Account created as " & Smid & "-" & Profile & " from template :" & Template, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
                retClient.Update()

            Catch ex As Exception
                Medscreen.LogError(ex, True, "Cloning")
            Finally
                CConnection.SetConnClosed()
            End Try
            Return retClient
        End Function


        Private Sub Docreate(ByVal custNode As Treenodes.MedTreeNode)
            Try  ' Protecting
                Dim Smidfrm As MedscreenCommonGui.frmSmidProfile = New frmSmidProfile()
                Smidfrm.FormMode = frmSmidProfile.FormModes.Template
                If Not custNode Is Nothing Then
                    If TypeOf custNode Is Treenodes.CustomerNodes.CustNode Then
                        Smidfrm.SMID = CType(custNode, Treenodes.CustomerNodes.CustNode).client.SMID()
                        Smidfrm.CompanyProfileName = CType(custNode, Treenodes.CustomerNodes.CustNode).client.CompanyName()
                    ElseIf TypeOf custNode Is Treenodes.CustomerNodes.CustSMIDNode Then
                        Smidfrm.SMID = CType(custNode, Treenodes.CustomerNodes.CustSMIDNode).SMID.SMID
                    ElseIf TypeOf custNode Is Treenodes.AlphaNode Then
                        Smidfrm.SMID = ""
                        Smidfrm.CompanyProfileName = ""
                    End If
                Else
                    Smidfrm.SMID = ""
                    Smidfrm.CompanyProfileName = ""
                End If
                If Not custNode Is Nothing Then Smidfrm.SmidProfiles = CType(custNode.TreeView, CustomTreeBase).TemplateNode.SMID.Profiles
                If Smidfrm.ShowDialog = DialogResult.OK Then
                    Dim strTemplate As String = Smidfrm.TemplateID
                    Dim strCommand As String
                    Dim tmpClient As Intranet.intranet.customerns.Client
                    strCommand = MedscreenLib.CConnection.NextSequence("SEQ_CUSTOMER_ID")
                    Dim CustomerDetails As New frmCustomerDetails()
                    CustomerDetails.NewClient = True
                    If Not custNode Is Nothing Then
                        If TypeOf custNode Is Treenodes.CustomerNodes.CustNode Then
                            tmpClient = CType(custNode, Treenodes.CustomerNodes.CustNode).client.CreateFromTemplate(strCommand, strTemplate, Smidfrm.SMID, Smidfrm.Profile, Smidfrm.CompanyProfileName)
                        Else
                            tmpClient = CreateFromTemplate(strCommand, strTemplate, Smidfrm.SMID, Smidfrm.Profile, Smidfrm.CompanyProfileName)
                        End If
                    Else
                        tmpClient = CreateFromTemplate(strCommand, strTemplate, Smidfrm.SMID, Smidfrm.Profile, Smidfrm.CompanyProfileName)
                    End If
                    If tmpClient Is Nothing Then 'Failed for some reason
                        MsgBox("Could not create the customer profile", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                        Exit Sub
                    End If 'Clone success check
                    tmpClient.Refresh()
                    tmpClient.NewCustomer = True
                    tmpClient.AccStatus = Constants.GCST_Cust_Status_New
                    tmpClient.SMID = Smidfrm.SMID
                    tmpClient.Profile = Smidfrm.Profile
                    'tmpClient.IndustryCode = Smidfrm.IndustryCode
                    tmpClient.CompanyName = Smidfrm.CompanyProfileName
                    tmpClient.Update()

                    'See about parenting 

                    'Sort out test panel selected details 
                    'Dim objTP As Intranet.intranet.TestPanels.TestPanel
                    ' Dim strTP As String
                    'objTP = Intranet.intranet.Support.cSupport.TestPanels.Item(Smidfrm.TestPanel)
                    'If objTP Is Nothing Then    'Can't match panel in the list 
                    '    strTP = Smidfrm.TestPanel
                    'Else
                    '    strTP = objTP.TestPanelDescription      'Use panel description 
                    'End If 'objtp nothing
                    Dim cNode As Treenodes.CustomerNodes.CustSMIDNode = Nothing
                    Dim chNode As Treenodes.CustomerNodes.CustNode
                    If TypeOf CType(custNode.TreeView, CustomTreeBase).SelectedNode Is Treenodes.CustomerNodes.CustSMIDNode Then
                        cNode = CType(custNode.TreeView, CustomTreeBase).SelectedNode
                    ElseIf TypeOf CType(custNode.TreeView, CustomTreeBase).SelectedNode Is Treenodes.CustomerNodes.CustNode Then
                        chNode = CType(custNode.TreeView, CustomTreeBase).SelectedNode
                        cNode = chNode.Parent
                    ElseIf TypeOf CType(custNode.TreeView, CustomTreeBase).SelectedNode Is Treenodes.AlphaNode Then
                        cNode = New Treenodes.CustomerNodes.CustSMIDNode(New Intranet.intranet.customerns.SMID(Smidfrm.SMID))

                        'See if we can find the right header node 
                        Dim trNode As TreeNode = CType(custNode.TreeView, CustomTreeBase).AlphaNodes.FindHeaderNode(Smidfrm.SMID.Chars(0))
                        If Not trNode Is Nothing Then
                            trNode.Nodes.Add(cNode)
                        Else
                            CType(custNode.TreeView, CustomTreeBase).SelectedNode.Nodes.Add(cNode)
                        End If
                    End If
                    'See if we have a header node
                    If Not cNode Is Nothing Then
                        cNode.Refresh()
                        Dim anode As Treenodes.CustomerNodes.CustNode = cNode.Find(tmpClient.Identity)
                        If Not anode Is Nothing Then
                            CType(custNode.TreeView, CustomTreeBase).SelectedNode = anode

                        End If
                    End If
                    MsgBox("The services, prices, options, data audits, collection defaults and account details have been copied " & vbCrLf & _
                        "Please ensure that all other details are set up correctly.  You should be positioned on the new account ready for editing.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                End If 'Dialog ok 

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "TestPanelForm-Docreate-6562")
            Finally
            End Try
        End Sub

        Friend Sub mnuCreateCustomer_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Try  ' Protecting
                If SenderType Is Nothing OrElse TypeOf SenderType Is Treenodes.AlphaNode Then
                    If MsgBox("Do you want to create a new SMID-Profile using a template ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        Docreate(sender)
                    End If
                    Exit Sub 'Bog off nothing selected 
                End If
                Dim onode As TreeNode = SenderType
                Dim custNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode
                If TypeOf onode Is MedscreenCommonGui.Treenodes.CustomerNodes.CustNode Then 'Do we have a customer profile node?
                    custNode = onode
                    If MsgBox("Do you want to create a new Profile for this SMID " & custNode.client.SMID & " using a template ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        Docreate(custNode)
                    End If 'Am cloning
                ElseIf TypeOf onode Is MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode Then                 'Do we have a customer profile node?
                    Dim custsmidNode As Treenodes.CustomerNodes.CustSMIDNode = onode
                    If MsgBox("Do you want to create a new Profile for this SMID " & custsmidNode.SMID.SMID & " using a template ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        Docreate(custsmidNode)
                    End If 'Am cloning

                Else
                    Exit Sub
                End If
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "TestPanelForm-mnuCreateCustomer_click-6625")
            Finally
            End Try

        End Sub



        Friend Sub mnuAddSite(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim anode As TreeNode = sender
            If TypeOf anode Is Treenodes.CustomerNodes.CustSMIDNode Then
                Dim cNode As Treenodes.CustomerNodes.CustSMIDNode = anode
                Dim objName As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "New Site name", "Fixed site name", "")
                If objName Is Nothing Then Exit Sub
                Dim aPort As Intranet.intranet.jobs.Port = Intranet.FixedSiteBookings.cCentres.CreateCentre(CStr(objName))
                aPort.Owner = cNode.SMID.SMID
                aPort.OwnerType = "SMID"
                aPort.Update()
                If Not cNode.SitesNode Is Nothing Then
                    Dim sNode As Treenodes.LocationNodes.FixedSiteNode = New Treenodes.LocationNodes.FixedSiteNode(aPort)
                    cNode.SitesNode.Nodes.Add(sNode)
                End If

            End If
        End Sub


    End Class

    Public Class AlphaMenu
        Inherits CustomerMenuBase

        Private Shared myInstance As AlphaMenu

        Private Sub New()

        End Sub

        Public Shared Function GetAlphaMenu() As AlphaMenu
            If myInstance Is Nothing Then
                myInstance = New AlphaMenu()

            End If
            Return myInstance
        End Function

        Private Sub mnuNewCustomer_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            MyBase.mnuCreateCustomer_click(SenderType, e)
            MyBase.UpdateSender(SenderType)

        End Sub
    End Class

    Public Class CustomerMenu
        Inherits CustomerMenuBase
        Private Shared myInstance As CustomerMenu

        Private Sub New()

        End Sub

        Public Shared Function GetCustomerMenu() As CustomerMenu
            If myInstance Is Nothing Then
                myInstance = New CustomerMenu()

            End If
            Return myInstance
        End Function


        Private Sub Report_handler_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Debug.WriteLine(sender.ToString)
            Dim menuItem As TaggedMenuItem
            Dim objSmidNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode
            Dim objCustNode As Treenodes.CustomerNodes.CustNode
            If TypeOf sender Is TaggedMenuItem Then
                menuItem = sender
            End If
            Dim mnuItem As New MedscreenLib.CRMenuItem(menuItem.MenuTag)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim blnGroup As Boolean = False
            Dim LocalClient As Intranet.intranet.customerns.Client      'Local variable
            If TypeOf SenderType Is MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode Then
                blnGroup = True
                objSmidNode = SenderType
                LocalClient = CType(objSmidNode.Nodes.Item(0), MedscreenCommonGui.Treenodes.CustomerNodes.CustNode).client

            Else
                objCustNode = SenderType
                LocalClient = objCustNode.client
            End If
            'Check we have a valid client 
            If LocalClient Is Nothing Then Exit Sub
            Dim MyAssembly As [Assembly]
            'Dim SampleAssembly As [Assembly]
            'dim Object As Intranet.intranet.jobs.clsJobs

            If mnuItem.MenuType = "RPT" Then
                Try
                    Dim strParameterList As String = "METHOD=" & mnuItem.MenuIdentity & " REPORTTYPE=CRYSTAL EMAIL=" & MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                    If blnGroup Then
                        strParameterList += " CUSTOMERID=" & LocalClient.SMID
                    Else
                        strParameterList += " CUSTOMERID=" & LocalClient.Identity
                    End If
                    mnuItem.FillFormula(strParameterList, objCustNode.client.SMID)
                    Dim strOutput As String = ""
                    Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.Constants.SendMethod))
                    Dim obj As Object = Medscreen.GetParameter(GetType(MedscreenLib.Constants.SendMethod), "Test")
                    If Not obj Is Nothing Then
                        strOutput = CStr(obj)
                    End If

                    Dim myBackReport As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    myBackReport.CommandString = strParameterList
                    myBackReport.Status = "W"
                    myBackReport.OutPutMethod = strOutput
                    myBackReport.DoUpdate()
                    MsgBox("Your report will be sent to a background report server and will be emailed to you when complete", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)


                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Crystal Report")
                End Try

            Else
                Try
                    Dim strCommand As String = "METHOD=" & mnuItem.MenuPath & " CUSTOMERID=" & LocalClient.Identity & " IDENTITY=" & mnuItem.MenuIdentity & " GROUP=" & blnGroup.ToString
                    'add report type info 
                    If menuItem.MenuType = "REPORTHTML" Then
                        strCommand += " REPORTTYPE=HTML"
                    ElseIf menuItem.MenuType = "REPORTCRYSTAL" Then
                        strCommand += " REPORTTYPE=CRYSTAL"
                    End If
                    Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.CurrentSMUser 'Find who is logged in 
                    Dim strParam As String
                    Dim Email As String = ""
                    Dim strEmail As String = ""
                    If Not up Is Nothing Then
                        strEmail = up.Email
                    End If
                    If Email.Trim.Length = 0 Then Email = strEmail
                    If Email.Trim.Length = 0 Then
                        strParam = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email Address", "Email Address", strEmail)
                    Else
                        strParam = Email
                    End If
                    strCommand += " EMAIL=" & strEmail
                    Dim objDate As Object
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Start Date", Now.AddMonths(-3))
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " STARTDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "End Date", "End Date", Now)
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " ENDDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    Dim objRep As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    objRep.CommandString = strCommand
                    objRep.DoUpdate()
                    objRep.DoUpdate()
                    MsgBox("Your report will be processed by a report server and emailed to you when it is complete.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    'Dim strRunCommand As String = "C:\Program Files\Medscreen\ReportService\ReportService.exe"
                    'Dim intStart As Integer = Shell(strRunCommand & " " & strCommand, AppWinStyle.MinimizedNoFocus)

                    'Dim repHandler As New Intranet.ReportHandler(mnuItem.MenuPath, LocalClient.Identity, blnGroup, mnuItem.MenuIdentity)
                    'MsgBox("Your report will be produced in background and emailed to you", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                    'repHandler.InvokeMethod()
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Running as assewmbly")
                End Try
            End If





        End Sub

        Public Shared Sub doSetChild(ByVal sender As Object, Optional ByVal aNode As TreeNode = Nothing)
            Dim SenderType As Object = Nothing
            If aNode Is Nothing Then
                SenderType = MenuBase.GetSenderType(sender)
            Else
                SenderType = aNode
            End If
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If

            Dim intBarcodes As Integer = CConnection.PackageStringList("lib_customer2.CountBarcodeAccounts", custNode.client.SMID)
            If intBarcodes = 0 Then
                MsgBox("There are no barcode accounts for this SMID", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            ElseIf intBarcodes > 1 Then
                MsgBox("There are more than one barcode accounts for this SMID, contact IT with the SMID so the problem can be fixed.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            Dim strProfile As String = CConnection.PackageStringList("lib_customer2.FindBarcodeAccount", custNode.client.SMID)
            If strProfile IsNot Nothing AndAlso strProfile.Length > 0 Then
                'check that we are not recursing
                If custNode.client.Identity = strProfile Then
                    MsgBox("The barcode account can't be made a child of itself.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                    Exit Sub
                End If
            Else
                MsgBox("There is a problem identifying the Barcode account for this SMID, ask IT to help.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            'if we get here we can proceed
            custNode.client.Parent = strProfile
            custNode.client.Update()
            MenuBase.UpdateSender(sender)
            'MyBase.UpdateSender(sender)
        End Sub
        Private Sub mnuSetChild_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            doSetChild(sender)
            ' MenuBase.UpdateSender(sender)
        End Sub


        Public Shared Sub doUnSetChild(ByVal sender As Object, Optional ByVal aNode As TreeNode = Nothing)
            Dim SenderType As Object = Nothing
            If aNode Is Nothing Then
                SenderType = MenuBase.GetSenderType(sender)
            Else
                SenderType = aNode
            End If
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If



            'if we get here we can proceed
            custNode.client.Parent = ""
            custNode.client.Update()
            MenuBase.UpdateSender(sender)
            'MyBase.UpdateSender(sender)
        End Sub
        Private Sub mnuUnSetChild_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            doUnSetChild(sender)
            'MenuBase.UpdateSender(sender)
        End Sub

        Private Sub mnuCloneCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            Try  ' Protecting
                If MsgBox("Do you want to clone this customer " & custNode.client.ClientID & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    Dim Smidfrm As MedscreenCommonGui.frmSmidProfile = New frmSmidProfile()
                    Smidfrm.FormMode = frmSmidProfile.FormModes.Clone
                    Smidfrm.SMID = custNode.client.SMID
                    Smidfrm.CompanyProfileName = custNode.client.CompanyName
                    If Smidfrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                        'Get next available customer ID number 
                        Dim strCommand As String
                        Dim tmpClient As Intranet.intranet.customerns.Client
                        strCommand = MedscreenLib.CConnection.NextSequence("SEQ_CUSTOMER_ID")
                        Dim CustomerDetails As New frmCustomerDetails()
                        CustomerDetails.NewClient = True
                        'Call PL/SQL block to do actual cloning 
                        tmpClient = custNode.client.clone(strCommand, Smidfrm.SMID, Smidfrm.Profile, Smidfrm.CompanyProfileName)
                        If tmpClient Is Nothing Then 'Failed for some reason
                            MsgBox("Could not clone the customer profile", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                            Exit Sub
                        End If 'Clone success check
                        tmpClient.NewCustomer = True
                        tmpClient.AccStatus = Constants.GCST_Cust_Status_New
                        tmpClient.SMID = Smidfrm.SMID
                        tmpClient.Profile = Smidfrm.Profile
                        'tmpClient.IndustryCode = Smidfrm.IndustryCode
                        tmpClient.CompanyName = Smidfrm.CompanyProfileName
                        tmpClient.Update()
                        'See about parenting 
                        If MsgBox("Make this account a child of " & custNode.client.ClientID & "?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            tmpClient.Parent = custNode.client.Identity
                            tmpClient.Update()
                        End If 'Child check

                        

                        If MsgBox("Duplicate Customers test panel?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                            Dim ocmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("pckCustomer.CloneTestPanel", CConnection.DbConnection)
                            Try
                                If CConnection.ConnOpen Then            'See if we can open the connection 
                                    ocmd.CommandType = CommandType.StoredProcedure
                                    ocmd.Parameters.Add(New OleDb.OleDbParameter("id", tmpClient.Identity))
                                    ocmd.Parameters.Add(New OleDb.OleDbParameter("ClonedID", custNode.client.Identity))
                                    ocmd.Parameters.Add(New OleDb.OleDbParameter("InOperator", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity))
                                    ocmd.ExecuteNonQuery()
                                End If 'CConnection.ConnOpen
                            Catch ex As Exception
                            Finally
                                CConnection.SetConnClosed()

                            End Try
                        End If 'MsgBox duplicate
                        'End If
                        Dim cNode As Treenodes.CustomerNodes.CustSMIDNode = Nothing
                        If CType(custNode.TreeView, CustomTreeBase).CheckNodeForSMID(tmpClient, cNode) Then
                            cNode.Refresh()
                            'Dim custNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode
                            If CType(custNode.TreeView, CustomTreeBase).CheckNodeForCustomer(tmpClient, custNode) Then
                                custNode.TreeView.SelectedNode = custNode
                                custNode.TreeView.SelectedNode.EnsureVisible()
                                custNode.TreeView.Invalidate()
                                custNode.TreeView.Select()
                            Else
                                Dim tcustNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode = New MedscreenCommonGui.Treenodes.CustomerNodes.CustNode(tmpClient, 0)
                                custNode.Parent.Nodes.Add(tcustNode)
                                custNode.TreeView.SelectedNode = tcustNode
                                custNode.TreeView.SelectedNode.EnsureVisible()
                                custNode.TreeView.Invalidate()
                                custNode.TreeView.Select()
                            End If 'TreeView1.CheckNodeForCustomer(tmpClient, custNode)

                        End If 'Me.TreeView1.CheckNodeForSMID(tmpClient, cNode)


                    End If 'Dialog ok 
                    MsgBox("The services, prices, options, data audits, collection defaults and account details have been copied " & vbCrLf & _
                        "Please ensure that all other details are set up correctly.  You should be positioned on the new account ready for editing.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                End If 'Am cloning
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "TestPanelForm-mnuCloneCustomer_Click-6646")
            Finally
            End Try

        End Sub

        Private Sub mnuShowInvoices_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            If custNode Is Nothing Then Exit Sub
            Dim strCID As String = custNode.client.Identity
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader

            Dim aValue As Object = MedscreenLib.Medscreen.GetParameter(MyTypes.typDate, "Invoices From", , Now.AddDays(-180))
            If aValue Is Nothing Then Exit Sub
            Dim datFrom As Date = CDate(aValue)

            Try
                oCmd.Connection = CConnection.DbConnection
                oCmd.CommandText = "Select invoice_number,transaction_type from invoices where invoice_customer = '" & _
                strCID & "' and invoiced_on > to_date('" & datFrom.ToString("ddMMyyyy") & "','DDMMYYYY')"
                MedscreenLib.CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                Dim strID As String
                Dim strTranstype As String
                Dim oInv As Intranet.intranet.jobs.Invoice
                'Me.TreeView1.BeginUpdate()
                Dim oColl As New Collection()
                While oRead.Read
                    strID = oRead.GetString(0)
                    strTranstype = oRead.GetValue(1)
                    oColl.Add(strID & "-" & strTranstype)

                End While
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
                Dim i As Integer
                For i = 1 To oColl.Count

                    Dim strIDFull As String = oColl.Item(i)
                    Dim intPos As Integer = strIDFull.IndexOf("-")
                    strID = Mid(strIDFull, 1, intPos)
                    Dim strTrans As String = Mid(strIDFull, intPos + 2)
                    oInv = Intranet.Intranet.Support.cSupport.Invoices.Item(strID, strTrans)
                    If Not oInv Is Nothing Then
                        CType(custNode.TreeView, CustomTreeBase).PositionInvoice(oInv)
                    End If
                Next
            Catch ex As Exception
                LogError(ex)

            Finally
                'Me.TreeView1.EndUpdate()

            End Try

        End Sub

        Private Sub mnuLast45Invoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            '
            If custNode Is Nothing Then Exit Sub
            Dim strCID As String = custNode.client.Identity
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader

            Try
                oCmd.Connection = CConnection.DbConnection
                oCmd.CommandText = "Select invoice_number from invoices where invoice_customer = '" & _
                strCID & "' and invoiced_on > to_date('" & Now.AddDays(-45).ToString("ddMMyyyy") & "','DDMMYYYY')"
                MedscreenLib.CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                Dim strID As String
                Dim oInv As Intranet.intranet.jobs.Invoice
                'Me.TreeView1.BeginUpdate()
                Dim oColl As New Collection()
                While oRead.Read
                    strID = oRead.GetString(0)
                    oColl.Add(strID)

                End While
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
                Dim i As Integer
                For i = 1 To oColl.Count
                    strID = oColl.Item(i)
                    oInv = Intranet.Intranet.Support.cSupport.Invoices.Item(strID)
                    If Not oInv Is Nothing Then
                        CType(custNode.TreeView, CustomTreeBase).PositionInvoice(oInv)
                    End If
                Next
            Catch ex As Exception
                LogError(ex)

            Finally
                'Me.TreeView1.EndUpdate()

            End Try
        End Sub

        Private Sub mnuShowCollections_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            If custNode.CollectionHeaderNode Is Nothing Then
                Dim myColl As Intranet.intranet.jobs.CMCollection = New Intranet.intranet.jobs.CMCollection("Customer_id = '" & custNode.client.Identity & "' and coll_date > sysdate -60")
                custNode.CollectionHeaderNode = New Treenodes.CollectionNodes.CollectionHeaderNode(myColl, custNode.client.Identity)
            End If
            custNode.CollectionHeaderNode.EnsureVisible()
            custNode.TreeView.SelectedNode = custNode.CollectionHeaderNode
        End Sub

        Public Shared Sub DoIndustry(ByVal sender As Object, Optional ByVal aNode As TreeNode = Nothing)
            Dim SenderType As Object = Nothing
            If aNode Is Nothing Then
                SenderType = MenuBase.GetSenderType(sender)
            Else
                SenderType = aNode
            End If
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            If custNode.client Is Nothing Then Exit Sub
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_EDIT_CUSTOMERID") Or MedscreenLib.Glossary.Glossary.UserHasRole("MED_CREATE_CUSTOMER") Then
                Dim objRet As Object
                Dim ClientId As String = custNode.client.Identity
                Dim objCustStatus As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("INDUSTRY")
                objCustStatus.Load("PHRASE_ID")
                objRet = Medscreen.GetParameter(objCustStatus, "Customer Industry", custNode.client.IndustryCode)
                If objRet Is Nothing Then  'User cancelled
                    Exit Sub
                End If
                custNode.client.IndustryCode = CStr(objRet)
                custNode.client.DoUpdate()
                Dim CustSmid As Treenodes.CustomerNodes.CustSMIDNode = custNode.Parent
                CustSmid.Refresh()
                custNode = CType(CustSmid.TreeView, CustomTreeBase).PositionClient(ClientId)
                If Not custNode Is Nothing Then
                    custNode.TreeView.SelectedNode = custNode
                Else
                    CustSmid.TreeView.SelectedNode = CustSmid
                End If


            Else
                MsgBox("Sorry you need to be in either the edit or create customer roles", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        End Sub
        Private Sub mnuIndustry_click(ByVal sender As Object, ByVal e As System.EventArgs)

            DoIndustry(sender)


        End Sub

        Public Shared Sub DoChangeStatus(ByVal sender As Object, Optional ByVal Anode As TreeNode = Nothing)
            Dim SenderType As Object = Nothing
            If Anode Is Nothing Then
                SenderType = MenuBase.GetSenderType(sender)
            Else
                SenderType = Anode
            End If
            Dim custNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                custNode = SenderType
            Else
                Exit Sub
            End If
            If custNode.client Is Nothing Then Exit Sub
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_EDIT_CUSTOMERID") Or MedscreenLib.Glossary.Glossary.UserHasRole("MED_CREATE_CUSTOMER") Then
                Dim objRet As Object
                Dim ClientId As String = custNode.client.Identity
                Dim objCustStatus As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection("CUST_STAT")
                objCustStatus.Load()
                objRet = Medscreen.GetParameter(objCustStatus, "Customer Status")
                If objRet Is Nothing Then  'User cancelled
                    Exit Sub
                End If
                custNode.client.AccStatus = CStr(objRet)
                custNode.client.DoUpdate()
                Dim CustSmid As Treenodes.CustomerNodes.CustSMIDNode = custNode.Parent
                CustSmid.Refresh()
                custNode = CType(CustSmid.TreeView, CustomTreeBase).PositionClient(ClientId)
                If Not custNode Is Nothing Then
                    custNode.TreeView.SelectedNode = custNode
                Else
                    CustSmid.TreeView.SelectedNode = CustSmid
                End If


            Else
                MsgBox("Sorry you need to be in either the edit or create customer roles", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        End Sub
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' change the status of the customer
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/04/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuChangeStatus_click(ByVal sender As Object, ByVal e As System.EventArgs)
            DoChangeStatus(sender)
        End Sub

        Private Sub mnuCustCommitted_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Try
                Dim SenderType As Object = MenuBase.GetSenderType(sender)
                Dim custNode As Treenodes.CustomerNodes.CustNode
                If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                    custNode = SenderType
                Else
                    Exit Sub
                End If
                If custNode.client Is Nothing Then Exit Sub

                If custNode.client.Identity = "MEDSCREEN" Then
                    If Not MedscreenLib.Glossary.Glossary.SampleMangerAssignments.UserHasRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "PERSONNEL") Then
                        MessageBox.Show("This account can only be viewed by personnel", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        Exit Sub
                    End If
                End If
                Dim startDate As Date = DateSerial(Now.Year, 1, 1)
                startDate = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Start date selection", startDate)
                Dim Enddate As Date = DateSerial(Now.Year, 12, 31)
                Enddate = Medscreen.GetParameter(MyTypes.typDate, "End Date", "End date selection", Enddate)

                custNode.TreeView.Cursor = Cursors.WaitCursor
                Dim aSamp As New Intranet.intranet.sample.oSampleCollection()
                aSamp.LoadCustomerSamples(custNode.client.Identity, startDate, Enddate, False)
                Dim strXml As String = aSamp.ToXML(, False, True)
                Dim args As New Xml.Xsl.XsltArgumentList()
                args.AddParam("startdate", "", startDate.ToString("dd-MMM-yyyy"))
                args.AddParam("enddate", "", Enddate.ToString("dd-MMM-yyyy"))
                args.AddParam("customer", "", custNode.client.CompanyName)

                Dim strHTMl As String = Medscreen.ResolveStyleSheet(strXml, "samplelist.xsl", 0, args, )
                Medscreen.BlatEmail("Committed samples", strHTMl, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)
                custNode.TreeView.Cursor = Cursors.Default

                'asamp.LoadCustomerSamples(
            Catch ex As Exception
            End Try
        End Sub



        Private Sub mnuCustDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            'Dim strMessage As String
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim CustNode As Treenodes.CustomerNodes.CustNode
            Dim aTreeView As TreeView
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                CustNode = SenderType
                aTreeView = CustNode.TreeView
            Else
                Exit Sub
            End If
            If CustNode.client Is Nothing Then Exit Sub
            aTreeView.Cursor = Cursors.WaitCursor              'Let the world know this might take some time

            CustNode.client.Refresh()                                 'Make sure the data is up todate
            Dim CustomerDetails As New frmCustomerDetails()
            CustomerDetails.cClient = CustNode.client
            Dim strOldProfile As String = CustNode.client.ClientID       'Save current client profile
            Dim strXML1 As String = CustNode.client.ToXML(True, True, True)
            Dim strXML2 As String
            LogTiming("Opening Form")
            Try  ' Protecting editing customer
                If CustomerDetails.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    CustNode.client = CustomerDetails.cClient
                    strXML2 = CustNode.client.ToXML(True, True, True)

                    'Write out the pre and post edit XML values 
                    Me.WriteOutClientXML(CustNode.client, strXML1, strXML2)
                    If strOldProfile <> CustNode.client.ClientID Then
                        LogAction("Smid Profile changed" & strOldProfile, MedscreenLib.Constants.GCST_X_DRIVE & "\error\cctool\error.txt")
                    End If
                    Try
                        CustNode.client.Update()
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                    End Try
                Else
                    'Client.ExtraInfo.Rollback()
                End If
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "TestPanelForm-mnuCustDetails_Click-2907")
            Finally
                CustNode.client.Refresh()
                Dim anode As TreeNode = CustNode
                Dim ASmidNode As Treenodes.CustomerNodes.CustSMIDNode = CustNode.Parent
                ASmidNode.Refresh()
                Dim aCustnode As Treenodes.CustomerNodes.CustNode = ASmidNode.FindCustomerNode(CustNode.client.Identity)
                aTreeView.SelectedNode = aCustnode
                aTreeView.Cursor = Windows.Forms.Cursors.Default               'Let the world know this might take some time
            End Try

        End Sub



        Private Sub mnuCustCollOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType

                Dim frmCOP As New frmCollDefaults()

                If cNode.client.CollectionInfo Is Nothing Then
                    cNode.client.CollectionInfo = New Intranet.intranet.customerns.CustColl(cNode.client.Identity)
                End If

                frmCOP.CollectionInfo = cNode.client.CollectionInfo
                If frmCOP.ShowDialog = Windows.Forms.DialogResult.OK Then
                    cNode.client.CollectionInfo = frmCOP.CollectionInfo
                    cNode.client.CollectionInfo.Update()
                    cNode.client.Refresh()
                    cNode.NodeHasChanged()
                End If
            End If
            MyBase.UpdateSender(sender)
        End Sub


        Private Sub mnuCustPositives_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                Dim strCmd As String = "http://andrew/WebGCMSLate/WebFormPositives.aspx?Action=POSITIVE&startdate=" & Now.AddYears(-1).ToString("dd-MMM-yyyy") & _
                               "&enddate=" & Now.ToString("dd-MMM-yyyy") & "&ID=" & cNode.client.Identity

                Try
                    Dim strHTML As String = ""
                    Try
                        strHTML = Intranet.Intranet.Support.cSupport.DoNavigate(strCmd)
                    Catch ex As Exception
                    End Try
                    Medscreen.BlatEmail("Positives for " & cNode.client.SMIDProfile, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)
                Catch ex As Exception
                End Try
            End If
        End Sub


        Private WithEvents CustOptions As frmCustomerOptions
        Private LocalClient As Intranet.intranet.customerns.Client
        Private Sub mnuCustServices_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            CustOptions = New frmCustomerOptions()
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType

                Try
                    With CustOptions
                        .FormType = frmCustomerOptions.COFormType.Services
                        .Client = cNode.client
                        LocalClient = cNode.client
                        .ShowDialog()
                    End With
                    cNode.client.Refresh()
                    cNode.NodeHasChanged()

                Catch ex As Exception
                End Try
            End If
        End Sub

        Private Sub mnuCardiffSamples_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                Dim vesselnode As MedscreenCommonGui.Treenodes.VesselNodes.VesselNode
                Dim objClient As Intranet.intranet.customerns.Client = Nothing
                objClient = cNode.client

                If Not objClient Is Nothing Then
                    Dim myService As New MedscreenLib.ts011.CardiffService()
                    Dim FromDate As Object = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting Date of Report", Today.AddDays(-30))
                    If FromDate Is Nothing Then FromDate = Today.AddDays(-30)
                    Dim strXML As String = myService.GetReceivedSamplesBySMIDProfile(FromDate, objClient.SMIDProfile, "")
                    Dim strStylesheet As String = MedscreenCommonGUIConfig.NodeStyleSheets("SMIDCardiffView")
                    Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, strStylesheet, 0)
                    Medscreen.ShowHtml(strHTML)

                End If
            End If
            MyBase.UpdateSender(SenderType)
        End Sub
        Private Sub mnuAddSched_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                Dim NewSched As MedscreenLib.ReportSchedule = MedscreenLib.ReportSchedules.CreateSchedule
                NewSched.ScheduleType = MedscreenLib.ReportSchedule.GCST_CUSTOMER_REPORT
                NewSched.CustomerID = cNode.client.Identity
                NewSched.Update()
                Dim frm As New frmReportSchedule()

                frm.Schedule = NewSched
                If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    NewSched = frm.Schedule
                    NewSched.Update()
                End If
                cNode.client.ReportSchedules.Add(NewSched)
                UpdateSender(sender)
            End If
        End Sub



        Private Sub mnuNewCustomer_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            MyBase.mnuCreateCustomer_click(sender, e)
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuNewCustComment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                'Dim vesselnode As MedscreenCommonGui.treenodes.vesselnodes.vesselnode
                Dim objClient As Intranet.intranet.customerns.Client = Nothing
                objClient = cNode.client

                MyBase.AddComment(objClient)

            End If
        End Sub


        Private Sub mnuSMIDCollectionReport_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tmpVar As Object
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim anode As TreeNode
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim ProfileNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is TreeNode Then
                anode = SenderType
            End If
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
            ElseIf TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                ProfileNode = SenderType
            Else
                Exit Sub
            End If
            'Get start date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting point of report", DateSerial(Today.AddMonths(-3).Year, Today.AddMonths(-3).Month, 1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datStart As Date = CDate(tmpVar)
            'Get end date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "End Date", "Starting point of report", datStart.AddMonths(4).AddDays(-1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datEnd As Date = CDate(tmpVar).AddDays(1)
            anode.TreeView.Cursor = Cursors.WaitCursor
            'now get the list of ids of customers managed
            Dim aCustColl As customerns.SMClientCollection = New customerns.SMClientCollection()
            aCustColl.Add(ProfileNode.client)


            Dim strHTML As String = aCustColl.CollectionReport(datStart, datEnd, "SMIDCollReport.xsl")
            anode.TreeView.Cursor = Cursors.Default
            Medscreen.BlatEmail("Collection report for " & ProfileNode.client.SMIDProfile, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)

        End Sub

        Private Sub mnuFleetReport_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim SMID As String
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim CustNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
                SMID = SmidNode.SMID.SMID
            ElseIf TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                CustNode = SenderType
                SMID = CustNode.client.SMID
            Else
                Exit Sub
            End If
            Dim tmpVar As Object
            'Get start date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting point of report", DateSerial(Today.AddMonths(-3).Year, Today.AddMonths(-3).Month, 1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datStart As Date = CDate(tmpVar)
            'Get end date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "End Date", "Starting point of report", datStart.AddMonths(4).AddDays(-1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datEnd As Date = CDate(tmpVar).AddDays(1)
            CType(SenderType, TreeNode).TreeView.Cursor = Cursors.WaitCursor
            'now get the list of ids of customers managed
            'Set up XML
            Dim strXML As String = "<GROUP>"
            Dim strTail As String = "</CollectionList>"
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                If SmidNode.SMID Is Nothing Then Exit Sub
                Dim oChild As Intranet.intranet.customerns.SMIDProfile
                Dim i As Integer
                For i = 0 To SmidNode.SMID.Profiles.Count - 1
                    oChild = SmidNode.SMID.Profiles.Item(i)
                    If Not oChild.Client.Removed Then        'Establish there are vessels and the account is active
                        ' If oChild.Client.Vessels.Count > 0 Then        'Make sure they are really there 
                        Dim objParamid As OleDb.OleDbParameter = New OleDb.OleDbParameter("ID", oChild.Client.Identity.ToUpper)
                        objParamid.DbType = DbType.String
                        objParamid.Size = 10
                        Dim objParamStart As OleDb.OleDbParameter = New OleDb.OleDbParameter("Coll_date", datStart)
                        objParamStart.DbType = DbType.DateTime
                        Dim objParamEnd As OleDb.OleDbParameter = New OleDb.OleDbParameter("Coll_date2", datEnd)
                        objParamEnd.DbType = DbType.DateTime
                        Dim ParamCollection As New Collection()
                        ParamCollection.Add(objParamid)
                        ParamCollection.Add(objParamStart)
                        ParamCollection.Add(objParamEnd)
                        'Get a collection of collections in Date Range
                        Dim objColl As New jobs.CMCollection("customer_id = ?  and coll_date >  ? " & _
                             " and coll_date <= ? ", ParamCollection)

                        'objColl.Client = objClient          ' populate the customer property 
                        objColl.Client = oChild.Client
                        strXML += objColl.ToXML(False, True, True, False, , True)
                        '  End If
                    End If
                Next
                strXML += "<HEADER>"
                strXML += "<StartDate>" & datStart.ToString("dd-MMM-yy") & "</StartDate>"
                strXML += "<EndDate>" & datEnd.AddDays(-1).ToString("dd-MMM-yy") & "</EndDate>"
                strXML += "</HEADER>"

                strXML += "</GROUP>"
                Dim strFileName As String = Medscreen.XMLDirectory & "custcoll.xml"
                Dim s As IO.StreamWriter = New IO.StreamWriter(strFileName)

                'Dim strXml As String = s.ReadToEnd()
                s.WriteLine(strXML)
                s.Flush()
                s.Close()
                Dim strHTML As String = ResolveStyleSheet(strXML, "custnorbulkgp.xsl", 0, , Medscreen.XSLTransformsDirectory)
                Medscreen.BlatEmail("SMID Collection report for " & SmidNode.SMID.SMID, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)

            End If
            CType(SenderType, TreeNode).TreeView.Cursor = Cursors.Default
        End Sub


        Public Shared Sub domnuTestPanels(ByVal sender As Object)
            Try
                Dim SenderType As Object = MenuBase.GetSenderType(sender)
                If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                    Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                    If cNode Is Nothing Then Exit Sub 'Bog off nothing selected
                    Dim cstTestPanel As New MedscreenCommonGui.frmTestPanel()
                    cstTestPanel.Client = cNode.client
                    If cstTestPanel.ShowDialog = DialogResult.OK Then
                        cNode.client.Refresh()
                        cNode.NodeHasChanged()
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Private Sub mnuTestPanels_click(ByVal sender As Object, ByVal e As System.EventArgs)
            domnuTestPanels(sender)
        End Sub

        Public Shared Sub doPasswordLetter(ByVal sender As Object, Optional ByVal aNode As TreeNode = Nothing)
            Dim SenderType As Object = Nothing
            If aNode Is Nothing Then
                SenderType = MenuBase.GetSenderType(sender)
            Else
                SenderType = aNode
            End If
            Try
                If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                    Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                    If cNode Is Nothing Then Exit Sub 'Bog off nothing selected
                    SendPasswordLetter(cNode.client, cNode.client.ReportingInfo.EmailPassword)
                    cNode.client.Refresh()
                    cNode.NodeHasChanged()

                End If
            Catch ex As Exception
            End Try
        End Sub
        Private Sub mnuPasswordLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            doPasswordLetter(sender)
        End Sub

        Public Shared Function SendPasswordLetter(ByVal myClient As Intranet.intranet.customerns.Client, ByVal EmailPassword As String) As Boolean
            'Send pasword letter to customer
            Dim frmPass As New frmPrintPassword()
            frmPass.Client = myClient
            Dim myTextFile As Reporter.clsTextFile
            Dim objReporter As Reporter.Report = New Reporter.Report()

            'objReporter.PurgeOldFiles()
            'Need to set email body file from config settings
            'Dim mySeetings As System.Configuration.ConfigurationSettings
            objReporter.ReportSettings.EmailBodyFile = MedscreenLib.Constants.PasswordBodyFile
            If frmPass.ShowDialog = DialogResult.OK Then
                myTextFile = myClient.ReportingInfo.PasswordLetter(frmPass.ConfirmMethod, EmailPassword, Now, frmPass.txtOutput.Text)
                Dim strPath As String = "\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates|\\corp.concateno.com\medscreen\common\Lab Programs\DBReports\Templates\Scanned Documents"


                Dim PathArray As String() = strPath.Split(New Char() {"|"})
                objReporter.ReportSettings.SetTemplatePaths(PathArray)                  'Set up Remplate locations

                objReporter.SetToCustomerDirectory()
                objReporter.ReportSettings.DocumentPath += myClient.SMID & "\"
                If Not IO.Directory.Exists(objReporter.ReportSettings.DocumentPath) Then
                    IO.Directory.CreateDirectory(objReporter.ReportSettings.DocumentPath)
                End If
                Dim strDocName As String = frmPass.ConfirmMethod.ToString & "-" & myClient.Profile & "-" & Now.ToString("yyyyMMddHHmm") & ".doc"
                Dim objCReport As Reporter.clsWordReport = objReporter.GetDatafile(myTextFile.ExportStringArray, strDocName)
                objReporter.WordApplication.Visible = (frmPass.OutputAs = frmPrintPassword.OutputOption.Printer)
                objReporter.WordApplication.ActiveDocument.Save()
                Intranet.Support.AdditionalDocumentColl.CreateCustomerDocument(myClient.Identity, _
                 objReporter.ReportSettings.DocumentPath & strDocName, "PASSWORD")
                If frmPass.OutputAs = frmPrintPassword.OutputOption.Email Then
                    Dim objRecip As Reporter.clsRecipient = New Reporter.clsRecipient()
                    objRecip.address = myClient.EmailResults.Replace(";", ",")
                    objCReport.SendByEmail(objRecip)
                End If

                'Okay deal with send log
                Dim objSendLog As MedscreenLib.SendLogInfo = MedscreenLib.SendLogCollection.CreateSendLog()
                objSendLog.Filename = strDocName
                objSendLog.ReferenceID = myClient.Identity
                objSendLog.ReferenceTo = "PASSWORD"
                objSendLog.SentOn = Now
                objSendLog.SendingProgram = "CCTOOL"
                objSendLog.SendMethod = "DISK"
                objSendLog.Status = "C"
                objSendLog.Update()
                'Now we need to set up a comment.
                MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, myClient.Identity, "DOC", frmPass.ConfirmMethod.ToString & " created", _
    MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)

                'Me.Client.CustomerComments.CreateComment(frmPass.ConfirmMethod.ToString & " created", "DOC", jobs.CollectionCommentList.CommentType.Customer)

            End If
        End Function

        Private Sub mnuDocView_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 
                Dim frmView As New frmViewDocs
                frmView.CustomerId = cNode.client.Identity
                frmView.ShowDialog()
            End If
        End Sub

        Private Sub mnuNewVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 

                Dim ovessel As Intranet.intranet.customerns.Vessel

                ovessel = Intranet.intranet.Support.cSupport.AddNewVessel(cNode.client.Identity)
                If Not ovessel Is Nothing Then
                    cNode.client.Vessels.Clear()
                    Dim objParamid As OleDb.OleDbParameter = New OleDb.OleDbParameter("ID", cNode.client.Identity.Trim)
                    objParamid.DbType = DbType.String
                    objParamid.Size = 10
                    Dim objParamct As OleDb.OleDbParameter = New OleDb.OleDbParameter("TYPE", "VESSEL")
                    objParamct.DbType = DbType.String
                    objParamct.Size = 10
                    Dim Oparamlist As New Collection()
                    Oparamlist.Add(objParamid)
                    Oparamlist.Add(objParamct)
                    Oparamlist.Add(CConnection.StringParameter("removeflag", "F", 1))
                    objParamct.Value = "CUSTSITE"
                    cNode.client.Vessels.Load("customer_id = ?  and vesseltype  <>  ? and removeflag = ?", , Oparamlist)
                    If Not cNode.VesselHeaderNode Is Nothing Then

                        cNode.VesselHeaderNode.Refresh()        'If we have a vessel header node then refresh it 

                        CType(cNode.TreeView, CustomTreeBase).CurrentVessel = ovessel
                        cNode.TreeView.SelectedNode = cNode.VesselHeaderNode.FindVesselNode(ovessel.VesselID)
                    End If

                End If
            End If
        End Sub

        Private Sub mnuNewCustomerSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 
                If cNode.client Is Nothing Then Exit Sub

                Dim oSite As Intranet.intranet.Address.CustomerSite

                oSite = cNode.client.Sites.Create(cNode.client.Identity)
                If Not oSite Is Nothing Then

                    oSite.VesselType = MedscreenLib.Constants.GCSTCustSite
                    Dim edFrm As New frmVessel()
                    edFrm.Vessel = oSite

                    If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                        CType(cNode.TreeView, CustomTreeBase).CurrentVessel = edFrm.Vessel
                    End If
                End If
            End If
        End Sub



        Private Sub mnuNewCustIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 
                Dim afrm As New FrmCollComment()
                afrm.FormType = FrmCollComment.FCCFormTypes.CustomerIssue
                afrm.CommentType = "PROB"
                afrm.CMNumber = cNode.client.Identity
                If cNode.client.ProgrammeManager.Trim.Length > 0 Then
                    afrm.ProgamManager = cNode.client.SalesPerson
                End If
                If afrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'Dim aComm As Intranet.intranet.jobs.CollectionComment = _
                    '   custNode.client.CustomerComments.CreateComment(afrm.Comment, afrm.CommentType, jobs.CollectionCommentList.CommentType.Customer)
                    MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, cNode.client.Identity, _
     afrm.CommentType, afrm.Comment, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)

                    'If Not aComm Is Nothing Then
                    '    aComm.ActionBy = afrm.ActionBy
                    '    aComm.Update()

                    'End If
                End If

            End If
        End Sub


        Private Sub mnuProgManContact_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 
                If cNode.client Is Nothing Then Exit Sub
                If cNode.client.ProgrammeManagement Is Nothing Then Exit Sub
                If cNode.client.ProgrammeManagement.ContactSendEmail.Trim.Length = 0 Then
                    Exit Sub
                End If

                Dim objOutlook As MsOutlook.Application

                Try
                    objOutlook = New MsOutlook.Application()
                    Dim objMessage As MsOutlook.MailItem = objOutlook.CreateItem(MsOutlook.OlItemType.olMailItem)
                    objMessage.Subject = "Re : Your Collection Programme"
                    objMessage.Recipients.Add(cNode.client.ProgrammeManagement.ContactSendEmail)
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


        Private Sub mnuCustProgMan_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                If cNode Is Nothing Then Exit Sub 'Bog off nothing selected 
                If cNode.client Is Nothing Then Exit Sub

                Dim objFrm As New MedscreenCommonGui.frmProgrammeManage()
                objFrm.Program = cNode.client.ProgrammeManagement      'Provide the data 
                objFrm.Company = cNode.client.CompanyName              'Let the yobs know how it is for 
                If objFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    cNode.client.ProgrammeManagement = objFrm.Program  'Get the changes 
                    cNode.client.ProgrammeManagement.DoUpdate()        'Update the changes
                End If

                'Me.TreeView1.CustomerNode.client.Refresh()          'Update the GUI
                'Dim anode As TreeNode = Me.TreeView1.CustomerNode
                'Me.TreeView1.SelectedNode = anode.Parent
                'Me.TreeView1.SelectedNode = anode
                UpdateSender(sender)
            End If
        End Sub


        Private Sub mnuCustOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim CustOptions As frmCustomerOptions = New frmCustomerOptions()
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType

                Try
                    With CustOptions
                        .Client = cNode.client
                        .ShowDialog()
                    End With
                    cNode.client.Refresh()
                    cNode.NodeHasChanged()
                    'HtmlEditor1.LoadDocument(Me.TreeView1.CustomerNode.client.ToHTML("customer1.xslt", MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\transforms\xsl\"))
                Catch ex As Exception
                End Try
            End If
        End Sub

        Private Sub mnuEditPrices_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType
                Dim frmep As New FrmPricing()
                frmep.Client = cNode.client
                If frmep.ShowDialog = Windows.Forms.DialogResult.OK Then
                    frmep.Client.Update()
                    cNode.Refresh()
                    cNode.NodeHasChanged()
                    'Me.HtmlEditor1.LoadDocument(cNode.ToHTML)
                End If
            End If
        End Sub

        Private Sub mnuEditCustAddress_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim frmAddr As New MedscreenCommonGui.frmAddress()
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Dim cNode As Treenodes.CustomerNodes.CustNode = SenderType

                frmAddr.Client = cNode.client
                If frmAddr.ShowDialog = DialogResult.OK Then
                    cNode.NodeHasChanged()
                End If
            End If
        End Sub


        Private Sub WriteOutClientXML(ByVal Client As Intranet.intranet.customerns.Client, ByVal strXML1 As String, ByVal strXML2 As String)
            Try

                Dim iof As System.IO.StreamWriter
                iof = New System.IO.StreamWriter(MedscreenLib.Constants.GCST_X_DRIVE & "\xmltemp\" & Client.Identity & Now.ToString("yyyyMMddHHmmss") & "A.xml")
                iof.WriteLine(strXML1)
                iof.Flush()
                iof.Close()
                iof = New System.IO.StreamWriter(MedscreenLib.Constants.GCST_X_DRIVE & "\xmltemp\" & Client.Identity & Now.ToString("yyyyMMddHHmmss") & "b.xml")
                iof.WriteLine(strXML2)
                iof.Flush()
                iof.Close()
            Catch ex As System.Exception
                ' Handle exception here
                MedscreenLib.Medscreen.LogError(ex, , "Exception log message")
            Finally
                ' Finally code goes here
            End Try

        End Sub

        Private Sub CustOptions_NotifyServiceChange(ByVal Service As String, ByVal Change As String) Handles CustOptions.NotifyServiceChange
            CommonForms.NotifyServiceChange(LocalClient, Service, Change)

        End Sub
    End Class

    Public Class SMIDMenu
        Inherits CustomerMenuBase
        Private Shared myInstance As SMIDMenu
        Private Sub New()

        End Sub

        Public Shared Function GetSMIDMenu() As SMIDMenu
            If myInstance Is Nothing Then
                myInstance = New SMIDMenu()

            End If
            Return myInstance
        End Function


        Private Sub mnuNewCustomer_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            MyBase.mnuCreateCustomer_click(sender, e)
            MyBase.UpdateSender(SenderType)
        End Sub

        Private Sub mnuAddSite_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            MyBase.mnuAddSite(SenderType, e)
            MyBase.UpdateSender(SenderType)

        End Sub

        Private Sub mnuSetRef_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim anode As TreeNode
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim ProfileNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is TreeNode Then
                anode = SenderType
            End If
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
            ElseIf TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                Exit Sub
            Else
                Exit Sub
            End If
            Dim intRet As Integer = 0
            Dim tmpVar As Object = Medscreen.GetParameter(MyTypes.typString, "Customer Reference", "New Customer Reference for all profiles")
            If tmpVar Is Nothing Then Exit Sub
            'now can update all accounts in one hit
            Dim strCommand As String = "update customer set contract_num = ? where smid = ?"
            Dim ocmd As New OleDb.OleDbCommand(strCommand, CConnection.DbConnection)
            Try
                ocmd.Parameters.Add(CConnection.StringParameter("CustRef", CStr(tmpVar), 25))
                ocmd.Parameters.Add(CConnection.StringParameter("SMID", SmidNode.SMID.SMID, 10))
                If CConnection.ConnOpen Then
                    intRet = ocmd.ExecuteNonQuery()
                End If
            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()
            End Try
            If intRet = 0 Then
                MsgBox("No profiles were changed", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            Else
                MsgBox(CStr(intRet) & " Profiles were updated with the new reference", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            End If
        End Sub

        Private Sub mnuNewCustComment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode = SenderType

                CustomerMenuBase.AddComment(SmidNode.SMID.SMID)
            End If
        End Sub


        Private Sub mnuSMIDCollectionReport_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tmpVar As Object
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim anode As TreeNode
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim ProfileNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is TreeNode Then
                anode = SenderType
            End If
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
            ElseIf TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                ProfileNode = SenderType
            Else
                Exit Sub
            End If
            'Get start date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting point of report", DateSerial(Today.AddMonths(-3).Year, Today.AddMonths(-3).Month, 1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datStart As Date = CDate(tmpVar)
            'Get end date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "End Date", "Starting point of report", datStart.AddMonths(4).AddDays(-1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datEnd As Date = CDate(tmpVar).AddDays(1)
            anode.TreeView.Cursor = Cursors.WaitCursor
            'now get the list of ids of customers managed
            Dim aCustColl As customerns.SMClientCollection = Intranet.intranet.Support.cSupport.CustList.Customers( _
                SmidNode.SMID.SMID, customerns.SMClientCollection.CustIndexType.SMID)
            Dim strHTML As String = aCustColl.CollectionReport(datStart, datEnd, "SMIDCollReport.xsl")
            anode.TreeView.Cursor = Cursors.Default
            Medscreen.BlatEmail("SMID Collection report for " & SmidNode.SMID.SMID, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)

        End Sub

        Private Sub mnuFleetReport_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim SMID As String
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim CustNode As Treenodes.CustomerNodes.CustNode
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
                SMID = SmidNode.SMID.SMID
            ElseIf TypeOf SenderType Is Treenodes.CustomerNodes.CustNode Then
                CustNode = SenderType
                SMID = CustNode.client.SMID
            Else
                Exit Sub
            End If
            Dim tmpVar As Object
            'Get start date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting point of report", DateSerial(Today.AddMonths(-3).Year, Today.AddMonths(-3).Month, 1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datStart As Date = CDate(tmpVar)
            'Get end date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "End Date", "Starting point of report", datStart.AddMonths(4).AddDays(-1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datEnd As Date = CDate(tmpVar).AddDays(1)
            CType(SenderType, TreeNode).TreeView.Cursor = Cursors.WaitCursor
            'now get the list of ids of customers managed
            'Set up XML
            Dim strXML As String = "<GROUP>"
            Dim strTail As String = "</CollectionList>"
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                If SmidNode.SMID Is Nothing Then Exit Sub
                Dim oChild As Intranet.intranet.customerns.SMIDProfile
                Dim i As Integer
                For i = 0 To SmidNode.SMID.Profiles.Count - 1
                    oChild = SmidNode.SMID.Profiles.Item(i)
                    If Not oChild.Client.Removed Then        'Establish there are vessels and the account is active
                        ' If oChild.Client.Vessels.Count > 0 Then        'Make sure they are really there 
                        Dim objParamid As OleDb.OleDbParameter = New OleDb.OleDbParameter("ID", oChild.Client.Identity.ToUpper)
                        objParamid.DbType = DbType.String
                        objParamid.Size = 10
                        Dim objParamStart As OleDb.OleDbParameter = New OleDb.OleDbParameter("Coll_date", datStart)
                        objParamStart.DbType = DbType.DateTime
                        Dim objParamEnd As OleDb.OleDbParameter = New OleDb.OleDbParameter("Coll_date2", datEnd)
                        objParamEnd.DbType = DbType.DateTime
                        Dim ParamCollection As New Collection()
                        ParamCollection.Add(objParamid)
                        ParamCollection.Add(objParamStart)
                        ParamCollection.Add(objParamEnd)
                        'Get a collection of collections in Date Range
                        Dim objColl As New jobs.CMCollection("customer_id = ?  and coll_date >  ? " & _
                             " and coll_date <= ? ", ParamCollection)

                        'objColl.Client = objClient          ' populate the customer property 
                        objColl.Client = oChild.Client
                        strXML += objColl.ToXML(False, True, True, False, , True)
                        '  End If
                    End If
                Next
                strXML += "<HEADER>"
                strXML += "<StartDate>" & datStart.ToString("dd-MMM-yy") & "</StartDate>"
                strXML += "<EndDate>" & datEnd.AddDays(-1).ToString("dd-MMM-yy") & "</EndDate>"
                strXML += "</HEADER>"

                strXML += "</GROUP>"
                Dim strFileName As String = Medscreen.XMLDirectory & SmidNode.SMID.SMID & " -custcoll.xml"
                Dim s As IO.StreamWriter = New IO.StreamWriter(strFileName)

                'Dim strXml As String = s.ReadToEnd()
                s.WriteLine(strXML)
                s.Flush()
                s.Close()
                Dim strHTML As String = ResolveStyleSheet(strXML, "custnorbulkgp.xsl", 0, , Medscreen.XSLTransformsDirectory)
                Medscreen.BlatEmail("SMID Collection report for " & SmidNode.SMID.SMID, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)

            End If
            CType(SenderType, TreeNode).TreeView.Cursor = Cursors.Default
        End Sub


        Private Sub Report_handler_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Debug.WriteLine(sender.ToString)
            Dim menuItem As TaggedMenuItem = Nothing
            Dim objSmidNode As MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode
            If TypeOf sender Is TaggedMenuItem Then
                menuItem = sender
            End If
            Dim mnuItem As New MedscreenLib.CRMenuItem(menuItem.MenuTag)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim blnGroup As Boolean = False
            Dim LocalClient As Intranet.intranet.customerns.Client = Nothing                            'Local variable
            If TypeOf SenderType Is MedscreenCommonGui.Treenodes.CustomerNodes.CustSMIDNode Then
                blnGroup = True
                objSmidNode = SenderType
                If objSmidNode.Nodes.Count > 0 Then                                                     'Check we have nodes
                    Dim objCustNode As Treenodes.CustomerNodes.CustNode = objSmidNode.FirstClientNode   'fIND FIRST CLIENT NODE
                    If objCustNode IsNot Nothing Then                                                   ' If we have found one get the client
                        LocalClient = objCustNode.client
                    End If
                End If
            Else
                LocalClient = CType(SenderType, Treenodes.CustomerNodes.CustNode).client
            End If
            'Check we have a valid client 
            If LocalClient Is Nothing Then Exit Sub
            Dim MyAssembly As [Assembly]
            'Dim SampleAssembly As [Assembly]
            'dim Object As Intranet.intranet.jobs.clsJobs

            If mnuItem.MenuType = "RPT" Then
                Try
                    Dim strParameterList As String = "METHOD=" & mnuItem.MenuIdentity & " REPORTTYPE=CRYSTAL EMAIL=" & MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                    If blnGroup Then
                        strParameterList += " CUSTOMERID=" & LocalClient.SMID
                    Else
                        strParameterList += " CUSTOMERID=" & LocalClient.Identity
                    End If
                    mnuItem.FillFormula(strParameterList, objSmidNode.SMID.SMID)
                    Dim strOutput As String = ""
                    Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.Constants.SendMethod))
                    Dim obj As Object = Medscreen.GetParameter(GetType(MedscreenLib.Constants.SendMethod), "Test")
                    If Not obj Is Nothing Then
                        strOutput = CStr(obj)
                    End If

                    Dim myBackReport As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    myBackReport.CommandString = strParameterList
                    myBackReport.Status = "W"
                    myBackReport.OutPutMethod = strOutput
                    myBackReport.DoUpdate()
                    MsgBox("Your report will be sent to a background report server and will be emailed to you when complete", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)

                    'It is s crystal report. 
                    'Dim cr As CrystalDecisions.CrystalReports.Engine.ReportDocument
                    'cr = LogInCrystal(mnuItem)
                    'If Not cr Is Nothing Then
                    '    'fill formulae if present

                    '    mnuItem.Formulae.Load(CConnection.DbConnection)
                    '    FillFormula(mnuItem, cr)
                    '    Crf = New MedscreenLib.CRFormulaItem()
                    '    Crf.Formula = "CustomerID"

                    '    If blnGroup Then
                    '        FillFormulaDet(cr, Crf, LocalClient.SMID)
                    '    Else
                    '        FillFormulaDet(cr, Crf, LocalClient.Identity)
                    '    End If
                    '    '<Added code modified 22-Apr-2007 15:58 by CONCATENO\taylor> 
                    '    If mnuItem.Formulae.Count = 1 Then
                    '        Crf = mnuItem.Formulae.Item(0)
                    '    End If

                    '    '</Added code modified 22-Apr-2007 15:58 by CONCATENO\taylor> 
                    '    Crf.ParentReport = cr
                    '    Crf.ReportName = mnuItem.MenuText
                    '    Crf.EmailAddress = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
                    '    MsgBox("Your report will be produced in the background and emailed to you." & vbCrLf _
                    '       & "You will be able to continue working while this happens.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)

                    '    Dim thrMyThread As New System.Threading.Thread( _
                    '        AddressOf Crf.DoReport)   'objCReport.SendByPDF(myRecipient))
                    '    thrMyThread.Start()
                    '    System.Threading.Thread.Sleep(1000)



                    'Else
                    '    MsgBox("Could not create Crystal Report", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)

                    'End If
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Crystal Report")
                End Try

            Else
                Try
                    Dim strCommand As String = "METHOD=" & mnuItem.MenuPath & " CUSTOMERID=" & LocalClient.Identity & " IDENTITY=" & mnuItem.MenuIdentity & " GROUP=" & blnGroup.ToString
                    'add report type info 
                    If menuItem.MenuType = "REPORTHTML" Then
                        strCommand += " REPORTTYPE=HTML"
                    ElseIf menuItem.MenuType = "REPORTCRYSTAL" Then
                        strCommand += " REPORTTYPE=CRYSTAL"
                    End If
                    Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.CurrentSMUser 'Find who is logged in 
                    Dim strParam As String
                    Dim Email As String = ""
                    Dim strEmail As String = ""
                    If Not up Is Nothing Then
                        strEmail = up.Email
                    End If
                    If Email.Trim.Length = 0 Then Email = strEmail
                    If Email.Trim.Length = 0 Then
                        strParam = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email Address", "Email Address", strEmail)
                    Else
                        strParam = Email
                    End If
                    strCommand += " EMAIL=" & strEmail
                    Dim objDate As Object
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Start Date", Now.AddMonths(-3))
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " STARTDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "End Date", "End Date", Now)
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " ENDDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    Dim objRep As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    objRep.CommandString = strCommand
                    objRep.DoUpdate()
                    objRep.DoUpdate()
                    MsgBox("Your report will be processed by a report server and emailed to you when it is complete.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    Dim strRunCommand As String = "C:\Program Files\Medscreen\ReportService\ReportService.exe"
                    Dim intStart As Integer = Shell(strRunCommand & " " & strCommand, AppWinStyle.MinimizedNoFocus)

                    'Dim repHandler As New Intranet.ReportHandler(mnuItem.MenuPath, LocalClient.Identity, blnGroup, mnuItem.MenuIdentity)
                    'MsgBox("Your report will be produced in background and emailed to you", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                    'repHandler.InvokeMethod()
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Running as assewmbly")
                End Try
            End If





        End Sub


        Private Sub mnuCustGroupPositives_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim SmidNode As Treenodes.CustomerNodes.CustSMIDNode
            Dim tmpVar As Object
            'Get start date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Starting point of report", DateSerial(Today.AddMonths(-3).Year, Today.AddMonths(-3).Month, 1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datStart As Date = CDate(tmpVar)
            'Get end date
            tmpVar = Medscreen.GetParameter(MyTypes.typDate, "End Date", "Starting point of report", datStart.AddMonths(4).AddDays(-1))
            If tmpVar Is Nothing Then Exit Sub
            Dim datEnd As Date = CDate(tmpVar).AddDays(1)
            If TypeOf SenderType Is Treenodes.CustomerNodes.CustSMIDNode Then
                SmidNode = SenderType
            Else
                Exit Sub
            End If
            SmidNode.TreeView.Cursor = Cursors.WaitCursor
            Dim strCmd As String = "http://andrew/WebGCMSLate/WebFormPositives.aspx?Action=POSITIVE&startdate=" & datStart.ToString("dd-MMM-yyyy") & _
              "&enddate=" & datEnd.ToString("dd-MMM-yyyy") & "&ID=" & SmidNode.SMID.SMID & "&OPTION=SMID"
            Dim strHTML As String = ""
            Try
                strHTML = Intranet.intranet.Support.cSupport.DoNavigate(strCmd)
            Catch ex As Exception
            End Try
            Medscreen.BlatEmail("Positives for " & SmidNode.SMID.SMID, strHTML, MedscreenLib.Glossary.Glossary.CurrentSMUser.Email)
            SmidNode.TreeView.Cursor = Cursors.Default

        End Sub
    End Class

    Public Class VesselMenu
        Inherits CustomerMenuBase
        Private Shared myInstance As VesselMenu
        Private Sub New()

        End Sub

        Public Shared Function GetVesselMenu() As VesselMenu
            If myInstance Is Nothing Then
                myInstance = New VesselMenu()

            End If
            Return myInstance
        End Function

        Private Sub mnuChangeFleet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselNode = SenderType
                If aVessNode.Vessel Is Nothing Then Exit Sub
                'Get the new customer
                Dim objCust As Intranet.intranet.customerns.Client
                Dim objSelCust As New MedscreenCommonGui.frmCustomerSelection2()
                If objSelCust.ShowDialog = DialogResult.OK Then
                    objCust = objSelCust.Client
                    aVessNode.Vessel.CustomerID = objCust.Identity
                    'aVessNode.Client = objCust
                    aVessNode.Vessel.Update()
                    aVessNode.Refresh()
                End If
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Allow a new vessel to be added 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/12/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuNewVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselNode = SenderType
                If aVessNode.Client Is Nothing Then Exit Sub

                Dim ovessel As Intranet.intranet.customerns.Vessel

                ovessel = Intranet.intranet.Support.cSupport.AddNewVessel(aVessNode.Client.Identity)
                If Not ovessel Is Nothing Then
                    aVessNode.Client.Vessels.Clear()
                    Dim objParamid As OleDb.OleDbParameter = New OleDb.OleDbParameter("ID", aVessNode.Client.Identity.Trim)
                    objParamid.DbType = DbType.String
                    objParamid.Size = 10
                    Dim objParamct As OleDb.OleDbParameter = New OleDb.OleDbParameter("TYPE", "VESSEL")
                    objParamct.DbType = DbType.String
                    objParamct.Size = 10
                    Dim Oparamlist As New Collection()
                    Oparamlist.Add(objParamid)
                    Oparamlist.Add(objParamct)
                    Oparamlist.Add(CConnection.StringParameter("removeflag", "F", 1))
                    objParamct.Value = "CUSTSITE"
                    aVessNode.Client.Vessels.Load("customer_id = ?  and vesseltype  <>  ? and removeflag = ?", , Oparamlist)
                    Dim aVesssNode As Treenodes.VesselNodes.VesselNode = New Treenodes.VesselNodes.VesselNode(ovessel)
                    aVessNode.Parent.Nodes.Add(aVesssNode)
                    aVessNode.TreeView.SelectedNode = aVesssNode
                    'Dim cnode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode = _
                    'Me.aVessNode.CustomerNode
                    'Me.TreeView1.SelectedNode = cnode
                    'If Not cnode Is Nothing Then                    'Get current customer node 
                    '    If Not cnode.VesselHeaderNode Is Nothing Then

                    '        cnode.VesselHeaderNode.Refresh()        'If we have a vessel header node then refresh it 
                    '        Me.TreeView1.CurrentVessel = ovessel
                    '        Me.TreeView1.SelectedNode = cnode.VesselHeaderNode.FindVesselNode(ovessel.VesselID)
                    '    End If
                    'End If

                End If
            End If
        End Sub


        Private Sub Report_handler_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Debug.WriteLine(sender.ToString)
            Dim menuItem As TaggedMenuItem
            If TypeOf sender Is TaggedMenuItem Then
                menuItem = sender
            End If
            Dim mnuItem As New MedscreenLib.CRMenuItem(menuItem.MenuTag)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            Dim blnGroup As Boolean = False
            If Not TypeOf SenderType Is MedscreenCommonGui.Treenodes.VesselNodes.VesselNode Then
                Exit Sub
            End If

            Dim vNode As Treenodes.VesselNodes.VesselNode = SenderType
            'Check we have a valid client 
            Dim MyAssembly As [Assembly]
            'Dim SampleAssembly As [Assembly]
            'dim Object As Intranet.intranet.jobs.clsJobs

            If mnuItem.MenuType = "RPT" Then
                Try
                    Dim strParameterList As String = "METHOD=" & mnuItem.MenuIdentity & " REPORTTYPE=CRYSTAL EMAIL=" & MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                    strParameterList += " VESSELID=" & vNode.Vessel.VesselID
                    'mnuItem.FillFormula(strParameterList, objCustNode.client.SMID)
                    Dim strOutput As String = ""
                    Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.Constants.SendMethod))
                    Dim obj As Object = Medscreen.GetParameter(GetType(MedscreenLib.Constants.SendMethod), "Test")
                    If Not obj Is Nothing Then
                        strOutput = CStr(obj)
                    End If

                    Dim myBackReport As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    myBackReport.CommandString = strParameterList
                    myBackReport.Status = "W"
                    myBackReport.OutPutMethod = strOutput
                    myBackReport.DoUpdate()
                    MsgBox("Your report will be sent to a background report server and will be emailed to you when complete", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)


                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Crystal Report")
                End Try

            Else
                Try
                    Dim strCommand As String = "METHOD=" & mnuItem.MenuPath & " VESSELID=" & vNode.Vessel.VesselID & " IDENTITY=" & mnuItem.MenuIdentity & " GROUP=" & blnGroup.ToString
                    'add report type info 
                    If menuItem.MenuType = "REPORTHTML" Then
                        strCommand += " REPORTTYPE=HTML"
                    ElseIf menuItem.MenuType = "REPORTCRYSTAL" Then
                        strCommand += " REPORTTYPE=CRYSTAL"
                    End If
                    Dim up As MedscreenLib.personnel.UserPerson = MedscreenLib.Glossary.Glossary.CurrentSMUser 'Find who is logged in 
                    Dim strParam As String
                    Dim Email As String = ""
                    Dim strEmail As String = ""
                    If Not up Is Nothing Then
                        strEmail = up.Email
                    End If
                    If Email.Trim.Length = 0 Then Email = strEmail
                    If Email.Trim.Length = 0 Then
                        strParam = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email Address", "Email Address", strEmail)
                    Else
                        strParam = Email
                    End If
                    strCommand += " EMAIL=" & strEmail
                    Dim objDate As Object
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "Start Date", "Start Date", Now.AddMonths(-3))
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " STARTDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    objDate = Medscreen.GetParameter(MyTypes.typDate, "End Date", "End Date", Now)
                    If objDate Is Nothing Then Exit Sub
                    strCommand += " ENDDATE=" & CDate(objDate).ToString("dd-MMM-yyyy")
                    Dim objRep As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                    objRep.CommandString = strCommand
                    objRep.DoUpdate()
                    objRep.DoUpdate()
                    MsgBox("Your report will be processed by a report server and emailed to you when it is complete.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    'Dim strRunCommand As String = "C:\Program Files\Medscreen\ReportService\ReportService.exe"
                    'Dim intStart As Integer = Shell(strRunCommand & " " & strCommand, AppWinStyle.MinimizedNoFocus)

                    'Dim repHandler As New Intranet.ReportHandler(mnuItem.MenuPath, LocalClient.Identity, blnGroup, mnuItem.MenuIdentity)
                    'MsgBox("Your report will be produced in background and emailed to you", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                    'repHandler.InvokeMethod()
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Running as assewmbly")
                End Try
            End If





        End Sub


        Friend Sub mnuEditVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'TODO add vessel editing capability
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselNode = SenderType
                If aVessNode.Vessel Is Nothing Then Exit Sub


                Dim edFrm As New frmVessel()
                edFrm.Vessel = aVessNode.Vessel

                If edFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'aVessNode.Vessel = edFrm.Vessel
                    aVessNode.Vessel.Update()

                    aVessNode.Refresh()
                    'Me.TreeView1.SelectedNode = vNode.Parent
                    aVessNode.EnsureVisible()
                End If

            End If
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 'Routine to deal with booking a vessel 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuBookVessel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            'Routine to deal with booking a vessel 

            Dim objCm As Intranet.intranet.jobs.CMJob
            Dim blnret As Boolean
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselNode = SenderType

                Dim strCustID As String = aVessNode.Vessel.CustomerID
                Dim strVesselID As String = aVessNode.Vessel.VesselID
                objCm = CommonForms.BookRoutine(blnret, strCustID, strVesselID)
            End If
        End Sub

        Private Sub mnuVesselReport_click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselNode = SenderType
                If aVessNode.Vessel Is Nothing Then Exit Sub
                Dim strHTML As String = aVessNode.Vessel.ToHTML("Vesselrep.xslt", MedscreenLib.Medscreen.XSLTransformsDirectory, customerns.Vessel.XMLOptions.WithCollections)
                Dim strEmail As String = ""
                strEmail = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email

                Dim vntRet As Object = GetParameter(MyTypes.typString, "Email Address(s)", , strEmail)
                If vntRet Is Nothing Then Exit Sub
                strEmail = vntRet
                Dim strFilename As String = IO.Path.GetTempFileName
                Dim iof As New IO.StreamWriter(strFilename)
                iof.WriteLine(strHTML)
                iof.Flush()
                iof.Close()
                Dim strSubject As String = "Vessel report for " & aVessNode.Vessel.VesselName
                'Medscreen.BlatEmail(strSubject, strHTML, strEmail)
                Dim strCommand As String = "C:\Program Files\Medscreen\HTMLToPDF\HTMLToPDF.exe /CONVERT " & strFilename & "," & strFilename & ".pdf," & strEmail & "," & strSubject

                Shell(strCommand, AppWinStyle.MaximizedFocus, True, 5000)

            End If
        End Sub
    End Class

    Public Class VesselHeaderMenu
        Inherits CustomerMenuBase
        Private Shared myInstance As VesselHeaderMenu
        Private Sub New()

        End Sub

        Public Shared Function GetVesselMenu() As VesselHeaderMenu
            If myInstance Is Nothing Then
                myInstance = New VesselHeaderMenu()

            End If
            Return myInstance
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Allow a new vessel to be added 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/12/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuNewVessel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselHeaderNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselHeaderNode = SenderType
                If aVessNode.Client Is Nothing Then Exit Sub

                Dim ovessel As Intranet.intranet.customerns.Vessel

                ovessel = Intranet.intranet.Support.cSupport.AddNewVessel(aVessNode.Client.Identity)
                If Not ovessel Is Nothing Then
                    aVessNode.Client.Vessels.Clear()
                    Dim objParamid As OleDb.OleDbParameter = New OleDb.OleDbParameter("ID", aVessNode.Client.Identity.Trim)
                    objParamid.DbType = DbType.String
                    objParamid.Size = 10
                    Dim objParamct As OleDb.OleDbParameter = New OleDb.OleDbParameter("TYPE", "VESSEL")
                    objParamct.DbType = DbType.String
                    objParamct.Size = 10
                    Dim Oparamlist As New Collection()
                    Oparamlist.Add(objParamid)
                    Oparamlist.Add(objParamct)
                    Oparamlist.Add(CConnection.StringParameter("removeflag", "F", 1))
                    objParamct.Value = "CUSTSITE"
                    aVessNode.Client.Vessels.Load("customer_id = ?  and vesseltype  <>  ? and removeflag = ?", , Oparamlist)
                    Dim aVesssNode As Treenodes.VesselNodes.VesselNode = New Treenodes.VesselNodes.VesselNode(ovessel)
                    aVessNode.Nodes.Add(aVesssNode)
                    aVessNode.TreeView.SelectedNode = aVesssNode
                    'Dim cnode As MedscreenCommonGui.Treenodes.CustomerNodes.CustNode = _
                    'Me.aVessNode.CustomerNode
                    'Me.TreeView1.SelectedNode = cnode
                    'If Not cnode Is Nothing Then                    'Get current customer node 
                    '    If Not cnode.VesselHeaderNode Is Nothing Then

                    '        cnode.VesselHeaderNode.Refresh()        'If we have a vessel header node then refresh it 
                    '        Me.TreeView1.CurrentVessel = ovessel
                    '        Me.TreeView1.SelectedNode = cnode.VesselHeaderNode.FindVesselNode(ovessel.VesselID)
                    '    End If
                    'End If

                End If
            End If
        End Sub

        Private Sub mnuVesselReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselHeaderNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselHeaderNode = SenderType
                If aVessNode.Client Is Nothing Then Exit Sub

                Dim AccInt As MedscreenLib.AccountsInterface = New MedscreenLib.AccountsInterface()

                With AccInt
                    .RequestCode = "RVDI"
                    .Reference = aVessNode.Client.Identity
                    .SendAddress = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
                    .Status = Constants.GCST_IFace_StatusRequest
                    .Update()

                    Dim chRet As Char
                    Dim ErrMess As String = ""

                    chRet = .Wait(ErrMess)
                    While InStr(Constants.GCST_IFace_StatusCreated & Constants.GCST_IFace_StatusFailed, chRet) = 0
                        chRet = .Wait(ErrMess)
                    End While
                    If ErrMess.Length <> 0 Then
                        MsgBox("report for " & .Reference & " " & ErrMess)
                    End If
                End With
            End If
        End Sub

        Private Sub mnuCustomerVesselReport2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            'Need to pop up dialog boxes
            Dim strCust As String = ""
            Dim vntRet As Object
            Dim datStart As Date
            Dim EndStart As Date
            Dim strEmail As String = ""
            Dim blnGroup As Boolean = False
            Dim SenderType As Object = MenuBase.GetSenderType(sender)
            If TypeOf SenderType Is Treenodes.VesselNodes.VesselHeaderNode Then
                Dim aVessNode As Treenodes.VesselNodes.VesselHeaderNode = SenderType
                strEmail = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                If aVessNode._type = Treenodes.CollectionNodes.CollectionHeaderNode.HeaderType.Profile Then
                    strCust = aVessNode.Client.Identity
                    If strCust = "" Then Exit Sub

                    vntRet = GetParameter(MyTypes.typDate, "Start Date")
                    If vntRet Is Nothing Then Exit Sub
                    datStart = vntRet
                    vntRet = GetParameter(MyTypes.typDate, "End Date")
                    If vntRet Is Nothing Then Exit Sub
                    EndStart = vntRet
                    vntRet = GetParameter(MyTypes.typString, "Email Address(s)", , strEmail)
                    If vntRet Is Nothing Then Exit Sub
                    strEmail = vntRet

                    Dim ActionString As String = "http://andrew/WebGCMSLate/WebformPositives.aspx?Action=CUSTVESSCOLL&id=" & strCust.ToUpper & _
                        "&startdate=" & datStart.ToString("dd-MMM-yyyy") & "&EndDate=" & EndStart.ToString("dd-MMM-yyyy") & _
                        "&email=" & strEmail & "&stylesheet=custcoll.xsl"
                    If blnGroup Then
                        ActionString += "&GROUP=TRUE"
                    End If
                    'While (myThread.ThreadState And Threading.ThreadState.WaitSleepJoin) <> Threading.ThreadState.WaitSleepJoin
                    'S'ystem.Threading.Thread.CurrentThread.Sleep(10)
                    'End While
                    'myThread.Suspend()
                    Application.DoEvents()
                    Dim strHTML As String = DoNavigate(ActionString)
                    Medscreen.BlatEmail("Customer Vessel Report for " & aVessNode.Client.SMIDProfile, strHTML, strEmail)
                    'If myThread.ThreadState = Threading.ThreadState.Suspended Then myThread.Resume()
                    'Me.StatusBarPanel1.Text = "Background process started"
                    'Me.StatusBar1.Invalidate()
                    Application.DoEvents()
                    Exit Sub
                ElseIf aVessNode._type = Treenodes.CollectionNodes.CollectionHeaderNode.HeaderType.Smid Then
                    strCust = aVessNode.Vessels.CustomerID
                    blnGroup = True
                    If TypeOf aVessNode.Parent Is Treenodes.CustomerNodes.CustSMIDNode Then
                        Dim asmidnode As Treenodes.CustomerNodes.CustSMIDNode = aVessNode.Parent
                        vntRet = GetParameter(MyTypes.typDate, "Start Date")
                        If vntRet Is Nothing Then Exit Sub
                        datStart = vntRet
                        vntRet = GetParameter(MyTypes.typDate, "End Date")
                        If vntRet Is Nothing Then Exit Sub
                        EndStart = vntRet
                        vntRet = GetParameter(MyTypes.typString, "Email Address(s)", , strEmail)
                        If vntRet Is Nothing Then Exit Sub
                        strEmail = vntRet
                        Dim strStyleSheet As String = "custcollg.xsl"                                      'Set default spreadsheet
                        Dim isMaritimeCheck As String
                        If asmidnode.CurrentCustomerViewState = Treenodes.AlphaNodeList.CustomerViewState.ViewBySage Then           'See how we are viewing the data and find out if maritime
                            isMaritimeCheck = CConnection.PackageStringList("Lib_customer.IsMaritimeSage", asmidnode.SMID.SMID)
                        Else
                            isMaritimeCheck = CConnection.PackageStringList("Lib_customer.IsMaritimeSMID", asmidnode.SMID.SMID)
                        End If
                        'if not maritime change the styleSheet
                        If isMaritimeCheck = "F" Then strStyleSheet = "custcollgSite.xsl"
                        Dim strHTML As String = asmidnode.SMID.VesselReport(datStart, EndStart.ToString("dd-MMM-yyyy"), strStyleSheet, strEmail)
                        Medscreen.BlatEmail("Customer Vessel Report for " & asmidnode.SMID.SMID, strHTML, strEmail)
                        'MyBase.UpdateSender(sender)
                    End If
                End If
                If strCust = "" Then Exit Sub

            End If

        End Sub


    End Class
End Namespace