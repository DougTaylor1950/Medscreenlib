Imports MedscreenLib
Imports Intranet.intranet.Support
Imports Intranet.intranet
Imports System.Windows.Forms
Namespace menus

    Public Class SampleMenu
        Inherits MenuBase
        Private Shared myInstance As SampleMenu
        Private Shared blnCanCancel As Boolean = False
        Private Const cstMROFilePos As Integer = 2
        Private Const cstScreenFilePos As Integer = 3
        Private Sub New()

        End Sub

        Public Shared Function GetSampleMenu() As SampleMenu
            If myInstance Is Nothing Then
                myInstance = New SampleMenu()

            End If
            Return myInstance
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Check to see if the sample is part of a job
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[ taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Function doJobCheck(ByVal aSample As Intranet.intranet.sample.OSample) As Boolean
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                Dim strReturn As String = CConnection.PackageStringList("Lib_uncommit.GetIsInJob", aSample.Barcode)
                If strReturn = "T" Then 'We need to check that the user knows they can deal with a job 
                    Dim intRet As MsgBoxResult = MsgBox("This sample - " & aSample.Barcode & " - is part of a collection job - " & aSample.JobName & _
                    " if you are wanting to do all the samples in the collection use the menus available on the collection node (above).  It is much easier - Do you want to use the collection menus?", _
                     MsgBoxStyle.YesNo Or MsgBoxStyle.Information)
                    If intRet = MsgBoxResult.Yes Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return True
            End If
        End Function
        ''' <developer></developer>
        ''' <summary>
        ''' Display the audit trail for a sample at Cardiff
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision><created>05-oct-2011</created></revision></revisionHistory>
        Private Sub mnuCardiffAudit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If SampleNode.sample.Barcode(0) = "H" Then
                'Get the sample pack no 
                Dim SampPackRow As DAOCardiffStarLims.Sample.ORDERSRow = DAOCardiffStarLims.DataTables.GetSample(SampleNode.sample.Barcode)
                Dim SampPackNo As String
                If SampPackRow IsNot Nothing Then                   'Sample exists
                    SampPackNo = SampPackRow.SAMPACKNO              'Find the SamplePackNo
                    Dim Report As New MedscreenLib.CRMenuItem With {.MenuPath = MedscreenCommonGui.My.Settings.CardiffAuditReport, .Instance = "SERVER4", .Username = "NULL", .Password = "NULL", .IsFormula = False}
                    Dim Parameter = New MedscreenLib.CRFormulaItem() With {.Formula = "Sample Pack No", .ParamType = "PARAM", .Value = SampPackNo}
                    Report.Formulae.Add(Parameter)
                    Dim cr As CrystalDecisions.CrystalReports.Engine.ReportDocument = Report.LogInCrystal("")
                    Report.FillFormulaFromValue(cr)
                    Dim FileName As String = MedscreenLib.CrystalSupport.ExportReport(cr, Constants.SendMethod.PDF)
                    Medscreen.BlatEmail("Audit trail for barcode : " & SampleNode.sample.Barcode, "Enclosed your Audit trail for barcode : " & SampleNode.sample.Barcode, MedscreenLib.Glossary.Glossary.CurrentSMUserEmail, , , , FileName)
                    System.Threading.Thread.Sleep(1000)
                    If IO.File.Exists(FileName) Then IO.File.Delete(FileName)

                End If

            End If

        End Sub

        Private Sub mnuCardiffCert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If SampleNode.sample.Barcode(0) = "H" Then
                'Get the sample pack no 

                Dim strFileName As String = BAOCardiff2.Certificates.GetLatestOrderCertificate(Mid(SampleNode.sample.Barcode, 3))
                If strFileName.Trim.Length > 0 Then
                    Medscreen.ShowPDF(strFileName, "Barcode no " & SampleNode.sample.Barcode)
                End If
                'Dim SampPackRow As DAOCardiffStarLims.DataSet1.ORDERSRow
                'Try
                '    SampPackRow = DAOCardiffStarLims.DataTables.GetSample(SampleNode.sample.Barcode)
                '    Dim SampPackNo As String
                '    If SampPackRow IsNot Nothing Then                   'Sample exists



                '        SampPackNo = SampPackRow.SAMPACKNO              'Find the SamplePackNo

                '        Dim CommandString As String = "REPORTTYPE=CRYSTAL METHOD=Cardiffstarlims\TESTCERTD.RPT  INSTANCE=SERVER4 PARAMETERS=Samppackno:" & SampPackNo & " BYPARAMETER=TRUE EMAIL=" & MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
                '        Dim BackReport As MedscreenLib.BackgroundReport = MedscreenLib.BackgroundReport.CreateBackgroundReport
                '        BackReport.CommandString = CommandString
                '        Dim strOutput As String = ""
                '        Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.Constants.SendMethod))
                '        Dim obj As Object = Medscreen.GetParameter(GetType(MedscreenLib.Constants.SendMethod), "Test")
                '        If Not obj Is Nothing Then
                '            strOutput = CStr(obj)
                '        End If
                '        BackReport.OutPutMethod = strOutput
                '        BackReport.Status = "R"
                '        BackReport.DoUpdate()
                '        MsgBox("Your report will be sent to a background report server and will be emailed to you when complete", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)



                '    End If
                'Catch ex As Exception
                'MsgBox("Can't show certificate - error " & ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                'End Try
            End If


        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory>
        ''' <revision><modified>6-oct-2011  Added checks to sample status, moved app to config file</modified><Author>Taylor</Author></revision>
        ''' </revisionHistory>
        Private Sub mnuCardiffImport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If SampleNode.sample.Barcode(0) = "H" Then                                      'Need to check that it is an Evidence Number
                Dim CmdString As String = MedscreenCommonGui.My.Settings.CardiffImportApp   'Move to config file
                'Do Sample status checks
                Dim SampleStatus As String = SampleNode.sample.Status

                If SampleStatus = "M" OrElse SampleStatus = "A" Then
                    MsgBox("This sample needs to be unautorised first! Terminating import", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

                If Not (SampleStatus = "V" OrElse SampleStatus = "C") Then
                    MsgBox("There may be issues if updating results!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
                'Create process rather than shelling to gain more control
                Dim myProcess As Process = New Process()
                'set program to use and arguments and window style then start process
                Dim StrBlat As String = "/BARCODE = " & SampleNode.sample.Barcode

                myProcess.StartInfo.FileName = CmdString
                myProcess.StartInfo.Arguments = StrBlat
                myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.UseShellExecute = False                 'Required for redirecting
                MedscreenLib.Medscreen.LogAction("starting process - " & StrBlat)

                myProcess.Start()
                Dim iso As IO.StreamReader = myProcess.StandardOutput

                'Wait 5 secs for the process to exit naturally if not kill the process
                myProcess.WaitForExit(5000)
                If Not myProcess.HasExited Then
                    myProcess.Kill()
                End If
                Dim StrLog As String = iso.ReadToEnd
                MedscreenLib.Medscreen.LogAction(myProcess.ExitTime & "." & "  Exit Code: " & myProcess.ExitCode & " - " & StrLog & " Barcode : " & SampleNode.sample.Barcode)
                iso.Close()



            Else
                MsgBox("this Barcode " & SampleNode.sample.Barcode & " doesn't appear to be a Trichotech Evidence No", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Unauthorise a sample
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [22/06/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuUnAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If Not doJobCheck(SampleNode.sample) Then
                Try
                    Dim oCmd As New OleDb.OleDbCommand("Lib_Uncommit.UnAuthoriseSample", CConnection.DbConnection)
                    Dim intRet As Integer
                    Dim objNonc As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Non Conformance no", , "NONC")
                    If objNonc Is Nothing Then Exit Sub
                    Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                    If objReason Is Nothing Then
                        objReason = ""
                    End If
                    Dim strReason As String = CStr(objReason)
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("noncRef", CStr(objNonc), 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery()
                    End If
                    If intRet = 1 Then
                        MsgBox("This Sample - " & SampleNode.sample.Barcode & " - has been succesfully unauthorised.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                    SampleNode.NotifyChange()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)

                End Try
            End If
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' un report a sample calling PL/SQL uncommit library
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuUnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If Not doJobCheck(SampleNode.sample) Then
                Try
                    Dim oCmd As New OleDb.OleDbCommand("Lib_Uncommit.UnreportSample", CConnection.DbConnection)
                    Dim intRet As Integer
                    Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                    If objReason Is Nothing Then
                        Exit Sub
                    End If
                    Dim strReason As String = CStr(objReason)
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery()
                    End If
                    If intRet = 1 Then
                        MsgBox("This Sample - " & SampleNode.sample.Barcode & " - has been succesfully uncreported.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                    'Unreport if possible
                    oCmd = New OleDb.OleDbCommand("lib_export.Unexport", CConnection.DbConnection)
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("SampleId", SampleNode.sample.IDNumeric, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("Status", "A", 10))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery()
                    End If


                    SampleNode.NotifyChange()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)

                End Try
            End If
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' uncommit a sample 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' calls PL/SQL routin UNCOMMIT library
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuUnCommit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If Not doJobCheck(SampleNode.sample) Then
                Try
                    Dim oCmd As New OleDb.OleDbCommand("Lib_Uncommit.UncommitSample", CConnection.DbConnection)
                    Dim intRet As Integer
                    Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                    If objReason Is Nothing Then
                        Exit Sub
                    End If
                    Dim strReason As String = CStr(objReason)
                    oCmd.CommandType = CommandType.StoredProcedure
                    oCmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                    If CConnection.ConnOpen Then
                        intRet = oCmd.ExecuteNonQuery()
                    End If
                    If intRet = 1 Then
                        MsgBox("This Sample - " & SampleNode.sample.Barcode & " - has been succesfully uncommitted.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                    SampleNode.NotifyChange()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)

                End Try
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' uncancel a sample
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [29/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuUnCancel_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            Try
                Dim oCmd As New OleDb.OleDbCommand("Lib_Uncommit.UncancelSample", CConnection.DbConnection)
                Dim intRet As Integer
                Dim objNonc As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Non Conformance no", , "NONC")
                If objNonc Is Nothing Then Exit Sub
                Dim objReason As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Reason to uncommit", , , False)
                If objReason Is Nothing Then
                    objReason = ""
                End If
                Dim strReason As String = CStr(objReason)
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                oCmd.Parameters.Add(CConnection.StringParameter("noncRef", CStr(objNonc), 10))
                oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                oCmd.Parameters.Add(CConnection.StringParameter("reason", strReason, strReason.Length))
                If CConnection.ConnOpen Then
                    intRet = oCmd.ExecuteNonQuery()
                End If
                'need to add acknowledge code here if success
                If intRet = 1 Then
                    MsgBox("This Sample - " & SampleNode.sample.Barcode & " - has been succesfully uncancelled. This will now be invoiced to the customer, get accounts to check!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
                SampleNode.NotifyChange()
            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            End Try
            'End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Change sample customer and ready sample for re-reporting
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuChangeCust_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeSampleCustomer("REPORT", sender)
            End If

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Change sample customer, set sample on hold ready for re-testing
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuChangeCustRetest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeSampleCustomer("RETEST", sender)
            End If

        End Sub

        Private Sub mnuLateConfirm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            Dim myAnalysis As New MedscreenLib.Glossary.PhraseCollection("lib_cctool.AnalysisList", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction)
            Dim oAnalysis As Object = Medscreen.GetParameter(myAnalysis, "Test needing to be added")
            If Not oAnalysis Is Nothing Then
                Dim intReturn As Integer
                Try
                    Dim AnalName As String = CConnection.PackageStringList("lib_sample8.GetAnalysisName", CStr(oAnalysis))
                    Dim ocmd As New OleDb.OleDbCommand("lib_uncommit.CreateGCMSRequest", CConnection.DbConnection)
                    ocmd.CommandType = CommandType.StoredProcedure
                    If CConnection.ConnOpen Then
                        ocmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                        ocmd.Parameters.Add(CConnection.StringParameter("testName", AnalName, AnalName.Length))
                        ocmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                        intReturn = ocmd.ExecuteNonQuery
                    End If
                Catch ex As Exception
                    Medscreen.LogError(ex, True, "Sample Late Confirmation")
                End Try
                If intReturn > 0 Then
                    MsgBox("The process appears to have worked, the sample should be set up for late confirmations", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                Else
                    MsgBox("The process may not have worked, we didn't get a response from the background process", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If

            End If

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' change sample customer and ready for re-invoicing
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuChangeCustInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            If MedscreenLib.Glossary.Glossary.UserHasRole("MED_UNDO_SAMPLES") Or MedscreenLib.Glossary.Glossary.UserHasRole("IT_SUPPORT") Then
                ChangeSampleCustomer("INVOICE", sender)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Common point to sample customer change
        ''' 
        ''' </summary>
        ''' <param name="Action"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/07/2009]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub ChangeSampleCustomer(ByVal Action As String, ByVal sender As Object)
            'First task get new customer 
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If

            Dim myParent As Treenodes.SampleHeaderNode = SampleNode.Parent
            If Not doJobCheck(SampleNode.sample) Then
                Dim objCust As Intranet.intranet.customerns.Client
                Dim objSelCust As New MedscreenCommonGui.frmCustomerSelection2()
                'Show AO customers
                objSelCust.AnalysisOnly = True
                If objSelCust.ShowDialog = DialogResult.OK Then
                    objCust = objSelCust.Client

                    Dim objNonc As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Non Conformance no", , "NONC")
                    If objNonc Is Nothing Then Exit Sub
                    Dim oCmd As New OleDb.OleDbCommand("lib_uncommit.ChangeSampleCustomer", CConnection.DbConnection)
                    Dim intRet As Integer
                    Dim strReason As String = Action
                    'if there are samples we need to test if the test panel is similar
                    Dim oColl As New Collection()
                    oColl.Add(SampleNode.sample.Barcode)
                    oColl.Add(objCust.Identity)
                    Dim strTestPanel As String = CConnection.PackageStringList("lib_uncommit.GetTestPanelStatus", oColl)
                    'Check to see if there is any problem with the panel
                    If Not strTestPanel Is Nothing AndAlso strTestPanel.Trim.Length > 0 AndAlso Action <> "RETEST" Then
                        Dim strMsg = "Tests that may need to be performed are " & vbCrLf & strTestPanel & vbCrLf & vbCrLf & _
                        "If there are any Xs(Not Done) or Ts(Not on Panel)listed the sample should be retested" & vbCrLf & _
                        "You chose " & Action & " Initially" & vbCrLf & _
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
                    Dim objInvoice As Object
                    Dim strInvoice As String = ""
                    Dim strOldSmid As String = SampleNode.sample.Client.SMIDProfile

                    'If Action = "INVOICE" Then
                    objInvoice = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Invoice number", "If left blank invoice will be created", myParent.InvoiceNumber)
                    If Not objInvoice Is Nothing Then strInvoice = CStr(objInvoice).ToUpper
                    'End 'If
                    Try
                        oCmd.CommandType = CommandType.StoredProcedure
                        oCmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 40))
                        oCmd.Parameters.Add(CConnection.StringParameter("newCustID", objCust.Identity, 10))
                        oCmd.Parameters.Add(CConnection.StringParameter("noncRef", CStr(objNonc), 10))
                        oCmd.Parameters.Add(CConnection.StringParameter("userID", MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, 10))
                        oCmd.Parameters.Add(CConnection.StringParameter("action", strReason, 10))
                        Dim oParam As OleDb.OleDbParameter = CConnection.StringParameter("InvNum", strInvoice, 10)
                        oCmd.Parameters.Add(oParam)
                        If CConnection.ConnOpen Then
                            intRet = oCmd.ExecuteNonQuery
                        End If
                        'update the local copy
                        SampleNode.sample.TableSet = "SAMPLE"
                        SampleNode.sample.Refresh()
                        Dim strMessage As String = "Barcode - " & SampleNode.sample.Barcode & " has been moved from " & strOldSmid & " to " & objCust.SMIDProfile
                        strMessage += " Invoiced on - " & SampleNode.sample.InvoiceNumber
                        myParent.InvoiceNumber = SampleNode.sample.InvoiceNumber
                        SampleNode.sample.CustomerID = objCust.Identity
                        SampleNode.sample.Client = Nothing
                        SampleNode.Refresh()
                        MsgBox(strMessage)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
                SampleNode.raiseChange()
            End If
        End Sub

        Private Sub mnuResendAsEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            'Check we have a fax message
            Dim strSend As String = CConnection.PackageStringList("Lib_sample8.ReportSentBy", SampleNode.sample.Barcode)
            If strSend = "MANFAX" Or strSend = "FAX" Then
            Else
                MsgBox("Not faxed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub

            End If
            Dim strRepFilename As String = Intranet.intranet.sample.OSample.GetReportFile(SampleNode.sample.Barcode)
            If strRepFilename.Trim.Length > 0 Then
                If CStr(strRepFilename).Trim.Length > 0 Then
                    Dim StrNewName As String

                    Dim strfilename As String = sample.OSample.FindReportFile(strRepFilename, StrNewName)
                    'We should have already checked but double checking won't do any harm.
                    If strfilename.Trim.Length > 0 AndAlso IO.File.Exists(strfilename) Then
                        Dim strFile As String
                        'open and read in file.
                        Dim iof As New IO.StreamReader(strfilename)
                        strFile = iof.ReadToEnd
                        iof.Close()
                        Dim iPosEmail As Integer = InStr(strFile, "FAX:")
                        Dim iCRET As Integer
                        Dim strEmail As String = ""
                        If iPosEmail > 0 Then
                            'We need to replace the Fax with email and customer and subject info
                            iPosEmail += 4
                            iCRET = InStr(iPosEmail, strFile, vbCr)
                            If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - 1 - iPosEmail)
                            Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Fax address", , strEmail.Trim)
                            If objNewAddress Is Nothing Then
                                Exit Sub
                            End If
                            Dim strOutFile As String = Mid(strFile, 1, iPosEmail)
                            strOutFile += CStr(objNewAddress) & Mid(strFile, iCRET - 1)
                            Dim iofo As New IO.StreamWriter(StrNewName)
                            iofo.Write(strOutFile)
                            iofo.Flush()
                            iofo.Close()
                        End If
                    Else
                        MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                Else
                    MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
            Else
                MsgBox("No report to send", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
            End If

        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' Create new output file 
        ''' </summary>
        ''' <param name="StrNewName"></param>
        ''' <param name="strOutFile"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory></revisionHistory>
        Private Shared Sub CreateNewOutputFile(ByVal StrNewName As String, ByVal strOutFile As String)
            Dim iofo As New IO.StreamWriter(StrNewName)
            iofo.Write(strOutFile)
            iofo.Flush()
            iofo.Close()
        End Sub
        Private Shared Function GetStrOutFile(ByVal strFile As String, ByRef iPosEmail As Integer, ByRef iCRET As Integer, ByVal objNewAddress As Object) As String
            Dim strOutFile As String = Mid(strFile, 1, iPosEmail)
            Dim i As Integer = strOutFile.Length - 1
            While strOutFile.Chars(i) <> "="
                i -= 1
            End While
            strOutFile = Mid(strOutFile, 1, i + 1)
            strOutFile += CStr(objNewAddress) & Mid(strFile, iCRET)
            Return strOutFile
        End Function

        ''' <developer></developer>
        ''' <summary>
        ''' Get one new email address
        ''' </summary>
        ''' <param name="strFile"></param>
        ''' <param name="iPosEmail"></param>
        ''' <param name="iCRET"></param>
        ''' <param name="strEmail"></param>
        ''' <param name="strOutFile"></param>
        ''' <param name="shouldReturn"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory></revisionHistory>
        Private Shared Function GetNewEmailAddress(ByVal strFile As String, ByRef iPosEmail As Integer, ByRef iCRET As Integer, ByRef strEmail As String, ByRef shouldReturn As Boolean) As String
            shouldReturn = False
            iPosEmail += 7
            iCRET = InStr(iPosEmail, strFile, vbCr)
            If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - 1 - iPosEmail)
            Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email address", , strEmail.Trim)
            If objNewAddress Is Nothing Then
                shouldReturn = True : Exit Function
            End If

            Dim strOutFile As String = GetStrOutFile(strFile, iPosEmail, iCRET, objNewAddress)
            Return strOutFile
        End Function

        Private Shared Sub PrisonsGetRecipient(ByVal StrNewName As String, ByVal strFile As String)
            Dim iPosEmail As Integer = InStr(strFile.ToLower, "output_address=")
            Dim iCRET As Integer
            Dim strEmail As String = ""
            If iPosEmail > 0 Then
                iPosEmail += CStr("output_address=").Length
                iCRET = InStr(iPosEmail, strFile, vbCr)
                If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - iPosEmail)
                Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "MDT address", , strEmail.Trim)
                If objNewAddress Is Nothing Then
                    Exit Sub
                End If
                Dim strOutFile As String = GetStrOutFile(strFile, iPosEmail, iCRET, objNewAddress)
                CreateNewOutputFile(StrNewName, strOutFile)
            End If
        End Sub

        Public Shared Sub DoResendCertificate(ByVal strRepFilename As String)
            If strRepFilename.Trim.Length > 0 Then
                If CStr(strRepFilename).Trim.Length > 0 Then
                    Dim StrNewName As String

                    Dim strfilename As String = sample.OSample.FindReportFile(strRepFilename, StrNewName)
                    'We should have already checked but double checking won't do any harm.
                    If InStr(strfilename.ToUpper, "CERTS_HMP") > 0 Then
                        Dim strFile As String
                        'open and read in file.
                        Dim iof As New IO.StreamReader(strfilename)
                        strFile = iof.ReadToEnd
                        iof.Close()
                        PrisonsGetRecipient(StrNewName, strFile)
                    ElseIf strfilename.Trim.Length > 0 AndAlso IO.File.Exists(strfilename) AndAlso InStr(strfilename, "HMP-") = 0 Then
                        Dim strFile As String
                        'open and read in file.
                        Dim iof As New IO.StreamReader(strfilename)
                        strFile = iof.ReadToEnd
                        iof.Close()
                        Dim iPosEmail As Integer = InStr(strFile, "E-MAIL:")
                        Dim iCRET As Integer
                        Dim strEmail As String = ""
                        If iPosEmail > 0 Then
                            Dim lShouldReturn As Boolean

                            Dim strOutFile As String = GetNewEmailAddress(strFile, iPosEmail, iCRET, strEmail, lShouldReturn)
                            If lShouldReturn Then
                                Return
                            End If
                            strFile = strOutFile
                            iPosEmail = InStr(iPosEmail + 8, strFile, "E-MAIL:")
                            While iPosEmail > 0
                                strOutFile = GetNewEmailAddress(strFile, iPosEmail, iCRET, strEmail, lShouldReturn)
                                If lShouldReturn Then
                                    Return
                                End If
                                strFile = strOutFile
                                iPosEmail = InStr(iPosEmail + 8, strFile, "E-MAIL:")

                            End While
                            CreateNewOutputFile(StrNewName, strOutFile)
                        Else 'look for a FAX number
                            iPosEmail = InStr(strFile, "FAX:")
                            If iPosEmail > 0 Then
                                iPosEmail += 4
                                iCRET = InStr(iPosEmail, strFile, vbCr)
                                If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - 1 - iPosEmail)
                                Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Fax address", , strEmail.Trim)
                                If objNewAddress Is Nothing Then
                                    Exit Sub
                                End If
                                Dim strOutFile As String = GetStrOutFile(strFile, iPosEmail, iCRET, objNewAddress)

                                'Dim strOutFile As String = Mid(strFile, 1, iPosEmail)
                                'strOutFile += CStr(objNewAddress) & Mid(strFile, iCRET - 1)
                                CreateNewOutputFile(StrNewName, strOutFile)
                                'Dim iofo As New IO.StreamWriter(StrNewName)
                                'iofo.Write(strOutFile)
                                'iofo.Flush()
                                'iofo.Close()
                            End If
                        End If
                    ElseIf InStr(strfilename, "HMP-") > 0 Then
                        Dim strFile As String
                        'open and read in file.
                        Dim iof As New IO.StreamReader(strfilename)
                        strFile = iof.ReadToEnd
                        iof.Close()
                        PrisonsGetRecipient(StrNewName, strFile)

                    Else
                        MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                Else
                    MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
                Else
                    MsgBox("No report to send", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
        End Sub
        Private Shared Sub ResendCertificate(ByVal Barcode As String)
            Dim strRepFilename As String = Intranet.intranet.sample.OSample.GetReportFile(Barcode)
            DoResendCertificate(strRepFilename)
        End Sub
        Private Sub mnuResendCert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            ResendCertificate(SampleNode.sample.Barcode)
        End Sub

        Private Sub mnuResendBUPACert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            'need to cancel export
            'update sample_action set status = 'X' where action_id = 'DATAEXPORT' and sample_id  in (select id_numeric from view_all_samples where bar_code = '&barcode');
            Dim cmd As New OleDb.OleDbCommand
            Dim intRet As Integer = 0
            Try
                cmd.CommandText = "update sample_action set status = 'X' where action_id = 'DATAEXPORT' and sample_id  in (select id_numeric from view_all_samples where bar_code = ?)"
                cmd.Connection = CConnection.DbConnection
                cmd.Parameters.Add(CConnection.StringParameter("barcode", SampleNode.sample.Barcode, 10))
                If CConnection.ConnOpen Then
                    intRet = cmd.ExecuteNonQuery()
                End If
            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()

            End Try

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Taylor]</Author><date> [22/02/2010]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub mnuResendScreen_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            Dim strResultFiles As String = CConnection.PackageStringList("Lib_cctool.GeTReportFiles", SampleNode.sample.Barcode)
            If Not strResultFiles Is Nothing AndAlso strResultFiles.Trim.Length > 0 Then 'Split files
                Dim strResFArray As String() = strResultFiles.Split(New Char() {","})
                If strResFArray.Length > 0 Then
                    Dim strFilename As String = ""
                    If CStr(strResFArray.GetValue(cstScreenFilePos)).Trim.Length > 0 Then
                        Dim StrNewName As String
                        strFilename = MedscreenLib.Constants.WordOutput & "HMP-" & strResFArray.GetValue(cstScreenFilePos) & ".all_done"
                        StrNewName = MedscreenLib.Constants.WordOutput & "HMP-" & strResFArray.GetValue(cstScreenFilePos)
                        StrNewName += "R-" & Now.ToString("ddHH") & ".out"
                        Dim tmpDate As Date = Today
                        If Not IO.File.Exists(strFilename) Then     'Can't find file
                            'get date info out of sample
                            Dim intPos As Integer = InStr(strFilename, ".")
                            If intPos > 0 Then      'Filename in the right format
                                Dim strmonth As String = SampleNode.sample.DateReported.ToString("MMM")
                                Dim strYear As String = SampleNode.sample.DateReported.ToString("yyyy")
                                tmpDate = Date.Parse("01-" & strmonth & "-" & strYear)
                                Dim newPath As String = MedscreenLib.Constants.BackupFolder & "\sm_data_" & tmpDate.ToString("yyyy-MM")
                                strFilename = Medscreen.ReplaceString(strFilename.ToLower, "\\john\live", newPath)
                            End If
                            'ensure that we have a drive mapped
                            Try
                                Dim myNetDrive As New MedscreenLib.Network.NetworkDrive()
                                myNetDrive.LocalDrive = Medscreen.BackupDrive
                                myNetDrive.ShareName = Medscreen.BackupFolder
                                myNetDrive.Force = True
                                myNetDrive.MapDrive("admin", "antioch")
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If
                        'Try and deal with end of month issues
                        If Not IO.File.Exists(strFilename) AndAlso tmpDate <> Today Then
                            Dim strCurrMonth As String = tmpDate.ToString("yyyy-MM") & "\"
                            tmpDate = tmpDate.AddMonths(1)
                            Dim strNewMonth As String = tmpDate.ToString("yyyy-MM") & "\"
                            strFilename = Medscreen.ReplaceString(strFilename, strCurrMonth, strNewMonth)

                        End If
                        If strFilename.Trim.Length > 0 AndAlso IO.File.Exists(strFilename) Then
                            Dim strFile As String
                            'open and read in file.
                            Dim iof As New IO.StreamReader(strFilename)
                            strFile = iof.ReadToEnd
                            iof.Close()
                            Dim iPosEmail As Integer = InStr(strFile, "output_address=")
                            Dim iCRET As Integer
                            Dim strEmail As String = ""
                            If iPosEmail > 0 Then
                                iPosEmail += 15
                                iCRET = InStr(iPosEmail, strFile, vbCr)
                                If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - iPosEmail)
                                Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email address", , strEmail.Trim)
                                If objNewAddress Is Nothing Then
                                    Exit Sub
                                End If
                                Dim strOutFile As String = Mid(strFile, 1, iPosEmail - 1)
                                strOutFile += CStr(objNewAddress) & Mid(strFile, iCRET)
                                Dim iofo As New IO.StreamWriter(StrNewName)
                                iofo.Write(strOutFile)
                                iofo.Flush()
                                iofo.Close()
                            End If
                        Else
                            MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                        End If
                    Else
                        MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                Else
                    MsgBox("No report to send", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
            End If
        End Sub

        Private Sub mnuResendMRO_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            Dim strResultFiles As String = CConnection.PackageStringList("Lib_cctool.GeTReportFiles", SampleNode.sample.Barcode)
            If Not strResultFiles Is Nothing AndAlso strResultFiles.Trim.Length > 0 Then 'Split files
                Dim strResFArray As String() = strResultFiles.Split(New Char() {","})
                If strResFArray.Length > 0 Then
                    Dim strFilename As String = ""
                    If CStr(strResFArray.GetValue(cstMROFilePos)).Trim.Length > 0 Then
                        Dim StrNewName As String
                        strFilename = Medscreen.ReplaceString(strResFArray.GetValue(cstMROFilePos), "SMP$CERTS:", MedscreenLib.Constants.CertsDirectory)
                        StrNewName = Medscreen.ReplaceString(strResFArray.GetValue(cstMROFilePos), "SMP$CERTS:", MedscreenLib.Constants.EmailResults) ' MedConnection.Instance.ServerPath & "\Dm_datafile\emails\")
                        StrNewName = Medscreen.ReplaceString(StrNewName, "DAILYMRO.", "MRO-")
                        StrNewName += "R-" & Now.ToString("ddHH") & ".ema"
                        Dim tmpDate As Date = Today
                        If Not IO.File.Exists(strFilename) Then     'Can't find file
                            'get date info out of filename and change to archive path
                            Dim intPos As Integer = InStr(strFilename, ".")
                            If intPos > 0 Then      'Filename in the right format
                                Dim intSep As Integer = 1
                                If strFilename.Chars(intPos + 1) = "-" Then intSep = 0
                                Dim strmonth As String = Mid(strFilename, intPos + 3 + intSep, 3)
                                Dim strYear As String = Mid(strFilename, intPos + 7 + intSep, 4)
                                tmpDate = Date.Parse("01-" & strmonth & "-" & strYear)
                                Dim newPath As String = MedscreenLib.Constants.BackupFolder & "\sm_data_" & tmpDate.ToString("yyyy-MM")
                                strFilename = Medscreen.ReplaceString(strFilename.ToLower, "\\john\live", newPath)
                            End If
                            'ensure that we have a drive mapped
                            Try
                                'Dim myNetDrive As New MedscreenLib.Network.NetworkDrive()
                                'myNetDrive.LocalDrive = Medscreen.BackupDrive
                                'myNetDrive.ShareName = Medscreen.BackupFolder
                                'myNetDrive.Force = True
                                'myNetDrive.MapDrive("admin", "antioch")
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If
                        'Try and deal with end of month issues
                        If Not IO.File.Exists(strFilename) AndAlso tmpDate <> Today Then
                            Dim strCurrMonth As String = tmpDate.ToString("yyyy-MM") & "\"
                            tmpDate = tmpDate.AddMonths(1)
                            Dim strNewMonth As String = tmpDate.ToString("yyyy-MM") & "\"
                            strFilename = Medscreen.ReplaceString(strFilename, strCurrMonth, strNewMonth)

                        End If
                        If strFilename.Trim.Length > 0 AndAlso IO.File.Exists(strFilename) Then
                            Dim strFile As String
                            'open and read in file.
                            Dim iof As New IO.StreamReader(strFilename)
                            strFile = iof.ReadToEnd
                            iof.Close()
                            Dim iPosEmail As Integer = InStr(strFile, "E-MAIL:")
                            Dim iCRET As Integer
                            Dim strEmail As String = ""
                            If iPosEmail > 0 Then
                                iPosEmail += 7
                                iCRET = InStr(iPosEmail, strFile, vbCr)
                                If iCRET > 0 Then strEmail = Mid(strFile, iPosEmail, iCRET - 1 - iPosEmail)
                                Dim objNewAddress As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Email address", , strEmail.Trim)
                                If objNewAddress Is Nothing Then
                                    Exit Sub
                                End If
                                Dim strOutFile As String = Mid(strFile, 1, iPosEmail)
                                strOutFile += CStr(objNewAddress) & Mid(strFile, iCRET - 1)
                                Dim iofo As New IO.StreamWriter(StrNewName)
                                iofo.Write(strOutFile)
                                iofo.Flush()
                                iofo.Close()
                            End If
                        Else
                            MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                        End If
                    Else
                        MsgBox("Can't find report file contact IT.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                Else
                    MsgBox("No report to send", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                End If
            End If
        End Sub


        Private frmMoveBC As frmBarcodes

        Private Sub mnuMoveBarcode_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim SenderType As Object = GetSenderType(sender)
            Dim SampleNode As Treenodes.SampleNode
            If TypeOf SenderType Is Treenodes.SampleNode Then
                SampleNode = SenderType
            Else
                Exit Sub
            End If
            If frmMoveBC Is Nothing Then
                frmMoveBC = New frmBarcodes()
            End If
            Dim aNode As Treenodes.SampleNode
            For Each aNode In CType(SampleNode.TreeView, CustomTreeBase).SelectedNodes
                frmMoveBC.AddBarCode(aNode.sample.Barcode)
            Next
            Dim nSamplesProcessed As Integer = 0
            If frmMoveBC.ShowDialog = DialogResult.OK Then


            End If
        End Sub
    End Class
End Namespace
