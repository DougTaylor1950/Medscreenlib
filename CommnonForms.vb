'$Revision: 1.1 $
'$Author: taylor $
'$Date: 2005-11-09 10:06:56+00 $
'$Log: frmCustomerSelection2.vb,v $
'Revision 1.1  2005-11-09 10:06:56+00  taylor
'Loging header added
'
Imports System.Windows.Forms
Imports MedscreenLib
Imports System.Drawing
Imports System.Reflection

Imports Intranet
Imports Intranet.intranet
Imports Intranet.intranet.Support

Imports MedscreenLib.Glossary
Imports MedscreenLib.Medscreen
Imports Intranet.intranet.customerns

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : CommonForms
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Exposes common forms as shared entities
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
<CLSCompliant(False)> _
Public Class CommonForms

    ' Declaration of shared variables
    Private Shared myFindAddress As FindAddress
    Private Shared oCollheaders As UserDefaults.ColumnDefaults
    Private Shared myUserOptions As UserDefaults.UserOptions
    Private Shared myAppPath As String
    Private Shared myEdCollOfficer As frmCollOfficerEd
    Private Shared CollForm As frmCollectionForm
    Private Shared custSel As MedscreenCommonGui.frmCustomerSelection2
    Private Shared CollOfficerForm As MedscreenCommonGui.FrmCollOfficer

    Public Shared Function ApplicationPath() As String
        Dim fi As IO.FileInfo = New System.IO.FileInfo(Application.ExecutablePath)
        Dim strPath As String = fi.DirectoryName
        Return strPath

    End Function

    Public Shared Function HuntVessel() As Intranet.intranet.customerns.Vessel
        Dim Frm As MedscreenCommonGui.frmVesselSelect = New MedscreenCommonGui.frmVesselSelect()

        'ModTpanel.Vessels.RefreshChanged()
        Frm.Vessels = Nothing 'ModTpanel.Vessels


        Dim oVessel As Intranet.intranet.customerns.Vessel = Nothing

        If Frm.ShowDialog = Windows.Forms.DialogResult.OK Then

            oVessel = Frm.Vessel

        End If
        Return oVessel
    End Function

    Public Shared Property FormCollection() As frmCollectionForm
        Get
            If CollForm Is Nothing Then
                CollForm = New frmCollectionForm()
            End If
            Return CollForm
        End Get
        Set(ByVal Value As frmCollectionForm)
            CollForm = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Book a call out collection 
    ''' </summary>
    ''' <param name="objCm"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function BookCallOut(ByRef objCm As Intranet.intranet.jobs.CMJob) As Boolean
        'objCm.FieldValues.Initialise()

        Dim blnBookAgain As Boolean = True
        Dim blnReturn As Boolean = False
        Intranet.intranet.Support.cSupport.Ports.RefreshChanged()    'See if the ports list needs updating 

        While blnBookAgain
            Dim objF As New MedscreenCommonGui.frmCustomerSelection2()
            objF.CallOutOnly = True
            If objF.ShowDialog = Windows.Forms.DialogResult.OK Then
                If objF.Client Is Nothing Then Exit Function

                If Not SetupService(objF.Client, CustomerServices.GCST_CALLOUT) Then Exit Function

                objCm = New Intranet.intranet.jobs.CMJob()
                objCm.SetId = MedscreenLib.CConnection.NextSequence("SEQ_CM_JOB_NUMBER")  'Assign the ID to the collection
                objCm.FieldValues.Initialise()
                objCm.ClientId = objF.Client.Identity                       'Add client id 
                objCm.Status = Constants.GCST_JobStatusTemp                 'Status indicates that collection hasn't been actually creayted
                objCm.CollectionType = Constants.GCST_JobTypeCallout        'Set it to a call out 
                objCm.CollectionDate = Now
                objCm.UK = True                                             'Must be a UK collection

            Else
                Exit Function
            End If
            objF.Hide()
            objF = Nothing
            Dim frmCall As New MedscreenCommonGui.frmCallOut()
            With frmCall
                .Collection = objCm
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    blnReturn = True
                Else
                    blnReturn = False
                End If
            End With
            frmCall = Nothing
            If MsgBox("Do you wish to create another callout?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.No Then
                blnBookAgain = False
            End If
        End While
        Return blnReturn
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a routine collection
    ''' </summary>
    ''' <param name="blnRet">Success or not</param>
    ''' <param name="CustomerID">Customer_id, usually a null string, which signals throwing up customer selection form</param>
    ''' <param name="Vessel_id">A vessel to create collection for, usually null string</param>
    ''' <returns>Collection created</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function BookRoutine(ByRef blnRet As Boolean, _
    Optional ByVal CustomerID As String = "", Optional ByVal Vessel_id As String = "") As Intranet.intranet.jobs.CMJob

        Dim objCm As Intranet.intranet.jobs.CMJob = Intranet.intranet.jobs.CMCollection.Create("W")
        'objCm.FieldValues.Initialise()


        Cursor.Current = Cursors.WaitCursor

        Try
            'Intranet.intranet.Support.cSupport.Ports.RefreshChanged()    'See if the ports list needs updating 
            If CustomerID.Trim.Length = 0 Then
                Dim objF As MedscreenCommonGui.frmCustomerSelection2 = CommonForms.FormCustSelection
                objF.CallOutOnly = False
                If objF.ShowDialog = Windows.Forms.DialogResult.OK Then       'Get the customer
                    If objF.Client Is Nothing Then
                        objCm = Nothing
                        Return objCm
                        Exit Function
                    End If 'Bog off selection cancelled
                    If Not SetupService(objF.Client, CustomerServices.GCST_ROUTINE) Then
                        Return objCm
                        Exit Function
                    End If
                    'Need to add a check for the customer status 
                    If objF.Client.AccStatus = "D" Then
                        MsgBox("This account has been removed, you can not book collections for it!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical)
                        objCm = Nothing
                        Return objCm
                        Exit Function       'The user has cancelled 
                    ElseIf objF.Client.AccStatus = "O" Then
                        MsgBox("This account has been maked Dormant, please let sales know that the account is being reactivated.", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)
                    ElseIf objF.Client.AccStatus = "S" Then
                        MsgBox("This account has been suspended, you can not book collections for it!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical)
                        objCm = Nothing
                        Return objCm
                        Exit Function       'The user has cancelled 
                    End If
                    SetCollectionCustomerDefaults(objCm, objF.Client)
                Else
                    objCm = Nothing
                    Return objCm
                    Exit Function       'The user has cancelled 
                End If
            Else
                Dim oClient As Intranet.intranet.customerns.Client = Intranet.intranet.Support.cSupport.CustList.Item(CustomerID)
                If oClient Is Nothing Then
                    objCm = Nothing
                    Return objCm
                    Exit Function
                End If
                SetCollectionCustomerDefaults(objCm, oClient)
                If Vessel_id.Trim.Length > 0 Then
                    objCm.vessel = Vessel_id
                End If
            End If
            Cursor.Current = Cursors.WaitCursor
            With CommonForms.FormCollection
                .Collection = objCm
                .FormMode = frmCollectionForm.FormModes.Create
                If .ShowDialog = Windows.Forms.DialogResult.OK Then 'If there is a low level issue will be set to retry
                    'TODO add to list offer to assign to collector
                    blnRet = True
                Else
                    blnRet = False
                End If
            End With
            CommonForms.FormCollection = Nothing
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try
        Return objCm


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deal with credit card payments
    ''' </summary>
    ''' <param name="objjob"></param>
    ''' <param name="HasCollection"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function DoCardPay(ByVal objjob As Intranet.intranet.jobs.CMJob, Optional ByVal HasCollection As Boolean = False) As Boolean

        Dim frmCardPay As New MedscreenCommonGui.frmCardPay2()
        With frmCardPay
            Try
                '               Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                '               Me.Invalidate()
                Dim strCustId As String = objjob.ClientId
                Dim objClient As Intranet.intranet.customerns.Client = Intranet.intranet.Support.cSupport.CustList.Item(strCustId)
                '.Client = objClient
                '.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser
                .Text = "Record payment for collection"
                If Not objjob Is Nothing Then
                    .CMNumber = objjob.ID
                    .Text = "Record payment for collection : " & objjob.ID
                End If
                .ShowDialog()
            Catch ex As Exception
                Medscreen.LogError(ex, , "Recording card payment")
            End Try
            'Me.Cursor = System.Windows.Forms.Cursors.Default
            'Me.Invalidate()

        End With

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up a service for a customer
    ''' </summary>
    ''' <param name="Client"></param>
    ''' <param name="Service"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	17/10/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function SetupService(ByVal Client As Intranet.intranet.customerns.Client, ByVal Service As String) As Boolean
        Dim blnRet As Boolean = True
        Dim oColl As New Collection()
        oColl.Add("SERVICE")
        oColl.Add(Service)
        Dim strServicelist As String = MedscreenCommonGUIConfig.NodePLSQL.Item("ServiceList")
        Dim strService As String = CConnection.PackageStringList(strServicelist, oColl)

        If Not Client.HasService(Service) Then
            Dim objRet As MsgBoxResult = MsgBox("This customer does not have the '" & strService & "' service do you want to add it? " & vbCrLf & _
              "Sales will be informed.  If you select no you will not be able to create the collection!", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)
            If objRet = MsgBoxResult.No Then
                blnRet = False
                Return blnRet
                Exit Function
            End If
            'Create the new service
            Dim objSer As Intranet.intranet.customerns.CustomerService
            objSer = Client.Services.Item(Service)

            'Define a variable to say what is happening 
            Dim strChange As String = ""
            If objSer Is Nothing Then
                objSer = New Intranet.intranet.customerns.CustomerService()
                objSer.CustomerId = Client.Identity
                objSer.Service = Service
                strChange = "added"
            Else
                objSer.Removed = False
                objSer.Fields.Item("DATE_END").SetNull()
                strChange = "reinstated"
            End If
            objSer.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            Client.Services.Add(objSer)
            objSer.update()
            'Notify the change
            NotifyServiceChange(Client, strService, strChange)
        End If
        Return blnRet

    End Function

    Public Shared Sub NotifyServiceChange(ByVal Client As Intranet.intranet.customerns.Client, ByVal Service As String, ByVal Change As String)
        'Email sales
        Dim strPriceXML As String = Client.InvoiceInfo.PriceListXML
        Dim strHTML As String = ""
        Try
            Dim strStyleSheet1 As String = MedscreenCommonGUIConfig.NodeStyleSheets.Item("SetupService")
            strHTML = Medscreen.ResolveStyleSheet(strPriceXML, strStyleSheet1, 0)

        Catch ex As Exception
            'LogError(ex, , "Fill prices - customer form ")
        Finally
            strPriceXML = Nothing
        End Try
        Dim strEmailTo As String = MedscreenCommonGUIConfig.NodeStyleSheets.Item("SetupServiceEmail")
        Dim strMessage As String = MedscreenCommonGUIConfig.Messages.Item("ServiceEmail")
        Dim oColl As New Collection()                                                           'Convert service id to description 
        oColl.Add("SERVICE")
        oColl.Add(Service)
        Dim ServiceDescription As String = CConnection.PackageStringList("lib_utils.decodephrase", oColl)
        If Not ServiceDescription Is Nothing AndAlso ServiceDescription.Trim.Length > 0 Then
            Service = ServiceDescription
        End If
        Dim strBody As String = MedscreenLib.Medscreen.ExpandQuery(strMessage, Service, Change, Client.SMIDProfile) & strHTML
        'Dim strBody As String = "The service <b>'" & strService & "'</b> has been " & strChange & " for this customer " & _
        '  Client.SMIDProfile & ". Please check that the prices are correct for this service.<BR/>" & strHTML
        Medscreen.BlatEmail("Services changed for - " & Client.SMIDProfile, strBody, strEmailTo, , , , )

    End Sub

    Public Shared Sub SetCollectionCustomerDefaults(ByRef objCm As jobs.CMJob, ByVal Client As customerns.Client)
        objCm.ClientId = Client.Identity
        objCm.PurchaseOrder = Client.ContractNumber
        If Not Client.CollectionInfo Is Nothing Then
            With Client.CollectionInfo
                objCm.WhoToTest = .WhoToTest
                objCm.NoOfDonors = .NoDonors
                objCm.ConfirmationContact = .ConfirmationContactList
                objCm.ConfirmationNumber = .ConfirmationSendAddress
                objCm.ConfirmationMethod = .ConfirmationMethod
                objCm.CollType = .ReasonForTest
                objCm.ConfirmationType = .ConfirmationType
                objCm.SpecialProcedures = .Procedures
                objCm.CollectionComments = .AccessToSite
                '<Added code modified 08-Dec-2007 11:58 by LOGOS\taylor> objCm.CollectionComments = .AccessToSite
                objCm.LaboratoryId = .LaboratoryId
                '</Added code modified 08-Dec-2007 11:58 by LOGOS\taylor> 
                Dim objPhrase As MedscreenLib.Glossary.Phrase = _
                    MedscreenLib.Glossary.Glossary.TestTypeList.Item(.TestType)
                If Not objPhrase Is Nothing Then
                    objCm.Tests = objPhrase.PhraseText
                End If
                objCm.UK = .UKCollection
                If objCm.UK Then objCm.CollectionType = "U"
            End With
        Else
            If Client.IndustryCode.Length > 3 Then
                If Mid(Client.IndustryCode, 1, 3) = "MAR" Then
                    objCm.UK = False
                Else
                    objCm.UK = True
                End If
            Else
                objCm.UK = True         'If we have no info probably work place
            End If

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Book an external collection
    ''' </summary>
    ''' <param name="blnRet"></param>
    ''' <param name="CustomerID"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	12/10/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function BookExternal(ByRef blnRet As Boolean, _
    Optional ByVal CustomerID As String = "") As Intranet.intranet.jobs.CMJob
        Dim objCm As Intranet.intranet.jobs.CMJob = Intranet.intranet.jobs.CMCollection.Create("E")
        'objCm.FieldValues.Initialise()


        Cursor.Current = Cursors.WaitCursor

        Try
            Intranet.intranet.Support.cSupport.Ports.RefreshChanged()    'See if the ports list needs updating 
            If CustomerID.Trim.Length = 0 Then
                Dim objF As MedscreenCommonGui.frmCustomerSelection2 = CommonForms.FormCustSelection
                objF.CallOutOnly = False
                If objF.ShowDialog = Windows.Forms.DialogResult.OK Then       'Get the customer
                    If objF.Client Is Nothing Then
                        objCm = Nothing
                        Return objCm
                        Exit Function
                    End If
                    If Not SetupService(objF.Client, CustomerServices.GCST_EXTERNAL) Then
                        objCm = Nothing
                        Return objCm
                        Exit Function
                    End If
                    SetCollectionCustomerDefaults(objCm, objF.Client)
                Else
                    objCm = Nothing
                    Return objCm
                    Exit Function       'The user has cancelled 
                End If
            Else
                Dim oClient As Intranet.intranet.customerns.Client = Intranet.intranet.Support.cSupport.CustList.Item(CustomerID)
                If oClient Is Nothing Then
                    objCm = Nothing
                    Return objCm
                    Exit Function
                End If
                SetCollectionCustomerDefaults(objCm, oClient)
                'If Vessel_id.Trim.Length > 0 Then
                '    objCm.vessel = Vessel_id
                'End If
            End If
            Cursor.Current = Cursors.WaitCursor
            CommonForms.FormCollection = Nothing
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try
        'objCm.CollectionType = "E"
        Dim objNC As New frmCallOutCMNO()
        Dim dr As DialogResult
        objNC.DefineCMNumber = True
        objNC.DateTimePicker1.Value = Today
        'Me.DTSentToOfficer.Value = Me.DtCallTime.Value
        dr = objNC.ShowDialog
        If dr = Windows.Forms.DialogResult.Yes Then
            Try
                'Dim OISID As Integer
                'Dim objC As Intranet.intranet.jobs.CMJob = Intranet.intranet.Support.cSupport.ICollection(False).CreateCollection(Me.Collection, OISID)
                'If Not objC Is Nothing Then 'Check we managed to create one
                '    Me.CopyCollection(objc)
                'End If
                'Me.txtCmNumber.Text = objc.ID
                Dim strID As String = MedscreenLib.CConnection.NextSequence("SEQ_CM_JOB_NUMBER")
                objCm.SetId = strID

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "Getting ID")

            End Try
            Try  ' Protecting
                Cursor.Current = Cursors.WaitCursor
                With CommonForms.FormCollection

                    .Collection = objCm
                    If .ShowDialog = Windows.Forms.DialogResult.OK Then 'If there is a low level issue will be set to retry
                        'TODO add to list offer to assign to collector
                        blnRet = True
                    Else
                        blnRet = False
                    End If
                End With
                CommonForms.FormCollection = Nothing
                Cursor.Current = Cursors.Default

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "CommonForms-BookExternal-318")
            Finally
            End Try

        ElseIf dr = Windows.Forms.DialogResult.No Or dr = Windows.Forms.DialogResult.OK Then
            '    'user provided cm number 
            Try

                If Not Intranet.intranet.Support.cSupport.ICollection(False).CollectionExists(objNC.UserCMNo.ToUpper) Then
                    objCm.SetId = objNC.UserCMNo.ToUpper

                Else
                    MsgBox("That collection " & objNC.UserCMNo & " already exists!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
                    objCm = Nothing
                    Return objCm
                    Exit Function
                End If
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "Getting user generated ID")

            End Try
            Try  ' Protecting
                Cursor.Current = Cursors.WaitCursor
                With CommonForms.FormCollection

                    .Collection = objCm
                    If .ShowDialog = Windows.Forms.DialogResult.OK Then 'If there is a low level issue will be set to retry
                        'TODO add to list offer to assign to collector
                        blnRet = True
                    Else
                        blnRet = False
                    End If
                End With
                CommonForms.FormCollection = Nothing
                Cursor.Current = Cursors.Default

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "CommonForms-BookExternal-318")
            Finally
            End Try
        Else
            objCm = Nothing
            Return objCm
            Exit Function   ' User wants to bog off.

        End If

        Return objCm

    End Function

    Public Shared Property FormCustSelection() As MedscreenCommonGui.frmCustomerSelection2
        Get
            If custSel Is Nothing Then
                custSel = New MedscreenCommonGui.frmCustomerSelection2()
            End If
            Return custSel
        End Get
        Set(ByVal Value As MedscreenCommonGui.frmCustomerSelection2)
            custSel = Value
        End Set
    End Property

    Public Shared Property FormCollOfficer() As MedscreenCommonGui.FrmCollOfficer
        Get
            If CollOfficerForm Is Nothing Then
                CollOfficerForm = New MedscreenCommonGui.FrmCollOfficer()
            End If
            Return CollOfficerForm
        End Get
        Set(ByVal Value As MedscreenCommonGui.FrmCollOfficer)
            CollOfficerForm = Value
        End Set
    End Property

    'Send collection to a collecting officer, this is called from various menus
    Public Overloads Shared Sub SendToCollector(ByVal myCollection As Intranet.intranet.jobs.CMJob, ByVal blnPrintOnly As Boolean, _
        Optional ByVal SendAndPrint As Boolean = False, _
        Optional ByVal sendtype As String = "Resend", Optional ByVal DestType As String = "EMAIL")
        If Not blnPrintOnly Then
            Dim frm As New frmSendOptions()
            With frm          'Taqke a look at the options 
                .DestType = DestType
                .CMNumber = myCollection.CMID
                Dim strDefaultDestination As String = CConnection.PackageStringList("Lib_cctool.Coll_officerSendMethod", myCollection.CollOfficer)
                'Add default address
                If strDefaultDestination IsNot Nothing AndAlso strDefaultDestination.Trim.Length > 0 Then .DestAddress = strDefaultDestination
                blnPrintOnly = SendAndPrint
                Dim sendAction As String = sendtype & " " & myCollection.CMID & " Print Request " & blnPrintOnly & "  -  sendandprint " & SendAndPrint
                Medscreen.LogAction(sendAction, "SendActions.txt")
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    If .Modify Then                 'User wants to send to a non default address
                        myCollection.SendToCollector(.DestType, .DestAddress, .PrintOnly, True, sendtype, .DestType)
                    Else    'Use default address 
                        myCollection.SendToCollector("", "", blnPrintOnly, True, sendtype, .DestType)
                    End If
                    'Record what happened
                    If sendtype.ToUpper = "SENT" Then
                        myCollection.Comments.CreateComment("Collection sent to Officer " & myCollection.CollOfficer, Constants.GCST_COMM_CollSendToOfficer)
                        myCollection.LogAction(jobs.CMJob.Actions.SendToCollOfficer)
                    Else
                        myCollection.Comments.CreateComment("Collection resent to Officer " & myCollection.CollOfficer, Constants.GCST_COMM_CollReSendToOfficer)
                        myCollection.LogAction(jobs.CMJob.Actions.ReSendToCollOfficer)
                    End If
                    myCollection.Refresh()         'Get modified data
                End If
            End With
        Else
            'Just printing
            myCollection.SendToCollector("", "", True, False, sendtype)
        End If
    End Sub


    'One of the possible overloads, this takes an array of destinations so it cam multi send
    Public Overloads Shared Sub SendToCollector(ByVal myCollection As Intranet.intranet.jobs.CMJob, ByVal Destinations As String(), ByVal SendType As String)
        Dim frm As New frmSendOptions()
        With frm
            .CMNumber = myCollection.CMID
            Dim sendAction As String = SendType & " " & myCollection.CMID & " Print Request " & "  -  sendandprint "
            Medscreen.LogAction(sendAction, "SendActions.txt")
            'Call collection method 
            myCollection.SendToCollector("", "", True, True, SendType)
            myCollection.Comments.CreateComment("Collection sent to Officer " & myCollection.CollOfficer, Constants.GCST_COMM_CollSendToOfficer)
            myCollection.Refresh()         'Get modified data
        End With
    End Sub


    Public Shared Sub EditCollection(ByVal mycollection As Intranet.intranet.jobs.CMJob, Optional ByVal FormState As frmCollectionForm.FormModes = frmCollectionForm.FormModes.Edit)
        'First check is to see if it is a call out  (callout form) or routine 
        If (mycollection.CollectionType = Constants.GCST_JobTypeCallout) Or _
        (mycollection.CollectionType = Constants.GCST_JobTypeCalloutOverseas) Then
            Dim FormCallOut As New frmCallOut()
            With FormCallOut          'Use call out form 
                .Collection = mycollection
                'Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                'Me.Invalidate()
                .Collection.Refresh()
                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    mycollection = .Collection
                    If .Collection.CollectionType = Constants.GCST_JobTypeCallout Then
                        'If .Collection.Status = Constants.GCST_JobStatusCreated Then
                        '    AssignCallOutToCollector(.Collection)
                        'End If
                    End If
                End If
                'Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                'Me.Invalidate()
            End With
        Else            'Routine collection 
            With CommonForms.FormCollection
                .FormMode = FormState           'Set the form into a particular state 
                .Collection = mycollection      'Set the colelction property of the form to list view item collection 
                'Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                'Me.Invalidate()
                Try
                    .ShowDialog()       'Show form, no need to get a return as the DB will be updated if the user wants 
                Catch ex As Exception

                End Try
                'Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                'Me.Invalidate()
            End With
        End If
        'Ensure that we refresh the display 

    End Sub


    'Assign a collecting officer to a collection 
    Public Shared Function AssignOfficer(ByVal myCollection As Intranet.intranet.jobs.CMJob) As Boolean
        If myCollection.Status = Constants.GCST_JobStatusInterrupted Then
            MsgBox("Processing on this collection is interrupted, Job can't be assigned.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
            Exit Function
        End If
        Cursor.Current = Cursors.WaitCursor

        'Use the collecting officer form 
        MedscreenCommonGui.CommonForms.FormCollOfficer.collection = myCollection
        'Set collection date to collection date of colelction this is used for feedback 
        MedscreenCommonGui.CommonForms.FormCollOfficer.CollectionDate = myCollection.CollectionDate
        'Set port this is used with coll port table to identify officers 
        MedscreenCommonGui.CommonForms.FormCollOfficer.Port = myCollection.Port
        'If we editing info set current officer 
        MedscreenCommonGui.CommonForms.FormCollOfficer.Collector = myCollection.CollOfficer
        'Show form modally 
        If MedscreenCommonGui.CommonForms.FormCollOfficer.ShowDialog = Windows.Forms.DialogResult.OK Then
            If Not MedscreenCommonGui.CommonForms.FormCollOfficer.Collector Is Nothing Then
                'Check that we have a valid officer 
                myCollection.CollOfficer = MedscreenCommonGui.CommonForms.FormCollOfficer.Collector
                Try
                    If InStr("CRDM", myCollection.Status) = 0 Then
                        'Set status to assigned and save so that we have a fall back position 
                        myCollection.SetStatus(MedscreenLib.Constants.GCST_JobStatusAssigned)
                    Else
                        MsgBox("Collection has been confirmed previously, the officer can be changed but it can't be reconfirmed.", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                    End If
                    myCollection.Update()
                Catch cex As MedscreenExceptions.CanNotChangeCollectionStatus
                    MsgBox(cex.ToString)
                    Exit Function
                End Try
                If myCollection.CollectionType <> Constants.GCST_JobTypeCallout Then     'Call outs don't get sent to CO
                    If MsgBox("Do you wish to send this collection (" & myCollection.ID & _
                        ") to the collecting officer (" & _
                        myCollection.CollOfficerDescription & ") now?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                        Try
                            If InStr("CRDM", myCollection.Status) = 0 Then              'Job has already been confirmed or samples collected
                                myCollection.Status = MedscreenLib.Constants.GCST_JobStatusSent
                                myCollection.Update()
                            End If
                            'SendToCollector(False, True) send and print collection
                            myCollection.SendToCollector("", "", False, True, "Send")
                            myCollection.Comments.CreateComment("Collection sent to Officer " & myCollection.CollOfficer, Constants.GCST_COMM_CollSendToOfficer)
                            myCollection.Refresh()         'Get modified data
                            myCollection.LogAction(jobs.CMJob.Actions.SendToCollOfficer)

                        Catch cex As MedscreenExceptions.CanNotChangeCollectionStatus
                            MsgBox(cex.ToString)
                            Exit Function
                        End Try
                    End If

                Else
                    If InStr("CRDM", myCollection.Status) = 0 Then              'Job has already been confirmed or samples collected
                        myCollection.Status = MedscreenLib.Constants.GCST_JobStatusSent
                        myCollection.Update()
                    End If
                    myCollection.Update()
                    'objColl.FieldValues.Rollback()  'Remove the changes made
                End If
            Else
                MsgBox("Problem with getting officer : please retry ")
                'myCollection.Status = Constants.GCST_JobJobStatusAvailable
            End If
        End If
        MedscreenCommonGui.CommonForms.FormCollOfficer = Nothing
        myCollection.Refresh()                   'Ensure any changes have been updated

    End Function

    Public Shared Sub SendEmailToManager(ByVal objColl As Intranet.intranet.jobs.CMJob, Optional ByVal Action As String = "Cancellation")
        Dim strBody As String = objColl.CancellationMessage(False)
        MedscreenLib.Medscreen.BlatEmail("Collection Reference " & objColl.ID & " " & Action, strBody)
    End Sub

    Public Shared Sub ConfirmCollection(ByVal Collection As Intranet.intranet.jobs.CMJob)
        If Collection Is Nothing Then Exit Sub
        With FormCollOfficer
            .CollectionDate = Collection.CollectionDate
            .Port = Collection.Port
            .Collector = Collection.CollOfficer
            .FormType = MedscreenCommonGui.FrmCollOfficer.CollOffFrmType.Confirm
            .collection = Collection
            If .ShowDialog = Windows.Forms.DialogResult.OK Then       'Found who will do the collection
                Collection.CollOfficer = .Collector  'Assign collecting officer to collection 
                Collection.DateCollConfirm = Now     'Fill date confirmed 
                Dim vntDate As Object = Medscreen.GetParameter(Medscreen.MyTypes.typDate, "Collection Date", "Change Collection Date ", Collection.CollectionDate)

                'Need to add checks to see if the collection date is more than a month ago
                Dim tmpTs As TimeSpan = Today.Subtract(CDate(vntDate))

                If tmpTs.TotalDays > 30 Then
                    MsgBox("Sorry this collection date is too long ago " & CDate(vntDate).ToShortDateString & " can not proceed." & vbCrLf & _
                    "Use defer collection to change the collection date first", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
                    Exit Sub
                End If

                Collection.CollectionDate = vntDate
                If Collection.CollectionType = "C" Then Collection.DateCalloutToOfficer = Now
                Collection.Update()

                'Also need to SM to create JOB
                If Collection.ConfirmCollection() Then      'Pass to SM 
                    'Collection.Status = Constants.GCST_JobStatusConfirmed    'Update the data 
                    'Collection.Comments.CreateComment("Confirmation received from " & .Collector, MedscreenLib.Constants.GCST_COMM_CollConfirm)

                    Collection.Update()
                    Collection.LogAction(jobs.CMJob.Actions.ConfirmCollection)

                Else
                    MsgBox("Problem in confirming collection")
                End If

            End If
        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancel a collection
    ''' </summary>
    ''' <param name="Collection">Collection to be cancelled</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2006]</date><Action>Added clean exit</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CancelCollection(ByVal Collection As Intranet.intranet.jobs.CMJob) As Boolean
        Dim cancStatus As Char = Constants.GCST_CANC_NotCancelled

        Dim FR As New frmCancelReason()


        Dim cancReasons As New MedscreenLib.Glossary.PhraseCollection("CANC_COLL")
        cancReasons.Load()

        Dim ds As DataSet = cancReasons.ToDataSet   'Convert XML to a dataset to make life easier
        FR.DoBinding(ds)                            'Bind the data to the controls
        Dim lngReason As Long = 0
        Dim strReason As String = ""
        If FR.ShowDialog = Windows.Forms.DialogResult.OK Then     'Show Dialog if not cancelled then work out who is at fault 
            lngReason = FR.NumericReason            'Reason as a hex number 
            strReason = FR.ReasonString             'Reason as text 

            Dim blnCustFault As Boolean = False
            Dim lngCustFault As Long = (lngReason And CLng("&H" & "FFFFF0000")) 'AND reason with mask 

            blnCustFault = (lngCustFault > 0)       'Customer is at fault if not zero 
            Dim FS As New MedscreenCommonGui.frmStatusSelect()

            With FS
                FS.FormTypeIs = MedscreenCommonGui.frmStatusSelect.FormType.Calendar
                FS.SelectedDate = Now
                FS.Text = "Date Collection Cancelled"
                If FS.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Collection.CancellationDate = FS.SelectedDate
                Else
                    Exit Function
                End If
            End With
            Dim ts As TimeSpan = Collection.CollectionDate.Subtract(Collection.CancellationDate)  'Find out how long to go 
            If ts.TotalHours < 48 Then              'Cancellation charges may apply 
                If blnCustFault Then                'Possibly invoice 
                    If Collection.CollectorPaid Then
                        cancStatus = Constants.GCST_CANC_Ready              'We have collectors pay details
                    Else
                        cancStatus = Constants.GCST_CANC_WaitingTimeSheet   'No waiting on time sheet
                    End If
                Else
                    cancStatus = Constants.GCST_CANC_NoInvoice              'Our fault they got off
                End If
            Else
                cancStatus = Constants.GCST_CANC_NoInvoice                  'Sufficient warning 
            End If
            Collection.HexErrorCode = Medscreen.DecimalToHex(lngReason)
            Collection.Update()
            Dim blnInform As Boolean = False
            If MessageBox.Show("Inform collecting officer?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                blnInform = True
            End If
            Try
                Collection.CancelCollection(strReason, cancStatus, blnInform, _
                "Please be advised that the above collection has been cancelled.  Please destroy all relevant paperwork.")              'Cancel the collection now
            Catch ex As Intranet.Exceptions.CollectionException
                If Collection.Job.JobStatus = "X" Then
                    Collection.SetStatus("X")
                    Collection.Update()
                Else
                    MsgBox(ex.Message)
                End If
            End Try
            'Redraw workflow 
        End If

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Application path 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>This property if it is to be used should be filled in by the containing application
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property AppPath() As String
        Get
            Return myAppPath
        End Get
        Set(ByVal Value As String)
            myAppPath = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes edit collecting officer form 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property FormEdCollOfficer() As frmCollOfficerEd
        Get
            If myEdCollOfficer Is Nothing Then
                myEdCollOfficer = New frmCollOfficerEd()
            End If
            Return myEdCollOfficer
        End Get
        Set(ByVal Value As frmCollOfficerEd)
            myEdCollOfficer = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes the FindAddress Form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property FindAddress() As FindAddress
        Get
            Return myFindAddress
        End Get
        Set(ByVal Value As FindAddress)
            myFindAddress = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List view columns in use
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property ListViewColumns() As UserDefaults.ColumnDefaults
        Get
            If oCollheaders Is Nothing OrElse oCollheaders.LVCustSelectList Is Nothing Then
                oCollheaders = New UserDefaults.ColumnDefaults()
            End If
            Return oCollheaders
        End Get
        Set(ByVal Value As UserDefaults.ColumnDefaults)
            oCollheaders = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' User options for Customer Centre tool
    ''' </summary>
    ''' <value></value>
    ''' <remarks>This has been moved to this dynamic library so it can be used in common forms
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property UserOptions() As UserDefaults.UserOptions
        Get
            Return myUserOptions
        End Get
        Set(ByVal Value As UserDefaults.UserOptions)
            myUserOptions = Value
        End Set
    End Property
End Class

Namespace Comparers

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItemComparer
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Compare list view items
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ListViewItemComparer
        Implements IComparer


        Private myDirection As Integer

        Private col As Integer

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new comparer
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            col = 0
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new comparer with column and direction 
        ''' </summary>
        ''' <param name="column">Column to sort</param>
        ''' <param name="intDirection">1 ascending, -1 descending</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal column As Integer, ByVal intDirection As Integer)
            col = column
            myDirection = intDirection
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Compare handle date items differently 
        ''' </summary>
        ''' <param name="x">Row 1</param>
        ''' <param name="y">Row 2</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            Try
                If TypeOf CType(x, System.Windows.Forms.ListViewItem).SubItems(col) Is ListViewItems.SubItems.DateItem Then
                    Return [Date].Compare(CType(CType(x, ListViewItem).SubItems(col), ListViewItems.SubItems.DateItem).ItemDate, CType(CType(y, ListViewItem).SubItems(col), ListViewItems.SubItems.DateItem).ItemDate) * myDirection
                ElseIf TypeOf CType(x, System.Windows.Forms.ListViewItem).SubItems(col) Is ListViewItems.SubItems.IntegerListViewSubItem Then
                    If CType(CType(x, ListViewItem).SubItems(col), ListViewItems.SubItems.IntegerListViewSubItem).ItemInteger = CType(CType(y, ListViewItem).SubItems(col), ListViewItems.SubItems.IntegerListViewSubItem).ItemInteger Then
                        Return 0
                    ElseIf CType(CType(x, ListViewItem).SubItems(col), ListViewItems.SubItems.IntegerListViewSubItem).ItemInteger > CType(CType(y, ListViewItem).SubItems(col), ListViewItems.SubItems.IntegerListViewSubItem).ItemInteger Then
                        Return 1 * myDirection
                    Else
                        Return -1 * myDirection
                    End If

                Else
                    Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text) * myDirection
                End If
            Catch ex As Exception
                Return -1
            End Try

        End Function

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : CustSortViewItemComparer
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Custom sorter for Customer form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CustSortViewItemComparer
        Implements IComparer

        Private myDirection As Integer

        Private col As Integer

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new default comparer
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            col = 0
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new comparer 
        ''' </summary>
        ''' <param name="column">Column to sort on </param>
        ''' <param name="intDirection">Direction to sort</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal column As Integer, ByVal intDirection As Integer)
            col = column
            myDirection = intDirection
        End Sub

        '''  -----------------------------------------------------------------------------
        ''' <summary>
        ''' Do compare 
        ''' </summary>
        ''' <param name="x">First Item</param>
        ''' <param name="y">Second Item</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare
            Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text) * myDirection
        End Function

    End Class
End Namespace


Namespace ListViewItems

    ''' <summary>
    ''' Manages Customer Services
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    ''' <author></author>
    Public Class lvServiceItem
        Inherits ListViewItem


        Private myService As Intranet.intranet.customerns.CustomerService
        ''' <summary>
        ''' Expose the Service as a property
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Property Service() As Intranet.intranet.customerns.CustomerService
            Get
                Return myService
            End Get
            Set(ByVal value As Intranet.intranet.customerns.CustomerService)
                myService = value
            End Set
        End Property
        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Create New Instance of Service ListViewItem
        ''' </summary>
        ''' <param name="service"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Sub New(ByVal service As Intranet.intranet.customerns.CustomerService)
            MyBase.New()
            myService = service
            Refresh()
        End Sub

        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Refresh Item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Function Refresh() As Boolean
            Dim blnRet As Boolean = True
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            Me.SubItems.Clear()
            If myService Is Nothing Then
                blnRet = False
                Exit Function
            End If
            objPhrase = MedscreenLib.Glossary.Glossary.ServicesSold.Item(myService.Service)
            If objPhrase Is Nothing Then
                Text = myService.Service
            Else
                Text = objPhrase.PhraseText
            End If
            Dim lvSubValue As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, "")
            SubItems.Add(lvSubValue)
            lvSubValue = New ListViewItem.ListViewSubItem(Me, "")
            SubItems.Add(lvSubValue)
            Tag = myService.Service
            Dim lvStartDate As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, myService.StartDateText)
            Dim lvEndDate As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, myService.EndDateText)
            SubItems.Add(lvStartDate)
            SubItems.Add(lvEndDate)
            Dim lvAmend As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, myService.ContractAmendment)
            Dim lvOperator As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, myService.[Operator])
            SubItems.Add(lvAmend)
            SubItems.Add(lvOperator)


            Return blnRet
        End Function
    End Class
    Public Class lvRandomColumn
        Inherits ListViewItem

        Private myCol As RandomNumberColumn

        Public Sub New(ByVal Col As RandomNumberColumn, ByVal txt As String, ByVal index As Integer)
            myCol = Col
            Text = txt
            Col.Index = index
            Dim subit As New ListViewItem.ListViewSubItem(Me, Col.NumberOfDonors)
            Me.SubItems.Add(subit)
            subit = New ListViewItem.ListViewSubItem(Me, Column.From)
            Me.SubItems.Add(subit)
            subit = New ListViewItem.ListViewSubItem(Me, Column.DataImportMethod.ToString)
            Me.SubItems.Add(subit)
        End Sub

        Public Property Column() As RandomNumberColumn
            Get
                Return myCol
            End Get
            Set(ByVal value As RandomNumberColumn)
                myCol = value
            End Set
        End Property
    End Class

    Public Class lvRandomSubColumn
        Inherits ListViewItem

        Private SubCol As SubColumn

        Public Sub New(ByVal RandomSubColumn As SubColumn)
            SubCol = RandomSubColumn

            Refresh()
        End Sub

        Public Property SubColumn() As SubColumn
            Get
                Return SubCol
            End Get
            Set(ByVal value As SubColumn)
                SubCol = value
                Refresh()
            End Set
        End Property

        Public Function Refresh()
            If SubCol Is Nothing Then Exit Function
            Dim S As String() = [Enum].GetNames(GetType(MedscreenLib.SubColumn.ColumnType))

            Me.SubItems.Clear()
            Me.Text = SubCol.Header
            Dim lsc As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, S.GetValue(SubCol.ColType))
            SubItems.Add(lsc)
            lsc = New ListViewItem.ListViewSubItem(Me, SubCol.DataStart)
            SubItems.Add(lsc)

        End Function
    End Class
    ''' <summary>
    ''' Class to display an option
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
    ''' <author></author>
    Public Class LvOptionItem
        Inherits ListViewItem

#Region "Declarations"
        Private myOption As MedscreenLib.Glossary.Options
        Private objOpt As Intranet.intranet.customerns.SmClientOption
        Private objPhraseColl As MedscreenLib.Glossary.PhraseCollection
        Private objPhrase As MedscreenLib.Glossary.Phrase
        Private myOptionEditableBackColour As String = "Red"
        Private myOptionEditableForeColour As String = "White"
#End Region

        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Sub New()
            MyBase.New()
            myOptionEditableBackColour = MedscreenCommonGUIConfig.Misc("OptionEditableBackColour")
            myOptionEditableForeColour = MedscreenCommonGUIConfig.Misc("OptionEditableForeColour")
        End Sub

        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Constructor taking an Option as a parameter
        ''' </summary>
        ''' <param name="_option"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Sub New(ByVal _option As Intranet.intranet.customerns.SmClientOption)
            MyClass.New()
            objOpt = _option
            Refresh()
        End Sub

        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Deal with Range Options
        ''' </summary>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Private Sub SetupRange()


            'Dim lvSubBool As ListViewItem.ListViewSubItem
            Dim lvSubValue As ListViewItem.ListViewSubItem

            objPhrase = Nothing
            If myOption.Reference.Trim.Length = 0 Then
                objPhraseColl = MedscreenLib.Glossary.Glossary.PhraseList.Item(myOption.OptionID)
            Else
                objPhraseColl = MedscreenLib.Glossary.Glossary.PhraseList.Item(myOption.Reference)
            End If

            If objPhraseColl Is Nothing Then
                If myOption.Reference.Trim.Length = 0 Then
                    objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(myOption.OptionID)
                Else
                    objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(myOption.Reference)
                End If
                objPhraseColl.Load()
                MedscreenLib.Glossary.Glossary.PhraseList.Add(objPhraseColl) 'This is a collection of phrases
            End If
            If Not objPhraseColl Is Nothing Then
                objPhrase = objPhraseColl.Item(objOpt.Value) 'Get the phrase item
            End If
            If Not objPhrase Is Nothing Then
                lvSubValue = New ListViewItem.ListViewSubItem(Me, objPhrase.PhraseText)
                SubItems.Add(lvSubValue)
                'We can use the phrase text 
                'Dr = Me.DataSet11.Customer_options.AddCustomer_optionsRow(myoption.OptionDescription, objOpt.optionBool, objPhrase.PhraseText, objOpt.OptionId)
            ElseIf myOption.OptionType = OptionCollection.GCST_OPTION_MULTI Then
                Dim PhraseText As String = ""
                For Each phr As Phrase In objPhraseColl

                    If phr.PhraseID.Trim.Length > 0 Then  'skip blank ones added to list 
                        If InStr(objOpt.Value, phr.PhraseID) <> 0 Then
                            If PhraseText.Trim.Length > 0 Then PhraseText += ";"
                            PhraseText += phr.PhraseText
                        End If
                    End If
                Next
                lvSubValue = New ListViewItem.ListViewSubItem(Me, PhraseText)
                SubItems.Add(lvSubValue)

            Else
                lvSubValue = New ListViewItem.ListViewSubItem(Me, objOpt.Value)
                SubItems.Add(lvSubValue)
                'We'll have to make do with a value of the phrase element 
                'Dr = Me.DataSet11.Customer_options.AddCustomer_optionsRow(myoption.OptionDescription, objOpt.optionBool, objOpt.Value, objOpt.OptionId)
            End If

        End Sub

        ''' <developer>Taylor</developer>
        ''' <summary>
        ''' Redraw Option Item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' <revisionHistory>
        ''' <revision><modified>15 Apr 2011</modified><Author>Taylor</Author>Added support for Multi types</revision>
        ''' <revision>Created 20th Jan 2011</revision>
        ''' </revisionHistory>
        Public Function Refresh() As Boolean
            Dim blnRet As Boolean = False
            Me.SubItems.Clear()
            Dim lvSubBool As ListViewItem.ListViewSubItem
            Dim lvSubValue As ListViewItem.ListViewSubItem
            If myOption Is Nothing Then
                myOption = MedscreenLib.Glossary.Glossary.Options.Item(objOpt.OptionId) 'Identify type of option
                If myOption Is Nothing Then
                    myOption = New Options(objOpt.OptionId)
                End If
            End If
            Me.Text = myOption.OptionDescription
            If myOption.OptionType = OptionCollection.GCST_OPTION_RANGE OrElse myOption.OptionType = OptionCollection.GCST_OPTION_TABLE OrElse myOption.OptionType = OptionCollection.GCST_OPTION_MULTI Then    'Deal with range options, they have an associated phrase
                lvSubBool = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(lvSubBool)
                SetupRange()
            ElseIf myOption.OptionType = OptionCollection.GCST_OPTION_BOOLRANGE Then                                       'Deal with range options, they have an associated phrase
                lvSubBool = New ListViewItem.ListViewSubItem(Me, objOpt.optionBool)
                Me.SubItems.Add(lvSubBool)
                SetupRange()

            ElseIf myOption.OptionType = "VALUE" Then 'Deal with values and boolean + value
                lvSubBool = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(lvSubBool)
                lvSubValue = New ListViewItem.ListViewSubItem(Me, objOpt.Value)
                Me.SubItems.Add(lvSubValue)
            ElseIf myOption.OptionType = "BLNVL" Then
                lvSubBool = New ListViewItem.ListViewSubItem(Me, objOpt.optionBool)
                Me.SubItems.Add(lvSubBool)
                lvSubValue = New ListViewItem.ListViewSubItem(Me, objOpt.Value)
                Me.SubItems.Add(lvSubValue)

            Else 'Boolean only 
                lvSubBool = New ListViewItem.ListViewSubItem(Me, objOpt.optionBool)
                Me.SubItems.Add(lvSubBool)
                lvSubValue = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(lvSubValue)
            End If
            Me.Tag = objOpt.OptionId
            Dim lvStartDate As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, objOpt.StartDateText)
            Dim lvEndDate As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, objOpt.EndDateText)
            Me.SubItems.Add(lvStartDate)
            Me.SubItems.Add(lvEndDate)
            Dim lvAmend As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, objOpt.ContractAmendment)
            Dim lvOperator As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, objOpt.[Operator])
            Me.SubItems.Add(lvAmend)
            Me.SubItems.Add(lvOperator)
            If objOpt.optionEndDate <> DateField.ZeroDate AndAlso objOpt.optionEndDate <= Today Then
                Me.BackColor = Color.Red
                Me.ForeColor = Color.Yellow
            ElseIf objOpt.optionStartDate <> DateField.ZeroDate AndAlso objOpt.optionStartDate > Today Then
                Me.BackColor = Color.Green
                Me.ForeColor = Color.Yellow
            End If
            If myOption Is Nothing OrElse myOption.RemoveFlag Then
                Me.Font = New Font("Arial", 8, FontStyle.Strikeout)
                Me.BackColor = Color.Red
                Me.ForeColor = Color.Yellow
            ElseIf Not (myOption.Editable OrElse myOption.Expireable) Then
                Me.BackColor = Color.FromName(myOption.ColorList.FindColour("EXPIRABLE").BackColour)
                Me.ForeColor = Color.FromName(myOption.ColorList.FindColour("EXPIRABLE").ForeColour)
            End If
            Return blnRet
        End Function

        ''' <summary>
        ''' Expose Option as a property
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        ''' <revisionHistory><revision>Created 20th Jan 2011</revision></revisionHistory>
        Public Property _Option() As MedscreenLib.Glossary.Options
            Get
                Return myOption
            End Get
            Set(ByVal value As MedscreenLib.Glossary.Options)
                myOption = value
            End Set
        End Property

        Public Property CustOption() As Intranet.intranet.customerns.SmClientOption
            Get
                Return objOpt
            End Get
            Set(ByVal value As Intranet.intranet.customerns.SmClientOption)
                objOpt = value
            End Set
        End Property

    End Class

    Public Class InfoReportItem
        Inherits ListViewItem
        Private myReportItem As MedscreenLib.Glossary.InfomakerLink

        Public Sub New(ByVal report As MedscreenLib.Glossary.InfomakerLink)
            MyBase.new(report.Description)
            myReportItem = report
        End Sub

        Public Property Report() As InfomakerLink
            Set(ByVal value As InfomakerLink)
                myReportItem = value
            End Set
            Get
                Return myReportItem
            End Get

        End Property
    End Class

    Public Class InfoParameterItem
        Inherits ListViewItem
        Private myParameter As MedscreenLib.Glossary.InfomakerParameter

        Public Sub New(ByVal Param As MedscreenLib.Glossary.InfomakerParameter)
            Me.Text = Param.DESCRIPTION
            Dim asub As ListViewItem.ListViewSubItem
            asub = New ListViewItem.ListViewSubItem(Me, Param.DEFAULTVALUE)
            Me.SubItems.Add(asub)
            asub = New ListViewItem.ListViewSubItem(Me, Param.TYPE)
            Me.SubItems.Add(asub)
            asub = New ListViewItem.ListViewSubItem(Me, Param.PROMPT1)
            Me.SubItems.Add(asub)
            asub = New ListViewItem.ListViewSubItem(Me, Param.PROMPT2)
            Me.SubItems.Add(asub)

        End Sub

    End Class

    Namespace SubItems


        ''' -----------------------------------------------------------------------------
        ''' Project	 : MedscreenCommonGui
        ''' Class	 : DateItem
        ''' 
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A list view sub item of date type
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Class DateItem
            Inherits ListViewItem.ListViewSubItem
            Private myDate As Date

            ''' -----------------------------------------------------------------------------
            ''' <summary>
            ''' The items date value 
            ''' </summary>
            ''' <value></value>
            ''' <remarks>
            ''' </remarks>
            ''' <revisionHistory>
            ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
            ''' </revisionHistory>
            ''' -----------------------------------------------------------------------------
            Public Property ItemDate() As Date
                Get
                    Return myDate
                End Get
                Set(ByVal Value As Date)
                    myDate = Value
                End Set
            End Property

            ''' -----------------------------------------------------------------------------
            ''' <summary>
            ''' create a new Date item with a passed date
            ''' </summary>
            ''' <param name="inDate">Date to use</param>
            ''' <remarks>
            ''' </remarks>
            ''' <revisionHistory>
            ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
            ''' </revisionHistory>
            ''' -----------------------------------------------------------------------------
            Public Sub New(ByVal inDate As Date)
                MyClass.New(inDate, "dd-MMM-yyyy")
            End Sub

            ''' -----------------------------------------------------------------------------
            ''' <summary>
            ''' Create a new date item with a supplied date and format
            ''' </summary>
            ''' <param name="inDate">Date to use</param>
            ''' <param name="DateFormat">Format of date </param>
            ''' <remarks>
            ''' </remarks>
            ''' <revisionHistory>
            ''' <revision><Author>[taylor]</Author><date> [04/11/2005]</date><Action></Action></revision>
            ''' </revisionHistory>
            ''' -----------------------------------------------------------------------------
            Public Sub New(ByVal inDate As Date, ByVal DateFormat As String)
                MyBase.new()
                myDate = inDate
                Me.Text = myDate.ToString(DateFormat)

            End Sub
        End Class
        ''' <summary>
        ''' Currency item sub class 
        ''' </summary>
        ''' <remarks></remarks>
        ''' <revisionHistory></revisionHistory>
        ''' <author></author>
        Public Class CurrencyItem
            Inherits ListViewItem.ListViewSubItem
            Private myCurrency As String
            Private myvalue As Double
            Private myCulture As String
            Public Sub New(ByVal Value As Double, ByVal Currency As String)
                MyBase.New()
                myCurrency = Currency
                myvalue = Value
                myCulture = CConnection.PackageStringList("lib_utils.GetCulture", Currency)
                Dim FMT As New Globalization.CultureInfo(myCulture)
                Me.Text = myvalue.ToString(FMT)
            End Sub
        End Class

        ''' -----------------------------------------------------------------------------
        ''' Project	 : MedscreenCommonGui
        ''' Class	 : ListViewItems.IntegerListViewSubItem
        ''' 
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' An integer based list view sub item for sorting 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Class IntegerListViewSubItem
            Inherits ListViewItem.ListViewSubItem
            Private myInteger As Integer

            ''' -----------------------------------------------------------------------------
            ''' <summary>
            ''' Create a new listview sub item 
            ''' </summary>
            ''' <param name="inInteger">Integer supplied </param>
            ''' <remarks>
            ''' </remarks>
            ''' <revisionHistory>
            ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
            ''' </revisionHistory>
            ''' -----------------------------------------------------------------------------
            Public Sub New(ByVal inInteger As Integer)
                myInteger = inInteger
                Me.Text = myInteger.ToString
            End Sub

            ''' -----------------------------------------------------------------------------
            ''' <summary>
            ''' Integer associated with this item 
            ''' </summary>
            ''' <value></value>
            ''' <remarks>
            ''' </remarks>
            ''' <revisionHistory>
            ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
            ''' </revisionHistory>
            ''' -----------------------------------------------------------------------------
            Public Property ItemInteger() As Integer
                Get
                    Return myInteger
                End Get
                Set(ByVal Value As Integer)
                    myInteger = Value
                End Set
            End Property
        End Class
    End Namespace

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvMessageItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' A list view item to deal with messages held in the SM database
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvMessageItem
        Inherits Windows.Forms.ListViewItem

        Private myMessage As MedscreenLib.Messaging.Messaging

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a new Messaging List view item
        ''' </summary>
        ''' <param name="Message">Message to add</param>
        ''' <remarks>Calls refresh to construct sub items
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Message As MedscreenLib.Messaging.Messaging)
            MyBase.new()
            myMessage = Message
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Message object associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Message() As MedscreenLib.Messaging.Messaging
            Get
                Return myMessage
            End Get
            Set(ByVal Value As MedscreenLib.Messaging.Messaging)
                myMessage = Value
                Refresh()
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' redraw the contents of this item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()                                 'Clear all items
            Me.Text = myMessage.MessageId                       'Set text property to Message ID
            Dim aSub As ListViewSubItem

            aSub = New ListViewSubItem(Me, myMessage.[Operator])  'Add operator
            Me.SubItems.Add(aSub)

            aSub = New ListViewSubItem(Me, myMessage.Read)      'Add Read flag
            Me.SubItems.Add(aSub)


            aSub = New SubItems.DateItem(myMessage.Expires)              'Set expires item 
            Me.SubItems.Add(aSub)

            aSub = New SubItems.DateItem(myMessage.Sent)                 'Set sent date 
            Me.SubItems.Add(aSub)

            If Not (myMessage.MessageType = MedscreenLib.Messaging.Messaging.GCST_MessageTypeSystem _
                Or myMessage.MessageType = MedscreenLib.Messaging.Messaging.GCST_MessageTypeAttachment) Then
                aSub = New ListViewSubItem(Me, myMessage.MessageText.Trim.Length)
                Me.SubItems.Add(aSub)
            End If

        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvInstrument
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Instrument display item
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	08/07/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class lvInstrument
        Inherits ListViewItem
        Private myInstrument As Intranet.INSTRUMENT.Instrument

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' base constructor
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	08/07/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' constructor with instrument 
        ''' </summary>
        ''' <param name="inst"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	08/07/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal inst As Intranet.INSTRUMENT.Instrument)
            MyClass.New()
            myInstrument = inst
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw the item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	08/07/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.Text = ""
            Me.SubItems.Clear()
            If myInstrument Is Nothing Then Exit Function
            Me.Text = myInstrument.Model
            Dim objSub As New ListViewItem.ListViewSubItem(Me, myInstrument.SerialNo)
            Me.SubItems.Add(objSub)
            objSub = New ListViewItem.ListViewSubItem(Me, myInstrument.ServiceDate.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(objSub)
            objSub = New ListViewItem.ListViewSubItem(Me, myInstrument.CalibDate.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(objSub)
            objSub = New ListViewItem.ListViewSubItem(Me, myInstrument.Removeflag)
            Me.SubItems.Add(objSub)

        End Function
    End Class

    Public Class lvInvoice
        Inherits ListViewItem
        Private myInvoice As Intranet.intranet.jobs.Invoice

        Public Sub New()
            MyBase.new()

        End Sub

        Public Sub New(ByVal invoice As Intranet.intranet.jobs.Invoice)
            MyClass.new()
            myInvoice = invoice
            Refresh()
        End Sub

        Public Sub New(ByVal invoiceId As String)
            MyClass.new()
            myInvoice = New Intranet.intranet.jobs.Invoice(invoiceId)
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Me.Text = ""
            Me.SubItems.Clear()
            If myInvoice Is Nothing Then Exit Function
            Me.Text = myInvoice.InvoiceID
            'Do sub items
            Dim invDate As SubItems.DateItem = New SubItems.DateItem(myInvoice.InvoiceDate)
            Me.SubItems.Add(invDate)
            Dim myValue As SubItems.CurrencyItem = New SubItems.CurrencyItem(myInvoice.InvoiceAmount, myInvoice.InvoiceCurrency)
            Me.SubItems.Add(myValue)
            If myInvoice.DispatchDue > DateField.ZeroDate Then
                Dim DevDate As SubItems.DateItem = New SubItems.DateItem(myInvoice.DispatchDue)
                Me.SubItems.Add(DevDate)
            End If
        End Function

        Public Property Invoice() As Intranet.intranet.jobs.Invoice
            Get
                Return myInvoice
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.Invoice)
                myInvoice = Value
            End Set
        End Property
    End Class

    Public Class lvInvoiceItem
        Inherits ListViewItem

        Private myInvoiceItem As Intranet.intranet.jobs.InvoiceItem

        Public Sub New(ByVal InvItem As Intranet.intranet.jobs.InvoiceItem)
            MyBase.New()
            myInvoiceItem = InvItem
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = myInvoiceItem.Quantity
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myInvoiceItem.LineItem.Description))

        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvIsoportItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A ISOPort List item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action>Added support for ISO port info</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [17/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvIsoportItem
        Inherits ListViewItem
        Private myName As String = ""
        Private myId As String
        Private intMyCountryId As Integer = -1
        Private UNLocode As String = ""
        Private strISOInfo As String = ""
        Private strDivision As String = ""
        Private strFunction As String = ""
        Private strCoordinates As String = ""

        Public Sub New(ByVal ID As String, ByVal name As String)
            MyBase.New()
            myId = ID
            myName = name
            Refresh()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Port Name 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Overloads Property Name() As String
            Get
                Return Me.myName
            End Get
            Set(ByVal Value As String)
                myName = Value
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ID() As String
            Get
                Return Me.myId
            End Get
            Set(ByVal Value As String)
                myId = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Country of the Port 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property CountryId() As Integer
            Get
                Return Me.intMyCountryId
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw the listview item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [17/12/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = myId
            Dim objCountry As MedscreenLib.Address.Country = _
            MedscreenLib.Glossary.Glossary.Countries.Item(Mid(myId, 1, 2), _
            MedscreenLib.Address.CountryCollection.indexBy.Code)
            If Not objCountry Is Nothing Then
                Me.intMyCountryId = objCountry.CountryId
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, objCountry.CountryName))
            Else
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            End If
            Me.GetISOInfo()
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myName))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myId))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, Me.strDivision))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, Me.strCoordinates))
            Me.BackColor = System.Drawing.Color.Red
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get information about this port from the UN database.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function GetISOInfo() As String
            Dim retString As String = ""

            Try

                If Me.myId.Trim.Length > 0 Then


                    If Me.strISOInfo.Trim.Length = 0 Then
                        'need to get info, use pl/sql code indicated in config file
                        Dim strIsoinfo As String = MedscreenCommonGUIConfig.NodePLSQL.Item("GetIsoInfo")
                        retString = CConnection.PackageStringList(strIsoinfo, Me.myId)
                        If retString.Trim.Length > 0 Then
                            Dim strArray As String() = retString.Split(New Char() {"|"})
                            Me.strDivision = strArray.GetValue(2)
                            Me.strFunction = strArray.GetValue(3)
                            Me.strCoordinates = strArray.GetValue(4)
                            If Me.strCoordinates.Trim.Length > 0 Then
                                'SplitCoordinates 
                                Dim strCoords As String() = Me.strCoordinates.Split(New Char() {" "})
                                Dim strACoord As String
                                If strCoords.Length = 2 Then
                                    strACoord = strCoords.GetValue(0)
                                    strACoord = strACoord.Trim
                                    Dim intLength As Integer
                                    intLength = strACoord.Length
                                    If intLength = 6 Then
                                        Me.strCoordinates = strACoord.Chars(5) & strACoord.Chars(0) & strACoord.Chars(1) & _
                                            strACoord.Chars(2) & ":" & strACoord.Chars(3) & strACoord.Chars(4) & ":00"
                                    Else
                                        Me.strCoordinates = strACoord.Chars(4) & strACoord.Chars(0) & strACoord.Chars(1) & _
                                             ":" & strACoord.Chars(2) & strACoord.Chars(3) & ":00"

                                    End If

                                    strACoord = strCoords.GetValue(1)
                                    strACoord = strACoord.Trim

                                    intLength = strACoord.Length
                                    If intLength = 6 Then
                                        Me.strCoordinates += "," & strACoord.Chars(5) & strACoord.Chars(0) & strACoord.Chars(1) & _
                                            strACoord.Chars(2) & ":" & strACoord.Chars(3) & strACoord.Chars(4) & ":00"
                                    Else
                                        Me.strCoordinates += "," & strACoord.Chars(4) & strACoord.Chars(0) & strACoord.Chars(1) & _
                                             ":" & strACoord.Chars(2) & strACoord.Chars(3) & ":00"

                                    End If


                                End If
                            End If
                            strIsoinfo = retString
                        End If
                    End If
                End If
            Catch ex As System.Exception
                ' Handle exception here
                'MedscreenLib.Medscreen.LogError(ex, , "Exception log message")
            End Try
            Return retString
        End Function

        ''' <summary>
        ''' Created by taylor on ANDREW at 1/25/2007 14:27:18
        '''     
        ''' </summary>
        ''' <returns>
        '''     A System.String value...
        ''' </returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>taylor</Author><date> 1/25/2007</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' 
        Public Function PortFunction() As String
            GetISOInfo()
            Return strFunction
        End Function

        Public Function PortCoordinate() As String
            GetISOInfo()
            Return Me.strCoordinates
        End Function

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvOfficerPortItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An officer port display item
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvOfficerPortItem
        Inherits ListViewItem
        Public Enum OfficerPortDisplayMode
            OfficerMode
            PortMode
        End Enum

        Private myDisplayMode As OfficerPortDisplayMode = OfficerPortDisplayMode.PortMode

        Public Property DisplayMode() As OfficerPortDisplayMode
            Get
                Return Me.myDisplayMode

            End Get
            Set(ByVal Value As OfficerPortDisplayMode)
                myDisplayMode = Value
                Me.Refresh()
            End Set
        End Property
        Private myPort As Intranet.intranet.jobs.CollPort

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Port and Officer Link 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OfficerPort() As Intranet.intranet.jobs.CollPort
            Get
                Return myPort
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollPort)
                myPort = Value
            End Set
        End Property

        Public Sub New(ByVal port As Intranet.intranet.jobs.CollPort)
            MyBase.New()
            myPort = port
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw port 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action>deal with added enumeration</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            If Me.myPort.Port Is Nothing Then
                If Me.myDisplayMode = OfficerPortDisplayMode.PortMode Then
                    Me.Text = myPort.PortId
                Else
                    Me.Text = myPort.OfficerId
                End If
                Me.BackColor = Color.Red
                Me.ForeColor = Color.Yellow
                Me.myPort.Comments = "Invalid Port - Does not exist in port table"
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Else
                If Me.myDisplayMode = OfficerPortDisplayMode.PortMode Then
                    Me.Text = myPort.Port.SiteName
                Else
                    If Not myPort.Officer Is Nothing Then Me.Text = myPort.Officer.Name
                End If
                If myPort.Port.CountryId = 44 Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Port.PortRegion))
                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Port.CountryString))
                End If
            End If

            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.UNLocode))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.PortName))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Agreed))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Costs))
            If Not myPort.Officer Is Nothing Then
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Officer.PayCurrency))
            Else
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            End If
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Comments))
            Me.Checked = Not myPort.Removed
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.NoCollections))
            If myPort.NoDonors > 0 Then Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, CDbl((myPort.TotalPay + myPort.TotalExpenses) / myPort.NoDonors).ToString("0.00")))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.isoCoordinates))
            If myPort.Removed Then
                Me.BackColor = Color.Yellow
                Me.ForeColor = Color.Black
            End If

        End Function

    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvPortItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A listview item specialised for the handling of ports
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvPortItem
        Inherits ListViewItem       'Inherit from list view item so has all the same methods and properties

        Private myPort As Intranet.intranet.jobs.Port

        '''<summary>
        ''' Port which this node describes <see cref="Intranet.intranet.jobs.Port" />
        ''' </summary>
        Public Property Port() As Intranet.intranet.jobs.Port
            Get
                Return myPort
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.Port)
                myPort = Value                          'Set local variable 
                Refresh()                               'Call draw method
            End Set
        End Property

        '''<summary>
        ''' Create a new PortTreeNode
        ''' </summary>
        ''' <param name='port'>Port to create node for </param>
        Public Sub New(ByVal port As Intranet.intranet.jobs.Port)
            MyBase.New()
            myPort = port                               'Initialise port object
            Refresh()                                   'Draw port object 
        End Sub

        '''<summary>
        ''' Refresh the display of this node
        ''' </summary>
        ''' <returns>TRUE if succesful</returns>
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()                         'Remove all items
            Me.Text = myPort.PortId                     'Set text property to port id 
            'Text property shows in all modes of the list view, sub items only show in the detail view.
            'Create subitems for all other items in the same order as their columns
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.PortRegion))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.SiteCharge))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.PortDescription))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.UNLocode))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.SubDivision))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myPort.Coordinates))

            'Now tart it up depending on the item type
            If myPort.PortType = "P" Then                       'Port
                Me.BackColor = System.Drawing.Color.Yellow
            ElseIf myPort.PortType = "S" Then                   'Site
                Me.BackColor = System.Drawing.Color.Yellow
            ElseIf myPort.PortType = "R" Then                   'Region
                Me.BackColor = System.Drawing.Color.Wheat
            ElseIf myPort.PortType = "C" Then                   'Country
                Me.BackColor = System.Drawing.Color.LightGreen
            End If
            'check for port being removed (removed ports will be displayed differently)
            If myPort.Removed Then
                Me.BackColor = System.Drawing.Color.Red
                Me.ForeColor = System.Drawing.Color.Yellow
                Me.Text += " (removed)"
            End If
        End Function
    End Class



    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.ExpenseTotal
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collecting officer expense Total
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class ExpenseTotal
        Inherits ExpenseDescription
        Private myTotal As Double
        Private myCulture As Globalization.CultureInfo

        Public Sub New(ByVal Total As Double, ByVal culture As Globalization.CultureInfo)
            MyBase.New()
            myTotal = Total
            myCulture = culture
            Me.Refresh()
        End Sub

        Public Shadows Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = "Total"

            'If myExpense.Removed Then Exit Function

            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Total"))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myTotal.ToString("c", myCulture)))

        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.ExpenseDescription
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collecting officer expense item
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class ExpenseDescription
        Inherits ListViewItem
        Private myExpense As Intranet.collectorpay.CollPayItem
        Private objCollOfficer As Intranet.intranet.jobs.CollectingOfficer
        'Private myLineItemDescription As Intranet.collectorpay.ExpenseLineItem

        Public Property Expense() As Intranet.collectorpay.CollPayItem
            Get
                Return myExpense
            End Get
            Set(ByVal Value As Intranet.collectorpay.CollPayItem)
                myExpense = Value
            End Set
        End Property

        Public ReadOnly Property CollOfficer() As Intranet.intranet.jobs.CollectingOfficer
            Get
                Return objCollOfficer
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = myExpense.ExpenseTypeID

            If myExpense.Removed Then Exit Function

            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.Quantity))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.ItemCodeDescription))
            If Not objCollOfficer Is Nothing Then
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.ItemPrice.ToString("0.00")))
                If myExpense.ItemDiscount > 0 Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, (1 - myExpense.ItemDiscount) * 100))
                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                End If
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, (myExpense.ItemValue / myExpense.Quantity).ToString("c", objCollOfficer.Culture)))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.ItemValue.ToString("c", objCollOfficer.Culture)))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.PayStatusText))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.Comments))
            Else
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.ItemPrice.ToString("c", objCollOfficer.Culture)))
                If myExpense.ItemDiscount > 0 Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.ItemDiscount * 100))
                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                End If
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, (myExpense.ItemValue / myExpense.Quantity)))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.PayStatusText))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myExpense.Comments))

            End If

            'deal with removed items and colouring
            If myExpense.Removed Then
                Me.BackColor = Color.Red
                Me.ForeColor = Color.Yellow
            Else
                If myExpense.PayStatus = Intranet.collectorpay.CollPayItem.GCST_Pay_Entered Then
                    Me.BackColor = Color.Wheat
                ElseIf myExpense.PayStatus = Intranet.collectorpay.CollPayItem.GCST_Pay_Locked Then
                    Me.BackColor = Color.PowderBlue
                ElseIf myExpense.PayStatus = Intranet.collectorpay.CollPayItem.GCST_Pay_Declined Then
                    Me.BackColor = Color.Yellow
                ElseIf myExpense.PayStatus = Intranet.collectorpay.CollPayItem.GCST_Pay_Approved Then
                    Me.BackColor = Color.LimeGreen
                End If
            End If

        End Function

        Private myCollPayEntry As Intranet.collectorpay.CollPayItem

        Public Property MonthPaid() As Intranet.collectorpay.CollPayItem
            Get
                Return myCollPayEntry
            End Get
            Set(ByVal Value As Intranet.collectorpay.CollPayItem)
                myCollPayEntry = Value
            End Set
        End Property

        Public Sub New(ByVal Expense As Intranet.collectorpay.CollPayItem, ByVal CollOfficer As Intranet.intranet.jobs.CollectingOfficer)
            MyBase.New()
            myExpense = Expense
            objCollOfficer = CollOfficer
            Refresh()
        End Sub

        Public Sub New()

        End Sub
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.LVCustomerContact
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class LVCustomerContact
        Inherits ListViewItem
#Region "Declarations"
        Private myContact As MedscreenLib.contacts.CustomerContact

#End Region
#Region "Public Instance"

#Region "Public Procedures"

        Public Sub New(ByVal CustomerContact As contacts.CustomerContact)
            MyBase.New()
            Me.myContact = CustomerContact
            Me.Refresh()
        End Sub
#End Region
#Region "Public Properties"
        Public Property Contact() As contacts.CustomerContact
            Get
                Return Me.myContact
            End Get
            Set(ByVal Value As contacts.CustomerContact)
                myContact = Value
            End Set
        End Property


#End Region
#Region "Public Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the contents of the control
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	11/07/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = ""
            If Me.myContact Is Nothing Then
                Exit Function
            End If
            Me.Text = Me.myContact.Contact
            If Me.Text = "" Then Me.Text = Me.myContact.ContactId

            Dim objSubItem As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.AddressId)
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.ContactTypeDescription)
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.ContactStatusDescription)
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.ContactMethodDescription)
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.ReportFormat)
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myContact.ReportOptionDescription)
            Me.SubItems.Add(objSubItem)
            If myContact.ContactTypePhrase Is Nothing Then Exit Function
            If myContact.ContactTypePhrase.Formatter Is Nothing Then Exit Function
            myContact.ContactTypePhrase.Formatter.Decorate(Me)
        End Function


#End Region
#End Region
    End Class

    Public Class LVCustBasicSelectItem
        Inherits ListViewItem

        Private myClientId As String
        Private myClientData As String()
        Private myRemoved As Boolean = False
        Private myAccStatus As String = ""
        Private mySMID As String = ""
        Private myProfile As String = ""
        Private myName As String = ""
        Private myPostCodes As String = ""
        Private myContactList As String = ""
        Private myCollectionTypeDescription As String = ""
        Private myIndustryCode As String = ""
        Private myProgrammeManager As String = ""
        Private myPinNumber As String = ""
        Private myCreditCard_pay As Boolean = False
        Private myComments As String = ""
        Private myCustomerAccountType As String = ""
        Private myAccountCollectionType As String = ""
        Private mySageID As String = ""
        Private myClient As Intranet.intranet.customerns.Client


        Private Sub GetInfo()
            Dim strRet As String = CConnection.PackageStringList("lib_customer2.ClientInfo", myClientId)
            If Not strRet Is Nothing AndAlso strRet.Trim.Length > 0 Then
                myClientData = strRet.Split(New Char() {"|"})
                If myClientData.Length > 0 Then mySMID = myClientData(0)
                If myClientData.Length > 1 Then myProfile = myClientData(1)
                If myClientData.Length > 2 Then myName = myClientData(2)
                If myClientData.Length > 3 Then mySageID = myClientData(3)
                If myClientData.Length > 4 Then myRemoved = (myClientData(4) = "T")
                If myClientData.Length > 5 Then myIndustryCode = myClientData(5)
                If myClientData.Length > 6 Then myPinNumber = myClientData(6)
                If myClientData.Length > 7 Then myAccountCollectionType = myClientData(7)
                If myClientData.Length > 8 Then myCreditCard_pay = (myClientData(8) = "T")
                If myClientData.Length > 9 Then myCustomerAccountType = myClientData(9)
                If myClientData.Length > 10 Then myComments = myClientData(10)
                'If myClientData.Length > 10 Then myCustomerAccountType = myClientData(10)
                If myClientData.Length > 12 Then myAccStatus = myClientData(12)
                If myClientData.Length > 13 Then myProgrammeManager = myClientData(13)
                If myClientData.Length > 14 Then myPostCodes = myClientData(14)
                If myClientData.Length > 15 Then myCollectionTypeDescription = myClientData(15)
                If myClientData.Length > 16 Then myContactList = myClientData(16)

            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new entity
        ''' </summary>
        ''' <param name="Client">Client to pass in </param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Clientid As String)
            MyBase.New()
            myClientId = Clientid
            GetInfo()
            DrawItem()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw this list view item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function DrawItem() As Boolean
            Dim blnRet As Boolean = False
            Me.UseItemStyleForSubItems = True
            Try
                Me.SubItems.Clear()
                Me.BackColor = Drawing.Color.FromArgb(255, 252, 250, 255)   'Set item back colour
                Select Case myAccountCollectionType                 'Set fore colour depending on account type 
                    Case "ANALONLY"
                        Me.ForeColor = Drawing.Color.Red
                    Case "COLLCONPM", "COLLCOPM"
                        Me.ForeColor = Drawing.Color.DarkBlue
                    Case "COLLNPM", "COLLCONPM"
                        Me.ForeColor = Drawing.Color.DarkGreen
                    Case "COLLCONPM", "COLLCOPM"
                        Me.ForeColor = Drawing.Color.DarkSlateBlue
                End Select

                If myRemoved Then
                    Me.BackColor = Drawing.Color.Black
                    Me.ForeColor = Drawing.Color.Yellow
                End If
                Dim oColInfo As UserDefaults.ColumnInfo

                If Not MedscreenCommonGui.CommonForms.UserOptions Is Nothing Then
                    MedscreenCommonGui.CommonForms.UserOptions.CustSelectColumns.Sort()
                    Me.Text = myClientId
                    If myAccStatus = "O" Then
                        Me.Font = New Font("Arial", 8.25, FontStyle.Strikeout Or FontStyle.Italic, GraphicsUnit.Point, Font.GdiCharSet)
                    End If
                    For Each oColInfo In MedscreenCommonGui.CommonForms.UserOptions.CustSelectColumns
                        If oColInfo.Show Then
                            Select Case oColInfo.FieldName
                                Case "SMID"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, mySMID)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "PROFILE"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myProfile)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "COMPANY_NAME"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myName)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "SAGE_ID"
                                    Dim Cust As ListViewItem.ListViewSubItem = Nothing
                                    Cust = New ListViewItem.ListViewSubItem(Me, mySageID)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "ADDRESS6"
                                    Dim Cust As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, myPostCodes)

                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                    Me.SubItems.Add(Cust)

                                    'Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myContactList)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                    Me.SubItems.Add(Cust)
                                Case "COLLECTIONTYPE"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myCollectionTypeDescription)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "INDUSTRY"

                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myIndustryCode)

                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                    Me.SubItems.Add(Cust)
                                Case "PMAN"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myProgrammeManager)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor

                                Case "PIN"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myPinNumber)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor

                                Case "CCard"

                                    Dim Cust As ListViewItem.ListViewSubItem
                                    If myCreditCard_pay Then
                                        Cust = New ListViewItem.ListViewSubItem(Me, "Yes")
                                    Else
                                        Cust = New ListViewItem.ListViewSubItem(Me, "-")
                                    End If
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor

                                Case "COMMENTS"

                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myComments)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor

                                Case "CUSTACC"

                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, myCustomerAccountType)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = Me.ForeColor
                                    Cust.BackColor = Me.BackColor
                                Case "STATUS"
                                    Dim cust As ListViewItem.ListViewSubItem
                                    Dim oColl As New Collection()
                                    oColl.Add("CUST_STAT")
                                    oColl.Add(myAccStatus)
                                    Dim statusString As String = CConnection.PackageStringList("Lib_utils.DecodePhrase", oColl)
                                    Dim strColorList As String = MedscreenCommonGUIConfig.NodePLSQL("StatusColourList")
                                    Dim strColourElements As String() = strColorList.Split(New Char() {"|"})
                                    Dim strStatusColours As String() = strColourElements(1).Split(New Char() {","})
                                    Dim strStatusElements As String()
                                    Dim strStatuses As String = strColourElements.GetValue(0)
                                    strStatusElements = strStatuses.Split(New Char() {","})

                                    cust = New ListViewItem.ListViewSubItem(Me, statusString)
                                    Me.SubItems.Add(cust)

                                    Dim intPos As Integer
                                    For intPos = 0 To strStatusElements.Length - 1
                                        Dim currentStatus As String = strStatusElements(intPos)
                                        If myAccStatus = currentStatus Then
                                            cust.BackColor = Color.FromName(strStatusColours(intPos))
                                            'Me.ImageIndex = CInt(strIconElements.GetValue(intPos))
                                            Exit For
                                        End If
                                    Next
                                    Me.UseItemStyleForSubItems = False
                            End Select
                        End If

                    Next

                End If
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try
            'do extras for history item
            Return blnRet
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Client associated with this list view item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Client() As Intranet.intranet.customerns.Client
            Get
                If Client Is Nothing AndAlso myClientId.Trim.Length > 0 Then
                    myClient = New Intranet.intranet.customerns.Client(myClientId)
                End If
                Return myClient
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.Client)
                myClient = Value
                If Not Value Is Nothing Then
                    DrawItem()

                End If
            End Set
        End Property
    End Class



    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.LVCustSelectItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A list view item for Customer Selection screens
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LVCustSelectItem
        Inherits ListViewItem
        Private myClientbase As Intranet.intranet.customerns.ClientBase
        Private myClient As Intranet.intranet.customerns.Client
        Private myClientHistory As Intranet.intranet.customerns.ClientHistory

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new entity
        ''' </summary>
        ''' <param name="Client">Client to pass in </param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Client As Intranet.intranet.customerns.Client)
            MyBase.New()
            myClientbase = Client
            myClient = Client
            DrawItem()
        End Sub

        Public Sub New(ByVal Client As Intranet.intranet.customerns.ClientHistory)
            MyBase.New()
            myClientbase = Client
            myClientHistory = Client
            DrawItem()
        End Sub
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw this list view item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function DrawItem() As Boolean
            Dim blnRet As Boolean = False
            Me.UseItemStyleForSubItems = True
            Try
                Me.SubItems.Clear()
                If TypeOf myClientbase Is Intranet.intranet.customerns.Client Then
                    Me.BackColor = Drawing.Color.FromArgb(255, 252, 250, 255)   'Set item back colour
                    Select Case myClient.AccountCollectionType                 'Set fore colour depending on account type 
                        Case "ANALONLY"
                            Me.ForeColor = Drawing.Color.Red
                        Case "COLLCONPM", "COLLCOPM"
                            Me.ForeColor = Drawing.Color.DarkBlue
                        Case "COLLNPM", "COLLCONPM"
                            Me.ForeColor = Drawing.Color.DarkGreen
                        Case "COLLCONPM", "COLLCOPM"
                            Me.ForeColor = Drawing.Color.DarkSlateBlue
                    End Select
                End If

                If myClientbase.Removed Then
                    Me.BackColor = Drawing.Color.Black
                    Me.ForeColor = Drawing.Color.Yellow
                End If
                Dim oColInfo As UserDefaults.ColumnInfo

                If Not MedscreenCommonGui.CommonForms.UserOptions Is Nothing Then
                    MedscreenCommonGui.CommonForms.UserOptions.CustSelectColumns.Sort()
                    If Not Me.myClientbase Is Nothing Then
                        If TypeOf myClientbase Is Client Then
                            Me.Text = myClientbase.Identity
                        ElseIf Not Me.myClientHistory Is Nothing Then
                            Me.Text = Me.myClientHistory.AuditDate.ToString("dd-MMM-yyyy HH:mm")
                        End If
                        With Me.myClientbase
                            If .AccStatus = "O" Then
                                Me.Font = New Font("Arial", 8.25, FontStyle.Strikeout Or FontStyle.Italic, GraphicsUnit.Point, Font.GdiCharSet)
                            End If
                            For Each oColInfo In MedscreenCommonGui.CommonForms.UserOptions.CustSelectColumns
                                If oColInfo.Show Then
                                    Select Case oColInfo.FieldName
                                        Case "SMID"
                                            Dim Cust As ListViewItem.ListViewSubItem
                                            Cust = New ListViewItem.ListViewSubItem(Me, .SMID)
                                            Me.SubItems.Add(Cust)
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor
                                        Case "PROFILE"
                                            Dim Cust As ListViewItem.ListViewSubItem
                                            Cust = New ListViewItem.ListViewSubItem(Me, .Profile)
                                            Me.SubItems.Add(Cust)
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor
                                        Case "COMPANY_NAME"
                                            Dim Cust As ListViewItem.ListViewSubItem
                                            Cust = New ListViewItem.ListViewSubItem(Me, .Name)
                                            Me.SubItems.Add(Cust)
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor
                                        Case "SAGE_ID"
                                            Dim Cust As ListViewItem.ListViewSubItem = Nothing
                                            If Not myClient Is Nothing Then
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.InvoiceInfo.SageID)
                                                Me.SubItems.Add(Cust)
                                            End If
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor
                                        Case "ADDRESS6"
                                            If Not myClient Is Nothing Then
                                                Dim Cust As ListViewItem.ListViewSubItem = New ListViewItem.ListViewSubItem(Me, Me.myClient.PostCodes)
                                                'If myClient.MainAddress Is Nothing Then
                                                '    Cust = New ListViewItem.ListViewSubItem(Me, "")
                                                'Else
                                                '    Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.MainAddress.PostCode)
                                                'End If
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor
                                                Me.SubItems.Add(Cust)
                                            End If
                                            If Not myClient Is Nothing Then

                                                Dim Cust As ListViewItem.ListViewSubItem
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.ContactList)
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor
                                                Me.SubItems.Add(Cust)
                                            End If
                                        Case "COLLECTIONTYPE"
                                            If Not myClient Is Nothing Then
                                                Dim Cust As ListViewItem.ListViewSubItem
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.CollectionTypeDescription)
                                                Me.SubItems.Add(Cust)
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor
                                            End If
                                        Case "INDUSTRY"
                                            If Not myClient Is Nothing Then

                                                Dim Cust As ListViewItem.ListViewSubItem
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.IndustryCode)

                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor
                                                Me.SubItems.Add(Cust)
                                            End If
                                        Case "PMAN"
                                            If Not myClient Is Nothing Then
                                                Dim Cust As ListViewItem.ListViewSubItem
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.ProgrammeManager)
                                                Me.SubItems.Add(Cust)
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor

                                            End If
                                        Case "PIN"
                                            If Not myClient Is Nothing Then
                                                Dim Cust As ListViewItem.ListViewSubItem
                                                Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.PinNumber)
                                                Me.SubItems.Add(Cust)
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor

                                            End If
                                        Case "CCard"
                                            If Not myClient Is Nothing Then

                                                Dim Cust As ListViewItem.ListViewSubItem
                                                If Me.myClient.InvoiceInfo.CreditCard_pay Then
                                                    Cust = New ListViewItem.ListViewSubItem(Me, "Yes")
                                                Else
                                                    Cust = New ListViewItem.ListViewSubItem(Me, "-")
                                                End If
                                                Me.SubItems.Add(Cust)
                                                Cust.ForeColor = Me.ForeColor
                                                Cust.BackColor = Me.BackColor
                                            End If

                                        Case "COMMENTS"

                                            Dim Cust As ListViewItem.ListViewSubItem
                                            Cust = New ListViewItem.ListViewSubItem(Me, .Comments)
                                            Me.SubItems.Add(Cust)
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor

                                        Case "CUSTACC"

                                            Dim Cust As ListViewItem.ListViewSubItem
                                            Cust = New ListViewItem.ListViewSubItem(Me, .CustomerAccountType)
                                            Me.SubItems.Add(Cust)
                                            Cust.ForeColor = Me.ForeColor
                                            Cust.BackColor = Me.BackColor
                                        Case "STATUS"
                                            Dim cust As ListViewItem.ListViewSubItem
                                            Dim oColl As New Collection()
                                            oColl.Add("CUST_STAT")
                                            oColl.Add(.AccStatus)
                                            Dim statusString As String = CConnection.PackageStringList("Lib_utils.DecodePhrase", oColl)
                                            Dim strColorList As String = MedscreenCommonGUIConfig.NodePLSQL("StatusColourList")
                                            Dim strColourElements As String() = strColorList.Split(New Char() {"|"})
                                            Dim strStatusColours As String() = strColourElements(1).Split(New Char() {","})
                                            Dim strStatusElements As String()
                                            Dim strStatuses As String = strColourElements.GetValue(0)
                                            strStatusElements = strStatuses.Split(New Char() {","})

                                            cust = New ListViewItem.ListViewSubItem(Me, statusString)
                                            Me.SubItems.Add(cust)

                                            Dim intPos As Integer
                                            For intPos = 0 To strStatusElements.Length - 1
                                                Dim currentStatus As String = strStatusElements(intPos)
                                                If .AccStatus = currentStatus Then
                                                    cust.BackColor = Color.FromName(strStatusColours(intPos))
                                                    'Me.ImageIndex = CInt(strIconElements.GetValue(intPos))
                                                    Exit For
                                                End If
                                            Next
                                            Me.UseItemStyleForSubItems = False
                                    End Select
                                End If

                            Next

                        End With
                    End If
                Else
                    If Not Me.myClientbase Is Nothing Then
                        If TypeOf myClientbase Is Client Then
                            Me.Text = myClientbase.Identity
                        ElseIf Not Me.myClientHistory Is Nothing Then
                            Me.Text = Me.myClientHistory.AuditDate.ToString("dd-MMM-yyyy HH:mm")
                        End If

                        Dim Cust As ListViewItem.ListViewSubItem
                        Cust = New ListViewItem.ListViewSubItem(Me, Me.myClientbase.SMID)
                        Me.SubItems.Add(Cust)

                        Cust = New ListViewItem.ListViewSubItem(Me, Me.myClientbase.Profile)
                        Me.SubItems.Add(Cust)

                        Cust = New ListViewItem.ListViewSubItem(Me, Me.myClientbase.Name)
                        Me.SubItems.Add(Cust)
                        If Not Me.myClient Is Nothing Then
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.InvoiceInfo.SageID)
                            Me.SubItems.Add(Cust)
                            'If Not Me.myClient.MainAddress Is Nothing Then
                            '    Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.MainAddress.PostCode)
                            'Else
                            '    Cust = New ListViewItem.ListViewSubItem(Me, "")

                            'End If
                            Me.SubItems.Add(Cust)

                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.Contact)
                            Me.SubItems.Add(Cust)
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.CollectionTypeDescription)
                            Me.SubItems.Add(Cust)
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.IndustryCode)

                            Me.SubItems.Add(Cust)

                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.ProgrammeManager)
                            Me.SubItems.Add(Cust)

                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.PinNumber)
                            Me.SubItems.Add(Cust)
                            If Me.myClient.InvoiceInfo.CreditCard_pay Then
                                Cust = New ListViewItem.ListViewSubItem(Me, "Yes")
                            Else
                                Cust = New ListViewItem.ListViewSubItem(Me, "-")
                            End If
                            Me.SubItems.Add(Cust)
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.Comments)
                            Me.SubItems.Add(Cust)
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.myClient.CustomerAccountType)
                            Me.SubItems.Add(Cust)
                        End If
                    End If

                End If
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try
            'do extras for history item
            If (TypeOf Me.myClientbase Is ClientHistory) AndAlso (Not Me.myClientHistory Is Nothing) Then
                Dim hCust As ListViewItem.ListViewSubItem
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.AccStatus)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.Parent)
                Me.SubItems.Add(hCust)
                'Accounting info
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.SageID)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.PlatAccType)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.PoRequired)
                Me.SubItems.Add(hCust)

                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.CurrencyCode)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.VATRegno)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, "") ') Me.myClientHistory.VATRate)
                Me.SubItems.Add(hCust)
                'Email Info
                hCust = New ListViewItem.ListViewSubItem(Me, "") ' Me.myClientHistory.EmailResults)
                Me.SubItems.Add(hCust)
                hCust = New ListViewItem.ListViewSubItem(Me, Me.myClientHistory.ReportingInfo.MROID)
                Me.SubItems.Add(hCust)


            End If
            Return blnRet
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Client associated with this list view item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Client() As Intranet.intranet.customerns.Client
            Get
                Return myClient
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.Client)
                myClient = Value
                If Not Value Is Nothing Then DrawItem()
            End Set
        End Property
    End Class

    Public Class lvDonor
        Inherits System.Windows.Forms.ListViewItem

        Private myDonor As Intranet.intranet.customerns.DONOR
        Private JobName As String = ""
        Private Customer As String = ""

        Public Sub New(ByVal Donor As Intranet.intranet.customerns.DONOR)
            MyBase.new()
            myDonor = Donor
            Refresh()
        End Sub

        Public Sub New(ByVal Donor As String)
            MyBase.New()
            myDonor = New Intranet.intranet.customerns.DONOR()
            Dim DonorArray As String() = Donor.Split(New Char() {"|"})
            myDonor.FirstName = DonorArray(0)
            myDonor.Surname = DonorArray(1)
            myDonor.Gender = DonorArray(2)
            myDonor.CaseNumber = DonorArray(4)
            myDonor.CustJobRef = DonorArray(5)
            myDonor.EmployeeRef = DonorArray(6)

            Try
                myDonor.DateOfBirth = Date.Parse(DonorArray(3))
            Catch
            End Try
            If DonorArray.Length > 6 Then
                Customer = CConnection.PackageStringList("lib_customer.getsmidprofile", DonorArray(7))
            End If
            If DonorArray.Length > 9 Then
                JobName = DonorArray(10)
            End If
            Refresh()

        End Sub

        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            If Not Me.myDonor Is Nothing Then
                Me.Text = Me.myDonor.CaseNumber
                Dim objSubItem As ListViewItem.ListViewSubItem
                objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.Surname)
                Me.SubItems.Add(objSubItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.FirstName)
                Me.SubItems.Add(objSubItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.Gender)
                Me.SubItems.Add(objSubItem)

                Dim objDateItem As SubItems.DateItem = New SubItems.DateItem(myDonor.DateOfBirth)
                'objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.DateOfBirth.ToString("dd-MMM-yyyy"))
                Me.SubItems.Add(objDateItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.CustJobRef)
                Me.SubItems.Add(objSubItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, myDonor.EmployeeRef)
                Me.SubItems.Add(objSubItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, Customer)
                Me.SubItems.Add(objSubItem)

                objSubItem = New ListViewItem.ListViewSubItem(Me, JobName)
                Me.SubItems.Add(objSubItem)




            End If
        End Function

        Public ReadOnly Property Donor() As Intranet.intranet.customerns.DONOR
            Get
                Return Me.myDonor
            End Get
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvReportParameter
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Report parameter
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvReportParameter
        Inherits System.Windows.Forms.ListViewItem
        Private myReportParameter As MedscreenLib.CRFormulaItem

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' create new instance
        ''' </summary>
        ''' <param name="item"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal item As MedscreenLib.CRFormulaItem)
            MyBase.New()
            myReportParameter = item
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose report parameter
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ReportParameter() As MedscreenLib.CRFormulaItem
            Get
                Return myReportParameter
            End Get
            Set(ByVal Value As MedscreenLib.CRFormulaItem)
                myReportParameter = Value
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' refresh and draw item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [04/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            If Me.myReportParameter Is Nothing Then Exit Function
            Me.Text = myReportParameter.Formula
            Dim LvSubItem As ListViewSubItem
            LvSubItem = New ListViewSubItem(Me, myReportParameter.ParamType)
            Me.SubItems.Add(LvSubItem)
            LvSubItem = New ListViewSubItem(Me, myReportParameter.FieldName)
            Me.SubItems.Add(LvSubItem)
            LvSubItem = New ListViewSubItem(Me, myReportParameter.Value)
            Me.SubItems.Add(LvSubItem)

        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvProgUsageItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Program usage list view item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvProgUsageItem
        Inherits System.Windows.Forms.ListViewItem
        Private myProgramUsage As MedscreenLib.ProgramUsage

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Program usage list view item
        ''' </summary>
        ''' <param name="item"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal item As MedscreenLib.ProgramUsage)
            myProgramUsage = item
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Program usage item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ProgramUsage() As MedscreenLib.ProgramUsage
            Get
                Return Me.myProgramUsage
            End Get
            Set(ByVal Value As MedscreenLib.ProgramUsage)
                myProgramUsage = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh, i.e. redraw contents of list item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = myProgramUsage.ProgramName
            Dim LvSubItem As ListViewSubItem
            LvSubItem = New ListViewSubItem(Me, myProgramUsage.ComputerName)
            Me.SubItems.Add(LvSubItem)
            LvSubItem = New ListViewSubItem(Me, myProgramUsage.LastLogin.ToString("dd-MMM-yyyy HH:mm"))
            Me.SubItems.Add(LvSubItem)
            If myProgramUsage.LoggedIn Then
                LvSubItem = New ListViewSubItem(Me, "Logged In")
                Me.BackColor = System.Drawing.Color.LightPink
                'Me.ForeColor = System.Drawing.Color.Orange
            Else
                LvSubItem = New ListViewSubItem(Me, " - ")
                Me.BackColor = Drawing.Color.LightSeaGreen
                'Me.ForeColor = Drawing.Color.Yellow
            End If
            Me.SubItems.Add(LvSubItem)

        End Function
    End Class

    Public Class lvSampleDoc
        Inherits lvAdditionalDoc

        Public Sub New(ByVal Path As String, ByVal Barcode As String)

            MyBase.New()
            Dim addDoc As New Intranet.Support.AdditionalDocument()
            addDoc.Path = Path
            MyBase.AdditionalDocument = addDoc
            MyBase.Text = "COC - " & Barcode & "-" & Path
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvAdditionalDoc
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A listviewItem to hold additional documents
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [26/06/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvAdditionalDoc
        Inherits ListViewItem

        Private myAddDoc As Intranet.Support.AdditionalDocument

        Public Sub New(ByVal AddDoc As Intranet.Support.AdditionalDocument)
            myAddDoc = AddDoc
            Refresh()
        End Sub

        Friend Sub New()

        End Sub

        Public Function Refresh() As Boolean
            Me.Text = myAddDoc.Path
            Dim Asub As New ListViewItem.ListViewSubItem(Me, myAddDoc.Reference)
            Me.SubItems.Add(Asub)
            Asub = New ListViewItem.ListViewSubItem(Me, myAddDoc.CreatedOn.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(Asub)
            Asub = New ListViewItem.ListViewSubItem(Me, myAddDoc.Type)
            Me.SubItems.Add(Asub)
            Asub = New ListViewItem.ListViewSubItem(Me, myAddDoc.Linkfield)
            Me.SubItems.Add(Asub)
            Asub = New ListViewItem.ListViewSubItem(Me, myAddDoc.ModifiedOn.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(Asub)

        End Function

        Public Property AdditionalDocument() As Intranet.Support.AdditionalDocument
            Get
                Return myAddDoc
            End Get
            Set(ByVal Value As Intranet.Support.AdditionalDocument)
                myAddDoc = Value
            End Set
        End Property
    End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvComment
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List view item to display comments
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvComment
        Inherits ListViewItem

        Private myComment As Intranet.intranet.jobs.CollectionComment
        Private myMedComment As Comment

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new comment list view item 
        ''' </summary>
        ''' <param name="Comment"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Comment As Intranet.intranet.jobs.CollectionComment)
            MyBase.new()
            myComment = Comment
            Refresh()
        End Sub

        Public Sub New(ByVal Comment As Comment)
            MyBase.new()
            myMedComment = Comment
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.Text = ""
            Me.SubItems.Clear()
            If Not Me.myComment Is Nothing Then
                Me.Text = Me.myComment.CreatedOn.ToString("dd-MMM-yyyy HH:mm")
                Dim aSub As ListViewItem.ListViewSubItem
                aSub = New ListViewItem.ListViewSubItem(Me, myComment.CommentType)
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, myComment.CommentText)
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, myComment.Commenter)
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, myComment.ActionBy)
                Me.SubItems.Add(aSub)
            End If
            If Not Me.myMedComment Is Nothing Then
                Try
                    Me.Text = Me.myMedComment.Created.ToString("dd-MMM-yyyy HH:mm")
                    Dim aSub As ListViewItem.ListViewSubItem
                    aSub = New ListViewItem.ListViewSubItem(Me, myMedComment.TypeText)
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, myMedComment.Text)
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, myMedComment.CreatedBy)
                    Me.SubItems.Add(aSub)
                    'aSub = New ListViewItem.ListViewSubItem(Me, myMedComment.ActionB)
                    'Me.SubItems.Add(aSub)
                Catch
                End Try
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose comment to outside world
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Comment() As Intranet.intranet.jobs.CollectionComment
            Get
                Return myComment
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollectionComment)
                myComment = Value
                Me.Refresh()
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose generic comment type
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	28/09/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property MedComment() As Comment
            Get
                Return myMedComment
            End Get
            Set(ByVal Value As Comment)
                myMedComment = Value
            End Set
        End Property
    End Class


    Public Class ListViewCMDiaryItem
        Inherits ListViewItem
        Private strStatus As String
        Private strInvoiceid As String
        Private strCMJobNumber As String
        Private strClientID As String = ""
        Private strSMIDPRofile As String = ""
        Private strCollectionType As String
        Private strCollectionTypeCode As String
        Private strOfficer As String
        Private strVesselNAme As String
        Private strCollType As String
        Private strCollOfficer As String = ""
        Private strPurchaseOrder As String = ""
        Private strProgrammeManager As String
        Private strTableSet As String = ""
        Private strPortName As String
        Private strJobName As String
        Private blnUK As String
        Private datCollDate As Date
        Private objColl As Intranet.intranet.jobs.CMJob

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection list view item
        ''' </summary>
        ''' <param name="objCm">Collection to attach</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal CMJobNumber As String)
            MyBase.new()
            If CMJobNumber.Trim.Length > 0 Then
                Me.strCMJobNumber = CMJobNumber
                GetInfo()
                'strStatus = objCm.Status
                'strCMJobNumber = objCm.ID
                'strInvoice = objCm.Invoiceid
                'strOfficer = objCm.CollOfficer
                'blnUK = objCm.UK
                DrawItem()

            End If
        End Sub

        Private Function GetCardiffInfo(ByVal cmnumber As String) As String
            Dim Ticket As String = Mid(cmnumber, 3)
            Dim strret As String = BAOCardiff2.Collection.CollectionInfo(Ticket)
            
            Return strret

        End Function

        Private Sub GetInfo()
            If Me.strCMJobNumber.Trim.Length = 0 Then Exit Sub
            Dim strinfo As String = ""
            If strCMJobNumber.Chars(0) = "T" Then
                strinfo = GetCardiffInfo(strCMJobNumber)
            Else
                strinfo = CConnection.PackageStringList("lib_collectionxml8.GetCollectionInfo", Me.strCMJobNumber)

            End If
            Dim infoArray As String() = strinfo.Split(New Char() {","})
            Try
                If infoArray.Length > 0 Then strStatus = infoArray(0)
                If infoArray.Length > 1 Then Me.strInvoiceid = infoArray(1)
                If infoArray.Length > 2 Then Me.strClientID = infoArray(2)
                If infoArray.Length > 3 Then Me.strCollectionType = infoArray(3)
                If infoArray.Length > 4 Then Me.strOfficer = infoArray(4)
                If infoArray.Length > 5 Then Me.strVesselNAme = infoArray(5)
                If infoArray.Length > 6 Then Me.strCollType = infoArray(6)
                If infoArray.Length > 7 Then Me.strCollOfficer = infoArray(7)
                If infoArray.Length > 8 Then Me.strPurchaseOrder = infoArray(8)
                If infoArray.Length > 9 Then Me.strProgrammeManager = infoArray(9)
                If infoArray.Length > 10 Then Me.strTableSet = infoArray(10)
                If infoArray.Length > 11 Then Me.strJobName = infoArray(11)
                If infoArray.Length > 13 Then Me.strPortName = infoArray(13)
                If infoArray.Length > 14 Then Me.blnUK = (infoArray(14) = "T")
                If infoArray.Length > 15 Then Me.strCollectionTypeCode = infoArray(15)
                If infoArray.Length > 12 Then Me.CollDate = Date.Parse(infoArray(12))
            Catch
            End Try
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection list view item with additional info
        ''' </summary>
        ''' <param name="strId"></param>
        ''' <param name="Status"></param>
        ''' <param name="strInv"></param>
        ''' <param name="StrUK"></param>
        ''' <param name="Officer"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal strId As String, ByVal Status As String, _
        ByVal strInv As String, ByVal StrUK As String, ByVal Officer As String)
            MyBase.New()
            strStatus = Status
            strCMJobNumber = strId
            strInvoiceid = strInv
            strOfficer = Officer
            blnUK = (StrUK = "T")

            'Me.Collection = Intranet.intranet.Support.cSupport.ICollection(False).Item(strId)
            DrawItem()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw the item out 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/01/2006]</date>
        ''' <Action>Chaged reason for test to check for phrase id </Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub DrawItem()
            Me.SubItems.Clear()
            Me.Checked = False
            'If blnUK Then
            Me.Text = Me.strCollectionTypeCode & strCMJobNumber
            'Else
            'Me.Text = "O" & strCMJobNumber
            'End If
            Me.BackColor = Drawing.Color.Wheat


            Me.ImageIndex = 0
            Select Case strStatus
                Case "M"
                    Me.ImageIndex = 4
                Case "C"
                    Me.ImageIndex = 3
                Case "P"
                    Me.ImageIndex = 2
                Case "H"
                    Me.ImageIndex = 1

            End Select
            Dim oColInfo As UserDefaults.ColumnInfo

            MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Sort()
            'If Not Me.Collection Is Nothing Then
            'With Me.Collection
            'Port = Me.strPortName
            'CollDate = .CollectionDate
            Dim csItem As CollStatusSubItem = New CollStatusSubItem(strStatus, strInvoiceid, MedscreenCommonGui.CommonForms.UserOptions.DiaryColours)
            For Each oColInfo In MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns
                If oColInfo.Show Then
                    Select Case oColInfo.FieldName
                        Case "CLIENTID"
                            Dim Cust As ListViewItem.ListViewSubItem
                            Cust = New ListViewItem.ListViewSubItem(Me, Me.strClientID)
                            Me.SubItems.Add(Cust)
                            Cust.ForeColor = csItem.ForeColor
                            Cust.BackColor = csItem.BackColour


                        Case "PORT"
                            Dim siPort As ListViewItem.ListViewSubItem
                            siPort = New ListViewItem.ListViewSubItem(Me, strPortName)
                            Me.SubItems.Add(siPort)
                            'Item.Collection = objColl
                            siPort.ForeColor = csItem.ForeColour
                            siPort.BackColor = csItem.BackColour
                        Case "VESSEL"
                            Dim siVess As ListViewItem.ListViewSubItem
                            siVess = New ListViewItem.ListViewSubItem(Me, strVesselNAme)
                            SubItems.Add(siVess)
                            siVess.ForeColor = csItem.ForeColour
                            siVess.BackColor = csItem.BackColor

                        Case "STATUS"
                            SubItems.Add(csItem)
                        Case "COLLECTIONDATE"
                            Dim datItem As SubItems.DateItem = New SubItems.DateItem(CollDate)
                            datItem.BackColor = csItem.BackColor
                            datItem.ForeColor = csItem.ForeColor
                            SubItems.Add(datItem)
                        Case "COLLTYPE"
                            Dim siCollType As ListViewItem.ListViewSubItem
                            'see if using phrase id or phrase text 
                            Dim phraseText As String
                            Dim oColl As New Collection()
                            If strCollType.Length = 1 Then
                                oColl.Add("COLL_TYPE")
                            Else
                                oColl.Add("OTHTST_RSN")
                            End If
                            oColl.Add(strCollType)
                            phraseText = CConnection.PackageStringList("lib_utils.decodephrase", oColl)
                            'Dim objPhrase As MedscreenLib.Glossary.Phrase = MedscreenLib.Glossary.Glossary.ReasonForTestList.Item(.CollType)
                            If phraseText Is Nothing Then    'Using phrase text 
                                siCollType = New ListViewItem.ListViewSubItem(Me, "Unknown")
                            Else
                                siCollType = New ListViewItem.ListViewSubItem(Me, phraseText)
                            End If
                            SubItems.Add(siCollType)
                            siCollType.BackColor = csItem.BackColor
                            siCollType.ForeColor = csItem.ForeColor
                        Case "JOBSTATUS"
                            Dim jsItem As ListViewItems.JobStatusSubItem = New JobStatusSubItem(strTableSet, strJobName)
                            SubItems.Add(jsItem)
                            jsItem.BackColor = csItem.BackColor
                            jsItem.ForeColor = csItem.ForeColor
                        Case "JOBNAME"
                            If Not strJobName.Length = 0 Then
                                If strInvoiceid.Trim.Length = 0 Then
                                End If
                            End If

                            Dim siJN As New ListViewItem.ListViewSubItem()
                            siJN.BackColor = csItem.BackColor
                            siJN.ForeColor = csItem.ForeColor


                            If Not strJobName.Length = 0 Then
                                siJN.Text = strJobName
                                SubItems.Add(siJN)
                            Else
                                SubItems.Add(siJN)
                            End If
                        Case "COLLOFFICER"                                  'Display officer details
                            Dim siCollOff As ListViewItem.ListViewSubItem
                            If strCollOfficer.Trim.Length = 0 Then
                                siCollOff = New ListViewItem.ListViewSubItem(Me, "-")
                            Else
                                'Get officer description 
                                'Dim strDiaryOfficerName As String = MedscreenCommonGUIConfig.NodePLSQL.Item("DiaryOfficerName")
                                'Dim strOff As String = CConnection.PackageStringList(strDiaryOfficerName, strCollOfficer)
                                siCollOff = New ListViewItem.ListViewSubItem(Me, strCollOfficer)
                            End If
                            SubItems.Add(siCollOff)
                            siCollOff.BackColor = csItem.BackColor
                            siCollOff.ForeColor = csItem.ForeColor
                        Case "PONUMBER"
                            Dim siPO As ListViewItem.ListViewSubItem
                            If Not strPurchaseOrder.Length = 0 Then
                                siPO = New ListViewItem.ListViewSubItem(Me, strPurchaseOrder)
                                Me.SubItems.Add(siPO)
                            Else
                                siPO = New ListViewItem.ListViewSubItem(Me, "-")
                                Me.SubItems.Add(siPO)
                            End If
                            siPO.BackColor = csItem.BackColor
                            siPO.ForeColor = csItem.ForeColor
                        Case "PROGMANAGER"
                            Dim sPm As ListViewItem.ListViewSubItem
                            sPm = New ListViewItem.ListViewSubItem(Me, strProgrammeManager)
                            Me.SubItems.Add(sPm)

                            sPm.ForeColor = csItem.ForeColor
                            sPm.BackColor = csItem.BackColor
                        Case "ID"
                        Case Else
                            'Dim ne As ListViewItem.ListViewSubItem
                            'ne = New ListViewItem.ListViewSubItem(Me, .FieldValues.Item(oColInfo.FieldName).Value)
                            'Me.SubItems.Add(ne)

                            'ne.ForeColor = csItem.ForeColor
                            'ne.BackColor = csItem.BackColor
                    End Select
                End If
            Next

            ForeColor = System.Drawing.SystemColors.WindowText
            ForeColor = csItem.ForeColour
            BackColor = csItem.BackColour

            'Debug.WriteLine(strOff)
            'End With
            'End If

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' expose the collection tied to this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Collection() As Intranet.intranet.jobs.CMJob
            Get
                If objColl Is Nothing Then
                    objColl = New Intranet.intranet.jobs.CMJob(Me.strCMJobNumber)
                End If
                Return objColl
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CMJob)
                objColl = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tooltip to display 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToolTip() As String
            If strCMJobNumber.Trim.Length > 0 Then

                Dim strRet As String = CConnection.PackageStringList("Lib_collection.collectiontooltip", strCMJobNumber)
                Return strRet

            Else
                Return Me.CMJobNumber & " " & Me.strStatus
            End If
        End Function


        'Public Property Officer() As String
        '    Get
        '        Return strOfficer
        '    End Get
        '    Set(ByVal Value As String)
        '        strOfficer = Value
        '    End Set
        'End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Collection date associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CollDate() As Date
            Get
                Return datCollDate
            End Get
            Set(ByVal Value As Date)
                datCollDate = Value
            End Set
        End Property

        '''' -----------------------------------------------------------------------------
        '''' <summary>
        '''' Port associated with this item 
        '''' </summary>
        '''' <value></value>
        '''' <remarks>
        '''' </remarks>
        '''' <revisionHistory>
        '''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        '''' </revisionHistory>
        '''' -----------------------------------------------------------------------------
        'Public Property Port() As String
        '    Get
        '        Return strPort
        '    End Get
        '    Set(ByVal Value As String)
        '        strPort = Value
        '    End Set
        'End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Cm Job number associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CMJobNumber() As String
            Get
                Return strCMJobNumber
            End Get
            Set(ByVal Value As String)
                strCMJobNumber = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Status associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Status() As String
            Get
                Return strStatus
            End Get
            Set(ByVal Value As String)
                strStatus = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh(Optional ByVal SampleRefresh As Boolean = True) As Boolean
            Me.Collection.Refresh(SampleRefresh)
            Me.GetInfo()
            Me.DrawItem()

        End Function



        'Public Property BackColour() As System.Drawing.Color
        '    Get
        '        Return MyBase.BackColor
        '    End Get
        '    Set(ByVal Value As System.Drawing.Color)
        '        MyBase.BackColor = Value

        '    End Set
        'End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new item 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()

        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
            Me.Collection = Nothing
        End Sub
    End Class



    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.ListViewDiaryItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection item used in the diary (with small mods could be used elsewhere)
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ListViewDiaryItem
        Inherits ListViewItem
        Private strStatus As String
        Private strInvoice As String
        Private strCMJobNumber As String
        Private strOfficer As String
        Private strPort As String
        Private blnUK As String
        Private datCollDate As Date
        Private objColl As Intranet.intranet.jobs.CMJob

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection list view item
        ''' </summary>
        ''' <param name="objCm">Collection to attach</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal objCm As Intranet.intranet.jobs.CMJob)
            MyBase.new()
            If Not objCm Is Nothing Then
                Me.Collection = objCm
                strStatus = objCm.Status
                strCMJobNumber = objCm.ID
                strInvoice = objCm.Invoiceid
                strOfficer = objCm.CollOfficer
                blnUK = objCm.UK
                DrawItem()

            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection list view item with additional info
        ''' </summary>
        ''' <param name="strId"></param>
        ''' <param name="Status"></param>
        ''' <param name="strInv"></param>
        ''' <param name="StrUK"></param>
        ''' <param name="Officer"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal strId As String, ByVal Status As String, _
        ByVal strInv As String, ByVal StrUK As String, ByVal Officer As String)
            MyBase.New()
            strStatus = Status
            strCMJobNumber = strId
            strInvoice = strInv
            strOfficer = Officer
            blnUK = (StrUK = "T")

            Me.Collection = Intranet.intranet.Support.cSupport.ICollection(False).Item(strId)
            DrawItem()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Draw the item out 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [06/01/2006]</date>
        ''' <Action>Chaged reason for test to check for phrase id </Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub DrawItem()
            Me.SubItems.Clear()
            Me.Checked = False
            'If blnUK Then
            Me.Text = Me.Collection.CollectionType & strCMJobNumber
            'Else
            'Me.Text = "O" & strCMJobNumber
            'End If
            Me.BackColor = Drawing.Color.Wheat


            Me.ImageIndex = 0
            Select Case strStatus
                Case "M"
                    Me.ImageIndex = 4
                Case "C"
                    Me.ImageIndex = 3
                Case "P"
                    Me.ImageIndex = 2
                Case "H"
                    Me.ImageIndex = 1

            End Select
            Dim oColInfo As UserDefaults.ColumnInfo

            MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns.Sort()
            If Not Me.Collection Is Nothing Then
                With Me.Collection
                    Port = .PortName
                    CollDate = .CollectionDate
                    Dim csItem As CollStatusSubItem = New CollStatusSubItem(.Status, .Invoiceid, MedscreenCommonGui.CommonForms.UserOptions.DiaryColours)
                    For Each oColInfo In MedscreenCommonGui.CommonForms.UserOptions.DiaryColumns
                        If oColInfo.Show Then
                            Select Case oColInfo.FieldName
                                Case "CLIENTID"
                                    Dim Cust As ListViewItem.ListViewSubItem
                                    Cust = New ListViewItem.ListViewSubItem(Me, Me.Collection.SmClient.ClientID)
                                    Me.SubItems.Add(Cust)
                                    Cust.ForeColor = csItem.ForeColor
                                    Cust.BackColor = csItem.BackColour


                                Case "PORT"
                                    Dim siPort As ListViewItem.ListViewSubItem
                                    siPort = New ListViewItem.ListViewSubItem(Me, .PortName)
                                    Me.SubItems.Add(siPort)
                                    'Item.Collection = objColl
                                    siPort.ForeColor = csItem.ForeColour
                                    siPort.BackColor = csItem.BackColour
                                Case "VESSEL"
                                    Dim siVess As ListViewItem.ListViewSubItem
                                    siVess = New ListViewItem.ListViewSubItem(Me, .VesselName)
                                    SubItems.Add(siVess)
                                    'If .vessel.Trim.Length > 0 Then
                                    '    If (.vessel.Chars(0) < "A") Or (InStr(.vessel, "VESS") = 1) Or (InStr(.vessel, "IMO") = 1) Then
                                    '        If Not .VesselObject Is Nothing Then
                                    '            siVess = New ListViewItem.ListViewSubItem(Me, .VesselObject.VesselName)
                                    '            SubItems.Add(siVess)
                                    '        Else
                                    '            siVess = New ListViewItem.ListViewSubItem(Me, "-")
                                    '            SubItems.Add(siVess)
                                    '        End If
                                    '    ElseIf (InStr(.vessel, "SITE") = 1) Then
                                    '        If Not .SiteObject Is Nothing Then
                                    '            siVess = New ListViewItem.ListViewSubItem(Me, .SiteObject.VesselName)
                                    '            SubItems.Add(siVess)
                                    '        Else
                                    '            siVess = New ListViewItem.ListViewSubItem(Me, "-")
                                    '            SubItems.Add(siVess)
                                    '        End If
                                    '    Else
                                    '        siVess = New ListViewItem.ListViewSubItem(Me, .vessel)
                                    '        SubItems.Add(siVess)
                                    '    End If
                                    'Else
                                    '    siVess = New ListViewItem.ListViewSubItem(Me, "-")
                                    '    SubItems.Add(siVess)
                                    'End If
                                    siVess.ForeColor = csItem.ForeColour
                                    siVess.BackColor = csItem.BackColor

                                Case "STATUS"
                                    SubItems.Add(csItem)
                                Case "COLLECTIONDATE"
                                    Dim datItem As SubItems.DateItem = New SubItems.DateItem(.CollectionDate)
                                    datItem.BackColor = csItem.BackColor
                                    datItem.ForeColor = csItem.ForeColor
                                    SubItems.Add(datItem)
                                Case "COLLTYPE"
                                    Dim siCollType As ListViewItem.ListViewSubItem
                                    'see if using phrase id or phrase text 
                                    Dim phraseText As String
                                    Dim oColl As New Collection()
                                    If .CollType.Length = 1 Then
                                        oColl.Add("COLL_TYPE")
                                    Else
                                        oColl.Add("OTHTST_RSN")
                                    End If
                                    oColl.Add(.CollType)
                                    phraseText = CConnection.PackageStringList("lib_utils.decodephrase", oColl)
                                    'Dim objPhrase As MedscreenLib.Glossary.Phrase = MedscreenLib.Glossary.Glossary.ReasonForTestList.Item(.CollType)
                                    If phraseText Is Nothing Then    'Using phrase text 
                                        siCollType = New ListViewItem.ListViewSubItem(Me, "Unknown")
                                    Else
                                        siCollType = New ListViewItem.ListViewSubItem(Me, phraseText)
                                    End If
                                    SubItems.Add(siCollType)
                                    siCollType.BackColor = csItem.BackColor
                                    siCollType.ForeColor = csItem.ForeColor
                                Case "JOBSTATUS"
                                    Dim jsItem As ListViewItems.JobStatusSubItem = New JobStatusSubItem(.TableSet, .JobName)
                                    SubItems.Add(jsItem)
                                    jsItem.BackColor = csItem.BackColor
                                    jsItem.ForeColor = csItem.ForeColor
                                Case "JOBNAME"
                                    If Not .JobName.Length = 0 Then
                                        If .Invoiceid.Trim.Length = 0 Then
                                        End If
                                    End If

                                    Dim siJN As New ListViewItem.ListViewSubItem()
                                    siJN.BackColor = csItem.BackColor
                                    siJN.ForeColor = csItem.ForeColor


                                    If Not .JobName.Length = 0 Then
                                        siJN.Text = .JobName
                                        SubItems.Add(siJN)
                                    Else
                                        SubItems.Add(siJN)
                                    End If
                                Case "COLLOFFICER"                                  'Display officer details
                                    Dim siCollOff As ListViewItem.ListViewSubItem
                                    If .CollOfficer.Trim.Length = 0 Then
                                        siCollOff = New ListViewItem.ListViewSubItem(Me, "-")
                                    Else
                                        'Get officer description 
                                        Dim strDiaryOfficerName As String = MedscreenCommonGUIConfig.NodePLSQL.Item("DiaryOfficerName")
                                        Dim strOff As String = CConnection.PackageStringList(strDiaryOfficerName, .CollOfficer)
                                        siCollOff = New ListViewItem.ListViewSubItem(Me, strOff)
                                    End If
                                    SubItems.Add(siCollOff)
                                    siCollOff.BackColor = csItem.BackColor
                                    siCollOff.ForeColor = csItem.ForeColor
                                Case "PONUMBER"
                                    Dim siPO As ListViewItem.ListViewSubItem
                                    If Not .PurchaseOrder.Length = 0 Then
                                        siPO = New ListViewItem.ListViewSubItem(Me, .PurchaseOrder)
                                        Me.SubItems.Add(siPO)
                                    Else
                                        siPO = New ListViewItem.ListViewSubItem(Me, "-")
                                        Me.SubItems.Add(siPO)
                                    End If
                                    siPO.BackColor = csItem.BackColor
                                    siPO.ForeColor = csItem.ForeColor
                                Case "PROGMANAGER"
                                    Dim sPm As ListViewItem.ListViewSubItem
                                    sPm = New ListViewItem.ListViewSubItem(Me, .ProgrammeManager)
                                    Me.SubItems.Add(sPm)

                                    sPm.ForeColor = csItem.ForeColor
                                    sPm.BackColor = csItem.BackColor
                                Case "ID"
                                Case Else
                                    Dim ne As ListViewItem.ListViewSubItem
                                    ne = New ListViewItem.ListViewSubItem(Me, .FieldValues.Item(oColInfo.FieldName).Value)
                                    Me.SubItems.Add(ne)

                                    ne.ForeColor = csItem.ForeColor
                                    ne.BackColor = csItem.BackColor
                            End Select
                        End If
                    Next

                    ForeColor = System.Drawing.SystemColors.WindowText
                    ForeColor = csItem.ForeColour
                    BackColor = csItem.BackColour

                    'Debug.WriteLine(strOff)
                End With
            End If

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' expose the collection tied to this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Collection() As Intranet.intranet.jobs.CMJob
            Get
                Return objColl
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CMJob)
                objColl = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tooltip to display 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToolTip() As String
            If Not Me.objColl Is Nothing Then

                Dim strRet As String = Me.Collection.ToolTip
                Return strRet

            Else
                Return Me.CMJobNumber & " " & Me.strStatus
            End If
        End Function


        'Public Property Officer() As String
        '    Get
        '        Return strOfficer
        '    End Get
        '    Set(ByVal Value As String)
        '        strOfficer = Value
        '    End Set
        'End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Collection date associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CollDate() As Date
            Get
                Return datCollDate
            End Get
            Set(ByVal Value As Date)
                datCollDate = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Port associated with this item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Port() As String
            Get
                Return strPort
            End Get
            Set(ByVal Value As String)
                strPort = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Cm Job number associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CMJobNumber() As String
            Get
                Return strCMJobNumber
            End Get
            Set(ByVal Value As String)
                strCMJobNumber = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Status associated with this item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Status() As String
            Get
                Return strStatus
            End Get
            Set(ByVal Value As String)
                strStatus = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh(Optional ByVal SampleRefresh As Boolean = True) As Boolean
            Me.Collection.Refresh(SampleRefresh)
            Me.DrawItem()

        End Function

        'Public Property BackColour() As System.Drawing.Color
        '    Get
        '        Return MyBase.BackColor
        '    End Get
        '    Set(ByVal Value As System.Drawing.Color)
        '        MyBase.BackColor = Value

        '    End Set
        'End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new item 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvCustomerTestPanel
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Class to draw a test panel 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvCustomerTestPanel
        Inherits ListViewItem
        Private myCustTestPanel As Intranet.intranet.TestPanels.CustomerTestPanel

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new customer panel list view item
        ''' </summary>
        ''' <param name="CustTestPanel"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal CustTestPanel As Intranet.intranet.TestPanels.CustomerTestPanel)
            MyBase.New()
            myCustTestPanel = CustTestPanel
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Customer Test panel exposed
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CustomerTestPanel() As Intranet.intranet.TestPanels.CustomerTestPanel
            Get
                Return Me.myCustTestPanel
            End Get
            Set(ByVal Value As Intranet.intranet.TestPanels.CustomerTestPanel)
                myCustTestPanel = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the view of this panel
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            If Not Me.myCustTestPanel Is Nothing Then
                Me.Text = Me.myCustTestPanel.TestPanelID
                Dim objSubItem As ListViewItem.ListViewSubItem
                With Me.myCustTestPanel
                    'test panel description 
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .TestPanel.TestPanelDescription)
                    Me.SubItems.Add(objSubItem)
                    'default
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .DefaultPanel)
                    Me.SubItems.Add(objSubItem)
                    'Matrix
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .TestPanel.Matrix)
                    Me.SubItems.Add(objSubItem)
                    'alcohol
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .Alcohol.ToString("###"))
                    Me.SubItems.Add(objSubItem)
                    'cannabis
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .Cannabis.ToString("###"))
                    Me.SubItems.Add(objSubItem)
                    'Breath 
                    'alcohol
                    objSubItem = New ListViewItem.ListViewSubItem(Me, .BreathCutoff.ToString("###") & .BreathUnits)
                    Me.SubItems.Add(objSubItem)

                End With
            End If
        End Function
    End Class

    'Public Class lvMatrix
    '    Inherits ListViewItem
    '    Private myMatrix As Intranet.intranet.customerns.CustomerMatrix

    '    Public Sub New(ByVal Matrix As Intranet.intranet.customerns.CustomerMatrix)
    '        MyBase.new()
    '        myMatrix = Matrix
    '        refresh()
    '    End Sub

    '    Public Function refresh() As Boolean
    '        If myMatrix Is Nothing Then Exit Function
    '        Me.SubItems.Clear()
    '        Me.Text = myMatrix.Matrix
    '        Dim aSub As ListViewItem.ListViewSubItem
    '        aSub = New ListViewItem.ListViewSubItem(Me, myMatrix.ScheduleDescription)
    '        Me.SubItems.Add(aSub)
    '        aSub = New ListViewItem.ListViewSubItem(Me, myMatrix.ProductDescription)
    '        Me.SubItems.Add(aSub)
    '        aSub = New ListViewItem.ListViewSubItem(Me, myMatrix.FlawScheduleDescription)
    '        Me.SubItems.Add(aSub)
    '        aSub = New ListViewItem.ListViewSubItem(Me, myMatrix.FlawSpecDescription)
    '        Me.SubItems.Add(aSub)
    '    End Function

    '    'Public Property CustomerMatrix() As Intranet.intranet.customerns.CustomerMatrix
    '    '    Get
    '    '        Return myMatrix
    '    '    End Get
    '    '    Set(ByVal Value As Intranet.intranet.customerns.CustomerMatrix)
    '    '        myMatrix = Value
    '    '    End Set
    '    'End Property
    'End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.lvTestSchedule
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List view item for test schedule
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvTestSchedule
        Inherits ListViewItem
        Private mySchedule As Intranet.intranet.TestResult.TestScheduleEntry
        Private myTest As Intranet.intranet.TestPanels.TestPanelTest
        Private strProduct As String = ""
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' create a new list view item
        ''' </summary>
        ''' <param name="Schedule"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Schedule As Intranet.intranet.TestResult.TestScheduleEntry, Optional ByVal Product As String = "")
            MyBase.new()
            mySchedule = Schedule
            myTest = Nothing
            strProduct = Product
            Refresh()

        End Sub

        Public Sub New(ByVal Test As Intranet.intranet.TestPanels.TestPanelTest)
            MyBase.new()
            mySchedule = Nothing
            myTest = Test
            Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh/redraw item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [02/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            If Not mySchedule Is Nothing Then
                Me.Text = mySchedule.AnalysisID
                Dim objSubItem As ListViewItem.ListViewSubItem
                objSubItem = New ListViewItem.ListViewSubItem(Me, mySchedule.AnalysisDecription)
                Me.SubItems.Add(objSubItem)
                objSubItem = New ListViewItem.ListViewSubItem(Me, mySchedule.AnalysisTypeDescription)
                Me.SubItems.Add(objSubItem)
                objSubItem = New ListViewItem.ListViewSubItem(Me, mySchedule.StandardTest)
                Me.SubItems.Add(objSubItem)
                objSubItem = New ListViewItem.ListViewSubItem(Me, mySchedule.ReplicateCount)
                Me.SubItems.Add(objSubItem)
                If strProduct.Trim.Length > 0 Then
                    objSubItem = New ListViewItem.ListViewSubItem(Me, mySchedule.ProductAnalysisDescription(Me.strProduct))
                    Me.SubItems.Add(objSubItem)
                End If
                'Set colours 
                If mySchedule.StandardTest = True Then
                    Me.BackColor = Color.LightGreen
                Else
                    Me.BackColor = Color.LightBlue
                End If
            End If
            If Not myTest Is Nothing Then
                Me.Text = myTest.Test
                Dim objSubItem As ListViewItem.ListViewSubItem
                If Not myTest.ComponentView Is Nothing Then
                    objSubItem = New ListViewItem.ListViewSubItem(Me, myTest.ComponentView.Name)
                    Me.SubItems.Add(objSubItem)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, myTest.ComponentView.IsConfirmatory)
                    Me.SubItems.Add(objSubItem)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, "")
                    Me.SubItems.Add(objSubItem)
                End If

            End If
        End Function

        ''' <summary>
        '''     Schedule details 
        ''' </summary>
        ''' <value>
        '''     <para>
        '''         
        '''     </para>
        ''' </value>
        ''' <remarks>
        '''     
        ''' </remarks>
        Public Property Schedule() As Intranet.intranet.TestResult.TestScheduleEntry
            Get
                Return mySchedule
            End Get
            Set(ByVal Value As Intranet.intranet.TestResult.TestScheduleEntry)
                mySchedule = Value
            End Set
        End Property
    End Class

    Public Class lviDonorTemplateRow
        Inherits ListViewItem

        Private Const strTick As String = "þ"
        Private Const strCross As String = "ý"

        Public Sub New(ByVal DonorRow As String)
            MyBase.New()
            Dim backColour As Color

            Me.UseItemStyleForSubItems = False
            Dim strSubItems As String() = DonorRow.Split(New Char() {","})
            If strSubItems.GetValue(9) = "T" Then
                backColour = Color.Beige
            Else
                backColour = Color.LightGreen
            End If
            Me.Text = strSubItems.GetValue(1)
            Dim aSub As ListViewItem.ListViewSubItem
            Dim strId As String = CStr(strSubItems.GetValue(0)).PadLeft(3)
            strId = strId.Replace(" ", "0")
            aSub = New ListViewItem.ListViewSubItem(Me, strId)
            aSub.BackColor = backColour
            Me.SubItems.Add(aSub)
            aSub = New ListViewItem.ListViewSubItem(Me, strSubItems.GetValue(3))
            aSub.BackColor = backColour

            Me.SubItems.Add(aSub)
            Dim strValue As String = strSubItems.GetValue(8)
            Dim aFont As Font
            If strValue = "T" Then
                strValue = strTick
                aFont = New Font("WINGDINGS", 10, FontStyle.Bold)
            Else
                strValue = strCross
                aFont = New Font("WINGDINGS", 10, FontStyle.Italic)
            End If
            aSub = New ListViewItem.ListViewSubItem(Me, strValue)
            aSub.BackColor = backColour
            aSub.Font = aFont
            Me.SubItems.Add(aSub)

            strValue = strSubItems.GetValue(9)
            If strValue = "T" Then
                strValue = strTick
                aFont = New Font("WINGDINGS", 10, FontStyle.Bold)
            Else
                strValue = strCross
                aFont = New Font("WINGDINGS", 10, FontStyle.Regular)
            End If
            aSub.BackColor = backColour
            If strSubItems.GetValue(8) = "T" Then
                aSub.ForeColor = Color.Green
            Else
                aSub.ForeColor = Color.Red
            End If

            aSub = New ListViewItem.ListViewSubItem(Me, strValue)
            aSub.Font = aFont


            Me.SubItems.Add(aSub)
            If strSubItems.GetValue(9) = "T" Then
                Me.BackColor = Color.Beige
                aSub.ForeColor = Color.Green
            Else
                Me.BackColor = Color.LightGreen
                aSub.ForeColor = Color.Red
            End If

            aSub.BackColor = backColour
            Try
                Dim strPhraseType As String = strSubItems.GetValue(10)
                If strPhraseType.Trim.Length > 0 Then 'We have a phrase or a default value 
                    'A default value will be prefixed by a tilda ~
                    If strPhraseType.Chars(0) = "~" Then 'default value 
                        aSub = New ListViewItem.ListViewSubItem(Me, Mid(strPhraseType, 2))
                    Else
                        aSub = New ListViewItem.ListViewSubItem(Me, CConnection.PackageStringList("Lib_utils.PhraseDescription", strPhraseType))
                    End If
                    Me.SubItems.Add(aSub)
                    aSub.BackColor = backColour
                Else
                    aSub = New ListViewItem.ListViewSubItem(Me, "")
                    Me.SubItems.Add(aSub)
                    aSub.BackColor = backColour
                End If
            Catch
            End Try
        End Sub
    End Class

    Public Class lviLoginTemplate
        Inherits ListViewItem

        Private myCustomerLoginTemplate As Intranet.Templates.CustomerLoginTemplate

        Public Sub New(ByVal CustomerLoginTemplate As Intranet.Templates.CustomerLoginTemplate)
            MyBase.New()
            myCustomerLoginTemplate = CustomerLoginTemplate
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.Text = myCustomerLoginTemplate.OrderNumber

            Dim objSub As ListViewItem.ListViewSubItem

            objSub = New ListViewItem.ListViewSubItem(Me, myCustomerLoginTemplate.SampleTemplateDescription)
            Me.SubItems.Add(objSub)
            objSub = New ListViewItem.ListViewSubItem(Me, myCustomerLoginTemplate.DonorTemplateDescription)
            Me.SubItems.Add(objSub)
            If Me.Selected Then
                Me.BackColor = SystemColors.ActiveCaption
                Me.ForeColor = SystemColors.ActiveCaptionText
            Else
                Me.BackColor = SystemColors.Window
                Me.ForeColor = SystemColors.WindowText

            End If
        End Function



        Public Property CustomerLoginTemplate() As Intranet.Templates.CustomerLoginTemplate
            Get
                Return myCustomerLoginTemplate
            End Get
            Set(ByVal Value As Intranet.Templates.CustomerLoginTemplate)
                myCustomerLoginTemplate = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.HolidayItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A holiday item used in collector form list view 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class HolidayItem
        Inherits ListViewItem

        Private myHol As Intranet.intranet.jobs.CollectorHoliday

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor for holiday item 
        ''' </summary>
        ''' <param name="Holiday">Holiday supplied</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Holiday As Intranet.intranet.jobs.CollectorHoliday)
            myHol = Holiday
            Build()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Build control
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub Build()
            Me.SubItems.Clear()
            Me.Text = myHol.HolidayStart.ToString("dd-MMM-yyyy")
            Me.ImageIndex = 0
            Me.StateImageIndex = 0
            Me.SubItems.Add(myHol.HolidayEnd.ToString("dd-MMM-yyyy"))
            Me.SubItems.Add(myHol.HolidayEnd.Subtract(myHol.HolidayStart).Days + 1)
            Me.SubItems.Add(myHol.DeputyId.Trim)
            Me.SubItems.Add(myHol.Comment.Trim)
            Me.SubItems.Add(myHol.AbsenceType.Trim)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Pass Holiday in 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Holiday() As Intranet.intranet.jobs.CollectorHoliday
            Get
                Return myHol
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollectorHoliday)
                myHol = Value

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw element 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.Build()
            Return True
        End Function
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : CMListItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An inherited list view item for use with collection cancellations
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CMListItem
        Inherits ListViewItem

#Region "Declarations"

        Private myCollection As Intranet.intranet.jobs.CMJob
        Private objClient As Intranet.intranet.customerns.Client
        Private strCurr As String = "GBP"
        Private strComment As String = ""
        Private CancelCharge As Double = 0
        Private cancelFee As Double = 0
        Private myDone As Boolean = False
        Private myListViewType As ListViewItemType = ListViewItemType.CancelInvoice
#End Region

#Region "Public enumerations"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Possible list view types supported
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Enum ListViewItemType
            ''' <summary>Cancel Invoice</summary>
            CancelInvoice
            ''' <summary>Late Callout</summary>
            LateCallout
        End Enum
#End Region

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Deals with the collection associated with this item, the set method will build the item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Collection() As Intranet.intranet.jobs.CMJob
            Get
                Return Me.myCollection
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CMJob)
                myCollection = Value
                Me.Refresh()
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw the item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Try
                If Not myCollection Is Nothing Then                 'Do we have a valid collection 
                    Me.SubItems.Clear()
                    Me.UseItemStyleForSubItems = True
                    Me.Text = myCollection.ID                       'Set the CM JOB NUMBER as the text property
                    Me.myCollection.Load()
                    'create a sub item to build the rest of the columns 
                    Dim objSub As ListViewItem.ListViewSubItem
                    'Set client id (column 2)
                    Dim strGetSMIDProfile As String = MedscreenCommonGUIConfig.NodePLSQL.Item("GetSMIDProfile")

                    Dim strCustID As String = CConnection.PackageStringList(strGetSMIDProfile, myCollection.ClientId)
                    If strCustID Is Nothing Then
                        objSub = New ListViewItem.ListViewSubItem(Me, "Unknown")
                    Else
                        objSub = New ListViewItem.ListViewSubItem(Me, strCustID)
                    End If
                    Me.SubItems.Add(objSub)
                    'Set collection date (column 3)
                    objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(myCollection.CollectionDate, "dd-MMM-yyyy")
                    Me.SubItems.Add(objSub)


                    'See if we have a comment, if we have we will store it in a local variable so it can be exposed as a property 
                    'of this object 
                    If Not myCollection.Comments Is Nothing Then
                        Dim I As Integer
                        Dim oComm As Intranet.intranet.jobs.CollectionComment  ' go through all comments
                        For I = 0 To myCollection.Comments.Count - 1
                            oComm = myCollection.Comments.item(I)
                            If oComm.CommentType = "CANC" And Me.myListViewType = ListViewItemType.CancelInvoice Then                  'If we have a cancellation comment Good oh 
                                strComment = oComm.CommentText
                                Exit For
                            ElseIf oComm.CommentType = "CALT" And Me.myListViewType = ListViewItemType.LateCallout Then                                          'If we have a cancellation comment Good oh 
                                strComment = oComm.CommentText
                                Exit For

                            End If
                        Next
                    End If

                    'Set collection date (column 4)
                    objSub = New ListViewItem.ListViewSubItem(Me, strComment)
                    Me.SubItems.Add(objSub)


                    If Me.myListViewType = ListViewItemType.CancelInvoice Then
                        'Set agreed charges (column 5)
                        objSub = New ListViewItem.ListViewSubItem(Me, myCollection.AgreedExpenses.ToString("c", myCollection.SmClient.Culture))
                        Me.SubItems.Add(objSub)

                        'Set Pre paid (column 6)
                        objSub = New ListViewItem.ListViewSubItem(Me, myCollection.PrePaid)
                        Me.SubItems.Add(objSub)



                        'Get currency
                        strCurr = myCollection.SmClient.CurrencyCode


                        If myCollection.CancellationCharge = 0 Then
                            CancelCharge = myCollection.AgreedExpenses + myCollection.CollectorsPayAmount
                        End If
                        If Not myCollection.SmClient Is Nothing Then
                            If myCollection.SmClient.HasStdPricing Or myCollection.SmClient.IsLondonUnderground Then
                                CancelCharge = 0
                            End If
                        End If

                        'Set currency (column 7)
                        'objSub = New ListViewItem.ListViewSubItem(Me, strCurr)
                        'Me.SubItems.Add(objSub)
                        'Set extra charge (column 8)
                        objSub = New ListViewItem.ListViewSubItem(Me, CancelCharge.ToString("c", myCollection.SmClient.Culture))

                        Me.SubItems.Add(objSub)
                        'Set cancellation fee(column 9)
                        Me.cancelFee = myCollection.CancellationCharge              'Set property ready 
                        objSub = New ListViewItem.ListViewSubItem(Me, myCollection.CancellationCharge.ToString("c", myCollection.SmClient.Culture))
                        Me.SubItems.Add(objSub)
                        'Set total cost (column 10)
                        objSub = New ListViewItem.ListViewSubItem(Me, (CancelCharge + myCollection.CancellationCharge).ToString("c", myCollection.SmClient.Culture))
                        Me.SubItems.Add(objSub)
                        'Set cancellation date (column 11)
                        objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(myCollection.CancellationDate, "dd-MMM-yyyy HH:mm")
                        Me.SubItems.Add(objSub)
                        'Set cancellation status (column 12)
                        objSub = New ListViewItem.ListViewSubItem(Me, myCollection.CancelStatusString)
                        Me.SubItems.Add(objSub)
                        'invoice number
                        objSub = New ListViewItem.ListViewSubItem(Me, myCollection.Invoiceid)
                        Me.SubItems.Add(objSub)

                        If myDone Then
                            Me.ImageIndex = 6
                            Me.BackColor = Color.DarkBlue
                            Me.ForeColor = Color.DarkGray
                        Else
                            Dim myImageIndex As Integer = 0
                            'Check that we have a valid Collection type 
                            If myCollection.CollectionType.Trim.Length > 0 Then
                                If myCollection.CollClassType = Intranet.intranet.jobs.CMJob.CollectionTypes.GCST_JobTypeCallout Or _
                                    myCollection.CollClassType = Intranet.intranet.jobs.CMJob.CollectionTypes.GCST_JobTypeCalloutOverseas Then
                                    myImageIndex = 1
                                ElseIf myCollection.CollClassType = Intranet.intranet.jobs.CMJob.CollectionTypes.GCST_JobTypeOverseas Then
                                    myImageIndex = 2
                                End If
                            End If
                            If myCollection.CancellationStatus = "W" Then
                                Me.BackColor = Color.PeachPuff
                                Me.ImageIndex = myImageIndex + 3

                            ElseIf myCollection.CancellationStatus = "R" Then
                                Me.BackColor = Color.PaleGreen
                                Me.ImageIndex = myImageIndex
                            End If
                        End If  'Mydone
                    Else
                        'Set booking date (column 5)
                        objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(myCollection.RequestReceived, "ddd, dd-MMM-yyyy HH:mm")
                        Me.SubItems.Add(objSub)
                        'Set date to officer (column 5)
                        objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(myCollection.DateCalloutToOfficer, "dd-MMM-yyyy HH:mm")
                        Me.SubItems.Add(objSub)
                        'Do we have pay info ?
                        Dim objPay As Intranet.intranet.jobs.CollPay = myCollection.CollectorsPay
                        If Not objPay Is Nothing Then       'Yes
                            'Set left site (column 6)
                            objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(objPay.TravelStart, "dd-MMM-yyyy HH:mm")
                            Me.SubItems.Add(objSub)
                            'Set left site (column 7)
                            objSub = New MedscreenCommonGui.ListViewItems.SubItems.DateItem(objPay.ArriveOnSite, "dd-MMM-yyyy HH:mm")
                            Me.SubItems.Add(objSub)
                            Dim ts As TimeSpan = objPay.ArriveOnSite.Subtract(myCollection.RequestReceived)
                            objSub = New MedscreenCommonGui.ListViewItems.SubItems.IntegerListViewSubItem(ts.TotalMinutes)
                            Me.SubItems.Add(objSub)
                            If myDone Then
                                Me.ImageIndex = 6
                                Me.BackColor = Color.DarkBlue
                                Me.ForeColor = Color.DarkGray
                            Else
                                If ts.TotalMinutes < 0 Then
                                    Me.BackColor = Color.Red
                                    Me.ForeColor = Color.Yellow
                                ElseIf ts.TotalMinutes > 1000 Then
                                    Me.BackColor = Color.Blue 'Color.Plum
                                ElseIf ts.TotalMinutes > 190 Then
                                    Me.BackColor = Color.Green 'Color.PeachPuff

                                Else
                                    Me.BackColor = Color.Yellow 'Color.PowderBlue
                                End If
                            End If
                        End If

                    End If      'Listviewitemtype 
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "Error in CMListItem - Refresh")
            End Try
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection cancellation invoice list item 
        ''' </summary>
        ''' <param name="Coll"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Coll As Intranet.intranet.jobs.CMJob)
            MyBase.New()

            Me.Collection = Coll
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Alternate constructor with listviewtype 
        ''' </summary>
        ''' <param name="Coll">Collection</param>
        ''' <param name="ItemType">Form Type</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Coll As Intranet.intranet.jobs.CMJob, ByVal ItemType As ListViewItemType)

            MyBase.new()
            Me.myListViewType = ItemType
            Me.Collection = Coll

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Reason given for cancellation 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [02/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ReasonForCancellation() As String
            Get
                Return strComment
            End Get
            Set(ByVal Value As String)
                strComment = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' User defined cancellation charge 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [02/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ActualCancelCharge() As Double
            Get
                Return Me.CancelCharge
            End Get
            Set(ByVal Value As Double)
                CancelCharge = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Actual fee to be charged to customer 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CancellationFee() As Double
            Get
                Return Me.cancelFee
            End Get
            Set(ByVal Value As Double)
                Me.cancelFee = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Have we dealt with this collection 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Done() As Boolean
            Get
                Return Me.myDone

            End Get
            Set(ByVal Value As Boolean)
                myDone = Value
                Me.Collection.Refresh(False)
                Me.Refresh()
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.SiteSubItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A Sub item type used specifically by the collecting officer List view item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class SiteSubItem
        Inherits ListViewItem.ListViewSubItem

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Ctreate a new site sub item 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            Me.BackColor = System.Drawing.SystemColors.Info
            Me.ForeColor = System.Drawing.SystemColors.InfoText

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.LVVesselItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A vessel list view item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LVVesselItem
        Inherits ListViewItem
        Private myVessel As Intranet.intranet.customerns.Vessel

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new Vessel List view item 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()

        End Sub

        Public Sub Refresh()
            Me.SubItems.Clear()
            Me.Text = Vessel.VesselName
            Me.BackColor = Color.AliceBlue
            Dim sit As ListViewItem.ListViewSubItem

            'Change it so the customer is described better
            Dim strID As String = CConnection.PackageStringList("lib_customer.GetSMIDProfile", myVessel.CustomerID)
            If strID.Trim.Length = 0 Then
                sit = New ListViewItem.ListViewSubItem(Me, myVessel.CustomerID)
            Else
                sit = New ListViewItem.ListViewSubItem(Me, strID)
            End If
            sit.BackColor = Me.BackColor
            Me.SubItems.Add(sit)
            sit = New ListViewItem.ListViewSubItem(Me, myVessel.VesselID)
            sit.BackColor = Me.BackColor
            Me.SubItems.Add(sit)
            sit = New ListViewItem.ListViewSubItem(Me, myVessel.IMONumber)
            sit.BackColor = Me.BackColor
            Me.SubItems.Add(sit)
            sit = New ListViewItem.ListViewSubItem(Me, myVessel.Owner)
            sit.BackColor = Me.BackColor
            Me.SubItems.Add(sit)
            sit = New ListViewItem.ListViewSubItem(Me, myVessel.VesselStatus)
            sit.BackColor = Me.BackColor
            Me.SubItems.Add(sit)


        End Sub

        Public Sub New(ByVal inVessel As Intranet.intranet.customerns.Vessel)
            myVessel = inVessel
            Me.Refresh()
        End Sub

        Public Property Vessel() As Intranet.intranet.customerns.Vessel
            Get
                Return myVessel
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.Vessel)
                myVessel = Value
            End Set
        End Property
    End Class

    Public Class LVOldCollectorpayHeader
        Inherits ListViewItem
        Private myCollPayHeader As Intranet.intranet.jobs.CollPay
        Private CollDesc As String = ""
        Private strCollDate As String
        Private strCustomer As String
        Private strVessel As String
        Private strPort As String
        Private myCulture As System.Globalization.CultureInfo

        Public Sub New(ByVal CollPayHeader As Intranet.intranet.jobs.CollPay, ByVal Culture As System.Globalization.CultureInfo)
            MyBase.new()
            myCollPayHeader = CollPayHeader
            myCulture = Culture
            If CollPayHeader.CID.Trim.Length > 0 Then
                strCollDate = CConnection.PackageStringList("Lib_collection.GetCollectionDate", CollPayHeader.CID)
                strCustomer = CConnection.PackageStringList("lib_collectionxml8.GetCustomer", CollPayHeader.CID)
                strCustomer = Medscreen.ReplaceString(strCustomer, "Undefined", "")
                strVessel = CConnection.PackageStringList("Lib_Collection.GetCollectionVessel", CollPayHeader.CID)
                strPort = CConnection.PackageStringList("Lib_Collection.GetCollectionPort", CollPayHeader.CID)
                strPort = Medscreen.ReplaceString(strPort, "Undefined", "")
            End If
            Me.Refresh()

        End Sub

        Public Property CollPayHeader() As Intranet.intranet.jobs.CollPay
            Get
                Return myCollPayHeader
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollPay)
                myCollPayHeader = Value
            End Set
        End Property

        Public Function Refresh() As Boolean
            'Dim objSubItem As ListViewItem.ListViewSubItem

            Me.SubItems.Clear()

            If Not Me.myCollPayHeader Is Nothing Then
                Me.Text = strCollDate
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, strCustomer))
                'check pay type
                If myCollPayHeader.PayType.Trim.Length = 0 OrElse myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_HALFDAY _
                OrElse myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_OTHERCOST Then   ' Null is pay 

                    If myCollPayHeader.Collection.Trim.Length <> 0 Then
                        Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.Collection & "/" & myCollPayHeader.CID))

                    Else
                        If strPort.Trim.Length > 0 Then
                            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, strVessel & "/" & strPort))
                        Else
                            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, strVessel))

                        End If

                    End If
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.Donors))
                ElseIf myCollPayHeader.PayType = "O" Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "OCR"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_TRANSFER Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Transfer Pay"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_AREA Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Area Coordinator Pay"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = "PC" Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Computer Equipment"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = "TEL" Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Telephone costs"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = jobs.CollPay.CST_PAYTYPE_STATIONERY Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Stationery"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = jobs.CollPay.CST_PAYTYPE_HOLIDAYCODE Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Holiday Pay"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                ElseIf myCollPayHeader.PayType = jobs.CollPay.CST_PAYTYPE_DEDUCTION Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, "Pay Deduction"))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))


                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.Collection & "-" & CConnection.PackageStringList("Lib_CollPay.GetPayComment", myCollPayHeader.CID)))
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))

                End If

                'If InStr(myCollPayHeader.Collection, "DED") > 0 Then
                '    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, CDbl(myCollPayHeader.CollTotal * -1).ToString("c", myCulture)))
                '    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, CDbl(myCollPayHeader.ExpensesTotal * -1).ToString("c", myCulture)))
                '    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, CDbl(myCollPayHeader.TotalCost * -1).ToString("c", myCulture)))
                'Else
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.CollTotal.ToString("c", myCulture)))
                'Expenses need to divide out mileage
                Dim Miles As Double = myCollPayHeader.Mileage
                Dim expTotal As Double = myCollPayHeader.ExpensesTotal
                Dim ItemCode As String = myCollPayHeader.ItemCode
                If Miles > 0 Then                    'We have a mileage
                    Dim ExpensesString As String = CDbl(expTotal - myCollPayHeader.TravelCost).ToString("c", myCulture) & " (" & myCollPayHeader.TravelCost.ToString("c", myCulture) & ")travel"
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ExpensesString))
                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.ExpensesTotal.ToString("c", myCulture)))
                End If
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.TotalCost.ToString("c", myCulture)))
                'End If
                If myCollPayHeader.PayType.Trim.Length = 0 OrElse myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_HALFDAY _
                OrElse myCollPayHeader.PayType = Intranet.intranet.jobs.CollPay.CST_PAYTYPE_OTHERCOST Then
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.CID))
                Else
                    Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
                End If
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.InvoiceStatusString))
                Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, myCollPayHeader.Comments))

            End If

        End Function

        Public ReadOnly Property Expenses() As Double
            Get
                Return myCollPayHeader.ExpensesTotal
            End Get
        End Property

        Public ReadOnly Property PayTotal() As Double
            Get
                If InStr(myCollPayHeader.Collection, "DED") > 0 Then
                    Return myCollPayHeader.CollTotal * -1
                Else
                    Return myCollPayHeader.CollTotal
                End If
            End Get
        End Property


        Public Property Total() As Double
            Get
                If InStr(myCollPayHeader.Collection, "DED") > 0 Then
                    Return myCollPayHeader.TotalCost * -1
                Else
                    Return myCollPayHeader.TotalCost
                End If
            End Get
            Set(ByVal Value As Double)

            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.LVCollectorpayHeader
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/09/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LVCollectorpayHeader
        Inherits ListViewItem
        Private myCollPayHeader As Intranet.collectorpay.CollectorPay
        Private CollDesc As String = ""

        Public Sub New(ByVal CollPayHeader As Intranet.collectorpay.CollectorPay)
            MyBase.new()
            myCollPayHeader = CollPayHeader
            Me.Refresh()
        End Sub

        Public Property CollPayHeader() As Intranet.collectorpay.CollectorPay
            Get
                Return myCollPayHeader
            End Get
            Set(ByVal Value As Intranet.collectorpay.CollectorPay)
                myCollPayHeader = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh contents of item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/03/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Dim objSubItem As ListViewItem.ListViewSubItem

            Me.SubItems.Clear()

            If Not Me.myCollPayHeader Is Nothing Then
                If Me.CollPayHeader.PayType = Intranet.collectorpay.CollectorPay.GCST_COLLECTION_PAY Then  'If a collection item draw the CM number in column 1
                    Me.Text = Me.CollPayHeader.CMidentity
                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PaymentMonth.ToString("dd-MMM-yyyy"))
                    Me.SubItems.Add(objSubItem)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PayStatusString)
                    Me.SubItems.Add(objSubItem)
                    'Me.BackColor = Color.AliceBlue
                    'See if we have a description 
                    If Me.CollDesc Is Nothing OrElse Me.CollDesc.Trim.Length = 0 Then 'Try and get description
                        If Me.CollPayHeader.CMidentity.Trim.Length > 0 Then     'Only attempt to get info if a genuine collection 
                            Dim strCollectionBrowseText As String = MedscreenCommonGUIConfig.NodePLSQL.Item("CollectionBrowseText")
                            Me.CollDesc = CConnection.PackageStringList(strCollectionBrowseText, Me.CollPayHeader.CMidentity)
                        End If
                    End If
                    objSubItem = New ListViewItem.ListViewSubItem(Me, CollDesc)
                    Me.SubItems.Add(objSubItem)
                    'We need to add three rows 
                    Dim TotValue As Double = CollPayHeader.PayItems.TotalValue
                    Dim ExpValue As Double = CollPayHeader.PayItems.ExpenseValue
                    objSubItem = New ListViewItem.ListViewSubItem(Me, (TotValue - ExpValue))
                    Me.SubItems.Add(objSubItem)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, (ExpValue))
                    Me.SubItems.Add(objSubItem)

                    objSubItem = New ListViewItem.ListViewSubItem(Me, (TotValue))
                    Me.SubItems.Add(objSubItem)
                ElseIf Me.CollPayHeader.PayType = Intranet.collectorpay.CollectorPay.GCST_OCR_PAY Then
                    Me.Text = "O C R - " & CollPayHeader.CollPayID
                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PaymentMonth.ToString("MMM-yyyy"))
                    Me.SubItems.Add(objSubItem)

                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PayStatusString)
                    Me.SubItems.Add(objSubItem)
                    Me.BackColor = Color.Beige
                Else
                    Me.Text = Me.myCollPayHeader.PaymentMonth.ToString("dd-MMM-yyyy")
                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PayStatusString)
                    Me.SubItems.Add(objSubItem)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.Comment)
                    Me.SubItems.Add(objSubItem)

                    'see if we can dig out a comment
                    If Me.myCollPayHeader.PayItems.Count > 0 Then
                        objSubItem = New ListViewItem.ListViewSubItem(Me, Me.myCollPayHeader.PayItems.Item(0).Comments)
                        Me.SubItems.Add(objSubItem)
                    End If
                End If
                Select Case Me.CollPayHeader.PayStatus
                    Case "E"
                        Me.BackColor = Color.LightBlue
                    Case "R"
                        Me.BackColor = Color.Fuchsia
                    Case "A"
                        Me.BackColor = Color.LimeGreen
                    Case "X"
                        Me.BackColor = Color.Yellow
                    Case "L"
                        Me.BackColor = Color.PowderBlue
                End Select
                If Me.Selected Then
                    Me.ForeColor = Color.White
                    Me.BackColor = Color.Purple
                End If
            End If

        End Function


    End Class

    Public Class lvMoveBarcode
        Inherits ListViewItem
        Private strBarcode As String
        Private strTartgetJob As String
        Private blnInUse As Boolean = True

#Region "procedures"

        Public Sub New(ByVal BarCode As String)
            MyBase.new()
            strBarcode = BarCode
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Me.SubItems.Add(New ListViewItem.ListViewSubItem(Me, ""))
            Redraw()
        End Sub
#End Region

#Region "properties"

        Public ReadOnly Property InUse() As Boolean
            Get
                Return blnInUse
            End Get
        End Property

        Public Property CompanyID() As String
            Get
                Return Me.SubItems(6).Text

            End Get
            Set(ByVal Value As String)
                Me.SubItems(6).Text = Value
            End Set
        End Property

        Public Property SMIDProfile() As String
            Get
                Return Me.SubItems(5).Text
            End Get
            Set(ByVal Value As String)
                Me.SubItems(5).Text = Value
            End Set
        End Property

        Public Property BarCode() As String
            Get
                Return strBarcode
            End Get
            Set(ByVal Value As String)
                strBarcode = Value
            End Set
        End Property

        Public Property TargetJob() As String
            Get
                Return Me.strTartgetJob
            End Get
            Set(ByVal Value As String)
                strTartgetJob = Value
            End Set
        End Property
#End Region

#Region "Functions"
        Public Function Redraw() As Boolean
            Me.Text = strBarcode
            Me.SubItems(4).Text = CConnection.PackageStringList("Lib_sample8.GetJobName", Me.strBarcode)
            Dim strInfo As String = CConnection.PackageStringList("Lib_sample8.GetCustomerInfoFromBarcode", Me.strBarcode)
            If Not strInfo Is Nothing AndAlso strInfo.Trim.Length > 0 Then
                Dim infoArray As String() = strInfo.Split(New Char() {","})
                Me.SubItems(5).Text = infoArray(1)
                Me.SubItems(6).Text = infoArray(0)
                Me.SubItems(7).Text = infoArray(5)
            End If
            ValidateBarcode()
        End Function

        ''' <summary>
        '''   Determines if the passed barcode is valid according to Mescreen's check digit rules.
        ''' </summary>
        ''' <param name="Barcode"></param>
        ''' <returns>Boolean - True if barcode is valid</returns>
        ''' <remarks>
        '''   This routine does NOT cater for knwon mis-printed barcodes or other special cases
        ''' </remarks>
        Public Function IsValidBarcode(ByVal Barcode As String) As Boolean
            Dim blnRet As Boolean = False
            If Barcode.Length < 2 Then
                Return blnRet
                Exit Function
            ElseIf Mid(Barcode, 1, 2) = "HA" Then
                blnRet = Barcode.Length >= 6 AndAlso _
                   Barcode.Length <= 10
            Else
                blnRet = (Not Barcode Is Nothing) AndAlso _
                   IsNumeric(Barcode) AndAlso _
                   Barcode.Length Mod 2 = 0 AndAlso _
                   Barcode.Length >= 6 AndAlso _
                   Barcode.Length <= 10 AndAlso _
                   Barcode.EndsWith(GetCheckDigit(Barcode).ToString)
            End If
            Return blnRet
        End Function

        ''' <summary>
        '''   Returns the correct check digit digit for the passed barcode
        ''' </summary>
        ''' <param name="Barcode"></param>
        ''' <returns>Integer in range 0 - 9, the value of the correct check digit</returns>
        ''' <remarks>
        ''' </remarks>
        Public Function GetCheckDigit(ByVal Barcode As String) As Integer
            Dim charPos As Integer, thisDigit As Integer, total As Integer
            ' Never include check digit (if included in barcode) in calculation
            For charPos = 1 To Barcode.Length - (1 - (Barcode.Length Mod 2))
                thisDigit = CInt(Barcode.Substring(charPos - 1, 1)) * ((charPos Mod 2) + 1)
                If thisDigit > 9 Then thisDigit -= 9
                total += thisDigit
                ' This line is equivalent, but harder to read and offers no performance increase!
                'total += thisDigit + ((thisDigit + (-9 * ((thisDigit * 2) \ 10))) * (charPos Mod 2))
            Next
            Return (10 - (total Mod 10)) Mod 10
        End Function

        Private Function ValidateBarcode() As Boolean
            Dim blnReturn As Boolean = True
            If Me.IsValidBarcode(strBarcode) Then
                Me.ImageIndex = 3
                Me.SubItems(1).Text = "OK"
            Else
                Me.ImageIndex = 2
                Me.SubItems(1).Text = "Failed"
                blnReturn = False
            End If
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(100)
            If blnReturn Then
                If Not ValidateBarcodeInUse() Then
                    If Me.SubItems(4).Text = Me.TargetJob Then
                        Me.ImageIndex = 1
                    Else
                        Me.ImageIndex = 5
                    End If
                    Me.SubItems(2).Text = "In Use"
                    blnReturn = False
                    blnInUse = True
                Else
                    blnInUse = False
                    Me.SubItems(2).Text = "OK"
                    Me.ImageIndex = 4
                End If
            Else
                Me.ImageIndex = 2
                Me.SubItems(2).Text = "not checked"


            End If
            Return blnReturn
        End Function

        Private Function ValidateBarcodeInUse() As Boolean
            Dim ocmd As New OleDb.OleDbCommand()
            Dim objRet As Object = Nothing
            Dim blnReturn As Boolean = True
            Try
                ocmd.CommandText = "Select count(bar_code) FROM View_AllUsedBarcodes WHERE Bar_Code = ?"
                ocmd.Connection = CConnection.DbConnection
                ocmd.Parameters.Add(CConnection.StringParameter("Barcode", strBarcode, 10))
                If CConnection.ConnOpen Then
                    objRet = ocmd.ExecuteScalar
                End If
                blnReturn = (CInt(objRet) = 0)
            Catch ex As Exception
                blnReturn = False
            Finally
                CConnection.SetConnClosed()
            End Try
            Return blnReturn
        End Function


#End Region

    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.LVCofficerItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collecting officer List view item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class LVCofficerItem
        Inherits ListViewItem
        Private myPortId As String
        ' Private blnAgreed As Boolean
        ' Private myCost As Double
        Private MyOfficer As Intranet.intranet.jobs.CollectingOfficer
        Private myPortColl As Intranet.intranet.jobs.CollPort
        Private myPortCollPlus As BOSampleManager.CollectorPort
        Private myCollectionDate As Date


        Public Enum DisplayOptions
            standard
            search
        End Enum

        Private myDisplayOption As DisplayOptions = DisplayOptions.standard

        Public Property DisplayOption() As DisplayOptions
            Get
                Return myDisplayOption
            End Get
            Set(ByVal Value As DisplayOptions)
                myDisplayOption = Value
                Refresh()
            End Set
        End Property

        Public Sub New(ByVal Cofficer As Intranet.intranet.jobs.CollectingOfficer, ByVal PortId As String, ByVal CollectionDate As Date)
            MyBase.New()
            MyOfficer = Cofficer
            myCollectionDate = CollectionDate
            myPortId = PortId
            Me.Refresh()
        End Sub

        Public Sub New(ByVal collport As BOSampleManager.CollectorPort, ByVal CollectionDate As Date)
            MyBase.new()
            MyOfficer = New Intranet.intranet.jobs.CollectingOfficer(collport.OfficerID)
            myCollectionDate = CollectionDate
            myPortId = collport.PortID
            myPortCollPlus = collport
            Me.Refresh()


        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw Officer Item
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/03/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Try
                If Me.myDisplayOption = DisplayOptions.standard Then
                    MyOfficer.Load()
                    Me.Text = MyOfficer.OfficerID
                    'Me.UseItemStyleForSubItems = True
                    Dim sItem As SiteSubItem
                    sItem = New SiteSubItem()
                    sItem.Text = Me.MyOfficer.FirstName
                    Me.BackColor = System.Drawing.SystemColors.Info
                    Me.SubItems.Add(sItem)
                    Me.SubItems.Add(MyOfficer.Surname)

                    Dim blnret As Boolean = MyOfficer.Holidays.HolidayIn(myCollectionDate)
                    If MyOfficer.Holidays.HolidayIn(myCollectionDate) Then
                        Me.Font = New Drawing.Font(System.Drawing.FontFamily.GenericSansSerif, 8, Drawing.FontStyle.Strikeout, Drawing.GraphicsUnit.Point)
                        Me.ImageIndex = 1

                        If MyOfficer.IsPinPerson Then
                            Me.Checked = True
                            Me.BackColor = System.Drawing.Color.Red
                            Me.ForeColor = System.Drawing.Color.Yellow
                            Me.SubItems.Add(MyOfficer.IsPinPerson)
                        Else
                            Me.BackColor = System.Drawing.Color.Red
                            Me.ForeColor = System.Drawing.Color.Yellow
                            Me.SubItems.Add(" ")
                        End If
                    Else
                        If MyOfficer.IsPinPerson Then
                            Me.Checked = True
                            Me.ImageIndex = 3
                            Me.BackColor = System.Drawing.Color.PaleGreen
                            Me.SubItems.Add(MyOfficer.IsPinPerson)
                        Else
                            Me.SubItems.Add(" ")
                            Me.ImageIndex = 0

                        End If
                    End If

                    If myPortId.Trim.Length > 0 Then
                        myPortColl = MyOfficer.Ports.Item(myPortId, 0)
                        If Not myPortColl Is Nothing Then
                            Me.SubItems.Add(myPortColl.Costs)
                            If myPortColl.Agreed Then
                                Me.SubItems.Add("Costs agreed")
                            Else
                                Me.SubItems.Add("Costs estimated")

                            End If
                        End If
                    End If
                    If myPortCollPlus Is Nothing Then
                        myPortCollPlus = New BOSampleManager.CollectorPort(Me.OfficerId, myPortId)
                    End If
                    
                    If myPortCollPlus IsNot Nothing Then
                        Me.SubItems.Add(myPortCollPlus.AverageExpense.ToString("#,##0.00"))
                        Dim avgcps As String = myPortCollPlus.AverageCostPerDonor.ToString
                        Me.SubItems.Add(myPortCollPlus.AverageCostPerDonor.ToString("#,##0.00"))
                        Me.SubItems.Add(myPortCollPlus.AverageCost.ToString("#,##0.00"))
                        Me.SubItems.Add(myPortCollPlus.NoCollections)
                        Me.SubItems.Add(myPortCollPlus.DistanceFromBase)
                    Else
                        'Get number of collection in last 30 days.
                        Dim strNoColl As String = CConnection.PackageStringList("LIB_OFFICER.OfficerJobsInPeriod", Me.OfficerId)
                        Me.SubItems.Add(strNoColl)

                    End If
                Else            'Not Standard
                    Me.Text = MyOfficer.Surname
                    If MyOfficer.Removed Then
                        Dim newFont As New Font("Arial", 8, FontStyle.Strikeout Or FontStyle.Italic, GraphicsUnit.Point)
                        Me.Font = newFont
                        Me.ForeColor = Color.DarkBlue
                        Me.BackColor = Color.Beige
                    End If
                    Dim aSubItem As ListViewItem.ListViewSubItem
                    aSubItem = New ListViewItem.ListViewSubItem(Me, MyOfficer.FirstName)
                    Me.SubItems.Add(aSubItem)
                    aSubItem = New ListViewItem.ListViewSubItem(Me, MyOfficer.Region)
                    Me.SubItems.Add(aSubItem)
                    Dim strLookupCountry As String = MedscreenCommonGUIConfig.NodePLSQL.Item("LookupCountry")
                    Dim Country As String = CConnection.PackageStringList(strLookupCountry, MyOfficer.CountryId)
                    aSubItem = New ListViewItem.ListViewSubItem(Me, Country)
                    Me.SubItems.Add(aSubItem)

                    Dim oRead As Object = Nothing
                    Dim oCmd As New OleDb.OleDbCommand()
                    oCmd.CommandText = "Select profile from customer where identity = ?"
                    oCmd.Parameters.Add(CConnection.StringParameter("id", MyOfficer.Division, 10))

                    Try ' Protecting 
                        oCmd.Connection = CConnection.DbConnection
                        If CConnection.ConnOpen Then            'Attempt to open reader
                            oRead = oCmd.ExecuteScalar
                        End If
                        If Not oRead Is System.DBNull.Value AndAlso Not oRead Is Nothing Then
                            aSubItem = New ListViewItem.ListViewSubItem(Me, oRead)
                            Me.SubItems.Add(aSubItem)

                        End If
                    Catch ex As Exception
                        MedscreenLib.Medscreen.LogError(ex, , "LVCofficerItem-Refresh-4331")
                    Finally
                        CConnection.SetConnClosed()             'Close connection
                    End Try

                End If
            Catch ex As Exception
            End Try
        End Function

        Public ReadOnly Property Agreed() As Boolean
            Get
                If Not Me.myPortColl Is Nothing Then
                    Return myPortColl.Agreed
                Else
                    Return False
                End If
            End Get

        End Property

        Public Property CollectionPort() As Intranet.intranet.jobs.CollPort
            Get
                Return Me.myPortColl
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollPort)
                Me.myPortColl = Value
            End Set
        End Property

        Public Property Officer() As Intranet.intranet.jobs.CollectingOfficer
            Get
                Return MyOfficer
            End Get
            Set(ByVal Value As Intranet.intranet.jobs.CollectingOfficer)
                Me.MyOfficer = Value
            End Set
        End Property

        Public ReadOnly Property Cost() As Double
            Get
                If Not Me.myPortColl Is Nothing Then
                    Return myPortColl.Costs
                Else
                    Return 0
                End If
            End Get

        End Property

        Public ReadOnly Property LastName() As String
            Get
                Return Me.MyOfficer.Surname
            End Get

        End Property

        Public ReadOnly Property FirstName() As String
            Get
                Return MyOfficer.FirstName
            End Get

        End Property

        Public ReadOnly Property Mobile() As String
            Get
                Return Me.MyOfficer.MobilePhone
            End Get
        End Property

        Public ReadOnly Property HomePhone() As String
            Get
                Return Me.MyOfficer.HomePhone
            End Get

        End Property

        Public ReadOnly Property OfficerId() As String
            Get
                Return MyOfficer.OfficerID
            End Get
        End Property

        Public ReadOnly Property OfficePhone() As String
            Get
                Return Me.MyOfficer.OfficePhone
            End Get

        End Property

        Public ReadOnly Property Beeper() As String
            Get
                Return Me.MyOfficer.Bleeper
            End Get

        End Property

        Public Sub New()

        End Sub
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.JobStatusSubItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A job status sub item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class JobStatusSubItem
        Inherits ListViewItem.ListViewSubItem

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Creat enew Job Status sub item 
        ''' </summary>
        ''' <param name="status">Status </param>
        ''' <param name="strIn">Invoice number</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal status As Object, ByVal strIn As String)
            MyBase.new()
            If strIn Is Nothing Then
                strIn = ""
            End If
            If strIn.Trim.Length > 8 Then
                If status Is System.DBNull.Value Then
                    Me.Text = "Active"
                Else
                    Me.Text = "Committed"
                End If
            ElseIf strIn.Trim.Length > 0 Then
                Me.Text = "Invoiced"
            Else
                Try
                    If status Is System.DBNull.Value Then
                        Me.Text = ""
                    ElseIf status = "SAMPLE" Then
                        Me.Text = ""
                    ElseIf status = 0 Then
                        Me.Text = ""
                    Else
                        Me.Text = "Committed"
                    End If
                Catch ex As Exception
                End Try
            End If

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.CollStatusSubItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection status sub item 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CollStatusSubItem
        Inherits ListViewItem.ListViewSubItem
        Private myBackColor As System.Drawing.Color
        Private myForeColor As System.Drawing.Color
        Private strStatus As String = ""
        Private Invoice As String = ""
        Private myUserColours As MedscreenCommonGui.UserDefaults.UserColourCollection

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Fore colour of item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ForeColour() As System.Drawing.Color
            Get
                Return Me.myForeColor
            End Get
            Set(ByVal Value As System.Drawing.Color)
                myForeColor = Value
                Me.ForeColor = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Back Colour of item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property BackColour() As System.Drawing.Color
            Get
                Return Me.myBackColor
            End Get
            Set(ByVal Value As System.Drawing.Color)
                Me.BackColor = Value
                myBackColor = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection status sub item
        ''' </summary>
        ''' <param name="status">Status</param>
        ''' <param name="strInvoice">Invoice number</param>
        ''' <param name="mUserColours">Colours</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal status As Object, ByVal strInvoice As String, _
                ByVal mUserColours As UserDefaults.UserColourCollection)
            MyBase.new()
            strStatus = status
            Invoice = strInvoice
            Me.UserColours = mUserColours
            SetStyle()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Status associated with item
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Status() As String
            Get
                Return strStatus
            End Get
            Set(ByVal Value As String)
                strStatus = Value
                SetStyle()
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' User colours 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property UserColours() As UserDefaults.UserColourCollection
            Get
                Return myUserColours
            End Get
            Set(ByVal Value As UserDefaults.UserColourCollection)
                myUserColours = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Set the style of the item 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [10/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Sub SetStyle()
            Me.BackColour = System.Drawing.Color.Wheat
            Me.ForeColour = System.Drawing.SystemColors.ControlText

            Dim uc As UserDefaults.UserColours

            If Invoice.Trim.Length > 0 Then
                Me.Text = "Invoiced - "
                uc = Me.UserColours.Item("I")
                If Not uc Is Nothing Then
                    Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                    Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                Else
                    Me.BackColour = System.Drawing.Color.LightBlue
                    Me.ForeColour = System.Drawing.Color.Black
                End If
                Select Case Status
                    Case "C"
                        Me.Text += "Confirmed C"
                    Case "A"
                        Me.Text += "Approved A"
                    Case "P"
                        Me.Text += "Sent P"
                    Case "V"
                        Me.Text += "Available V"
                    Case "X"
                        Me.Text += "Cancelled X"
                    Case Constants.GCST_JobStatusOnHold
                        Me.Text += "OnHold"
                    Case Constants.GCST_JobStatusCollected
                        Me.Text += "Samples Collected"
                    Case Constants.GCST_JobStatusReceived
                        Me.Text += "Samples Received"
                    Case Else
                        Me.Text += "Status " & Status
                End Select
            Else
                Select Case strStatus
                    Case "C"

                        Me.Text = "Confirmed C"
                        uc = Me.UserColours.Item("C")
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.ForeColour = System.Drawing.Color.Green
                            Me.BackColour = Drawing.Color.White
                        End If

                    Case "A"
                        Me.Text = "Approved A"
                        uc = Me.UserColours.Item("A")
                        If Not uc Is Nothing Then
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                        Else
                            Me.ForeColour = System.Drawing.Color.SeaGreen
                        End If
                    Case "P"
                        Me.Text = "Sent P"
                        uc = Me.UserColours.Item("P")
                        If Not uc Is Nothing Then
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                        Else

                            Me.ForeColour = System.Drawing.Color.Blue
                        End If
                    Case "V"
                        Me.Text = "Available V"
                        uc = Me.UserColours.Item("V")
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(250, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            'Debug.WriteLine(Me.BackColour.ToArgb)
                            'Me.BackColour = Drawing.Color.LightBlue
                            'Debug.WriteLine(Me.BackColour.ToArgb)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.ForeColour = System.Drawing.Color.Black
                        End If
                    Case "X"
                        Me.Text = "Cancelled X"
                        uc = Me.UserColours.Item("X")
                        If Not uc Is Nothing Then
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                        Else

                            Me.BackColour = System.Drawing.Color.Red
                            Me.ForeColour = System.Drawing.Color.Yellow
                        End If
                    Case Constants.GCST_JobStatusOnHold
                        uc = Me.UserColours.Item("H")
                        Me.Text = "On Hold"
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.BackColour = System.Drawing.Color.Yellow
                            Me.ForeColour = System.Drawing.Color.Tomato
                        End If
                    Case Constants.GCST_JobStatusCommitted
                        uc = Me.UserColours.Item(Constants.GCST_JobStatusCommitted)
                        Me.Text = "Committed"
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.BackColour = System.Drawing.Color.Wheat
                            Me.ForeColour = System.Drawing.Color.Black
                        End If
                    Case Constants.GCST_JobStatusCollected
                        Me.Text += "Samples Collected"
                        uc = Me.UserColours.Item(Constants.GCST_JobStatusCollected)
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.BackColour = System.Drawing.Color.Wheat
                            Me.ForeColour = System.Drawing.Color.Black
                        End If
                    Case Constants.GCST_JobStatusReceived
                        Me.Text += "Samples Received"
                        uc = Me.UserColours.Item(Constants.GCST_JobStatusReceived)
                        If Not uc Is Nothing Then
                            Me.BackColour = Drawing.Color.FromArgb(255, uc.BackRed, uc.BackGreen, uc.BackBlue)
                            Me.ForeColour = Drawing.Color.FromArgb(255, uc.foreRed, uc.foreGreen, uc.foreBlue)
                        Else
                            Me.BackColour = System.Drawing.Color.Wheat
                            Me.ForeColour = System.Drawing.Color.Black
                        End If
                    Case Else
                        Me.Text = "Status " & strStatus
                        Me.ForeColour = System.Drawing.Color.GreenYellow
                End Select
            End If

        End Sub
    End Class

End Namespace

Namespace ListViewItems.DiscountSchemes

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.DiscountSchemes.lvDiscountScheme
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' discount scheme class, deals with discount scheme entries
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvDiscountScheme
        Inherits lvDsItem

        Private myDiscountScheme As DiscountSchemeEntry
        Private objPrice As MedscreenLib.Glossary.LineItemPrice
        Private myClient As Intranet.intranet.customerns.Client

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh display of item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            If Not Me.myDiscountScheme Is Nothing Then
                Me.Text = Me.myDiscountScheme.DiscountSchemeId & "-" & _
                    Me.myDiscountScheme.EntryNumber
                Dim aSub As ListViewItem.ListViewSubItem
                aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.DiscountType)
                Me.SubItems.Add(aSub)
                'Deal with Code type items 
                If Me.myDiscountScheme.DiscountType = "CODE" Then
                    Dim strLineItem As String = Me.myDiscountScheme.ApplyTo
                    Dim objLineItem As MedscreenLib.Glossary.LineItem = MedscreenLib.Glossary.Glossary.LineItems.Item(strLineItem)
                    If Not objLineItem Is Nothing Then
                        aSub = New ListViewItem.ListViewSubItem(Me, objLineItem.Service & " " & objLineItem.Type & _
                        " " & objLineItem.Description & " (" & Me.myDiscountScheme.ApplyTo & ")")
                        Me.SubItems.Add(aSub)
                    Else
                        aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.ApplyTo)
                        Me.SubItems.Add(aSub)
                    End If
                Else
                    aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.ApplyTo)
                    Me.SubItems.Add(aSub)
                End If
                Dim tmpDisc As Double = Me.myDiscountScheme.Discount
                If tmpDisc < 0 Then
                    Me.UseItemStyleForSubItems = False      'Set so each sub item can have a different style
                    'Make the item negative accountancy style
                    aSub = New ListViewItem.ListViewSubItem(Me, "(" & Math.Abs(Me.myDiscountScheme.Discount) & ")")
                    aSub.BackColor = Color.Red      'Make it stand out 
                    aSub.ForeColor = Color.Yellow
                Else
                    aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.Discount)
                End If
                Me.SubItems.Add(aSub)

                'Deal with discounts 
                If Not objPrice Is Nothing Then
                    aSub = New ListViewItem.ListViewSubItem(Me, objPrice.Price.ToString("c", myClient.Culture))
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, (objPrice.Price - (objPrice.Price * tmpDisc / 100)).ToString("c", myClient.Culture))
                    Me.SubItems.Add(aSub)
                End If

            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new dicsount scheme list view item and display 
        ''' </summary>
        ''' <param name="DiscountScheme"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal DiscountScheme As DiscountSchemeEntry, ByVal client As Intranet.intranet.customerns.Client)
            myDiscountScheme = DiscountScheme
            myClient = client               'Set local variable
            If Not myDiscountScheme Is Nothing AndAlso _
            myDiscountScheme.DiscountType = "CODE" AndAlso Not myClient Is Nothing Then
                objPrice = MedscreenLib.Glossary.Glossary.PriceList.Item(myDiscountScheme.ApplyTo, Now, myClient.CurrencyCode)

            End If

            MyClass.Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose Discount scheme
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiscountScheme() As DiscountSchemeEntry
            Get
                Return myDiscountScheme
            End Get
            Set(ByVal Value As DiscountSchemeEntry)
                myDiscountScheme = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.DiscountSchemes.lvCustDiscountScheme
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Class for a list view item that displays header details 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class lvCustDiscountScheme
        Inherits lvDsItem

        Private myDiscountScheme As Intranet.intranet.customerns.CustomerDiscount

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh the contents of the item 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean
            Me.SubItems.Clear()
            Me.BackColor = Drawing.Color.Wheat

            If Not Me.myDiscountScheme Is Nothing Then
                Me.Text = Me.myDiscountScheme.DiscountSchemeId
                If Not Me.myDiscountScheme.DiscountScheme Is Nothing Then


                    Dim aSub As ListViewItem.ListViewSubItem
                    aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.DiscountScheme.Description)
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.ContractAmendment)
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, "")
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, "")
                    Me.SubItems.Add(aSub)
                    aSub = New ListViewItem.ListViewSubItem(Me, "")
                    Me.SubItems.Add(aSub)
                    If Me.myDiscountScheme.Startdate > DateField.ZeroDate Then
                        aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.Startdate.ToString("dd-MMM-yyyy"))
                    Else
                        aSub = New ListViewItem.ListViewSubItem(Me, "")
                    End If
                    Me.SubItems.Add(aSub)
                    If Me.myDiscountScheme.EndDate > DateField.ZeroDate Then
                        aSub = New ListViewItem.ListViewSubItem(Me, Me.myDiscountScheme.EndDate.ToString("dd-MMM-yyyy"))
                    Else
                        aSub = New ListViewItem.ListViewSubItem(Me, "")
                    End If
                    Me.SubItems.Add(aSub)
                End If
                If Me.myDiscountScheme.Startdate > Today Then
                    Me.BackColor = Drawing.Color.Red
                    Me.ForeColor = Color.White
                    Me.Text += " Active from - " & Me.myDiscountScheme.Startdate.ToString("dd-MMM-yyyy")
                End If
                If Me.myDiscountScheme.EndDate < Today And Me.myDiscountScheme.EndDate <> DateField.ZeroDate Then
                    Me.BackColor = Drawing.Color.Red
                    Me.ForeColor = Color.White
                    Me.Text += " Disabled from - " & Me.myDiscountScheme.EndDate.ToString("dd-MMM-yyyy")
                End If
            End If
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new discount scheme list view item
        ''' </summary>
        ''' <param name="DiscountScheme"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal DiscountScheme As Intranet.intranet.customerns.CustomerDiscount)
            myDiscountScheme = DiscountScheme
            MyClass.Refresh()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose discount scheme
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiscountScheme() As Intranet.intranet.customerns.CustomerDiscount
            Get
                Return myDiscountScheme
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.CustomerDiscount)
                myDiscountScheme = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.DiscountSchemes.lvDsItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Base class for discount scheme list view items
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public MustInherit Class lvDsItem
        Inherits ListViewItem
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.DiscountSchemes.lvVolDsItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Base class for Volume Discount schemes 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/05/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public MustInherit Class lvVolDsItem
        Inherits lvDsItem
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : ListViewItems.DiscountSchemes.TestDiscountEntry
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Represents as discount on as test
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class TestDiscountEntry
        Inherits ListViewItem
        Private myClient As Intranet.intranet.customerns.Client
        Private myOfficer As Intranet.intranet.jobs.CollectingOfficer
        Private myItemCode As String
        Private myObjLineItem As MedscreenLib.Glossary.LineItem
        Private intSold As Integer = 0
        Private intKitSold As Integer = 0
        Private dblPrice As Double = 0
        Private strDiscountId As String = ""
        Private blnCustomer As Boolean = True

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new item 
        ''' </summary>
        ''' <param name="Client"></param>
        ''' <param name="LineItem"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Client As Intranet.intranet.customerns.Client, _
            ByVal LineItem As String, Optional ByVal sold As Integer = 0, _
            Optional ByVal KitSold As Integer = 0, Optional ByVal Price As Double = 0)
            myClient = Client
            myItemCode = LineItem
            intSold = sold
            intKitSold = KitSold
            dblPrice = Price
            myObjLineItem = MedscreenLib.Glossary.Glossary.LineItems.Item(myItemCode)
            If Not myObjLineItem Is Nothing Then
                Me.Text = myObjLineItem.Description

                Me.Refresh()
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a collecting officer discount
        ''' </summary>
        ''' <param name="Officer"></param>
        ''' <param name="LineItem"></param>
        ''' <param name="sold"></param>
        ''' <param name="KitSold"></param>
        ''' <param name="Price"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	25/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Officer As Intranet.intranet.jobs.CollectingOfficer, _
    ByVal LineItem As String, Optional ByVal sold As Integer = 0, _
    Optional ByVal KitSold As Integer = 0, Optional ByVal Price As Double = 0)
            myOfficer = Officer
            myItemCode = LineItem
            intSold = sold
            intKitSold = KitSold
            dblPrice = Price
            myObjLineItem = MedscreenLib.Glossary.Glossary.LineItems.Item(myItemCode)
            Me.blnCustomer = False
            If Not myObjLineItem Is Nothing Then
                Me.Text = myObjLineItem.Description

                Me.Refresh()
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Refresh view 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Refresh() As Boolean

            'Separators
            Dim backColour As System.Drawing.Color = SystemColors.Control
            'Active cells
            Dim backColourA As System.Drawing.Color = SystemColors.Control
            If intSold > 0 Then
                backColourA = Drawing.Color.PaleGreen
            Else
                backColourA = Drawing.SystemColors.Window
            End If

            Me.SubItems.Clear()
            Me.UseItemStyleForSubItems = False
            If myObjLineItem Is Nothing Then Exit Function
            If Me.myClient Is Nothing And Me.myOfficer Is Nothing Then Exit Function
            Me.Text = myObjLineItem.Description
            Me.BackColor = backColourA
            Dim objSubItem As ListViewItem.ListViewSubItem
            objSubItem = New ListViewItem.ListViewSubItem(Me, "")
            Me.SubItems.Add(objSubItem)
            objSubItem.BackColor = backColour
            'See if we have a client
            If Not myClient Is Nothing Then
                objSubItem = New ListViewItem.ListViewSubItem(Me, myClient.InvoiceInfo.Price(myItemCode).ToString("c", myClient.Culture))
            End If
            'see if we have an officer
            If Not myOfficer Is Nothing Then
                objSubItem = New ListViewItem.ListViewSubItem(Me, myOfficer.PriceForItem(myItemCode, Now).ToString("c", myOfficer.Culture))
            End If
            Me.SubItems.Add(objSubItem)
            objSubItem.BackColor = backColourA
            objSubItem = New ListViewItem.ListViewSubItem(Me, "")
            Me.SubItems.Add(objSubItem)
            objSubItem.BackColor = backColour
            Dim disc As Double
            'See if we have a client
            If Not myClient Is Nothing Then
                disc = myClient.InvoiceInfo.Discount(myItemCode)
            End If
            'see if we have an officer
            If Not myOfficer Is Nothing Then
                disc = Me.myOfficer.DiscountONItem(myItemCode, Now)
            End If
            'Convert to %
            Dim pDisc As Double = (1 - disc) * 100
            'If disc >= 1 Then disc()
            If Not myOfficer Is Nothing Then
                If disc = 1 Then
                    'No discount in place
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " -")                   'Blank out 0 discounts
                    objSubItem.BackColor = Color.Yellow
                    objSubItem.ForeColor = Color.Black
                ElseIf disc = 0 Then                  '100% discount
                    objSubItem.BackColor = Color.Red
                    objSubItem.ForeColor = Color.Yellow

                Else
                    'Discount exists
                    objSubItem = New ListViewItem.ListViewSubItem(Me, pDisc.ToString("0.00"))
                    objSubItem.BackColor = backColourA

                End If
            Else
                If disc = 0 Then
                    'No discount in place
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " -")                   'Blank out 0 discounts
                    objSubItem.BackColor = Color.Yellow
                    objSubItem.ForeColor = Color.Black
                ElseIf disc = 100 Then                  '100% discount
                    objSubItem.BackColor = Color.Red
                    objSubItem.ForeColor = Color.Yellow

                Else
                    'Discount exists
                    objSubItem = New ListViewItem.ListViewSubItem(Me, pDisc.ToString("0.00"))
                    objSubItem.BackColor = backColourA

                End If
            End If
            Me.SubItems.Add(objSubItem)
            objSubItem = New ListViewItem.ListViewSubItem(Me, "")
            Me.SubItems.Add(objSubItem)
            objSubItem.BackColor = backColour
            'See if we have a client
            If Not myClient Is Nothing Then
                objSubItem = New ListViewItem.ListViewSubItem(Me, myClient.InvoiceInfo.DiscountedPrice(myItemCode))
            End If
            'see if we have an officer
            If Not myOfficer Is Nothing Then
                objSubItem = New ListViewItem.ListViewSubItem(Me, myOfficer.PayForItem(myItemCode, Now).ToString("c", myOfficer.Culture))
            End If

            objSubItem.BackColor = backColourA
            Me.SubItems.Add(objSubItem)
            If Me.blnCustomer Then                  'If we are dealing with a customer items sold makes sense, not for officers.
                If intSold > 0 Then
                    objSubItem = New ListViewItem.ListViewSubItem(Me, CStr(intSold))
                    objSubItem.BackColor = backColourA
                Else
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " ")
                End If
            Else
                objSubItem = New ListViewItem.ListViewSubItem(Me, " ")
                Dim pColl As New Collection()
                pColl.Add(Me.myItemCode)
                pColl.Add(Me.myOfficer.OfficerID)
                pColl.Add(Intranet.intranet.customerns.CustomerDiscount.GCST_Officer_link)
                Dim strGetDiscountSchemes As String = MedscreenCommonGUIConfig.NodePLSQL.Item("GetDiscountSchemes")
                Me.strDiscountId = CConnection.PackageStringList(strGetDiscountSchemes, pColl)
                If Not Me.strDiscountId Is Nothing AndAlso strDiscountId <> "-1" Then
                    pColl = New Collection()
                    pColl.Add(Me.strDiscountId)
                    pColl.Add(Me.myOfficer.OfficerID)
                    pColl.Add(Intranet.intranet.customerns.CustomerDiscount.GCST_Officer_link)
                    Dim strGetCustomerSchemeDates As String = MedscreenCommonGUIConfig.NodePLSQL.Item("GetCustomerSchemeDates")
                    Dim strDates As String = CConnection.PackageStringList(strGetCustomerSchemeDates, pColl)
                    objSubItem.Text = strDates & " , scheme id " & strDiscountId
                End If
            End If
            Me.SubItems.Add(objSubItem)
            'Different between officer and customer.
            If Me.blnCustomer Then
                If intSold > 0 Then
                    objSubItem = New ListViewItem.ListViewSubItem(Me, CStr(intKitSold))
                    objSubItem.BackColor = backColourA
                Else
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " ")
                End If
                Me.SubItems.Add(objSubItem)
                If intSold > 0 Then
                    objSubItem = New ListViewItem.ListViewSubItem(Me, CStr(Me.dblPrice))
                    objSubItem.BackColor = backColourA
                Else
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " ")
                End If
                Me.SubItems.Add(objSubItem)
            Else
                'Deal with officer banded payments
                If Not myOfficer Is Nothing Then
                    Dim oColl As New Collection()
                    oColl.Add(Me.myOfficer.OfficerID)
                    oColl.Add(Me.LineItem.ItemCode)
                    Dim strR As String = CConnection.PackageStringList("Lib_collpay.HasLineVolumeDiscount", oColl)
                    objSubItem = New ListViewItem.ListViewSubItem(Me, " ")
                    If strR <> "F" Then
                        objSubItem.Text = "+," & strR
                        objSubItem.BackColor = System.Drawing.Color.LightGreen
                        Me.Checked = True
                    End If
                    Me.SubItems.Add(objSubItem)
                End If
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Find the discount scheme
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	20/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function DiscountScheme() As String
            Dim strTemp As String = ""
            If Not Me.myClient Is Nothing Then
                strTemp = Me.myClient.InvoiceInfo.DiscountScheme(Me.myItemCode)
            Else
                Dim oColl As New Collection()
                oColl.Add(Me.myItemCode)
                oColl.Add(Me.myOfficer.OfficerID)

                oColl.Add(Intranet.intranet.customerns.CustomerDiscount.GCST_Officer_link)
                Dim strGetDiscountSchemes As String = MedscreenCommonGUIConfig.NodePLSQL.Item("GetDiscountSchemes")

                strTemp = CConnection.PackageStringList(strGetDiscountSchemes, oColl)
            End If
            Return strTemp
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Line item encapsulated by this list item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [12/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property LineItem() As MedscreenLib.Glossary.LineItem
            Get
                Return Me.myObjLineItem
            End Get
            Set(ByVal Value As MedscreenLib.Glossary.LineItem)
                Me.myObjLineItem = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Find Price 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [13/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function GetPrice() As Double
            Dim strTemp As String
            strTemp = Me.SubItems(6).Text
            Return CDbl(strTemp)
        End Function
    End Class

    Public Class lvVolumeDiscountEntry
        Inherits lvVolDsItem

        Private myScheme As VolumeDiscountEntry
        Private myClient As Intranet.intranet.customerns.Client

        Public Sub New(ByVal VDEntry As VolumeDiscountEntry, ByVal Client As Intranet.intranet.customerns.Client)
            myScheme = VDEntry
            myClient = Client
            'Now redraw item 
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Dim myReturn As Boolean = False
            Me.SubItems.Clear()
            Me.Text = ""
            If Not Me.myScheme Is Nothing Then
                Me.Text = Me.myScheme.DiscountSchemeId & "-" & Me.myScheme.EntryNumber
                Dim objEntry As ListViewItem.ListViewSubItem
                'Type 
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.Type)
                Me.SubItems.Add(objEntry)
                'Apply to 
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.ApplyTo)
                Me.SubItems.Add(objEntry)
                'Min Samples
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.MinSamples)
                Me.SubItems.Add(objEntry)
                'Max Samples
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.MaxSamples)
                Me.SubItems.Add(objEntry)
                'Min Charge
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.MinChargeDiscount.ToString("0.00"))
                Me.SubItems.Add(objEntry)
                'Sample Discount
                objEntry = New ListViewItem.ListViewSubItem(Me, myScheme.SampleDiscount.ToString("0.00"))
                Me.SubItems.Add(objEntry)

                'Do example 
                If myClient Is Nothing Then Exit Function 'Check we have a customer 
                Dim CurrencyCode As String = myClient.CurrencyCode
                Dim objPrice As LineItemPrice = MedscreenLib.Glossary.Glossary.PriceList.Item("APCOLS002", Now, myClient.CurrencyCode)
                Dim MinCharge As LineItemPrice = MedscreenLib.Glossary.Glossary.PriceList.Item("APCOL0003", Now, CurrencyCode)
                'Deal with discounts to sample and the minimum charge
                Dim minChargeDiscount As Object = Nothing
                Dim SampleDiscount As Object = Nothing

                If objPrice Is Nothing Then Exit Function
                Dim ocmd As New OleDb.OleDbCommand()
                ocmd.Connection = CConnection.DbConnection
                Try
                    CConnection.SetConnOpen()
                    ocmd.CommandText = "Select pckwarehouse.getdiscount(?,?) from dual"
                    ocmd.Parameters.Add(CConnection.StringParameter("CustId", myClient.Identity, 10))
                    ocmd.Parameters.Add(CConnection.StringParameter("ItemCode", "APCOL0003", 10))
                    minChargeDiscount = ocmd.ExecuteScalar
                    If minChargeDiscount Is Nothing OrElse minChargeDiscount Is DBNull.Value Then
                        minChargeDiscount = 1
                    End If
                    ocmd.Parameters(1).Value = "APCOLS002"
                    SampleDiscount = ocmd.ExecuteScalar
                    If SampleDiscount Is Nothing OrElse SampleDiscount Is DBNull.Value Then
                        SampleDiscount = 1
                    End If
                Catch ex As Exception
                Finally
                    CConnection.SetConnClosed()
                End Try

                'Dim Price As Double = objPrice.Price * SampleDiscount
                'Dim ActualPrice As Double = myScheme.AdditionalCharge
                'Dim ActMinCharge As Double = (MinCharge.Price * minChargeDiscount) * (1 - myScheme.MinChargeDiscount / 100) + myScheme.AdditionalCharge
                'Dim actualPriceMin As Double = ActualPrice + (Price * (1 - myScheme.SampleDiscount / 100)) * myScheme.MinSamples
                'Dim actualPriceMax As Double = ActualPrice + (Price * (1 - myScheme.SampleDiscount / 100)) * myScheme.MaxSamples

                'If actualPriceMin < ActMinCharge Then actualPriceMin = ActMinCharge
                'If actualPriceMax < ActMinCharge Then actualPriceMax = ActMinCharge

                'Dim strSamples As String
                'Dim strRet As String = ""
                'If myScheme.Type = "COUNTRY" Then         'Country based discount scheme
                '    Dim objCountry As MedscreenLib.Address.Country = MedscreenLib.Glossary.Glossary.Countries.Item(CInt(myScheme.ApplyTo), 0)
                '    strSamples = "Charge for samples in "
                '    If objCountry Is Nothing Then
                '        strSamples += myScheme.ApplyTo
                '    Else
                '        strSamples += objCountry.CountryName
                '    End If
                '    strRet += strSamples & " is " & (Price * (1 - myScheme.SampleDiscount / 100)).ToString("c", myClient.Culture) & " per sample"
                'ElseIf myScheme.Type = "COLL_TYPE" Then
                '    strSamples = "Charge for samples in " & myScheme.ApplyTo & " collections is "
                '    strRet += strSamples & " is " & (Price * (1 - myScheme.SampleDiscount / 100)).ToString("c", myClient.Culture) & " per sample"
                'ElseIf myScheme.Type = "PORT" Then
                '    strSamples = "Charge for samples collected at " & myScheme.ApplyTo & " is "
                '    strRet += strSamples & (Price * (1 - myScheme.SampleDiscount / 100)).ToString("c", myClient.Culture) & " per sample"
                'ElseIf myScheme.Type = "VESSEL" Then
                '    Dim strQuery As String = MedscreenCommonGUIConfig.NodeQuery("VesselName")
                '    strQuery = String.Format(strQuery, myScheme.ApplyTo.Trim)

                '    ocmd = New OleDb.OleDbCommand(strQuery, CConnection.DbConnection)
                '    Dim strVName As String = myScheme.ApplyTo
                '    Try
                '        If CConnection.ConnOpen Then
                '            strVName = ocmd.ExecuteScalar
                '        End If
                '    Catch
                '    Finally
                '        CConnection.SetConnClosed()
                '    End Try
                '    strSamples = "Charge for samples collected from " & strVName & " is "
                '    strRet += strSamples & (Price * (1 - myScheme.SampleDiscount / 100)).ToString("c", myClient.Culture) & " per sample"
                'Else
                '    If (myScheme.MaxSamples = myScheme.MinSamples) Or (myScheme.MaxSamples = 0) Then
                '        strSamples = "Charge for " & myScheme.MinSamples & " samples is "
                '    Else
                '        strSamples = "Charge for between " & myScheme.MinSamples & " and " & myScheme.MaxSamples & " samples is "
                '    End If
                '    If myScheme.MaxSamples > 0 Then
                '        strRet += strSamples & _
                '              actualPriceMin.ToString("c", myClient.Culture)
                '        If myScheme.MaxSamples <> myScheme.MinSamples Then
                '            strRet += " - " & actualPriceMax.ToString("c", myClient.Culture)
                '        End If
                '    Else
                '        strRet += strSamples & _
                '             actualPriceMin.ToString("c", myClient.Culture) & " -  and then " & _
                '             (Price * (1 - myScheme.SampleDiscount / 100)).ToString("c", myClient.Culture) & " per sample"
                '    End If
                'End If
                Dim strRet As String = ""
                Dim oColl As New Collection
                oColl.Add(myClient.Identity)
                oColl.Add(myScheme.DiscountSchemeId)
                Dim strXML As String = CConnection.PackageStringList("Lib_customer.xmlexample", oColl)
                If strXML IsNot Nothing AndAlso strXML.Trim.Length > 0 Then             'We have some xml
                    strRet = Medscreen.ResolveStyleSheet(strXML, "VolDiscountExampleTxt.xsl", 0)
                End If
                objEntry = New ListViewItem.ListViewSubItem(Me, strRet)
                Me.SubItems.Add(objEntry)

            End If
            Return myReturn
        End Function
    End Class

    Public Class dgvVolumeDiscountEntry
        Inherits DataGridViewRow

        Private myScheme As VolumeDiscountEntry
        Private myClient As Intranet.intranet.customerns.Client

        Public Sub New(ByVal VDEntry As VolumeDiscountEntry, ByVal Client As Intranet.intranet.customerns.Client)
            myScheme = VDEntry
            myClient = Client
            Me.Height = Me.Height * 3
            'Now redraw item 
            DefaultCellStyle.BackColor = Color.AntiqueWhite
            DefaultCellStyle.ForeColor = Color.Black
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Dim myReturn As Boolean = False
            Me.Cells.Clear()
            If Not Me.myScheme Is Nothing Then
                Dim entryId As New DataGridViewTextBoxCell
                entryId.Value = Me.myScheme.DiscountSchemeId & "-" & Me.myScheme.EntryNumber
                Me.Cells.Add(entryId)

                Dim SchemeType As New DataGridViewTextBoxCell
                SchemeType.Value = Me.myScheme.Type
                Me.Cells.Add(SchemeType)


                'Apply to 
                Dim ApplyTo As New DataGridViewTextBoxCell
                ApplyTo.Value = Me.myScheme.ApplyTo
                Me.Cells.Add(ApplyTo)

                'Min Samples
                Dim MinSamples As New DataGridViewTextBoxCell
                MinSamples.Value = Me.myScheme.MinSamples
                Me.Cells.Add(MinSamples)

                'Max Samples
                Dim MaxSamples As New DataGridViewTextBoxCell
                MaxSamples.Value = Me.myScheme.MaxSamples
                Me.Cells.Add(MaxSamples)

                'Min Charge
                Dim MinChargeDiscount As New DataGridViewTextBoxCell
                MinChargeDiscount.Value = myScheme.MinChargeDiscount.ToString("0.00")
                Me.Cells.Add(MinChargeDiscount)

                'Sample Discount
                Dim SampleDiscount As New DataGridViewTextBoxCell
                SampleDiscount.Value = myScheme.SampleDiscount.ToString("0.00")
                Me.Cells.Add(SampleDiscount)
                'Do example 
                If myClient Is Nothing Then Exit Function 'Check we have a customer 

                Dim strRet As String = ""
                Dim oColl As New Collection
                oColl.Add(myClient.Identity)
                oColl.Add(myScheme.DiscountSchemeId)
                Dim strXML As String = CConnection.PackageStringList("Lib_customer.xmlexample", oColl)
                If strXML IsNot Nothing AndAlso strXML.Trim.Length > 0 Then             'We have some xml
                    strRet = Medscreen.ResolveStyleSheet(strXML, "VolDiscountExampleTxt.xsl", 0)
                End If
                Dim Example As New DataGridViewTextBoxCell
                Example.Value = strRet
                Example.Style.WrapMode = DataGridViewTriState.True
                Example.ToolTipText = strRet
                Me.Cells.Add(Example)

            End If
            Return myReturn
        End Function

        Property Scheme() As VolumeDiscountEntry
            Get
                Return myScheme
            End Get
            Set(ByVal value As VolumeDiscountEntry)
                myScheme = value
            End Set
        End Property
    End Class



    Public Class lvVolumeDiscountHeader
        Inherits lvVolDsItem

        Private myScheme As Intranet.intranet.customerns.CustomerVolumeDiscount

        Public Sub New(ByVal VDEntry As Intranet.intranet.customerns.CustomerVolumeDiscount)
            myScheme = VDEntry
            'Now redraw item 
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Dim myReturn As Boolean = False
            Me.SubItems.Clear()
            Me.Text = ""
            Me.BackColor = Color.Wheat
            If Not Me.myScheme Is Nothing Then
                Me.Text = Me.myScheme.DiscountSchemeId
                Dim aSub As ListViewItem.ListViewSubItem
                aSub = New ListViewItem.ListViewSubItem(Me, Me.myScheme.DiscountScheme.Description)
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, "")
                Me.SubItems.Add(aSub)
                aSub = New ListViewItem.ListViewSubItem(Me, Me.myScheme.ContractAmendment)
                Me.SubItems.Add(aSub)

                '<Added code modified 26-Apr-2007 06:49 by LOGOS\Taylor> 
                If Me.myScheme.EndDate <= Today AndAlso Me.myScheme.EndDate > DateField.ZeroDate Then
                    Me.Text += " - Expired on - " & Me.myScheme.Startdate.ToString("dd-MMM-yyyy")
                    Me.BackColor = Color.Red
                    Me.ForeColor = Color.White
                End If

                If Me.myScheme.Startdate > Today Then
                    Me.Text += " starts - " & Me.myScheme.Startdate.ToString("dd-MMM-yyyy")
                    Me.BackColor = Color.Yellow

                End If
                '</Added code modified 26-Apr-2007 06:49 by LOGOS\Taylor> 

            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose volume discount scheme
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/04/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Scheme() As Intranet.intranet.customerns.CustomerVolumeDiscount
            Get
                Return Me.myScheme
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.CustomerVolumeDiscount)
                myScheme = Value
            End Set
        End Property
    End Class


    Public Class dgvVolumeDiscountHeader
        Inherits DataGridViewRow

        Private myScheme As Intranet.intranet.customerns.CustomerVolumeDiscount

        Public Sub New(ByVal VDEntry As Intranet.intranet.customerns.CustomerVolumeDiscount)
            myScheme = VDEntry
            'Now redraw item 
            DefaultCellStyle.BackColor = Color.Wheat
            DefaultCellStyle.ForeColor = Color.Black
            'DefaultCellStyle.Font.Bold = True
            Refresh()
        End Sub

        Public Function Refresh() As Boolean
            Dim myReturn As Boolean = False
            Me.Cells.Clear()
            If Not Me.myScheme Is Nothing Then
                Dim SchemeId As New DataGridViewTextBoxCell
                SchemeId.Value = CStr(Me.myScheme.DiscountSchemeId)
                Me.Cells.Add(SchemeId)

                Dim Description As New DataGridViewTextBoxCell
                If myScheme.DiscountScheme Is Nothing Then
                    myScheme.DiscountScheme = New MedscreenLib.Glossary.VolumeDiscountHeader()
                    myScheme.DiscountScheme.DiscountSchemeId = myScheme.DiscountSchemeId

                End If

                Description.Value = Me.myScheme.DiscountScheme.Description
                Me.Cells.Add(Description)


                Dim asub As New DataGridViewTextBoxCell


                Me.Cells.Add(asub)
                ''asub = New DataGridViewTextBoxCell()
                'Me.Cells.Add(asub)
                'asub = New DataGridViewTextBoxCell()
                'Me.Cells.Add(asub)
                'asub = New DataGridViewTextBoxCell()
                'Me.Cells.Add(asub)
                'asub = New DataGridViewTextBoxCell()
                'Me.Cells.Add(asub)
                'asub = New DataGridViewTextBoxCell()
                'Me.Cells.Add(asub)
                'asub = New DataGridViewTextBoxCell()
                asub.Value = Me.myScheme.ContractAmendment
                'Me.Cells.Add(asub)
                Dim sdate As New DataGridViewTextBoxCell
                If myScheme.Startdate > DateField.ZeroDate Then
                    sdate.Value = myScheme.Startdate
                    sdate.Style.Format = "dd-MMM-yyyy"
                Else
                    sdate.Value = "Not Set"
                End If
                Me.Cells.Add(sdate)

                Dim edate As New DataGridViewTextBoxCell
                If myScheme.EndDate > DateField.ZeroDate Then
                    edate.Value = myScheme.EndDate
                    edate.Style.Format = "dd-MMM-yyyy"
                Else
                    edate.Value = "Not Set"
                End If
                Me.Cells.Add(edate)

                '<Added code modified 26-Apr-2007 06:49 by LOGOS\Taylor> 
                If Me.myScheme.EndDate <= Today AndAlso Me.myScheme.EndDate > DateField.ZeroDate Then
                    SchemeId.Value += " - Expired on - " & Me.myScheme.Startdate.ToString("dd-MMM-yyyy")
                    SchemeId.Style.BackColor = Color.Red
                    SchemeId.Style.ForeColor = Color.White
                End If

                If Me.myScheme.Startdate > Today Then
                    SchemeId.Value += " starts - " & Me.myScheme.Startdate.ToString("dd-MMM-yyyy")
                    SchemeId.Style.BackColor = Color.Yellow

                End If
                '</Added code modified 26-Apr-2007 06:49 by LOGOS\Taylor> 

            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose volume discount scheme
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/04/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Scheme() As Intranet.intranet.customerns.CustomerVolumeDiscount
            Get
                Return Me.myScheme
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.CustomerVolumeDiscount)
                myScheme = Value
            End Set
        End Property
    End Class
End Namespace
