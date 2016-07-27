'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 08:32:43+00 $
'$Log: frmCollPay.vb,v $
'Revision 1.0  2005-11-29 08:32:43+00  taylor
'checked in after commenting
'
'Revision 1.2  2005-06-16 17:39:17+01  taylor
'Milestone Check in
'
'Revision 1.1  2004-05-04 10:06:25+01  taylor
'<>
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports Intranet
Imports Intranet.intranet.Support
Imports System.Reflection
Imports Intranet.intranet
Imports Intranet.collectorpay
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCollPay
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Collecting Officers Pay form 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCollPay
    Inherits System.Windows.Forms.Form

#Region "Constants"
    Const cstPAYCALLOUTINRS001 As String = "COLLCOSTCI"      'Call Out In Hours Pay                           
    Const cstPAYCALLOUTINRS002 As String = "TRAVCOSTCI"      'Call Out In Hours Pay                           
    Const cstPAYCALLOUTOUTS001 As String = "COLLCOSTCO"      'Call Out In Hours Pay       
    Const cstPAYCALLOUTOUTS002 As String = "TRAVCOSTCO"      'Call Out In Hours Pay                           

    Const cstPAYDAYRATE As String = "PAYDAYRATE"                    'Collection Attendance Day rate                  
    Const cstPAYHALFDAY As String = "PAYHALFDAY"                    'Collection Attendance Half Day rate
    Const cstPAYROUNTINHRS001 As String = "COLLCOST"        'Routine In Hours Pay. 
    Const cstPAYROUNTINHRS002 As String = "TRAVCOST"        'Routine In Hours Pay. 
    Const cstPAYROUNTOUTHRS001 As String = "COLLCOSTO"        'Routine In Hours Pay. 
    Const cstPAYROUNTOUTHRS002 As String = "TRAVCOSTO"        'Routine In Hours Pay. 
    Const cstPAYSAMPLE001 As String = "PAYSAMPLE001"                'Routine per DOA Sample Rate C/O pay   
    Const cstPAYSAMPLE002 As String = "PAYSAMPLE002"                'Benzene per sample 
    Const cstPAYSAMPLE003 As String = "PAYSAMPLE003"                'Routine DOA and Benzene per sample  
    Const cstPAYSAMPLE004 As String = "PAYSAMPLE004"                'Breath Test per sample  
    Const cstPAYSAMPLE005 As String = "PAYSAMPLE005"                'Breath Test (Witness) per sample  
    Const cstPAYSAMPLE006 As String = "PAYSAMPLE006"                'Breath test with Urine sample  
    Const cstPAYWRKTIME As String = "PAYWRKTIME"                    'Collecting Officer Working Time pr hr  
    Const cstPAYHOURMIN As String = "PAYHOURMIN"                    'Collecting Officer Minimum Hourly Amount
    Const cstSAMPLE As String = "SAMPLE"                            'PerSample Rate
    Const cstMILEAGE As String = "PAYPERMILE"                          'Pay per mile
    Const cstStartPayMonth As Date = #7/1/2011#
    Friend WithEvents ckCallOutRate As System.Windows.Forms.CheckBox
    Friend WithEvents lblOtherOfficers As System.Windows.Forms.Label
    Friend WithEvents lblPayType As System.Windows.Forms.Label
    Friend WithEvents udKMMiles As System.Windows.Forms.DomainUpDown
    Friend WithEvents udMileage As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblMileage As System.Windows.Forms.Label
    Friend WithEvents lblDonorPay As System.Windows.Forms.Label
    Friend WithEvents udDonors As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblNDonors As System.Windows.Forms.Label
    Friend WithEvents udHalfDays As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblHalfDays As System.Windows.Forms.Label
    Friend WithEvents lblmileage3 As System.Windows.Forms.Label
    Friend WithEvents lblAverageMileage As System.Windows.Forms.Label
    Friend WithEvents lbllmileage4 As System.Windows.Forms.Label
    Friend WithEvents lblOutOfHoursTravelRate As System.Windows.Forms.Label
    Friend WithEvents lblInHoursTravelPayRate As System.Windows.Forms.Label
    Friend WithEvents lblOutHoursTravel As System.Windows.Forms.Label
    Friend WithEvents lblinHoursTravel As System.Windows.Forms.Label
    Friend WithEvents dtoutTravelTime As System.Windows.Forms.TextBox
    Friend WithEvents dtInTravelTime As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents dtTravelTime As System.Windows.Forms.TextBox
    'Start of new system
#End Region

#Region "Declarations"
    Private SystemChangeDate As Date
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim strFilename As String
        Dim MyAssembly As [Assembly]
        MyAssembly = MyAssembly.GetAssembly(GetType(MedscreenCommonGui.Controls.ListColumn))
        Dim directory As String = IO.Path.GetDirectoryName(MyAssembly.CodeBase)
        strFilename = Mid(directory, 7) & "\CollPayRates.xml"
        Me.CollPayPricing1.ReadXml(strFilename)
        Me.cbInvStatus.PhraseType = "COLLINVSTA"

        'get the date on which the paysystem changes

        SystemChangeDate = MedscreenLib.Glossary.ConfigItem.GetConfigValue("MEDPAYCHANGE")
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
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lkPort As System.Windows.Forms.LinkLabel
    Friend WithEvents lblJob As System.Windows.Forms.Label
    Friend WithEvents lblCollection As System.Windows.Forms.Label
    Friend WithEvents lkOfficer As System.Windows.Forms.LinkLabel
    Friend WithEvents pnlTimes As System.Windows.Forms.Panel
    Friend WithEvents dtLeave As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblSiteLeave As System.Windows.Forms.Label
    Friend WithEvents dtArrive As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtTravelEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTravelEnd As System.Windows.Forms.Label
    Friend WithEvents dtTravelStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents gbCosts As System.Windows.Forms.GroupBox
    Friend WithEvents lblAgreedCosts As System.Windows.Forms.Label
    Friend WithEvents gbExpenses As System.Windows.Forms.GroupBox
    Friend WithEvents udOther As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblOther As System.Windows.Forms.Label
    Friend WithEvents udAgreed As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents udTaxi As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblTaxi As System.Windows.Forms.Label
    Friend WithEvents udHotel As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblHotel As System.Windows.Forms.Label
    Friend WithEvents UDRail As System.Windows.Forms.NumericUpDown
    Friend WithEvents udAir As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblAir As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gbCalculation As System.Windows.Forms.GroupBox
    Friend WithEvents dtTimeOnJob As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtTimeOnSite As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DTExpectedTimeOnSite As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtOutTime As System.Windows.Forms.TextBox
    Friend WithEvents dtHometime As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtavgSpeedHome As System.Windows.Forms.TextBox
    Friend WithEvents txtAvgSpeedOut As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents txtCollCost As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtExpTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DTMonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents CollPayPricing1 As MedscreenCommonGui.CollPayPricing
    Friend WithEvents cmdCalc As System.Windows.Forms.Button
    Friend WithEvents txtOutofHours As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblinhrspay As System.Windows.Forms.Label
    Friend WithEvents lblouthourspay As System.Windows.Forms.Label
    Friend WithEvents lblCurrency1 As System.Windows.Forms.Label
    Friend WithEvents lblCurrency2 As System.Windows.Forms.Label
    Friend WithEvents lblCurrency3 As System.Windows.Forms.Label
    Friend WithEvents ckPerSample As System.Windows.Forms.CheckBox
    Friend WithEvents ckHalfDay As System.Windows.Forms.CheckBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lblMilagePay As System.Windows.Forms.Label
    Friend WithEvents lblTotalPay As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents lblCollType As System.Windows.Forms.Label
    Friend WithEvents lblMileage2 As System.Windows.Forms.Label
    Friend WithEvents cbInvStatus As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents lblinhoursPayRate As System.Windows.Forms.Label
    Friend WithEvents lblOutOfHoursPayRate As System.Windows.Forms.Label
    Friend WithEvents lblCollCost As System.Windows.Forms.Label
    Friend WithEvents lblCollectionDate As System.Windows.Forms.Label
    Friend WithEvents lblHoliday As System.Windows.Forms.Label
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents udNonItem As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCollPay))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCalc = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lblPayType = New System.Windows.Forms.Label
        Me.lblHoliday = New System.Windows.Forms.Label
        Me.lblCollectionDate = New System.Windows.Forms.Label
        Me.lblCollType = New System.Windows.Forms.Label
        Me.lkPort = New System.Windows.Forms.LinkLabel
        Me.lblJob = New System.Windows.Forms.Label
        Me.lblCollection = New System.Windows.Forms.Label
        Me.lkOfficer = New System.Windows.Forms.LinkLabel
        Me.pnlTimes = New System.Windows.Forms.Panel
        Me.lblOtherOfficers = New System.Windows.Forms.Label
        Me.ckCallOutRate = New System.Windows.Forms.CheckBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtavgSpeedHome = New System.Windows.Forms.TextBox
        Me.txtAvgSpeedOut = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtHometime = New System.Windows.Forms.TextBox
        Me.dtOutTime = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.dtLeave = New System.Windows.Forms.DateTimePicker
        Me.lblSiteLeave = New System.Windows.Forms.Label
        Me.dtArrive = New System.Windows.Forms.DateTimePicker
        Me.dtTravelEnd = New System.Windows.Forms.DateTimePicker
        Me.lblTravelEnd = New System.Windows.Forms.Label
        Me.dtTravelStart = New System.Windows.Forms.DateTimePicker
        Me.gbCosts = New System.Windows.Forms.GroupBox
        Me.lblDonorPay = New System.Windows.Forms.Label
        Me.udDonors = New System.Windows.Forms.NumericUpDown
        Me.lblNDonors = New System.Windows.Forms.Label
        Me.lblMilagePay = New System.Windows.Forms.Label
        Me.lblAgreedCosts = New System.Windows.Forms.Label
        Me.gbExpenses = New System.Windows.Forms.GroupBox
        Me.udNonItem = New System.Windows.Forms.NumericUpDown
        Me.Label19 = New System.Windows.Forms.Label
        Me.udOther = New System.Windows.Forms.NumericUpDown
        Me.lblOther = New System.Windows.Forms.Label
        Me.udAgreed = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.udTaxi = New System.Windows.Forms.NumericUpDown
        Me.lblTaxi = New System.Windows.Forms.Label
        Me.udHotel = New System.Windows.Forms.NumericUpDown
        Me.lblHotel = New System.Windows.Forms.Label
        Me.UDRail = New System.Windows.Forms.NumericUpDown
        Me.udAir = New System.Windows.Forms.NumericUpDown
        Me.lblAir = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.gbCalculation = New System.Windows.Forms.GroupBox
        Me.lblAverageMileage = New System.Windows.Forms.Label
        Me.udKMMiles = New System.Windows.Forms.DomainUpDown
        Me.udMileage = New System.Windows.Forms.NumericUpDown
        Me.lblMileage = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtComment = New System.Windows.Forms.TextBox
        Me.lblCollCost = New System.Windows.Forms.Label
        Me.lblOutOfHoursPayRate = New System.Windows.Forms.Label
        Me.lblinhoursPayRate = New System.Windows.Forms.Label
        Me.lblMileage2 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.lblTotalPay = New System.Windows.Forms.Label
        Me.udHalfDays = New System.Windows.Forms.NumericUpDown
        Me.lblHalfDays = New System.Windows.Forms.Label
        Me.ckHalfDay = New System.Windows.Forms.CheckBox
        Me.ckPerSample = New System.Windows.Forms.CheckBox
        Me.lblCurrency3 = New System.Windows.Forms.Label
        Me.lblCurrency2 = New System.Windows.Forms.Label
        Me.lblCurrency1 = New System.Windows.Forms.Label
        Me.lblouthourspay = New System.Windows.Forms.Label
        Me.lblinhrspay = New System.Windows.Forms.Label
        Me.txtOutofHours = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.DTMonth = New System.Windows.Forms.DateTimePicker
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtExpTotal = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtCollCost = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtTotalCost = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.DTExpectedTimeOnSite = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtTimeOnJob = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtTimeOnSite = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblmileage3 = New System.Windows.Forms.Label
        Me.lbllmileage4 = New System.Windows.Forms.Label
        Me.CollPayPricing1 = New MedscreenCommonGui.CollPayPricing
        Me.lblOutOfHoursTravelRate = New System.Windows.Forms.Label
        Me.lblInHoursTravelPayRate = New System.Windows.Forms.Label
        Me.lblOutHoursTravel = New System.Windows.Forms.Label
        Me.lblinHoursTravel = New System.Windows.Forms.Label
        Me.dtoutTravelTime = New System.Windows.Forms.TextBox
        Me.dtInTravelTime = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.dtTravelTime = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.cbInvStatus = New MedscreenCommonGui.PhraseComboBox
        Me.Panel1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.pnlTimes.SuspendLayout()
        Me.gbCosts.SuspendLayout()
        CType(Me.udDonors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbExpenses.SuspendLayout()
        CType(Me.udNonItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udOther, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udAgreed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udTaxi, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udHotel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UDRail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udAir, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbCalculation.SuspendLayout()
        CType(Me.udMileage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udHalfDays, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CollPayPricing1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCalc)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 825)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(720, 32)
        Me.Panel1.TabIndex = 4
        '
        'cmdCalc
        '
        Me.cmdCalc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdCalc.Image = CType(resources.GetObject("cmdCalc.Image"), System.Drawing.Image)
        Me.cmdCalc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCalc.Location = New System.Drawing.Point(24, 4)
        Me.cmdCalc.Name = "cmdCalc"
        Me.cmdCalc.Size = New System.Drawing.Size(72, 24)
        Me.cmdCalc.TabIndex = 6
        Me.cmdCalc.Text = "Calculate"
        Me.cmdCalc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(636, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(556, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lblPayType)
        Me.Panel2.Controls.Add(Me.lblHoliday)
        Me.Panel2.Controls.Add(Me.lblCollectionDate)
        Me.Panel2.Controls.Add(Me.cbInvStatus)
        Me.Panel2.Controls.Add(Me.lblCollType)
        Me.Panel2.Controls.Add(Me.lkPort)
        Me.Panel2.Controls.Add(Me.lblJob)
        Me.Panel2.Controls.Add(Me.lblCollection)
        Me.Panel2.Controls.Add(Me.lkOfficer)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(720, 106)
        Me.Panel2.TabIndex = 5
        '
        'lblPayType
        '
        Me.lblPayType.Location = New System.Drawing.Point(21, 83)
        Me.lblPayType.Name = "lblPayType"
        Me.lblPayType.Size = New System.Drawing.Size(376, 23)
        Me.lblPayType.TabIndex = 24
        '
        'lblHoliday
        '
        Me.lblHoliday.Location = New System.Drawing.Point(416, 56)
        Me.lblHoliday.Name = "lblHoliday"
        Me.lblHoliday.Size = New System.Drawing.Size(296, 23)
        Me.lblHoliday.TabIndex = 23
        '
        'lblCollectionDate
        '
        Me.lblCollectionDate.Location = New System.Drawing.Point(24, 56)
        Me.lblCollectionDate.Name = "lblCollectionDate"
        Me.lblCollectionDate.Size = New System.Drawing.Size(376, 23)
        Me.lblCollectionDate.TabIndex = 22
        '
        'lblCollType
        '
        Me.lblCollType.Location = New System.Drawing.Point(424, 0)
        Me.lblCollType.Name = "lblCollType"
        Me.lblCollType.Size = New System.Drawing.Size(176, 20)
        Me.lblCollType.TabIndex = 20
        '
        'lkPort
        '
        Me.lkPort.Location = New System.Drawing.Point(20, 31)
        Me.lkPort.Name = "lkPort"
        Me.lkPort.Size = New System.Drawing.Size(200, 16)
        Me.lkPort.TabIndex = 19
        Me.lkPort.TabStop = True
        Me.lkPort.Text = "Port"
        '
        'lblJob
        '
        Me.lblJob.Location = New System.Drawing.Point(244, 31)
        Me.lblJob.Name = "lblJob"
        Me.lblJob.Size = New System.Drawing.Size(224, 23)
        Me.lblJob.TabIndex = 18
        Me.lblJob.Text = "Job:"
        '
        'lblCollection
        '
        Me.lblCollection.Location = New System.Drawing.Point(236, 3)
        Me.lblCollection.Name = "lblCollection"
        Me.lblCollection.Size = New System.Drawing.Size(176, 20)
        Me.lblCollection.TabIndex = 17
        Me.lblCollection.Text = "Collection"
        '
        'lkOfficer
        '
        Me.lkOfficer.Location = New System.Drawing.Point(20, 3)
        Me.lkOfficer.Name = "lkOfficer"
        Me.lkOfficer.Size = New System.Drawing.Size(200, 23)
        Me.lkOfficer.TabIndex = 16
        Me.lkOfficer.TabStop = True
        Me.lkOfficer.Text = "Officer: "
        '
        'pnlTimes
        '
        Me.pnlTimes.Controls.Add(Me.lblOtherOfficers)
        Me.pnlTimes.Controls.Add(Me.ckCallOutRate)
        Me.pnlTimes.Controls.Add(Me.Label15)
        Me.pnlTimes.Controls.Add(Me.Label16)
        Me.pnlTimes.Controls.Add(Me.txtavgSpeedHome)
        Me.pnlTimes.Controls.Add(Me.txtAvgSpeedOut)
        Me.pnlTimes.Controls.Add(Me.Label9)
        Me.pnlTimes.Controls.Add(Me.dtHometime)
        Me.pnlTimes.Controls.Add(Me.dtOutTime)
        Me.pnlTimes.Controls.Add(Me.Label8)
        Me.pnlTimes.Controls.Add(Me.Label7)
        Me.pnlTimes.Controls.Add(Me.Label6)
        Me.pnlTimes.Controls.Add(Me.dtLeave)
        Me.pnlTimes.Controls.Add(Me.lblSiteLeave)
        Me.pnlTimes.Controls.Add(Me.dtArrive)
        Me.pnlTimes.Controls.Add(Me.dtTravelEnd)
        Me.pnlTimes.Controls.Add(Me.lblTravelEnd)
        Me.pnlTimes.Controls.Add(Me.dtTravelStart)
        Me.pnlTimes.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTimes.Location = New System.Drawing.Point(0, 106)
        Me.pnlTimes.Name = "pnlTimes"
        Me.pnlTimes.Size = New System.Drawing.Size(720, 144)
        Me.pnlTimes.TabIndex = 0
        '
        'lblOtherOfficers
        '
        Me.lblOtherOfficers.AutoSize = True
        Me.lblOtherOfficers.Location = New System.Drawing.Point(442, 11)
        Me.lblOtherOfficers.Name = "lblOtherOfficers"
        Me.lblOtherOfficers.Size = New System.Drawing.Size(72, 13)
        Me.lblOtherOfficers.TabIndex = 30
        Me.lblOtherOfficers.Text = "Other Officers"
        '
        'ckCallOutRate
        '
        Me.ckCallOutRate.AutoSize = True
        Me.ckCallOutRate.Location = New System.Drawing.Point(476, 80)
        Me.ckCallOutRate.Name = "ckCallOutRate"
        Me.ckCallOutRate.Size = New System.Drawing.Size(129, 17)
        Me.ckCallOutRate.TabIndex = 29
        Me.ckCallOutRate.TabStop = False
        Me.ckCallOutRate.Text = "CallOut Rate (Agreed)"
        Me.ckCallOutRate.UseVisualStyleBackColor = True
        Me.ckCallOutRate.Visible = False
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(368, 104)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(64, 16)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "Avg Speed"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(272, 104)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 16)
        Me.Label16.TabIndex = 27
        Me.Label16.Text = "Travel Times"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtavgSpeedHome
        '
        Me.txtavgSpeedHome.Enabled = False
        Me.txtavgSpeedHome.Location = New System.Drawing.Point(344, 120)
        Me.txtavgSpeedHome.Name = "txtavgSpeedHome"
        Me.txtavgSpeedHome.Size = New System.Drawing.Size(69, 20)
        Me.txtavgSpeedHome.TabIndex = 26
        '
        'txtAvgSpeedOut
        '
        Me.txtAvgSpeedOut.Enabled = False
        Me.txtAvgSpeedOut.Location = New System.Drawing.Point(152, 120)
        Me.txtAvgSpeedOut.Name = "txtAvgSpeedOut"
        Me.txtAvgSpeedOut.Size = New System.Drawing.Size(69, 20)
        Me.txtAvgSpeedOut.TabIndex = 25
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(160, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(64, 16)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Avg Speed"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dtHometime
        '
        Me.dtHometime.Enabled = False
        Me.dtHometime.Location = New System.Drawing.Point(272, 120)
        Me.dtHometime.Name = "dtHometime"
        Me.dtHometime.Size = New System.Drawing.Size(69, 20)
        Me.dtHometime.TabIndex = 23
        '
        'dtOutTime
        '
        Me.dtOutTime.Enabled = False
        Me.dtOutTime.Location = New System.Drawing.Point(72, 120)
        Me.dtOutTime.Name = "dtOutTime"
        Me.dtOutTime.Size = New System.Drawing.Size(69, 20)
        Me.dtOutTime.TabIndex = 22
        Me.dtOutTime.TabStop = False
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(64, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 16)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Travel Times"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(272, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 16)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Time Arrived Home"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(80, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Time Set Off"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dtLeave
        '
        Me.dtLeave.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.dtLeave.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtLeave.Location = New System.Drawing.Point(272, 80)
        Me.dtLeave.Name = "dtLeave"
        Me.dtLeave.ShowUpDown = True
        Me.dtLeave.Size = New System.Drawing.Size(120, 20)
        Me.dtLeave.TabIndex = 2
        Me.dtLeave.Value = New Date(2003, 11, 13, 12, 43, 0, 592)
        '
        'lblSiteLeave
        '
        Me.lblSiteLeave.Location = New System.Drawing.Point(256, 56)
        Me.lblSiteLeave.Name = "lblSiteLeave"
        Me.lblSiteLeave.Size = New System.Drawing.Size(88, 23)
        Me.lblSiteLeave.TabIndex = 17
        Me.lblSiteLeave.Text = "Time Left Site"
        Me.lblSiteLeave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtArrive
        '
        Me.dtArrive.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.dtArrive.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtArrive.Location = New System.Drawing.Point(72, 80)
        Me.dtArrive.Name = "dtArrive"
        Me.dtArrive.ShowUpDown = True
        Me.dtArrive.Size = New System.Drawing.Size(120, 20)
        Me.dtArrive.TabIndex = 1
        Me.dtArrive.Value = New Date(2003, 11, 13, 12, 43, 0, 642)
        '
        'dtTravelEnd
        '
        Me.dtTravelEnd.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.dtTravelEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTravelEnd.Location = New System.Drawing.Point(272, 32)
        Me.dtTravelEnd.Name = "dtTravelEnd"
        Me.dtTravelEnd.ShowUpDown = True
        Me.dtTravelEnd.Size = New System.Drawing.Size(120, 20)
        Me.dtTravelEnd.TabIndex = 3
        Me.dtTravelEnd.Value = New Date(2003, 11, 13, 12, 43, 0, 682)
        '
        'lblTravelEnd
        '
        Me.lblTravelEnd.Location = New System.Drawing.Point(48, 54)
        Me.lblTravelEnd.Name = "lblTravelEnd"
        Me.lblTravelEnd.Size = New System.Drawing.Size(140, 23)
        Me.lblTravelEnd.TabIndex = 13
        Me.lblTravelEnd.Text = "Time Arrived on Site"
        Me.lblTravelEnd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtTravelStart
        '
        Me.dtTravelStart.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.dtTravelStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtTravelStart.Location = New System.Drawing.Point(72, 32)
        Me.dtTravelStart.Name = "dtTravelStart"
        Me.dtTravelStart.ShowUpDown = True
        Me.dtTravelStart.Size = New System.Drawing.Size(120, 20)
        Me.dtTravelStart.TabIndex = 0
        Me.dtTravelStart.Value = New Date(2003, 11, 13, 12, 43, 0, 732)
        '
        'gbCosts
        '
        Me.gbCosts.Controls.Add(Me.lblDonorPay)
        Me.gbCosts.Controls.Add(Me.udDonors)
        Me.gbCosts.Controls.Add(Me.lblNDonors)
        Me.gbCosts.Controls.Add(Me.lblMilagePay)
        Me.gbCosts.Controls.Add(Me.lblAgreedCosts)
        Me.gbCosts.Dock = System.Windows.Forms.DockStyle.Top
        Me.gbCosts.Location = New System.Drawing.Point(0, 250)
        Me.gbCosts.Name = "gbCosts"
        Me.gbCosts.Size = New System.Drawing.Size(720, 56)
        Me.gbCosts.TabIndex = 2
        Me.gbCosts.TabStop = False
        Me.gbCosts.Text = "Costs"
        '
        'lblDonorPay
        '
        Me.lblDonorPay.Location = New System.Drawing.Point(215, 16)
        Me.lblDonorPay.Name = "lblDonorPay"
        Me.lblDonorPay.Size = New System.Drawing.Size(100, 23)
        Me.lblDonorPay.TabIndex = 27
        Me.lblDonorPay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'udDonors
        '
        Me.udDonors.Location = New System.Drawing.Point(111, 16)
        Me.udDonors.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udDonors.Name = "udDonors"
        Me.udDonors.Size = New System.Drawing.Size(96, 20)
        Me.udDonors.TabIndex = 25
        '
        'lblNDonors
        '
        Me.lblNDonors.Location = New System.Drawing.Point(39, 16)
        Me.lblNDonors.Name = "lblNDonors"
        Me.lblNDonors.Size = New System.Drawing.Size(77, 23)
        Me.lblNDonors.TabIndex = 26
        Me.lblNDonors.Text = "No of Donors "
        '
        'lblMilagePay
        '
        Me.lblMilagePay.Location = New System.Drawing.Point(232, 24)
        Me.lblMilagePay.Name = "lblMilagePay"
        Me.lblMilagePay.Size = New System.Drawing.Size(64, 13)
        Me.lblMilagePay.TabIndex = 3
        '
        'lblAgreedCosts
        '
        Me.lblAgreedCosts.Location = New System.Drawing.Point(473, 16)
        Me.lblAgreedCosts.Name = "lblAgreedCosts"
        Me.lblAgreedCosts.Size = New System.Drawing.Size(136, 23)
        Me.lblAgreedCosts.TabIndex = 2
        Me.lblAgreedCosts.Text = "No Costs Agreed"
        Me.lblAgreedCosts.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'gbExpenses
        '
        Me.gbExpenses.Controls.Add(Me.udNonItem)
        Me.gbExpenses.Controls.Add(Me.Label19)
        Me.gbExpenses.Controls.Add(Me.udOther)
        Me.gbExpenses.Controls.Add(Me.lblOther)
        Me.gbExpenses.Controls.Add(Me.udAgreed)
        Me.gbExpenses.Controls.Add(Me.Label3)
        Me.gbExpenses.Controls.Add(Me.udTaxi)
        Me.gbExpenses.Controls.Add(Me.lblTaxi)
        Me.gbExpenses.Controls.Add(Me.udHotel)
        Me.gbExpenses.Controls.Add(Me.lblHotel)
        Me.gbExpenses.Controls.Add(Me.UDRail)
        Me.gbExpenses.Controls.Add(Me.udAir)
        Me.gbExpenses.Controls.Add(Me.lblAir)
        Me.gbExpenses.Controls.Add(Me.Label2)
        Me.gbExpenses.Dock = System.Windows.Forms.DockStyle.Top
        Me.gbExpenses.Location = New System.Drawing.Point(0, 306)
        Me.gbExpenses.Name = "gbExpenses"
        Me.gbExpenses.Size = New System.Drawing.Size(720, 112)
        Me.gbExpenses.TabIndex = 3
        Me.gbExpenses.TabStop = False
        Me.gbExpenses.Text = "Expenses"
        '
        'udNonItem
        '
        Me.udNonItem.DecimalPlaces = 2
        Me.udNonItem.Location = New System.Drawing.Point(545, 77)
        Me.udNonItem.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udNonItem.Name = "udNonItem"
        Me.udNonItem.Size = New System.Drawing.Size(96, 20)
        Me.udNonItem.TabIndex = 14
        Me.udNonItem.Visible = False
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(473, 77)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(56, 23)
        Me.Label19.TabIndex = 13
        Me.Label19.Text = "Non Itemised"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label19.Visible = False
        '
        'udOther
        '
        Me.udOther.DecimalPlaces = 2
        Me.udOther.Location = New System.Drawing.Point(312, 77)
        Me.udOther.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udOther.Name = "udOther"
        Me.udOther.Size = New System.Drawing.Size(96, 20)
        Me.udOther.TabIndex = 12
        '
        'lblOther
        '
        Me.lblOther.Location = New System.Drawing.Point(240, 77)
        Me.lblOther.Name = "lblOther"
        Me.lblOther.Size = New System.Drawing.Size(56, 23)
        Me.lblOther.TabIndex = 11
        Me.lblOther.Text = "Other"
        Me.lblOther.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udAgreed
        '
        Me.udAgreed.DecimalPlaces = 2
        Me.udAgreed.Location = New System.Drawing.Point(88, 77)
        Me.udAgreed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.udAgreed.Name = "udAgreed"
        Me.udAgreed.Size = New System.Drawing.Size(96, 20)
        Me.udAgreed.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Agreed"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udTaxi
        '
        Me.udTaxi.DecimalPlaces = 2
        Me.udTaxi.Location = New System.Drawing.Point(312, 48)
        Me.udTaxi.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udTaxi.Name = "udTaxi"
        Me.udTaxi.Size = New System.Drawing.Size(96, 20)
        Me.udTaxi.TabIndex = 8
        '
        'lblTaxi
        '
        Me.lblTaxi.Location = New System.Drawing.Point(240, 48)
        Me.lblTaxi.Name = "lblTaxi"
        Me.lblTaxi.Size = New System.Drawing.Size(56, 23)
        Me.lblTaxi.TabIndex = 7
        Me.lblTaxi.Text = "Taxi"
        Me.lblTaxi.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udHotel
        '
        Me.udHotel.DecimalPlaces = 2
        Me.udHotel.Location = New System.Drawing.Point(88, 48)
        Me.udHotel.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udHotel.Name = "udHotel"
        Me.udHotel.Size = New System.Drawing.Size(96, 20)
        Me.udHotel.TabIndex = 6
        '
        'lblHotel
        '
        Me.lblHotel.Location = New System.Drawing.Point(16, 48)
        Me.lblHotel.Name = "lblHotel"
        Me.lblHotel.Size = New System.Drawing.Size(56, 23)
        Me.lblHotel.TabIndex = 5
        Me.lblHotel.Text = "Hotel"
        Me.lblHotel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'UDRail
        '
        Me.UDRail.DecimalPlaces = 2
        Me.UDRail.Location = New System.Drawing.Point(88, 18)
        Me.UDRail.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.UDRail.Name = "UDRail"
        Me.UDRail.Size = New System.Drawing.Size(96, 20)
        Me.UDRail.TabIndex = 4
        '
        'udAir
        '
        Me.udAir.DecimalPlaces = 2
        Me.udAir.Location = New System.Drawing.Point(312, 18)
        Me.udAir.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udAir.Name = "udAir"
        Me.udAir.Size = New System.Drawing.Size(96, 20)
        Me.udAir.TabIndex = 3
        '
        'lblAir
        '
        Me.lblAir.Location = New System.Drawing.Point(216, 18)
        Me.lblAir.Name = "lblAir"
        Me.lblAir.Size = New System.Drawing.Size(80, 23)
        Me.lblAir.TabIndex = 2
        Me.lblAir.Text = "Air Travel"
        Me.lblAir.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Rail"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'gbCalculation
        '
        Me.gbCalculation.Controls.Add(Me.Label25)
        Me.gbCalculation.Controls.Add(Me.Label24)
        Me.gbCalculation.Controls.Add(Me.dtTravelTime)
        Me.gbCalculation.Controls.Add(Me.lblOutOfHoursTravelRate)
        Me.gbCalculation.Controls.Add(Me.lblInHoursTravelPayRate)
        Me.gbCalculation.Controls.Add(Me.lblOutHoursTravel)
        Me.gbCalculation.Controls.Add(Me.lblinHoursTravel)
        Me.gbCalculation.Controls.Add(Me.dtoutTravelTime)
        Me.gbCalculation.Controls.Add(Me.dtInTravelTime)
        Me.gbCalculation.Controls.Add(Me.Label27)
        Me.gbCalculation.Controls.Add(Me.lblAverageMileage)
        Me.gbCalculation.Controls.Add(Me.udKMMiles)
        Me.gbCalculation.Controls.Add(Me.udMileage)
        Me.gbCalculation.Controls.Add(Me.lblMileage)
        Me.gbCalculation.Controls.Add(Me.Label18)
        Me.gbCalculation.Controls.Add(Me.txtComment)
        Me.gbCalculation.Controls.Add(Me.lblCollCost)
        Me.gbCalculation.Controls.Add(Me.lblOutOfHoursPayRate)
        Me.gbCalculation.Controls.Add(Me.lblinhoursPayRate)
        Me.gbCalculation.Controls.Add(Me.lblMileage2)
        Me.gbCalculation.Controls.Add(Me.Label17)
        Me.gbCalculation.Controls.Add(Me.lblTotalPay)
        Me.gbCalculation.Controls.Add(Me.udHalfDays)
        Me.gbCalculation.Controls.Add(Me.lblHalfDays)
        Me.gbCalculation.Controls.Add(Me.ckHalfDay)
        Me.gbCalculation.Controls.Add(Me.ckPerSample)
        Me.gbCalculation.Controls.Add(Me.lblCurrency3)
        Me.gbCalculation.Controls.Add(Me.lblCurrency2)
        Me.gbCalculation.Controls.Add(Me.lblCurrency1)
        Me.gbCalculation.Controls.Add(Me.lblouthourspay)
        Me.gbCalculation.Controls.Add(Me.lblinhrspay)
        Me.gbCalculation.Controls.Add(Me.txtOutofHours)
        Me.gbCalculation.Controls.Add(Me.Label14)
        Me.gbCalculation.Controls.Add(Me.DTMonth)
        Me.gbCalculation.Controls.Add(Me.Label13)
        Me.gbCalculation.Controls.Add(Me.txtExpTotal)
        Me.gbCalculation.Controls.Add(Me.Label12)
        Me.gbCalculation.Controls.Add(Me.txtCollCost)
        Me.gbCalculation.Controls.Add(Me.Label11)
        Me.gbCalculation.Controls.Add(Me.txtTotalCost)
        Me.gbCalculation.Controls.Add(Me.Label10)
        Me.gbCalculation.Controls.Add(Me.DTExpectedTimeOnSite)
        Me.gbCalculation.Controls.Add(Me.Label5)
        Me.gbCalculation.Controls.Add(Me.dtTimeOnJob)
        Me.gbCalculation.Controls.Add(Me.Label4)
        Me.gbCalculation.Controls.Add(Me.dtTimeOnSite)
        Me.gbCalculation.Controls.Add(Me.Label1)
        Me.gbCalculation.Controls.Add(Me.lblmileage3)
        Me.gbCalculation.Controls.Add(Me.lbllmileage4)
        Me.gbCalculation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbCalculation.Location = New System.Drawing.Point(0, 418)
        Me.gbCalculation.Name = "gbCalculation"
        Me.gbCalculation.Size = New System.Drawing.Size(720, 407)
        Me.gbCalculation.TabIndex = 1
        Me.gbCalculation.TabStop = False
        Me.gbCalculation.Text = "Calculation"
        '
        'lblAverageMileage
        '
        Me.lblAverageMileage.Location = New System.Drawing.Point(106, 192)
        Me.lblAverageMileage.Name = "lblAverageMileage"
        Me.lblAverageMileage.Size = New System.Drawing.Size(200, 20)
        Me.lblAverageMileage.TabIndex = 40
        Me.lblAverageMileage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'udKMMiles
        '
        Me.udKMMiles.Items.Add("Miles")
        Me.udKMMiles.Items.Add("Km")
        Me.udKMMiles.Location = New System.Drawing.Point(200, 169)
        Me.udKMMiles.Name = "udKMMiles"
        Me.udKMMiles.Size = New System.Drawing.Size(40, 20)
        Me.udKMMiles.TabIndex = 38
        Me.udKMMiles.TabStop = False
        Me.udKMMiles.Text = "Miles"
        '
        'udMileage
        '
        Me.udMileage.Location = New System.Drawing.Point(104, 169)
        Me.udMileage.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.udMileage.Name = "udMileage"
        Me.udMileage.Size = New System.Drawing.Size(96, 20)
        Me.udMileage.TabIndex = 37
        '
        'lblMileage
        '
        Me.lblMileage.Location = New System.Drawing.Point(32, 169)
        Me.lblMileage.Name = "lblMileage"
        Me.lblMileage.Size = New System.Drawing.Size(56, 23)
        Me.lblMileage.TabIndex = 36
        Me.lblMileage.Text = "Mileage"
        Me.lblMileage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(16, 300)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 32)
        Me.Label18.TabIndex = 35
        Me.Label18.Text = "Comment 50 characters"
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(96, 292)
        Me.txtComment.MaxLength = 50
        Me.txtComment.Multiline = True
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(576, 40)
        Me.txtComment.TabIndex = 34
        Me.txtComment.TabStop = False
        '
        'lblCollCost
        '
        Me.lblCollCost.Location = New System.Drawing.Point(520, 205)
        Me.lblCollCost.Name = "lblCollCost"
        Me.lblCollCost.Size = New System.Drawing.Size(184, 23)
        Me.lblCollCost.TabIndex = 33
        Me.lblCollCost.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOutOfHoursPayRate
        '
        Me.lblOutOfHoursPayRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOutOfHoursPayRate.Location = New System.Drawing.Point(536, 48)
        Me.lblOutOfHoursPayRate.Name = "lblOutOfHoursPayRate"
        Me.lblOutOfHoursPayRate.Size = New System.Drawing.Size(178, 23)
        Me.lblOutOfHoursPayRate.TabIndex = 32
        Me.lblOutOfHoursPayRate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblinhoursPayRate
        '
        Me.lblinhoursPayRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblinhoursPayRate.Location = New System.Drawing.Point(536, 24)
        Me.lblinhoursPayRate.Name = "lblinhoursPayRate"
        Me.lblinhoursPayRate.Size = New System.Drawing.Size(178, 23)
        Me.lblinhoursPayRate.TabIndex = 31
        Me.lblinhoursPayRate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMileage2
        '
        Me.lblMileage2.Location = New System.Drawing.Point(536, 140)
        Me.lblMileage2.Name = "lblMileage2"
        Me.lblMileage2.Size = New System.Drawing.Size(100, 23)
        Me.lblMileage2.TabIndex = 30
        Me.lblMileage2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(352, 140)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 23)
        Me.Label17.TabIndex = 29
        Me.Label17.Text = "Total Pay"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalPay
        '
        Me.lblTotalPay.Location = New System.Drawing.Point(424, 140)
        Me.lblTotalPay.Name = "lblTotalPay"
        Me.lblTotalPay.Size = New System.Drawing.Size(100, 23)
        Me.lblTotalPay.TabIndex = 28
        Me.lblTotalPay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'udHalfDays
        '
        Me.udHalfDays.Location = New System.Drawing.Point(392, 172)
        Me.udHalfDays.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.udHalfDays.Name = "udHalfDays"
        Me.udHalfDays.Size = New System.Drawing.Size(96, 20)
        Me.udHalfDays.TabIndex = 27
        Me.udHalfDays.TabStop = False
        '
        'lblHalfDays
        '
        Me.lblHalfDays.Location = New System.Drawing.Point(312, 172)
        Me.lblHalfDays.Name = "lblHalfDays"
        Me.lblHalfDays.Size = New System.Drawing.Size(88, 23)
        Me.lblHalfDays.TabIndex = 26
        Me.lblHalfDays.Text = "No of Half Days "
        '
        'ckHalfDay
        '
        Me.ckHalfDay.Location = New System.Drawing.Point(144, 244)
        Me.ckHalfDay.Name = "ckHalfDay"
        Me.ckHalfDay.Size = New System.Drawing.Size(104, 24)
        Me.ckHalfDay.TabIndex = 25
        Me.ckHalfDay.TabStop = False
        Me.ckHalfDay.Text = "Half Day Pay"
        '
        'ckPerSample
        '
        Me.ckPerSample.Checked = True
        Me.ckPerSample.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckPerSample.Location = New System.Drawing.Point(24, 244)
        Me.ckPerSample.Name = "ckPerSample"
        Me.ckPerSample.Size = New System.Drawing.Size(104, 24)
        Me.ckPerSample.TabIndex = 22
        Me.ckPerSample.Text = "Pay per Sample"
        '
        'lblCurrency3
        '
        Me.lblCurrency3.Location = New System.Drawing.Point(424, 260)
        Me.lblCurrency3.Name = "lblCurrency3"
        Me.lblCurrency3.Size = New System.Drawing.Size(16, 23)
        Me.lblCurrency3.TabIndex = 21
        Me.lblCurrency3.Text = "£"
        '
        'lblCurrency2
        '
        Me.lblCurrency2.Location = New System.Drawing.Point(424, 228)
        Me.lblCurrency2.Name = "lblCurrency2"
        Me.lblCurrency2.Size = New System.Drawing.Size(16, 23)
        Me.lblCurrency2.TabIndex = 20
        Me.lblCurrency2.Text = "£"
        '
        'lblCurrency1
        '
        Me.lblCurrency1.Location = New System.Drawing.Point(424, 204)
        Me.lblCurrency1.Name = "lblCurrency1"
        Me.lblCurrency1.Size = New System.Drawing.Size(16, 23)
        Me.lblCurrency1.TabIndex = 19
        Me.lblCurrency1.Text = "£"
        '
        'lblouthourspay
        '
        Me.lblouthourspay.Location = New System.Drawing.Point(424, 48)
        Me.lblouthourspay.Name = "lblouthourspay"
        Me.lblouthourspay.Size = New System.Drawing.Size(100, 23)
        Me.lblouthourspay.TabIndex = 18
        Me.lblouthourspay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblinhrspay
        '
        Me.lblinhrspay.Location = New System.Drawing.Point(424, 24)
        Me.lblinhrspay.Name = "lblinhrspay"
        Me.lblinhrspay.Size = New System.Drawing.Size(100, 23)
        Me.lblinhrspay.TabIndex = 17
        Me.lblinhrspay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtOutofHours
        '
        Me.txtOutofHours.Enabled = False
        Me.txtOutofHours.Location = New System.Drawing.Point(312, 48)
        Me.txtOutofHours.Name = "txtOutofHours"
        Me.txtOutofHours.Size = New System.Drawing.Size(100, 20)
        Me.txtOutofHours.TabIndex = 16
        Me.txtOutofHours.TabStop = False
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(224, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(96, 23)
        Me.Label14.TabIndex = 15
        Me.Label14.Text = "Out of Hours (hr)"
        '
        'DTMonth
        '
        Me.DTMonth.CustomFormat = "MMMM yyyy"
        Me.DTMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTMonth.Location = New System.Drawing.Point(120, 211)
        Me.DTMonth.Name = "DTMonth"
        Me.DTMonth.Size = New System.Drawing.Size(200, 20)
        Me.DTMonth.TabIndex = 14
        Me.DTMonth.TabStop = False
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(16, 212)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(100, 23)
        Me.Label13.TabIndex = 13
        Me.Label13.Text = "Invoice Month"
        '
        'txtExpTotal
        '
        Me.txtExpTotal.Location = New System.Drawing.Point(440, 234)
        Me.txtExpTotal.Name = "txtExpTotal"
        Me.txtExpTotal.Size = New System.Drawing.Size(72, 20)
        Me.txtExpTotal.TabIndex = 12
        Me.txtExpTotal.TabStop = False
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(344, 234)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 23)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Exp Total"
        '
        'txtCollCost
        '
        Me.txtCollCost.Location = New System.Drawing.Point(440, 205)
        Me.txtCollCost.Name = "txtCollCost"
        Me.txtCollCost.Size = New System.Drawing.Size(72, 20)
        Me.txtCollCost.TabIndex = 10
        Me.txtCollCost.TabStop = False
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(344, 205)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 23)
        Me.Label11.TabIndex = 9
        Me.Label11.Text = "Coll Cost"
        '
        'txtTotalCost
        '
        Me.txtTotalCost.Location = New System.Drawing.Point(440, 260)
        Me.txtTotalCost.Name = "txtTotalCost"
        Me.txtTotalCost.Size = New System.Drawing.Size(72, 20)
        Me.txtTotalCost.TabIndex = 8
        Me.txtTotalCost.TabStop = False
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(344, 260)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(100, 23)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "Total Cost"
        '
        'DTExpectedTimeOnSite
        '
        Me.DTExpectedTimeOnSite.Enabled = False
        Me.DTExpectedTimeOnSite.Location = New System.Drawing.Point(104, 48)
        Me.DTExpectedTimeOnSite.Name = "DTExpectedTimeOnSite"
        Me.DTExpectedTimeOnSite.Size = New System.Drawing.Size(100, 20)
        Me.DTExpectedTimeOnSite.TabIndex = 6
        Me.DTExpectedTimeOnSite.TabStop = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 23)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Expected Time"
        '
        'dtTimeOnJob
        '
        Me.dtTimeOnJob.Enabled = False
        Me.dtTimeOnJob.Location = New System.Drawing.Point(312, 24)
        Me.dtTimeOnJob.Name = "dtTimeOnJob"
        Me.dtTimeOnJob.Size = New System.Drawing.Size(100, 20)
        Me.dtTimeOnJob.TabIndex = 3
        Me.dtTimeOnJob.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(224, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Time on job"
        '
        'dtTimeOnSite
        '
        Me.dtTimeOnSite.Enabled = False
        Me.dtTimeOnSite.Location = New System.Drawing.Point(104, 24)
        Me.dtTimeOnSite.Name = "dtTimeOnSite"
        Me.dtTimeOnSite.Size = New System.Drawing.Size(100, 20)
        Me.dtTimeOnSite.TabIndex = 1
        Me.dtTimeOnSite.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Time on site"
        '
        'lblmileage3
        '
        Me.lblmileage3.Location = New System.Drawing.Point(241, 166)
        Me.lblmileage3.Name = "lblmileage3"
        Me.lblmileage3.Size = New System.Drawing.Size(447, 23)
        Me.lblmileage3.TabIndex = 39
        Me.lblmileage3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbllmileage4
        '
        Me.lbllmileage4.Location = New System.Drawing.Point(309, 185)
        Me.lbllmileage4.Name = "lbllmileage4"
        Me.lbllmileage4.Size = New System.Drawing.Size(405, 23)
        Me.lbllmileage4.TabIndex = 41
        Me.lbllmileage4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CollPayPricing1
        '
        Me.CollPayPricing1.DataSetName = "CollPayPricing"
        Me.CollPayPricing1.Locale = New System.Globalization.CultureInfo("en-US")
        Me.CollPayPricing1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'lblOutOfHoursTravelRate
        '
        Me.lblOutOfHoursTravelRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOutOfHoursTravelRate.Location = New System.Drawing.Point(533, 112)
        Me.lblOutOfHoursTravelRate.Name = "lblOutOfHoursTravelRate"
        Me.lblOutOfHoursTravelRate.Size = New System.Drawing.Size(178, 23)
        Me.lblOutOfHoursTravelRate.TabIndex = 53
        Me.lblOutOfHoursTravelRate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblInHoursTravelPayRate
        '
        Me.lblInHoursTravelPayRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInHoursTravelPayRate.Location = New System.Drawing.Point(533, 88)
        Me.lblInHoursTravelPayRate.Name = "lblInHoursTravelPayRate"
        Me.lblInHoursTravelPayRate.Size = New System.Drawing.Size(178, 23)
        Me.lblInHoursTravelPayRate.TabIndex = 52
        Me.lblInHoursTravelPayRate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOutHoursTravel
        '
        Me.lblOutHoursTravel.Location = New System.Drawing.Point(421, 112)
        Me.lblOutHoursTravel.Name = "lblOutHoursTravel"
        Me.lblOutHoursTravel.Size = New System.Drawing.Size(100, 23)
        Me.lblOutHoursTravel.TabIndex = 51
        Me.lblOutHoursTravel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblinHoursTravel
        '
        Me.lblinHoursTravel.Location = New System.Drawing.Point(421, 88)
        Me.lblinHoursTravel.Name = "lblinHoursTravel"
        Me.lblinHoursTravel.Size = New System.Drawing.Size(100, 23)
        Me.lblinHoursTravel.TabIndex = 50
        Me.lblinHoursTravel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtoutTravelTime
        '
        Me.dtoutTravelTime.Enabled = False
        Me.dtoutTravelTime.Location = New System.Drawing.Point(311, 112)
        Me.dtoutTravelTime.Name = "dtoutTravelTime"
        Me.dtoutTravelTime.Size = New System.Drawing.Size(100, 20)
        Me.dtoutTravelTime.TabIndex = 49
        Me.dtoutTravelTime.TabStop = False
        '
        'dtInTravelTime
        '
        Me.dtInTravelTime.Enabled = False
        Me.dtInTravelTime.Location = New System.Drawing.Point(312, 85)
        Me.dtInTravelTime.Name = "dtInTravelTime"
        Me.dtInTravelTime.Size = New System.Drawing.Size(100, 20)
        Me.dtInTravelTime.TabIndex = 47
        Me.dtInTravelTime.TabStop = False
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(29, 88)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(72, 23)
        Me.Label27.TabIndex = 42
        Me.Label27.Text = "Travel Time"
        '
        'dtTravelTime
        '
        Me.dtTravelTime.Enabled = False
        Me.dtTravelTime.Location = New System.Drawing.Point(104, 85)
        Me.dtTravelTime.Name = "dtTravelTime"
        Me.dtTravelTime.Size = New System.Drawing.Size(100, 20)
        Me.dtTravelTime.TabIndex = 54
        Me.dtTravelTime.TabStop = False
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(219, 84)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(80, 23)
        Me.Label24.TabIndex = 55
        Me.Label24.Text = "In Hours (hr)"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(216, 109)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(96, 23)
        Me.Label25.TabIndex = 56
        Me.Label25.Text = "Out of Hours (hr)"
        '
        'cbInvStatus
        '
        Me.cbInvStatus.DisplayMember = "PhraseText"
        Me.cbInvStatus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbInvStatus.Location = New System.Drawing.Point(520, 24)
        Me.cbInvStatus.Name = "cbInvStatus"
        Me.cbInvStatus.PhraseType = Nothing
        Me.cbInvStatus.Size = New System.Drawing.Size(121, 21)
        Me.cbInvStatus.TabIndex = 21
        Me.cbInvStatus.TabStop = False
        Me.cbInvStatus.ValueMember = "PhraseID"
        '
        'frmCollPay
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 857)
        Me.Controls.Add(Me.gbCalculation)
        Me.Controls.Add(Me.gbExpenses)
        Me.Controls.Add(Me.gbCosts)
        Me.Controls.Add(Me.pnlTimes)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmCollPay"
        Me.Text = "frmCollPay"
        Me.Panel1.ResumeLayout(False)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.pnlTimes.ResumeLayout(False)
        Me.pnlTimes.PerformLayout()
        Me.gbCosts.ResumeLayout(False)
        CType(Me.udDonors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbExpenses.ResumeLayout(False)
        CType(Me.udNonItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udOther, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udAgreed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udTaxi, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udHotel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UDRail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udAir, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbCalculation.ResumeLayout(False)
        Me.gbCalculation.PerformLayout()
        CType(Me.udMileage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udHalfDays, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CollPayPricing1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CollPay As Intranet.intranet.jobs.CollPay
    Private objColl As Intranet.intranet.jobs.CMJob
    Private objOfficer As Intranet.intranet.jobs.CollectingOfficer
    Private dblAverage As Double = 0
    Private blnUk As Boolean = True
    Private Payrates As CollPayPricing.PayRatesRow
    Private blnDrawing As Boolean = True

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in the pay details to display 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PayDetails() As Boolean
        Get
            If CollPay Is Nothing Then

                Return False
            Else
                Return Not CollPay.NoPayDetails
            End If
        End Get
    End Property

    Public Property PayItem() As Intranet.intranet.jobs.CollPay
        Get
            Return Me.CollPay
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CollPay)
            Me.CollPay = Value
            Me.FillForm()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pass in the collection for the job 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Collection() As Intranet.intranet.jobs.CMJob

        Set(ByVal Value As Intranet.intranet.jobs.CMJob)
            objColl = Value
            Me.objOfficer = Nothing
            If Not objColl Is Nothing Then
                'CollPay = New Intranet.intranet.jobs.CollPay(objColl.ID, objColl.CollOfficer)
                CollPay = objColl.CollectorsPay
                If Not CollPay Is Nothing Then
                    FillForm()          'Fill the form 
                End If
            Else
                ClearForm()
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sort out display of the times 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DoTimes()
        Dim outTime As TimeSpan
        Dim homeTime As TimeSpan
        'Will need to add travel times
        With Me.CollPay
            Me.dtTimeOnSite.Text = .TimeOnSite.ToString()
            Me.dtTimeOnJob.Text = .TimeSpent.ToString
            outTime = Me.dtArrive.Value.Subtract(Me.dtTravelStart.Value)
            Me.dtOutTime.Text = outTime.ToString
            homeTime = dtTravelEnd.Value.Subtract(dtLeave.Value)
            Me.dtHometime.Text = homeTime.ToString
            If .Mileage > 0 Then
                Me.txtAvgSpeedOut.Text = (((.Mileage / 2) / outTime.TotalMinutes) * 60).ToString("0.00")
                Me.txtavgSpeedHome.Text = (((.Mileage / 2) / homeTime.TotalMinutes) * 60).ToString("0.00")
                If Me.dblAverage > 0 Then
                    If Math.Abs(dblAverage - .Mileage) / .Mileage > 0.2 Then
                        Me.ErrorProvider1.SetError(Me.udMileage, "Check Mileage")
                    End If
                End If
            End If
            If TimeSpan.Compare(.TimeOnSite, .ExpectedTimeOnSite) > 0 Then
                'Possible time problem 
                If .TimeOnSite.Subtract(.ExpectedTimeOnSite).TotalMinutes > 60 Then
                    'more than an hour over
                    Me.ErrorProvider1.SetError(Me.dtTimeOnSite, "Time on site seems long")
                End If
            End If
        End With
        DoHours()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the contents of the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillCollPay()
        With Me.objColl
            Me.CollPay.Officer = .CollOfficer
            'Me.CollPay.ArriveOnSite = .CollectionDate
            'Me.CollPay.TravelStart = .CollectionDate.AddHours(-1)
            If .NoReceived = 0 Then
                Me.CollPay.Donors = .NoOfDonors
            Else
                Me.CollPay.Donors = .NoReceived
            End If
            'Me.CollPay.LeaveSite = .CollectionDate.AddMinutes(60 + .NoOfDonors * 20)
            'Me.CollPay.TravelEnd = Me.CollPay.LeaveSite.AddHours(1)
            Me.CollPay.Mileage = 0
            Me.CollPay.PayPerSample = True
            If .CollectionType = "O" OrElse .CollectionType = "C" Then
                Me.CollPay.PayPerSample = False
            End If
            CollPay.ArriveOnSite = .CollectionDate
            CollPay.LeaveSite = .CollectionDate
            CollPay.TravelEnd = .CollectionDate
            CollPay.TravelStart = .CollectionDate
            Me.Text = "Entering new Pay Details"
        End With
        Me.UpdateMileage()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillForm()

        If Me.CollPay Is Nothing Then Exit Sub
        Me.Text = "Viewing modifying pay details"
        If objColl Is Nothing Then objColl = New Intranet.intranet.jobs.CMJob(CollPay.CID)
        Me.lblCollection.Text = "Collection : " & objColl.ID
        Me.lblCollectionDate.Text = objColl.CollectionDate.ToString("ddd dd-MMM-yyyy")

        If CollPay.NoPayDetails Then
            FillCollPay()
        End If
        With Me.CollPay
            Me.ckCallOutRate.Hide()
            Me.ckCallOutRate.Checked = False
            If .PayType = CollPay.CST_PAYTYPE_OTHERCOST Then
                Me.ckCallOutRate.Show()
            End If
            Try
                If .TravelStart = MedscreenLib.DateField.ZeroDate Then
                    .TravelStart = objColl.CollectionDate
                    .TravelEnd = objColl.CollectionDate
                    .LeaveSite = objColl.CollectionDate
                    .ArriveOnSite = objColl.CollectionDate
                End If
                Me.dtTravelStart.Value = .TravelStart()
                Me.dtArrive.Value = .ArriveOnSite
                Me.dtLeave.Value = .LeaveSite
                Me.dtTravelEnd.Value = .TravelEnd
            Catch ex As Exception
            End Try
            If .ChargesAgreed() Then
                Me.lblAgreedCosts.Text = "Charges agreed for this site"
            Else
                Me.lblAgreedCosts.Text = "No agreed charges"
            End If


            If .InvoiceMonth.Trim.Length > 0 Then
                Me.DTMonth.Value = Date.Parse("01 " & .InvoiceMonth)
            Else
                Me.DTMonth.Value = Date.Parse("01-" & Today.ToString("MMM-yyyy"))  'Set to start of month 
            End If

            'make sure we have an officer for what follows.
            If objOfficer Is Nothing Then

                If .Officer.Length > 0 Then
                    objOfficer = cSupport.CollectingOfficers.Item(.Officer)
                    Dim strCurrency As String = CDbl("0.0").ToString("c", objOfficer.Culture).Chars(0)

                    Me.lblCurrency1.Text = strCurrency
                    Me.lblCurrency2.Text = strCurrency
                    Me.lblCurrency3.Text = strCurrency

                End If
            End If
            Try
                lblOtherOfficers.Text = .OtherOfficers(objColl.ID, .Officer)
            Catch
            End Try
            'Set up default pay rates
            SetBasePayRates(objColl.CollectionDate)
            'check to see if we have passed the point where per sample pay is allowed 

            If DTMonth.Value >= cstStartPayMonth OrElse Today > SystemChangeDate AndAlso (objOfficer.CountryId = 44 Or objOfficer.CountryId = 353) Then
                Me.ckPerSample.Checked = False
                Me.ckPerSample.Hide()
            Else
                Me.ckPerSample.Checked = .PayPerSample
            End If


            Me.udAgreed.Value = .AgreedExpenses
            Me.udMileage.Value = .Mileage
            Dim intIndex As Integer = CType(Me.cbInvStatus.DataSource, MedscreenLib.Glossary.PhraseCollection).IndexOf(.InvoiceStatus)

            Me.cbInvStatus.SelectedIndex = intIndex 'Me.cbInvStatus.DataSource.indexof(.InvoiceStatus)

            Me.dblAverage = cSupport.CollectingOfficers.AvgMileage(.Officer, Me.objColl.Port)
            Me.lblAverageMileage.Text = "average : " & dblAverage.ToString
            Me.udHotel.Value = .HotelExp
            Me.UDRail.Value = .RailExp
            Me.udAir.Value = .AirExpenses
            Me.udOther.Value = .OtherExpenses
            Me.udTaxi.Value = .TaxiExp
            Me.udNonItem.Value = Math.Abs(.ExpensesTotal - (.HotelExp + .RailExp + .AirExpenses + .OtherExpenses + .TaxiExp) - (Me.udMileage.Value * Me.Mileage))
            Me.ErrorProvider1.SetError(Me.dtTimeOnSite, "")
            Me.txtTotalCost.Text = .TotalCost.ToString("0.00")
            Me.txtExpTotal.Text = .ExpensesTotal.ToString("0.00")
            Me.txtCollCost.Text = .CollTotal.ToString("0.00")
            'Mimimum pay?
            If .CollTotal = HourlyPayMinimum Then
                lblCollCost.Text = "Minimum Pay"
            End If
            txtComment.Text = .Comments

            DoTimes()
            Me.DTExpectedTimeOnSite.Text = .ExpectedTimeOnSite.ToString
            If objOfficer.CountryId = 44 OrElse objOfficer.CountryId = 1 OrElse objOfficer.CountryId = 353 Then
                Me.udKMMiles.SelectedIndex = 0
                Mileage = RatePerMile
            Else
                Me.udKMMiles.SelectedIndex = 1
            End If
            If Me.CollPay.ItemCode = ExpenseTypeCollections.GCST_PayPerMile Then Me.udKMMiles.SelectedIndex = 0
            If Me.CollPay.ItemCode = ExpenseTypeCollections.GCST_PayPerKM Then Me.udKMMiles.SelectedIndex = 1

            Dim strCountryCode As String = MedscreenLib.CConnection.PackageStringList("Lib_utils.CountryISOCodeFromID", objOfficer.CountryId)
            If strCountryCode Is Nothing Then strCountryCode = "UK"
            If strCountryCode = "GB" Then strCountryCode = "UK"
            'Find out if it is a holiday
            Dim cColl As New Collection()
            cColl.Add(objColl.CollectionDate.ToString("dd-MMM-yy"))
            cColl.Add(strCountryCode)
            Me.lblHoliday.Text = MedscreenLib.CConnection.PackageStringList("Lib_utils.HolidayIs", cColl)

            Me.lkOfficer.Text = "Officer " & .Officer
            Me.lkOfficer.Links.Clear()
            If Not .Officer Is Nothing Then
                Payrates = Nothing
                Me.lkOfficer.Links.Add(8, .Officer.Length)
                Dim pr As CollPayPricing.PayRatesRow
                Dim strCountry As String = objOfficer.GetMyCountry
                SetBasePayRates(objColl.CollectionDate)
                If strCountry.Trim.Length > 0 Then
                    For Each pr In CollPayPricing1.PayRates.Rows
                        If pr.Country = strCountry Then
                            Me.Payrates = pr
                            GetPayRates()
                            Exit For
                        End If
                    Next
                End If
            End If
            Me.lkPort.Text = "Port " & objColl.Port
            Me.lkPort.Links.Clear()
            Me.lkPort.Links.Add(5, objColl.Port.Length)
            If .JobName.Trim.Length = 0 AndAlso Not objColl Is Nothing Then
                Me.lblJob.Text = "Job : " & objColl.JobName
            Else
                Me.lblJob.Text = "Job : " & .JobName
            End If
            Me.lblNDonors.Text = "No of Donors : " & .Donors
            Me.udDonors.Value = .Donors
            'see if we are going to pay half daily
            Dim oColl As New Collection()
            oColl.Add(objColl.ClientId)
            oColl.Add("COLLBOOK")
            Dim strBooking As String = MedscreenLib.CConnection.PackageStringList("Lib_customer.GetOptionDefault", oColl)
            'Do pay type LAbels
            lblPayType.Text = ""
            If .ExpenseTypeID = .CST_PAYTYPE_DAY Then
                lblPayType.Text = "Day Rate"
            ElseIf .ExpenseTypeID = .CST_PAYTYPE_HALFDAY Then
                lblPayType.Text = "Half Day Rate"
            ElseIf .ExpenseTypeID = .CST_PAYTYPE_FIXED Then
                lblPayType.Text = "Fixed Rate (variable)"
            End If
            lblmileage3.Width = 100
            'lblmileage3.AutoSize = False
            If strBooking <> "NOTUSED" Then
                Me.ckHalfDay.Show()
                Me.ckHalfDay.Checked = True
                Me.lblHalfDays.Show()
                Me.udHalfDays.Show()
                Me.udHalfDays.Value = .Quantity
                Me.lblDonorPay.Left = 500
                Me.lblDonorPay.Width = 100
            Else
                ckCallOutRate.Show()
                If .ExpenseTypeID = .CST_PAYTYPE_HALFDAY Or .ExpenseTypeID = .CST_PAYTYPE_DAY Then
                    Me.ckHalfDay.Show()
                    Me.ckHalfDay.Checked = True
                    Me.lblHalfDays.Show()
                    Me.udHalfDays.Show()
                    'Me.lblMileage4.Hide()
                    If .ExpenseTypeID = .CST_PAYTYPE_HALFDAY Then
                        Me.lblHalfDays.Text = "Half Days"
                    Else
                        Me.lblHalfDays.Text = "Days"
                    End If
                Else
                    Me.ckHalfDay.Hide()
                    Me.ckHalfDay.Checked = False
                    Me.lblHalfDays.Hide()
                    Me.udHalfDays.Hide()
                    ' Me.lblMileage4.Show()

                End If
                Me.lblDonorPay.Width = 400
                Me.udHalfDays.Value = 0
                Me.lblDonorPay.Left = 208
            End If
            Me.DoHours()
            Me.DoSample()

            Try
                Try
                    If objColl.CollectionType = "C" Or objColl.CollectionType = "O" Then
                        Me.lblTotalPay.Text = TotalCallOutOnSiteCost.ToString("c", objOfficer.Culture)
                    Else
                        Me.lblTotalPay.Text = TotalRoutineOnSiteCost.ToString("c", objOfficer.Culture)
                    End If
                Catch
                End Try
                oColl = New Collection()
                oColl.Add("COLLTYPE")
                oColl.Add(objColl.CollectionType)
                Me.lblCollType.Text = MedscreenLib.CConnection.PackageStringList("Lib_utils.DecodePhrase", oColl)
            Catch
            End Try
            Try
                lbllmileage4.Text = ""
                Me.lblMilagePay.Text = CDbl(Me.udMileage.Value * Me.Mileage).ToString("c", objOfficer.Culture)
                Me.lblMileage2.Text = "Travel " & CDbl(Me.udMileage.Value * Me.Mileage).ToString("c", objOfficer.Culture)
                Me.lblmileage3.Text = lblMileage2.Text
                oColl = New Collection
                oColl.Add(Me.objOfficer.OfficerID)
                If objColl.CollectionDate.Month >= 4 Then
                    oColl.Add(objColl.CollectionDate.Year)
                Else
                    oColl.Add(objColl.CollectionDate.Year - 1)
                End If
                Dim TotMiles As String = MedscreenLib.CConnection.PackageStringList("lib_collpay.OfficerTotalMileage", oColl)
                Me.lbllmileage4.Text += lblmileage3.Text & "@ " & Mileage.ToString("c", objOfficer.Culture) & " after " & TotMiles & " Miles"
                lblmileage3.Width = 300
                DisplayPayRates()
                Me.UpdateMileage()
                If .CollTotal = HourlyPayMinimum Then
                    lblCollCost.Text = "Minimum Pay"
                End If
            Catch
            End Try
        End With
    End Sub
    'On site rates
    Dim inHrsRoutineHourlyRate As Double
    Dim outHrsRoutineHourlyRate As Double
    Dim inHrsCallOutHourlyRate As Double
    Dim outHrsCallOutHourlyRate As Double
    ' travel rates
    Dim inHrsRoutineTravelHourlyRate As Double
    Dim outHrsRoutineTravelHourlyRate As Double
    Dim inHrsCallOutTravelHourlyRate As Double
    Dim outHrsCallOutTravelHourlyRate As Double

    Dim HourlyPayMinimum As Double
    Dim Mileage As Double
    Dim RatePerMile As Double
    Dim RatePerKM As Double
    Dim HalfDay As Double
    Dim DayRate As Double
    Dim SampleRate As Double

    Private Function GetOColl(ByVal LinkCode As String, ByVal CollDate As Date) As Collection
        Dim oColl As New Collection()
        oColl.Add(Me.objOfficer.OfficerID)
        oColl.Add(LinkCode)
        oColl.Add(CollDate.ToString("dd-MMM-yy"))
        Return oColl
    End Function
    Private Sub SetBasePayRates(ByVal CollDate As Date)
        'Const cstPAYCALLOUTINRS001 As String = "PAYCALLOUTINRS001"      'Call Out In Hours Pay                           
        'Const cstPAYDAYRATE As String = "PAYDAYRATE"                    'Collection Attendance Day rate                  
        'Const cstPAYHALFDAY As String = "PAYHALFDAY"                    'Collection Attendance Half Day rate
        'Const cstPAYROUNTINHRS001 As String = "PAYROUNTINHRS001"        'Routine In Hours Pay. 
        'Const cstPAYSAMPLE001 As String = "PAYSAMPLE001"                'Routine per DOA Sample Rate C/O pay   
        'Const cstPAYSAMPLE002 As String = "PAYSAMPLE002"                'Benzene per sample 
        'Const cstPAYSAMPLE003 As String = "PAYSAMPLE003"                'Routine DOA and Benzene per sample  
        'Const cstPAYSAMPLE004 As String = "PAYSAMPLE004"                'Breath Test per sample  
        'Const cstPAYSAMPLE005 As String = "PAYSAMPLE005"                'Breath Test (Witness) per sample  
        'Const cstPAYSAMPLE006 As String = "PAYSAMPLE006"                'Breath test with Urine sample  
        'Const cstPAYWRKTIME As String = "PAYWRKTIME"                    'Collecting Officer Working Time pr hr  
        'Const cstPAYHOURMIN As String = "PAYHOURMIN"                    'Collecting Officer Minimum Hourly Amount

        If CollDate.Year = 1850 Then CollDate = DTMonth.Value : objColl.CollectionDate = DTMonth.Value

        Dim oColl As Collection = GetOColl(cstSAMPLE, CollDate)
        SampleRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        'On site rates
        oColl = GetOColl(cstPAYROUNTINHRS001, CollDate)
        inHrsRoutineHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstPAYROUNTOUTHRS001, CollDate)
        outHrsRoutineHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstPAYCALLOUTINRS001, CollDate)
        inHrsCallOutHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstPAYCALLOUTOUTS001, CollDate)
        outHrsCallOutHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        'Travel rates
        oColl = GetOColl(cstPAYROUNTINHRS002, CollDate)
        Try
            inHrsRoutineTravelHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
            oColl = GetOColl(cstPAYROUNTOUTHRS002, CollDate)
            outHrsRoutineTravelHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
            oColl = GetOColl(cstPAYCALLOUTINRS002, CollDate)
            inHrsCallOutTravelHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
            oColl = GetOColl(cstPAYCALLOUTOUTS002, CollDate)
            outHrsCallOutTravelHourlyRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        Catch ex As Exception
        End Try
        'others
        oColl = GetOColl(cstPAYHOURMIN, CollDate)
        HourlyPayMinimum = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstMILEAGE, CollDate)
        Mileage = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstPAYHALFDAY, CollDate)
        HalfDay = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)
        oColl = GetOColl(cstPAYDAYRATE, CollDate)
        DayRate = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetPayRate", oColl)

    End Sub

    Private Function GetPayRates() As Boolean
        'inHrsRoutineHourlyRate = 0
        'outHrsRoutineHourlyRate = 0
        'Mileage = 0
        'RatePerMile = 0
        'RatePerKM = 0
        'HalfDay = 0
        'Dim CollectionType As String = "ROUTINE"
        'If Me.objColl.CollectionType = "C" Or Me.objColl.CollectionType = "O" Then
        '    CollectionType = "CALLOUT"
        '    Dim objRow As CollPayPricing.PayRateRow
        '    Dim payrateTable As CollPayPricing.PayRateRow() = Payrates.GetPayRateRows
        '    For Each objRow In payrateTable
        '        If objRow.PayRateType = CollectionType Then
        '            If objRow.OutofHours Then
        '                outHrsRoutineHourlyRate = objRow.PayRate
        '            Else
        '                inHrsRoutineHourlyRate = objRow.PayRate
        '            End If
        '        ElseIf objRow.PayRateType = "MILEAGE" Then
        '            RatePerMile = objRow.PayRate
        '        ElseIf objRow.PayRateType = "PERKM" Then
        '            RatePerKM = objRow.PayRate
        '        ElseIf objRow.PayRateType = "HALFDAY" Then
        '            HalfDay = objRow.PayRate
        '        End If
        '    Next
        'Else
        '    CollectionType = "ROUTINE"
        '    Dim objRow As CollPayPricing.PayRateRow
        '    Dim payrateTable As CollPayPricing.PayRateRow() = Payrates.GetPayRateRows
        '    For Each objRow In payrateTable
        '        If objRow.PayRateType = CollectionType Then
        '            If objRow.OutofHours Then
        '                outHrsRoutineHourlyRate = objRow.PayRate
        '            Else
        '                inHrsRoutineHourlyRate = objRow.PayRate
        '            End If
        '        ElseIf objRow.PayRateType = "MILEAGE" Then
        '            RatePerMile = objRow.PayRate
        '        ElseIf objRow.PayRateType = "PERKM" Then
        '            RatePerKM = objRow.PayRate
        '        ElseIf objRow.PayRateType = "HALFDAY" Then
        '            HalfDay = objRow.PayRate
        '        ElseIf objRow.PayRateType = "SAMPLE" Then
        '            SampleRate = objRow.PayRate
        '        End If
        '    Next
        '    Mileage = RatePerMile
        'End If
    End Function


    Private Function GetForm() As Boolean
        Dim blnRet As Boolean = True

        With Me.CollPay
            Try
                .ArriveOnSite = Me.dtArrive.Value
                .LeaveSite = Me.dtLeave.Value
                .TravelEnd = Me.dtTravelEnd.Value
                .TravelStart = Me.dtTravelStart.Value
                .TotalCost = Me.txtTotalCost.Text
                .ExpensesTotal = Me.txtExpTotal.Text
                .CollTotal = Me.txtCollCost.Text
                .Mileage = Me.udMileage.Value
                .AirExpenses = Me.udAir.Value
                .HotelExp = Me.udHotel.Value
                .RailExp = Me.UDRail.Value
                .OtherExpenses = Me.udOther.Value
                .TaxiExp = Me.udTaxi.Value
                .AgreedExpenses = Me.udAgreed.Value
                .Donors = Me.udDonors.Value
                .InvoiceMonth = Me.DTMonth.Value.ToString("MMMM yyyy")
                .PayPerSample = Me.ckPerSample.Checked
                .InvoiceStatus = "E"
                .JobName = objColl.JobName
                .Comments = txtComment.Text
                If Me.ckHalfDay.Checked Then
                    .Quantity = Me.udHalfDays.Value
                    .PayType = "PAYHALFDAY"
                End If
            Catch ex As Exception
                blnRet = False
            End Try
            Return blnRet
        End With
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Remove the contents of the form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ClearForm()
        Me.lblCollection.Text = "Collection : "
        Me.dtArrive.Value = DateSerial(1970, 1, 1)
        Me.dtLeave.Value = DateSerial(1970, 1, 1)
        Me.dtTravelEnd.Value = DateSerial(1970, 1, 1)
        Me.dtTravelStart.Value = DateSerial(1970, 1, 1)
        Me.udAgreed.Value = 0
        Me.udMileage.Value = 0
        Me.udHotel.Value = 0
        Me.UDRail.Value = 0
        Me.lblNDonors.Text = "No of Donors"
        Me.dtTimeOnSite.Text = ""
        Me.dtTimeOnJob.Text = ""
        Me.lkOfficer.Text = "Officer "
        Me.lkOfficer.Links.Clear()
        Me.lkPort.Text = "Port"
        Me.lkPort.Links.Clear()
        Me.lblJob.Text = "Job : "
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTExpectedTimeOnSite.TextChanged

    End Sub

    Private Sub dtArrive_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtArrive.Validating
        If Date.Compare(Me.dtTravelStart.Value, Me.dtArrive.Value) > 0 Then
            e.Cancel = True
            Me.ErrorProvider1.SetError(Me.dtArrive, "Arrive on site must be after travel start")
        Else
            Me.ErrorProvider1.SetError(Me.dtArrive, "")
            CollPay.ArriveOnSite = Me.dtArrive.Value
            Me.DoTimes()
        End If
        If Me.dtLeave.Value < Me.dtArrive.Value Then Me.dtLeave.Value = Me.dtArrive.Value
        If Me.dtTravelEnd.Value < Me.dtLeave.Value Then Me.dtTravelEnd.Value = Me.dtLeave.Value
        ValidateOtherCollections(dtArrive)

    End Sub




    Private Sub dtLeave_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtLeave.Validating
        If Date.Compare(Me.dtArrive.Value, Me.dtLeave.Value) > 0 Then
            e.Cancel = True
            Me.ErrorProvider1.SetError(Me.dtLeave, "Can't leave before arrival")
        Else
            Me.ErrorProvider1.SetError(Me.dtLeave, "")
            Me.CollPay.LeaveSite = Me.dtLeave.Value
            DoTimes()
        End If
        If Date.Compare(Me.dtLeave.Value, Me.dtTravelEnd.Value) > 0 Then
            If Me.dtTravelEnd.Value < Me.dtLeave.Value Then Me.dtTravelEnd.Value = Me.dtLeave.Value
        End If
        ValidateOtherCollections(dtLeave)

    End Sub

    Private Sub ValidateOtherCollections(ByVal DTPicker As Windows.Forms.DateTimePicker)
        Dim ocoll As New Collection
        ocoll.Add(objOfficer.OfficerID)
        ocoll.Add(DTPicker.Value.ToString("dd-MMM-yy HH:mm"))
        ocoll.Add(objColl.ID)
        Me.ErrorProvider1.SetError(DTPicker, "")
        Dim strResp As String = MedscreenLib.CConnection.PackageStringList("Lib_collpay.OverlappingJobs", ocoll, True)
        If strResp.Trim.Length > 0 Then
            ocoll = New Collection
            ocoll.Add(objOfficer.OfficerID)
            ocoll.Add(strResp)
            Dim strTimes As String = MedscreenLib.CConnection.PackageStringList("Lib_collpay.GetJobTimes", ocoll, True)
            Me.ErrorProvider1.SetError(DTPicker, "Overlaps with collection " & strResp & "  " & strTimes)
            'e.Cancel = True
        End If
    End Sub
    Private Sub dtTravelStart_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtTravelStart.Validating
        If Me.dtArrive.Value < Me.dtTravelStart.Value Then Me.dtArrive.Value = Me.dtTravelStart.Value
        If Me.dtLeave.Value < Me.dtArrive.Value Then Me.dtLeave.Value = Me.dtArrive.Value
        If Me.dtTravelEnd.Value < Me.dtLeave.Value Then Me.dtTravelEnd.Value = Me.dtLeave.Value
        If dtArrive.Value.Subtract(dtTravelStart.Value).TotalDays > 2 Then
            dtLeave.Value = dtTravelStart.Value
            dtTravelEnd.Value = dtTravelStart.Value
            dtArrive.Value = dtTravelStart.Value
        End If
        ValidateOtherCollections(dtTravelStart)
        'Me.CollPay.TravelStart = Me.dtTravelStart.Value
        DoTimes()
    End Sub

    Private Sub udMileage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles udMileage.Leave
        Me.CollPay.Mileage = Me.udMileage.Value
        DoTimes()
        UpdateMileage()
    End Sub


    Private Sub udMileage_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles udMileage.Validated
        UpdateMileage()

    End Sub

    Private Sub lkOfficer_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkOfficer.LinkClicked
        Dim strOfficer As String


        Me.lkOfficer.Links(lkOfficer.Links.IndexOf(e.Link)).Visited = True

        With CommonForms.FormEdCollOfficer
            .CollectingOfficer = Nothing
            .CollectingOfficer = objOfficer
            .ShowDialog()
        End With
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.GetForm Then
            CollPay.DoUpdate()
        End If
    End Sub

    Dim TotalTimeOnSite As Double
    Dim OutHoursOnSite As Double
    Dim inHoursOnSite As Double
    Dim TotalTravelTime As Double
    Dim TotalOutofHoursTravelTime As Double
    Dim TotalTravelinHours As Double

    Private Function InhoursCallOutTravelPayAsString() As String
        Return CDbl(TotalTravelinHours * Me.inHrsCallOutTravelHourlyRate).ToString("c", objOfficer.Culture)
    End Function

    Private Function OutHoursCallOutTravelPayAsString() As String
        Return CDbl(TotalOutofHoursTravelTime * Me.outHrsCallOutTravelHourlyRate).ToString("c", objOfficer.Culture)
    End Function


    Private Function InHoursRoutineTravelPayAsString() As String
        Return CDbl(TotalTravelinHours * Me.inHrsRoutineTravelHourlyRate).ToString("c", objOfficer.Culture)
    End Function

    Private Function OutHoursRoutineTravelPayasString() As String
        Return CDbl(TotalOutofHoursTravelTime * Me.outHrsRoutineTravelHourlyRate).ToString("c", objOfficer.Culture)
    End Function

    Private Function InhoursCallOutPayAsString() As String
        Return CDbl(inHoursOnSite * Me.inHrsCallOutHourlyRate).ToString("c", objOfficer.Culture)
    End Function

    Private Function OutHoursCallOutPayAsString() As String
        Return CDbl(OutHoursOnSite * Me.outHrsCallOutHourlyRate).ToString("c", objOfficer.Culture)
    End Function


    Private Function InHoursRoutinePayAsString() As String
        Return CDbl(inHoursOnSite * Me.inHrsRoutineHourlyRate).ToString("c", objOfficer.Culture)
    End Function

    Private Function OutHoursRoutinePayasString() As String
        Return CDbl(OutHoursOnSite * Me.outHrsRoutineHourlyRate).ToString("c", objOfficer.Culture)
    End Function


    Private Function TotalCallOutOnSiteCost() As Double
        Return CDbl(OutHoursOnSite * Me.outHrsCallOutHourlyRate + inHoursOnSite * Me.inHrsCallOutHourlyRate + TotalTravelinHours * Me.inHrsCallOutTravelHourlyRate + TotalOutofHoursTravelTime * Me.outHrsCallOutTravelHourlyRate)
    End Function

    Private Function TotalRoutineOnSiteCost() As Double
        Return CDbl(OutHoursOnSite * Me.outHrsRoutineHourlyRate + inHoursOnSite * Me.inHrsRoutineHourlyRate + TotalTravelinHours * Me.inHrsRoutineTravelHourlyRate + inHoursOnSite * Me.inHrsRoutineHourlyRate)
    End Function

    Private Sub DoHours()
        'This will need upgrading to deal with travel times, also need to deal with variable time frames
        Dim tsTravelOutOutofHours As TimeSpan = MedscreenLib.Medscreen.HoursOutOfHours(Me.dtTravelStart.Value, Me.dtArrive.Value, objOfficer.GetMyCountry)
        Dim tsTravelHomeOutofHours As TimeSpan = MedscreenLib.Medscreen.HoursOutOfHours(Me.dtLeave.Value, Me.dtTravelEnd.Value, objOfficer.GetMyCountry)
        'Get total travel time
        Dim tsTotalTravelTime As TimeSpan = dtArrive.Value.Subtract(dtTravelStart.Value).Add(dtTravelEnd.Value.Subtract(dtLeave.Value))
        TotalTravelTime = dtArrive.Value.Subtract(dtTravelStart.Value).TotalHours + dtTravelEnd.Value.Subtract(dtLeave.Value).TotalHours
        TotalOutofHoursTravelTime = tsTravelHomeOutofHours.TotalHours + tsTravelOutOutofHours.TotalHours
        TotalTravelinHours = TotalTravelTime - TotalOutofHoursTravelTime

        Dim TSOutofHoursOnSite As TimeSpan = MedscreenLib.Medscreen.HoursOutOfHours(Me.dtArrive.Value, Me.dtLeave.Value, objOfficer.GetMyCountry)
        OutHoursOnSite = TSOutofHoursOnSite.TotalHours
        inHoursOnSite = (TotalTimeOnSite - OutHoursOnSite)
        TotalTimeOnSite = Me.dtLeave.Value.Subtract(Me.dtArrive.Value).TotalHours
        'needs decorating with info 
        Me.txtOutofHours.Text = TSOutofHoursOnSite.ToString
        Me.dtoutTravelTime.Text = tsTravelHomeOutofHours.Add(tsTravelOutOutofHours).ToString
        Me.dtTravelTime.Text = tsTotalTravelTime.ToString
        Me.dtInTravelTime.Text = tsTotalTravelTime.Subtract(tsTravelHomeOutofHours.Add(tsTravelOutOutofHours)).ToString

        If objColl Is Nothing Then Exit Sub
        If objColl.CollectionType = "C" Or objColl.CollectionType = "O" OrElse ckCallOutRate.Checked Then                       'If call out or paid at call out rate
            Me.lblouthourspay.Text = OutHoursCallOutPayAsString()
            Me.lblinhrspay.Text = InhoursCallOutPayAsString()
            Me.lblOutHoursTravel.Text = OutHoursCallOutTravelPayAsString()
            Me.lblinHoursTravel.Text = InhoursCallOutTravelPayAsString()
            Me.lblCollCost.Text = OutHoursCallOutPayAsString() & " + " & InhoursCallOutPayAsString()
        Else
            Me.lblouthourspay.Text = OutHoursRoutinePayasString()
            Me.lblinhrspay.Text = InHoursRoutinePayAsString()
            Me.lblOutHoursTravel.Text = OutHoursRoutineTravelPayasString()
            Me.lblinHoursTravel.Text = InHoursRoutineTravelPayAsString()
            Me.lblCollCost.Text = OutHoursRoutinePayasString() & " + " & InHoursRoutinePayAsString()
        End If
        DisplayPayRates()
    End Sub

    Dim SamplePay As Double
    Private Sub DoSample()
        If Me.ckHalfDay.Checked Then
            If Me.PayItem.ExpenseTypeID = CollPay.CST_PAYTYPE_HALFDAY Then SamplePay = CInt(Me.udHalfDays.Value) * Me.HalfDay
            If Me.PayItem.ExpenseTypeID = CollPay.CST_PAYTYPE_DAY Then SamplePay = CInt(Me.udHalfDays.Value) * Me.DayRate

            Me.lblDonorPay.Text = SamplePay.ToString("c", objOfficer.Culture)
        Else
            SamplePay = Me.udDonors.Value * Me.SampleRate
            Me.lblDonorPay.Text = Me.udDonors.Value.ToString("0") & " Samples @ " & Me.SampleRate.ToString("c", objOfficer.Culture) & " per sample " & SamplePay.ToString("c", objOfficer.Culture)
        End If
    End Sub

    Private Sub UpdateMileage()
        Me.txtExpTotal.Text = CDbl(Me.udAgreed.Value + Me.udAir.Value + Me.udHotel.Value + Me.udOther.Value + Me.UDRail.Value + Me.udTaxi.Value + (Me.udMileage.Value * Me.Mileage)).ToString("0.00")
        Try
            Me.lblMilagePay.Text = CDbl(Me.udMileage.Value * Me.Mileage).ToString("c", objOfficer.Culture)
            Me.lblMileage2.Text = "Travel " & CDbl(Me.udMileage.Value * Me.Mileage).ToString("c", objOfficer.Culture)
            Me.lblmileage3.Text = lblMileage2.Text
            Me.txtTotalCost.Text = CDbl(txtExpTotal.Text) + CDbl(txtCollCost.Text)

        Catch
        End Try

    End Sub

    Private Sub DoCalculations()
        If Me.PayItem Is Nothing Then Exit Sub
        If Me.PayItem.ExpenseTypeID = CollPay.CST_PAYTYPE_FIXED Then
            ' we don't do calculations for fixed rate
        ElseIf PayItem.ExpenseTypeID = CollPay.CST_PAYTYPE_HALFDAY Then
            Me.txtCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("0.00")
            Me.lblCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
            Me.lblTotalPay.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
        ElseIf PayItem.ExpenseTypeID = "PAYDAYRATE" Then
            Me.txtCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("0.00")
            Me.lblCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
            Me.lblTotalPay.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)

        Else
            If Me.ckPerSample.Checked Then
                DoSample()
                Me.txtCollCost.Text = SamplePay
                Me.lblCollCost.Text = SamplePay.ToString("c", objOfficer.Culture)
            ElseIf ckHalfDay.Checked Then
                Me.txtCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("0.00")
                Me.lblCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
                Me.lblTotalPay.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
            Else
                DoHours()
                If objColl.CollectionType = "C" Or objColl.CollectionType = "O" OrElse ckCallOutRate.Checked Then
                    Me.txtCollCost.Text = TotalCallOutOnSiteCost.ToString("c", objOfficer.Culture)
                Else
                    Me.txtCollCost.Text = TotalRoutineOnSiteCost.ToString("c", objOfficer.Culture)
                End If
                Try
                    If objColl.CollectionType = "C" Or objColl.CollectionType = "O" OrElse ckCallOutRate.Checked Then
                        Me.lblTotalPay.Text = TotalCallOutOnSiteCost.ToString("c", objOfficer.Culture)
                    Else
                        Me.lblTotalPay.Text = TotalRoutineOnSiteCost.ToString("c", objOfficer.Culture)
                    End If
                Catch
                End Try
            End If
        End If
        'Check for minimum payment
        If objOfficer.CountryId = 353 OrElse objOfficer.CountryId = 44 Then  'British officers
            If CDbl(Me.txtCollCost.Text) < HourlyPayMinimum Then
                Me.txtCollCost.Text = HourlyPayMinimum.ToString("0.00")
                Me.lblCollCost.Text = "Minimum Pay"
            End If
        End If

        Me.txtTotalCost.Text = CDbl(CDbl(Me.txtCollCost.Text) + CDbl(Me.txtExpTotal.Text)).ToString("0.00")
        UpdateMileage()
        DisplayPayRates()
    End Sub
    Private Sub cmdCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCalc.Click
        DoCalculations()
    End Sub

    Private Sub DisplayPayRates()
        If objColl Is Nothing Then Exit Sub
        If objColl.CollectionType = "C" OrElse objColl.CollectionType = "O" OrElse Me.ckCallOutRate.Checked Then
            Me.lblinhoursPayRate.Text = inHoursOnSite.ToString("0.00") & "hours @ " & Me.inHrsCallOutHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblOutOfHoursPayRate.Text = OutHoursOnSite.ToString("0.00") & "hours @ " & Me.outHrsCallOutHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblInHoursTravelPayRate.Text = TotalTravelinHours.ToString("0.00") & "hours @ " & Me.inHrsCallOutTravelHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblOutOfHoursTravelRate.Text = TotalOutofHoursTravelTime.ToString("0.00") & "hours @ " & Me.outHrsCallOutTravelHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
        Else
            Me.lblinhoursPayRate.Text = inHoursOnSite.ToString("0.00") & "hours @ " & Me.inHrsRoutineHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblOutOfHoursPayRate.Text = OutHoursOnSite.ToString("0.00") & "hours @ " & Me.outHrsRoutineHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblInHoursTravelPayRate.Text = TotalTravelinHours.ToString("0.00") & "hours @ " & Me.inHrsRoutineTravelHourlyRate.ToString("c", objOfficer.Culture) & " per hour"
            Me.lblOutOfHoursTravelRate.Text = TotalOutofHoursTravelTime.ToString("0.00") & "hours @ " & Me.outHrsRoutineTravelHourlyRate.ToString("c", objOfficer.Culture) & " per hour"

        End If
    End Sub

    Private Sub dtTravelEnd_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtTravelEnd.Validating
        Me.ErrorProvider1.SetError(Me.dtTravelEnd, "")
        If Date.Compare(Me.dtLeave.Value, dtTravelEnd.Value) > 0 Then
            Me.ErrorProvider1.SetError(Me.dtTravelEnd, "Must be after the time site left")
            Me.dtTravelEnd.Value = Me.dtLeave.Value.AddHours(1)
        End If
        Me.CollPay.TravelEnd = Me.dtTravelEnd.Value
        DoTimes()
        ValidateOtherCollections(dtTravelEnd)

    End Sub

    Private Sub dtTravelStart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtTravelStart.ValueChanged

    End Sub

    Private Sub dtArrive_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtArrive.ValueChanged

    End Sub

    Private Sub udKMMiles_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles udKMMiles.SelectedItemChanged
        If Me.blnDrawing Then Exit Sub
        If udKMMiles.SelectedIndex = 0 Then
            'Label4.Text = "Mileage:"
            'If Not Me.MyMileageCost Is Nothing Then
            '    MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerMile
            '    MyMileageCost.DoUpdate()
            'End If
            Mileage = RatePerMile
            If Not Me.CollPay Is Nothing Then Me.CollPay.ItemCode = ExpenseTypeCollections.GCST_PayPerMile
        Else
            'Label4.Text = "Travel:"
            'If Not Me.MyMileageCost Is Nothing Then
            '    MyMileageCost.ExpenseTypeID = ExpenseTypeCollections.GCST_PayPerKM
            '    MyMileageCost.DoUpdate()
            'End If
            Mileage = RatePerKM
            If Not Me.CollPay Is Nothing Then Me.CollPay.ItemCode = ExpenseTypeCollections.GCST_PayPerKM
        End If
        UpdateMileage()
    End Sub

    Private Sub frmCollPay_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        blnDrawing = False
    End Sub

    Private Sub udNonItem_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles udNonItem.ValueChanged, udAir.ValueChanged, udOther.ValueChanged, UDRail.ValueChanged, udHotel.ValueChanged, udTaxi.ValueChanged
        UpdateMileage()
    End Sub

    Private Sub txtCollCost_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCollCost.Leave
        Try
            Me.txtTotalCost.Text = CDbl(txtExpTotal.Text) + CDbl(txtCollCost.Text)
        Catch
        End Try
    End Sub

    Private Sub txtExpTotal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExpTotal.Leave
        Try
            Me.txtTotalCost.Text = CDbl(txtExpTotal.Text) + CDbl(txtCollCost.Text)
        Catch
        End Try
    End Sub

    Private Sub udHalfDays_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles udHalfDays.ValueChanged
        If PayItem.ExpenseTypeID = CollPay.CST_PAYTYPE_HALFDAY Then
            Me.txtCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("0.00")
            Me.lblCollCost.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
            Me.lblDonorPay.Text = CDbl(udHalfDays.Value * HalfDay).ToString("c", objOfficer.Culture)
        Else
            Me.txtCollCost.Text = CDbl(udHalfDays.Value * DayRate).ToString("0.00")
            Me.lblCollCost.Text = CDbl(udHalfDays.Value * DayRate).ToString("c", objOfficer.Culture)
            Me.lblDonorPay.Text = CDbl(udHalfDays.Value * DayRate).ToString("c", objOfficer.Culture)
        End If
        Me.txtTotalCost.Text = CDbl(CDbl(Me.txtCollCost.Text) + CDbl(Me.txtExpTotal.Text)).ToString("0.00")
    End Sub

    Private Sub ckCallOutRate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckCallOutRate.CheckedChanged
        DoHours()
        DoCalculations()
    End Sub

    Private Sub lblOtherOfficers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblOtherOfficers.Click

    End Sub

    Private Sub DTMonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTMonth.ValueChanged

    End Sub
End Class
