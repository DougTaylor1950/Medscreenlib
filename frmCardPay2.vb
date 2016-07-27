Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports Intranet.intranet.jobs

Public Class frmCardPay2
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents txtCreditCard As System.Windows.Forms.TextBox
    Friend WithEvents lblCreditCard As System.Windows.Forms.Label
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtGrossPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtAuthCode As System.Windows.Forms.TextBox
    Friend WithEvents lblAuthCode As System.Windows.Forms.Label
    Friend WithEvents txtEmailAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtHolder As System.Windows.Forms.TextBox
    Friend WithEvents LblHolder As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents lblCmNumber As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCardPay2))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.txtCreditCard = New System.Windows.Forms.TextBox()
        Me.lblCreditCard = New System.Windows.Forms.Label()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtGrossPrice = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtAuthCode = New System.Windows.Forms.TextBox()
        Me.lblAuthCode = New System.Windows.Forms.Label()
        Me.txtEmailAddress = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtHolder = New System.Windows.Forms.TextBox()
        Me.LblHolder = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.label1 = New System.Windows.Forms.Label()
        Me.lblCmNumber = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 349)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(448, 32)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(364, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(284, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCreditCard
        '
        Me.txtCreditCard.Location = New System.Drawing.Point(120, 48)
        Me.txtCreditCard.Name = "txtCreditCard"
        Me.txtCreditCard.Size = New System.Drawing.Size(150, 20)
        Me.txtCreditCard.TabIndex = 33
        Me.txtCreditCard.Text = ""
        '
        'lblCreditCard
        '
        Me.lblCreditCard.Location = New System.Drawing.Point(24, 48)
        Me.lblCreditCard.Name = "lblCreditCard"
        Me.lblCreditCard.TabIndex = 32
        Me.lblCreditCard.Text = "Credit Card No"
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(120, 80)
        Me.txtReference.Name = "txtReference"
        Me.txtReference.TabIndex = 35
        Me.txtReference.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(48, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 34
        Me.Label6.Text = "Reference"
        '
        'txtGrossPrice
        '
        Me.txtGrossPrice.Location = New System.Drawing.Point(120, 152)
        Me.txtGrossPrice.Name = "txtGrossPrice"
        Me.txtGrossPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtGrossPrice.TabIndex = 37
        Me.txtGrossPrice.Text = "0"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(40, 152)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 23)
        Me.Label9.TabIndex = 36
        Me.Label9.Text = "Gross Price"
        '
        'txtAuthCode
        '
        Me.txtAuthCode.Location = New System.Drawing.Point(120, 112)
        Me.txtAuthCode.Name = "txtAuthCode"
        Me.txtAuthCode.Size = New System.Drawing.Size(96, 20)
        Me.txtAuthCode.TabIndex = 56
        Me.txtAuthCode.Text = ""
        '
        'lblAuthCode
        '
        Me.lblAuthCode.Location = New System.Drawing.Point(48, 112)
        Me.lblAuthCode.Name = "lblAuthCode"
        Me.lblAuthCode.Size = New System.Drawing.Size(72, 24)
        Me.lblAuthCode.TabIndex = 55
        Me.lblAuthCode.Text = "Auth. Code"
        '
        'txtEmailAddress
        '
        Me.txtEmailAddress.Location = New System.Drawing.Point(120, 208)
        Me.txtEmailAddress.Name = "txtEmailAddress"
        Me.txtEmailAddress.Size = New System.Drawing.Size(200, 20)
        Me.txtEmailAddress.TabIndex = 60
        Me.txtEmailAddress.Text = "Accounts@medscreen.com"
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(64, 216)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(48, 16)
        Me.lblEmail.TabIndex = 59
        Me.lblEmail.Text = "Email :"
        '
        'txtHolder
        '
        Me.txtHolder.Location = New System.Drawing.Point(120, 184)
        Me.txtHolder.Name = "txtHolder"
        Me.txtHolder.Size = New System.Drawing.Size(200, 20)
        Me.txtHolder.TabIndex = 58
        Me.txtHolder.Text = ""
        '
        'LblHolder
        '
        Me.LblHolder.Location = New System.Drawing.Point(64, 184)
        Me.LblHolder.Name = "LblHolder"
        Me.LblHolder.Size = New System.Drawing.Size(48, 16)
        Me.LblHolder.TabIndex = 57
        Me.LblHolder.Text = "Holder"
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(32, 8)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(88, 23)
        Me.label1.TabIndex = 61
        Me.label1.Text = "Cm Job Number"
        '
        'lblCmNumber
        '
        Me.lblCmNumber.Location = New System.Drawing.Point(128, 8)
        Me.lblCmNumber.Name = "lblCmNumber"
        Me.lblCmNumber.Size = New System.Drawing.Size(136, 23)
        Me.lblCmNumber.TabIndex = 62
        Me.lblCmNumber.Text = "Label2"
        '
        'frmCardPay2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 381)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblCmNumber, Me.label1, Me.txtEmailAddress, Me.lblEmail, Me.txtHolder, Me.LblHolder, Me.txtAuthCode, Me.lblAuthCode, Me.txtGrossPrice, Me.Label9, Me.txtReference, Me.Label6, Me.txtCreditCard, Me.lblCreditCard, Me.Panel1})
        Me.Name = "frmCardPay2"
        Me.Text = "frmCardPay2"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private myCmNumber As String
    Private objJob As Intranet.intranet.jobs.CMJob

    Public Property CMNumber() As String
        Get
            Return myCmNumber
        End Get
        Set(ByVal Value As String)
            myCmNumber = Value
            If myCmNumber.Trim.Length > 0 Then
                objJob = New Intranet.intranet.jobs.CMJob(myCmNumber)
            End If
            Me.txtCreditCard.Text = "XXXX-XXXX-XXXX-"
            Me.txtCreditCard.SelectionStart = 15
            Me.txtCreditCard.SelectionLength = 1
            lblCmNumber.Text = myCmNumber
            If Not objJob Is Nothing AndAlso Not objJob.Invoice Is Nothing Then
                MsgBox("This collection has already been invoiced: - " & objJob.Invoiceid, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            End If
        End Set
    End Property

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        'do validation checks
        Me.ErrorProvider1.SetError(Me.txtAuthCode, "")
        Me.ErrorProvider1.SetError(Me.txtCreditCard, "")
        Me.ErrorProvider1.SetError(Me.txtGrossPrice, "")
        Me.ErrorProvider1.SetError(Me.txtReference, "")
        Me.ErrorProvider1.SetError(Me.txtHolder, "")
        Me.DialogResult = Windows.Forms.DialogResult.None
        Dim strCreditCard As String = txtCreditCard.Text.Replace("-", "")
        Dim isValid As Boolean = True
        If Me.txtAuthCode.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtAuthCode, "The card company's transaction authorisation code must be provided")
            isValid = False
        End If

        If Me.txtGrossPrice.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtGrossPrice, "The Gross cost of the transaction must be entered.")
            isValid = False
        End If

        If Me.txtReference.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtReference, "The Medscreen Reference must be provided.")
            isValid = False
        End If

        If Me.txtHolder.Text.Trim.Length = 0 Then
            Me.ErrorProvider1.SetError(Me.txtHolder, "The Card Owner must be provided.")
            isValid = False
        End If


        If strCreditCard.Trim.Length < 16 Then
            Me.ErrorProvider1.SetError(Me.txtCreditCard, "The card number must be entered, and be 16 characters long, it can be secured by using XXXXs instead of the first 16 numbers.")
            isValid = False
        End If
        If Not isValid Then Exit Sub
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Dim ocmd As New OleDb.OleDbCommand("Lib_collection.SetCardDetails", MedscreenLib.CConnection.DbConnection)
        Dim intRet As Integer
        Try
            ocmd.CommandType = CommandType.StoredProcedure
            ocmd.Parameters.Add(CConnection.StringParameter("cmjobnumber", myCmNumber, 10))
            ocmd.Parameters.Add(CConnection.StringParameter("Cnumber", strCreditCard, 30))
            ocmd.Parameters.Add(CConnection.StringParameter("ccauth", Me.txtAuthCode.Text, 10))
            ocmd.Parameters.Add(CConnection.DateParameter("ccexp", DateField.ZeroDate))
            ocmd.Parameters.Add(CConnection.StringParameter("cardholder", Me.txtHolder.Text, 20))
            ocmd.Parameters.Add(CConnection.StringParameter("paytype", "1", 10))
            ocmd.Parameters.Add(CConnection.StringParameter("refnc", Me.txtReference.Text, 20))
            If CConnection.ConnOpen Then
                intRet = ocmd.ExecuteNonQuery
            End If
            objJob.Load()
            'now see if we can do the accounts get 

            'Get a collection Invoice Object 
            Dim AccInterface As MedscreenLib.AccountsInterface = New MedscreenLib.AccountsInterface()
            AccInterface.Status = "R"
            AccInterface.UserId = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            AccInterface.Reference = objJob.ID
            'At this stage we need to know whether we are dealing with nett or gross transactions
            Dim strSendMethod As String = CConnection.PackageStringList("lib_cctool.GetConfigItem", "CM_PAY_VALUE")
            If strSendMethod = "NET" Then
                AccInterface.Value = 0
            Else
                AccInterface.Value = Me.txtGrossPrice.Text
            End If
            AccInterface.RequestCode = "RCPY"
            AccInterface.SetSendDateNull()

            AccInterface.SendAddress = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
            If Me.txtEmailAddress.Text.Trim.Length > 0 Then AccInterface.SendAddress += ";" & Me.txtEmailAddress.Text()

            AccInterface.Update()
            'cTemp.Update()                  'Start invoicing process
            'if collection is on hold we need to take off hold
            If objJob.Status = MedscreenLib.Constants.GCST_JobStatusOnHold Then
                'We need to take off of hold
                If objJob.CollOfficer.Trim.Length > 0 Then 'Collecting officer assigned so must be sent
                    If objJob.JobName.Trim.Length > 0 Then
                        objJob.ConfirmCollection()
                        objJob.SetStatus(MedscreenLib.Constants.GCST_JobStatusConfirmed)
                    Else
                        objJob.SetStatus(MedscreenLib.Constants.GCST_JobStatusSent)
                    End If
                Else
                    If objJob.JobName.Trim.Length > 0 Then
                        objJob.ConfirmCollection()
                        objJob.SetStatus(MedscreenLib.Constants.GCST_JobStatusConfirmed)
                    Else
                        objJob.SetStatus(MedscreenLib.Constants.GCST_JobStatusCreated)
                    End If
                End If
            End If

            objJob.Update() '  Update collection to show new status 
            Dim intMaxRepeat As Integer = 60
            While (AccInterface.Status = "R") AndAlso (intMaxRepeat > 0)
                AccInterface.Load()
                If AccInterface.Status = "C" Then
                    MsgBox("Invoice number " & AccInterface.Reference & " created")
                    Medscreen.LogAction("Invoice number " & AccInterface.Reference & " created - Collection : " & Me.objJob.ID)
                ElseIf AccInterface.Status = "F" Then
                    MsgBox("Transaction failed : - " & AccInterface.Message)
                    LogAction("Transaction failed : - " & AccInterface.Message & " - Collection : " & Me.objJob.ID)
                End If
                intMaxRepeat -= 1
                System.Threading.Thread.CurrentThread.Sleep(1000)
            End While
        Catch ex As Exception
        Finally

        End Try

    End Sub
End Class
