'Imports oneStopClasses.CreditCards
Imports System.Windows.Forms
Imports System.Drawing

Imports Intranet.invoicing.Creditcards
<ToolboxBitmap("C:\Documents and Settings\taylor\My Documents\Visual Studio Projects\MedscreenCommonGUI\Branch1\CCPic.bmp")> _
Public Class CreditCard
    Inherits System.Windows.Forms.UserControl

    Public Event NewCCard()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Dim tmpDate As Date
        Dim I As Integer
        Dim assemble As System.Reflection.Assembly
        Dim assVer As System.Reflection.AssemblyVersionAttribute

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        tmpDate = now
        For I = 1 To 10
            Me.cbYear.Items.Add(tmpDate.ToString("yyyy"))
            tmpDate = tmpDate.AddYears(1)
        Next

        Me.lblVersion.Text = "1.0.0.3"

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents txtPostcode As System.Windows.Forms.TextBox
    Friend WithEvents txtHouseNo As System.Windows.Forms.TextBox
    Friend WithEvents txtSecCode As System.Windows.Forms.TextBox
    Friend WithEvents lblSecCode As System.Windows.Forms.Label
    Friend WithEvents txtIssue As System.Windows.Forms.TextBox
    Friend WithEvents lblIssue As System.Windows.Forms.Label
    Friend WithEvents lblCardType As System.Windows.Forms.Label
    Friend WithEvents cbYear As System.Windows.Forms.ComboBox
    Friend WithEvents cbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents lblCCEXpiry As System.Windows.Forms.Label
    Friend WithEvents txtCreditCard As System.Windows.Forms.TextBox
    Friend WithEvents lblCreditCard As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblTransaction As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents cmdNewCard As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cbClientCreditCards As System.Windows.Forms.ComboBox
    Friend WithEvents txtHolder As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthCode As System.Windows.Forms.TextBox
    Friend WithEvents lblAuthCode As System.Windows.Forms.Label
    Friend WithEvents txtEmailAddress As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents LblHolder As System.Windows.Forms.Label
    Friend WithEvents lblPost As System.Windows.Forms.Label
    Friend WithEvents lblHouse As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CreditCard))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.txtEmailAddress = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtAuthCode = New System.Windows.Forms.TextBox()
        Me.lblAuthCode = New System.Windows.Forms.Label()
        Me.txtHolder = New System.Windows.Forms.TextBox()
        Me.LblHolder = New System.Windows.Forms.Label()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdNewCard = New System.Windows.Forms.Button()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.lblTransaction = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtPostcode = New System.Windows.Forms.TextBox()
        Me.txtHouseNo = New System.Windows.Forms.TextBox()
        Me.lblPost = New System.Windows.Forms.Label()
        Me.lblHouse = New System.Windows.Forms.Label()
        Me.txtSecCode = New System.Windows.Forms.TextBox()
        Me.lblSecCode = New System.Windows.Forms.Label()
        Me.txtIssue = New System.Windows.Forms.TextBox()
        Me.lblIssue = New System.Windows.Forms.Label()
        Me.lblCardType = New System.Windows.Forms.Label()
        Me.cbYear = New System.Windows.Forms.ComboBox()
        Me.cbMonth = New System.Windows.Forms.ComboBox()
        Me.lblCCEXpiry = New System.Windows.Forms.Label()
        Me.txtCreditCard = New System.Windows.Forms.TextBox()
        Me.lblCreditCard = New System.Windows.Forms.Label()
        Me.cbClientCreditCards = New System.Windows.Forms.ComboBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblVersion, Me.txtEmailAddress, Me.lblEmail, Me.txtAuthCode, Me.lblAuthCode, Me.txtHolder, Me.LblHolder, Me.cmdSave, Me.cmdNewCard, Me.txtAmount, Me.lblTransaction, Me.Button1, Me.txtPostcode, Me.txtHouseNo, Me.lblPost, Me.lblHouse, Me.txtSecCode, Me.lblSecCode, Me.txtIssue, Me.lblIssue, Me.lblCardType, Me.cbYear, Me.cbMonth, Me.lblCCEXpiry, Me.txtCreditCard, Me.lblCreditCard, Me.cbClientCreditCards, Me.lblStatus})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 120)
        Me.Panel1.TabIndex = 0
        '
        'lblVersion
        '
        Me.lblVersion.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.lblVersion.Location = New System.Drawing.Point(672, 96)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.TabIndex = 57
        '
        'txtEmailAddress
        '
        Me.txtEmailAddress.Location = New System.Drawing.Point(72, 48)
        Me.txtEmailAddress.Name = "txtEmailAddress"
        Me.txtEmailAddress.Size = New System.Drawing.Size(200, 20)
        Me.txtEmailAddress.TabIndex = 56
        Me.txtEmailAddress.Text = "Accounts@medscreen.com"
        Me.ToolTip1.SetToolTip(Me.txtEmailAddress, "Email address to send invoice to (for fax enter fax number @star.co.uk e.g. 02077" & _
        "128046@starfax.co.uk)")
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(15, 56)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(48, 16)
        Me.lblEmail.TabIndex = 55
        Me.lblEmail.Text = "Email :"
        '
        'txtAuthCode
        '
        Me.txtAuthCode.Location = New System.Drawing.Point(608, 24)
        Me.txtAuthCode.Name = "txtAuthCode"
        Me.txtAuthCode.Size = New System.Drawing.Size(48, 20)
        Me.txtAuthCode.TabIndex = 54
        Me.txtAuthCode.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtAuthCode, "Authorisation code from Commidea")
        '
        'lblAuthCode
        '
        Me.lblAuthCode.Location = New System.Drawing.Point(536, 24)
        Me.lblAuthCode.Name = "lblAuthCode"
        Me.lblAuthCode.Size = New System.Drawing.Size(72, 24)
        Me.lblAuthCode.TabIndex = 53
        Me.lblAuthCode.Text = "Auth. Code"
        '
        'txtHolder
        '
        Me.txtHolder.Location = New System.Drawing.Point(72, 24)
        Me.txtHolder.Name = "txtHolder"
        Me.txtHolder.Size = New System.Drawing.Size(200, 20)
        Me.txtHolder.TabIndex = 52
        Me.txtHolder.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtHolder, "Name of card holder")
        '
        'LblHolder
        '
        Me.LblHolder.Location = New System.Drawing.Point(16, 24)
        Me.LblHolder.Name = "LblHolder"
        Me.LblHolder.Size = New System.Drawing.Size(48, 16)
        Me.LblHolder.TabIndex = 51
        Me.LblHolder.Text = "Holder"
        '
        'cmdSave
        '
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Bitmap)
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(568, 56)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(88, 23)
        Me.cmdSave.TabIndex = 49
        Me.cmdSave.Text = "Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSave, "Save and verify the credit card details")
        '
        'cmdNewCard
        '
        Me.cmdNewCard.Image = CType(resources.GetObject("cmdNewCard.Image"), System.Drawing.Bitmap)
        Me.cmdNewCard.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewCard.Location = New System.Drawing.Point(680, 56)
        Me.cmdNewCard.Name = "cmdNewCard"
        Me.cmdNewCard.Size = New System.Drawing.Size(88, 23)
        Me.cmdNewCard.TabIndex = 48
        Me.cmdNewCard.Text = "New Card"
        Me.cmdNewCard.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdNewCard, "Create new card")
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(664, 0)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.TabIndex = 47
        Me.txtAmount.Text = ""
        '
        'lblTransaction
        '
        Me.lblTransaction.Location = New System.Drawing.Point(544, 0)
        Me.lblTransaction.Name = "lblTransaction"
        Me.lblTransaction.Size = New System.Drawing.Size(112, 16)
        Me.lblTransaction.TabIndex = 46
        Me.lblTransaction.Text = "Transaction Amount"
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(680, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 44
        Me.Button1.Text = " Bill"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Button1, "Attempt online billing of charge")
        '
        'txtPostcode
        '
        Me.txtPostcode.Location = New System.Drawing.Point(482, 49)
        Me.txtPostcode.Name = "txtPostcode"
        Me.txtPostcode.Size = New System.Drawing.Size(48, 20)
        Me.txtPostcode.TabIndex = 43
        Me.txtPostcode.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtPostcode, "Post code of card holder's registered address")
        '
        'txtHouseNo
        '
        Me.txtHouseNo.Location = New System.Drawing.Point(337, 49)
        Me.txtHouseNo.Name = "txtHouseNo"
        Me.txtHouseNo.Size = New System.Drawing.Size(48, 20)
        Me.txtHouseNo.TabIndex = 42
        Me.txtHouseNo.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtHouseNo, "House no or name")
        '
        'lblPost
        '
        Me.lblPost.Location = New System.Drawing.Point(410, 49)
        Me.lblPost.Name = "lblPost"
        Me.lblPost.Size = New System.Drawing.Size(64, 24)
        Me.lblPost.TabIndex = 41
        Me.lblPost.Text = "Post Code"
        '
        'lblHouse
        '
        Me.lblHouse.Location = New System.Drawing.Point(274, 49)
        Me.lblHouse.Name = "lblHouse"
        Me.lblHouse.Size = New System.Drawing.Size(56, 24)
        Me.lblHouse.TabIndex = 40
        Me.lblHouse.Text = "House No"
        '
        'txtSecCode
        '
        Me.txtSecCode.Location = New System.Drawing.Point(482, 25)
        Me.txtSecCode.Name = "txtSecCode"
        Me.txtSecCode.Size = New System.Drawing.Size(48, 20)
        Me.txtSecCode.TabIndex = 39
        Me.txtSecCode.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSecCode, "3 or 4 digit code usually on reverse of card")
        '
        'lblSecCode
        '
        Me.lblSecCode.Location = New System.Drawing.Point(410, 25)
        Me.lblSecCode.Name = "lblSecCode"
        Me.lblSecCode.Size = New System.Drawing.Size(72, 24)
        Me.lblSecCode.TabIndex = 38
        Me.lblSecCode.Text = "Sec. Code"
        '
        'txtIssue
        '
        Me.txtIssue.Location = New System.Drawing.Point(338, 25)
        Me.txtIssue.Name = "txtIssue"
        Me.txtIssue.Size = New System.Drawing.Size(48, 20)
        Me.txtIssue.TabIndex = 37
        Me.txtIssue.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtIssue, "Applicable to Switch cards only")
        '
        'lblIssue
        '
        Me.lblIssue.Location = New System.Drawing.Point(274, 25)
        Me.lblIssue.Name = "lblIssue"
        Me.lblIssue.Size = New System.Drawing.Size(56, 24)
        Me.lblIssue.TabIndex = 36
        Me.lblIssue.Text = "Issue No"
        '
        'lblCardType
        '
        Me.lblCardType.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblCardType.Location = New System.Drawing.Point(16, 74)
        Me.lblCardType.Name = "lblCardType"
        Me.lblCardType.Size = New System.Drawing.Size(246, 16)
        Me.lblCardType.TabIndex = 35
        '
        'cbYear
        '
        Me.cbYear.Location = New System.Drawing.Point(432, 0)
        Me.cbYear.Name = "cbYear"
        Me.cbYear.Size = New System.Drawing.Size(96, 21)
        Me.cbYear.TabIndex = 34
        Me.cbYear.Text = "1900"
        '
        'cbMonth
        '
        Me.cbMonth.Items.AddRange(New Object() {"01 Jan", "02 Feb", "03 Mar", "04 Apr", "05 May", "06 Jun", "07 Jul", "08 Aug", "09 Sep", "10 Oct", "11 Nov", "12 Dec"})
        Me.cbMonth.Location = New System.Drawing.Point(370, 0)
        Me.cbMonth.Name = "cbMonth"
        Me.cbMonth.Size = New System.Drawing.Size(48, 21)
        Me.cbMonth.TabIndex = 33
        Me.cbMonth.Text = "01 Jan"
        '
        'lblCCEXpiry
        '
        Me.lblCCEXpiry.Location = New System.Drawing.Point(274, 0)
        Me.lblCCEXpiry.Name = "lblCCEXpiry"
        Me.lblCCEXpiry.TabIndex = 32
        Me.lblCCEXpiry.Text = "Credit Card Expiry"
        '
        'txtCreditCard
        '
        Me.txtCreditCard.Location = New System.Drawing.Point(106, 0)
        Me.txtCreditCard.Name = "txtCreditCard"
        Me.txtCreditCard.Size = New System.Drawing.Size(150, 20)
        Me.txtCreditCard.TabIndex = 31
        Me.txtCreditCard.Text = ""
        '
        'lblCreditCard
        '
        Me.lblCreditCard.Location = New System.Drawing.Point(10, 0)
        Me.lblCreditCard.Name = "lblCreditCard"
        Me.lblCreditCard.TabIndex = 30
        Me.lblCreditCard.Text = "Credit Card No"
        '
        'cbClientCreditCards
        '
        Me.cbClientCreditCards.Location = New System.Drawing.Point(104, 0)
        Me.cbClientCreditCards.Name = "cbClientCreditCards"
        Me.cbClientCreditCards.Size = New System.Drawing.Size(168, 21)
        Me.cbClientCreditCards.TabIndex = 50
        Me.cbClientCreditCards.Text = "ComboBox1"
        '
        'lblStatus
        '
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblStatus.Location = New System.Drawing.Point(0, 97)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(552, 16)
        Me.lblStatus.TabIndex = 45
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 500
        Me.ToolTip1.ReshowDelay = 100
        Me.ToolTip1.ShowAlways = True
        '
        'CreditCard
        '
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Name = "CreditCard"
        Me.Size = New System.Drawing.Size(784, 120)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private myCard As Intranet.invoicing.Creditcards.CreditCard
    Private myClient As Intranet.intranet.customerns.Client
    'Private MyTxn As New eftdll2.Transaction()
    Private strCardholder As String
    Private strReference As String

    Public Enum CcPaymentTypes
        CCCardPay = 1
        CCChequePay = 2
        CCCashPay = 3
    End Enum

    Private ccPayMethod As CcPaymentTypes


    Public Property PaymentType() As CcPaymentTypes
        Get
            Return ccPayMethod
        End Get
        Set(ByVal Value As CcPaymentTypes)
            ccPayMethod = Value
            Select Case ccPayMethod
                Case CcPaymentTypes.CCCardPay
                    Me.cbYear.Show()
                    Me.cbMonth.Show()
                    Me.cmdNewCard.Show()
                    Me.cbClientCreditCards.Show()
                    Me.txtAuthCode.Enabled = True
                    Me.txtAuthCode.Text = ""
                    Me.cmdSave.Show()
                    Me.Button1.Show()
                    Me.txtAuthCode.Show()
                    Me.lblAuthCode.Show()
                    Me.lblHouse.Show()
                    Me.lblIssue.Show()
                    Me.lblPost.Show()
                    Me.lblSecCode.Show()
                    Me.lblAuthCode.Text = "Auth. Code"
                    Me.lblCreditCard.Text = "Credit Card No"
                    Me.lblCCEXpiry.Show()
                    Me.txtHouseNo.Show()
                    Me.txtIssue.Show()
                    Me.txtSecCode.Show()
                    Me.txtPostcode.Show()
                Case Else
                    Me.cbYear.Hide()
                    Me.cbMonth.Hide()
                    Me.cmdNewCard.Hide()
                    Me.cmdSave.Hide()
                    Me.cbClientCreditCards.Hide()
                    Me.Button1.Hide()
                    Me.lblHouse.Hide()
                    Me.lblIssue.Hide()
                    Me.lblPost.Hide()
                    Me.lblAuthCode.Text = "Payment by"
                    Me.lblSecCode.Hide()
                    Me.txtHouseNo.Hide()
                    Me.txtIssue.Hide()
                    Me.txtSecCode.Hide()
                    Me.txtPostcode.Hide()
                    Me.lblCreditCard.Text = "Cheque R No."
                    Me.lblCCEXpiry.Hide()

                    If ccPayMethod = CcPaymentTypes.CCCashPay Then
                        Me.txtAuthCode.Text = "CASH"
                    Else
                        Me.txtAuthCode.Text = "CHEQUE"
                    End If
                    Me.txtAuthCode.Enabled = False

            End Select
        End Set
    End Property


    'Public WithEvents SSLTest As eftdll2.SSL



    Private blnFirst As Boolean = True
    Private blnBilled As Boolean = False
    Private blnNewCard = False
    Private dblPrice As Double

    Public ReadOnly Property Billcommand() As Button
        Get
            Return Me.Button1
        End Get
    End Property

    Public ReadOnly Property CardHolder() As String
        Get
            strCardholder = Me.txtHolder.Text
            Return strCardholder
        End Get
    End Property



    Public ReadOnly Property Billed() As Boolean
        Get
            Return blnBilled
        End Get
    End Property

    Public Property Reference() As String
        Get
            Return strReference
        End Get
        Set(ByVal Value As String)
            strReference = "CM" & Value
            Me.lblStatus.Text = Value
        End Set
    End Property

    Public Property AuthorisationCode() As String
        Get
            'If MyTxn Is Nothing Or Me.Button1.Visible = False Then
            Return Me.txtAuthCode.Text
            'Else
            'Return MyTxn.AuthCode
            'End If
        End Get
        Set(ByVal Value As String)
            Me.txtAuthCode.Text = Value
        End Set
    End Property

    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            Dim cCard As Intranet.intranet.Creditcard.ClientCard

            myClient = Value
            If myClient Is Nothing Then
                Me.cbClientCreditCards.Hide()
            Else
                Me.cbClientCreditCards.Show()
                cbClientCreditCards.Items.Clear()
                For Each cCard In myClient.CreditCards
                    cbClientCreditCards.Items.Add(cCard.CardNo)
                Next
                cbClientCreditCards.Refresh()
                If myClient.CreditCards.Count > 0 Then
                    cbClientCreditCards.SelectedIndex = 0
                End If

            End If
        End Set
    End Property

    Public Property Price() As Double
        Get
            Return dblPrice
        End Get
        Set(ByVal Value As Double)
            Dim dfs As New System.Globalization.CultureInfo("en-GB")
            dblPrice = Value
            dfs.NumberFormat.NumberDecimalDigits = 2
            Me.txtAmount.Text = dblPrice.ToString("F", dfs)
        End Set
    End Property



    Public Property Card() As Intranet.invoicing.Creditcards.CreditCard
        Get
            Return myCard
        End Get
        Set(ByVal Value As Intranet.invoicing.Creditcards.CreditCard)
            myCard = Value
            If Not myCard Is Nothing Then
                'If Not myCard.CardInstance Is Nothing Then
                FillControl()
                'End If
            Else
                EnableControl(False)
            End If
        End Set
    End Property

    Private Sub EnableControl(ByVal state As Boolean)
        Me.txtCreditCard.Enabled = True
        Me.lblCardType.Enabled = state
        Me.txtIssue.Enabled = state
        Me.txtSecCode.Enabled = state
        Me.txtHouseNo.Enabled = state
        Me.txtPostcode.Enabled = state


        Me.cbMonth.Enabled = True
        Me.cbYear.Enabled = True
        Me.Button1.Enabled = state

    End Sub

    Public Sub FillControl()
        Dim datTemp As Date
        ErrorProvider1.SetError(Me.txtCreditCard, "")
        ErrorProvider1.SetError(Me.cbYear, "")
        ErrorProvider1.SetError(Me.cbMonth, "")
        ErrorProvider1.SetError(Me.txtEmailAddress, "")
        Me.txtHolder.Text = ""
        If Not myCard Is Nothing Then
            Me.txtCreditCard.Text = myCard.ToString
            Me.lblCardType.Show()
            Me.lblCardType.Text = "Card Type : " & myCard.CardType
            Me.txtIssue.Text = myCard.IssueNumber
            Me.txtSecCode.Text = myCard.SecCode
            Me.txtHouseNo.Text = myCard.HouseNo
            Me.txtPostcode.Text = myCard.PostCode
            Me.txtEmailAddress.Text = myCard.EmailAddress
            Me.txtHolder.Text = myCard.Holder


            datTemp = myCard.ExpiryDate
            Me.cbMonth.SelectedIndex = datTemp.Month - 1
            If Date.Compare(datTemp, Now) >= 0 Then
                Me.cbMonth.ForeColor = System.Drawing.SystemColors.ControlText
                Me.cbYear.ForeColor = System.Drawing.SystemColors.ControlText
                Me.cbYear.SelectedIndex = datTemp.Year - Now.Year
            Else
                ErrorProvider1.SetError(Me.cbYear, "Invalid Date")
                Me.cbMonth.ForeColor = System.Drawing.Color.Red
                Me.cbYear.ForeColor = System.Drawing.Color.Red
            End If
            Me.txtCreditCard.Text = myCard.ToString
        Else
            Me.lblCardType.Hide()
        End If
        EnableControl(True)
        Me.txtCreditCard.Focus()

    End Sub

    Public Overloads Function ValidateCard(ByVal intOpt As Integer) As Boolean
        ErrorProvider1.SetError(Me.txtCreditCard, "")
        ErrorProvider1.SetError(Me.cbYear, "")
        ErrorProvider1.SetError(Me.cbMonth, "")
        ErrorProvider1.SetError(Me.txtEmailAddress, "")
        ErrorProvider1.SetError(Me.txtHolder, "")
        If intOpt = 1 Then
            If Me.txtCreditCard.Text.Length = 0 Then
                Throw New CCNumberInvalidException()
                Return False
            Else
            End If

            If Me.txtHolder.Text.Length = 0 Then
                Throw New CCMissingCardHolder()
                Return False
            Else
            End If
        End If

        If intOpt = 2 Then
            If Me.cbYear.Text < Year(Now) Then
                Throw New CCExpiredException()
                Return False
            End If
        End If
        Return True

    End Function

    Public Overloads Function ValidateCard(ByVal blnNumberOnly As Boolean) As Boolean
        ErrorProvider1.SetError(Me.txtCreditCard, "")
        ErrorProvider1.SetError(Me.cbYear, "")
        ErrorProvider1.SetError(Me.cbMonth, "")
        ErrorProvider1.SetError(Me.txtEmailAddress, "")
        ErrorProvider1.SetError(Me.txtHolder, "")
        If Me.txtCreditCard.Text.Length = 0 Then
            Throw New CCNumberInvalidException()
            Return False
        Else
        End If


        Return True

    End Function


    Public Function GetControl()
        ErrorProvider1.SetError(Me.txtCreditCard, "")
        ErrorProvider1.SetError(Me.cbYear, "")
        ErrorProvider1.SetError(Me.cbMonth, "")
        ErrorProvider1.SetError(Me.txtPostcode, "")
        ErrorProvider1.SetError(Me.txtHolder, "")
        If Not myCard Is Nothing Then
            Try
                myCard.Number = txtCreditCard.Text
                strCardholder = Me.txtHolder.Text
                Me.lblStatus.Text = "Validating email address"
                myCard.EmailAddress = Me.txtEmailAddress.Text
                Me.lblStatus.Text = ""
                myCard.Holder = Me.txtHolder.Text


            Catch ex As CCNumberInvalidException
                'Me.ErrorProvider1.SetError(Me.txtCreditCard, ex.Message)
            Catch ex As CCInvalidEmailAddress
                Me.ErrorProvider1.SetError(Me.txtEmailAddress, ex.Message)
            Catch ex As CCMissingCardHolder
                Me.ErrorProvider1.SetError(Me.txtHolder, ex.Message)
            Catch ex As CCExpiredException
                'Me.ErrorProvider1.SetError(Me.cbYear, ex.Message)
            Catch ex As Exception

            End Try
            myCard.ExpiryYear = Now.Year + Me.cbYear.SelectedIndex
            myCard.ExpiryMonth = Me.cbMonth.SelectedIndex + 1
            myCard.IssueNumber = Me.txtIssue.Text
            myCard.SecCode = Me.txtSecCode.Text
            myCard.HouseNo = Me.txtHouseNo.Text
            myCard.PostCode = Me.txtPostcode.Text
        End If

    End Function
    Private Sub txtCreditCard_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCreditCard.Validating
        Try
            Me.ValidateCard(1)
            If blnNewCard Then
                If IsValidated(1) Then
                    Me.CreditCard1_NewCCard()
                    blnNewCard = False
                End If
            End If
        Catch ex As CCNumberInvalidException
            ' Cancel the event and select the text to be corrected by the user.
            e.Cancel = True
            txtCreditCard.Select(0, txtCreditCard.Text.Length)

            ' Set the ErrorProvider error with the text to display. 
            Me.ErrorProvider1.SetError(txtCreditCard, ex.Message)
        Catch ex As CCExpiredException
            e.Cancel = False
            Me.cbYear.Select(0, cbYear.Text.Length)
            Me.ErrorProvider1.SetError(cbYear, ex.Message)
        Catch ex As CCInvalidEmailAddress
            e.Cancel = False
            Me.ErrorProvider1.SetError(Me.txtEmailAddress, ex.Message)
        Catch ex As Exception
            e.Cancel = False
        End Try

    End Sub

    Private Sub txtCreditCard_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCreditCard.Validated
        ErrorProvider1.SetError(txtCreditCard, "")
    End Sub

    Private Sub cbYear_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cbYear.Validating
        If Not myCard Is Nothing Then
            Try
                Me.ValidateCard(2)
                If blnNewCard Then
                    If IsValidated(2) Then
                        Me.CreditCard1_NewCCard()
                    End If
                End If
            Catch ex As CCNumberInvalidException
                ' Cancel the event and select the text to be corrected by the user.
                e.Cancel = True
                txtCreditCard.Select(0, txtCreditCard.Text.Length)

                ' Set the ErrorProvider error with the text to display. 
                Me.ErrorProvider1.SetError(txtCreditCard, ex.Message)
            Catch ex As CCExpiredException
                e.Cancel = True
                Me.cbYear.Select(0, cbYear.Text.Length)
                Me.ErrorProvider1.SetError(cbYear, ex.Message)
            Catch ex As CCMissingCardHolder
                e.Cancel = True
                Me.ErrorProvider1.SetError(Me.txtHolder, ex.Message)
                Me.txtHolder.Focus()

            Catch ex As Exception
                e.Cancel = True
            End Try
        End If
    End Sub


    Private Sub cbYear_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbYear.Validated
        ErrorProvider1.SetError(cbYear, "")
    End Sub

    Private Sub txtCreditCard_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCreditCard.Leave
        If Me.PaymentType = CcPaymentTypes.CCCardPay Then
            If blnNewCard Then
                If IsNumberValidated() Then
                    Me.CreditCard1_NewCCard()
                Else
                    Me.txtCreditCard.Focus()
                End If
            End If
        Else
            Me.Reference = Me.txtCreditCard.Text
        End If
    End Sub

    Private Sub txtCreditCard_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCreditCard.TextChanged
        Try
            'myCard.Number = txtCreditCard.Text
            ErrorProvider1.SetError(txtCreditCard, "")
            Me.lblCardType.Text = "Card Type : " & myCard.CardType
        Catch ex As Exception
            ErrorProvider1.SetError(txtCreditCard, ex.Message)
        End Try
    End Sub

    Public Function IsValidated(ByVal intOpt As Integer) As Boolean
        Dim blnFail As Boolean = False
        Try
            Me.ValidateCard(intOpt)
            Return True
        Catch ex As CCInvalidEmailAddress
            Me.ErrorProvider1.SetError(Me.txtEmailAddress, ex.Message)
            Me.txtEmailAddress.Focus()
        Catch ex As CCNumberInvalidException
            Me.ErrorProvider1.SetError(Me.txtCreditCard, ex.Message)
            Me.txtCreditCard.Focus()
        Catch ex As CCExpiredException
            Me.ErrorProvider1.SetError(Me.cbYear, ex.Message)
            Me.cbYear.Focus()
        Catch ex As CCMissingCardHolder
            Me.ErrorProvider1.SetError(Me.txtHolder, ex.Message)
            Me.txtHolder.Focus()
        Catch ex As Exception
        End Try
        Return blnFail

    End Function

    Public Function IsNumberValidated() As Boolean
        Dim blnFail As Boolean = False
        Try
            Me.ValidateCard(True)
            Return True
        Catch ex As CCInvalidEmailAddress
            Me.ErrorProvider1.SetError(Me.txtEmailAddress, ex.Message)
            Me.txtEmailAddress.Focus()
        Catch ex As CCNumberInvalidException
            Me.ErrorProvider1.SetError(Me.txtCreditCard, ex.Message)
            Me.txtCreditCard.Focus()
        Catch ex As CCExpiredException
            Me.ErrorProvider1.SetError(Me.cbYear, ex.Message)
            Me.cbYear.Focus()
        Catch ex As CCMissingCardHolder
            Me.ErrorProvider1.SetError(Me.txtHolder, ex.Message)
            Me.txtHolder.Focus()
        Catch ex As Exception
        End Try
        Return blnFail

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        blnBilled = False
        blnFirst = True
        Me.lblStatus.Text = ""
        If myCard Is Nothing Then Exit Sub
        'CheckTransaction("0")
    End Sub

    'Private Sub SSLTest_GUIDCheckRecived(ByVal Data As String) Handles SSLTest.GUIDCheckRecived
    '    Me.lblStatus.Text = "Connected...."
    '    If blnFirst Then
    '        DoTransaction()
    '        blnFirst = False
    '    End If
    'End Sub


    'Private Sub DoTransaction()
    '    MyTxn = New eftdll2.Transaction()
    '    MyTxn.AccountNumber = CreditCardClass.cstAccountNo
    '    MyTxn.TxnType = "01"
    '    MyTxn.Modifier = CreditCardClass.cstModifier
    '    MyTxn.AVS = Me.txtHouseNo.Text & ";" & Me.txtPostcode.Text

    '    'MyTxn.CV2 = "000"
    '    MyTxn.PAN = Me.myCard.Number
    '    MyTxn.Expiry = Me.myCard.ExpiryDate.ToString("MMyy")
    '    MyTxn.Issue = Me.myCard.IssueNumber
    '    'MyTxn.Issue = ""
    '    'MyTxn.Start = Me.txtstart.text
    '    MyTxn.TxnValue = Me.txtAmount.Text
    '    MyTxn.Reference = strReference
    '    'MyTxn.IRecord = New CommsDLL.IRecord()
    '    MyTxn.Execute(SSLTest)
    '    If SSLTest.Errors.NumOfErrors > 0 Then
    '        MsgBox(SSLTest.Errors.GetError(1))
    '    End If

    'End Sub


    'Private Sub SSLTest_SSLError(ByVal Data As String) Handles SSLTest.SSLError
    '    lblStatus.Text = Data
    'End Sub


    'Private Sub SSLTest_TransactionRecived(ByVal Data As String) Handles SSLTest.TransactionRecived
    '    lblStatus.Text = MyTxn.Status
    '    blnBilled = (Len(MyTxn.AuthCode) > 0)
    '    If blnBilled Then
    '        Me.txtAuthCode.Text = MyTxn.AuthCode
    '    End If
    'End Sub


    'Private Sub SSLTest_TransactionStatusUpdate(ByVal Data As String) Handles SSLTest.TransactionStatusUpdate
    '    lblStatus.Text = MyTxn.Status
    'End Sub

    'Private Function ConnectToCommidea()
    '    Dim strGuid As String
    '    Try
    '        SSLTest = New eftdll2.SSL()
    '        SSLTest.LogTransactions = True
    '        SSLTest.Connect(CreditCardClass.cstIPAddress, CreditCardClass.cstPortNumber)
    '        strGuid = SSLTest.GetGUID()
    '        If Trim$(CreditCardClass.cstGuidToSend) = "" Then
    '            'cstGuidToSend = SSLTest.GetGUID()
    '            SSLTest.SendString("GUID," & strGuid)
    '        Else
    '            SSLTest.SendString("GUID," & strGuid)

    '        End If
    '    Catch ex As Exception
    '    Finally
    '    End Try

    'End Function

    'Public Function CheckTransaction(ByVal strValue As String)


    '    'ConnectToCommidea()

    'End Function

    Public ReadOnly Property CNumber() As String
        Get
            Return Me.txtCreditCard.Text
        End Get
    End Property

    Public ReadOnly Property ExpiryDate() As Date
        Get
            Return DateSerial(Now.Year + Me.cbYear.SelectedIndex, _
                Me.cbMonth.SelectedIndex + 1, 1)
        End Get
    End Property
    Private Sub cmdNewCard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewCard.Click
        'Me.EnableControl(True)
        Me.txtCreditCard.Text = "--"
        Me.EnableControl(False)
        Me.txtCreditCard.Focus()
        blnNewCard = True

    End Sub

    Private Sub cbYear_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbYear.Leave
        If Not myCard Is Nothing Then
            myCard.ExpiryYear = Now.Year + Me.cbYear.SelectedIndex
        End If
    End Sub

    Private Sub cbMonth_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbMonth.Leave
        If Not myCard Is Nothing Then
            myCard.ExpiryMonth = Me.cbMonth.SelectedIndex + 1
        End If

    End Sub

    Private Sub CreditCard1_NewCCard()
        Dim datExp As Date
        Dim strCnumber As String
        Dim tmpCard As Intranet.invoicing.Creditcards.CreditCard

        strCnumber = Me.CNumber
        If strCnumber = "--" Then Exit Sub
        Try
            datExp = Me.ExpiryDate
            tmpCard = myClient.CreditCards.NewCard(strCnumber, datExp)
            If Not tmpCard Is Nothing Then
                Me.Card = tmpCard
                Me.cbClientCreditCards.Items.Add(strCnumber)
                Me.cbClientCreditCards.SelectedIndex = Me.cbClientCreditCards.Items.Count - 1
            End If
            blnNewCard = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            Me.GetControl()
            Me.Card.Update()
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        End Try
    End Sub

    Private Sub cbClientCreditCards_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbClientCreditCards.SelectedIndexChanged
        Dim strText
        Me.txtCreditCard.Text = Me.cbClientCreditCards.Text
        If myClient.CreditCard Is Nothing Then
            strText = Me.cbClientCreditCards.Text
        Else
            strText = myClient.CreditCard.Trimmed(Me.cbClientCreditCards.Text)
        End If
        If Not myClient Is Nothing Then
            myCard = Nothing
            myCard = myClient.CreditCards.Item(strText)
            If Not myCard Is Nothing Then
                'If Not myCard.CardInstance Is Nothing Then
                FillControl()
                'End If
            Else
                EnableControl(False)
            End If
        Else

        End If

    End Sub

    Private Sub Button1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.VisibleChanged
        'Me.lblAuthCode.Visible = Not Me.Button1.Visible
        'Me.txtAuthCode.Visible = Not Me.Button1.Visible
        Me.txtSecCode.Visible = Me.Button1.Visible
        Me.lblSecCode.Visible = Me.Button1.Visible
    End Sub

    Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
        cmdSave_Click(Nothing, Nothing)
    End Sub

    Private Sub txtEmailAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmailAddress.Enter
        Me.lblStatus.Text = Me.ToolTip1.GetToolTip(Me.txtEmailAddress)
        Me.lblStatus.BringToFront()
        Me.lblStatus.Show()

    End Sub

    Private Sub txtEmailAddress_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmailAddress.Leave
        Try
            Me.lblStatus.Text = "Validating email address"
            Me.lblStatus.BringToFront()
            myCard.EmailAddress = Me.txtEmailAddress.Text
            Me.lblStatus.Text = ""
            Me.lblStatus.SendToBack()

        Catch ex As CCInvalidEmailAddress
            Me.ErrorProvider1.SetError(Me.txtEmailAddress, ex.Message)
        Catch ex As CCNumberInvalidException
            Me.ErrorProvider1.SetError(Me.txtCreditCard, ex.Message)
        Catch ex As CCExpiredException
            Me.ErrorProvider1.SetError(Me.cbYear, ex.Message)
        Catch ex As CCMissingCardHolder
            Me.ErrorProvider1.SetError(Me.txtHolder, ex.Message)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub lblStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblStatus.Click

    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged

    End Sub



End Class
