'$Revision: 1.2 $
'$Author: taylor $
'$Date: 2005-06-16 17:39:47+01 $
'$Log: frmCardPay.vb,v $
'Revision 1.2  2005-06-16 17:39:47+01  taylor
'Milestone Check in
'
'Revision 1.1  2004-05-04 10:11:43+01  taylor
'Header added to form
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Imports MedscreenLib
Imports MedscreenLib.Medscreen
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCardPay
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' form to get payments by card cash or cheque
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCardPay
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
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents CreditCard1 As MedscreenCommonGui.CreditCard
    Friend WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbCollections As System.Windows.Forms.ComboBox
    Friend WithEvents txtNett As System.Windows.Forms.TextBox
    Friend WithEvents txtVAT As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtPO As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents DTCollDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents HtmlEditor1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtGrossPrice As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RBCCard As System.Windows.Forms.RadioButton
    Friend WithEvents RBCheque As System.Windows.Forms.RadioButton
    Friend WithEvents RBCash As System.Windows.Forms.RadioButton
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents lblWarn As System.Windows.Forms.Label
    Friend WithEvents lblClient As System.Windows.Forms.Label
    Friend WithEvents lblCustFull As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents SqlConnection1 As System.Data.SqlClient.SqlConnection
    Friend WithEvents SqlDataAdapter1 As System.Data.SqlClient.SqlDataAdapter
    Friend WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCardPay))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblWarn = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.CreditCard1 = New MedscreenCommonGui.CreditCard()
        Me.label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNett = New System.Windows.Forms.TextBox()
        Me.txtVAT = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbCollections = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtPO = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
        Me.DTCollDate = New System.Windows.Forms.DateTimePicker()
        Me.HtmlEditor1 = New System.Windows.Forms.WebBrowser
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtGrossPrice = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RBCash = New System.Windows.Forms.RadioButton()
        Me.RBCheque = New System.Windows.Forms.RadioButton()
        Me.RBCCard = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.lblClient = New System.Windows.Forms.Label()
        Me.lblCustFull = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SqlConnection1 = New System.Data.SqlClient.SqlConnection()
        Me.SqlDataAdapter1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblWarn, Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 445)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 48)
        Me.Panel1.TabIndex = 0
        '
        'lblWarn
        '
        Me.lblWarn.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarn.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblWarn.Location = New System.Drawing.Point(16, 16)
        Me.lblWarn.Name = "lblWarn"
        Me.lblWarn.Size = New System.Drawing.Size(496, 23)
        Me.lblWarn.TabIndex = 2
        Me.lblWarn.Text = "Payment already taken for collection"
        Me.lblWarn.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(696, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(584, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CreditCard1
        '
        Me.CreditCard1.AuthorisationCode = ""
        Me.CreditCard1.Card = Nothing
        Me.CreditCard1.Client = Nothing
        Me.CreditCard1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.CreditCard1.Location = New System.Drawing.Point(0, 325)
        Me.CreditCard1.Name = "CreditCard1"
        Me.CreditCard1.Price = 0
        Me.CreditCard1.Reference = "CMCMCMCMCMCMCMCMCMCMCMCMCMCMCMCMCM"
        Me.CreditCard1.Size = New System.Drawing.Size(784, 120)
        Me.CreditCard1.TabIndex = 1
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(8, 8)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(88, 23)
        Me.label1.TabIndex = 2
        Me.label1.Text = "Cm Job Number"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(248, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Nett Price"
        '
        'txtNett
        '
        Me.txtNett.Location = New System.Drawing.Point(320, 5)
        Me.txtNett.Name = "txtNett"
        Me.txtNett.Size = New System.Drawing.Size(56, 20)
        Me.txtNett.TabIndex = 5
        Me.txtNett.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtNett, "Nett Amount (Vat Exclusive) value")
        '
        'txtVAT
        '
        Me.txtVAT.Location = New System.Drawing.Point(472, 5)
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(56, 20)
        Me.txtVAT.TabIndex = 7
        Me.txtVAT.Text = "0"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(400, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "VAT"
        '
        'cbCollections
        '
        Me.cbCollections.Location = New System.Drawing.Point(104, 4)
        Me.cbCollections.Name = "cbCollections"
        Me.cbCollections.Size = New System.Drawing.Size(121, 21)
        Me.cbCollections.TabIndex = 8
        Me.cbCollections.Text = "ComboBox1"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 23)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "PO Num"
        '
        'txtPO
        '
        Me.txtPO.Location = New System.Drawing.Point(56, 32)
        Me.txtPO.Name = "txtPO"
        Me.txtPO.Size = New System.Drawing.Size(88, 20)
        Me.txtPO.TabIndex = 10
        Me.txtPO.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtPO, "Customer's Purchase Order number or reference")
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(336, 34)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Coll. Date"
        '
        'DTCollDate
        '
        Me.DTCollDate.Enabled = False
        Me.DTCollDate.Location = New System.Drawing.Point(408, 32)
        Me.DTCollDate.Name = "DTCollDate"
        Me.DTCollDate.Size = New System.Drawing.Size(136, 20)
        Me.DTCollDate.TabIndex = 12
        '
        'HtmlEditor1
        '
        Me.HtmlEditor1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        'Me.HtmlEditor1.DefaultComposeSettings.BackColor = System.Drawing.Color.White
        'Me.HtmlEditor1.DefaultComposeSettings.DefaultFont = New System.Drawing.Font("Arial", 10.0!)
        'Me.HtmlEditor1.DefaultComposeSettings.Enabled = False
        'Me.HtmlEditor1.DefaultComposeSettings.ForeColor = System.Drawing.Color.Black
        'Me.HtmlEditor1.DocumentEncoding = onlyconnect.EncodingType.WindowsCurrent
        'Me.HtmlEditor1.Location = New System.Drawing.Point(8, 112)
        Me.HtmlEditor1.Name = "HtmlEditor1"
        'Me.HtmlEditor1.OpenLinksInNewWindow = True
        'Me.HtmlEditor1.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left
        'Me.HtmlEditor1.SelectionBackColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionBullets = False
        'Me.HtmlEditor1.SelectionFont = Nothing
        'Me.HtmlEditor1.SelectionForeColor = System.Drawing.Color.Empty
        'Me.HtmlEditor1.SelectionNumbering = False
        Me.HtmlEditor1.Size = New System.Drawing.Size(768, 200)
        Me.HtmlEditor1.TabIndex = 13
        ''Me.HtmlEditor1.Text = "HtmlEditor1"
        Me.ToolTip1.SetToolTip(Me.HtmlEditor1, "Job Price Calculation")
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(544, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 23)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Gross Price"
        '
        'txtGrossPrice
        '
        Me.txtGrossPrice.Location = New System.Drawing.Point(624, 4)
        Me.txtGrossPrice.Name = "txtGrossPrice"
        Me.txtGrossPrice.Size = New System.Drawing.Size(72, 20)
        Me.txtGrossPrice.TabIndex = 20
        Me.txtGrossPrice.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtGrossPrice, "Gross Amount (VAT inclusive)")
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.RBCash, Me.RBCheque, Me.RBCCard})
        Me.GroupBox1.Location = New System.Drawing.Point(560, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(224, 40)
        Me.GroupBox1.TabIndex = 21
        Me.GroupBox1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.GroupBox1, "Select payment method")
        '
        'RBCash
        '
        Me.RBCash.Location = New System.Drawing.Point(152, 16)
        Me.RBCash.Name = "RBCash"
        Me.RBCash.Size = New System.Drawing.Size(56, 16)
        Me.RBCash.TabIndex = 2
        Me.RBCash.Text = "C&ash"
        '
        'RBCheque
        '
        Me.RBCheque.Location = New System.Drawing.Point(80, 16)
        Me.RBCheque.Name = "RBCheque"
        Me.RBCheque.Size = New System.Drawing.Size(64, 16)
        Me.RBCheque.TabIndex = 1
        Me.RBCheque.Text = "Che&que"
        '
        'RBCCard
        '
        Me.RBCCard.Checked = True
        Me.RBCCard.Location = New System.Drawing.Point(8, 16)
        Me.RBCCard.Name = "RBCCard"
        Me.RBCCard.Size = New System.Drawing.Size(64, 16)
        Me.RBCCard.TabIndex = 0
        Me.RBCCard.TabStop = True
        Me.RBCCard.Text = "&CCard"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(152, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Reference"
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(216, 32)
        Me.txtReference.Name = "txtReference"
        Me.txtReference.TabIndex = 23
        Me.txtReference.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtReference, "Enter Medscreen reference number here")
        '
        'lblClient
        '
        Me.lblClient.Location = New System.Drawing.Point(24, 64)
        Me.lblClient.Name = "lblClient"
        Me.lblClient.Size = New System.Drawing.Size(144, 23)
        Me.lblClient.TabIndex = 24
        '
        'lblCustFull
        '
        Me.lblCustFull.Location = New System.Drawing.Point(208, 64)
        Me.lblCustFull.Name = "lblCustFull"
        Me.lblCustFull.Size = New System.Drawing.Size(520, 23)
        Me.lblCustFull.TabIndex = 25
        '
        'SqlConnection1
        '
        Me.SqlConnection1.ConnectionString = "data source=andrew\netsdk;initial catalog=ICP;integrated security=SSPI;workstatio" & _
        "n id=TS02DOCK;packet size=4096"
        '
        'SqlDataAdapter1
        '
        Me.SqlDataAdapter1.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlDataAdapter1.InsertCommand = Me.SqlInsertCommand1
        Me.SqlDataAdapter1.SelectCommand = Me.SqlSelectCommand1
        Me.SqlDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Transactions", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TransactionID", "TransactionID"), New System.Data.Common.DataColumnMapping("UserName", "UserName"), New System.Data.Common.DataColumnMapping("TxnType", "TxnType"), New System.Data.Common.DataColumnMapping("SchemeName", "SchemeName"), New System.Data.Common.DataColumnMapping("Modifier", "Modifier"), New System.Data.Common.DataColumnMapping("CardNumber", "CardNumber"), New System.Data.Common.DataColumnMapping("Expiry", "Expiry"), New System.Data.Common.DataColumnMapping("TxnValue", "TxnValue"), New System.Data.Common.DataColumnMapping("AuthCode", "AuthCode"), New System.Data.Common.DataColumnMapping("DateTime", "DateTime"), New System.Data.Common.DataColumnMapping("EFTSeqNum", "EFTSeqNum"), New System.Data.Common.DataColumnMapping("Referance", "Referance"), New System.Data.Common.DataColumnMapping("TxnResult", "TxnResult"), New System.Data.Common.DataColumnMapping("AuthMessage", "AuthMessage"), New System.Data.Common.DataColumnMapping("CardholderName", "CardholderName")})})
        Me.SqlDataAdapter1.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT TransactionID, UserName, TxnType, SchemeName, Modifier, CardNumber, Expiry" & _
        ", TxnValue, AuthCode, DateTime, EFTSeqNum, Referance, TxnResult, AuthMessage, Ca" & _
        "rdholderName FROM Transactions"
        Me.SqlSelectCommand1.Connection = Me.SqlConnection1
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO Transactions(UserName, TxnType, SchemeName, Modifier, CardNumber, Exp" & _
        "iry, TxnValue, AuthCode, DateTime, EFTSeqNum, Referance, TxnResult, AuthMessage," & _
        " CardholderName) VALUES (@UserName, @TxnType, @SchemeName, @Modifier, @CardNumbe" & _
        "r, @Expiry, @TxnValue, @AuthCode, @DateTime, @EFTSeqNum, @Referance, @TxnResult," & _
        " @AuthMessage, @CardholderName); SELECT TransactionID, UserName, TxnType, Scheme" & _
        "Name, Modifier, CardNumber, Expiry, TxnValue, AuthCode, DateTime, EFTSeqNum, Ref" & _
        "erance, TxnResult, AuthMessage, CardholderName FROM Transactions WHERE (Transact" & _
        "ionID = @@IDENTITY)"
        Me.SqlInsertCommand1.Connection = Me.SqlConnection1
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserName", System.Data.SqlDbType.NVarChar, 20, "UserName"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnType", System.Data.SqlDbType.NVarChar, 20, "TxnType"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SchemeName", System.Data.SqlDbType.NVarChar, 20, "SchemeName"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Modifier", System.Data.SqlDbType.NVarChar, 20, "Modifier"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CardNumber", System.Data.SqlDbType.NVarChar, 30, "CardNumber"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Expiry", System.Data.SqlDbType.NVarChar, 20, "Expiry"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnValue", System.Data.SqlDbType.NVarChar, 20, "TxnValue"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AuthCode", System.Data.SqlDbType.NVarChar, 20, "AuthCode"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateTime", System.Data.SqlDbType.NVarChar, 20, "DateTime"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@EFTSeqNum", System.Data.SqlDbType.NVarChar, 20, "EFTSeqNum"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Referance", System.Data.SqlDbType.NVarChar, 100, "Referance"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnResult", System.Data.SqlDbType.NVarChar, 100, "TxnResult"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AuthMessage", System.Data.SqlDbType.NVarChar, 100, "AuthMessage"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CardholderName", System.Data.SqlDbType.NVarChar, 80, "CardholderName"))
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = "UPDATE Transactions SET UserName = @UserName, TxnType = @TxnType, SchemeName = @S" & _
        "chemeName, Modifier = @Modifier, CardNumber = @CardNumber, Expiry = @Expiry, Txn" & _
        "Value = @TxnValue, AuthCode = @AuthCode, DateTime = @DateTime, EFTSeqNum = @EFTS" & _
        "eqNum, Referance = @Referance, TxnResult = @TxnResult, AuthMessage = @AuthMessag" & _
        "e, CardholderName = @CardholderName WHERE (TransactionID = @Original_Transaction" & _
        "ID) AND (AuthCode = @Original_AuthCode OR @Original_AuthCode IS NULL AND AuthCod" & _
        "e IS NULL) AND (AuthMessage = @Original_AuthMessage OR @Original_AuthMessage IS " & _
        "NULL AND AuthMessage IS NULL) AND (CardNumber = @Original_CardNumber OR @Origina" & _
        "l_CardNumber IS NULL AND CardNumber IS NULL) AND (CardholderName = @Original_Car" & _
        "dholderName OR @Original_CardholderName IS NULL AND CardholderName IS NULL) AND " & _
        "(DateTime = @Original_DateTime OR @Original_DateTime IS NULL AND DateTime IS NUL" & _
        "L) AND (EFTSeqNum = @Original_EFTSeqNum OR @Original_EFTSeqNum IS NULL AND EFTSe" & _
        "qNum IS NULL) AND (Expiry = @Original_Expiry OR @Original_Expiry IS NULL AND Exp" & _
        "iry IS NULL) AND (Modifier = @Original_Modifier OR @Original_Modifier IS NULL AN" & _
        "D Modifier IS NULL) AND (Referance = @Original_Referance OR @Original_Referance " & _
        "IS NULL AND Referance IS NULL) AND (SchemeName = @Original_SchemeName OR @Origin" & _
        "al_SchemeName IS NULL AND SchemeName IS NULL) AND (TxnResult = @Original_TxnResu" & _
        "lt OR @Original_TxnResult IS NULL AND TxnResult IS NULL) AND (TxnType = @Origina" & _
        "l_TxnType OR @Original_TxnType IS NULL AND TxnType IS NULL) AND (TxnValue = @Ori" & _
        "ginal_TxnValue OR @Original_TxnValue IS NULL AND TxnValue IS NULL) AND (UserName" & _
        " = @Original_UserName OR @Original_UserName IS NULL AND UserName IS NULL); SELEC" & _
        "T TransactionID, UserName, TxnType, SchemeName, Modifier, CardNumber, Expiry, Tx" & _
        "nValue, AuthCode, DateTime, EFTSeqNum, Referance, TxnResult, AuthMessage, Cardho" & _
        "lderName FROM Transactions WHERE (TransactionID = @TransactionID)"
        Me.SqlUpdateCommand1.Connection = Me.SqlConnection1
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserName", System.Data.SqlDbType.NVarChar, 20, "UserName"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnType", System.Data.SqlDbType.NVarChar, 20, "TxnType"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SchemeName", System.Data.SqlDbType.NVarChar, 20, "SchemeName"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Modifier", System.Data.SqlDbType.NVarChar, 20, "Modifier"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CardNumber", System.Data.SqlDbType.NVarChar, 30, "CardNumber"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Expiry", System.Data.SqlDbType.NVarChar, 20, "Expiry"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnValue", System.Data.SqlDbType.NVarChar, 20, "TxnValue"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AuthCode", System.Data.SqlDbType.NVarChar, 20, "AuthCode"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateTime", System.Data.SqlDbType.NVarChar, 20, "DateTime"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@EFTSeqNum", System.Data.SqlDbType.NVarChar, 20, "EFTSeqNum"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Referance", System.Data.SqlDbType.NVarChar, 100, "Referance"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TxnResult", System.Data.SqlDbType.NVarChar, 100, "TxnResult"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AuthMessage", System.Data.SqlDbType.NVarChar, 100, "AuthMessage"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CardholderName", System.Data.SqlDbType.NVarChar, 80, "CardholderName"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TransactionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TransactionID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_AuthCode", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AuthCode", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_AuthMessage", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AuthMessage", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_CardNumber", System.Data.SqlDbType.NVarChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CardNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_CardholderName", System.Data.SqlDbType.NVarChar, 80, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CardholderName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_DateTime", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateTime", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_EFTSeqNum", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EFTSeqNum", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Expiry", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Expiry", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Modifier", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Modifier", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Referance", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Referance", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_SchemeName", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SchemeName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnResult", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnResult", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnType", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnType", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnValue", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnValue", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_UserName", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TransactionID", System.Data.SqlDbType.Int, 4, "TransactionID"))
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM Transactions WHERE (TransactionID = @Original_TransactionID) AND (Aut" & _
        "hCode = @Original_AuthCode OR @Original_AuthCode IS NULL AND AuthCode IS NULL) A" & _
        "ND (AuthMessage = @Original_AuthMessage OR @Original_AuthMessage IS NULL AND Aut" & _
        "hMessage IS NULL) AND (CardNumber = @Original_CardNumber OR @Original_CardNumber" & _
        " IS NULL AND CardNumber IS NULL) AND (CardholderName = @Original_CardholderName " & _
        "OR @Original_CardholderName IS NULL AND CardholderName IS NULL) AND (DateTime = " & _
        "@Original_DateTime OR @Original_DateTime IS NULL AND DateTime IS NULL) AND (EFTS" & _
        "eqNum = @Original_EFTSeqNum OR @Original_EFTSeqNum IS NULL AND EFTSeqNum IS NULL" & _
        ") AND (Expiry = @Original_Expiry OR @Original_Expiry IS NULL AND Expiry IS NULL)" & _
        " AND (Modifier = @Original_Modifier OR @Original_Modifier IS NULL AND Modifier I" & _
        "S NULL) AND (Referance = @Original_Referance OR @Original_Referance IS NULL AND " & _
        "Referance IS NULL) AND (SchemeName = @Original_SchemeName OR @Original_SchemeNam" & _
        "e IS NULL AND SchemeName IS NULL) AND (TxnResult = @Original_TxnResult OR @Origi" & _
        "nal_TxnResult IS NULL AND TxnResult IS NULL) AND (TxnType = @Original_TxnType OR" & _
        " @Original_TxnType IS NULL AND TxnType IS NULL) AND (TxnValue = @Original_TxnVal" & _
        "ue OR @Original_TxnValue IS NULL AND TxnValue IS NULL) AND (UserName = @Original" & _
        "_UserName OR @Original_UserName IS NULL AND UserName IS NULL)"
        Me.SqlDeleteCommand1.Connection = Me.SqlConnection1
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TransactionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TransactionID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_AuthCode", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AuthCode", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_AuthMessage", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AuthMessage", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_CardNumber", System.Data.SqlDbType.NVarChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CardNumber", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_CardholderName", System.Data.SqlDbType.NVarChar, 80, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CardholderName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_DateTime", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DateTime", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_EFTSeqNum", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EFTSeqNum", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Expiry", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Expiry", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Modifier", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Modifier", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_Referance", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Referance", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_SchemeName", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SchemeName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnResult", System.Data.SqlDbType.NVarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnResult", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnType", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnType", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_TxnValue", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TxnValue", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_UserName", System.Data.SqlDbType.NVarChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserName", System.Data.DataRowVersion.Original, Nothing))
        '
        'frmCardPay
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 493)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblCustFull, Me.lblClient, Me.txtReference, Me.Label6, Me.GroupBox1, Me.txtGrossPrice, Me.Label9, Me.DTCollDate, Me.Label5, Me.txtPO, Me.Label4, Me.cbCollections, Me.txtVAT, Me.Label3, Me.txtNett, Me.Label2, Me.label1, Me.CreditCard1, Me.Panel1, Me.HtmlEditor1})
        Me.Name = "frmCardPay"
        Me.Text = "Record Credit Card Payments"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declarations"



    Private myClient As Intranet.intranet.customerns.Client
    Private MyCollections As Intranet.intranet.jobs.CMCollection
    Private myUser As MedscreenLib.personnel.UserPerson
    Private blnClose As Boolean = False
    Private Collcentres As New Intranet.FixedSiteBookings.cCentres()
    'Private PortsList As New Intranet.intranet.jobs.Ports()
    Private blnJobSelect As Boolean = False
    Private objJob As Intranet.intranet.jobs.CMJob
    'Private myColl As intranet.in

    Private blnInitialising As Boolean = True
    Private blnNoCollections As Boolean = False

    'Interface variables
    Private cTest As New MedscreenLib.IntColl()
    Private cTemp As MedscreenLib.InterfaceClass2
    Private Oisid As Integer = 0
    Private icp As MedscreenLib.ICP
    Private icpR As MedscreenLib.ICPRow

#End Region


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates that it the collections need to be loaded
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property NoCollections() As Boolean
        Get
            Return blnNoCollections
        End Get
        Set(ByVal Value As Boolean)
            blnNoCollections = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Collection ID for payment
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property CollID() As String
        Set(ByVal Value As String)
            Dim i As Integer
            Dim strHtml As String
            Dim Cost As Double
            Dim smClient As Intranet.intranet.customerns.Client
            Dim myPort As Intranet.intranet.jobs.Port
            Dim blnFixedSite As Boolean
            Dim cCharge As Double
            'Dim icp As MedscreenLib.ICP
            'Dim icpR As MedscreenLib.ICPRow
            Dim strCm As String
            Dim intCount As Integer

            Try
                Oisid = 0
                blnClose = True
                objJob = New Intranet.intranet.jobs.CMJob(Value)
                If MyCollections Is Nothing Then
                    MyCollections = New Intranet.intranet.jobs.CMCollection(False)
                    MyCollections.Add(objJob)
                End If
                'objJob = MyCollections.Item(Value)
                Me.lblWarn.Visible = False

                If Not objJob Is Nothing Then
                    If objJob.PrePaid Then
                        Me.lblWarn.Show()
                        Me.txtNett.Text = objJob.Invoice.InvoiceAmount
                    End If
                    'DoContUpdate(objJob)
                    Me.BindingContext(MyCollections).Position = 0
                    For i = 0 To Me.BindingContext(MyCollections).Count - 1
                        If Me.BindingContext(MyCollections).Current.id = Value Then
                            Me.Text = "Record Credit Card Payments for collection " & Value
                            blnClose = False
                            Exit For
                        End If
                        Me.BindingContext(MyCollections).Position = Me.BindingContext(MyCollections).Position + 1
                    Next
                    blnJobSelect = True
                    Me.cbCollections.Enabled = False
                    strCm = objJob.ID
                    Client = objJob.SmClient
                    i = 0
                    While (i < strCm.Length) And strCm.Chars(i) = "0"
                        i += 1
                    End While
                    'Create a new interface record and set status = 'M'
                    Me.cbCollections.Text = objJob.ID

                    'Check to see if we are paying by credit card then go through the process
                    'This is the default setting 
                    If objJob.PaymentType = Constants.PaymentTypes.PaymentCard Then

                        icp = New ICP(objJob.CMID)
                        If Not icp Is Nothing Then
                            If icp.Count = 1 Then
                                icpR = icp.Item(0)
                            End If
                        End If

                        Dim oCI As Intranet.intranet.jobs.CollectionInvoice
                        If icpR Is Nothing Then
                            If objJob.CollectionInvoices.Count > 0 Then oCI = objJob.CollectionInvoices.Item(0)
                            objJob.PaymentType = Constants.PaymentTypes.PaymentCheque
                            icpR = New MedscreenLib.ICPRow()
                            icpR.CMJobNumber = objJob.ID
                        Else
                            oci = objJob.CollectionInvoices.Item(icpR.TransactionId, Intranet.intranet.jobs.CollectionInvoices.IndexBy.TransactionID)
                        End If

                        If oCI Is Nothing Then
                            oCI = objJob.CollectionInvoices.Item(icpR.AuthCode, Intranet.intranet.jobs.CollectionInvoices.IndexBy.Authcode)
                            If Not oCI Is Nothing Then
                                oCI.ICPReference = icpR.TransactionId
                                oCI.Update()
                            Else
                                If Not icpR Is Nothing Then
                                    oCI = objJob.CollectionInvoices.CreateInvoice()
                                    oCI.CCAuthCode = icpR.AuthCode
                                    oCI.CCNumber = icpR.Cardnumber
                                    oCI.MedscreenReference = icpR.MedscreenReference
                                    oCI.ICPReference = icpR.TransactionId
                                    oCI.CMJobNumber = objJob.ID
                                    If icpR.Expiry.Trim.Length > 0 Then oCI.CCExpiry = DateSerial("20" & Mid(icpR.Expiry, 3, 2), Mid(icpR.Expiry, 1, 2), 1)
                                    oCI.Update()

                                End If
                            End If
                        End If
                        If Not oCI Is Nothing Then
                            If icpR.Expiry.Trim.Length = 0 Then
                                SetCard(oCI.CCNumber, oCI.CCExpiry)
                            Else
                                SetCard(oCI.CCNumber, DateSerial(Mid(icpR.Expiry, 3, 2), Mid(icpR.Expiry, 1, 2), 1))
                            End If
                            Me.CreditCard1.Reference = objJob.ID
                            Me.CreditCard1.AuthorisationCode = oCI.CCAuthCode
                            DoContUpdate(objJob)
                            If icpR.TxnValue.Trim.Length = 0 Then
                                If Not objJob.Invoice Is Nothing Then
                                    icpR.TxnValue = objJob.Invoice.InvoiceAmount + objJob.Invoice.InvoiceVat
                                End If
                            End If
                            Me.txtGrossPrice.Text = icpR.TxnValue
                            Me.txtGrossPrice_Leave(Nothing, Nothing)
                            Me.txtReference.Text = oCI.MedscreenReference
                            Exit Property
                        End If
                        'Attemp to get info from ICP database
                        cTemp = cTest.NewInterface(strCm)
                        If Not cTemp Is Nothing Then
                            'Sort out collection details
                            If cTemp.CreditCardNumber.Trim.Length > 0 Then
                                Me.CreditCard1.Reference = objJob.ID
                                Me.CreditCard1.AuthorisationCode = cTemp.CreditAuthorisation

                                DoContUpdate(objJob)
                                Me.Oisid = cTemp.OisID
                                SetCard(cTemp.CreditCardNumber, cTemp.CreditExpiry)
                                Exit Property  ' We have card details bog out 
                            End If
                        End If
                        cTemp = cTest.NewInterface
                        cTemp.CMJobNumber = objJob.ID
                        cTemp.NoDonors = objJob.ExpectedSamples
                        cTemp.Customer = objJob.ClientId
                        cTemp.CollDate = objJob.CollectionDate
                        cTemp.Optional6 = MedscreenLib.Constants.GCST_Collection_Abbr
                        cTemp.Status = MedscreenLib.Constants.GCST_IFace_StatusICPRequest
                        cTemp.Update()
                        intCount = 240
                        Oisid = cTemp.OisID
                        'Wait for response from icp getter
                        While (Not cTest.ReturnFromSM(cTemp, MedscreenLib.Constants.GCST_IFace_StatusFailed)) _
                            And (intCount > 0)
                            System.Threading.Thread.Sleep(250)
                            intCount -= 1
                        End While
                        If intCount > 0 Then 'We have a succesful return 
                            cTemp = Nothing
                            System.Threading.Thread.Sleep(1000)

                            cTemp = cTest.NewInterface(strCm)
                            cTemp.Refresh()

                            'We might still not have a card 
                            If cTemp.PaymentType = Constants.PaymentTypes.PaymentCard Then
                                SetCard(cTemp.CreditCardNumber, cTemp.CreditExpiry)
                                Try
                                    Me.CreditCard1.AuthorisationCode = cTemp.CreditAuthorisation

                                    Me.txtReference.Text = cTemp.MedscreenReference
                                    Me.CreditCard1.Price = cTemp.NettCost
                                Catch ex As Exception
                                    LogError(ex, , "set CollId details : ")
                                End Try
                            Else
                                Me.RBCheque.Checked = True
                            End If
                        End If
                    Else
                        Me.RBCheque.Checked = True
                    End If
                    DoContUpdate(objJob)
                End If
            Catch ex As Exception
                LogError(ex)
            End Try
        End Set
    End Property

    Private Sub SetCard(ByVal CardNo As String, ByVal Expiry As Date)
        Dim myCard As Intranet.invoicing.Creditcards.CreditCard

        Try
            'If cTemp Is Nothing Then
            '    MedscreenLib.Medscreen.LogError("SetCard : Ctemp is nothing")
            '    Exit Sub
            'End If
            If Not myClient Is Nothing Then
                If Not myClient.CreditCards Is Nothing Then
                    If CardNo.Trim.Length = 0 Then
                        If myClient.CreditCards.Count > 0 Then
                            myCard = myClient.CreditCards.Item(0)
                        End If
                    Else
                        myCard = Me.myClient.CreditCards.Item(CardNo)
                    End If
                    If myCard Is Nothing Then
                        myCard = myClient.CreditCards.NewCard(CardNo, Expiry)
                        myCard.ExpiryDate = DateSerial(Mid(icpR.Expiry, 3, 2), Mid(icpR.Expiry, 1, 2), 1)
                        myCard.Update()
                    End If
                    Try
                        If Not myCard Is Nothing Then Me.CreditCard1.Card = myCard
                    Catch ex As Exception
                        MedscreenLib.Medscreen.LogError(ex, True, "det : ")
                    End Try
                End If
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "SetCard : ")
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current user of the program 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property [Operator]() As MedscreenLib.personnel.UserPerson
        Get
            Return myUser
        End Get
        Set(ByVal Value As MedscreenLib.personnel.UserPerson)
            myUser = Value
            blnInitialising = False
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Customer payment is for 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Client() As Intranet.intranet.customerns.Client
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            Dim strWhere As String
            Dim intRet As Windows.Forms.DialogResult
            Try
                Try
                    myClient = Value
                    Me.CreditCard1.Client = Value
                    Me.CreditCard1.AuthorisationCode = ""
                    Me.txtNett.Text = "0"
                    Me.txtVAT.Text = "0"
                    Me.lblClient.Text = myClient.Identity
                    Me.lblCustFull.Text = myClient.CompanyName
                    Me.lblCustFull.ForeColor = System.Drawing.SystemColors.ControlText
                    If Not myClient Is Nothing Then
                        If Not myClient.HasStandardPricing Then
                            intRet = MsgBox("This customer is not on standard pricing! " & vbCrLf & _
                                "This option is normally set for credit card payments. " & vbCrLf & _
                                "Do you want to set it now?", _
                                MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, "Standard Pricing Check")
                            If intRet = DialogResult.Yes Then
                                myClient.Options.AddOption(myClient.Identity, "NEWCHARGE", True, System.DBNull.Value)
                            Else
                                Me.lblCustFull.Text += " Does not have standard pricing"
                                Me.lblCustFull.ForeColor = System.Drawing.Color.Red
                            End If
                        End If
                        If myClient.VATRate = 0 Then
                            intRet = MsgBox("This customer has a zero rate for VAT, please ensure that this is correct", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)

                        End If
                    End If
                Catch ex As Exception
                    MedscreenLib.Medscreen.LogError(ex)

                End Try

                Me.CreditCard1.Billcommand.Hide()
                If Not Value Is Nothing Then
                    If MyCollections Is Nothing Then
                        If Not blnNoCollections Then

                            strWhere = "customer_id = '" & Value.Identity & "' and coll_date >= to_date('" & _
                            Now.AddYears(-1).ToString("yyyyMMdd") & "','yyyyMMdd') and status not in ('X') order by coll_date desc"

                            MyCollections = New Intranet.intranet.jobs.CMCollection(strWhere)
                            Me.cbCollections.DataSource = MyCollections
                            Me.DataBindings.Add("text", MyCollections, "ID")
                            AddHandler Me.BindingContextChanged, AddressOf cbCollections_SelectedIndexChanged

                            Me.cbCollections.DisplayMember = "ID"
                            blnClose = False
                            If MyCollections.Count = 0 Then
                                MsgBox("No collections require credit card details!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Information)
                                blnClose = True
                            End If
                        End If

                    End If
                End If
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex)
            End Try

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Don't open form 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DontOpen() As Boolean
        Get
            Return blnClose
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accept payment
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim intCount As Integer = 60
        Dim vbRet As MsgBoxResult
        Dim blnContinue As Boolean
        Dim nc As MedscreenLib.NetChat
        Dim Start As DateTime
        Dim current As DateTime
        Dim gap As TimeSpan

        Try
            'check to see if it has been invoiced 
            If objJob.Invoiceid.Trim.Length > 0 Then
                vbRet = MsgBox("This collection has already been invoiced, I can't continue!", MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            Start = Now
            'get credit card details and validate
            If CreditCard1.PaymentType = MedscreenCommonGui.CreditCard.CcPaymentTypes.CCCardPay Then
                If CreditCard1.Card Is Nothing Then
                    MsgBox("Select a valid credit card or create a new credit card!", MsgBoxStyle.Information And MsgBoxStyle.OKOnly)
                    Exit Sub
                End If
            End If
            'Check we have a value entered, this will normally be got from ICP
            If (Me.txtNett.Text.Trim.Length = 0) Then
                MsgBox("You must enter a value for the amount of the invoice", MsgBoxStyle.Exclamation And MsgBoxStyle.OKOnly)
                Exit Sub
            End If

            Dim oCI As Intranet.intranet.jobs.CollectionInvoice
            If CreditCard1.PaymentType = CreditCard.CcPaymentTypes.CCCardPay Then
                If icpR.TransactionId > 0 Then
                    oCI = objJob.CollectionInvoices.Item(icpR.TransactionId, Intranet.intranet.jobs.CollectionInvoices.IndexBy.TransactionID)
                Else
                    'At this stage we need to fill the ICP record and save it back to the database
                    icpR.TxnValue = Me.txtGrossPrice.Text
                    icpR.TransactionDate = Now
                    icpR.Cardnumber = Me.CreditCard1.Card.Number
                    icpR.AuthCode = Me.CreditCard1.AuthorisationCode
                    icpR.CardHolder = Me.CreditCard1.CardHolder
                    icpR.Expiry = Me.CreditCard1.ExpiryDate.ToString("yyMM")
                    icpR.Reference = Me.objJob.CMID & " " & Me.txtReference.Text & " " & Me.objJob.SmClient.SageID
                    icpR.UserName = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    icpR.SchemeName = Me.CreditCard1.Card.CardType
                    icpR.EFTSeqNum = icp.GetNextSeq
                    icpR.AuthMessage = "Authorised"
                    icp.insertRow(icpR)
                    icpR.TransactionId = icp.GetTransactionID(icpR.EFTSeqNum)
                    oCI = objJob.CollectionInvoices.CreateInvoice()
                    oCI.ICPReference = icpR.TransactionId
                    oCI.CCAuthCode = icpR.AuthCode
                    oCI.CCExpiry = Me.CreditCard1.ExpiryDate

                End If
            ElseIf CreditCard1.PaymentType = CreditCard.CcPaymentTypes.CCChequePay Then
                oCI = objJob.CollectionInvoices.CreateInvoice
            End If
            'Get a collection Invoice Object 
            Dim AccInterface As MedscreenLib.AccountsInterface = New MedscreenLib.AccountsInterface()
            AccInterface.Status = "R"
            AccInterface.UserId = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
            AccInterface.Reference = objJob.ID
            'At this stage we need to know whether we are dealing with nett or gross transactions
            Dim strSendMethod As String = CConnection.PackageStringList("lib_cctool.GetConfigItem", "CM_PAY_VALUE")
            If strSendMethod = "NET" Then
                AccInterface.Value = Me.txtNett.Text
            Else
                AccInterface.Value = icpR.TxnValue
            End If
            AccInterface.RequestCode = "RCPY"
            AccInterface.SetSendDateNull()


            'If Oisid = 0 Then        'Set up one stop interface code
            '    cTemp = cTest.NewInterface(objJob.ID)
            '    cTemp.CMJobNumber = objJob.ID
            'Else
            '    cTemp = cTest.NewInterface(Oisid)
            'End If
            'set up interface details
            'cTemp.PrePaid = True
            oCI.PrePaid = True
            objJob.PrePaid = True
            '.cTemp.CreditAuthorisation = CreditCard1.AuthorisationCode
            oCI.CCAuthCode = CreditCard1.AuthorisationCode
            objJob.CreditCardAuthCode = CreditCard1.AuthorisationCode
            'cTemp.CollDate = objJob.CollectionDate

            objJob.PaymentType = CreditCard1.PaymentType
            oCI.PaymentType = CreditCard1.PaymentType
            'Add timestamp to give more information
            'cTemp.Optional2 = Now.ToString("dd-MMM-yyyy HH:mm")
            objJob.Invoiceid = ""
            'check payment type
            If CreditCard1.PaymentType = MedscreenCommonGui.CreditCard.CcPaymentTypes.CCCardPay Then
                objJob.CreditCardNumber = CreditCard1.Card.Number
                oCI.CCNumber = CreditCard1.Card.Number
                objJob.CCExpiry = CreditCard1.Card.ExpiryDate
                oCI.CCExpiry = CreditCard1.Card.ExpiryDate
            Else
                objJob.CreditCardNumber = CreditCard1.Reference
                oCI.CMJobNumber = CreditCard1.Reference
                objJob.CCExpiry = Now
            End If
            objJob.CCPayee = CreditCard1.CardHolder
            oCI.CCPayee = CreditCard1.CardHolder

            'cTemp.NettCost = Me.txtNett.Text
            'cTemp.VatCost = Me.txtVAT.Text
            'cTemp.Customer = myClient.Identity
            If Not myUser Is Nothing Then
                'cTemp.operator = myUser.Identity

            Else
                'cTemp.operator = "Unknown"
            End If
            'cTemp.Status = Constants.GCST_IFace_TaskRequest
            'cTemp.Optional1 = Constants.GCST_IFace_TaskPrepay
            'cTemp.PurchaseOrder = objJob.PurchaseOrder
            objJob.Medscreenref = Me.txtReference.Text
            oCI.MedscreenReference = Me.txtReference.Text
            'cTemp.NoDonors = objJob.ExpectedSamples
            'cTemp.Optional4 = oCI.PaymentID
            AccInterface.SendAddress = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
            If Not Me.CreditCard1.Card Is Nothing Then AccInterface.SendAddress += ";" & Me.CreditCard1.Card.EmailAddress()
            If Not icpR Is Nothing Then
                oCI.ICPReference = icpR.TransactionId
            End If
            oCI.Update()
            objJob.Update()
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
            blnContinue = True
            gap = Now.Subtract(Start)
            LogAction("Finish : " & gap.TotalMilliseconds)

            Dim intMaxRepeat As Integer = 60
            While (AccInterface.Status = "R") AndAlso (intMaxRepeat > 0)
                AccInterface.Load()
                If AccInterface.Status = "C" Then
                    MsgBox("Invoice number " & AccInterface.Reference & " created")
                    oCI.InvoiceNumber = AccInterface.Reference
                    If Not icp Is Nothing AndAlso Not icpR Is Nothing Then
                        icp.SetInvoiceNumber(oCI.InvoiceNumber, icpR)
                    End If
                    LogAction("Invoice number " & AccInterface.Reference & " created - Collection : " & Me.objJob.ID)
                    oCI.Update()
                ElseIf AccInterface.Status = "F" Then
                    MsgBox("Transaction failed : - " & AccInterface.Message)
                    LogAction("Transaction failed : - " & AccInterface.Message & " - Collection : " & Me.objJob.ID)
                End If
                intMaxRepeat -= 1
                System.Threading.Thread.CurrentThread.Sleep(1000)
            End While


        Catch ex As Exception
            LogError(ex, , "Raising credit card payment")
        End Try
    End Sub

 

    Private Sub DoContUpdate(ByVal objJob As Intranet.intranet.jobs.CMJob)
        Dim strHtml As String
        Dim strTemp As String
        Dim Cost As Double
        Dim smClient As Intranet.intranet.customerns.Client
        Dim myPort As Intranet.intranet.jobs.Port
        Dim blnFixedSite As Boolean
        Dim cCharge As Double
        Dim dblSiteChg As Double
        Dim dblTotalCharge As Double
        Dim strCurr As String

        Try
            myPort = Nothing
            blnFixedSite = True
            If Not objJob Is Nothing Then
                Me.txtPO.Text = objJob.PurchaseOrder
                Me.DTCollDate.Value = objJob.CollectionDate
                If Len(Trim(objJob.Site)) > 0 Then
                    'myCentre = Collcentres.Item(objJob.Site)
                    myPort = objJob.CollectionPort
                    If Not myPort Is Nothing Then
                        blnFixedSite = myPort.IsFixedSite
                    End If
                End If
            End If

            strHtml = "<html><head><LINK REL=STYLESHEET TYPE=" & _
                Chr(34) & "text/css" & Chr(34) & " HREF = " & Chr(34) & _
            "\\Mark\home\MedScreen.css" & Chr(34) & " Title=" & Chr(34) & _
            "Style" & Chr(34) & "></head><body> "
            smClient = myClient
            If smClient Is Nothing Then smClient = myClient

            If Not myPort Is Nothing Then
                strHtml = strHtml & myPort.ToHTML(False, True)
            End If


            If Not smClient Is Nothing Then
                objJob.PriceJobHTML(smClient, strHtml, Cost)
                Me.txtNett.Text = Cost.ToString
            End If

            Try
                Me.HtmlEditor1.Refresh()
                Me.HtmlEditor1.DocumentText = (strHtml)
                Me.HtmlEditor1.Refresh()
                Me.HtmlEditor1.Invalidate()
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex)
            End Try
            Me.txtNett.Text = Cost
            Me.Oisid = cTemp.OisID
            Me.txtGrossPrice.Text = cTemp.NettCost
            Me.txtGrossPrice_Leave(Nothing, Nothing)
            Me.CreditCard1.Price = cTemp.NettCost
            Me.cbCollections.Text = objJob.ID
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        End Try
    End Sub

    Private Sub cbCollections_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCollections.SelectedIndexChanged
        Dim strHtml As String

        If blnInitialising Then Exit Sub
        If blnJobSelect Then Exit Sub
        Try
            Me.CreditCard1.Reference = Me.cbCollections.Text
            objJob = MyCollections.Item(Me.cbCollections.Text)
            DoContUpdate(objJob)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtNett_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNett.Leave

        Dim dblPrice As Double
        Dim vatRate As Double
        Dim nettPrice As Double
        Dim Vat As Double

        Try
            vatRate = Me.myClient.VATRate() / 100
            nettPrice = CDbl(txtNett.Text)
            Vat = (nettPrice * vatRate).ToString("0.00")
            Me.txtVAT.Text = Vat
            dblPrice = nettPrice + Vat
            Me.txtGrossPrice.Text = dblPrice
            Me.CreditCard1.Price = dblPrice
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtVAT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVAT.TextChanged
        Try
            Me.CreditCard1.Price = CDbl(txtNett.Text) + CDbl(txtVAT.Text)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub CreditCard1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreditCard1.Load

    End Sub



    Private Sub txtGrossPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGrossPrice.Leave
        Dim dblPrice As Double
        Dim vatRate As Double
        Dim nettPrice As Double
        Dim Vat As Double

        Try
            dblPrice = Me.txtGrossPrice.Text
            vatRate = Me.myClient.VATRate() / 100 + 1
            nettPrice = dblPrice / vatRate
            Me.txtNett.Text = nettPrice.ToString("0.00")
            Vat = dblPrice - nettPrice
            Me.txtVAT.Text = Vat.ToString("0.00")
            If Me.icpR Is Nothing Then Exit Sub
            If Me.icpR.Cardnumber.Trim.Length = 0 Then
                Me.icpR.TxnValue = Me.txtGrossPrice.Text
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub RBCheque_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBCheque.CheckedChanged
        Me.CreditCard1.PaymentType = MedscreenCommonGui.CreditCard.CcPaymentTypes.CCChequePay
        Me.CreditCard1.Invalidate()
    End Sub

    Private Sub RBCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBCash.CheckedChanged
        Me.CreditCard1.PaymentType = MedscreenCommonGui.CreditCard.CcPaymentTypes.CCCashPay
        Me.CreditCard1.Invalidate()
    End Sub

    Private Sub RBCCard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBCCard.CheckedChanged
        Me.CreditCard1.PaymentType = MedscreenCommonGui.CreditCard.CcPaymentTypes.CCCardPay
        Me.CreditCard1.Invalidate()

    End Sub

    Private Sub txtPO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPO.TextChanged
        If Me.objJob Is Nothing Then Exit Sub

        If Me.txtPO.Text.Trim.Length > 0 Then objJob.PurchaseOrder = Me.txtPO.Text
    End Sub
End Class
