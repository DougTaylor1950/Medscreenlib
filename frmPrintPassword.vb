Imports MedscreenLib
Public Class frmPrintPassword
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.rbNewPass.Checked = True
        Me.rbPrinter.Checked = True
        Dim inF As IniFile.IniFiles = New IniFile.IniFiles()

        'inF.FileName = MedscreenLib.Medscreen.LiveRoot & MedscreenLib.Medscreen.IniFiles & "reporterprinters.ini"
        'Dim PColl As Collection = New Collection()
        'PColl = inF.ReadSection("Printers")

        'Dim prColl As New IniFile.IniCollection()
        'Dim pStr As String
        'Dim i As Integer
        'Dim iItem As IniFile.IniElement
        'Dim inIndex As Integer

        'For i = 1 To PColl.Count
        '    pStr = PColl.Item(i)
        '    inIndex = InStr(pStr, "=")
        '    If inIndex > 0 Then
        '        iItem = New IniFile.IniElement()
        '        iItem.Header = Mid(pStr, 1, inIndex - 1)
        '        iItem.Item = Mid(pStr, inIndex + 1)
        '        prColl.Add(iItem)
        '        Me.cbPrinters.Items.Add(iItem.Item)
        '    End If
        'Next
        Dim prColl As MedscreenLib.Glossary.PrinterColl = MedscreenLib.Glossary.Glossary.Printers
        Me.cbPrinters.DataSource = prColl
        Me.cbPrinters.ValueMember = "Identity"
        Me.cbPrinters.DisplayMember = "Info"
        If prColl.Count > 0 Then
            Me.cbPrinters.SelectedIndex = 0
        End If


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
    Friend WithEvents grpLetterType As System.Windows.Forms.GroupBox
    Friend WithEvents rbNewPass As System.Windows.Forms.RadioButton
    Friend WithEvents rbConfirm As System.Windows.Forms.RadioButton
    Friend WithEvents rbChanged As System.Windows.Forms.RadioButton
    Friend WithEvents grpDestination As System.Windows.Forms.GroupBox
    Friend WithEvents rbPrinter As System.Windows.Forms.RadioButton
    Friend WithEvents rbFax As System.Windows.Forms.RadioButton
    Friend WithEvents rbEmail As System.Windows.Forms.RadioButton
    Friend WithEvents pnlPrinter As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbPrinters As System.Windows.Forms.ComboBox
    Friend WithEvents pnlEmail As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents pnkFax As System.Windows.Forms.Panel
    Friend WithEvents txtFaxNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents pnlText As System.Windows.Forms.Panel
    Friend WithEvents labelMsg As System.Windows.Forms.Label
    Friend WithEvents lblEmailAddress As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrintPassword))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.grpLetterType = New System.Windows.Forms.GroupBox()
        Me.rbChanged = New System.Windows.Forms.RadioButton()
        Me.rbConfirm = New System.Windows.Forms.RadioButton()
        Me.rbNewPass = New System.Windows.Forms.RadioButton()
        Me.grpDestination = New System.Windows.Forms.GroupBox()
        Me.rbPrinter = New System.Windows.Forms.RadioButton()
        Me.rbFax = New System.Windows.Forms.RadioButton()
        Me.rbEmail = New System.Windows.Forms.RadioButton()
        Me.pnlPrinter = New System.Windows.Forms.Panel()
        Me.cbPrinters = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlEmail = New System.Windows.Forms.Panel()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pnkFax = New System.Windows.Forms.Panel()
        Me.txtFaxNo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.pnlText = New System.Windows.Forms.Panel()
        Me.labelMsg = New System.Windows.Forms.Label()
        Me.lblEmailAddress = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.grpLetterType.SuspendLayout()
        Me.grpDestination.SuspendLayout()
        Me.pnlPrinter.SuspendLayout()
        Me.pnlEmail.SuspendLayout()
        Me.pnkFax.SuspendLayout()
        Me.pnlText.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 421)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(680, 40)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(592, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(512, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpLetterType
        '
        Me.grpLetterType.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbChanged, Me.rbConfirm, Me.rbNewPass})
        Me.grpLetterType.Location = New System.Drawing.Point(32, 40)
        Me.grpLetterType.Name = "grpLetterType"
        Me.grpLetterType.Size = New System.Drawing.Size(200, 136)
        Me.grpLetterType.TabIndex = 2
        Me.grpLetterType.TabStop = False
        Me.grpLetterType.Text = "Letter Type"
        '
        'rbChanged
        '
        Me.rbChanged.Location = New System.Drawing.Point(24, 104)
        Me.rbChanged.Name = "rbChanged"
        Me.rbChanged.Size = New System.Drawing.Size(168, 24)
        Me.rbChanged.TabIndex = 2
        Me.rbChanged.Text = "C&hanged Password Letter"
        '
        'rbConfirm
        '
        Me.rbConfirm.Location = New System.Drawing.Point(24, 68)
        Me.rbConfirm.Name = "rbConfirm"
        Me.rbConfirm.Size = New System.Drawing.Size(168, 24)
        Me.rbConfirm.TabIndex = 1
        Me.rbConfirm.Text = "&Confirm Password Letter"
        '
        'rbNewPass
        '
        Me.rbNewPass.Location = New System.Drawing.Point(24, 32)
        Me.rbNewPass.Name = "rbNewPass"
        Me.rbNewPass.Size = New System.Drawing.Size(144, 24)
        Me.rbNewPass.TabIndex = 0
        Me.rbNewPass.Text = "&New Password Letter"
        '
        'grpDestination
        '
        Me.grpDestination.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbPrinter, Me.rbFax, Me.rbEmail})
        Me.grpDestination.Location = New System.Drawing.Point(400, 40)
        Me.grpDestination.Name = "grpDestination"
        Me.grpDestination.Size = New System.Drawing.Size(200, 136)
        Me.grpDestination.TabIndex = 3
        Me.grpDestination.TabStop = False
        Me.grpDestination.Text = "Destination"
        '
        'rbPrinter
        '
        Me.rbPrinter.Location = New System.Drawing.Point(16, 72)
        Me.rbPrinter.Name = "rbPrinter"
        Me.rbPrinter.Size = New System.Drawing.Size(168, 24)
        Me.rbPrinter.TabIndex = 2
        Me.rbPrinter.Text = "&Printer"
        '
        'rbFax
        '
        Me.rbFax.Location = New System.Drawing.Point(16, 104)
        Me.rbFax.Name = "rbFax"
        Me.rbFax.Size = New System.Drawing.Size(168, 24)
        Me.rbFax.TabIndex = 1
        Me.rbFax.Text = "&Fax"
        Me.rbFax.Visible = False
        '
        'rbEmail
        '
        Me.rbEmail.Location = New System.Drawing.Point(16, 32)
        Me.rbEmail.Name = "rbEmail"
        Me.rbEmail.Size = New System.Drawing.Size(144, 24)
        Me.rbEmail.TabIndex = 0
        Me.rbEmail.Text = "&Email"
        '
        'pnlPrinter
        '
        Me.pnlPrinter.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbPrinters, Me.Label1})
        Me.pnlPrinter.Location = New System.Drawing.Point(32, 200)
        Me.pnlPrinter.Name = "pnlPrinter"
        Me.pnlPrinter.Size = New System.Drawing.Size(448, 48)
        Me.pnlPrinter.TabIndex = 4
        Me.pnlPrinter.Visible = False
        '
        'cbPrinters
        '
        Me.cbPrinters.Location = New System.Drawing.Point(144, 16)
        Me.cbPrinters.Name = "cbPrinters"
        Me.cbPrinters.Size = New System.Drawing.Size(296, 21)
        Me.cbPrinters.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Printers"
        '
        'pnlEmail
        '
        Me.pnlEmail.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtEmail, Me.Label2})
        Me.pnlEmail.Location = New System.Drawing.Point(32, 200)
        Me.pnlEmail.Name = "pnlEmail"
        Me.pnlEmail.Size = New System.Drawing.Size(448, 48)
        Me.pnlEmail.TabIndex = 5
        Me.pnlEmail.Visible = False
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(136, 14)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(280, 20)
        Me.txtEmail.TabIndex = 1
        Me.txtEmail.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(48, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "EmailAddress"
        '
        'pnkFax
        '
        Me.pnkFax.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtFaxNo, Me.Label3})
        Me.pnkFax.Location = New System.Drawing.Point(32, 199)
        Me.pnkFax.Name = "pnkFax"
        Me.pnkFax.Size = New System.Drawing.Size(448, 48)
        Me.pnkFax.TabIndex = 6
        Me.pnkFax.Visible = False
        '
        'txtFaxNo
        '
        Me.txtFaxNo.Location = New System.Drawing.Point(136, 14)
        Me.txtFaxNo.Name = "txtFaxNo"
        Me.txtFaxNo.Size = New System.Drawing.Size(280, 20)
        Me.txtFaxNo.TabIndex = 1
        Me.txtFaxNo.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(48, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Fax No"
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(176, 264)
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.Size = New System.Drawing.Size(280, 20)
        Me.txtOutput.TabIndex = 8
        Me.txtOutput.Text = ""
        Me.txtOutput.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(64, 264)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "output to"
        Me.Label4.Visible = False
        '
        'pnlText
        '
        Me.pnlText.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.pnlText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlText.Controls.AddRange(New System.Windows.Forms.Control() {Me.labelMsg, Me.lblEmailAddress})
        Me.pnlText.Location = New System.Drawing.Point(16, 184)
        Me.pnlText.Name = "pnlText"
        Me.pnlText.Size = New System.Drawing.Size(648, 216)
        Me.pnlText.TabIndex = 9
        '
        'labelMsg
        '
        Me.labelMsg.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.labelMsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelMsg.Location = New System.Drawing.Point(16, 16)
        Me.labelMsg.Name = "labelMsg"
        Me.labelMsg.Size = New System.Drawing.Size(600, 184)
        Me.labelMsg.TabIndex = 0
        Me.labelMsg.Text = "The password document will be created in Word, this will be open on the desktop l" & _
        "etting you correct any details prior to printing "
        '
        'lblEmailAddress
        '
        Me.lblEmailAddress.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.lblEmailAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmailAddress.Location = New System.Drawing.Point(18, 14)
        Me.lblEmailAddress.Name = "lblEmailAddress"
        Me.lblEmailAddress.Size = New System.Drawing.Size(606, 186)
        Me.lblEmailAddress.TabIndex = 1
        Me.lblEmailAddress.Visible = False
        '
        'frmPrintPassword
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 461)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlText, Me.txtOutput, Me.Label4, Me.grpDestination, Me.grpLetterType, Me.Panel1, Me.pnlPrinter, Me.pnkFax, Me.pnlEmail})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrintPassword"
        Me.Text = "Print Password"
        Me.Panel1.ResumeLayout(False)
        Me.grpLetterType.ResumeLayout(False)
        Me.grpDestination.ResumeLayout(False)
        Me.pnlPrinter.ResumeLayout(False)
        Me.pnlEmail.ResumeLayout(False)
        Me.pnkFax.ResumeLayout(False)
        Me.pnlText.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private myConfirm As Intranet.intranet.customerns.ClientReport.PasswordLetters
    Private myClient As Intranet.intranet.customerns.Client

    Public Enum OutputOption
        Printer
        Email
        Fax
    End Enum

    Private myOutputOption As OutputOption = OutputOption.Email

    Public Property OutputAs() As OutputOption
        Get
            Return myOutputOption
        End Get
        Set(ByVal Value As OutputOption)
            myOutputOption = Value
        End Set
    End Property

    Public Property ConfirmMethod() As Intranet.intranet.customerns.ClientReport.PasswordLetters
        Get
            Return myConfirm
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.ClientReport.PasswordLetters)
            myConfirm = Value
        End Set
    End Property

    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            myClient = Value
            If Not myClient Is Nothing Then
                'Me.txtEmail.Text = myClient.EmailResults
                'Me.txtFaxNo.Text = myClient.ReportingInfo.FaxNumber
            End If
        End Set
    End Property

    Private Sub rbNewPass_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbNewPass.CheckedChanged, rbChanged.CheckedChanged, rbConfirm.CheckedChanged
        If sender.name = Me.rbChanged.Name Then
            myConfirm = Intranet.intranet.customerns.ClientReport.PasswordLetters.ChangedPassword
        ElseIf sender.name = Me.rbNewPass.Name Then
            myConfirm = Intranet.intranet.customerns.ClientReport.PasswordLetters.NewPassword
        Else
            myConfirm = Intranet.intranet.customerns.ClientReport.PasswordLetters.Reminder
        End If

    End Sub

    Private Sub rbPrinter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbPrinter.CheckedChanged, rbEmail.CheckedChanged, rbFax.CheckedChanged
        pnlPrinter.Hide()
        If sender.name = Me.rbPrinter.Name Then

            myOutputOption = OutputOption.Printer
            Me.labelMsg.Show()
            Me.lblEmailAddress.Hide()
            Me.pnlPrinter.Show()
        ElseIf sender.name = Me.rbEmail.Name Then
            Me.labelMsg.Hide()
            Me.lblEmailAddress.Show()
            If Not Me.myClient Is Nothing Then
                If Not Me.myClient.ResultsAddress Is Nothing Then
                    Me.lblEmailAddress.Text = "Send to " & vbCrLf & myClient.ResultsAddress.BlockAddress(vbCrLf, True) & vbCrLf & myClient.ResultsAddress.Email
                End If
            End If
            myOutputOption = OutputOption.Email
            'Else
            '    Me.pnlPrinter.Hide()
            '    Me.pnlEmail.Hide()
            '    'Me.pnkFax.Show()
            '    Me.pnkFax.BringToFront()
            '    Me.txtOutput.Text = Medscreen.ReplaceString(Me.txtFaxNo.Text, " ", "") & "@starfax.co.uk"
            '    myOutputOption = OutputOption.Fax
        End If
    End Sub

    Private Sub cbPrinters_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPrinters.SelectedIndexChanged
        If Not Me.cbPrinters.SelectedValue Is Nothing Then Me.txtOutput.Text = CType(Me.cbPrinters.SelectedItem, MedscreenLib.Glossary.Printer).Identity
        'End If
    End Sub

    Private Sub txtFaxNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaxNo.Leave
        Me.txtOutput.Text = Medscreen.ReplaceString(Me.txtFaxNo.Text, " ", "") & "@starfax.co.uk"

    End Sub

    Private Sub txtEmail_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.Leave
        Me.txtOutput.Text = Me.txtEmail.Text

    End Sub

    Private Function ValidateForm() As Boolean
        Dim blnRet As Boolean = False
        Me.ErrorProvider1.SetError(Me.txtFaxNo, "")
        Me.ErrorProvider1.SetError(Me.txtEmail, "")
        If Me.myOutputOption = OutputOption.Email Then
            'If Medscreen.ValidateEmail(Me.txtEmail.Text) = Constants.EmailErrors.NoAddress Then
            '    Me.ErrorProvider1.SetError(Me.txtEmail, "No email address")
            '    blnRet = True
            'ElseIf Medscreen.ValidateEmail(Me.txtEmail.Text) = Constants.EmailErrors.Invalid Then
            '    Me.ErrorProvider1.SetError(Me.txtEmail, "Invalid email address")
            '    blnRet = True
            'ElseIf Medscreen.ValidateEmail(Me.txtEmail.Text) = Constants.EmailErrors.NoAt Then
            '    Me.ErrorProvider1.SetError(Me.txtEmail, "No @ in  address")
            '    blnRet = True
            'ElseIf Medscreen.ValidateEmail(Me.txtEmail.Text) = Constants.EmailErrors.NoDomain Then
            '    Me.ErrorProvider1.SetError(Me.txtEmail, "No Domain in  address")
            '    blnRet = True
            'End If
        ElseIf Me.myOutputOption = OutputOption.Fax Then
            'If Not Medscreen.IsNumber(Medscreen.ReplaceString(Me.txtFaxNo.Text, " ", "")) Then
            '    Me.ErrorProvider1.SetError(Me.txtFaxNo, "Invalid fax number")
            '    blnRet = True
            'End If
        End If
        Return blnRet
    End Function

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Me.DialogResult = DialogResult.OK
        If ValidateForm() Then
            Me.DialogResult = DialogResult.None
        End If
    End Sub
End Class
