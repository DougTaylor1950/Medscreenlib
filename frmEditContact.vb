Imports Intranet.intranet
Imports MedscreenLib

Public Class frmEditContact
    Inherits System.Windows.Forms.Form

    Private myContact As contacts.CustomerContact
    Private objContTypePhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContMethodPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objRepOptionPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContStatusPhrase As MedscreenLib.Glossary.PhraseCollection
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

    Private blnDrawn As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        blnDrawn = False

        '<Added code modified 05-Jul-2007 09:18 by LOGOS\Taylor> 
        'objContTypePhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_TYP")
        'objContTypePhrase.Load("Phrase_text")
        'objContMethodPhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_METH")
        'objContMethodPhrase.Load("Phrase_text")
        'objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("REP_OPTION")
        'objRepOptionPhrase.Load("Phrase_text")

        'objContStatusPhrase = New MedscreenLib.Glossary.PhraseCollection("CONT_STAT")
        'objContStatusPhrase.Load("Phrase_text")

        '</Added code modified 05-Jul-2007 09:18 by LOGOS\Taylor> 

        Me.cbContactType.PhraseType = "CNTCT_TYP"
        'Me.cbContactType.DisplayMember = "PhraseText"
        'Me.cbContactType.ValueMember = "PhraseID"

        Me.cbContMethod.PhraseType = "CNTCT_METH"
        'Me.cbContMethod.DisplayMember = "PhraseText"
        'Me.cbContMethod.ValueMember = "PhraseID"

        Me.cbReportOptions.PhraseType = "REP_OPTION"
        'Me.cbReportOptions.DisplayMember = "PhraseText"
        'Me.cbReportOptions.ValueMember = "PhraseID"


        Me.cbContactStatus.PhraseType = "CONT_STAT"
        'Me.cbContactStatus.DisplayMember = "PhraseText"
        'Me.cbContactStatus.ValueMember = "PhraseID"




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
    Friend WithEvents pnlContact As System.Windows.Forms.Panel
    Friend WithEvents cbContactStatus As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label87 As System.Windows.Forms.Label
    Friend WithEvents cbReportOptions As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label85 As System.Windows.Forms.Label
    Friend WithEvents cbReportFormat As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label86 As System.Windows.Forms.Label
    Friend WithEvents cbContMethod As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label84 As System.Windows.Forms.Label
    Friend WithEvents cbContactType As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents Label83 As System.Windows.Forms.Label
    Friend WithEvents AdrContact As MedscreenCommonGui.AddressPanel
    Friend WithEvents Label81 As System.Windows.Forms.Label
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents cmdFindContactAdr As System.Windows.Forms.Button
    Friend WithEvents cmdNewContactAdr As System.Windows.Forms.Button
    Friend WithEvents ckRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents pnlSecurity As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditContact))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.pnlContact = New System.Windows.Forms.Panel
        Me.ckRemoved = New System.Windows.Forms.CheckBox
        Me.cbContactStatus = New MedscreenCommonGui.PhraseComboBox
        Me.Label87 = New System.Windows.Forms.Label
        Me.cbReportOptions = New MedscreenCommonGui.PhraseComboBox
        Me.Label85 = New System.Windows.Forms.Label
        Me.cbReportFormat = New MedscreenCommonGui.PhraseComboBox
        Me.Label86 = New System.Windows.Forms.Label
        Me.cbContMethod = New MedscreenCommonGui.PhraseComboBox
        Me.Label84 = New System.Windows.Forms.Label
        Me.cbContactType = New MedscreenCommonGui.PhraseComboBox
        Me.Label83 = New System.Windows.Forms.Label
        Me.AdrContact = New MedscreenCommonGui.AddressPanel
        Me.pnlSecurity = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label81 = New System.Windows.Forms.Label
        Me.Label82 = New System.Windows.Forms.Label
        Me.cmdFindContactAdr = New System.Windows.Forms.Button
        Me.cmdNewContactAdr = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.pnlContact.SuspendLayout()
        Me.AdrContact.SuspendLayout()
        Me.pnlSecurity.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 389)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(832, 32)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(748, 4)
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
        Me.cmdOk.Location = New System.Drawing.Point(668, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlContact
        '
        Me.pnlContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlContact.Controls.Add(Me.ckRemoved)
        Me.pnlContact.Controls.Add(Me.cbContactStatus)
        Me.pnlContact.Controls.Add(Me.Label87)
        Me.pnlContact.Controls.Add(Me.cbReportOptions)
        Me.pnlContact.Controls.Add(Me.Label85)
        Me.pnlContact.Controls.Add(Me.cbReportFormat)
        Me.pnlContact.Controls.Add(Me.Label86)
        Me.pnlContact.Controls.Add(Me.cbContMethod)
        Me.pnlContact.Controls.Add(Me.Label84)
        Me.pnlContact.Controls.Add(Me.cbContactType)
        Me.pnlContact.Controls.Add(Me.Label83)
        Me.pnlContact.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContact.Location = New System.Drawing.Point(576, 0)
        Me.pnlContact.Name = "pnlContact"
        Me.pnlContact.Size = New System.Drawing.Size(256, 389)
        Me.pnlContact.TabIndex = 246
        '
        'ckRemoved
        '
        Me.ckRemoved.Location = New System.Drawing.Point(48, 304)
        Me.ckRemoved.Name = "ckRemoved"
        Me.ckRemoved.Size = New System.Drawing.Size(104, 24)
        Me.ckRemoved.TabIndex = 236
        Me.ckRemoved.Text = "Removed"
        Me.ToolTip1.SetToolTip(Me.ckRemoved, "If checked the contact will no longer be contacted, for temporary changes use the" & _
                " status.")
        '
        'cbContactStatus
        '
        Me.cbContactStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContactStatus.DisplayMember = "PhraseText"
        Me.cbContactStatus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContactStatus.Location = New System.Drawing.Point(44, 264)
        Me.cbContactStatus.Name = "cbContactStatus"
        Me.cbContactStatus.PhraseType = Nothing
        Me.cbContactStatus.Size = New System.Drawing.Size(201, 21)
        Me.cbContactStatus.TabIndex = 235
        Me.ToolTip1.SetToolTip(Me.cbContactStatus, "If a contact is made in active they will remain as a contact but will not be cont" & _
                "acted, use this for maternity leave etc.")
        Me.cbContactStatus.ValueMember = "PhraseID"
        '
        'Label87
        '
        Me.Label87.Location = New System.Drawing.Point(44, 240)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(100, 23)
        Me.Label87.TabIndex = 234
        Me.Label87.Text = "Contact Status"
        '
        'cbReportOptions
        '
        Me.cbReportOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReportOptions.DisplayMember = "PhraseText"
        Me.cbReportOptions.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbReportOptions.Location = New System.Drawing.Point(40, 212)
        Me.cbReportOptions.Name = "cbReportOptions"
        Me.cbReportOptions.PhraseType = Nothing
        Me.cbReportOptions.Size = New System.Drawing.Size(201, 21)
        Me.cbReportOptions.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.cbReportOptions, "There may be different options that can be set to modify the report.")
        Me.cbReportOptions.ValueMember = "PhraseID"
        '
        'Label85
        '
        Me.Label85.Location = New System.Drawing.Point(40, 188)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(100, 23)
        Me.Label85.TabIndex = 6
        Me.Label85.Text = "Report Options"
        '
        'cbReportFormat
        '
        Me.cbReportFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReportFormat.DisplayMember = "PhraseText"
        Me.cbReportFormat.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbReportFormat.Location = New System.Drawing.Point(40, 156)
        Me.cbReportFormat.Name = "cbReportFormat"
        Me.cbReportFormat.PhraseType = Nothing
        Me.cbReportFormat.Size = New System.Drawing.Size(201, 21)
        Me.cbReportFormat.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.cbReportFormat, "SElect the type of the report to send, different lists will be displayed for each" & _
                " contact type")
        Me.cbReportFormat.ValueMember = "PhraseID"
        '
        'Label86
        '
        Me.Label86.Location = New System.Drawing.Point(40, 132)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(100, 23)
        Me.Label86.TabIndex = 4
        Me.Label86.Text = "Report Format"
        '
        'cbContMethod
        '
        Me.cbContMethod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContMethod.DisplayMember = "PhraseText"
        Me.cbContMethod.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContMethod.Location = New System.Drawing.Point(40, 96)
        Me.cbContMethod.Name = "cbContMethod"
        Me.cbContMethod.PhraseType = Nothing
        Me.cbContMethod.Size = New System.Drawing.Size(201, 21)
        Me.cbContMethod.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cbContMethod, "Select the method of contact, usually email")
        Me.cbContMethod.ValueMember = "PhraseID"
        '
        'Label84
        '
        Me.Label84.Location = New System.Drawing.Point(40, 72)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(100, 23)
        Me.Label84.TabIndex = 2
        Me.Label84.Text = "Contact Method"
        '
        'cbContactType
        '
        Me.cbContactType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbContactType.DisplayMember = "PhraseText"
        Me.cbContactType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cbContactType.Location = New System.Drawing.Point(40, 40)
        Me.cbContactType.Name = "cbContactType"
        Me.cbContactType.PhraseType = Nothing
        Me.cbContactType.Size = New System.Drawing.Size(201, 21)
        Me.cbContactType.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cbContactType, "SElect the type of contact, the other options will change as a reult of this sele" & _
                "ction")
        Me.cbContactType.ValueMember = "PhraseID"
        '
        'Label83
        '
        Me.Label83.Location = New System.Drawing.Point(40, 16)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(100, 23)
        Me.Label83.TabIndex = 0
        Me.Label83.Text = "Contact Type"
        '
        'AdrContact
        '
        Me.AdrContact.Address = Nothing
        Me.AdrContact.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AdrContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AdrContact.Controls.Add(Me.pnlSecurity)
        Me.AdrContact.Controls.Add(Me.Label81)
        Me.AdrContact.Controls.Add(Me.Label82)
        Me.AdrContact.Controls.Add(Me.cmdFindContactAdr)
        Me.AdrContact.Controls.Add(Me.cmdNewContactAdr)
        Me.AdrContact.Dock = System.Windows.Forms.DockStyle.Left
        Me.AdrContact.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AdrContact.InfoVisible = False
        Me.AdrContact.isReadOnly = False
        Me.AdrContact.Location = New System.Drawing.Point(0, 0)
        Me.AdrContact.Name = "AdrContact"
        Me.AdrContact.ShowCancelCmd = False
        Me.AdrContact.ShowCloneCmd = False
        Me.AdrContact.ShowContact = True
        Me.AdrContact.ShowFindCmd = False
        Me.AdrContact.ShowNewCmd = False
        Me.AdrContact.ShowPanel = False
        Me.AdrContact.ShowSaveCmd = False
        Me.AdrContact.ShowStatus = False
        Me.AdrContact.Size = New System.Drawing.Size(576, 389)
        Me.AdrContact.TabIndex = 245
        Me.AdrContact.Tooltip = Nothing
        Me.ToolTip1.SetToolTip(Me.AdrContact, "Address for Contact, email must be filled in.")
        '
        'pnlSecurity
        '
        Me.pnlSecurity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlSecurity.Controls.Add(Me.Label3)
        Me.pnlSecurity.Controls.Add(Me.txtPassword)
        Me.pnlSecurity.Controls.Add(Me.txtUserName)
        Me.pnlSecurity.Controls.Add(Me.Label2)
        Me.pnlSecurity.Controls.Add(Me.Label1)
        Me.pnlSecurity.Location = New System.Drawing.Point(16, 280)
        Me.pnlSecurity.Name = "pnlSecurity"
        Me.pnlSecurity.Size = New System.Drawing.Size(536, 104)
        Me.pnlSecurity.TabIndex = 236
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(64, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(440, 32)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "* These will be supplied by IT when the mailbox is setup, and relates to the @Con" & _
            "cateno.com email account above"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(304, 32)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(160, 20)
        Me.txtPassword.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtPassword, "Pasword to Concateno email.")
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(304, 8)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(160, 20)
        Me.txtUserName.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtUserName, "Login name to Concateno Email")
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(64, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 23)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = " password *:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(64, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Secure email user name/:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label81
        '
        Me.Label81.Location = New System.Drawing.Point(392, 72)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(16, 16)
        Me.Label81.TabIndex = 235
        '
        'Label82
        '
        Me.Label82.Location = New System.Drawing.Point(48, 184)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(16, 16)
        Me.Label82.TabIndex = 234
        '
        'cmdFindContactAdr
        '
        Me.cmdFindContactAdr.Image = CType(resources.GetObject("cmdFindContactAdr.Image"), System.Drawing.Image)
        Me.cmdFindContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdFindContactAdr.Location = New System.Drawing.Point(192, 256)
        Me.cmdFindContactAdr.Name = "cmdFindContactAdr"
        Me.cmdFindContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdFindContactAdr.TabIndex = 233
        Me.cmdFindContactAdr.Text = "Find Address"
        Me.cmdFindContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdNewContactAdr
        '
        Me.cmdNewContactAdr.Image = CType(resources.GetObject("cmdNewContactAdr.Image"), System.Drawing.Image)
        Me.cmdNewContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewContactAdr.Location = New System.Drawing.Point(64, 256)
        Me.cmdNewContactAdr.Name = "cmdNewContactAdr"
        Me.cmdNewContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewContactAdr.TabIndex = 232
        Me.cmdNewContactAdr.Text = "New Address"
        Me.cmdNewContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.AlwaysBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'frmEditContact
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(832, 421)
        Me.Controls.Add(Me.pnlContact)
        Me.Controls.Add(Me.AdrContact)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEditContact"
        Me.Text = "Edit Contact"
        Me.Panel1.ResumeLayout(False)
        Me.pnlContact.ResumeLayout(False)
        Me.AdrContact.ResumeLayout(False)
        Me.AdrContact.PerformLayout()
        Me.pnlSecurity.ResumeLayout(False)
        Me.pnlSecurity.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub lvContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

 
    End Sub

    Public Property Contact() As contacts.CustomerContact
        Get
            Return myContact
        End Get
        Set(ByVal Value As contacts.CustomerContact)
            myContact = Value
            If Not myContact Is Nothing Then
                Me.ckRemoved.Checked = myContact.Removed
                Me.AdrContact.Address = myContact.Address
                Me.cbContactType.SelectedValue = myContact.ContactType
                Me.cbContactType.SelectionLength = 0
                SetReportOptions()
                Me.cbContMethod.SelectedValue = myContact.ContactMethod
                Me.cbReportOptions.SelectedValue = myContact.ReportOptions
                Me.cbContactStatus.SelectedValue = myContact.ContactStatus
                Me.txtPassword.Text = myContact.Password
                Me.txtUserName.Text = myContact.UserName
            Else
                Me.AdrContact.Address = Nothing
                Me.cbContactType.SelectedValue = ""
                Me.cbContMethod.SelectedValue = ""
                Me.cbReportOptions.SelectedValue = ""

            End If

        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUpdateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        'Update info 
        If myContact Is Nothing Then Exit Sub

        'validate Contact call PL/SQL code

        'Prepare for noraml return
        ErrorProvider1.SetError(Me.cbReportFormat, "")
        Me.DialogResult = Windows.Forms.DialogResult.OK

        Dim oColl As New Collection
        oColl.Add(myContact.CustomerId)
        oColl.Add(Me.cbReportOptions.SelectedValue)
        Dim strError As String = CConnection.PackageStringList("LIB_CONTACTS.ValidateContacts", oColl)
        If strError IsNot Nothing AndAlso strError.Trim.Length > 0 Then 'There are issues or error messages that need dealing with. 
            'There may be a pipe character | in the string if so following this is the severity level
            Dim intSeverity As Integer = 0
            If InStr(strError, "|") Then
                Dim Parts As String() = strError.Split(New Char() {"|"})
                If Parts.Length > 1 Then
                    strError = Parts(0)
                    intSeverity = CInt(Parts(1)) 'Could I also suggest level 0 for information (no action), 1 for warning (no action)
                    If intSeverity > 1 Then
                        ErrorProvider1.SetError(Me.cbReportFormat, strError)
                        Me.DialogResult = Windows.Forms.DialogResult.None
                        Exit Sub
                    Else
                        MsgBox(strError, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
                    End If
                End If
            End If
        End If

        If Not Me.AdrContact.Address Is Nothing Then myContact.AddressId = Me.AdrContact.Address.AddressId
        myContact.Address = Me.AdrContact.Address
        If Not myContact.Address Is Nothing Then myContact.Address.Update()
        myContact.ContactMethod = Me.cbContMethod.SelectedValue
        myContact.ContactType = Me.cbContactType.SelectedValue
        myContact.ReportOptions = Me.cbReportOptions.SelectedValue
        myContact.ContactStatus = Me.cbContactStatus.SelectedValue
        If myContact.ContactStatus.Trim.Length = 0 Then myContact.ContactStatus = myContact.GCST_STATUS_ACTIVE
        myContact.Removed = Me.ckRemoved.Checked
        myContact.Password = Me.txtPassword.Text
        myContact.UserName = Me.txtUserName.Text


        myContact.Update()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdNewContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewContactAdr.Click
        If Me.myContact Is Nothing Then Exit Sub
        Dim addr As MedscreenLib.Address.Caddress
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "", Me.myContact.CustomerId, MedscreenLib.Constants.AddressType.AddressMain)
        Me.AdrContact.Address = addr
        Me.myContact.Address = addr
        Me.myContact.AddressId = addr.AddressId
    End Sub

    Private Sub cmdFindContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindContactAdr.Click
        If Me.myContact Is Nothing Then Exit Sub

        Dim fndAddr As New MedscreenCommonGui.FindAddress()

        Dim intAddrID As Integer
        With fndAddr
            If .ShowDialog = DialogResult.OK Then               'If user has found a
                intAddrID = .AddressID                          ' a valid address put it in 
                'get address and fill control
                Dim addr As MedscreenLib.Address.Caddress
                If intAddrID > 0 Then
                    addr = MedscreenLib.Glossary.Glossary.SMAddresses.item(intAddrID, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                    Me.AdrContact.Address = addr
                    Me.myContact.Address = addr
                    Me.myContact.AddressId = addr.AddressId

                    'MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
                    'TODO: Need to change address links
                End If
                'make controls invisible
                'Me.cmdNewAddress.Visible = False
                'Me.cmdFindAddress.Visible = False

            End If
        End With

    End Sub


    Private Sub SetReportOptions()
        Me.cbReportOptions.Text = ""
        Me.cbReportOptions.SelectedIndex = -1
        objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("select phrase_id,phrase_text from phrase where phrase_type = '" & cbContactType.SelectedValue & "'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        cbReportOptions.DataSource = objRepOptionPhrase
        'The next line needs to get its information from config file
        Dim strContacttypes As String = MedscreenCommonGUIConfig.NodeQuery("CONTACTTYPE")
        If Medscreen.UserInRole(MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity, "MEDSECUREEMAIL") AndAlso InStr(cbContactType.SelectedValue, strContacttypes) > 0 Then
            Me.pnlSecurity.Show()
        Else
            Me.pnlSecurity.Hide()

        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' changing the contact type will change the report options
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/12/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cbContactType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbContactType.SelectedIndexChanged
        If Not (blnDrawn) Then Exit Sub
        SetReportOptions()
    End Sub

    Private Sub frmEditContact_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        blnDrawn = True
    End Sub

    Private Sub AdrContact_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles AdrContact.Paint

    End Sub
End Class
