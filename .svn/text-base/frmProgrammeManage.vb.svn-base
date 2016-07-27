'''$Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-22 06:30:28+00 $ <para/>
'''$Log: frmAbout.vb,v $ <para/>

Imports Intranet.intranet
Imports Intranet.intranet.customerns
Imports MedscreenLib

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : frmProgrammeManage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to manage Customers test programme
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [19/01/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmProgrammeManage
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.cbContactFrequency.DataSource = MedscreenLib.Glossary.Glossary.PMCustomerContactFrequency
        Me.cbContactFrequency.DisplayMember = "PhraseText"
        Me.cbContactFrequency.ValueMember = "PhraseID"

        Me.cbProgType.DataSource = MedscreenLib.Glossary.Glossary.PMAccountType
        Me.cbProgType.DisplayMember = "PhraseText"
        Me.cbProgType.ValueMember = "PhraseID"

        Me.cbTestFreq.DataSource = MedscreenLib.Glossary.Glossary.PMVesselTestFrequency
        Me.cbTestFreq.DisplayMember = "PhraseText"
        Me.cbTestFreq.ValueMember = "PhraseID"

        Me.cbProgMan.DataSource = MedscreenLib.Glossary.Glossary.SampleMangerAssignments.RoleHolders("PROGRAMMEMANAGER")
        Me.cbProgMan.DisplayMember = "User"
        Me.cbProgMan.ValueMember = "User"

        objContTypePhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_TYP")
        objContTypePhrase.Load()
        objContMethodPhrase = New MedscreenLib.Glossary.PhraseCollection("CNTCT_METH")
        objContMethodPhrase.Load()
        objRepOptionPhrase = New MedscreenLib.Glossary.PhraseCollection("REP_OPTION")
        objRepOptionPhrase.Load()


        '<Added code modified 05-Jul-2007 09:18 by LOGOS\Taylor> 
        objContStatusPhrase = New MedscreenLib.Glossary.PhraseCollection("CONT_STAT")
        objContStatusPhrase.Load()

        '</Added code modified 05-Jul-2007 09:18 by LOGOS\Taylor> 

        Me.cbContactType.DataSource = objContTypePhrase
        Me.cbContactType.DisplayMember = "PhraseText"
        Me.cbContactType.ValueMember = "PhraseID"

        Me.cbContMethod.DataSource = objContMethodPhrase
        Me.cbContMethod.DisplayMember = "PhraseText"
        Me.cbContMethod.ValueMember = "PhraseID"

        Me.cbReportOptions.DataSource = objRepOptionPhrase
        Me.cbReportOptions.DisplayMember = "PhraseText"
        Me.cbReportOptions.ValueMember = "PhraseID"


        Me.cbContactStatus.DataSource = objContStatusPhrase
        Me.cbContactStatus.DisplayMember = "PhraseText"
        Me.cbContactStatus.ValueMember = "PhraseID"



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
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents cbProgMan As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtContact As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFax As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DTProgStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTProgEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbContactFrequency As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents ckRemoved As System.Windows.Forms.CheckBox
    Friend WithEvents cbProgType As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cbTestFreq As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents cmdAddComment As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents pnlContactList As System.Windows.Forms.Panel
    Friend WithEvents cmdDelContact As System.Windows.Forms.Button
    Friend WithEvents cmdAddContact As System.Windows.Forms.Button
    Friend WithEvents lvContacts As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents pnlContact As System.Windows.Forms.Panel
    Friend WithEvents cbContactStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label87 As System.Windows.Forms.Label
    Friend WithEvents cmdUpdateContact As System.Windows.Forms.Button
    Friend WithEvents cbReportOptions As System.Windows.Forms.ComboBox
    Friend WithEvents Label85 As System.Windows.Forms.Label
    Friend WithEvents cbReportFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label86 As System.Windows.Forms.Label
    Friend WithEvents cbContMethod As System.Windows.Forms.ComboBox
    Friend WithEvents Label84 As System.Windows.Forms.Label
    Friend WithEvents cbContactType As System.Windows.Forms.ComboBox
    Friend WithEvents Label83 As System.Windows.Forms.Label
    Friend WithEvents AdrContact As MedscreenCommonGui.AddressPanel
    Friend WithEvents Label81 As System.Windows.Forms.Label
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents cmdFindContactAdr As System.Windows.Forms.Button
    Friend WithEvents cmdNewContactAdr As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProgrammeManage))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.cbProgMan = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtContact = New System.Windows.Forms.TextBox()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DTProgStart = New System.Windows.Forms.DateTimePicker()
        Me.DTProgEnd = New System.Windows.Forms.DateTimePicker()
        Me.cbContactFrequency = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cbProgType = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cbTestFreq = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ckRemoved = New System.Windows.Forms.CheckBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.cmdAddComment = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.pnlContact = New System.Windows.Forms.Panel()
        Me.cbContactStatus = New System.Windows.Forms.ComboBox()
        Me.Label87 = New System.Windows.Forms.Label()
        Me.cmdUpdateContact = New System.Windows.Forms.Button()
        Me.cbReportOptions = New System.Windows.Forms.ComboBox()
        Me.Label85 = New System.Windows.Forms.Label()
        Me.cbReportFormat = New System.Windows.Forms.ComboBox()
        Me.Label86 = New System.Windows.Forms.Label()
        Me.cbContMethod = New System.Windows.Forms.ComboBox()
        Me.Label84 = New System.Windows.Forms.Label()
        Me.cbContactType = New System.Windows.Forms.ComboBox()
        Me.Label83 = New System.Windows.Forms.Label()
        Me.AdrContact = New MedscreenCommonGui.AddressPanel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label81 = New System.Windows.Forms.Label()
        Me.Label82 = New System.Windows.Forms.Label()
        Me.cmdFindContactAdr = New System.Windows.Forms.Button()
        Me.cmdNewContactAdr = New System.Windows.Forms.Button()
        Me.pnlContactList = New System.Windows.Forms.Panel()
        Me.cmdDelContact = New System.Windows.Forms.Button()
        Me.cmdAddContact = New System.Windows.Forms.Button()
        Me.lvContacts = New System.Windows.Forms.ListView()
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.pnlContact.SuspendLayout()
        Me.AdrContact.SuspendLayout()
        Me.pnlContactList.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 605)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(752, 40)
        Me.Panel1.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(667, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(581, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCustomer
        '
        Me.lblCustomer.Location = New System.Drawing.Point(24, 16)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.Size = New System.Drawing.Size(280, 24)
        Me.lblCustomer.TabIndex = 2
        '
        'cbProgMan
        '
        Me.cbProgMan.Location = New System.Drawing.Point(144, 48)
        Me.cbProgMan.Name = "cbProgMan"
        Me.cbProgMan.Size = New System.Drawing.Size(121, 21)
        Me.cbProgMan.TabIndex = 51
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(20, 48)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 23)
        Me.Label13.TabIndex = 50
        Me.Label13.Text = "Programme Manager"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(288, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 52
        Me.Label1.Text = "Contact"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Visible = False
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(384, 48)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Size = New System.Drawing.Size(144, 20)
        Me.txtContact.TabIndex = 53
        Me.txtContact.Text = ""
        Me.txtContact.Visible = False
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(384, 80)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(144, 20)
        Me.txtPhone.TabIndex = 55
        Me.txtPhone.Text = ""
        Me.txtPhone.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(312, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 54
        Me.Label2.Text = "Phone"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label2.Visible = False
        '
        'txtFax
        '
        Me.txtFax.Location = New System.Drawing.Point(384, 112)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(144, 20)
        Me.txtFax.TabIndex = 57
        Me.txtFax.Text = ""
        Me.txtFax.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(312, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 56
        Me.Label3.Text = "Fax"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Visible = False
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(104, 144)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(424, 20)
        Me.txtEmail.TabIndex = 59
        Me.txtEmail.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 58
        Me.Label4.Text = "Email"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(20, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 23)
        Me.Label5.TabIndex = 60
        Me.Label5.Text = "Programme Start"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(44, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 23)
        Me.Label6.TabIndex = 61
        Me.Label6.Text = "Programme End"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTProgStart
        '
        Me.DTProgStart.Location = New System.Drawing.Point(144, 80)
        Me.DTProgStart.Name = "DTProgStart"
        Me.DTProgStart.Size = New System.Drawing.Size(136, 20)
        Me.DTProgStart.TabIndex = 62
        '
        'DTProgEnd
        '
        Me.DTProgEnd.Location = New System.Drawing.Point(144, 112)
        Me.DTProgEnd.Name = "DTProgEnd"
        Me.DTProgEnd.Size = New System.Drawing.Size(136, 20)
        Me.DTProgEnd.TabIndex = 63
        '
        'cbContactFrequency
        '
        Me.cbContactFrequency.Location = New System.Drawing.Point(144, 192)
        Me.cbContactFrequency.Name = "cbContactFrequency"
        Me.cbContactFrequency.Size = New System.Drawing.Size(384, 21)
        Me.cbContactFrequency.TabIndex = 65
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 192)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 23)
        Me.Label7.TabIndex = 64
        Me.Label7.Text = "Contact Frequency"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2, Me.TabPage3})
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(752, 605)
        Me.TabControl1.TabIndex = 71
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label11, Me.cbProgType, Me.Label9, Me.cbTestFreq, Me.Label8, Me.ckRemoved, Me.cbContactFrequency, Me.Label7, Me.DTProgEnd, Me.DTProgStart, Me.Label6, Me.Label5, Me.txtEmail, Me.Label4, Me.txtFax, Me.Label3, Me.txtPhone, Me.Label2, Me.txtContact, Me.Label1, Me.cbProgMan, Me.Label13, Me.lblCustomer})
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(744, 579)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Details"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.IndianRed
        Me.Label11.Location = New System.Drawing.Point(368, 48)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(168, 88)
        Me.Label11.TabIndex = 76
        Me.Label11.Text = "Contacts are on a separate tab, there can be more than one contact."
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbProgType
        '
        Me.cbProgType.Location = New System.Drawing.Point(144, 256)
        Me.cbProgType.Name = "cbProgType"
        Me.cbProgType.Size = New System.Drawing.Size(384, 21)
        Me.cbProgType.TabIndex = 75
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(32, 256)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 23)
        Me.Label9.TabIndex = 74
        Me.Label9.Text = "Programme Type"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbTestFreq
        '
        Me.cbTestFreq.Location = New System.Drawing.Point(144, 224)
        Me.cbTestFreq.Name = "cbTestFreq"
        Me.cbTestFreq.Size = New System.Drawing.Size(384, 21)
        Me.cbTestFreq.TabIndex = 73
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(32, 224)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 23)
        Me.Label8.TabIndex = 72
        Me.Label8.Text = "Test Frequency"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ckRemoved
        '
        Me.ckRemoved.Location = New System.Drawing.Point(24, 288)
        Me.ckRemoved.Name = "ckRemoved"
        Me.ckRemoved.TabIndex = 71
        Me.ckRemoved.Text = "Removed"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAddComment, Me.ListView1})
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(744, 579)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Comments"
        '
        'cmdAddComment
        '
        Me.cmdAddComment.Location = New System.Drawing.Point(48, 280)
        Me.cmdAddComment.Name = "cmdAddComment"
        Me.cmdAddComment.Size = New System.Drawing.Size(216, 23)
        Me.cmdAddComment.TabIndex = 1
        Me.cmdAddComment.Text = "Add Comment"
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(744, 248)
        Me.ListView1.SmallImageList = Me.ImageList1
        Me.ListView1.TabIndex = 0
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Created"
        Me.ColumnHeader1.Width = 80
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Width = 1
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Comment"
        Me.ColumnHeader3.Width = 600
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 40)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'TabPage3
        '
        Me.TabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlContact, Me.AdrContact, Me.pnlContactList})
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(744, 579)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Contacts"
        '
        'pnlContact
        '
        Me.pnlContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlContact.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbContactStatus, Me.Label87, Me.cmdUpdateContact, Me.cbReportOptions, Me.Label85, Me.cbReportFormat, Me.Label86, Me.cbContMethod, Me.Label84, Me.cbContactType, Me.Label83})
        Me.pnlContact.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContact.Location = New System.Drawing.Point(576, 256)
        Me.pnlContact.Name = "pnlContact"
        Me.pnlContact.Size = New System.Drawing.Size(168, 323)
        Me.pnlContact.TabIndex = 240
        '
        'cbContactStatus
        '
        Me.cbContactStatus.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbContactStatus.Location = New System.Drawing.Point(44, 264)
        Me.cbContactStatus.Name = "cbContactStatus"
        Me.cbContactStatus.Size = New System.Drawing.Size(113, 21)
        Me.cbContactStatus.TabIndex = 235
        '
        'Label87
        '
        Me.Label87.Location = New System.Drawing.Point(44, 240)
        Me.Label87.Name = "Label87"
        Me.Label87.TabIndex = 234
        Me.Label87.Text = "Contact Status"
        '
        'cmdUpdateContact
        '
        Me.cmdUpdateContact.Image = CType(resources.GetObject("cmdUpdateContact.Image"), System.Drawing.Bitmap)
        Me.cmdUpdateContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdateContact.Location = New System.Drawing.Point(40, 296)
        Me.cmdUpdateContact.Name = "cmdUpdateContact"
        Me.cmdUpdateContact.Size = New System.Drawing.Size(112, 23)
        Me.cmdUpdateContact.TabIndex = 233
        Me.cmdUpdateContact.Text = "Save"
        Me.cmdUpdateContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbReportOptions
        '
        Me.cbReportOptions.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbReportOptions.Location = New System.Drawing.Point(40, 212)
        Me.cbReportOptions.Name = "cbReportOptions"
        Me.cbReportOptions.Size = New System.Drawing.Size(113, 21)
        Me.cbReportOptions.TabIndex = 7
        '
        'Label85
        '
        Me.Label85.Location = New System.Drawing.Point(40, 188)
        Me.Label85.Name = "Label85"
        Me.Label85.TabIndex = 6
        Me.Label85.Text = "Report Options"
        '
        'cbReportFormat
        '
        Me.cbReportFormat.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbReportFormat.Location = New System.Drawing.Point(40, 156)
        Me.cbReportFormat.Name = "cbReportFormat"
        Me.cbReportFormat.Size = New System.Drawing.Size(113, 21)
        Me.cbReportFormat.TabIndex = 5
        '
        'Label86
        '
        Me.Label86.Location = New System.Drawing.Point(40, 132)
        Me.Label86.Name = "Label86"
        Me.Label86.TabIndex = 4
        Me.Label86.Text = "Report Format"
        '
        'cbContMethod
        '
        Me.cbContMethod.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbContMethod.Location = New System.Drawing.Point(40, 96)
        Me.cbContMethod.Name = "cbContMethod"
        Me.cbContMethod.Size = New System.Drawing.Size(113, 21)
        Me.cbContMethod.TabIndex = 3
        '
        'Label84
        '
        Me.Label84.Location = New System.Drawing.Point(40, 72)
        Me.Label84.Name = "Label84"
        Me.Label84.TabIndex = 2
        Me.Label84.Text = "Contact Method"
        '
        'cbContactType
        '
        Me.cbContactType.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbContactType.Location = New System.Drawing.Point(40, 40)
        Me.cbContactType.Name = "cbContactType"
        Me.cbContactType.Size = New System.Drawing.Size(113, 21)
        Me.cbContactType.TabIndex = 1
        '
        'Label83
        '
        Me.Label83.Location = New System.Drawing.Point(40, 16)
        Me.Label83.Name = "Label83"
        Me.Label83.TabIndex = 0
        Me.Label83.Text = "Contact Type"
        '
        'AdrContact
        '
        Me.AdrContact.Address = Nothing
        Me.AdrContact.AddressType = MedscreenLib.Constants.AddressType.AddressMain
        Me.AdrContact.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AdrContact.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label10, Me.Label81, Me.Label82, Me.cmdFindContactAdr, Me.cmdNewContactAdr})
        Me.AdrContact.Dock = System.Windows.Forms.DockStyle.Left
        Me.AdrContact.HiLiteControl = MedscreenCommonGui.AddressPanel.HighlightedControl.none
        Me.AdrContact.InfoVisible = False
        Me.AdrContact.isReadOnly = False
        Me.AdrContact.Location = New System.Drawing.Point(0, 256)
        Me.AdrContact.Name = "AdrContact"
        Me.AdrContact.ShowContact = True
        Me.AdrContact.ShowStatus = False
        Me.AdrContact.Size = New System.Drawing.Size(576, 323)
        Me.AdrContact.TabIndex = 239
        Me.AdrContact.Tooltip = Nothing
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 296)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(440, 23)
        Me.Label10.TabIndex = 236
        Me.Label10.Text = "Press Save button on contact panel to save changes to an address"
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
        Me.cmdFindContactAdr.Image = CType(resources.GetObject("cmdFindContactAdr.Image"), System.Drawing.Bitmap)
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
        Me.cmdNewContactAdr.Image = CType(resources.GetObject("cmdNewContactAdr.Image"), System.Drawing.Bitmap)
        Me.cmdNewContactAdr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdNewContactAdr.Location = New System.Drawing.Point(64, 256)
        Me.cmdNewContactAdr.Name = "cmdNewContactAdr"
        Me.cmdNewContactAdr.Size = New System.Drawing.Size(112, 23)
        Me.cmdNewContactAdr.TabIndex = 232
        Me.cmdNewContactAdr.Text = "New Address"
        Me.cmdNewContactAdr.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlContactList
        '
        Me.pnlContactList.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdDelContact, Me.cmdAddContact, Me.lvContacts})
        Me.pnlContactList.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactList.Name = "pnlContactList"
        Me.pnlContactList.Size = New System.Drawing.Size(744, 256)
        Me.pnlContactList.TabIndex = 76
        '
        'cmdDelContact
        '
        Me.cmdDelContact.Image = CType(resources.GetObject("cmdDelContact.Image"), System.Drawing.Bitmap)
        Me.cmdDelContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDelContact.Location = New System.Drawing.Point(152, 216)
        Me.cmdDelContact.Name = "cmdDelContact"
        Me.cmdDelContact.Size = New System.Drawing.Size(96, 23)
        Me.cmdDelContact.TabIndex = 2
        Me.cmdDelContact.Text = "Del Contact"
        Me.cmdDelContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAddContact
        '
        Me.cmdAddContact.Image = CType(resources.GetObject("cmdAddContact.Image"), System.Drawing.Bitmap)
        Me.cmdAddContact.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAddContact.Location = New System.Drawing.Point(40, 216)
        Me.cmdAddContact.Name = "cmdAddContact"
        Me.cmdAddContact.Size = New System.Drawing.Size(96, 23)
        Me.cmdAddContact.TabIndex = 1
        Me.cmdAddContact.Text = "Add Contact"
        Me.cmdAddContact.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvContacts
        '
        Me.lvContacts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader23, Me.ColumnHeader24, Me.ColumnHeader28, Me.ColumnHeader25, Me.ColumnHeader26, Me.ColumnHeader27})
        Me.lvContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvContacts.FullRowSelect = True
        Me.lvContacts.Name = "lvContacts"
        Me.lvContacts.Size = New System.Drawing.Size(744, 200)
        Me.lvContacts.TabIndex = 0
        Me.lvContacts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "Contact"
        Me.ColumnHeader23.Width = 180
        '
        'ColumnHeader24
        '
        Me.ColumnHeader24.Text = "Contact Type"
        Me.ColumnHeader24.Width = 100
        '
        'ColumnHeader28
        '
        Me.ColumnHeader28.Text = "Status"
        Me.ColumnHeader28.Width = 80
        '
        'ColumnHeader25
        '
        Me.ColumnHeader25.Text = "Contact Method"
        Me.ColumnHeader25.Width = 121
        '
        'ColumnHeader26
        '
        Me.ColumnHeader26.Text = "Report Format"
        Me.ColumnHeader26.Width = 120
        '
        'ColumnHeader27
        '
        Me.ColumnHeader27.Text = "Report Options"
        Me.ColumnHeader27.Width = 144
        '
        'frmProgrammeManage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(752, 645)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.Panel1})
        Me.Name = "frmProgrammeManage"
        Me.Text = "Collection Programme Management"
        Me.Panel1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.pnlContact.ResumeLayout(False)
        Me.AdrContact.ResumeLayout(False)
        Me.pnlContactList.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declarations"
    Private myProgramme As ProgrammeManagement
    Private myClient As Client
    Private objContTypePhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContMethodPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objRepOptionPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objContStatusPhrase As MedscreenLib.Glossary.PhraseCollection

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the programme 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Program() As ProgrammeManagement
        Get
            Return Me.myProgramme
        End Get
        Set(ByVal Value As ProgrammeManagement)
            Me.myProgramme = Value

            If Not Value Is Nothing Then
                With myProgramme
                    Me.myClient = New Client(.CustomerID, True)


                    Me.cbProgMan.SelectedIndex = -1
                    Me.cbProgMan.Text = ""
                    Dim ik As Integer
                    For ik = 0 To Me.cbProgMan.Items.Count - 1
                        If Me.cbProgMan.Items(ik).user = myProgramme.ProgrammeManager Then
                            Me.cbProgMan.SelectedIndex = ik
                            Me.cbProgMan.Text = Me.cbProgMan.Items(ik).user
                            Exit For
                        End If
                    Next

                    Me.DTProgStart.Value = .ProgramStartDate
                    Me.DTProgEnd.Value = .ProgramEndDate

                    Me.txtContact.Text = .ContactList
                    Me.txtEmail.Text = .ContactSendEmail
                    Me.txtFax.Text = .ContactFaxNumber
                    Me.txtPhone.Text = .ContactPhoneNumber

                    Me.cbContactFrequency.SelectedIndex = MedscreenLib.Glossary.Glossary.PMCustomerContactFrequency.IndexOf(.ContactFrequency)
                    Me.cbProgType.SelectedIndex = MedscreenLib.Glossary.Glossary.PMAccountType.IndexOf(.ProgramType)
                    Me.cbTestFreq.SelectedIndex = MedscreenLib.Glossary.Glossary.PMVesselTestFrequency.IndexOf(.VesselTestFrequency)

                    Me.ckRemoved.Checked = .Removed
                    DrawComments()
                    Dim objContact As contacts.CustomerContact = myClient.CustomerContacts.Item(1, customerns.CustColl.GCST_Coll_contact_PM)
                    If objContact Is Nothing And myProgramme.ContactName.Trim.Length > 0 Then
                        objContact = Me.myClient.CustomerContacts.CreateCustomerContact()
                        objContact.ContactType = customerns.CustColl.GCST_Coll_contact_PM       'Set to right type
                        objContact.Contact() = myProgramme.ContactName
                        objContact.ContactStatus = "ACTIVE"
                        If objContact.Address Is Nothing Then                                   'We will need an address
                            objContact.Address = New MedscreenLib.Address.Caddress()
                            objContact.Address.AddressId = MedscreenLib.Address.AddressCollection.NextAddressId
                            objContact.Address.AddressType = "CONT"
                            objContact.Address.Update()                                         'Update the address
                            objContact.AddressId = objContact.Address.AddressId
                            objContact.Address.Email = myProgramme.Email
                            myProgramme.Email = ""
                            objContact.Address.Phone = myProgramme.ContactPhoneNumber
                            myProgramme.ContactPhoneNumber = ""
                            objContact.Address.Fax = myProgramme.ContactFaxNumber
                            myProgramme.ContactPhoneNumber = ""
                            objContact.ContactMethod = "EMAIL"
                            objContact.Address.Update()
                            objContact.Update()
                        End If


                    End If

                    Me.drawContactList()

                End With
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose Company name property 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Company() As String
        Set(ByVal Value As String)
            Me.lblCustomer.Text = Value

        End Set
    End Property

#End Region
#End Region

#Region "Private Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get form content and bugger off
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [19/01/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim blnNoError As Boolean = True
        Dim phoneError As MedscreenLib.Constants.PhoneNoErrors
        Dim emailError As MedscreenLib.Constants.EmailErrors

        With myProgramme
            .ProgramStartDate = Me.DTProgStart.Value
            .ProgramEndDate = Me.DTProgEnd.Value

            '.Contact = Me.txtContact.Text

            Me.ErrorProvider1.SetError(Me.txtEmail, "")
            emailError = MedscreenLib.Medscreen.ValidateEmail(Me.txtEmail.Text)
            'Check email address is valid 
            If emailError = MedscreenLib.Constants.EmailErrors.None Then
                '.ContactEmail = Me.txtEmail.Text
            ElseIf emailError = MedscreenLib.Constants.EmailErrors.NoDomain Then
                Me.ErrorProvider1.SetError(Me.txtEmail, "An email address must have a valid domain")
                blnNoError = False
            ElseIf emailError = MedscreenLib.Constants.EmailErrors.NoAt Then
                Me.ErrorProvider1.SetError(Me.txtEmail, "An email address must have an @ symbol")
                blnNoError = False
            End If
            'Fax and phone numbers
            Me.ErrorProvider1.SetError(Me.txtFax, "")
            phoneError = MedscreenLib.Address.Caddress.ValidatePhoneNum(Me.txtFax.Text)
            If phoneError = MedscreenLib.Constants.PhoneNoErrors.NoError Then
                '.ContactFax = Me.txtFax.Text
            ElseIf phoneError = MedscreenLib.Constants.PhoneNoErrors.IllegalCharacterPresent Then
                Me.ErrorProvider1.SetError(Me.txtFax, "Illegal characters in phone number")
                blnNoError = False
            ElseIf phoneError = MedscreenLib.Constants.PhoneNoErrors.InTDialCodePresent Then
                '.ContactFax = Me.txtFax.Text
            End If

            Me.ErrorProvider1.SetError(Me.txtPhone, "")
            phoneError = MedscreenLib.Address.Caddress.ValidatePhoneNum(Me.txtPhone.Text)

            If phoneError = MedscreenLib.Constants.PhoneNoErrors.NoError Then
                '.ContactPhone = Me.txtPhone.Text
            ElseIf phoneError = MedscreenLib.Constants.PhoneNoErrors.IllegalCharacterPresent Then
                Me.ErrorProvider1.SetError(Me.txtPhone, "Illegal characters in phone number")
                blnNoError = False
            ElseIf phoneError = MedscreenLib.Constants.PhoneNoErrors.InTDialCodePresent Then
                '.ContactPhone = Me.txtPhone.Text
            End If

            Dim objPhrase As MedscreenLib.Glossary.Phrase
            objPhrase = MedscreenLib.Glossary.Glossary.PMCustomerContactFrequency.Item(Me.cbContactFrequency.SelectedIndex)
            .ContactFrequency = objPhrase.PhraseID
            objPhrase = MedscreenLib.Glossary.Glossary.PMAccountType.Item(Me.cbProgType.SelectedIndex)
            .ProgramType = objPhrase.PhraseID
            objPhrase = MedscreenLib.Glossary.Glossary.PMVesselTestFrequency.Item(Me.cbTestFreq.SelectedIndex)
            .VesselTestFrequency = objPhrase.PhraseID

            .Removed = Me.ckRemoved.Checked
            .ProgrammeManager = Me.cbProgMan.Text

        End With

        If Not blnNoError Then
            Me.DialogResult = Windows.Forms.DialogResult.None
        Else
            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
    End Sub

    Private Sub DrawComments()
        Dim objCommList As New Comment.Collection(Comment.CommentOn.Customer, Me.Program.CustomerID, "PM")
        objCommList.Load()

        'Dim strQuery As String = "Select * from MED_CUSTOMER_COMMENT where customer_id = ? and COMMENT_TYPE = ? order by COMMENT_NO "
        Me.ListView1.Items.Clear()
        '   Dim strQuery As String = "Select * from MED_CUSTOMER_COMMENT where customer_id = ? and COMMENT_TYPE = ? order by COMMENT_NO "
        'Me.ListView1.Items.Clear()
        'Dim oRead As OleDb.OleDbDataReader
        'Dim oCmd As New OleDb.OleDbCommand()
        'Dim oColl As New Collection()
        Try ' Protecting 
            'oCmd.Connection = CConnection.DbConnection
            'oCmd.CommandText = strQuery
            'oCmd.Parameters.Add(CConnection.StringParameter("CustId", Me.Program.CustomerID, 10))
            'oCmd.Parameters.Add(CConnection.StringParameter("Type", "PM", 10))
            'If CConnection.ConnOpen Then            'Attempt to open reader
            '    oRead = oCmd.ExecuteReader          'Get Date
            Dim objPm As Comment
            For Each objPm In objCommList
                Dim lvitem As New ListViewItems.lvComment(objPm)
                Me.ListView1.Items.Add(lvitem)
                'End While
                'End If
            Next

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "frmProgrammeManage-DrawComments-588")
        Finally
            'CConnection.SetConnClosed()             'Close connection
            'If Not oRead Is Nothing Then            'Try and close reader
            '    If Not oRead.IsClosed Then oRead.Close()
            'End If
        End Try
    End Sub
#End Region

    Dim Currentlvitem As ListViewItems.lvComment
    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Currentlvitem = Nothing
        If Me.ListView1.SelectedItems.Count = 1 Then
            Currentlvitem = Me.ListView1.SelectedItems.Item(0)
        End If
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If Currentlvitem Is Nothing Then Exit Sub
        If Currentlvitem.Comment Is Nothing Then Exit Sub
        Dim objRet As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", , Currentlvitem.Comment.CommentText, False)
        If Not objRet Is Nothing Then
            Currentlvitem.Comment.CommentText = objRet
            Currentlvitem.Comment.Update()
        End If
        DrawComments()
    End Sub

    Private Sub cmdAddComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddComment.Click
        If Me.Program Is Nothing Then Exit Sub
        Dim objCust As New Intranet.intranet.customerns.Client(Me.Program.CustomerID, True)
        'Dim objComment As jobs.CollectionComment = objCust.CustomerComments.CreateComment("", "PM", jobs.CollectionCommentList.CommentType.Customer)
        'Dim objRet As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", , objComment.CommentText, False)
        Dim objRet As Object = Medscreen.GetParameter(Medscreen.MyTypes.typString, "Comment", , "", False)

        If Not objRet Is Nothing Then
            MedscreenLib.Comment.CreateComment(Comment.CommentOn.Customer, objCust.Identity, _
"PM", objRet, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity)
            'objComment.CommentText = objRet
            'objComment.Update()
        End If
        DrawComments()
    End Sub

    Private Sub drawContactList()
        Me.lvContacts.Items.Clear()

        Dim objContact As contacts.CustomerContact
        For Each objContact In Me.myClient.CustomerContacts
            If objContact.ContactType = "PM" Then
                Dim objLvCont As New MedscreenCommonGui.ListViewItems.LVCustomerContact(objContact)
                Me.lvContacts.Items.Add(objLvCont)
            End If
        Next
        If lvContacts.Items.Count > 0 Then
            Me.lvContacts.Focus()
            Me.lvContacts.Items(0).Selected = True
            Me.lvContacts_Click(Nothing, Nothing)
        End If

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
    Private Sub cmdAddContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddContact.Click
        If Me.myClient Is Nothing Then Exit Sub

        Dim objContact As contacts.CustomerContact = myClient.CustomerContacts.CreateCustomerContact()
        objContact.ContactType = contacts.CustomerContact.GCST_PMContact
        objContact.Update()
        Me.drawContactList()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Variable to hold contact row
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private myLvContact As MedscreenCommonGui.ListViewItems.LVCustomerContact

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
    Private Sub lvContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvContacts.Click
        myLvContact = Nothing
        Me.cmdDelContact.Enabled = False
        If Me.lvContacts.SelectedItems.Count = 1 Then
            myLvContact = Me.lvContacts.SelectedItems.Item(0)

            If Not myLvContact Is Nothing Then
                Me.AdrContact.Address = myLvContact.Contact.Address
                Me.cbContactType.SelectedValue = myLvContact.Contact.ContactType
                Me.cbContMethod.SelectedValue = myLvContact.Contact.ContactMethod
                Me.cbReportOptions.SelectedValue = myLvContact.Contact.ReportOptions
                Me.cbContactStatus.SelectedValue = myLvContact.Contact.ContactStatus
                Me.cmdDelContact.Enabled = True
            Else
                Me.AdrContact.Address = Nothing
                Me.cbContactType.SelectedValue = ""
                Me.cbContMethod.SelectedValue = ""
                Me.cbReportOptions.SelectedValue = ""

            End If
        End If

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
    Private Sub cmdUpdateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateContact.Click
        'Update info 
        If myLvContact Is Nothing Then Exit Sub
        myLvContact.Contact.AddressId = Me.AdrContact.Address.AddressId
        myLvContact.Contact.Address = Me.AdrContact.Address
        myLvContact.Contact.Address.Update()
        myLvContact.Contact.ContactMethod = Me.cbContMethod.SelectedValue
        myLvContact.Contact.ContactType = Me.cbContactType.SelectedValue
        myLvContact.Contact.ReportOptions = Me.cbReportOptions.SelectedValue
        myLvContact.Contact.ContactStatus = Me.cbContactStatus.SelectedValue

        myLvContact.Contact.Update()
        mylvcontact.Refresh()
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
        If Me.myLvContact Is Nothing Then Exit Sub
        Dim addr As MedscreenLib.Address.Caddress
        addr = MedscreenLib.Glossary.Glossary.SMAddresses.CreateAddress(, "", Me.myClient.Identity, MedscreenLib.Constants.AddressType.AddressMain)
        Me.AdrContact.Address = addr
        Me.myLvContact.Contact.Address = addr
        Me.myLvContact.Contact.AddressId = addr.AddressId


    End Sub

    Private Sub cmdFindContactAdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindContactAdr.Click
        If Me.myLvContact Is Nothing Then Exit Sub

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
                    Me.myLvContact.Contact.Address = addr
                    Me.myLvContact.Contact.AddressId = addr.AddressId

                    'MedscreenLib.Address.AddressCollection.AddLink(addr.AddressId, "Customer.Identity", Me.Client.Identity, Me.AddressUsed)
                    'TODO: Need to change address links
                End If
                'make controls invisible
                'Me.cmdNewAddress.Visible = False
                'Me.cmdFindAddress.Visible = False

            End If
        End With

    End Sub
End Class
