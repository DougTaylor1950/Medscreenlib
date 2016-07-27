Imports Intranet
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmHolidays
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for entering staff holidays
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action>
''' $Revision: 1.1 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-30 06:23:44+00 $ <para/>
'''$Log: frmHolidays.vb,v $
'''Revision 1.1  2005-11-30 06:23:44+00  taylor
'''Commented
''' <para/>
'''Revision 1.0  2005-11-30 06:21:15+00  taylor
'''checked in to fix header
''' <para/>
''' </Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmHolidays
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateCombo1 As UtilityLibrary.Combos.DateCombo
    Friend WithEvents DateCombo2 As UtilityLibrary.Combos.DateCombo
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents Deputy As System.Windows.Forms.Label
    Friend WithEvents txtDeputy As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHolidays))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateCombo1 = New UtilityLibrary.Combos.DateCombo()
        Me.DateCombo2 = New UtilityLibrary.Combos.DateCombo()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.Deputy = New System.Windows.Forms.Label()
        Me.txtDeputy = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.DockPadding.All = 3
        Me.Panel1.Location = New System.Drawing.Point(0, 225)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(464, 48)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtDeputy, Me.Deputy, Me.txtComments, Me.Label3, Me.DateCombo2, Me.Label2, Me.DateCombo1, Me.Label1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(464, 225)
        Me.Panel2.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Start Date"
        '
        'DateCombo1
        '
        Me.DateCombo1.Location = New System.Drawing.Point(160, 8)
        Me.DateCombo1.Name = "DateCombo1"
        Me.DateCombo1.Size = New System.Drawing.Size(144, 23)
        Me.DateCombo1.TabIndex = 1
        Me.DateCombo1.Text = "DateCombo1"
        Me.DateCombo1.Value = New Date(CType(0, Long))
        '
        'DateCombo2
        '
        Me.DateCombo2.Location = New System.Drawing.Point(159, 43)
        Me.DateCombo2.Name = "DateCombo2"
        Me.DateCombo2.Size = New System.Drawing.Size(144, 23)
        Me.DateCombo2.TabIndex = 3
        Me.DateCombo2.Text = "DateCombo2"
        Me.DateCombo2.Value = New Date(CType(0, Long))
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "End Date"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(30, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Comments"
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(160, 72)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(280, 88)
        Me.txtComments.TabIndex = 5
        Me.txtComments.Text = ""
        '
        'Deputy
        '
        Me.Deputy.Location = New System.Drawing.Point(35, 168)
        Me.Deputy.Name = "Deputy"
        Me.Deputy.TabIndex = 6
        Me.Deputy.Text = "Deputy"
        '
        'txtDeputy
        '
        Me.txtDeputy.Location = New System.Drawing.Point(160, 168)
        Me.txtDeputy.Name = "txtDeputy"
        Me.txtDeputy.TabIndex = 7
        Me.txtDeputy.Text = ""
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(376, 12)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(264, 12)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmHolidays
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1})
        Me.Name = "frmHolidays"
        Me.Text = "Holidays"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myHoliday As Intranet.intranet.jobs.CollectorHoliday

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holiday passed to form
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Holiday() As Intranet.intranet.jobs.CollectorHoliday
        Get
            Return myHoliday
        End Get
        Set(ByVal Value As Intranet.intranet.jobs.CollectorHoliday)
            myHoliday = Value
            If myHoliday Is Nothing Then
            Else
                DrawForm()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get Holiday info back into object
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub GetForm()
        myHoliday.HolidayStart = Me.DateCombo1.Value
        myHoliday.HolidayEnd = Me.DateCombo2.Value
        myHoliday.Comment = Me.txtComments.Text
        myHoliday.DeputyId = Me.txtDeputy.Text

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill controls with Holiday information
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawForm()
        Me.DateCombo1.Value = myHoliday.HolidayStart
        Me.DateCombo2.Value = myHoliday.HolidayEnd
        Me.txtComments.Text = myHoliday.Comment
        Me.txtDeputy.Text = myHoliday.DeputyId
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exit and return 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        GetForm()
    End Sub
End Class
