'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-09-02 07:41:45+01 $
'$Log: Class1.vb,v $
'<>

'''<summary>
''' Form to allow users to enter parameters
''' </summary>
Public Class frmParameter
    Inherits System.Windows.Forms.Form


    Private MyType As Medscreen.MyTypes
    Private CurrentDate As Date

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
    Friend WithEvents Parameter As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents ckRemember As System.Windows.Forms.CheckBox
    Friend WithEvents cmbOption As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmParameter))
        Me.Parameter = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.ckRemember = New System.Windows.Forms.CheckBox()
        Me.cmbOption = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Parameter
        '
        Me.Parameter.Location = New System.Drawing.Point(8, 8)
        Me.Parameter.Name = "Parameter"
        Me.Parameter.Size = New System.Drawing.Size(288, 40)
        Me.Parameter.TabIndex = 4
        Me.Parameter.Text = "Label1"
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.Button1})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 225)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(292, 48)
        Me.Panel1.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(208, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(120, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Ok"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Location = New System.Drawing.Point(20, 41)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 3
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(48, 43)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(192, 152)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = ""
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(24, 56)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000000, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(176, 20)
        Me.NumericUpDown1.TabIndex = 2
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(24, 56)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.TabIndex = 1
        Me.CheckBox1.Text = "CheckBox1"
        Me.CheckBox1.Visible = False
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd-MMM-yyyy HH:mm"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(16, 200)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.TabIndex = 6
        Me.DateTimePicker1.Visible = False
        '
        'ckRemember
        '
        Me.ckRemember.Location = New System.Drawing.Point(40, 200)
        Me.ckRemember.Name = "ckRemember"
        Me.ckRemember.Size = New System.Drawing.Size(200, 24)
        Me.ckRemember.TabIndex = 7
        Me.ckRemember.Text = "Remember Entry"
        '
        'cmbOption
        '
        Me.cmbOption.Location = New System.Drawing.Point(24, 56)
        Me.cmbOption.Name = "cmbOption"
        Me.cmbOption.Size = New System.Drawing.Size(232, 21)
        Me.cmbOption.TabIndex = 8
        '
        'frmParameter
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmbOption, Me.ckRemember, Me.DateTimePicker1, Me.CheckBox1, Me.NumericUpDown1, Me.TextBox1, Me.MonthCalendar1, Me.Panel1, Me.Parameter})
        Me.Name = "frmParameter"
        Me.Text = "frmParameter"
        Me.Panel1.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    '''<summary>
    ''' Name of the Parameter
    ''' </summary>
    Public WriteOnly Property ParameterName()
        Set(ByVal Value)
            Me.Text = "Entry of " & Value
            Me.Parameter.Text = Value
            Me.CheckBox1.Text = Value

        End Set
    End Property

    '''<summary>
    ''' Value both the Default Value (in) and the return (out)
    ''' </summary>
    Public Property Value()
        Get
            Select Case Me.MyType
                Case Medscreen.MyTypes.typDate
                    Return CurrentDate
                Case Medscreen.MyTypes.typeInteger
                    Return Me.NumericUpDown1.Value
                Case Medscreen.MyTypes.typString
                    Return Me.TextBox1.Text
                Case Medscreen.MyTypes.typBoolean
                    Return Me.CheckBox1.Checked
                Case Medscreen.MyTypes.typItem
                    Return Me.cmbOption.Text
            End Select
        End Get
        Set(ByVal Value)
            Select Case Me.MyType
                Case Medscreen.MyTypes.typDate
                    CurrentDate = Value
                    Me.MonthCalendar1.TodayDate = Value
                    Me.MonthCalendar1.SetDate(Value)
                    Me.DateTimePicker1.Value = Value
                Case Medscreen.MyTypes.typeInteger
                    Me.NumericUpDown1.Value = Value
                Case Medscreen.MyTypes.typString
                    Me.TextBox1.Text = Value
                Case Medscreen.MyTypes.typBoolean
                    Me.CheckBox1.Checked = Value
                Case Medscreen.MyTypes.typItem
                    '
            End Select
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show the remember control or not 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property ShowRemember() As Boolean

        Set(ByVal Value As Boolean)
            Me.ckRemember.Visible = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Do we remember the value between uses 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Remebered() As Boolean
        Get
            Return Me.ckRemember.Checked
        End Get
        Set(ByVal Value As Boolean)
            Me.ckRemember.Checked = Value
        End Set
    End Property

    '''<summary>
    ''' The type of the Parameter needed
    ''' </summary>
    ''' 
    Property ParameterType() As Medscreen.MyTypes
        Get
            Return MyType
        End Get
        Set(ByVal Value As Medscreen.MyTypes)
            MyType = Value
            Me.ckRemember.Hide()
            Select Case MyType
                Case Medscreen.MyTypes.typDate
                    Me.TextBox1.Visible = False
                    Me.CheckBox1.Hide()
                    Me.NumericUpDown1.Hide()
                    Me.MonthCalendar1.Show()
                    cmbOption.Hide()
                    Me.MonthCalendar1.Focus()
                    Me.DateTimePicker1.Show()


                    CurrentDate = Today
                Case Medscreen.MyTypes.typeInteger
                    Me.TextBox1.Hide()
                    Me.MonthCalendar1.Hide()
                    Me.NumericUpDown1.Show()
                    Me.DateTimePicker1.Hide()
                    Me.CheckBox1.Hide()
                    Me.NumericUpDown1.Focus()
                    cmbOption.Hide()
                Case Medscreen.MyTypes.typString
                    Me.MonthCalendar1.Hide()
                    Me.NumericUpDown1.Hide()
                    Me.CheckBox1.Hide()
                    Me.DateTimePicker1.Hide()
                    Me.TextBox1.Show()
                    cmbOption.Hide()
                    Me.TextBox1.Focus()

                Case Medscreen.MyTypes.typBoolean
                    Me.MonthCalendar1.Hide()
                    Me.NumericUpDown1.Hide()
                    Me.CheckBox1.Show()
                    Me.TextBox1.Hide()
                    Me.DateTimePicker1.Hide()
                    cmbOption.Hide()
                    Me.CheckBox1.Focus()
                Case Medscreen.MyTypes.typItem
                    Me.MonthCalendar1.Hide()
                    Me.NumericUpDown1.Hide()
                    Me.CheckBox1.Hide()
                    Me.TextBox1.Hide()
                    Me.DateTimePicker1.Hide()
                    cmbOption.Show()
                    Me.cmbOption.Focus()

            End Select
        End Set
    End Property

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        CurrentDate = e.Start
        If Me.DateTimePicker1.Value <> CurrentDate Then Me.DateTimePicker1.Value = CurrentDate
    End Sub

    Private Sub frmParameter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List of items
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [21/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Itemlist() As Collection
        Set(ByVal Value As Collection)
            Dim i As Integer
            Me.cmbOption.Items.Clear()
            For i = 1 To Value.Count
                Me.cmbOption.Items.Add(Value.Item(i))


            Next
        End Set
    End Property

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        CurrentDate = Me.DateTimePicker1.Value
        MonthCalendar1.SetDate(CurrentDate)
    End Sub
End Class
